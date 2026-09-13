"""Use real processes to exercise shared recording-directory cleanup."""

import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

CLIENT_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(CLIENT_ROOT / "sms"))
sys.path.insert(0, str(CLIENT_ROOT))

from sms_core.call_recordings import CallRecordingRepository, SerialCallRecordingReceiver
from tests.test_call_recordings import AMR_PAYLOAD, begin_frame, chunk_frame, end_frame


IMEI_A = "000000000000001"
IMEI_B = "000000000000002"
RECORDING_ID = "synthetic-shared-recording"


def run_child(root, operation):
    process = subprocess.run(
        [sys.executable, "-X", "utf8", "-B", str(Path(__file__).resolve()),
         "--child", str(root), operation],
        cwd=CLIENT_ROOT, capture_output=True, text=True, encoding="utf-8", timeout=15, check=True,
    )
    if process.stderr:
        raise AssertionError(process.stderr)
    return json.loads(process.stdout)


class RecordingIncomingFileTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="sms-recording-files-")
        self.root = Path(self.temp.name)
        self.repository = CallRecordingRepository(self.root)
        self.saved = []
        self.aborted = []
        self.receiver = SerialCallRecordingReceiver(
            self.repository, on_saved=self.saved.append,
            on_aborted=lambda _metadata, reason: self.aborted.append(reason),
        )
        self.addCleanup(self.temp.cleanup)
        self.addCleanup(self.receiver.abort)

    def begin(self):
        self.receiver.consume_line(begin_frame(RECORDING_ID, imei=IMEI_A))
        self.assertIsNotNone(self.receiver.current)
        self.receiver.consume_line(chunk_frame(RECORDING_ID, 1, AMR_PAYLOAD))
        return Path(self.receiver.current.temp_path)

    def end(self):
        self.receiver.consume_line(end_frame(RECORDING_ID, 1))

    def assert_saved(self):
        self.assertEqual(self.aborted, [])
        self.assertEqual(len(self.saved), 1)
        self.assertEqual(Path(self.saved[0].path).read_bytes(), AMR_PAYLOAD)

    def test_other_process_cannot_delete_closed_file_before_commit_even_if_old(self):
        original = self.repository.commit
        observed = []

        def commit_at_boundary(path, metadata, digest):
            # Age alone must not decide whether a modern transfer is active.
            old = time.time() - 7 * 86400
            os.utime(path, (old, old))
            run_child(self.root, "cleanup")
            observed.append(Path(path).is_file())
            return original(path, metadata, digest)

        self.repository.commit = commit_at_boundary
        self.begin()
        self.end()
        self.assertEqual(observed, [True])
        self.assert_saved()
        self.assertEqual(list((self.root / ".incoming").iterdir()), [])

    def test_active_receiver_survives_other_process_startup(self):
        partial = self.begin()
        run_child(self.root, "cleanup")
        self.assertTrue(partial.is_file())
        self.end()
        self.assert_saved()

    def test_same_recording_id_on_two_devices_has_independent_temporary_files(self):
        partial = self.begin()
        child = run_child(self.root, "receive")
        self.assertEqual(child["saved"], 1)
        self.assertNotEqual(partial.name, child["partial_name"])
        self.end()
        self.assert_saved()
        reloaded = CallRecordingRepository(self.root)
        self.assertIsNotNone(reloaded.find(RECORDING_ID, IMEI_A))
        self.assertIsNotNone(reloaded.find(RECORDING_ID, IMEI_B))
        self.assertEqual(len(list(self.root.rglob("*.amr"))), 2)

    def test_same_process_repository_cleanup_respects_active_guard(self):
        original = self.repository.commit

        def commit_at_boundary(path, metadata, digest):
            CallRecordingRepository(self.root)
            return original(path, metadata, digest)

        self.repository.commit = commit_at_boundary
        self.begin()
        self.end()
        self.assert_saved()

    def test_process_crash_releases_guard_and_next_start_reclaims_partial(self):
        child = run_child(self.root, "crash")
        partial = self.root / ".incoming" / child["partial_name"]
        self.assertTrue(partial.is_file())
        CallRecordingRepository(self.root)
        self.assertFalse(partial.exists())
        self.assertEqual(list((self.root / ".incoming").iterdir()), [])

    def test_fresh_legacy_partial_is_preserved_and_stale_legacy_partial_is_reclaimed(self):
        fresh = Path(self.repository.incoming_path("legacy-fresh"))
        stale = Path(self.repository.incoming_path("legacy-stale"))
        fresh.write_bytes(b"synthetic partial")
        stale.write_bytes(b"synthetic partial")
        old = time.time() - 7 * 86400
        os.utime(stale, (old, old))
        CallRecordingRepository(self.root)
        self.assertTrue(fresh.exists())
        self.assertFalse(stale.exists())

    def test_abort_from_another_thread_releases_files_and_guard(self):
        self.begin()
        worker = threading.Thread(target=self.receiver.abort, args=("serial_disconnect",))
        worker.start()
        worker.join(timeout=3)
        self.assertFalse(worker.is_alive())
        self.assertIsNone(self.receiver.current)
        self.assertEqual(list((self.root / ".incoming").iterdir()), [])

    def test_timeout_and_replacement_release_old_guard(self):
        partial = self.begin()
        self.receiver.consume_line(begin_frame("replacement", imei=IMEI_A))
        self.assertFalse(partial.exists())
        self.receiver.current.last_activity -= 31
        self.assertTrue(self.receiver.expire_stale())
        self.assertEqual(list((self.root / ".incoming").iterdir()), [])

    def test_commit_failure_releases_guard_and_reports_save_failure(self):
        self.begin()
        with patch.object(self.repository, "commit", side_effect=OSError("synthetic commit failure")):
            self.end()
        self.assertEqual(self.saved, [])
        self.assertEqual(self.aborted, ["save_failed"])
        self.assertEqual(list((self.root / ".incoming").iterdir()), [])

    def test_invalid_audio_releases_guard(self):
        invalid = b"not an AMR file"
        self.receiver.consume_line(begin_frame(RECORDING_ID, payload=invalid, imei=IMEI_A))
        self.receiver.consume_line(chunk_frame(RECORDING_ID, 1, invalid))
        self.receiver.consume_line(end_frame(RECORDING_ID, 1, payload=invalid))
        self.assertEqual(self.saved, [])
        self.assertEqual(self.aborted, ["save_failed"])
        self.assertEqual(list((self.root / ".incoming").iterdir()), [])


def child_main(root, operation):
    repository = CallRecordingRepository(root)
    if operation == "cleanup":
        return {"cleaned": True}
    saved = []
    receiver = SerialCallRecordingReceiver(repository, on_saved=saved.append)
    receiver.consume_line(begin_frame(RECORDING_ID, imei=IMEI_B))
    partial_name = Path(receiver.current.temp_path).name
    receiver.consume_line(chunk_frame(RECORDING_ID, 1, AMR_PAYLOAD))
    if operation == "crash":
        receiver.current.file.flush()
        print(json.dumps({"partial_name": partial_name}), flush=True)
        os._exit(0)
    receiver.consume_line(end_frame(RECORDING_ID, 1))
    return {"partial_name": partial_name, "saved": len(saved)}


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "--child":
        print(json.dumps(child_main(Path(sys.argv[2]), sys.argv[3])), flush=True)
        sys.exit(0)
    unittest.main()
