import asyncio
import errno
import hashlib
import json
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch

from sms_core.call_recordings import CallRecordingRepository, CloudCallRecordingUploader
from sms_core.call_recordings import SerialCallRecordingReceiver
from tests.test_call_recordings import begin_frame, chunk_frame, end_frame, AMR_PAYLOAD


IMEI = "000000000000001"
AUDIO = b"#!AMR\n" + (b"\x04" + bytes(12)) * 20
PRIVATE_TEXT = "password=synthetic-private-path"


class OpaqueDiagnosticError(Exception):
    def __str__(self):
        raise AssertionError("Diagnostic must not format untrusted exception text")

    def __repr__(self):
        raise AssertionError("Diagnostic must not format untrusted exception text")


class SerialRecordingDiagnosticSafetyTests(unittest.TestCase):
    def test_callback_failures_preserve_valid_recording_and_report_safe_types(self):
        with tempfile.TemporaryDirectory() as directory:
            repository = CallRecordingRepository(directory)
            logs = []
            def fail(*_args):
                raise OpaqueDiagnosticError()
            receiver = SerialCallRecordingReceiver(repository, on_started=fail,
                on_saved=fail, log_error=logs.append)
            self.assertTrue(receiver.consume_line(begin_frame('synthetic-callback')))
            self.assertIsNotNone(receiver.current)
            receiver.consume_line(chunk_frame('synthetic-callback', 1, AMR_PAYLOAD))
            receiver.consume_line(end_frame('synthetic-callback', 1))
            self.assertIsNone(receiver.current)
            self.assertIsNotNone(repository.find('synthetic-callback'))
            self.assertEqual(logs, [
                'Call recording start callback failed: OpaqueDiagnosticError',
                'Call recording saved callback failed: OpaqueDiagnosticError'])

    def test_abort_callback_failure_releases_file_and_allows_next_transfer(self):
        with tempfile.TemporaryDirectory() as directory:
            repository = CallRecordingRepository(directory)
            logs = []
            def fail(*_args):
                raise OpaqueDiagnosticError()
            receiver = SerialCallRecordingReceiver(repository, on_aborted=fail, log_error=logs.append)
            receiver.consume_line(begin_frame('synthetic-abort'))
            partial = Path(receiver.current.temp_path)
            self.assertTrue(receiver.abort('synthetic-aborted'))
            self.assertFalse(partial.exists())
            self.assertIsNone(receiver.current)
            self.assertIn('Call recording abort callback failed: OpaqueDiagnosticError', logs)
            receiver.consume_line(begin_frame('synthetic-next'))
            self.assertIsNotNone(receiver.current)
            receiver.abort()

    def test_callback_file_error_does_not_expose_filename(self):
        with tempfile.TemporaryDirectory() as directory:
            logs = []
            def fail(*_args):
                Path(directory, PRIVATE_TEXT).read_bytes()
            receiver = SerialCallRecordingReceiver(CallRecordingRepository(directory),
                on_started=fail, log_error=logs.append)
            receiver.consume_line(begin_frame('synthetic-path'))
            self.assertNotIn(PRIVATE_TEXT, '\n'.join(logs))
            self.assertEqual(logs, ['Call recording start callback failed: FileNotFoundError'])
            receiver.abort()


class RecordingDiagnosticSafetyTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.directory = tempfile.TemporaryDirectory(prefix="sms-recording-diagnostic-")
        self.logs = []
        self.repository = CallRecordingRepository(self.directory.name)
        partial = Path(self.repository.incoming_path("synthetic-recording"))
        partial.write_bytes(AUDIO)
        self.recording = self.repository.commit(partial, {
            "recording_id": "synthetic-recording", "imei": IMEI,
            "started_at": 1789171200, "duration_ms": 400,
        }, hashlib.sha256(AUDIO).hexdigest())
        self.uploader = CloudCallRecordingUploader(self.repository, log_error=self.logs.append)
        self.ws = object()

    async def asyncTearDown(self):
        await self.uploader.stop()
        self.directory.cleanup()

    def metadata(self):
        return json.loads(Path(self.recording.metadata_path).read_text(encoding="utf-8"))

    def arguments(self, send_payload):
        return (self.ws, send_payload, lambda: {"imei": IMEI},
                lambda ws: ws is self.ws, lambda: True)

    async def send_success(self, ws, payload):
        kind = payload["type"]
        if kind == "call_recording_offer":
            self.uploader.handle_server_message({"type": "call_recording_offer_ack",
                "recording_id": self.recording.recording_id, "ok": True,
                "accepted": True}, websocket=ws)
        elif kind == "call_recording_end":
            self.uploader.handle_server_message({"type": "call_recording_result",
                "recording_id": self.recording.recording_id, "ok": True}, websocket=ws)
        return "sent"

    async def test_cancelled_file_failure_is_joined_without_exposing_path(self):
        entered, release = threading.Event(), threading.Event()
        completed = threading.Event()
        loop = asyncio.get_running_loop()
        diagnostics = []
        previous_handler = loop.get_exception_handler()
        loop.set_exception_handler(lambda _loop, context: diagnostics.append(context))

        def failed_file_read():
            entered.set()
            try:
                release.wait(5)
                with open(Path(self.directory.name) / PRIVATE_TEXT, "rb") as stream:
                    return stream.read()
            finally:
                completed.set()

        task = asyncio.create_task(self.uploader._io(failed_file_read))
        try:
            self.assertTrue(await asyncio.to_thread(entered.wait, 5))
            task.cancel()
            await asyncio.sleep(0)
            self.assertFalse(task.done())
            release.set()
            with self.assertRaises(asyncio.CancelledError):
                await task
            self.assertTrue(completed.is_set())
            self.assertEqual(self.logs, ["Cancelled recording I/O failed: FileNotFoundError"])
            self.assertEqual(diagnostics, [])
        finally:
            release.set()
            await asyncio.gather(task, return_exceptions=True)
            loop.set_exception_handler(previous_handler)

    async def test_cancelled_opaque_file_error_never_formats_exception(self):
        entered, release = threading.Event(), threading.Event()

        def fail():
            entered.set()
            release.wait(5)
            raise OpaqueDiagnosticError()

        task = asyncio.create_task(self.uploader._io(fail))
        try:
            self.assertTrue(await asyncio.to_thread(entered.wait, 5))
            task.cancel()
            await asyncio.sleep(0)
            release.set()
            with self.assertRaises(asyncio.CancelledError):
                await task
            self.assertEqual(self.logs, ["Cancelled recording I/O failed: OpaqueDiagnosticError"])
        finally:
            release.set()
            await asyncio.gather(task, return_exceptions=True)

    async def test_upload_transport_error_retains_pending_and_cooldown_without_raw_text(self):
        async def fail(_ws, _payload):
            raise OpaqueDiagnosticError()

        await self.uploader._drain(*self.arguments(fail))
        self.assertEqual(self.metadata()["upload_status"], "pending")
        self.assertEqual(self.metadata()["upload_error"], "unexpected_upload_error")
        self.assertGreater(self.metadata()["upload_retry_at"], 0)
        self.assertEqual(self.logs, ["Upload call recording failed for synthetic-recording: OpaqueDiagnosticError"])
        self.assertIsNotNone(self.uploader._retry_handle)
        self.assertFalse(self.uploader._offer_waiters)
        self.assertTrue(Path(self.recording.path).is_file())

    async def test_upload_persistence_error_preserves_sidecar_and_hides_file_path(self):
        before = Path(self.recording.metadata_path).read_bytes()
        failure = PermissionError(errno.EACCES, "synthetic denied", PRIVATE_TEXT)
        with patch("sms_core.call_recordings.os.fsync", side_effect=failure):
            await self.uploader._drain(*self.arguments(self.send_success))
        self.assertEqual(Path(self.recording.metadata_path).read_bytes(), before)
        self.assertFalse(Path(self.recording.metadata_path + ".tmp").exists())
        self.assertTrue(Path(self.recording.path).is_file())
        self.assertTrue(self.logs)
        self.assertNotIn(PRIVATE_TEXT, "\n".join(self.logs))
        self.assertTrue(any("PermissionError" in message for message in self.logs))
        self.assertIsNotNone(self.uploader._retry_handle)

    async def test_scan_and_retry_failures_release_task_without_formatting_exception(self):
        with patch.object(self.repository, "pending_with_delays", side_effect=OpaqueDiagnosticError()):
            await self.uploader._drain(*self.arguments(self.send_success))
        self.assertEqual(self.logs, ["Upload call recording failed: OpaqueDiagnosticError"])
        self.assertIsNone(self.uploader._task)
        self.assertIsNone(self.uploader._retry_handle)
        self.logs.clear()

        async def reject(_ws, payload):
            self.uploader.handle_server_message({"type": "call_recording_offer_ack",
                "recording_id": payload["recording_id"], "ok": False,
                "accepted": False, "reason": "storage_unavailable"}, websocket=self.ws)
            return "sent"

        with patch.object(self.uploader, "_schedule_retry", side_effect=OpaqueDiagnosticError()):
            await self.uploader._drain(*self.arguments(reject))
        self.assertEqual(self.logs, ["Schedule call recording retry failed: OpaqueDiagnosticError"])
        self.assertIsNone(self.uploader._task)
        self.assertEqual(self.metadata()["upload_status"], "pending")

    async def test_schedule_submission_failure_uses_safe_error_type(self):
        loop = unittest.mock.Mock()
        loop.is_running.return_value = True
        loop.call_soon_threadsafe.side_effect = OpaqueDiagnosticError()
        self.assertFalse(self.uploader.schedule(loop, self.ws, send_payload=self.send_success,
            identity_payload=lambda: {"imei": IMEI}, is_current=lambda _ws: True,
            is_authorized=lambda: True))
        self.assertEqual(self.logs, ["Schedule call recording upload failed: OpaqueDiagnosticError"])
        self.assertIsNone(self.uploader._task)
