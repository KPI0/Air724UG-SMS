"""Recording direction survives serial reception, restart and cloud upload."""
import json
from pathlib import Path
import tempfile
import unittest

from sms_core.call_recordings import (
    CallRecordingRepository,
    CloudCallRecordingUploader,
    SerialCallRecordingReceiver,
)
from sms_core.cloud_payloads import build_call_recording_status_payload
from tests.test_call_recordings import AMR_PAYLOAD, IMEI_A, begin_frame, chunk_frame, end_frame


class CallRecordingDirectionTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "recordings"
        self.repository = CallRecordingRepository(self.path)

    def receive(self, receiver, recording_id, *, imei=IMEI_A):
        for frame in (
            begin_frame(recording_id, imei=imei),
            chunk_frame(recording_id, 1, AMR_PAYLOAD),
            end_frame(recording_id, 1),
        ):
            self.assertTrue(receiver.consume_line(frame))
        recording = self.repository.find(recording_id, IMEI_A)
        self.assertIsNotNone(recording)
        self.assertEqual(Path(recording.path).read_bytes(), AMR_PAYLOAD)
        return recording

    async def test_direction_survives_serial_status_restart_and_offer(self):
        for direction in ("incoming", "outgoing"):
            with self.subTest(direction=direction):
                started = []
                recording_id = "direction-" + direction
                receiver = SerialCallRecordingReceiver(self.repository, on_started=started.append)
                receiver.consume_line("@@CALL_RECORD_META|" + recording_id + "|" + direction)
                recording = self.receive(receiver, recording_id)
                self.assertEqual(started[0].get("direction"), direction)
                metadata = json.loads(Path(recording.metadata_path).read_text(encoding="utf-8"))
                self.assertEqual(metadata.get("direction"), direction)
                status = build_call_recording_status_payload(
                    "uploading", recording_id, "10086", 1788048000, 3200,
                    len(AMR_PAYLOAD), 1788048001, {"imei": IMEI_A}, direction=direction,
                )
                self.assertEqual(status.get("direction"), direction)
                repository = CallRecordingRepository(self.path)
                reloaded = repository.find(recording_id, IMEI_A)
                self.assertEqual(reloaded.direction, direction)
                uploader = CloudCallRecordingUploader(repository)
                ws, offers = object(), []

                async def send_payload(websocket, payload):
                    offers.append(payload)
                    uploader.handle_server_message({
                        "type": "call_recording_offer_ack", "ok": True,
                        "accepted": True, "already_uploaded": True,
                        "recording_id": recording_id,
                    }, websocket=websocket)
                    return True

                ok, reason = await uploader._send_recording(
                    reloaded, ws, send_payload, lambda: {"imei": IMEI_A},
                )
                self.assertTrue(ok, reason)
                self.assertEqual(offers[0].get("direction"), direction)

    def test_old_begin_formats_do_not_invent_direction(self):
        for index, imei in enumerate((None, IMEI_A)):
            receiver = SerialCallRecordingReceiver(self.repository, source_imei=lambda: IMEI_A)
            saved = self.receive(receiver, "legacy-" + str(index), imei=imei)
            self.assertEqual(saved.direction, "")
            sidecar = Path(saved.metadata_path)
            metadata = json.loads(sidecar.read_text(encoding="utf-8"))
            metadata.pop("direction", None)
            sidecar.write_text(json.dumps(metadata), encoding="utf-8")
            reopened = CallRecordingRepository(self.path).find(saved.recording_id, IMEI_A)
            self.assertEqual(reopened.direction, "")

    def test_direction_prelude_only_applies_to_the_next_matching_begin(self):
        for index, prelude in enumerate((
            ["@@CALL_RECORD_META|other-recording|outgoing"],
            ["@@CALL_RECORD_META|matching|invalid"],
            ["@@CALL_RECORD_META|matching|outgoing", "@@CALL_RECORD_ABORT|matching"],
        )):
            with self.subTest(index=index):
                receiver = SerialCallRecordingReceiver(self.repository)
                recording_id = "matching-" + str(index)
                for line in prelude:
                    receiver.consume_line(line.replace("|matching", "|" + recording_id))
                saved = self.receive(receiver, recording_id)
                self.assertEqual(saved.direction, "")

    def test_interleaved_ordinary_log_does_not_discard_direction(self):
        receiver = SerialCallRecordingReceiver(self.repository)
        receiver.consume_line("@@CALL_RECORD_META|interleaved|outgoing")
        self.assertFalse(receiver.consume_line("unrelated serial log"))
        saved = self.receive(receiver, "interleaved")
        self.assertEqual(saved.direction, "outgoing")

    def test_expired_prelude_does_not_affect_later_recording(self):
        now = [10.0]
        receiver = SerialCallRecordingReceiver(self.repository, monotonic=lambda: now[0])
        receiver.consume_line("@@CALL_RECORD_META|expired|outgoing")
        now[0] += 31.0
        saved = self.receive(receiver, "expired")
        self.assertEqual(saved.direction, "")

    def test_disconnect_discards_unconsumed_direction_prelude(self):
        receiver = SerialCallRecordingReceiver(self.repository)
        receiver.consume_line("@@CALL_RECORD_META|after-disconnect|outgoing")
        receiver.abort("serial_disconnect")
        saved = self.receive(receiver, "after-disconnect")
        self.assertEqual(saved.direction, "")

    def test_upload_failure_carries_the_recordings_own_direction(self):
        aborted = []
        receiver = SerialCallRecordingReceiver(
            self.repository, on_aborted=lambda metadata, _reason: aborted.append(metadata),
        )
        receiver.consume_line("@@CALL_RECORD_META|failed-outgoing|outgoing")
        receiver.consume_line(begin_frame("failed-outgoing"))
        receiver.abort("serial_disconnect")
        self.assertEqual(aborted[0].get("direction"), "outgoing")
