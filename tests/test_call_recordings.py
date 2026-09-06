import asyncio
import base64
import hashlib
import json
import os
from pathlib import Path
import tempfile
import unittest

from sms_core.call_recordings import (
    CallRecordingRepository,
    CloudCallRecordingUploader,
    MAX_CALL_RECORDING_BYTES,
    MAX_SERIAL_RECORDING_CHUNKS,
    SerialCallRecordingReceiver,
)


AMR_PAYLOAD = b"#!AMR\n" + bytes(range(96))
IMEI_A = "123456789012345"
IMEI_B = "543210987654321"


def begin_frame(recording_id, payload=AMR_PAYLOAD, phone="10086", imei=IMEI_A):
    parts = [
        "@@CALL_RECORD_BEGIN",
        recording_id,
        base64.b64encode(phone.encode("utf-8")).decode("ascii"),
        "1788048000",
        "3200",
        str(len(payload)),
        "amr",
    ]
    if imei is not None:
        parts.append(imei)
    return "|".join(parts)


def chunk_frame(recording_id, sequence, payload):
    return "|".join(
        [
            "@@CALL_RECORD_CHUNK",
            recording_id,
            str(sequence),
            base64.b64encode(payload).decode("ascii"),
        ]
    )


def end_frame(recording_id, count, payload=AMR_PAYLOAD):
    return "|".join(
        [
            "@@CALL_RECORD_END",
            recording_id,
            str(count),
            str(len(payload)),
        ]
    )


class CallRecordingRepositoryTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.repository = CallRecordingRepository(self.root / "recordings")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_serial_receiver_saves_recording_and_sidecar(self):
        saved = []
        receiver = SerialCallRecordingReceiver(
            self.repository,
            on_saved=saved.append,
        )

        self.assertTrue(receiver.consume_line(begin_frame("recording-a")))
        self.assertTrue(receiver.consume_line(chunk_frame("recording-a", 1, AMR_PAYLOAD)))
        self.assertTrue(receiver.consume_line(end_frame("recording-a", 1)))

        self.assertEqual(len(saved), 1)
        recording = saved[0]
        self.assertEqual(Path(recording.path).read_bytes(), AMR_PAYLOAD)
        metadata = json.loads(Path(recording.metadata_path).read_text(encoding="utf-8"))
        self.assertEqual(metadata["upload_status"], "pending")
        self.assertEqual(metadata["sha256"], hashlib.sha256(AMR_PAYLOAD).hexdigest())
        self.assertEqual(recording.imei, IMEI_A)
        self.assertEqual(metadata["imei"], IMEI_A)
        self.assertIn(IMEI_A, Path(recording.path).parts)

    def test_serial_receiver_notifies_start_before_save_and_abort_on_failure(self):
        started = []
        aborted = []
        receiver = SerialCallRecordingReceiver(
            self.repository,
            on_started=lambda metadata: started.append(dict(metadata)),
            on_aborted=lambda metadata, reason: aborted.append((dict(metadata), reason)),
        )

        receiver.consume_line(begin_frame("recording-status"))
        self.assertEqual(started[0]["recording_id"], "recording-status")
        self.assertEqual(started[0]["size"], len(AMR_PAYLOAD))
        self.assertEqual(aborted, [])
        receiver.consume_line(chunk_frame("recording-status", 2, AMR_PAYLOAD))

        self.assertEqual(aborted[0][0]["recording_id"], "recording-status")
        self.assertEqual(aborted[0][1], "ValueError")

    def test_serial_receiver_accepts_line_safe_chunk_frames(self):
        payload = b"#!AMR\n" + bytes(range(256)) * 4
        recording_id = "rec-" + IMEI_A + "-1788140340-84809"
        receiver = SerialCallRecordingReceiver(self.repository)
        receiver.consume_line(begin_frame(recording_id, payload=payload))
        chunk_size = 24
        chunks = [
            payload[index : index + chunk_size]
            for index in range(0, len(payload), chunk_size)
        ]
        for sequence, chunk in enumerate(chunks, 1):
            frame = chunk_frame(recording_id, sequence, chunk)
            self.assertLessEqual(len(frame), 127)
            receiver.consume_line(frame)
        receiver.consume_line(end_frame(recording_id, len(chunks), payload=payload))

        recording = self.repository.find(recording_id, IMEI_A)
        self.assertIsNotNone(recording)
        self.assertEqual(Path(recording.path).read_bytes(), payload)

    def test_line_safe_chunk_frame_holds_for_max_id_and_sequence(self):
        recording_id = "X" * 64
        frame = chunk_frame(recording_id, MAX_SERIAL_RECORDING_CHUNKS, b"x" * 24)
        self.assertLessEqual(len(frame), 127)

    def test_serial_receiver_accepts_recordings_above_legacy_limit_but_rejects_over_max(self):
        receiver = SerialCallRecordingReceiver(self.repository)
        accepted_payload = b"#!AMR\n" + b"x" * (300 * 1024 - 6)
        receiver.consume_line(begin_frame("recording-512k", payload=accepted_payload))
        self.assertIsNotNone(receiver.current)
        self.assertEqual(MAX_CALL_RECORDING_BYTES, 512 * 1024)

        receiver.abort()
        oversized_payload = b"#!AMR\n" + b"x" * (MAX_CALL_RECORDING_BYTES - 5)
        receiver.consume_line(begin_frame("recording-oversized", payload=oversized_payload))
        self.assertIsNone(receiver.current)

    def test_legacy_begin_frame_uses_runtime_or_recording_id_imei(self):
        runtime_receiver = SerialCallRecordingReceiver(
            self.repository,
            source_imei=lambda: IMEI_A,
        )
        runtime_receiver.consume_line(begin_frame("legacy-runtime", imei=None))
        runtime_receiver.consume_line(chunk_frame("legacy-runtime", 1, AMR_PAYLOAD))
        runtime_receiver.consume_line(end_frame("legacy-runtime", 1))

        recording_id = "rec-{}-1788048000-1".format(IMEI_B)
        id_receiver = SerialCallRecordingReceiver(self.repository)
        id_receiver.consume_line(begin_frame(recording_id, imei=None))
        id_receiver.consume_line(chunk_frame(recording_id, 1, AMR_PAYLOAD))
        id_receiver.consume_line(end_frame(recording_id, 1))

        self.assertIsNotNone(self.repository.find("legacy-runtime", IMEI_A))
        self.assertIsNotNone(self.repository.find(recording_id, IMEI_B))

    def test_pending_recordings_are_scoped_by_source_imei(self):
        for recording_id, imei in (("recording-a", IMEI_A), ("recording-b", IMEI_B)):
            receiver = SerialCallRecordingReceiver(self.repository)
            receiver.consume_line(begin_frame(recording_id, imei=imei))
            receiver.consume_line(chunk_frame(recording_id, 1, AMR_PAYLOAD))
            receiver.consume_line(end_frame(recording_id, 1))

        self.assertEqual(
            [recording.recording_id for recording in self.repository.pending(IMEI_A)],
            ["recording-a"],
        )
        self.assertEqual(
            [recording.recording_id for recording in self.repository.pending(IMEI_B)],
            ["recording-b"],
        )

    def test_invalid_chunk_aborts_and_removes_partial_file(self):
        receiver = SerialCallRecordingReceiver(self.repository)
        receiver.consume_line(begin_frame("recording-invalid"))
        partial = Path(self.repository.incoming_path("recording-invalid"))
        self.assertTrue(partial.exists())

        receiver.consume_line(chunk_frame("recording-invalid", 2, AMR_PAYLOAD))

        self.assertIsNone(receiver.current)
        self.assertFalse(partial.exists())

    def test_duplicate_recording_is_not_saved_twice(self):
        saved = []
        receiver = SerialCallRecordingReceiver(
            self.repository,
            on_saved=saved.append,
        )
        frames = [
            begin_frame("recording-duplicate"),
            chunk_frame("recording-duplicate", 1, AMR_PAYLOAD),
            end_frame("recording-duplicate", 1),
        ]
        for frame in frames:
            receiver.consume_line(frame)
        for frame in frames:
            receiver.consume_line(frame)

        self.assertEqual(len(saved), 1)
        self.assertEqual(len(self.repository.pending()), 1)

    def test_repository_rejects_non_amr_file(self):
        partial = Path(self.repository.incoming_path("invalid-amr"))
        partial.write_bytes(b"not-amr")
        with self.assertRaisesRegex(ValueError, "format"):
            self.repository.commit(
                str(partial),
                {
                    "recording_id": "invalid-amr",
                    "phone": "10086",
                    "started_at": 1788048000,
                },
                "",
            )

    def test_repository_removes_stale_partial_files_on_startup(self):
        partial = Path(self.repository.incoming_path("stale-recording"))
        partial.write_bytes(b"partial")

        reloaded = CallRecordingRepository(self.root / "recordings")

        self.assertFalse(partial.exists())
        self.assertEqual(reloaded.cleanup_incoming(), 0)

    def test_receiver_abort_and_timeout_remove_partial_file(self):
        now = [10.0]
        receiver = SerialCallRecordingReceiver(
            self.repository,
            monotonic=lambda: now[0],
        )
        receiver.consume_line(begin_frame("recording-abort"))
        abort_path = Path(self.repository.incoming_path("recording-abort"))
        self.assertTrue(abort_path.exists())
        self.assertTrue(receiver.abort("serial_disconnect"))
        self.assertFalse(abort_path.exists())

        receiver.consume_line(begin_frame("recording-timeout"))
        timeout_path = Path(self.repository.incoming_path("recording-timeout"))
        now[0] += 31.0
        self.assertTrue(receiver.expire_stale())
        self.assertFalse(timeout_path.exists())


class CloudCallRecordingUploaderTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.repository = CallRecordingRepository(self.root / "recordings")
        partial = Path(self.repository.incoming_path("cloud-recording"))
        partial.write_bytes(AMR_PAYLOAD)
        self.recording = self.repository.commit(
            str(partial),
            {
                "recording_id": "cloud-recording",
                "phone": "10086",
                "started_at": 1788048000,
                "duration_ms": 3200,
                "imei": IMEI_A,
            },
            hashlib.sha256(AMR_PAYLOAD).hexdigest(),
        )

    async def asyncTearDown(self):
        self.temp_dir.cleanup()

    async def test_upload_success_marks_sidecar_uploaded(self):
        uploader = CloudCallRecordingUploader(self.repository)
        ws = object()

        async def send_payload(_ws, payload):
            if payload["type"] == "call_recording_offer":
                uploader.handle_server_message(
                    {
                        "type": "call_recording_offer_ack",
                        "ok": True,
                        "accepted": True,
                        "recording_id": payload["recording_id"],
                    }
                )
            elif payload["type"] == "call_recording_end":
                uploader.handle_server_message(
                    {
                        "type": "call_recording_result",
                        "ok": True,
                        "recording_id": payload["recording_id"],
                    }
                )
            return True

        ok, _reason = await uploader._send_recording(
            self.recording,
            ws,
            send_payload,
            lambda: {"imei": IMEI_A},
        )
        self.assertTrue(ok)
        self.repository.mark_uploaded(self.recording)
        metadata = json.loads(Path(self.recording.metadata_path).read_text(encoding="utf-8"))
        self.assertEqual(metadata["upload_status"], "uploaded")

    async def test_delayed_ack_after_poll_timeout_still_completes_upload(self):
        uploader = CloudCallRecordingUploader(self.repository)
        ws = object()
        delayed_tasks = []

        async def send_payload(_ws, payload):
            if payload["type"] == "call_recording_offer":
                async def send_delayed_offer_ack():
                    await asyncio.sleep(0.65)
                    uploader.handle_server_message(
                        {
                            "type": "call_recording_offer_ack",
                            "ok": True,
                            "accepted": True,
                            "recording_id": payload["recording_id"],
                        },
                        websocket=ws,
                    )

                delayed_tasks.append(asyncio.create_task(send_delayed_offer_ack()))
            elif payload["type"] == "call_recording_end":
                async def send_delayed_result():
                    await asyncio.sleep(0.65)
                    uploader.handle_server_message(
                        {
                            "type": "call_recording_result",
                            "ok": True,
                            "recording_id": payload["recording_id"],
                        },
                        websocket=ws,
                    )

                delayed_tasks.append(asyncio.create_task(send_delayed_result()))
            return True

        try:
            ok, reason = await uploader._send_recording(
                self.recording,
                ws,
                send_payload,
                lambda: {"imei": IMEI_A},
                is_current=lambda current_ws: current_ws is ws,
            )
        finally:
            if delayed_tasks:
                await asyncio.gather(*delayed_tasks)

        self.assertTrue(ok, reason)

    async def test_new_connection_schedule_runs_after_old_upload_finishes(self):
        uploader = CloudCallRecordingUploader(self.repository)
        loop = asyncio.get_running_loop()
        old_ws = object()
        new_ws = object()
        current_ws = [old_ws]
        old_offer_started = asyncio.Event()
        release_old_offer = asyncio.Event()
        new_upload_finished = asyncio.Event()

        async def old_send(_ws, payload):
            if payload["type"] == "call_recording_offer":
                old_offer_started.set()
                await release_old_offer.wait()
                return False
            return False

        async def new_send(_ws, payload):
            if payload["type"] == "call_recording_offer":
                uploader.handle_server_message(
                    {
                        "type": "call_recording_offer_ack",
                        "ok": True,
                        "accepted": True,
                        "recording_id": payload["recording_id"],
                    }
                )
            elif payload["type"] == "call_recording_end":
                uploader.handle_server_message(
                    {
                        "type": "call_recording_result",
                        "ok": True,
                        "recording_id": payload["recording_id"],
                    }
                )
                new_upload_finished.set()
            return True

        is_current = lambda ws: ws is current_ws[0]
        is_authorized = lambda: True
        self.assertTrue(
            uploader.schedule(
                loop,
                old_ws,
                send_payload=old_send,
                identity_payload=lambda: {"imei": IMEI_A},
                is_current=is_current,
                is_authorized=is_authorized,
            )
        )
        await asyncio.wait_for(old_offer_started.wait(), 1)
        current_ws[0] = new_ws
        self.assertTrue(
            uploader.schedule(
                loop,
                new_ws,
                send_payload=new_send,
                identity_payload=lambda: {"imei": IMEI_A},
                is_current=is_current,
                is_authorized=is_authorized,
            )
        )
        release_old_offer.set()
        await asyncio.wait_for(new_upload_finished.wait(), 2)

        for _ in range(20):
            metadata = json.loads(
                Path(self.recording.metadata_path).read_text(encoding="utf-8")
            )
            if metadata.get("upload_status") == "uploaded":
                break
            await asyncio.sleep(0.01)
        self.assertEqual(metadata["upload_status"], "uploaded")

    async def test_uploader_rejects_recording_from_another_device(self):
        uploader = CloudCallRecordingUploader(self.repository)
        sent = []

        async def send_payload(_ws, payload):
            sent.append(payload)
            return True

        ok, reason = await uploader._send_recording(
            self.recording,
            object(),
            send_payload,
            lambda: {"imei": IMEI_B},
        )

        self.assertFalse(ok)
        self.assertEqual(reason, "source_imei_mismatch")
        self.assertEqual(sent, [])

    async def test_late_response_from_old_connection_does_not_complete_new_waiter(self):
        uploader = CloudCallRecordingUploader(self.repository)
        old_ws = object()
        new_ws = object()
        loop = asyncio.get_running_loop()
        future = loop.create_future()
        uploader._offer_waiters[self.recording.recording_id] = (new_ws, future)

        self.assertTrue(
            uploader.handle_server_message(
                {
                    "type": "call_recording_offer_ack",
                    "ok": True,
                    "accepted": True,
                    "recording_id": self.recording.recording_id,
                },
                websocket=old_ws,
            )
        )
        self.assertFalse(future.done())
        self.assertIn(self.recording.recording_id, uploader._offer_waiters)

        uploader.handle_server_message(
            {
                "type": "call_recording_offer_ack",
                "ok": True,
                "accepted": True,
                "recording_id": self.recording.recording_id,
            },
            websocket=new_ws,
        )
        self.assertTrue(future.done())

    async def test_unexpected_upload_exception_returns_recording_to_pending(self):
        errors = []
        uploader = CloudCallRecordingUploader(
            self.repository,
            log_error=errors.append,
        )

        async def send_payload(_ws, _payload):
            raise RuntimeError("transport exploded")

        await uploader._drain(
            object(),
            send_payload,
            lambda: {"imei": IMEI_A},
            lambda _ws: True,
            lambda: True,
        )

        metadata = json.loads(
            Path(self.recording.metadata_path).read_text(encoding="utf-8")
        )
        self.assertEqual(metadata["upload_status"], "pending")
        self.assertEqual(metadata["upload_error"], "unexpected_upload_error")
        self.assertTrue(any("transport exploded" in item for item in errors))

    async def test_uploaded_state_exception_returns_recording_to_pending(self):
        errors = []
        uploader = CloudCallRecordingUploader(
            self.repository,
            log_error=errors.append,
        )

        async def send_payload(_ws, payload):
            if payload["type"] == "call_recording_offer":
                uploader.handle_server_message(
                    {
                        "type": "call_recording_offer_ack",
                        "ok": True,
                        "accepted": True,
                        "recording_id": payload["recording_id"],
                    }
                )
            elif payload["type"] == "call_recording_end":
                uploader.handle_server_message(
                    {
                        "type": "call_recording_result",
                        "ok": True,
                        "recording_id": payload["recording_id"],
                    }
                )
            return True

        def fail_mark_uploaded(_recording):
            raise OSError("sidecar is temporarily unavailable")

        self.repository.mark_uploaded = fail_mark_uploaded
        await uploader._drain(
            object(),
            send_payload,
            lambda: {"imei": IMEI_A},
            lambda _ws: True,
            lambda: True,
        )

        metadata = json.loads(
            Path(self.recording.metadata_path).read_text(encoding="utf-8")
        )
        self.assertEqual(metadata["upload_status"], "pending")
        self.assertEqual(metadata["upload_error"], "upload_state_persist_error")
        self.assertTrue(any("sidecar is temporarily unavailable" in item for item in errors))


if __name__ == "__main__":
    unittest.main()
