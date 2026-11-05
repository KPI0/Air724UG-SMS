import asyncio
import hashlib
import json
from pathlib import Path
import tempfile
import threading
import unittest

from sms_core.call_recordings import CallRecordingRepository, CloudCallRecordingUploader
from sms_core.cloud_command_context import set_cloud_command_authorized
from sms_core.cloud_message_namespace_runtime import (
    handle_cloud_message_namespace_runtime,
    schedule_cloud_call_recording_upload_namespace_runtime,
)
from sms_core.cloud_sms_event_runtime import CloudSmsEventDrainState


class CallRecordingAuthorizationTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.repository = CallRecordingRepository(Path(self.directory.name) / "recordings")
        audio = b"#!AMR\n" + bytes(range(96))
        part = Path(self.directory.name) / "synthetic.part"
        part.write_bytes(audio)
        self.imei = "123456789012345"
        self.recording = self.repository.commit(str(part), {
            "recording_id": "synthetic-authorization", "phone": "",
            "started_at": 1788048000, "duration_ms": 1000, "imei": self.imei,
        }, hashlib.sha256(audio).hexdigest())
        self.ws = object()
        self.reached_read = threading.Event()
        self.release_read = threading.Event()
        reached, release = self.reached_read, self.release_read

        class PausedUploader(CloudCallRecordingUploader):
            @staticmethod
            def _read_audio(recording):
                reached.set()
                if not release.wait(5):
                    raise OSError("Synthetic recording read timed out")
                return CloudCallRecordingUploader._read_audio(recording)

        self.uploader = PausedUploader(self.repository)
        self.sent = []
        self.namespace = {
            "CALL_RECORDING_UPLOADER": self.uploader,
            "CLOUD_SMS_EVENT_DRAIN_STATE": CloudSmsEventDrainState(),
            "cloud_ws_loop": asyncio.get_running_loop(), "cloud_ws_conn": self.ws,
            "cloud_connected": True, "cloud_device_authorized": True,
            "CLOUD_CONTROL_ENABLED": True, "CLOUD_DEVICE_SECRET": "synthetic-ascii",
            "CLOUD_WS_URL": "ws://example.test/ws/device", "_cloud_runtime_imei": lambda: self.imei,
            "_cloud_identity_payload": lambda: {"imei": self.imei},
            "_cloud_log": lambda *_args, **_kwargs: None,
            "_cloud_reply": lambda *_args: None,
            "set_cloud_auth_status_from_ack": lambda *_args: None,
            "set_cloud_status": lambda *_args: None,
            "_cloud_check_replay_window": lambda *_args, **_kwargs: None,
            "_cloud_auth_matches": lambda *_args: True,
            "_cloud_now_ts": lambda: 0, "_cloud_send_status_payload": lambda: {},
            "show_window": lambda: None, "hide_window": lambda: None,
            "_cloud_send_payload": self.send_payload,
        }

    async def asyncTearDown(self):
        self.release_read.set()
        await self.uploader.stop()
        self.directory.cleanup()

    async def send_payload(self, ws, payload):
        self.assertIs(ws, self.ws)
        self.sent.append(payload["type"])
        if payload["type"] == "call_recording_offer":
            self.ack_offer()
        elif payload["type"] == "call_recording_end":
            self.uploader.handle_server_message({"type": "call_recording_result", "ok": True,
                "recording_id": self.recording.recording_id}, websocket=ws)
        return "sent"

    def ack_offer(self, **extra):
        self.uploader.handle_server_message({"type": "call_recording_offer_ack", "ok": True,
            "accepted": True, "recording_id": self.recording.recording_id, **extra}, websocket=self.ws)

    async def authorization_ack(self, status):
        await handle_cloud_message_namespace_runtime(self.namespace, self.ws,
            json.dumps({"type": "device_login_ack", "imei": self.imei,
                "auth_status": status, "ok": status == "authorized"}))

    async def wait_for(self, predicate):
        async def wait():
            while not predicate():
                await asyncio.sleep(0.005)
        await asyncio.wait_for(wait(), 5)

    def metadata(self):
        with self.repository.lock:
            return json.loads(Path(self.recording.metadata_path).read_text(encoding="utf-8"))

    async def test_reauthorization_requires_new_offer_after_blocked_file_read(self):
        self.assertTrue(schedule_cloud_call_recording_upload_namespace_runtime(self.namespace))
        await self.wait_for(self.reached_read.is_set)
        await self.authorization_ack("waiting")
        await self.authorization_ack("authorized")
        self.release_read.set()
        await self.wait_for(lambda: self.uploader._task is None)
        self.assertEqual(self.sent, ["call_recording_offer"])
        self.assertEqual(self.metadata()["upload_status"], "pending")
        self.assertEqual(self.metadata()["upload_retry_at"], 0)
        self.assertFalse(self.uploader._offer_waiters)
        self.assertFalse(self.uploader._result_waiters)
        self.assertTrue(schedule_cloud_call_recording_upload_namespace_runtime(self.namespace))
        await self.wait_for(lambda: self.metadata()["upload_status"] == "uploaded")
        self.assertEqual(self.sent, ["call_recording_offer", "call_recording_offer",
            "call_recording_chunk", "call_recording_end"])

    async def test_direct_upload_is_cancelled_without_sending_chunks_after_pending_is_cancelled(self):
        task = asyncio.create_task(self.uploader._send_recording(
            self.recording, self.ws, self.send_payload, lambda: {"imei": self.imei}))
        await self.wait_for(self.reached_read.is_set)
        self.uploader.cancel_pending()
        self.release_read.set()
        ok, reason = await asyncio.wait_for(task, 5)
        self.assertFalse(ok)
        self.assertEqual(reason, "connection_changed")
        self.assertEqual(self.sent, ["call_recording_offer"])

    async def test_direct_upload_with_stale_connection_does_not_send_offer(self):
        ok, reason = await self.uploader._send_recording(
            self.recording, self.ws, self.send_payload, lambda: {"imei": self.imei},
            is_current=lambda _ws: False)
        self.assertFalse(ok)
        self.assertEqual(reason, "connection_changed")
        self.assertEqual(self.sent, [])
        self.assertFalse(self.uploader._offer_waiters)

    async def test_cancel_pending_invalidates_active_upload_even_when_authorized_stays_true(self):
        task = asyncio.create_task(self.uploader._drain(self.ws, self.send_payload,
            lambda: {"imei": self.imei}, lambda ws: ws is self.ws, lambda: True))
        await self.wait_for(self.reached_read.is_set)
        self.uploader.cancel_pending()
        self.release_read.set()
        await asyncio.wait_for(task, 5)
        self.assertEqual(self.sent, ["call_recording_offer"])
        self.assertEqual(self.metadata()["upload_status"], "pending")
        self.assertEqual(self.metadata()["upload_retry_at"], 0)

    async def test_cancel_pending_invalidates_created_upload_before_drain_starts(self):
        created = []
        self.release_read.set()
        self.assertTrue(schedule_cloud_call_recording_upload_namespace_runtime(self.namespace))

        def cancel_before_start():
            created.append(self.uploader._task)
            self.assertIsNotNone(created[0])
            self.assertEqual(self.sent, [])
            self.uploader.cancel_pending()

        asyncio.get_running_loop().call_soon(cancel_before_start)
        await self.wait_for(lambda: bool(created) and created[0].done())
        self.assertEqual(self.sent, [])
        self.assertFalse(self.reached_read.is_set())
        self.assertEqual(self.metadata()["upload_status"], "pending")
        self.assertEqual(self.metadata()["upload_retry_at"], 0)
        self.assertFalse(self.uploader._retry_deadlines)
        self.assertIsNone(self.uploader._retry_handle)
        self.assertIsNone(self.uploader._task)
        self.assertFalse(self.uploader._offer_waiters)
        self.assertFalse(self.uploader._result_waiters)

    async def test_fresh_schedule_survives_cancelled_created_upload_cleanup(self):
        created = []
        self.assertTrue(schedule_cloud_call_recording_upload_namespace_runtime(self.namespace))

        def cancel_and_reschedule():
            created.append(self.uploader._task)
            self.assertIsNotNone(created[0])
            self.assertEqual(self.sent, [])
            self.uploader.cancel_pending()
            self.assertTrue(schedule_cloud_call_recording_upload_namespace_runtime(self.namespace))

        asyncio.get_running_loop().call_soon(cancel_and_reschedule)
        await self.wait_for(self.reached_read.is_set)
        self.assertTrue(created[0].done())
        fresh_task = self.uploader._task
        self.assertIsNotNone(fresh_task)
        self.assertIsNot(fresh_task, created[0])
        await asyncio.wait_for(asyncio.shield(created[0]), 5)
        self.assertIs(self.uploader._task, fresh_task)
        self.assertEqual(self.sent, ["call_recording_offer"])
        self.assertEqual(self.metadata()["upload_retry_at"], 0)
        self.release_read.set()
        await self.wait_for(lambda: self.uploader._task is None)
        self.assertEqual(self.metadata()["upload_status"], "uploaded")
        self.assertEqual(self.sent, ["call_recording_offer", "call_recording_chunk", "call_recording_end"])
        self.assertFalse(self.uploader._retry_deadlines)
        self.assertTrue(Path(self.recording.path).is_file())

    async def test_old_already_uploaded_ack_cannot_commit_after_authorization_changes(self):
        async def delayed_ack(_ws, payload):
            self.sent.append(payload["type"])

            def revoke_and_ack():
                set_cloud_command_authorized(self.namespace, False)
                set_cloud_command_authorized(self.namespace, True)
                self.ack_offer(already_uploaded=True)

            asyncio.get_running_loop().call_soon(revoke_and_ack)
            return "sent"

        self.namespace["_cloud_send_payload"] = delayed_ack
        self.assertTrue(schedule_cloud_call_recording_upload_namespace_runtime(self.namespace))
        await self.wait_for(lambda: bool(self.sent) and self.uploader._task is None)
        self.assertEqual(self.sent, ["call_recording_offer"])
        self.assertEqual(self.metadata()["upload_status"], "pending")
        self.assertFalse(self.reached_read.is_set())

    async def test_new_authorized_schedule_waits_for_old_upload_without_losing_pending_file(self):
        self.assertTrue(schedule_cloud_call_recording_upload_namespace_runtime(self.namespace))
        await self.wait_for(self.reached_read.is_set)
        await self.authorization_ack("waiting")
        await self.authorization_ack("authorized")
        self.assertTrue(schedule_cloud_call_recording_upload_namespace_runtime(self.namespace))
        await asyncio.sleep(0)
        self.release_read.set()
        await self.wait_for(lambda: self.metadata()["upload_status"] == "uploaded")
        self.assertEqual(self.sent, ["call_recording_offer", "call_recording_offer",
            "call_recording_chunk", "call_recording_end"])
        self.assertTrue(Path(self.recording.path).is_file())
