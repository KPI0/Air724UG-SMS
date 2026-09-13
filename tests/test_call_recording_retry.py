"""Recording retries must survive idle sockets without response-driven loops."""

import asyncio
import hashlib
import json
from pathlib import Path
import tempfile
import threading
import time
import unittest
from unittest.mock import patch

from sms_core.call_recordings import CallRecordingRepository, CloudCallRecordingUploader
from sms_core.cloud_ws_runtime import cloud_ws_main_runtime


IMEI = "000000000000001"
AUDIO = b"#!AMR\n" + (b"\x04" + bytes(12)) * 20
RETRY_SECONDS = 0.20


def create_recording(repository, recording_id="synthetic-recording", imei=IMEI):
    partial = Path(repository.incoming_path(recording_id))
    partial.write_bytes(AUDIO)
    return repository.commit(partial, {
        "recording_id": recording_id, "imei": imei, "started_at": 1789171200,
        "duration_ms": 400,
    }, hashlib.sha256(AUDIO).hexdigest())


async def finish_uploader(uploader):
    stop = getattr(uploader, "stop", None)
    if stop is not None:
        await stop()
    elif uploader._task is not None:
        uploader._next_schedule = None
        uploader._task.cancel()
        await asyncio.gather(uploader._task, return_exceptions=True)


async def replay_receive_loop(mode):
    """Keep the real receive loop and scheduler; only the transport is local."""
    with tempfile.TemporaryDirectory(prefix="sms-recording-retry-") as temp:
        repository = CallRecordingRepository(temp)
        recording = create_recording(repository)
        uploader = CloudCallRecordingUploader(repository)
        loop = asyncio.get_running_loop()
        stop = threading.Event()
        state = {"ws": None, "authorized": False}
        offers = []

        class QueueWebSocket:
            def __init__(self):
                self.queue = asyncio.Queue()

            async def __aenter__(self):
                return self

            async def __aexit__(self, *_args):
                return False

            async def recv(self):
                return await self.queue.get()

            async def close(self):
                return None

        ws = QueueWebSocket()

        def set_state(connection, *, connected, authorized):
            state.update(ws=connection, authorized=bool(connected and authorized))

        async def login(connection):
            set_state(connection, connected=True, authorized=True)
            return True

        async def register(_ws):
            pass

        async def send_payload(connection, payload):
            if connection is not ws:
                raise AssertionError("Unexpected recording connection")
            response = {"recording_id": recording.recording_id}
            if payload["type"] == "call_recording_offer":
                offers.append(loop.time())
                await asyncio.sleep(0.003)
                response.update(type="call_recording_offer_ack", ok=True, accepted=True)
                if mode == "already_uploaded":
                    response["already_uploaded"] = True
                elif mode == "recording_deleted" or (
                    len(offers) == 1 and mode not in ("success", "result_failure")
                ):
                    response.update(ok=False, accepted=False, reason=mode)
                await ws.queue.put(response)
            elif payload["type"] == "call_recording_end":
                response.update(type="call_recording_result", ok=True)
                if mode == "result_failure" and len(offers) == 1:
                    response.update(ok=False, reason="event_persistence_failed")
                await ws.queue.put(response)
            return True

        def finish():
            stop.set()
            ws.queue.put_nowait({"type": "test_stop"})

        async def handle_message(connection, payload):
            uploader.handle_server_message(payload, websocket=connection)
            if mode in ("already_uploaded", "recording_deleted") or (
                payload.get("type") == "call_recording_result" and payload.get("ok") is True
            ):
                loop.call_later(0.01, finish)

        def schedule():
            uploader.schedule(
                loop, state["ws"], send_payload=send_payload,
                identity_payload=lambda: {"imei": IMEI},
                is_current=lambda candidate: candidate is state["ws"],
                is_authorized=lambda: state["authorized"],
            )

        watchdog = loop.call_later(2, finish)
        try:
            await asyncio.wait_for(cloud_ws_main_runtime(
                "ws://test.invalid/ws/device", 1, stop_event=stop,
                runtime_imei=lambda: IMEI, request_cloud_device_imei=lambda: None,
                set_cloud_status=lambda *_args: None, log=lambda *_args, **_kwargs: None,
                connect=lambda *_args, **_kwargs: ws, set_connection_state=set_state,
                reset_serial_log_state=lambda: None, send_register=register,
                wait_login_ack=login, handle_message=handle_message,
                cloud_control_enabled=lambda: True, monotonic=time.monotonic,
                schedule_pending_sms_events=schedule,
            ), 3)
            metadata = json.loads(Path(recording.metadata_path).read_text(encoding="utf-8"))
            return offers, metadata["upload_status"], Path(recording.path).read_bytes()
        finally:
            watchdog.cancel()
            await finish_uploader(uploader)


class RecordingReceiveLoopRetryTests(unittest.IsolatedAsyncioTestCase):
    async def test_temporary_rejections_and_result_failures_retry_after_delay_on_idle_socket(self):
        with patch("sms_core.call_recordings.CLOUD_RECORDING_RETRY_SECONDS", RETRY_SECONDS):
            for mode in ("too_many_uploads", "storage_unavailable", "event_persistence_failed", "result_failure"):
                with self.subTest(mode=mode):
                    offers, status, audio = await replay_receive_loop(mode)
                    self.assertEqual(len(offers), 2)
                    self.assertGreaterEqual(offers[1] - offers[0], RETRY_SECONDS - 0.01)
                    self.assertEqual(status, "uploaded")
                    self.assertEqual(audio, AUDIO)

    async def test_success_and_terminal_responses_do_not_retry(self):
        with patch("sms_core.call_recordings.CLOUD_RECORDING_RETRY_SECONDS", RETRY_SECONDS):
            for mode in ("success", "already_uploaded", "recording_deleted"):
                with self.subTest(mode=mode):
                    offers, status, audio = await replay_receive_loop(mode)
                    self.assertEqual(len(offers), 1)
                    self.assertEqual(status, "uploaded")
                    self.assertEqual(audio, AUDIO)


class RecordingRetrySchedulingTests(unittest.IsolatedAsyncioTestCase):
    async def asyncSetUp(self):
        self.temp = tempfile.TemporaryDirectory(prefix="sms-recording-schedule-")
        self.repository = CallRecordingRepository(self.temp.name)
        self.recording = create_recording(self.repository)
        self.uploader = CloudCallRecordingUploader(self.repository)
        self.ws = object()
        self.current_ws = self.ws
        self.authorized = True
        self.imei = IMEI
        self.offers = []
        self.completed = asyncio.Event()
        self.retry_patch = patch("sms_core.call_recordings.CLOUD_RECORDING_RETRY_SECONDS", RETRY_SECONDS)
        self.retry_patch.start()

    async def asyncTearDown(self):
        await finish_uploader(self.uploader)
        self.retry_patch.stop()
        self.temp.cleanup()

    def args(self, send_payload=None):
        return dict(
            send_payload=send_payload or self.send_success,
            identity_payload=lambda: {"imei": self.imei},
            is_current=lambda ws: ws is self.current_ws,
            is_authorized=lambda: self.authorized,
        )

    def schedule(self, send_payload=None):
        return self.uploader.schedule(asyncio.get_running_loop(), self.current_ws, **self.args(send_payload))

    async def send_success(self, ws, payload):
        kind = payload["type"]
        if kind == "call_recording_offer":
            self.offers.append((payload["recording_id"], asyncio.get_running_loop().time()))
            self.uploader.handle_server_message({
                "type": "call_recording_offer_ack", "recording_id": payload["recording_id"],
                "ok": True, "accepted": True,
            }, websocket=ws)
        elif kind == "call_recording_end":
            self.uploader.handle_server_message({
                "type": "call_recording_result", "recording_id": payload["recording_id"], "ok": True,
            }, websocket=ws)
            self.completed.set()
        return True

    async def reject(self, ws, payload):
        if payload["type"] == "call_recording_offer":
            self.offers.append((payload["recording_id"], asyncio.get_running_loop().time()))
            self.uploader.handle_server_message({
                "type": "call_recording_offer_ack", "recording_id": payload["recording_id"],
                "ok": False, "accepted": False, "reason": "too_many_uploads",
            }, websocket=ws)
        return True

    async def test_repeated_schedules_cannot_bypass_backoff(self):
        for _ in range(6):
            self.schedule(self.reject)
            await asyncio.sleep(0.01)
        self.assertEqual(len(self.offers), 1)

    async def test_failed_recording_does_not_block_other_pending_recording(self):
        second = create_recording(self.repository, "z-second")

        async def send(ws, payload):
            if payload["recording_id"] == self.recording.recording_id:
                return await self.reject(ws, payload)
            return await self.send_success(ws, payload)

        self.schedule(send)
        await asyncio.wait_for(self.completed.wait(), 1)
        self.assertEqual([item[0] for item in self.offers[:2]], [self.recording.recording_id, second.recording_id])
        self.assertLess(self.offers[1][1] - self.offers[0][1], RETRY_SECONDS)

    async def test_sidecar_failure_keeps_memory_backoff(self):
        attempts = []
        original = self.repository.mark_uploading

        def mark_uploading(recording):
            attempts.append(asyncio.get_running_loop().time())
            if len(attempts) == 1:
                return False
            return original(recording)

        with patch.object(self.repository, "mark_uploading", side_effect=mark_uploading), patch.object(
            self.repository, "mark_pending", side_effect=OSError("synthetic sidecar failure")
        ):
            for _ in range(6):
                self.schedule()
                await asyncio.sleep(0.01)
            self.assertEqual(len(attempts), 1)
            await asyncio.wait_for(self.completed.wait(), 1)
        self.assertEqual(len(attempts), 2)
        self.assertGreaterEqual(attempts[1] - attempts[0], RETRY_SECONDS - 0.01)

    async def test_restarted_uploader_retries_abandoned_upload_without_new_messages(self):
        self.repository.mark_uploading(self.recording)
        self.repository = CallRecordingRepository(self.temp.name)
        self.uploader = CloudCallRecordingUploader(self.repository)
        started = asyncio.get_running_loop().time()
        self.schedule()
        await asyncio.wait_for(self.completed.wait(), 1)
        self.assertGreaterEqual(self.offers[0][1] - started, RETRY_SECONDS - 0.03)

    async def test_wall_clock_rollback_does_not_extend_active_retry(self):
        async def reject_once(ws, payload):
            if not self.offers:
                return await self.reject(ws, payload)
            return await self.send_success(ws, payload)

        now = time.time()
        with patch("sms_core.call_recordings.time.time", return_value=now) as clock:
            await self.uploader._drain(self.ws, **self.args(reject_once))
            clock.return_value = now - 3600
            await asyncio.wait_for(self.completed.wait(), 1)
        self.assertEqual(len(self.offers), 2)
        self.assertGreaterEqual(self.offers[1][1] - self.offers[0][1], RETRY_SECONDS - 0.01)

    async def test_future_sidecar_deadline_is_bounded_after_restart(self):
        self.repository.mark_uploading(self.recording)
        self.repository = CallRecordingRepository(self.temp.name)
        self.uploader = CloudCallRecordingUploader(self.repository)
        started = asyncio.get_running_loop().time()
        with patch("sms_core.call_recordings.time.time", return_value=time.time() - 3600):
            self.schedule()
            await asyncio.wait_for(self.completed.wait(), 1)
        self.assertEqual(len(self.offers), 1)
        self.assertGreaterEqual(self.offers[0][1] - started, RETRY_SECONDS - 0.03)

    async def test_failed_pending_state_retains_retry_deadline_across_restart(self):
        self.repository.mark_pending(self.recording, "temporary_failure", retry_after=RETRY_SECONDS)
        self.repository = CallRecordingRepository(self.temp.name)
        self.assertEqual(self.repository.pending(IMEI), [])
        self.uploader = CloudCallRecordingUploader(self.repository)
        self.schedule()
        await asyncio.wait_for(self.completed.wait(), 1)
        self.assertEqual(len(self.offers), 1)

    async def test_expired_retry_does_not_send_on_replaced_connection_or_revoked_auth(self):
        for change in ("connection", "authorization"):
            with self.subTest(change=change):
                self.current_ws = self.ws
                self.authorized = True
                await self.uploader._drain(self.ws, **self.args(self.reject))
                before = len(self.offers)
                if change == "connection":
                    self.current_ws = object()
                else:
                    self.authorized = False
                await asyncio.sleep(RETRY_SECONDS + 0.05)
                self.assertEqual(len(self.offers), before)
                self.assertIsNone(self.uploader._retry_handle)

    async def test_retry_reservation_does_not_follow_an_imei_change(self):
        other_imei = "000000000000002"
        create_recording(self.repository, "other-device", other_imei)
        await self.uploader._drain(self.ws, **self.args(self.reject))
        self.imei = other_imei
        await asyncio.sleep(RETRY_SECONDS + 0.05)
        self.assertEqual(len(self.offers), 1)
        self.schedule()
        await asyncio.wait_for(self.completed.wait(), 1)
        self.assertEqual(self.offers[-1][0], "other-device")

    async def test_stop_cancels_queued_start_and_timer_but_can_restart(self):
        self.schedule()
        await self.uploader.stop()
        await asyncio.sleep(0)
        self.assertEqual(self.offers, [])
        await self.uploader._drain(self.ws, **self.args(self.reject))
        await self.uploader.stop()
        before = len(self.offers)
        await asyncio.sleep(RETRY_SECONDS + 0.03)
        self.assertEqual(len(self.offers), before)
        self.assertIsNone(self.uploader._retry_handle)
        self.schedule()
        await asyncio.wait_for(self.completed.wait(), 1)

    async def test_stop_restores_interrupted_upload_and_cleans_response_waiters(self):
        offered = asyncio.Event()

        async def hold_offer(_ws, payload):
            if payload["type"] == "call_recording_offer":
                offered.set()
            return True

        self.schedule(hold_offer)
        await asyncio.wait_for(offered.wait(), 1)
        await self.uploader.stop()
        metadata = json.loads(Path(self.recording.metadata_path).read_text(encoding="utf-8"))
        self.assertEqual(metadata["upload_status"], "pending")
        self.assertFalse(self.uploader._offer_waiters)
        self.assertFalse(self.uploader._result_waiters)
        self.assertIsNone(self.uploader._task)
        self.schedule()
        await asyncio.wait_for(self.completed.wait(), 1)

    async def test_stop_cleans_task_cancelled_before_its_coroutine_starts(self):
        self.schedule()
        await asyncio.sleep(0)
        self.assertIsNotNone(self.uploader._task)
        await self.uploader.stop()
        self.assertIsNone(self.uploader._task)
        self.assertEqual(self.offers, [])

    async def test_sender_cleanup_cannot_start_another_upload_during_stop(self):
        offered = asyncio.Event()
        cleanup_schedules = []

        async def hold_send(_ws, _payload):
            offered.set()
            try:
                await asyncio.Event().wait()
            finally:
                cleanup_schedules.append(self.schedule())
                await asyncio.sleep(0)

        self.schedule(hold_send)
        await asyncio.wait_for(offered.wait(), 1)
        await self.uploader.stop()
        self.assertEqual(cleanup_schedules, [False])
        self.assertIsNone(self.uploader._task)
        self.assertIsNone(self.uploader._retry_handle)
        self.assertEqual(self.offers, [])


if __name__ == "__main__":
    unittest.main()
