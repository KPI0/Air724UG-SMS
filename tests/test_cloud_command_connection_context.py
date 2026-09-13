import asyncio
import json
import threading
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from sms_app.cloud_namespace_bindings import install_cloud_namespace_bindings
from sms_core import serial_sender
from sms_core.serial_startup_runtime import open_and_initialize_serial_runtime
from sms_core.threading_runtime import WorkerThreadRegistry
from sms_ui.app_infrastructure_namespace_runtime import safe_close_serial_namespace_runtime


COMMANDS = (
    ("ATI", {}),
    ("CLOUD:SEND_SMS", {"sms_phone": "10086", "sms_message": "synthetic queue test"}),
    ("CLOUD:SET_OWN_NUMBER", {"own_number": "+19990000001"}),
)


class Modem:
    def __init__(self, at_coordinator, sms_coordinator):
        self.at = at_coordinator
        self.sms = sms_coordinator
        self.is_open = True
        self.writes = []

    def write(self, payload):
        assert self.is_open
        self.writes.append(payload)
        if payload.startswith(b"AT+CMGS="):
            self.sms.observe_line("> ", connection=self)
        elif payload.endswith(b"\x1a"):
            self.sms.observe_line("+CMGS: 1", connection=self)
            self.sms.observe_line("OK", connection=self)
        else:
            self.at.observe_line("OK", connection=self)
        return len(payload)

    def flush(self):
        pass

    def close(self):
        self.is_open = False


class CloudCommandConnectionContextTests(unittest.IsolatedAsyncioTestCase):
    async def exercise(self, command, metadata, pause, change):
        loop = asyncio.get_running_loop()
        reached = asyncio.Event()
        async_release = asyncio.Event()
        worker_release = threading.Event()
        workers = []
        replies = []
        at = serial_sender.AtCommandResponseCoordinator()
        sms = serial_sender.SmsPduSendCoordinator()
        original, replacement = Modem(at, sms), Modem(at, sms)
        websocket = object()
        namespace = {
            "serial_lock": threading.RLock(), "serial_obj": original,
            "serial_connection_generation": 1, "cloud_imei_verified": True,
            "CLOUD_DEVICE_IMEI": "000000000000001", "CLOUD_DEVICE_SECRET": "synthetic-only",
            "CLOUD_WS_URL": "wss://example.invalid/ws/device", "CLOUD_CONTROL_ENABLED": True,
            "cloud_ws_conn": websocket, "cloud_connected": True, "cloud_device_authorized": True,
            "SMS_SEND_COORDINATOR": sms, "SERIAL_COMMAND_RESPONSE_COORDINATOR": at,
            "SERIAL_COMMAND_THREAD_REGISTRY": WorkerThreadRegistry(),
            "CLOUD_SENSITIVE_COMMAND_PERMISSIONS": {"sms": True, "phone_number": True},
            "unlock_port_mutex": lambda: None, "show_window": lambda: None,
            "hide_window": lambda: None, "set_cloud_auth_status_from_ack": lambda data: None,
            "set_cloud_status": lambda *args: None, "_normalize_imei": lambda value: value,
        }
        install_cloud_namespace_bindings(namespace)

        async def reply(_ws, payload):
            replies.append(payload)
            if pause == "started" and payload.get("status") == "started":
                reached.set()
                await async_release.wait()

        async def check_replay(_ws, _data, mark_seen=True):
            if pause == "replay" and not mark_seen:
                reached.set()
                await async_release.wait()
            return True

        def thread_factory(*args, target, **kwargs):
            def run():
                if pause == "worker":
                    loop.call_soon_threadsafe(reached.set)
                    assert worker_release.wait(5)
                target()
            thread = threading.Thread(*args, target=run, **kwargs)
            workers.append(thread)
            return thread

        namespace.update({
            "_cloud_reply": reply, "_cloud_log": lambda *args, **kwargs: None,
            "_cloud_check_replay_window": check_replay,
            "_cloud_auth_matches": lambda data: data.get("target_imei") == namespace["CLOUD_DEVICE_IMEI"]
                and data.get("secret") == namespace["CLOUD_DEVICE_SECRET"],
            "_cloud_send_status_payload": lambda: {}, "_notify_cloud_channel_status": lambda value: None,
            "_handle_cloud_call_recording_message": lambda *args, **kwargs: False,
            "threading": SimpleNamespace(Thread=thread_factory),
        })
        transaction_lock = serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK
        loop_thread = threading.current_thread()

        class ObservedLock:
            def __enter__(self):
                if threading.current_thread() is not loop_thread:
                    loop.call_soon_threadsafe(reached.set)
                transaction_lock.acquire()

            def __exit__(self, *_args):
                transaction_lock.release()

        patcher = patch.object(serial_sender, "DEFAULT_SERIAL_TRANSACTION_LOCK", ObservedLock())
        if pause == "transaction":
            transaction_lock.acquire()
            patcher.start()
        message = json.dumps({
            "type": "cmd", "task_id": "context-test", "target_imei": namespace["CLOUD_DEVICE_IMEI"],
            "secret": namespace["CLOUD_DEVICE_SECRET"], "command": command, **metadata,
        })
        handling = asyncio.create_task(namespace["_handle_cloud_message"](websocket, message))
        old_offset = new_offset = 0
        try:
            await asyncio.wait_for(reached.wait(), 3)
            if change in ("replaced", "disconnected", "same_object_reopened"):
                safe_close_serial_namespace_runtime(namespace)
                if change != "disconnected":
                    next_serial = original if change == "same_object_reopened" else replacement
                    next_serial.is_open = True
                    open_and_initialize_serial_runtime(
                        target_port="SYNTHETIC_PORT", baud=115200, mode="Manual",
                        serial_lock=namespace["serial_lock"], open_serial=lambda *args: next_serial,
                        set_serial_obj=lambda value: namespace.__setitem__("serial_obj", value),
                        set_port=lambda value: None, lock_port_mutex=lambda value: None,
                        set_cloud_imei_query_deadline=lambda value: None,
                        serial_error_ui=lambda *args, **kwargs: None,
                        set_status=lambda *args, **kwargs: None,
                    )
                    self.assertEqual(next_serial.writes, [b"AT+CLIP=1\r\n", b"AT+CGSN\r\n", b"AT+CNUM\r\n"])
            elif change == "imei":
                namespace["CLOUD_DEVICE_IMEI"] = "000000000000002"
            elif change == "websocket":
                namespace["cloud_ws_conn"] = object()
            elif change == "authorization":
                namespace["cloud_device_authorized"] = False
            elif change == "secret":
                namespace["CLOUD_DEVICE_SECRET"] = "changed-synthetic-only"
            elif change == "disabled":
                namespace["CLOUD_CONTROL_ENABLED"] = False
            old_offset, new_offset = len(original.writes), len(replacement.writes)
        finally:
            async_release.set()
            worker_release.set()
            if pause == "transaction":
                transaction_lock.release()
            try:
                await asyncio.wait_for(handling, 5)
            finally:
                for worker in workers:
                    await asyncio.to_thread(worker.join, 2)
                    self.assertFalse(worker.is_alive())
                if pause == "transaction":
                    patcher.stop()
        result = [item for item in replies if item.get("type") == "send_at_result"]
        self.assertEqual(len(result), 1)
        if change == "unchanged":
            self.assertTrue(result[0]["ok"], result[0].get("message"))
            self.assertTrue(original.writes)
        else:
            self.assertFalse(result[0]["ok"], "old cloud command reported success after context changed")
            self.assertEqual(original.writes[old_offset:], [], "stale command still wrote to original connection")
            self.assertEqual(replacement.writes[new_offset:], [], "old command reached replacement device")
        self.assertEqual(namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(), ())
        self.assertIsNone(at._active)
        self.assertIsNone(sms._active)

    async def test_switch_while_started_ack_is_blocked_rejects_all_transaction_types(self):
        for command, metadata in COMMANDS:
            with self.subTest(command=command):
                await self.exercise(command, metadata, "started", "replaced")

    async def test_switch_before_registered_worker_runs_rejects_all_transaction_types(self):
        for command, metadata in COMMANDS:
            with self.subTest(command=command):
                await self.exercise(command, metadata, "worker", "replaced")

    async def test_switch_while_waiting_for_transaction_lock_rejects_all_transaction_types(self):
        for command, metadata in COMMANDS:
            with self.subTest(command=command):
                await self.exercise(command, metadata, "transaction", "replaced")

    async def test_switch_during_replay_validation_cannot_change_target(self):
        await self.exercise("ATI", {}, "replay", "replaced")

    async def test_reopened_same_serial_object_is_a_different_generation(self):
        await self.exercise("ATI", {}, "transaction", "same_object_reopened")

    async def test_identity_and_cloud_access_changes_cancel_waiting_work(self):
        for change in ("imei", "websocket", "authorization", "secret", "disabled"):
            with self.subTest(change=change):
                await self.exercise("ATI", {}, "transaction", change)

    async def test_disconnect_without_replacement_remains_a_failure(self):
        await self.exercise("ATI", {}, "transaction", "disconnected")

    async def test_unchanged_connection_completes_all_transaction_types(self):
        for command, metadata in COMMANDS:
            with self.subTest(command=command):
                await self.exercise(command, metadata, "transaction", "unchanged")
