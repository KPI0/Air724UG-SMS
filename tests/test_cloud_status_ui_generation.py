import asyncio
import queue
import threading
from types import SimpleNamespace
import unittest

from sms_core.cloud_protocol import auth_status_from_ack
from sms_core.cloud_runtime import CloudControlSettings
from sms_core.cloud_ws_namespace_runtime import wait_cloud_login_ack_namespace_runtime
from sms_ui.app_ui_namespace_runtime import (
    set_cloud_auth_status_from_ack_namespace_runtime,
    set_cloud_status_namespace_runtime,
)
from sms_ui.cloud_control_namespace_runtime import (
    apply_cloud_control_settings_namespace_runtime,
    stop_cloud_control_namespace_runtime,
)
from sms_ui.thread_runtime import run_on_ui_thread, ui_pump_runtime
from sms_ui.ui_log_namespace_runtime import ui_post_namespace_runtime


class FakeVar:
    def __init__(self):
        self.value = "initial"
        self.updates = []

    def set(self, value):
        self.value = value
        self.updates.append((value, threading.get_ident()))


class CloudStatusUiGenerationTests(unittest.TestCase):
    def setUp(self):
        self.owner = threading.get_ident()
        self.task_queue = queue.Queue()
        self.var = FakeVar()
        self.colors = []
        self.errors = []
        self.ns = {
            "UI_TASK_QUEUE": self.task_queue,
            "cloud_var": self.var,
            "cloud_label": SimpleNamespace(config=lambda **values: self.colors.append(values)),
            "tk_alive": lambda: True,
            "run_on_ui_thread": run_on_ui_thread,
            "cloud_stop_event": threading.Event(),
            "cloud_ws_lock": threading.RLock(),
            "cloud_restart_seq": 0,
            "cloud_device_authorized": False,
            "cloud_connected": True,
            "cloud_ws_loop": None,
            "CLOUD_CONTROL_ENABLED": True,
            "CLOUD_WS_URL": "wss://example.invalid/ws/device",
            "CLOUD_DEVICE_SECRET": "synthetic-secret",
            "CLOUD_WS_RECONNECT_INTERVAL": 5,
            "CLOUD_AUTO_UPLOAD": False,
            "_cloud_runtime_imei": lambda: "861",
            "_cloud_log": lambda *args, **kwargs: None,
            "_cloud_safe_preview": lambda text: "hidden",
            "_cloud_auth_status_from_ack": auth_status_from_ack,
            "_reset_cloud_serial_log_state": lambda: None,
            "_cloud_unregister_then_close": lambda ws: None,
            "log_file_only": self.errors.append,
        }
        self.ns["ui_post"] = lambda callback, *args, **kwargs: ui_post_namespace_runtime(
            self.ns, callback, *args, **kwargs,
        )
        self.ns["set_cloud_status"] = lambda *args: set_cloud_status_namespace_runtime(self.ns, *args)
        self.ns["set_cloud_auth_status_from_ack"] = lambda data: set_cloud_auth_status_from_ack_namespace_runtime(
            self.ns, data,
        )

    def run_worker(self, callback):
        failures = []

        def work():
            try:
                callback()
            except Exception as exc:
                failures.append(exc)

        worker = threading.Thread(target=work)
        worker.start()
        worker.join(2)
        self.assertFalse(worker.is_alive())
        self.assertEqual(failures, [])

    def pump(self):
        return ui_pump_runtime(
            self.task_queue, SimpleNamespace(), lambda: False, lambda: None,
            log_error=self.errors.append,
        )

    def tearDown(self):
        self.assertEqual(self.errors, [])
        self.assertEqual(self.task_queue.unfinished_tasks, 0)
        self.assertTrue(all(owner == self.owner for _value, owner in self.var.updates))

    def test_queued_login_ack_cannot_replace_disabled_cloud_status(self):
        class Socket:
            async def recv(self):
                return '{"type":"device_login_ack","auth_status":"authorized","imei":"861"}'

        ws = Socket()
        self.ns["cloud_ws_conn"] = ws
        results = []
        self.run_worker(lambda: results.append(asyncio.run(
            wait_cloud_login_ack_namespace_runtime(self.ns, ws),
        )))
        self.assertEqual(results, [True])
        self.assertTrue(self.ns["cloud_device_authorized"])
        self.assertEqual(self.task_queue.qsize(), 1)
        apply_cloud_control_settings_namespace_runtime(self.ns, CloudControlSettings(
            False, self.ns["CLOUD_WS_URL"], 5, self.ns["CLOUD_DEVICE_SECRET"], False,
        ))
        stop_cloud_control_namespace_runtime(self.ns)
        stopped_status = self.var.value
        self.assertEqual(stopped_status, "🌐 已关闭")
        self.assertFalse(self.ns["CLOUD_CONTROL_ENABLED"])
        self.assertFalse(self.ns["cloud_device_authorized"])
        self.pump()
        self.assertEqual(self.var.value, stopped_status)
        self.assertEqual(self.colors, [{"fg": "#666666"}])

    def test_queued_old_failure_cannot_replace_new_connection_status(self):
        self.run_worker(lambda: self.ns["set_cloud_status"]("🌐 授权失败", "#cc0000"))
        self.assertTrue(self.ns["set_cloud_status"]("🌐 连接中", "#b26a00"))
        self.pump()
        self.assertEqual(self.var.value, "🌐 连接中")
        self.assertEqual(self.colors, [{"fg": "#b26a00"}])

    def test_replayed_stale_callback_cannot_invalidate_new_queued_status(self):
        self.run_worker(lambda: self.ns["set_cloud_status"]("old", "red"))
        old_callback, args, kwargs = self.task_queue.get_nowait()
        self.task_queue.task_done()
        self.run_worker(lambda: self.ns["set_cloud_status"]("current", "green"))
        old_callback(*args, **kwargs)
        old_callback(*args, **kwargs)
        self.assertEqual(self.var.value, "initial")
        self.pump()
        self.assertEqual(self.var.value, "current")
        old_callback(*args, **kwargs)
        self.assertEqual(self.var.value, "current")
        self.assertEqual(self.colors, [{"fg": "green"}])

    def test_latest_queued_cloud_status_updates_through_existing_ui_dispatcher(self):
        self.run_worker(lambda: self.ns["set_cloud_status"]("current", "green"))
        self.assertEqual(self.var.value, "initial")
        self.assertEqual(self.pump(), 1)
        self.assertEqual(self.var.value, "current")
        self.assertEqual(self.colors, [{"fg": "green"}])


if __name__ == "__main__":
    unittest.main()
