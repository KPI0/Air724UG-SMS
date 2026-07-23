import configparser
import threading
import unittest

from sms_ui.cloud_control_namespace_runtime import (
    cloud_log_namespace_runtime,
    open_cloud_control_window_namespace_runtime,
    refresh_cloud_control_settings_namespace_runtime,
    restart_cloud_control_namespace_runtime,
    save_cloud_control_setting_namespace_runtime,
    start_cloud_control_namespace_runtime,
    stop_cloud_control_namespace_runtime,
)


class FakeMessageBox:
    def showwarning(self, *args):
        return ("warning", args)


class FakeThreading:
    class Thread:
        pass


class FakeAsyncio:
    @staticmethod
    def run_coroutine_threadsafe(coro, loop):
        return ("future", coro, loop)


class FakeRoot:
    def after(self, *args):
        return ("after", args)


class CloudControlNamespaceRuntimeTests(unittest.TestCase):
    def base_namespace(self):
        config = configparser.ConfigParser(interpolation=None)
        return {
            "websockets": object(),
            "CLOUD_WS_URL": "ws://host",
            "CLOUD_DEVICE_SECRET": "secret",
            "CLOUD_WS_RECONNECT_INTERVAL": 5,
            "CLOUD_CONTROL_ENABLED": True,
            "CLOUD_AUTO_UPLOAD": False,
            "cloud_ws_lock": "lock",
            "cloud_ws_thread": "thread",
            "cloud_stop_event": "stop_event",
            "cloud_ws_loop": "loop",
            "cloud_ws_conn": "ws",
            "cloud_connected": True,
            "cloud_device_authorized": True,
            "cloud_restart_seq": 3,
            "cloud_control_win": "window",
            "cloud_var": "status_var",
            "config": config,
            "CONFIG_FILE": "config.ini",
            "CONFIG_LOCK": threading.RLock(),
            "root": FakeRoot(),
            "messagebox": FakeMessageBox(),
            "threading": FakeThreading,
            "asyncio": FakeAsyncio,
            "CLOUD_WS_RECONNECT_INTERVAL": 5,
            "set_cloud_status": lambda *args: ("status", args),
            "_cloud_log": lambda *args, **kwargs: ("log", args, kwargs),
            "apply_cloud_control_settings": lambda settings: ("apply", settings),
            "safe_save_config": lambda: ("save_config",),
            "system_ui": lambda *args: ("system_ui", args),
            "_cloud_runtime_imei": lambda: "861",
            "request_cloud_device_imei": lambda: "imei",
            "_cloud_thread_main": lambda *args: ("thread_main", args),
            "_reset_cloud_serial_log_state": lambda: "reset",
            "_cloud_unregister_then_close": lambda ws: ("close", ws),
            "tk_alive": lambda: True,
            "ui_post": lambda callback: callback(),
            "stop_cloud_control": lambda **kwargs: ("stop", kwargs),
            "start_cloud_control": lambda **kwargs: ("start", kwargs),
            "refresh_cloud_control_settings_from_config": lambda: "refresh",
            "save_cloud_control_setting": lambda **kwargs: ("save", kwargs),
            "_cloud_send_register": lambda ws: ("register", ws),
            "_cloud_schedule_unregister": lambda reason="hidden": ("unregister", reason),
            "restart_cloud_control": lambda **kwargs: ("restart", kwargs),
            "sync_and_focus_existing_window": lambda win, attr, **kwargs: ("sync", win, attr, kwargs["log_error"]("sync log")),
            "center_window": lambda win: ("center", win),
            "_cloud_file_notice": "file_notice",
            "_cloud_main_notice": "main_notice",
            "_cloud_repeat_filter": lambda notice, message: f"{notice}:{message}",
            "log_file_only": lambda message: ("file", message),
            "ui_only": lambda *args: ("ui", args),
        }

    def test_refresh_cloud_control_settings_applies_staged_values(self):
        namespace = self.base_namespace()
        staged = configparser.ConfigParser(interpolation=None)
        staged["cloud_control"] = {
            "enabled": "0",
            "url": "wss://new.example/ws/device",
            "device_secret": "next-secret",
            "reconnect_interval": "9",
            "auto_upload": "1",
        }
        applied = []
        namespace["apply_cloud_control_settings"] = applied.append

        def reload_config(**kwargs):
            self.assertIs(kwargs["config"], namespace["config"])
            self.assertIs(kwargs["config_lock"], namespace["CONFIG_LOCK"])
            return kwargs["read_values"](staged)

        result = refresh_cloud_control_settings_namespace_runtime(
            namespace,
            reload_config=reload_config,
        )

        self.assertTrue(result)
        self.assertEqual(len(applied), 1)
        self.assertFalse(applied[0].enabled)
        self.assertEqual(applied[0].url, "wss://new.example/ws/device")
        self.assertEqual(applied[0].reconnect_interval, 9)
        self.assertTrue(applied[0].auto_upload)

    def test_refresh_cloud_control_failure_preserves_runtime_and_redacts_error(self):
        namespace = self.base_namespace()
        logs = []
        applied = []
        namespace["log_file_only"] = logs.append
        namespace["apply_cloud_control_settings"] = applied.append

        result = refresh_cloud_control_settings_namespace_runtime(
            namespace,
            reload_config=lambda **_kwargs: (_ for _ in ()).throw(
                ValueError("device_secret=do-not-log")
            ),
        )

        self.assertFalse(result)
        self.assertEqual(applied, [])
        self.assertEqual(logs, ["Reload cloud-control config failed (ValueError)"])
        self.assertNotIn("do-not-log", logs[0])

    def test_start_cloud_control_namespace_runtime_forwards_state(self):
        namespace = self.base_namespace()
        calls = []

        result = start_cloud_control_namespace_runtime(
            namespace,
            show_errors=True,
            start_app_runtime=lambda **kwargs: calls.append(kwargs) or "started",
        )

        self.assertEqual(result, "started")
        self.assertTrue(calls[0]["websockets_available"])
        self.assertEqual(calls[0]["url"], "ws://host")
        self.assertEqual(calls[0]["device_secret"], "secret")
        self.assertTrue(calls[0]["show_errors"])
        self.assertIs(calls[0]["thread_factory"], FakeThreading.Thread)
        calls[0]["set_thread"]("next_thread")
        self.assertEqual(namespace["cloud_ws_thread"], "next_thread")

    def test_stop_cloud_control_namespace_runtime_forwards_mutators(self):
        namespace = self.base_namespace()
        calls = []

        result = stop_cloud_control_namespace_runtime(
            namespace,
            update_status=False,
            stop_app_runtime=lambda **kwargs: calls.append(kwargs) or "stopped",
        )

        self.assertEqual(result, "stopped")
        forwarded = calls[0]
        self.assertFalse(forwarded["update_status"])
        self.assertTrue(forwarded["enabled"])
        forwarded["set_connected"](False)
        forwarded["set_authorized"](False)
        forwarded["set_ws"](None)
        self.assertFalse(namespace["cloud_connected"])
        self.assertFalse(namespace["cloud_device_authorized"])
        self.assertIsNone(namespace["cloud_ws_conn"])
        self.assertEqual(forwarded["run_coroutine_threadsafe"]("coro", "loop"), ("future", "coro", "loop"))

    def test_restart_cloud_control_namespace_runtime_forwards_restart_state(self):
        namespace = self.base_namespace()
        calls = []

        result = restart_cloud_control_namespace_runtime(
            namespace,
            restart_app_runtime=lambda **kwargs: calls.append(kwargs) or "restarted",
        )

        self.assertEqual(result, "restarted")
        forwarded = calls[0]
        self.assertEqual(forwarded["get_restart_seq"](), 3)
        forwarded["set_restart_seq"](4)
        self.assertEqual(namespace["cloud_restart_seq"], 4)
        self.assertIs(forwarded["thread_factory"], FakeThreading.Thread)
        self.assertEqual(forwarded["schedule_after"]("tick"), ("after", ("tick",)))

    def test_save_cloud_control_setting_namespace_runtime_forwards_current_state(self):
        namespace = self.base_namespace()
        calls = []

        result = save_cloud_control_setting_namespace_runtime(
            namespace,
            enabled=False,
            save_app_runtime=lambda **kwargs: calls.append(kwargs) or kwargs["current_settings"](),
        )

        self.assertFalse(calls[0]["enabled"])
        self.assertIs(calls[0]["config"], namespace["config"])
        self.assertIs(calls[0]["apply_settings"], namespace["apply_cloud_control_settings"])
        self.assertEqual(result.enabled, True)
        self.assertEqual(result.url, "ws://host")
        self.assertEqual(result.reconnect_interval, 5)
        self.assertEqual(result.device_secret, "secret")
        self.assertEqual(result.auto_upload, False)

    def test_cloud_log_namespace_runtime_filters_file_and_main_messages(self):
        namespace = self.base_namespace()
        calls = []
        namespace["log_file_only"] = lambda message: calls.append(("file", message))
        namespace["ui_only"] = lambda *args: calls.append(("ui", args))

        cloud_log_namespace_runtime(namespace, "hello", show_main=True)

        self.assertEqual(calls[0], ("file", "file_notice:🌐 hello"))
        self.assertEqual(calls[1], ("ui", ("main_notice:🌐 hello", "normal")))

    def test_cloud_log_namespace_runtime_falls_back_when_filter_fails(self):
        namespace = self.base_namespace()
        calls = []
        namespace["_cloud_repeat_filter"] = lambda *_: (_ for _ in ()).throw(RuntimeError("filter"))
        namespace["log_file_only"] = lambda message: calls.append(("file", message))
        namespace["ui_only"] = lambda *args: calls.append(("ui", args))

        cloud_log_namespace_runtime(namespace, "hello", show_main=True)

        self.assertEqual(calls[0], ("file", "🌐 hello"))
        self.assertEqual(calls[1], ("ui", ("🌐 hello", "normal")))

    def test_cloud_log_namespace_runtime_suppresses_shutdown_messages(self):
        namespace = self.base_namespace()
        calls = []

        class SetEvent:
            @staticmethod
            def is_set():
                return True

        namespace["TK_SHUTDOWN"] = SetEvent()
        namespace["log_file_only"] = lambda message: calls.append(("file", message))
        namespace["ui_only"] = lambda *args: calls.append(("ui", args))

        self.assertIsNone(cloud_log_namespace_runtime(namespace, "late", show_main=True))
        self.assertEqual(calls, [])

    def test_open_cloud_control_window_namespace_runtime_forwards_values(self):
        namespace = self.base_namespace()
        calls = []

        result = open_cloud_control_window_namespace_runtime(
            namespace,
            open_values_app_runtime=lambda parent, **kwargs: calls.append((parent, kwargs)) or "opened",
        )

        self.assertEqual(result, "opened")
        parent, forwarded = calls[0]
        self.assertIs(parent, namespace["root"])
        self.assertEqual(forwarded["current_window"], "window")
        self.assertFalse(forwarded["auto_upload"])
        self.assertEqual(forwarded["url"], "ws://host")
        self.assertEqual(forwarded["status_var"], "status_var")
        self.assertTrue(forwarded["is_connected"]())
        self.assertEqual(forwarded["get_loop"](), "loop")
        self.assertEqual(forwarded["get_ws"](), "ws")
        self.assertTrue(forwarded["settings_provider"]()["enabled"])
        namespace["CLOUD_CONTROL_ENABLED"] = False
        namespace["CLOUD_WS_URL"] = "wss://new.example/ws/device"
        self.assertFalse(forwarded["settings_provider"]()["enabled"])
        self.assertEqual(forwarded["settings_provider"]()["url"], "wss://new.example/ws/device")
        forwarded["set_window"]("next_window")
        self.assertEqual(namespace["cloud_control_win"], "next_window")
        self.assertEqual(forwarded["run_coroutine_threadsafe"]("coro", "loop"), ("future", "coro", "loop"))
        self.assertEqual(
            forwarded["sync_existing_window"]("win", "_sync"),
            ("sync", "win", "_sync", ("file", "sync log")),
        )


if __name__ == "__main__":
    unittest.main()
