import configparser
import unittest
from unittest.mock import patch

from sms_core.cloud_runtime import CloudControlSettings
from sms_ui.cloud_control_app_runtime import (
    cloud_control_settings_from_values,
    cloud_window_connection_state,
    open_cloud_control_app_runtime,
    open_cloud_control_values_app_runtime,
    register_current_cloud_connection,
    restart_cloud_control_app_runtime,
    save_cloud_control_setting_runtime,
    start_cloud_control_app_runtime,
    stop_cloud_control_app_runtime,
)
from sms_ui.cloud_control_window import open_cloud_control_window_runtime


class CloudControlAppRuntimeTests(unittest.TestCase):
    def test_cloud_control_settings_from_values_builds_settings(self):
        settings = cloud_control_settings_from_values(
            True,
            "ws://host",
            7,
            "secret",
            False,
        )

        self.assertEqual(settings, CloudControlSettings(
            enabled=True,
            url="ws://host",
            reconnect_interval=7,
            device_secret="secret",
            auto_upload=False,
        ))

    def test_save_cloud_control_setting_runtime_updates_writes_and_saves(self):
        config = configparser.ConfigParser()
        calls = []

        result = save_cloud_control_setting_runtime(
            current_settings=lambda: CloudControlSettings(
                enabled=False,
                url="ws://old",
                reconnect_interval=5,
                device_secret="old",
                auto_upload=False,
            ),
            apply_settings=lambda settings: calls.append(("apply", settings)),
            config=config,
            save_config=lambda: calls.append(("save",)),
            system_ui=lambda *args: calls.append(("ui", args)),
            enabled=True,
            url="wss://example.test",
            reconnect_interval="9",
            device_secret=" secret ",
            auto_upload=True,
        )

        self.assertTrue(result.enabled)
        self.assertEqual(result.url, "wss://example.test/ws/device")
        self.assertEqual(result.reconnect_interval, 9)
        self.assertEqual(result.device_secret, "secret")
        self.assertTrue(result.auto_upload)
        self.assertEqual(calls[0], ("save",))
        self.assertEqual(calls[1], ("apply", result))
        self.assertEqual(config.get("cloud_control", "enabled"), "1")
        self.assertEqual(config.get("cloud_control", "auto_upload"), "1")

    def test_save_cloud_control_setting_runtime_reports_save_errors(self):
        calls = []
        config = configparser.ConfigParser()
        config["cloud_control"] = {
            "enabled": "0",
            "url": "ws://old",
            "device_secret": "old-secret",
            "device_imei": "legacy-imei",
            "extra_key": "keep-me",
        }
        before = dict(config.items("cloud_control", raw=True))

        result = save_cloud_control_setting_runtime(
            current_settings=lambda: CloudControlSettings(),
            apply_settings=lambda settings: calls.append(("apply", settings)),
            config=config,
            save_config=lambda: (_ for _ in ()).throw(RuntimeError("disk")),
            system_ui=lambda *args: calls.append(("ui", args)),
            enabled=True,
        )

        self.assertIsNone(result)
        self.assertFalse(any(call[0] == "apply" for call in calls))
        self.assertEqual(dict(config.items("cloud_control", raw=True)), before)
        self.assertEqual(calls[0][0], "ui")
        self.assertIn("disk", calls[0][1][0])

    def test_save_cloud_control_setting_runtime_reports_false_save_result(self):
        calls = []
        config = configparser.ConfigParser()
        config["cloud_control"] = {"enabled": "0", "url": "ws://old", "device_imei": "legacy"}
        before = dict(config.items("cloud_control", raw=True))

        result = save_cloud_control_setting_runtime(
            current_settings=lambda: CloudControlSettings(),
            apply_settings=lambda settings: calls.append(("apply", settings)),
            config=config,
            save_config=lambda: False,
            system_ui=lambda *args: calls.append(("ui", args)),
            enabled=True,
        )

        self.assertIsNone(result)
        self.assertFalse(any(call[0] == "apply" for call in calls))
        self.assertEqual(dict(config.items("cloud_control", raw=True)), before)
        self.assertEqual(calls[0][0], "ui")
        self.assertIn("配置保存失败", calls[0][1][0])

    def test_cloud_window_connection_state_reports_loop_and_socket(self):
        self.assertEqual(
            cloud_window_connection_state(lambda: True, lambda: "loop", lambda: None),
            (True, True, False),
        )
        self.assertEqual(
            cloud_window_connection_state(lambda: False, lambda: None, lambda: "ws"),
            (False, False, True),
        )

    def test_cloud_control_lifecycle_app_runtimes_forward_dependencies(self):
        calls = []
        restart_seq = {"value": 4}

        start_result = start_cloud_control_app_runtime(
            websockets_available=True,
            url="ws://host",
            device_secret="secret",
            reconnect_interval=5,
            show_errors=True,
            set_cloud_status=lambda *args: calls.append(("status", args)),
            cloud_log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            show_warning=lambda *args: calls.append(("warning", args)),
            runtime_imei=lambda: "imei",
            request_device_imei=lambda: calls.append(("imei",)),
            lock="lock",
            get_thread=lambda: None,
            set_thread=lambda thread: calls.append(("set_thread", thread)),
            stop_event="stop",
            thread_factory="thread_factory",
            thread_target="thread_target",
            start_runtime=lambda **kwargs: calls.append(("start_runtime", kwargs)) or "started",
        )
        stop_result = stop_cloud_control_app_runtime(
            update_status=True,
            enabled=True,
            stop_event="stop",
            set_connected=lambda value: calls.append(("connected", value)),
            set_authorized=lambda value: calls.append(("authorized", value)),
            reset_serial_log_state=lambda: calls.append(("reset",)),
            get_loop=lambda: "loop",
            get_ws=lambda: "ws",
            schedule_unregister_then_close=lambda ws: ("close", ws),
            set_ws=lambda ws: calls.append(("set_ws", ws)),
            set_cloud_status=lambda *args: calls.append(("status", args)),
            run_coroutine_threadsafe=lambda *args: calls.append(("future", args)),
            stop_runtime=lambda **kwargs: calls.append(("stop_runtime", kwargs)) or "stopped",
        )
        restart_result = restart_cloud_control_app_runtime(
            show_errors=False,
            lock="lock",
            get_restart_seq=lambda: restart_seq["value"],
            set_restart_seq=lambda value: restart_seq.__setitem__("value", value),
            get_thread=lambda: None,
            stop_control=lambda **kwargs: calls.append(("stop", kwargs)),
            tk_alive=lambda: True,
            stop_event="stop",
            set_cloud_status=lambda *args: calls.append(("status", args)),
            schedule_after=lambda *args: calls.append(("after", args)),
            ui_post=lambda callback: callback(),
            start_control=lambda **kwargs: calls.append(("start", kwargs)),
            thread_factory="thread_factory",
            restart_runtime=lambda **kwargs: calls.append(("restart_runtime", kwargs)) or kwargs["increment_restart_seq"](),
        )

        self.assertEqual(start_result, "started")
        self.assertEqual(stop_result, "stopped")
        self.assertEqual(restart_result, 5)
        self.assertEqual(restart_seq["value"], 5)
        self.assertEqual(calls[0][0], "start_runtime")
        self.assertEqual(calls[1][0], "stop_runtime")
        self.assertEqual(calls[2][0], "restart_runtime")

    def test_register_current_cloud_connection_schedules_register(self):
        calls = []

        async def send_register(ws):
            return ("sent", ws)

        result = register_current_cloud_connection(
            lambda: "loop",
            lambda: "ws",
            send_register,
            lambda coro, loop: calls.append((coro, loop)) or "future",
        )

        self.assertEqual(result, "future")
        self.assertEqual(calls[0][1], "loop")
        calls[0][0].close()

    def test_open_cloud_control_app_runtime_adapts_settings_and_callbacks(self):
        calls = []

        def open_window_runtime(
            parent,
            current_window,
            state_provider,
            status_var,
            refresh_settings,
            save_setting,
            get_connection_state,
            register_current,
            schedule_unregister,
            restart_control,
            stop_control,
            cloud_log,
            sync_existing_window,
            set_window,
            center_window,
        ):
            calls.append(("parent", parent))
            calls.append(("window", current_window))
            calls.append(("state", state_provider()))
            calls.append(("status_var", status_var))
            refresh_settings()
            save_setting(enabled=True)
            calls.append(("connection", get_connection_state()))
            register_current()
            schedule_unregister("reason")
            restart_control(show_errors=True)
            stop_control(update_status=False)
            cloud_log("saved")
            calls.append(("sync", sync_existing_window("win", "_sync")))
            set_window("new_win")
            calls.append(("center", center_window))
            return "new_win"

        result = open_cloud_control_app_runtime(
            "root",
            current_window="old_win",
            get_settings=lambda: {
                "enabled": True,
                "auto_upload": False,
                "url": "ws://host",
                "secret": "secret",
                "reconnect_interval": 5,
            },
            status_var="status_var",
            refresh_settings=lambda: calls.append(("refresh",)),
            save_setting=lambda **kwargs: calls.append(("save", kwargs)),
            get_connection_state=lambda: (True, True, False),
            register_current=lambda: calls.append(("register",)),
            schedule_unregister=lambda reason: calls.append(("unregister", reason)),
            restart_control=lambda **kwargs: calls.append(("restart", kwargs)),
            stop_control=lambda **kwargs: calls.append(("stop", kwargs)),
            cloud_log=lambda message: calls.append(("log", message)),
            sync_existing_window=lambda win, attr: (win, attr),
            set_window=lambda win: calls.append(("set_window", win)),
            center_window="center",
            open_window_runtime=open_window_runtime,
        )

        self.assertEqual(result, "new_win")
        self.assertEqual(calls[0], ("parent", "root"))
        self.assertEqual(calls[1], ("window", "old_win"))
        self.assertEqual(calls[2], ("state", {
            "enabled": True,
            "auto_upload": False,
            "url": "ws://host",
            "secret": "secret",
            "reconnect_interval": 5,
        }))
        self.assertIn(("refresh",), calls)
        self.assertIn(("save", {"enabled": True}), calls)
        self.assertIn(("connection", (True, True, False)), calls)
        self.assertIn(("register",), calls)
        self.assertIn(("unregister", "reason"), calls)
        self.assertIn(("restart", {"show_errors": True}), calls)
        self.assertIn(("stop", {"update_status": False}), calls)
        self.assertIn(("log", "saved"), calls)
        self.assertIn(("sync", ("win", "_sync")), calls)
        self.assertIn(("set_window", "new_win"), calls)
        self.assertIn(("center", "center"), calls)

    def test_open_cloud_control_app_runtime_builds_connection_callbacks(self):
        calls = []

        async def send_register(ws):
            return ws

        def open_window_runtime(
            parent,
            current_window,
            state_provider,
            status_var,
            refresh_settings,
            save_setting,
            get_connection_state,
            register_current,
            schedule_unregister,
            restart_control,
            stop_control,
            cloud_log,
            sync_existing_window,
            set_window,
            center_window,
        ):
            calls.append(("connection", get_connection_state()))
            calls.append(("register", register_current()))
            return "win"

        result = open_cloud_control_app_runtime(
            "root",
            current_window=None,
            get_settings=lambda: {
                "enabled": False,
                "auto_upload": True,
                "url": "ws://host",
                "secret": "secret",
                "reconnect_interval": 3,
            },
            status_var="status",
            refresh_settings=lambda: None,
            save_setting=lambda **_: None,
            is_connected=lambda: True,
            get_loop=lambda: "loop",
            get_ws=lambda: "ws",
            send_register=send_register,
            run_coroutine_threadsafe=lambda coro, loop: (coro.close(), ("future", loop))[1],
            schedule_unregister=lambda reason: None,
            restart_control=lambda **_: None,
            stop_control=lambda **_: None,
            cloud_log=lambda message: None,
            sync_existing_window=lambda win, attr: False,
            set_window=lambda win: None,
            center_window="center",
            open_window_runtime=open_window_runtime,
        )

        self.assertEqual(result, "win")
        self.assertIn(("connection", (True, True, True)), calls)
        self.assertIn(("register", ("future", "loop")), calls)

    def test_open_cloud_control_values_app_runtime_builds_state(self):
        calls = []

        def open_window_runtime(
            parent,
            current_window,
            state_provider,
            status_var,
            refresh_settings,
            save_setting,
            get_connection_state,
            register_current,
            schedule_unregister,
            restart_control,
            stop_control,
            cloud_log,
            sync_existing_window,
            set_window,
            center_window,
        ):
            calls.append(("parent", parent))
            calls.append(("window", current_window))
            calls.append(("state", state_provider()))
            calls.append(("connection", get_connection_state()))
            set_window("new_win")
            return "new_win"

        result = open_cloud_control_values_app_runtime(
            "root",
            current_window="old_win",
            enabled=True,
            auto_upload=True,
            url="ws://host",
            secret="secret",
            reconnect_interval=9,
            status_var="status",
            refresh_settings=lambda: None,
            save_setting=lambda **_: None,
            get_connection_state=lambda: (False, True, False),
            schedule_unregister=lambda reason: None,
            restart_control=lambda **_: None,
            stop_control=lambda **_: None,
            cloud_log=lambda message: None,
            sync_existing_window=lambda win, attr: False,
            set_window=lambda win: calls.append(("set_window", win)),
            center_window="center",
            open_window_runtime=open_window_runtime,
        )

        self.assertEqual(result, "new_win")
        self.assertEqual(calls[0], ("parent", "root"))
        self.assertEqual(calls[1], ("window", "old_win"))
        self.assertEqual(calls[2], ("state", {
            "enabled": True,
            "auto_upload": True,
            "url": "ws://host",
            "secret": "secret",
            "reconnect_interval": 9,
        }))
        self.assertIn(("connection", (False, True, False)), calls)
        self.assertIn(("set_window", "new_win"), calls)

    def test_open_cloud_control_window_reverts_auto_upload_after_save_failure(self):
        calls = []

        class FakeWindow:
            def _sync_form_from_state(self):
                calls.append(("sync",))

        def fake_dialog(
            _parent,
            state_provider,
            _status_var,
            on_auto_upload_changed,
            *_callbacks,
            **_kwargs,
        ):
            result = on_auto_upload_changed(True, FakeWindow())
            calls.append(("result", result, state_provider()))
            return "window"

        with patch("sms_ui.cloud_control_window.open_cloud_control_window_dialog", side_effect=fake_dialog):
            result = open_cloud_control_window_runtime(
                "root",
                None,
                lambda: {"enabled": False, "auto_upload": False},
                "status",
                lambda: None,
                lambda **_: None,
                lambda: (True, True, True),
                lambda: calls.append(("register",)),
                lambda reason: calls.append(("unregister", reason)),
                lambda **_: None,
                lambda **_: None,
                lambda *_: None,
                lambda *_: False,
                lambda *_: None,
                "center",
            )

        self.assertEqual(result, "window")
        self.assertIn(("sync",), calls)
        self.assertIn(("result", False, {"enabled": False, "auto_upload": False}), calls)
        self.assertNotIn(("register",), calls)
        self.assertFalse(any(item[0] == "unregister" for item in calls))

    def test_open_cloud_control_connect_callback_reports_save_failure(self):
        calls = []

        def fake_dialog(
            _parent,
            _state_provider,
            _status_var,
            _on_auto_upload_changed,
            _on_save,
            on_connect,
            _on_disconnect,
            _on_close,
            *_args,
            **_kwargs,
        ):
            calls.append(on_connect((True, "ws://host", 5, "secret", False), object()))
            return "window"

        with patch("sms_ui.cloud_control_window.open_cloud_control_window_dialog", side_effect=fake_dialog):
            result = open_cloud_control_window_runtime(
                "root",
                None,
                lambda: {"enabled": False, "auto_upload": False, "url": "ws://host", "secret": "secret", "reconnect_interval": 5},
                "status",
                lambda: None,
                lambda **_: None,
                lambda: (False, False, False),
                lambda: calls.append("register"),
                lambda _reason: calls.append("unregister"),
                lambda **_: calls.append("restart"),
                lambda **_: calls.append("stop"),
                lambda *_: None,
                lambda *_: False,
                lambda *_: None,
                "center",
            )

        self.assertEqual(result, "window")
        self.assertEqual(calls[0], False)
        self.assertNotIn("restart", calls)


if __name__ == "__main__":
    unittest.main()
