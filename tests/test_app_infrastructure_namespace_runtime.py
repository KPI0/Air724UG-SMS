import datetime
import unittest
from unittest.mock import patch

import sms_ui.app_infrastructure_namespace_runtime as runtime


class FakeRoot:
    def __init__(self):
        self.calls = []

    def after(self, *args):
        self.calls.append(("after", args))
        return ("after", args)


class FakeMessageBox:
    def __init__(self, calls):
        self.calls = calls

    def showinfo(self, *args):
        self.calls.append(("info", args))


class FakeTime:
    @staticmethod
    def monotonic():
        return 12.5


class FakeDateTime:
    @staticmethod
    def now():
        return datetime.datetime(2026, 6, 8, 23, 59, 0)


class AppInfrastructureNamespaceRuntimeTests(unittest.TestCase):
    def make_namespace(self):
        calls = []
        return {
            "calls": calls,
            "config": "config",
            "CONFIG_FILE": "config.ini",
            "CONFIG_LOCK": "lock",
            "APP_DIR": "E:\\sms-client",
            "log_file_only": lambda message: calls.append(("log", message)),
            "serial_lock": "serial_lock",
            "serial_obj": "serial",
            "serial_connection_generation": 4,
            "unlock_port_mutex": lambda: calls.append(("unlock",)),
            "APP_START_MONO": 10.0,
            "START_UI_DELAY": 2.0,
            "time": FakeTime,
            "root": FakeRoot(),
            "run_on_ui_thread": lambda callback, ui_post: calls.append(("run", ui_post)) or callback(),
            "ui_post": "ui_post",
            "current_port_mutex": "old_mutex",
            "ALLOW_MULTI_INSTANCE": False,
            "APP_WINDOW_TITLE": "SMS",
            "app_mutex": None,
            "POPUP_ENABLED": True,
            "messagebox": FakeMessageBox(calls),
            "sms_popup_win": None,
            "center_on_screen": lambda win: calls.append(("center", win)),
            "show_window": lambda: calls.append(("show_window",)),
            "tk_alive": lambda: True,
            "clear_text_area_for_new_day": lambda: calls.append(("clear_day",)),
            "datetime": FakeDateTime,
        }

    def test_safe_save_config_namespace_runtime_forwards_config_context(self):
        namespace = self.make_namespace()
        defaults = {"ui": {"voice_enabled": "1"}}

        with patch.object(runtime, "safe_save_config_runtime", return_value=True) as save_runtime:
            self.assertTrue(
                runtime.safe_save_config_namespace_runtime(
                    namespace,
                    defaults_by_section=defaults,
                )
            )

        kwargs = save_runtime.call_args.kwargs
        self.assertEqual(kwargs["config"], "config")
        self.assertEqual(kwargs["config_file"], "config.ini")
        self.assertEqual(kwargs["config_lock"], "lock")
        self.assertIs(kwargs["defaults_by_section"], defaults)
        kwargs["log_error"]("bad")
        self.assertIn(("log", "bad"), namespace["calls"])

    def test_safe_close_serial_namespace_runtime_updates_serial(self):
        namespace = self.make_namespace()
        namespace["SMS_SEND_COORDINATOR"] = type(
            "Coordinator",
            (),
            {"cancel_active": lambda _self, reason: namespace["calls"].append(("cancel_sms", reason))},
        )()

        def close_runtime(serial_lock, get_serial, set_serial, unlock):
            self.assertEqual(serial_lock, "serial_lock")
            self.assertEqual(get_serial(), "serial")
            set_serial(None)
            unlock()
            return "closed"

        with patch.object(runtime, "safe_close_serial_runtime", side_effect=close_runtime):
            self.assertEqual(runtime.safe_close_serial_namespace_runtime(namespace), "closed")

        self.assertIsNone(namespace["serial_obj"])
        self.assertEqual(namespace["serial_connection_generation"], 5)
        self.assertEqual(namespace["calls"][0][0], "cancel_sms")
        self.assertIn(("unlock",), namespace["calls"])

    def test_schedule_delayed_ui_namespace_runtime_forwards_timing(self):
        namespace = self.make_namespace()
        calls = []

        result = runtime.schedule_delayed_ui_namespace_runtime(
            namespace,
            lambda: calls.append("callback"),
        )

        self.assertEqual(result, None)
        self.assertIn(("run", "ui_post"), namespace["calls"])
        self.assertEqual(calls, ["callback"])

    def test_port_mutex_namespace_runtime_locks_and_unlocks(self):
        namespace = self.make_namespace()

        with patch.object(runtime, "close_windows_handle") as close_handle, \
                patch.object(runtime, "create_named_mutex", return_value="new_mutex") as create_mutex:
            runtime.unlock_port_mutex_namespace_runtime(namespace)
            self.assertIsNone(namespace["current_port_mutex"])
            close_handle.assert_called_once_with("old_mutex")

            result = runtime.lock_port_mutex_namespace_runtime(namespace, "COM5")

        self.assertEqual(result, "new_mutex")
        self.assertEqual(namespace["current_port_mutex"], "new_mutex")
        create_mutex.assert_called_once_with("Air724UG_PORT_COM5")

    def test_unlock_port_mutex_logs_close_failure_and_clears_state(self):
        namespace = self.make_namespace()

        with patch.object(runtime, "close_windows_handle", side_effect=RuntimeError("close failed")):
            runtime.unlock_port_mutex_namespace_runtime(namespace)

        self.assertIsNone(namespace["current_port_mutex"])
        self.assertEqual(namespace["calls"][0][0], "log")
        self.assertIn("close failed", namespace["calls"][0][1])

    def test_check_single_instance_namespace_runtime_stores_mutex(self):
        namespace = self.make_namespace()

        with patch.object(runtime, "check_single_instance_app_runtime", return_value="mutex") as check_runtime:
            self.assertEqual(runtime.check_single_instance_namespace_runtime(namespace), "mutex")

        self.assertEqual(namespace["app_mutex"], "mutex")
        self.assertFalse(check_runtime.call_args.kwargs["allow_multi_instance"])
        self.assertEqual(check_runtime.call_args.kwargs["window_title"], "SMS")
        self.assertEqual(check_runtime.call_args.kwargs["app_dir"], "E:\\sms-client")
        self.assertIs(check_runtime.call_args.kwargs["log_error"], namespace["log_file_only"])

    def test_show_sms_popup_namespace_runtime_posts_popup(self):
        namespace = self.make_namespace()

        with patch.object(runtime, "show_sms_popup_runtime", return_value="shown") as popup_runtime:
            self.assertEqual(runtime.show_sms_popup_namespace_runtime(namespace, "msg"), "shown")

        self.assertIn(("run", "ui_post"), namespace["calls"])
        kwargs = popup_runtime.call_args.kwargs
        self.assertTrue(kwargs["popup_enabled"])
        self.assertIs(kwargs["parent"], namespace["root"])
        self.assertIsNone(kwargs["current_popup"])
        kwargs["set_popup"]("popup")
        kwargs["center_on_screen"]("popup")
        kwargs["show_window"]()
        self.assertEqual(namespace["sms_popup_win"], "popup")
        self.assertIn(("center", "popup"), namespace["calls"])
        self.assertIn(("show_window",), namespace["calls"])

    def test_schedule_next_midnight_clear_namespace_runtime_forwards_datetime(self):
        namespace = self.make_namespace()

        with patch.object(runtime, "schedule_next_midnight_clear_runtime", return_value=60000) as schedule_runtime:
            self.assertEqual(runtime.schedule_next_midnight_clear_namespace_runtime(namespace), 60000)

        self.assertIn(("run", "ui_post"), namespace["calls"])
        kwargs = schedule_runtime.call_args.kwargs
        self.assertTrue(kwargs["tk_alive"]())
        self.assertEqual(kwargs["now_func"](), datetime.datetime(2026, 6, 8, 23, 59, 0))
        kwargs["clear_callback"]()
        self.assertIn(("clear_day",), namespace["calls"])


if __name__ == "__main__":
    unittest.main()
