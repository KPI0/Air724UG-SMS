import unittest
from unittest.mock import patch

import sms_ui.app_ui_namespace_runtime as runtime


class FakeRoot:
    def __init__(self):
        self.calls = []

    def deiconify(self):
        self.calls.append(("deiconify",))

    def lift(self):
        self.calls.append(("lift",))

    def focus_force(self):
        self.calls.append(("focus",))

    def withdraw(self):
        self.calls.append(("withdraw",))


class BrokenRoot:
    def deiconify(self):
        raise RuntimeError("cannot show")

    def withdraw(self):
        raise RuntimeError("cannot hide")


class FakeWindow:
    def __init__(self, width=80, height=40, req_width=100, req_height=50):
        self.width = width
        self.height = height
        self.req_width = req_width
        self.req_height = req_height
        self.geometry_value = None
        self.updated = False

    def update_idletasks(self):
        self.updated = True

    def winfo_reqwidth(self):
        return self.req_width

    def winfo_reqheight(self):
        return self.req_height

    def winfo_screenwidth(self):
        return 1000

    def winfo_screenheight(self):
        return 600

    def winfo_width(self):
        return self.width

    def winfo_height(self):
        return self.height

    def geometry(self, value):
        self.geometry_value = value


class FakeParent:
    def winfo_rootx(self):
        return 10

    def winfo_rooty(self):
        return 20

    def winfo_width(self):
        return 300

    def winfo_height(self):
        return 200


class FakeMessageBox:
    def __init__(self, calls):
        self.calls = calls

    def askyesno(self, *args, **kwargs):
        self.calls.append(("ask", args, kwargs))
        return True

    def showwarning(self, *args, **kwargs):
        self.calls.append(("warning", args, kwargs))


class FakeTk:
    END = "END"


class AppUiNamespaceRuntimeTests(unittest.TestCase):
    def make_namespace(self):
        calls = []
        root = FakeRoot()
        return {
            "calls": calls,
            "root": root,
            "tray_icon": "icon",
            "resource_path": lambda path: f"res/{path}",
            "APP_WINDOW_TITLE": "SMS",
            "APP_DISPLAY_TITLE": "SMS 2",
            "show_window": lambda: calls.append(("show_window",)),
            "hide_window": lambda: calls.append(("hide_window",)),
            "cleanup_and_exit": lambda: calls.append(("exit",)),
            "run_on_ui_thread": lambda callback, ui_post: calls.append(("run", ui_post)) or callback(),
            "ui_post": "ui_post",
            "messagebox": FakeMessageBox(calls),
            "tk_alive": lambda: True,
            "temp_var": "temp_var",
            "signal_var": "signal_var",
            "cloud_var": "cloud_var",
            "cloud_label": "cloud_label",
            "status_var": "status_var",
            "status_label": "status_label",
            "_cloud_auth_status_from_ack": lambda data: data["status"],
            "set_cloud_status": lambda *args: calls.append(("cloud_status", args)),
            "text_area": "text_area",
            "SMS_FONT_SIZE": 18,
            "SMS_FONT_COLOR": "#123456",
            "tk": FakeTk,
            "send_command_with_result_async": lambda *args, **kwargs: calls.append(("send", args, kwargs)),
            "serial_lock": "lock",
            "serial_obj": "serial",
            "system_ui": lambda *args: calls.append(("system", args)),
            "log_file_only": lambda message: ("log", message),
        }

    def test_tray_and_window_namespace_runtimes_forward_state(self):
        namespace = self.make_namespace()

        with patch.object(runtime, "stop_tray_icon_runtime", return_value=True) as stop_runtime:
            self.assertTrue(runtime.stop_tray_icon_namespace_runtime(namespace, wait_after=0.1))
        self.assertEqual(stop_runtime.call_args.kwargs["tray_icon"], "icon")
        self.assertEqual(stop_runtime.call_args.kwargs["wait_after"], 0.1)
        self.assertEqual(stop_runtime.call_args.kwargs["log_error"]("tray log"), ("log", "tray log"))
        stop_runtime.call_args.kwargs["clear_tray_icon"]()
        self.assertIsNone(namespace["tray_icon"])

        self.assertIsNone(runtime.show_window_namespace_runtime(namespace))
        self.assertIsNone(runtime.hide_window_namespace_runtime(namespace))
        self.assertEqual(namespace["root"].calls, [
            ("deiconify",),
            ("lift",),
            ("focus",),
            ("withdraw",),
        ])

        with patch.object(runtime, "create_tray_icon_runtime", return_value="new_icon") as create_runtime:
            self.assertEqual(runtime.create_tray_namespace_runtime(namespace), "new_icon")
        kwargs = create_runtime.call_args.kwargs
        self.assertEqual(kwargs["icon_path"], "res/icon.ico")
        self.assertEqual(kwargs["title"], "SMS 2")
        self.assertEqual(kwargs["log_error"]("create log"), ("log", "create log"))
        kwargs["set_tray_icon"]("stored")
        self.assertEqual(namespace["tray_icon"], "stored")
        kwargs["cleanup_and_exit"]()
        self.assertEqual(namespace["calls"][-2:], [("run", "ui_post"), ("exit",)])

    def test_window_namespace_runtimes_log_root_failures(self):
        logs = []
        namespace = self.make_namespace()
        namespace["root"] = BrokenRoot()
        namespace["log_file_only"] = logs.append

        self.assertIsNone(runtime.show_window_namespace_runtime(namespace))
        self.assertIsNone(runtime.hide_window_namespace_runtime(namespace))

        self.assertEqual(len(logs), 2)
        self.assertIn("cannot show", logs[0])
        self.assertIn("cannot hide", logs[1])

    def test_window_namespace_runtimes_ignore_logging_failures(self):
        namespace = self.make_namespace()
        namespace["root"] = BrokenRoot()
        namespace["log_file_only"] = lambda _message: (_ for _ in ()).throw(RuntimeError("log down"))

        self.assertIsNone(runtime.show_window_namespace_runtime(namespace))
        self.assertIsNone(runtime.hide_window_namespace_runtime(namespace))

    def test_center_helpers_compute_geometry(self):
        namespace = self.make_namespace()
        win = FakeWindow()

        runtime.center_on_screen_namespace_runtime(namespace, win, 200, 100)
        self.assertTrue(win.updated)
        self.assertEqual(win.geometry_value, "200x100+400+250")

        child = FakeWindow(width=1, height=1, req_width=100, req_height=50)
        runtime.center_window_namespace_runtime(namespace, child, FakeParent())
        self.assertEqual(child.geometry_value, "+110+95")

    def test_messagebox_and_status_namespace_runtimes_forward_dependencies(self):
        namespace = self.make_namespace()

        with patch.object(runtime, "ui_messagebox_runtime", return_value="info") as messagebox_runtime:
            self.assertEqual(runtime.ui_messagebox_namespace_runtime(namespace, "info", "T", "M"), "info")
        self.assertIs(messagebox_runtime.call_args.kwargs["messagebox"], namespace["messagebox"])

        with patch.object(runtime, "update_temperature_status_runtime", return_value=True) as temp_runtime:
            self.assertTrue(runtime.set_temperature_namespace_runtime(namespace, "31"))
        self.assertEqual(temp_runtime.call_args.args, ("31",))
        self.assertEqual(temp_runtime.call_args.kwargs["temp_var"], "temp_var")

        with patch.object(runtime, "update_signal_status_runtime", return_value=True) as signal_runtime:
            self.assertTrue(runtime.set_signal_namespace_runtime(namespace, 100))
        self.assertEqual(signal_runtime.call_args.kwargs["signal_var"], "signal_var")

        with patch.object(runtime, "update_label_status_runtime", return_value=True) as label_runtime:
            self.assertTrue(runtime.set_cloud_status_namespace_runtime(namespace, "cloud", "green"))
            self.assertTrue(runtime.set_status_namespace_runtime(namespace, "ready", "black"))
        self.assertEqual(label_runtime.call_args_list[0].kwargs["text_var"], "cloud_var")
        self.assertEqual(label_runtime.call_args_list[1].kwargs["text_var"], "status_var")

    def test_cloud_auth_status_selects_status_text(self):
        namespace = self.make_namespace()

        runtime.set_cloud_auth_status_from_ack_namespace_runtime(namespace, {"status": "authorized"})
        runtime.set_cloud_auth_status_from_ack_namespace_runtime(namespace, {"status": "failed"})
        runtime.set_cloud_auth_status_from_ack_namespace_runtime(namespace, {"status": "pending"})

        self.assertEqual([call[1][1] for call in namespace["calls"] if call[0] == "cloud_status"], [
            "#008000",
            "#cc0000",
            "#b26a00",
        ])

    def test_font_clear_and_reset_namespace_runtimes_forward_state(self):
        namespace = self.make_namespace()

        with patch.object(runtime, "apply_sms_font_style_runtime", return_value=True) as font_runtime:
            self.assertTrue(runtime.apply_sms_font_style_namespace_runtime(namespace))
        self.assertEqual(font_runtime.call_args.args, ("text_area", 18, "#123456"))

        with patch.object(runtime, "clear_text_widget_runtime", return_value=True) as clear_runtime:
            self.assertTrue(runtime.clear_window_namespace_runtime(namespace))
        self.assertEqual(clear_runtime.call_args.kwargs["end"], "END")

        with patch.object(runtime, "send_reset_command_runtime", return_value="submitted") as reset_runtime:
            self.assertEqual(runtime.send_reset_cmd_namespace_runtime(namespace), "submitted")
        kwargs = reset_runtime.call_args.kwargs
        self.assertTrue(kwargs["confirm_reset"]())
        self.assertIs(kwargs["serial_lock"], namespace["serial_lock"])
        self.assertEqual(kwargs["get_serial"](), "serial")
        kwargs["show_warning"]("T", "M")
        self.assertEqual(namespace["calls"][-1][0], "warning")


if __name__ == "__main__":
    unittest.main()
