import unittest
from unittest.mock import patch

import sms_ui.app_ui_namespace_bindings as bindings


class AppUiNamespaceBindingsTests(unittest.TestCase):
    def make_namespace(self):
        calls = []
        return {
            "calls": calls,
            "root": "root",
            "APP_VERSION": "3.6.6",
            "_ui_open_about_dialog": lambda *args: calls.append(("about", args)) or "about",
        }

    def test_install_registers_expected_names(self):
        namespace = self.make_namespace()

        result = bindings.install_app_ui_namespace_bindings(namespace)

        self.assertIs(result, namespace)
        for name in (
            "center_on_screen",
            "stop_tray_icon",
            "show_window",
            "hide_window",
            "create_tray",
            "center_window",
            "show_about",
            "ui_messagebox",
            "set_temperature",
            "set_signal",
            "set_cloud_status",
            "set_cloud_auth_status_from_ack",
            "set_status",
            "apply_sms_font_style",
            "clear_window",
            "send_reset_cmd",
        ):
            self.assertIn(name, namespace)
            self.assertTrue(callable(namespace[name]))

    def test_bindings_forward_namespace_and_arguments(self):
        namespace = self.make_namespace()
        bindings.install_app_ui_namespace_bindings(namespace)

        with patch.object(bindings, "center_on_screen_namespace_runtime", return_value="centered") as center_screen, \
                patch.object(bindings, "stop_tray_icon_namespace_runtime", return_value="stopped") as stop_tray, \
                patch.object(bindings, "show_window_namespace_runtime", return_value="shown") as show, \
                patch.object(bindings, "hide_window_namespace_runtime", return_value="hidden") as hide, \
                patch.object(bindings, "create_tray_namespace_runtime", return_value="tray") as tray, \
                patch.object(bindings, "center_window_namespace_runtime", return_value="center_window") as center_window, \
                patch.object(bindings, "ui_messagebox_namespace_runtime", return_value="message") as messagebox, \
                patch.object(bindings, "set_temperature_namespace_runtime", return_value="temp") as temp, \
                patch.object(bindings, "set_signal_namespace_runtime", return_value="signal") as signal, \
                patch.object(bindings, "set_cloud_status_namespace_runtime", return_value="cloud") as cloud, \
                patch.object(bindings, "set_cloud_auth_status_from_ack_namespace_runtime", return_value="auth") as auth, \
                patch.object(bindings, "set_status_namespace_runtime", return_value="status") as status, \
                patch.object(bindings, "apply_sms_font_style_namespace_runtime", return_value="font") as font, \
                patch.object(bindings, "clear_window_namespace_runtime", return_value="clear") as clear, \
                patch.object(bindings, "send_reset_cmd_namespace_runtime", return_value="reset") as reset:
            self.assertEqual(namespace["center_on_screen"]("win", 100, 50), "centered")
            self.assertEqual(namespace["stop_tray_icon"](wait_after=0.1), "stopped")
            self.assertEqual(namespace["show_window"](), "shown")
            self.assertEqual(namespace["hide_window"](), "hidden")
            self.assertEqual(namespace["create_tray"](), "tray")
            self.assertEqual(namespace["center_window"]("child", "parent"), "center_window")
            self.assertEqual(namespace["show_about"](), "about")
            self.assertEqual(namespace["ui_messagebox"]("info", "T", "M"), "message")
            self.assertEqual(namespace["set_temperature"]("31"), "temp")
            self.assertEqual(namespace["set_signal"](-80), "signal")
            self.assertEqual(namespace["set_cloud_status"]("ok", "green"), "cloud")
            self.assertEqual(namespace["set_cloud_auth_status_from_ack"]({"ok": True}), "auth")
            self.assertEqual(namespace["set_status"]("ready", "black"), "status")
            self.assertEqual(namespace["apply_sms_font_style"](), "font")
            self.assertEqual(namespace["clear_window"](), "clear")
            self.assertEqual(namespace["send_reset_cmd"](), "reset")

        center_screen.assert_called_once_with(namespace, "win", 100, 50)
        stop_tray.assert_called_once_with(namespace, wait_after=0.1)
        show.assert_called_once_with(namespace)
        hide.assert_called_once_with(namespace)
        tray.assert_called_once_with(namespace)
        center_window.assert_called_once_with(namespace, "child", "parent")
        self.assertEqual(namespace["calls"], [("about", ("root", "3.6.6", "https://github.com/KPI0/Air724UG-SMS", namespace["center_window"]))])
        messagebox.assert_called_once_with(namespace, "info", "T", "M")
        temp.assert_called_once_with(namespace, "31")
        signal.assert_called_once_with(namespace, -80)
        cloud.assert_called_once_with(namespace, "ok", "green")
        auth.assert_called_once_with(namespace, {"ok": True})
        status.assert_called_once_with(namespace, "ready", "black")
        font.assert_called_once_with(namespace)
        clear.assert_called_once_with(namespace)
        reset.assert_called_once_with(namespace)


if __name__ == "__main__":
    unittest.main()
