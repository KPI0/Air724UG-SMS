import unittest

from sms_ui.settings_namespace_runtime import (
    open_call_filter_setting_namespace_runtime,
    open_desktop_shortcut_dialog_namespace_runtime,
    open_keywords_setting_namespace_runtime,
    open_security_settings_namespace_runtime,
    open_serial_setting_namespace_runtime,
    open_sms_font_dialog_namespace_runtime,
    open_voice_text_dialog_namespace_runtime,
)


class FakeWakeEvent:
    def set(self):
        return "wakeup"


class SettingsNamespaceRuntimeTests(unittest.TestCase):
    def base_namespace(self):
        return {
            "root": "root",
            "config": "config",
            "safe_save_config": lambda: "saved",
            "system_ui": lambda *args: ("ui", args),
            "center_window": "center",
            "SMS_FONT_SIZE": 14,
            "SMS_FONT_COLOR": "#112233",
            "VOICE_TEXT": "hello",
            "MODE": "Auto",
            "PORT": "COM1",
            "BAUD": 115200,
            "KEYWORDS": ["code"],
            "LOG_UNMATCHED_SMS": False,
            "CALL_FILTER_MODE": "Disabled",
            "CALL_WHITELIST": ["10086"],
            "CALL_BLACKLIST": ["10010"],
            "CLOUD_SENSITIVE_COMMAND_PERMISSIONS": {"sms": False},
            "_ui_open_sms_font_dialog": "font_dialog",
            "_ui_open_voice_text_dialog": "voice_dialog",
            "_ui_open_desktop_shortcut_dialog": "shortcut_dialog",
            "_ui_open_security_settings_dialog": "security_dialog",
            "apply_sms_font_style": lambda: "font",
            "generate_alert_voice": lambda **kwargs: ("voice", kwargs),
            "scan_com_ports_all": lambda: ["COM1"],
            "set_status": lambda *args: ("status", args),
            "safe_close_serial": lambda: "close",
            "serial_wakeup_event": FakeWakeEvent(),
            "create_desktop_shortcut": lambda name: ("shortcut", name),
            "log_file_only": lambda message: ("log", message),
            "register_config_sync_refresher": lambda group, callback: (
                "unregister",
                group,
                callback,
            ),
        }

    def test_open_sms_font_dialog_namespace_runtime_forwards_and_sets_font(self):
        namespace = self.base_namespace()
        calls = []

        result = open_sms_font_dialog_namespace_runtime(
            namespace,
            open_dialog_runtime=lambda parent, size, color, **kwargs: calls.append((parent, size, color, kwargs)) or "opened",
        )

        self.assertEqual(result, "opened")
        parent, size, color, forwarded = calls[0]
        self.assertEqual((parent, size, color), ("root", 14, "#112233"))
        self.assertEqual(forwarded["open_dialog"], "font_dialog")
        self.assertEqual(forwarded["log_error"]("font log"), ("log", "font log"))
        forwarded["set_font"](18, "#445566")
        self.assertEqual(namespace["SMS_FONT_SIZE"], 18)
        self.assertEqual(namespace["SMS_FONT_COLOR"], "#445566")

    def test_open_voice_text_dialog_namespace_runtime_forwards_and_sets_text(self):
        namespace = self.base_namespace()
        calls = []

        result = open_voice_text_dialog_namespace_runtime(
            namespace,
            open_dialog_runtime=lambda parent, current_text, **kwargs: calls.append((parent, current_text, kwargs)) or "voice",
        )

        self.assertEqual(result, "voice")
        parent, text, forwarded = calls[0]
        self.assertEqual((parent, text), ("root", "hello"))
        self.assertEqual(forwarded["open_dialog"], "voice_dialog")
        self.assertEqual(forwarded["log_error"]("voice log"), ("log", "voice log"))
        forwarded["set_voice_text"]("next")
        self.assertEqual(namespace["VOICE_TEXT"], "next")

    def test_open_security_settings_namespace_runtime_forwards_and_sets_flag(self):
        namespace = self.base_namespace()
        calls = []

        result = open_security_settings_namespace_runtime(
            namespace,
            open_setting_runtime=lambda parent, enabled, **kwargs: calls.append(
                (parent, enabled, kwargs)
            ) or "security",
        )

        self.assertEqual(result, "security")
        parent, permissions, forwarded = calls[0]
        self.assertEqual((parent, permissions), ("root", {"sms": False}))
        self.assertEqual(forwarded["open_dialog"], "security_dialog")
        forwarded["set_permissions"]({"sms": True})
        self.assertEqual(namespace["CLOUD_SENSITIVE_COMMAND_PERMISSIONS"], {"sms": True})

    def test_open_security_settings_namespace_runtime_uses_explicit_parent(self):
        namespace = self.base_namespace()
        calls = []

        result = open_security_settings_namespace_runtime(
            namespace,
            "cloud_window",
            open_setting_runtime=lambda parent, permissions, **kwargs: calls.append(
                (parent, permissions, kwargs)
            ) or "security",
        )

        self.assertEqual(result, "security")
        self.assertEqual(calls[0][0], "cloud_window")

    def test_open_serial_setting_namespace_runtime_forwards_and_sets_serial_state(self):
        namespace = self.base_namespace()
        calls = []

        result = open_serial_setting_namespace_runtime(
            namespace,
            open_setting_runtime=lambda parent, **kwargs: calls.append((parent, kwargs)) or "serial",
        )

        self.assertEqual(result, "serial")
        parent, forwarded = calls[0]
        self.assertEqual(parent, "root")
        self.assertEqual(forwarded["current_mode"], "Auto")
        self.assertEqual(forwarded["current_port"], "COM1")
        self.assertEqual(forwarded["current_baud"], 115200)
        self.assertEqual(forwarded["wake_serial"](), "wakeup")
        forwarded["set_serial_state"]("Manual", "COM7", 9600)
        self.assertEqual(namespace["MODE"], "Manual")
        self.assertEqual(namespace["PORT"], "COM7")
        self.assertEqual(namespace["BAUD"], 9600)

    def test_open_desktop_shortcut_dialog_namespace_runtime_forwards_dependencies(self):
        namespace = self.base_namespace()
        calls = []

        result = open_desktop_shortcut_dialog_namespace_runtime(
            namespace,
            open_dialog_runtime=lambda parent, **kwargs: calls.append((parent, kwargs)) or "shortcut",
        )

        self.assertEqual(result, "shortcut")
        parent, forwarded = calls[0]
        self.assertEqual(parent, "root")
        self.assertEqual(forwarded["open_dialog"], "shortcut_dialog")
        self.assertEqual(forwarded["create_shortcut"]("Desk"), ("shortcut", "Desk"))

    def test_open_keywords_setting_namespace_runtime_forwards_and_sets_flag(self):
        namespace = self.base_namespace()
        calls = []

        result = open_keywords_setting_namespace_runtime(
            namespace,
            open_setting_runtime=lambda *args, **kwargs: calls.append((args, kwargs)) or "keywords",
        )

        self.assertEqual(result, "keywords")
        args, kwargs = calls[0]
        self.assertEqual(args[:6], ("root", ["code"], False, "config", namespace["safe_save_config"], namespace["system_ui"]))
        args[6](True)
        self.assertEqual(args[8]("keyword log"), ("log", "keyword log"))
        self.assertTrue(namespace["LOG_UNMATCHED_SMS"])
        refresh = lambda: None
        self.assertEqual(
            kwargs["register_external_refresh"](refresh),
            ("unregister", "keywords", refresh),
        )
        self.assertTrue(kwargs["get_log_unmatched"]())

    def test_open_call_filter_setting_namespace_runtime_forwards_and_sets_mode(self):
        namespace = self.base_namespace()
        calls = []

        result = open_call_filter_setting_namespace_runtime(
            namespace,
            open_setting_runtime=lambda *args, **kwargs: calls.append((args, kwargs)) or "filter",
        )

        self.assertEqual(result, "filter")
        args, kwargs = calls[0]
        self.assertEqual(args[:6], (
            "root",
            "Disabled",
            ["10086"],
            ["10010"],
            "config",
            namespace["safe_save_config"],
        ))
        args[7]("Whitelist")
        self.assertEqual(args[9]("filter log"), ("log", "filter log"))
        self.assertEqual(namespace["CALL_FILTER_MODE"], "Whitelist")
        refresh = lambda: None
        self.assertEqual(
            kwargs["register_external_refresh"](refresh),
            ("unregister", "call_filter", refresh),
        )
        self.assertEqual(kwargs["get_mode"](), "Whitelist")


if __name__ == "__main__":
    unittest.main()
