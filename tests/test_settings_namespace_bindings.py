import unittest
from unittest.mock import patch

import sms_ui.settings_namespace_bindings as bindings


class SettingsNamespaceBindingsTests(unittest.TestCase):
    def make_namespace(self):
        calls = []

        class Menu:
            def entryconfig(self, index, **kwargs):
                calls.append(("entryconfig", index, kwargs))

        return {
            "calls": calls,
            "config": "config",
            "VOICE_TEXT": "hello",
            "VOICE_ENABLED": True,
            "safe_save_config": lambda: calls.append(("save",)),
            "log_file_only": lambda message: calls.append(("log", message)),
            "menu_bar": Menu(),
            "voice_menu_index": 3,
        }

    def test_install_registers_expected_names(self):
        namespace = self.make_namespace()

        result = bindings.install_settings_namespace_bindings(namespace)

        self.assertIs(result, namespace)
        for name in (
            "save_voice_text_setting",
            "open_sms_font_dialog",
            "open_voice_text_dialog",
            "open_serial_setting",
            "open_desktop_shortcut_dialog",
            "open_keywords_setting",
            "open_call_filter_setting",
            "open_security_settings",
            "update_voice_menu_label",
            "save_voice_setting",
        ):
            self.assertIn(name, namespace)
            self.assertTrue(callable(namespace[name]))

    def test_dialog_bindings_forward_namespace(self):
        namespace = self.make_namespace()
        bindings.install_settings_namespace_bindings(namespace)

        with patch.object(bindings, "open_sms_font_dialog_namespace_runtime", return_value="font") as font, \
                patch.object(bindings, "open_voice_text_dialog_namespace_runtime", return_value="voice") as voice, \
                patch.object(bindings, "open_serial_setting_namespace_runtime", return_value="serial") as serial, \
                patch.object(bindings, "open_desktop_shortcut_dialog_namespace_runtime", return_value="shortcut") as shortcut, \
                patch.object(bindings, "open_keywords_setting_namespace_runtime", return_value="keywords") as keywords, \
                patch.object(bindings, "open_call_filter_setting_namespace_runtime", return_value="filter") as filter_dialog, \
                patch.object(bindings, "open_security_settings_namespace_runtime", return_value="security") as security:
            self.assertEqual(namespace["open_sms_font_dialog"](), "font")
            self.assertEqual(namespace["open_voice_text_dialog"](), "voice")
            self.assertEqual(namespace["open_serial_setting"](), "serial")
            self.assertEqual(namespace["open_desktop_shortcut_dialog"](), "shortcut")
            self.assertEqual(namespace["open_keywords_setting"](), "keywords")
            self.assertEqual(namespace["open_call_filter_setting"](), "filter")
            self.assertEqual(namespace["open_security_settings"](), "security")

        font.assert_called_once_with(namespace)
        voice.assert_called_once_with(namespace)
        serial.assert_called_once_with(namespace)
        shortcut.assert_called_once_with(namespace)
        keywords.assert_called_once_with(namespace)
        filter_dialog.assert_called_once_with(namespace)
        security.assert_called_once_with(namespace)

    def test_save_voice_settings_and_update_menu_label(self):
        namespace = self.make_namespace()
        bindings.install_settings_namespace_bindings(namespace)

        with patch.object(bindings, "save_ui_config_values", return_value="saved") as save_values:
            self.assertEqual(namespace["save_voice_text_setting"](), "saved")
            self.assertEqual(namespace["save_voice_setting"](), "saved")

        self.assertEqual(save_values.call_args_list[0].args, (
            "config",
            {"voice_text": "hello"},
            namespace["safe_save_config"],
        ))

        self.assertEqual(save_values.call_args_list[1].args, (
            "config",
            {"voice_enabled": "1"},
            namespace["safe_save_config"],
        ))
        self.assertIs(save_values.call_args_list[0].kwargs["log_error"], namespace["log_file_only"])
        self.assertIs(save_values.call_args_list[1].kwargs["log_error"], namespace["log_file_only"])

        namespace["update_voice_menu_label"]()
        namespace["VOICE_ENABLED"] = False
        namespace["update_voice_menu_label"]()

        self.assertIn(("entryconfig", 3, {"label": "🔊 语音播报"}), namespace["calls"])
        self.assertIn(("entryconfig", 3, {"label": "🔇 语音播报"}), namespace["calls"])

    def test_security_settings_binding_rejects_transient_parent(self):
        namespace = self.make_namespace()
        bindings.install_settings_namespace_bindings(namespace)

        with self.assertRaises(TypeError):
            namespace["open_security_settings"]("cloud_window")

    def test_update_voice_menu_label_logs_failures(self):
        namespace = self.make_namespace()

        class BrokenMenu:
            def entryconfig(self, index, **kwargs):
                raise RuntimeError("menu failed")

        namespace["menu_bar"] = BrokenMenu()
        bindings.install_settings_namespace_bindings(namespace)

        namespace["update_voice_menu_label"]()

        self.assertEqual(namespace["calls"][0][0], "log")
        self.assertIn("menu failed", namespace["calls"][0][1])


if __name__ == "__main__":
    unittest.main()
