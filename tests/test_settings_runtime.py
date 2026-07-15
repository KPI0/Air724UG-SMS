import configparser
import unittest
from unittest.mock import patch

from sms_ui.settings_runtime import (
    call_filter_list_status,
    call_filter_mode_status,
    keyword_change_message,
    log_unmatched_status,
    open_call_filter_setting_runtime,
    open_keywords_setting_runtime,
    open_sms_font_dialog_runtime,
    open_voice_text_dialog_runtime,
    save_call_filter_list,
    save_call_filter_mode,
    save_keywords_config,
    save_log_unmatched_config,
    save_ui_config_values,
    toggle_popup_runtime,
    toggle_voice_broadcast_runtime,
)


class SettingsRuntimeTests(unittest.TestCase):
    def test_keyword_and_log_messages(self):
        self.assertEqual(keyword_change_message("add", value="otp"), "💬 关键词 增加：otp")
        self.assertEqual(keyword_change_message("delete", value="otp"), "💬 关键词 删除：otp")
        self.assertEqual(keyword_change_message("edit", value="new", old_value="old"), "💬 关键词 修改：old -> new")
        self.assertEqual(keyword_change_message("noop"), "")
        self.assertEqual(log_unmatched_status(True), "⚙️ 未匹配短信写入COM日志：已开启")
        self.assertEqual(log_unmatched_status(False), "⚙️ 未匹配短信写入COM日志：已关闭")

    def test_call_filter_messages(self):
        self.assertEqual(call_filter_mode_status("Disabled"), "📞 防骚扰模式已切换为：关闭过滤")
        self.assertEqual(call_filter_mode_status("Whitelist"), "📞 防骚扰模式已切换为：白名单模式")
        self.assertEqual(call_filter_mode_status("Blacklist"), "📞 防骚扰模式已切换为：黑名单模式")
        self.assertEqual(call_filter_list_status("whitelist", "add", value="10086"), "📞 白名单 增加：10086")
        self.assertEqual(call_filter_list_status("blacklist", "delete", value="10010"), "📞 黑名单 删除：10010")
        self.assertEqual(
            call_filter_list_status("blacklist", "edit", value="10011", old_value="10010"),
            "📞 黑名单 修改：10010 -> 10011",
        )
        self.assertEqual(call_filter_list_status("blacklist", "noop"), "")

    def test_save_settings_to_config(self):
        config = configparser.ConfigParser()
        saved = []

        save_keywords_config(config, ["otp", "bank"], lambda: saved.append("keywords"))
        save_log_unmatched_config(config, True, lambda: saved.append("log"))
        save_call_filter_mode(config, "Whitelist", lambda: saved.append("mode"))
        save_call_filter_list(config, "call_whitelist", ["10086"], lambda: saved.append("list"))

        self.assertEqual(config.get("ui", "keywords"), '["otp", "bank"]')
        self.assertEqual(config.get("ui", "log_unmatched_sms"), "1")
        self.assertEqual(config.get("ui", "call_filter_mode"), "Whitelist")
        self.assertEqual(config.get("ui", "call_whitelist"), '["10086"]')
        self.assertEqual(saved, ["keywords", "log", "mode", "list"])

    def test_save_ui_config_values_writes_multiple_values(self):
        config = configparser.ConfigParser()
        saved = []

        ok = save_ui_config_values(
            config,
            {"voice_text": "hello", "sms_font_size": 30},
            lambda: saved.append("saved"),
        )

        self.assertTrue(ok)
        self.assertEqual(config.get("ui", "voice_text"), "hello")
        self.assertEqual(config.get("ui", "sms_font_size"), "30")
        self.assertEqual(saved, ["saved"])

    def test_save_ui_config_values_reports_save_errors(self):
        config = configparser.ConfigParser()
        config["ui"] = {"voice_text": "old", "extra": "keep"}
        logs = []

        ok = save_ui_config_values(
            config,
            {"voice_text": "hello"},
            lambda: (_ for _ in ()).throw(RuntimeError("boom")),
            log_error=logs.append,
        )

        self.assertFalse(ok)
        self.assertEqual(dict(config.items("ui")), {"voice_text": "old", "extra": "keep"})
        self.assertEqual(len(logs), 1)
        self.assertIn("boom", logs[0])

    def test_save_ui_config_values_reports_false_save_result(self):
        config = configparser.ConfigParser()
        logs = []

        ok = save_ui_config_values(
            config,
            {"voice_text": "hello"},
            lambda: False,
            log_error=logs.append,
        )

        self.assertFalse(ok)
        self.assertFalse(config.has_section("ui"))
        self.assertEqual(len(logs), 1)
        self.assertIn("配置保存失败", logs[0])

    def test_save_helpers_log_save_errors(self):
        config = configparser.ConfigParser()
        logs = []

        def fail_save():
            raise RuntimeError("disk full")

        save_keywords_config(config, ["otp"], fail_save, log_error=logs.append)
        save_log_unmatched_config(config, True, fail_save, log_error=logs.append)
        save_call_filter_mode(config, "Whitelist", fail_save, log_error=logs.append)
        save_call_filter_list(config, "call_whitelist", ["10086"], fail_save, log_error=logs.append)

        self.assertEqual(len(logs), 4)
        self.assertTrue(all("disk full" in message for message in logs))

    def test_toggle_voice_broadcast_runtime_flips_state_and_saves(self):
        config = configparser.ConfigParser()
        calls = []

        result = toggle_voice_broadcast_runtime(
            False,
            config,
            lambda: calls.append("save"),
            lambda value: calls.append(("enabled", value)),
            lambda: calls.append("label"),
            lambda *args: calls.append(("ui", args)),
        )

        self.assertTrue(result)
        self.assertEqual(config.get("ui", "voice_enabled"), "1")
        self.assertEqual(calls[0], "save")
        self.assertEqual(calls[1], ("enabled", True))
        self.assertEqual(calls[2], "label")
        self.assertIn("语音播报", calls[-1][1][0])

    def test_toggle_voice_broadcast_runtime_logs_save_errors(self):
        config = configparser.ConfigParser()
        config["ui"] = {"voice_enabled": "0"}
        logs = []
        states = []
        labels = []

        result = toggle_voice_broadcast_runtime(
            False,
            config,
            lambda: (_ for _ in ()).throw(RuntimeError("save failed")),
            states.append,
            lambda: labels.append("label"),
            lambda *args: None,
            log_error=logs.append,
        )

        self.assertFalse(result)
        self.assertEqual(config.get("ui", "voice_enabled"), "0")
        self.assertEqual(states, [])
        self.assertEqual(labels, [])
        self.assertEqual(len(logs), 1)
        self.assertIn("save failed", logs[0])

    def test_toggle_popup_runtime_sets_state_and_saves(self):
        config = configparser.ConfigParser()
        calls = []

        result = toggle_popup_runtime(
            True,
            config,
            lambda: calls.append("save"),
            lambda value: calls.append(("popup", value)),
            lambda *args: calls.append(("ui", args)),
        )

        self.assertTrue(result)
        self.assertEqual(config.get("ui", "popup_enabled"), "1")
        self.assertEqual(calls[0], "save")
        self.assertEqual(calls[1], ("popup", True))
        self.assertIn("短信弹窗", calls[-1][1][0])

    def test_toggle_popup_runtime_does_not_commit_on_save_failure(self):
        config = configparser.ConfigParser()
        config["ui"] = {"popup_enabled": "0"}
        states = []

        result = toggle_popup_runtime(
            True,
            config,
            lambda: False,
            states.append,
            lambda *_args: None,
        )

        self.assertIsNone(result)
        self.assertEqual(config.get("ui", "popup_enabled"), "0")
        self.assertEqual(states, [])

    def test_open_voice_text_dialog_runtime_wires_preview_and_save(self):
        config = configparser.ConfigParser()
        calls = []

        def open_dialog(parent, current_text, preview, save, center_window):
            calls.append(("open", parent, current_text, center_window))
            preview("preview text")
            save("saved text")

        open_voice_text_dialog_runtime(
            "root",
            "current",
            config=config,
            safe_save=lambda: calls.append(("save_config",)),
            set_voice_text=lambda text: calls.append(("voice_text", text)),
            generate_voice=lambda **kwargs: calls.append(("generate", kwargs)),
            system_ui=lambda *args: calls.append(("ui", args)),
            center_window="center",
            open_dialog=open_dialog,
        )

        self.assertEqual(calls[0], ("open", "root", "current", "center"))
        self.assertEqual(calls[1], ("generate", {"force": True, "text": "preview text", "play_after": True}))
        self.assertEqual(calls[2], ("save_config",))
        self.assertEqual(config.get("ui", "voice_text"), "saved text")
        self.assertEqual(calls[3], ("voice_text", "saved text"))
        self.assertEqual(calls[4], ("generate", {"force": True}))
        self.assertIn("saved text", calls[5][1][0])

    def test_open_voice_text_dialog_runtime_does_not_commit_on_save_failure(self):
        config = configparser.ConfigParser()
        config["ui"] = {"voice_text": "old"}
        calls = []

        def open_dialog(_parent, _current, _preview, save, _center):
            self.assertFalse(save("new"))

        open_voice_text_dialog_runtime(
            "root",
            "old",
            config=config,
            safe_save=lambda: False,
            set_voice_text=lambda text: calls.append(("voice_text", text)),
            generate_voice=lambda **kwargs: calls.append(("generate", kwargs)),
            system_ui=lambda *args: calls.append(("ui", args)),
            center_window="center",
            open_dialog=open_dialog,
        )

        self.assertEqual(config.get("ui", "voice_text"), "old")
        self.assertFalse(any(item[0] in ("voice_text", "generate") for item in calls))

    def test_open_sms_font_dialog_runtime_wires_save(self):
        config = configparser.ConfigParser()
        calls = []

        def open_dialog(parent, current_size, current_color, save_font, center_window):
            calls.append(("open", parent, current_size, current_color, center_window))
            save_font(26, "#123456")

        open_sms_font_dialog_runtime(
            "root",
            30,
            "#ff0000",
            config=config,
            safe_save=lambda: calls.append(("save_config",)),
            set_font=lambda size, color: calls.append(("font", size, color)),
            apply_font_style=lambda: calls.append(("apply",)),
            system_ui=lambda *args: calls.append(("ui", args)),
            center_window="center",
            open_dialog=open_dialog,
        )

        self.assertEqual(calls[0], ("open", "root", 30, "#ff0000", "center"))
        self.assertEqual(calls[1], ("save_config",))
        self.assertEqual(config.get("ui", "sms_font_size"), "26")
        self.assertEqual(config.get("ui", "sms_font_color"), "#123456")
        self.assertEqual(calls[2], ("font", 26, "#123456"))
        self.assertEqual(calls[3], ("apply",))
        self.assertIn("#123456", calls[4][1][0])

    def test_open_sms_font_dialog_runtime_does_not_commit_on_save_failure(self):
        config = configparser.ConfigParser()
        config["ui"] = {"sms_font_size": "30", "sms_font_color": "#ff0000"}
        calls = []

        def open_dialog(_parent, _size, _color, save_font, _center):
            self.assertFalse(save_font(26, "#123456"))

        open_sms_font_dialog_runtime(
            "root",
            30,
            "#ff0000",
            config=config,
            safe_save=lambda: False,
            set_font=lambda size, color: calls.append(("font", size, color)),
            apply_font_style=lambda: calls.append(("apply",)),
            system_ui=lambda *args: calls.append(("ui", args)),
            center_window="center",
            open_dialog=open_dialog,
        )

        self.assertEqual(config.get("ui", "sms_font_size"), "30")
        self.assertEqual(config.get("ui", "sms_font_color"), "#ff0000")
        self.assertFalse(any(item[0] in ("font", "apply") for item in calls))

    def test_keywords_runtime_does_not_commit_log_state_on_save_failure(self):
        config = configparser.ConfigParser()
        config["ui"] = {"log_unmatched_sms": "0", "keywords": '["otp"]'}
        states = []
        callback_results = []

        def open_dialog(
            _parent,
            _keywords,
            _log_unmatched,
            on_keywords_changed,
            on_log_unmatched_changed,
            _center,
        ):
            callback_results.append(on_keywords_changed("add", value="bank"))
            callback_results.append(on_log_unmatched_changed(True))

        with patch(
            "sms_ui.settings_runtime.open_keywords_setting_dialog",
            side_effect=open_dialog,
        ):
            open_keywords_setting_runtime(
                "root",
                ["otp", "bank"],
                False,
                config,
                lambda: False,
                lambda *_args: None,
                states.append,
                "center",
            )

        self.assertEqual(callback_results, [False, False])
        self.assertEqual(config.get("ui", "keywords"), '["otp"]')
        self.assertEqual(config.get("ui", "log_unmatched_sms"), "0")
        self.assertEqual(states, [])

    def test_call_filter_runtime_does_not_commit_mode_on_save_failure(self):
        config = configparser.ConfigParser()
        config["ui"] = {"call_filter_mode": "Disabled"}
        modes = []
        callback_results = []

        def open_dialog(
            _parent,
            _mode,
            _whitelist,
            _blacklist,
            on_mode_changed,
            on_list_changed,
            _center,
        ):
            callback_results.append(on_mode_changed("Whitelist"))
            callback_results.append(
                on_list_changed("whitelist", "add", value="10086")
            )

        with patch(
            "sms_ui.settings_runtime.open_call_filter_setting_dialog",
            side_effect=open_dialog,
        ):
            open_call_filter_setting_runtime(
                "root",
                "Disabled",
                ["10086"],
                [],
                config,
                lambda: False,
                lambda *_args: None,
                modes.append,
                "center",
            )

        self.assertEqual(callback_results, [False, False])
        self.assertEqual(config.get("ui", "call_filter_mode"), "Disabled")
        self.assertFalse(config.has_option("ui", "call_whitelist"))
        self.assertEqual(modes, [])


if __name__ == "__main__":
    unittest.main()
