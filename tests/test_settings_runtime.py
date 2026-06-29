import configparser
import unittest

from sms_ui.settings_runtime import (
    call_filter_list_status,
    call_filter_mode_status,
    keyword_change_message,
    log_unmatched_status,
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
        logs = []

        ok = save_ui_config_values(
            config,
            {"voice_text": "hello"},
            lambda: (_ for _ in ()).throw(RuntimeError("boom")),
            log_error=logs.append,
        )

        self.assertFalse(ok)
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
        self.assertEqual(calls[0], ("enabled", True))
        self.assertEqual(calls[1], "label")
        self.assertIn("语音播报", calls[-1][1][0])

    def test_toggle_voice_broadcast_runtime_logs_save_errors(self):
        config = configparser.ConfigParser()
        logs = []

        result = toggle_voice_broadcast_runtime(
            False,
            config,
            lambda: (_ for _ in ()).throw(RuntimeError("save failed")),
            lambda value: None,
            lambda: None,
            lambda *args: None,
            log_error=logs.append,
        )

        self.assertTrue(result)
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
        self.assertEqual(calls[0], ("popup", True))
        self.assertIn("短信弹窗", calls[-1][1][0])

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
        self.assertEqual(calls[2], ("voice_text", "saved text"))
        self.assertEqual(config.get("ui", "voice_text"), "saved text")
        self.assertEqual(calls[3], ("save_config",))
        self.assertEqual(calls[4], ("generate", {"force": True}))
        self.assertIn("saved text", calls[5][1][0])

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
        self.assertEqual(calls[1], ("font", 26, "#123456"))
        self.assertEqual(config.get("ui", "sms_font_size"), "26")
        self.assertEqual(config.get("ui", "sms_font_color"), "#123456")
        self.assertEqual(calls[2], ("save_config",))
        self.assertEqual(calls[3], ("apply",))
        self.assertIn("#123456", calls[4][1][0])


if __name__ == "__main__":
    unittest.main()
