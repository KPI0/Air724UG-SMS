import unittest

from sms_ui.main_status_runtime import (
    apply_sms_font_style_runtime,
    format_signal_text,
    update_label_status_runtime,
    update_signal_status_runtime,
    update_temperature_status_runtime,
)


class FakeVar:
    def __init__(self):
        self.value = None

    def set(self, value):
        self.value = value


class FakeLabel:
    def __init__(self, fail=False):
        self.fail = fail
        self.config_calls = []

    def config(self, **kwargs):
        if self.fail:
            raise RuntimeError("config failed")
        self.config_calls.append(kwargs)


class FakeText:
    def __init__(self, fail=False):
        self.fail = fail
        self.tag_calls = []

    def tag_config(self, tag, **kwargs):
        if self.fail:
            raise RuntimeError("tag failed")
        self.tag_calls.append((tag, kwargs))


class MainStatusRuntimeTests(unittest.TestCase):
    def test_format_signal_text(self):
        self.assertEqual(format_signal_text(100), "📶 -40 dBm")
        self.assertEqual(format_signal_text(255), "📶 未知")
        self.assertEqual(format_signal_text("bad"), "📶 -- dBm")

    def test_update_temperature_status_runtime_posts_update(self):
        temp_var = FakeVar()
        calls = []

        result = update_temperature_status_runtime(
            "31",
            tk_alive=lambda: True,
            temp_var=temp_var,
            run_on_ui_thread=lambda callback, ui_post: calls.append(ui_post) or callback(),
            ui_post="post",
        )

        self.assertTrue(result)
        self.assertEqual(calls, ["post"])
        self.assertEqual(temp_var.value, "🌡️ 31 ℃")

    def test_update_signal_status_runtime_skips_when_tk_dead(self):
        signal_var = FakeVar()

        result = update_signal_status_runtime(
            100,
            tk_alive=lambda: False,
            signal_var=signal_var,
            run_on_ui_thread=lambda callback, ui_post: callback(),
            ui_post=None,
        )

        self.assertFalse(result)
        self.assertIsNone(signal_var.value)

    def test_update_label_status_runtime_sets_text_and_color(self):
        text_var = FakeVar()
        label = FakeLabel()

        result = update_label_status_runtime(
            "ready",
            "green",
            tk_alive=lambda: True,
            text_var=text_var,
            label=label,
            run_on_ui_thread=lambda callback, ui_post: callback(),
            ui_post=None,
        )

        self.assertTrue(result)
        self.assertEqual(text_var.value, "ready")
        self.assertEqual(label.config_calls, [{"fg": "green"}])

    def test_apply_sms_font_style_runtime_reports_success_and_failure(self):
        text = FakeText()
        self.assertTrue(apply_sms_font_style_runtime(text, 24, "#123456"))
        self.assertEqual(text.tag_calls, [("sms", {"foreground": "#123456", "font": ("微软雅黑", 24)})])
        self.assertFalse(apply_sms_font_style_runtime(FakeText(fail=True), 24, "#123456"))


if __name__ == "__main__":
    unittest.main()
