import unittest

from sms_ui.sms_popup import (
    _pack_message_viewport,
    _pack_popup_sections,
    additional_message_notice,
    display_line_total,
    estimate_message_lines,
    initial_popup_size,
    message_viewport,
)


class RecordingWidget:
    def __init__(self, name, events):
        self.name = name
        self.events = events

    def pack(self, **kwargs):
        self.events.append((self.name, "pack", kwargs))

    def pack_propagate(self, enabled):
        self.events.append((self.name, "pack_propagate", enabled))


class SmsPopupTests(unittest.TestCase):
    def test_additional_message_notice_is_hidden_for_one_message(self):
        self.assertEqual(additional_message_notice(1), "")
        self.assertEqual(additional_message_notice(None), "")

    def test_additional_message_notice_reports_messages_in_main_window(self):
        self.assertEqual(
            additional_message_notice(2),
            "另有 1 条短信已显示在主窗口",
        )
        self.assertEqual(
            additional_message_notice(6),
            "另有 5 条短信已显示在主窗口",
        )

    def test_scrollbar_keeps_priority_at_the_right_of_the_message(self):
        events = []
        message_text = RecordingWidget("message_text", events)
        message_scrollbar = RecordingWidget("message_scrollbar", events)

        _pack_message_viewport(message_text, message_scrollbar)

        self.assertEqual(
            events,
            [
                (
                    "message_scrollbar",
                    "pack",
                    {"side": "right", "fill": "y", "padx": (8, 0)},
                ),
                (
                    "message_text",
                    "pack",
                    {"side": "left", "fill": "both", "expand": True},
                ),
            ],
        )

    def test_footer_keeps_priority_and_confirmation_button_is_centered(self):
        events = []
        body = RecordingWidget("body", events)
        footer = RecordingWidget("footer", events)
        close_button = RecordingWidget("close_button", events)

        _pack_popup_sections(body, footer, close_button)

        self.assertEqual(
            events,
            [
                ("footer", "pack", {"fill": "x", "side": "bottom"}),
                ("footer", "pack_propagate", False),
                ("close_button", "pack", {"pady": 13}),
                ("body", "pack", {"fill": "both", "expand": True}),
            ],
        )

    def test_short_messages_keep_native_like_minimum_height(self):
        self.assertEqual(message_viewport(1), (2, False))
        self.assertEqual(message_viewport(2), (2, False))

    def test_regular_messages_expand_without_height_cap(self):
        self.assertEqual(message_viewport(8), (8, False))
        self.assertEqual(message_viewport(10), (10, False))

    def test_long_messages_are_capped(self):
        self.assertEqual(message_viewport(11), (10, True))
        self.assertEqual(message_viewport(100), (10, True))

    def test_bad_line_counts_fall_back_safely(self):
        self.assertEqual(message_viewport(None), (2, False))
        self.assertEqual(message_viewport("bad"), (2, False))

    def test_tk_display_line_boundaries_include_the_starting_line(self):
        self.assertEqual(display_line_total(None), 1)
        self.assertEqual(display_line_total((0,)), 1)
        self.assertEqual(display_line_total((1,)), 2)
        self.assertEqual(display_line_total((2,)), 3)

    def test_line_estimate_accounts_for_chinese_and_explicit_newlines(self):
        self.assertEqual(estimate_message_lines("short message"), 1)
        self.assertEqual(estimate_message_lines("\u4e2d" * 31), 2)
        self.assertEqual(estimate_message_lines("one\ntwo\nthree"), 3)

    def test_verification_sms_reserves_its_wrapped_final_line(self):
        message = (
            "\u3010\u4e2d\u56fd\u7535\u4fe1\u3011\u9a8c\u8bc1\u7801418735\uff0c3\u5206\u949f\u5185\u6709\u6548\u3002"
            "\u60a8\u6b63\u5728\u4e2d\u56fd\u7535\u4fe1APP\u67e5\u8be2\u6570\u636e\u8be6\u5355\u4e1a\u52a1\uff0c"
            "\u5207\u52ff\u5c06\u9a8c\u8bc1\u7801\u6cc4\u9732\u4e8e\u7ed9\u4ed6\u4eba\u3002"
        )
        self.assertGreaterEqual(estimate_message_lines(message), 3)

    def test_initial_size_does_not_expose_a_partial_extra_line(self):
        self.assertEqual(initial_popup_size(500, 170), (536, 180))
        self.assertEqual(initial_popup_size(620, 240), (620, 240))


if __name__ == "__main__":
    unittest.main()
