import unittest

from sms_ui.serial_debug_sms_call_dialogs import format_sms_pdu_counter


class SerialDebugSmsCallDialogTests(unittest.TestCase):
    def test_format_sms_pdu_counter_allows_segmented_long_messages(self):
        text, too_long = format_sms_pdu_counter("A" * 100)

        self.assertEqual(text, "100 字 | UCS2 200 字节 | 2/255 段")
        self.assertFalse(too_long)

    def test_format_sms_pdu_counter_counts_emoji_by_ucs2_bytes(self):
        text, too_long = format_sms_pdu_counter("😀" * 36)

        self.assertEqual(text, "36 字 | UCS2 144 字节 | 2/255 段")
        self.assertFalse(too_long)

    def test_format_sms_pdu_counter_marks_over_segment_limit(self):
        text, too_long = format_sms_pdu_counter("A" * 17086)

        self.assertEqual(text, "17086 字 | UCS2 34172 字节 | 256/255 段")
        self.assertTrue(too_long)


if __name__ == "__main__":
    unittest.main()
