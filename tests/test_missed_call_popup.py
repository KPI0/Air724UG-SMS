import unittest
from datetime import datetime

from sms_ui.missed_call_popup import (
    additional_missed_call_notice,
    format_call_time,
)


class MissedCallPopupTests(unittest.TestCase):
    def test_additional_notice_is_hidden_for_one_call(self):
        self.assertEqual(additional_missed_call_notice(1), "")
        self.assertEqual(additional_missed_call_notice(None), "")

    def test_additional_notice_reports_calls_already_in_main_window(self):
        self.assertEqual(
            additional_missed_call_notice(2),
            "另有 1 个未接来电已显示在主窗口",
        )
        self.assertEqual(
            additional_missed_call_notice(6),
            "另有 5 个未接来电已显示在主窗口",
        )

    def test_call_time_uses_stable_display_format(self):
        self.assertEqual(
            format_call_time(datetime(2026, 8, 1, 10, 30, 45)),
            "2026-08-01 10:30:45",
        )


if __name__ == "__main__":
    unittest.main()
