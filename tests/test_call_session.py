import unittest
from datetime import datetime

from sms_core.call_session import IncomingCallSessionTracker


class IncomingCallSessionTrackerTests(unittest.TestCase):
    def test_unhandled_session_finishes_as_one_missed_call(self):
        started_at = datetime(2026, 8, 1, 10, 30, 0)
        tracker = IncomingCallSessionTracker(now_func=lambda: started_at)

        self.assertTrue(tracker.start("10086"))
        self.assertFalse(tracker.start("10086"))

        missed_call = tracker.finish()

        self.assertEqual(missed_call.caller_num, "10086")
        self.assertEqual(missed_call.started_at, started_at)
        self.assertIsNone(tracker.finish())
        self.assertEqual(tracker.snapshot().caller_num, "")

    def test_handled_session_does_not_finish_as_missed(self):
        tracker = IncomingCallSessionTracker()

        self.assertTrue(tracker.start("10010"))
        self.assertTrue(tracker.mark_handled())
        self.assertIsNone(tracker.finish())
        self.assertFalse(tracker.mark_handled())

    def test_reset_discards_incomplete_session(self):
        tracker = IncomingCallSessionTracker()
        tracker.start("10000")

        tracker.reset()

        self.assertIsNone(tracker.finish())
        self.assertFalse(tracker.snapshot().handled)

    def test_different_caller_replaces_session_and_returns_previous_missed_call(self):
        started_times = iter(("first", "second"))
        tracker = IncomingCallSessionTracker(now_func=lambda: next(started_times))

        first = tracker.start("10086")
        second = tracker.start("10010")

        self.assertTrue(first)
        self.assertFalse(first.replaced)
        self.assertTrue(second)
        self.assertTrue(second.replaced)
        self.assertEqual(second.replaced_missed_call.caller_num, "10086")
        self.assertEqual(second.replaced_missed_call.started_at, "first")
        self.assertEqual(tracker.snapshot().caller_num, "10010")
        self.assertEqual(tracker.snapshot().started_at, "second")

    def test_different_caller_does_not_report_handled_session_as_missed(self):
        tracker = IncomingCallSessionTracker()
        tracker.start("10086")
        tracker.mark_handled()

        result = tracker.start("10010")

        self.assertTrue(result)
        self.assertTrue(result.replaced)
        self.assertIsNone(result.replaced_missed_call)


if __name__ == "__main__":
    unittest.main()
