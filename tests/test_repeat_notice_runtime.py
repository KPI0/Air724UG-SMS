import unittest

from sms_ui.repeat_notice_runtime import (
    ConsecutiveRepeatNotice,
    TimedRepeatNotice,
    emit_repeat_notice,
)


class RepeatNoticeRuntimeTests(unittest.TestCase):
    def test_consecutive_notice_limits_repeated_messages(self):
        notice = ConsecutiveRepeatNotice(limit=4, suffix=" ignored")

        messages = [notice.next_message("same") for _ in range(5)]

        self.assertEqual(messages, ["same", "same", "same", "same ignored", None])

    def test_consecutive_notice_resets_when_message_changes(self):
        notice = ConsecutiveRepeatNotice(limit=3, suffix=" ignored")
        notice.next_message("first")
        notice.next_message("first")

        self.assertEqual(notice.next_message("second"), "second")
        self.assertEqual(notice.count, 1)

    def test_timed_notice_groups_by_repeat_key(self):
        now = [100.0]
        notice = TimedRepeatNotice(
            limit=3,
            reset_seconds=60.0,
            suffix=" ignored",
            monotonic=lambda: now[0],
        )

        first_key = [notice.next_message("first", repeat_key="A") for _ in range(3)]
        second_key = notice.next_message("second", repeat_key="B")

        self.assertEqual(first_key, ["first", "first", "first ignored"])
        self.assertEqual(second_key, "second")

    def test_timed_notice_resets_after_timeout(self):
        now = [100.0]
        notice = TimedRepeatNotice(
            limit=3,
            reset_seconds=60.0,
            suffix=" ignored",
            monotonic=lambda: now[0],
        )
        notice.next_message("same")
        notice.next_message("same")
        now[0] = 161.0

        self.assertEqual(notice.next_message("same"), "same")

    def test_emit_repeat_notice_falls_back_when_notice_fails(self):
        calls = []

        class BrokenNotice:
            def next_message(self, *_args, **_kwargs):
                raise RuntimeError("broken")

        emit_repeat_notice(
            BrokenNotice(),
            "fallback",
            lambda message, level: calls.append((message, level)),
        )

        self.assertEqual(calls, [("fallback", "normal")])

    def test_emit_repeat_notice_swallows_nested_fallback_errors(self):
        class BrokenNotice:
            def next_message(self, *_args, **_kwargs):
                raise RuntimeError("broken")

        emit_repeat_notice(
            BrokenNotice(),
            "fallback",
            lambda *_args: (_ for _ in ()).throw(RuntimeError("ui failed")),
        )


if __name__ == "__main__":
    unittest.main()
