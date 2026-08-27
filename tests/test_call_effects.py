import unittest
from types import SimpleNamespace

from sms_core.call_effects import (
    apply_call_answer_result,
    apply_call_decision,
    apply_call_hangup_result,
    apply_ring_timeout_expired,
)
from sms_core.call_events import CallLineDecision, CallState
from sms_core.call_session import IncomingCallStartResult


class CallEffectTests(unittest.TestCase):
    def test_apply_call_answer_result_success(self):
        calls = []

        ok = apply_call_answer_result(
            SimpleNamespace(ok=True, error=""),
            "10086",
            lambda: calls.append(("restore",)),
            lambda: calls.append(("connected",)),
            lambda text, level: calls.append(("ui", text, level)),
            lambda text, color: calls.append(("status", text, color)),
            lambda func: calls.append(("post", func())),
            lambda value: calls.append(("timeout", value)),
        )

        self.assertTrue(ok)
        self.assertIn(("ui", "📞 已发送接听指令 (ATA)", "normal"), calls)
        self.assertIn(("status", "📞 通话中：10086", "blue"), calls)
        self.assertIn(("timeout", -1.0), calls)
        self.assertIn(("post", None), calls)
        self.assertIn(("connected",), calls)
        self.assertNotIn(("restore",), calls)

    def test_apply_call_answer_result_failure(self):
        calls = []

        ok = apply_call_answer_result(
            SimpleNamespace(ok=False, error="closed"),
            "10086",
            lambda: calls.append(("restore",)),
            lambda: calls.append(("connected",)),
            lambda text, level: calls.append(("ui", text, level)),
            lambda text, color: calls.append(("status", text, color)),
            lambda func: calls.append(("post", func())),
            lambda value: calls.append(("timeout", value)),
        )

        self.assertFalse(ok)
        self.assertIn(("ui", "📞 接听失败：closed", "warning"), calls)
        self.assertIn(("restore",), calls)
        self.assertNotIn(("connected",), calls)

    def test_apply_call_hangup_result(self):
        calls = []

        ok = apply_call_hangup_result(
            SimpleNamespace(ok=True, error=""),
            lambda: calls.append(("restore",)),
            lambda text, level: calls.append(("ui", text, level)),
            lambda func: calls.append(("post", func())),
            lambda: calls.append(("close",)),
        )

        self.assertTrue(ok)
        self.assertEqual(calls, [
            ("ui", "📞 已发送挂机指令 (ATH)", "normal"),
            ("close",),
        ])

        calls = []
        ok = apply_call_hangup_result(
            SimpleNamespace(ok=False, error="closed"),
            lambda: calls.append(("restore",)),
            lambda text, level: calls.append(("ui", text, level)),
            lambda func: calls.append(("post", func())),
            lambda: calls.append(("close",)),
        )

        self.assertFalse(ok)
        self.assertIn(("ui", "📞 挂断失败：closed", "warning"), calls)
        self.assertIn(("restore",), calls)
        self.assertNotIn(("close",), calls)

    def test_apply_ring_timeout_expired_updates_ui(self):
        calls = []

        apply_ring_timeout_expired(
            "COM5",
            lambda text, level: calls.append(("ui", text, level)),
            lambda text, color: calls.append(("status", text, color)),
            lambda: calls.append(("close",)),
        )

        self.assertEqual(calls, [
            ("ui", "📞 呼叫已取消或未接听", "normal"),
            ("status", "🟢 已连接：COM5", "green"),
            ("close",),
        ])

    def test_apply_call_decision_handles_incoming_call(self):
        calls = []
        decision = CallLineDecision(
            state=CallState(),
            push_message="incoming",
            incoming_number="+8613123123123",
            show_popup_number="+8613123123123",
        )

        result = apply_call_decision(
            decision,
            "COM5",
            lambda: calls.append(("hangup",)),
            lambda message, event_type: calls.append(("push", message, event_type)),
            lambda text, level: calls.append(("ui", text, level)),
            lambda text, color: calls.append(("status", text, color)),
            lambda number: calls.append(("popup", number)),
            lambda: calls.append(("close",)),
        )

        self.assertFalse(result.stop_processing)
        self.assertEqual(calls[0], ("push", "incoming", "call"))
        self.assertIn(("ui", "📞 收到来电：来自 +8613123123123", "normal"), calls)
        self.assertIn(("status", "🔔 响铃中：+8613123123123", "blue"), calls)
        self.assertIn(("popup", "+8613123123123"), calls)
        self.assertNotIn(("hangup",), calls)

    def test_apply_call_decision_enqueues_structured_cloud_call_event_once(self):
        calls = []
        decision = CallLineDecision(
            state=CallState(),
            push_message="blocked",
            blocked_number="10086",
            block_reason="blacklist",
            stop_processing=True,
        )

        apply_call_decision(
            decision,
            "COM5",
            lambda: None,
            lambda *_args, **_kwargs: None,
            lambda *_args: None,
            lambda *_args: None,
            lambda *_args: None,
            lambda: None,
            send_cloud_call_event=lambda *args, **kwargs: calls.append((args, kwargs)),
        )

        self.assertEqual(len(calls), 1)
        self.assertEqual(calls[0][0], ("10086", "blocked"))
        self.assertTrue(calls[0][1]["blocked"])
        self.assertEqual(calls[0][1]["block_reason"], "blacklist")

    def test_apply_call_decision_passes_call_template_variables(self):
        calls = []
        decision = CallLineDecision(
            state=CallState(),
            push_message="incoming",
            incoming_number="+8613323312312",
            show_popup_number="+8613323312312",
        )

        apply_call_decision(
            decision,
            "COM5",
            lambda: calls.append(("hangup",)),
            lambda message, **kwargs: calls.append(("push", message, kwargs)),
            lambda text, level: calls.append(("ui", text, level)),
            lambda text, color: calls.append(("status", text, color)),
            lambda number: calls.append(("popup", number)),
            lambda: calls.append(("close",)),
        )

        push = calls[0]
        self.assertEqual(push[0], "push")
        self.assertEqual(push[1], "incoming")
        self.assertEqual(push[2]["event_type"], "call")
        self.assertEqual(push[2]["variables"]["caller"], "+8613323312312")
        self.assertEqual(push[2]["variables"]["phone"], "+8613323312312")

    def test_apply_call_decision_blocks_and_stops_processing(self):
        calls = []
        decision = CallLineDecision(
            state=CallState(),
            push_message="blocked",
            blocked_number="+8613123123123",
            block_reason="黑名单",
            stop_processing=True,
        )

        result = apply_call_decision(
            decision,
            "COM5",
            lambda: calls.append(("hangup",)),
            lambda message, event_type: calls.append(("push", message, event_type)),
            lambda text, level: calls.append(("ui", text, level)),
            lambda text, color: calls.append(("status", text, color)),
            lambda number: calls.append(("popup", number)),
            lambda: calls.append(("close",)),
        )

        self.assertTrue(result.stop_processing)
        self.assertEqual(calls[0], ("push", "blocked", "call"))
        self.assertIn(
            ("ui", "🚫 防骚扰拦截：拒接 +8613123123123 (黑名单)", "warning"),
            calls,
        )
        self.assertIn(("hangup",), calls)
        self.assertNotIn(("popup", "+8613123123123"), calls)

    def test_apply_call_decision_handles_hangup_and_connected(self):
        calls = []
        decision = CallLineDecision(
            state=CallState(),
            hangup_notify=True,
            connected_number="10086",
        )

        result = apply_call_decision(
            decision,
            "COM5",
            lambda: calls.append(("hangup",)),
            lambda message, event_type: calls.append(("push", message, event_type)),
            lambda text, level: calls.append(("ui", text, level)),
            lambda text, color: calls.append(("status", text, color)),
            lambda number: calls.append(("popup", number)),
            lambda: calls.append(("close",)),
        )

        self.assertFalse(result.stop_processing)
        self.assertIn(("ui", "📞 语音通话已结束", "normal"), calls)
        self.assertIn(("status", "🟢 已连接：COM5", "green"), calls)
        self.assertIn(("close",), calls)
        self.assertIn(("ui", "📞 对方已接听：10086", "normal"), calls)
        self.assertIn(("status", "📞 通话中：10086", "blue"), calls)

    def test_duplicate_incoming_session_does_not_repeat_push_or_popup(self):
        calls = []
        decision = CallLineDecision(
            state=CallState(),
            push_message="incoming",
            incoming_number="10086",
            show_popup_number="10086",
        )

        apply_call_decision(
            decision,
            "COM5",
            lambda: calls.append(("hangup",)),
            lambda *args, **kwargs: calls.append(("push", args, kwargs)),
            lambda *args: calls.append(("ui", args)),
            lambda *args: calls.append(("status", args)),
            lambda number: calls.append(("popup", number)),
            lambda: calls.append(("close",)),
            start_incoming_call=lambda _number: False,
        )

        self.assertEqual(calls, [])

    def test_hangup_closes_active_popup_before_showing_missed_call(self):
        calls = []
        missed_call = SimpleNamespace(caller_num="10086")
        decision = CallLineDecision(state=CallState(), hangup_notify=True)

        apply_call_decision(
            decision,
            "COM5",
            lambda: None,
            lambda *args, **kwargs: None,
            lambda text, level: calls.append(("ui", text, level)),
            lambda text, color: calls.append(("status", text, color)),
            lambda number: None,
            lambda: calls.append(("close",)),
            finish_incoming_call=lambda: missed_call,
            show_missed_call_popup=lambda value: calls.append(("missed", value)),
        )

        self.assertIn(("ui", "📵 未接来电：10086", "warning"), calls)
        self.assertLess(calls.index(("close",)), calls.index(("missed", missed_call)))

    def test_debounced_hangup_still_cleans_up_without_duplicate_end_log(self):
        calls = []
        missed_call = SimpleNamespace(caller_num="10010")
        decision = CallLineDecision(state=CallState(), call_ended=True)

        apply_call_decision(
            decision,
            "COM5",
            lambda: None,
            lambda *args, **kwargs: None,
            lambda text, level: calls.append(("ui", text, level)),
            lambda text, color: calls.append(("status", text, color)),
            lambda number: None,
            lambda: calls.append(("close",)),
            finish_incoming_call=lambda: calls.append(("finish",)) or missed_call,
            show_missed_call_popup=lambda value: calls.append(("missed", value)),
        )

        self.assertIn(("finish",), calls)
        self.assertIn(("close",), calls)
        self.assertIn(("missed", missed_call), calls)
        self.assertNotIn(("ui", "📞 语音通话已结束", "normal"), calls)

    def test_different_incoming_caller_replaces_popup_and_reports_previous_missed(self):
        calls = []
        previous_missed = SimpleNamespace(caller_num="10086")
        decision = CallLineDecision(
            state=CallState(),
            push_message="incoming",
            incoming_number="10010",
            show_popup_number="10010",
        )

        apply_call_decision(
            decision,
            "COM5",
            lambda: None,
            lambda *args, **kwargs: calls.append(("push",)),
            lambda text, level: calls.append(("ui", text, level)),
            lambda text, color: calls.append(("status", text, color)),
            lambda number: calls.append(("popup", number)),
            lambda: calls.append(("close",)),
            start_incoming_call=lambda _number: IncomingCallStartResult(
                started=True,
                replaced=True,
                replaced_missed_call=previous_missed,
            ),
            show_missed_call_popup=lambda value: calls.append(("missed", value)),
        )

        self.assertLess(calls.index(("close",)), calls.index(("missed", previous_missed)))
        self.assertLess(calls.index(("missed", previous_missed)), calls.index(("popup", "10010")))
        self.assertIn(("push",), calls)

    def test_timeout_does_not_show_missed_call_after_user_handled_it(self):
        calls = []

        apply_ring_timeout_expired(
            "COM5",
            lambda text, level: calls.append(("ui", text, level)),
            lambda text, color: calls.append(("status", text, color)),
            lambda: calls.append(("close",)),
            finish_incoming_call=lambda: None,
            show_missed_call_popup=lambda value: calls.append(("missed", value)),
        )

        self.assertFalse(any(item[0] == "missed" for item in calls))


if __name__ == "__main__":
    unittest.main()
