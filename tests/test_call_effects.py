import unittest
from types import SimpleNamespace

from sms_core.call_effects import (
    apply_call_answer_result,
    apply_call_decision,
    apply_call_hangup_result,
    apply_ring_timeout_expired,
)
from sms_core.call_events import CallLineDecision, CallState


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
            incoming_number="+8613812345678",
            show_popup_number="+8613812345678",
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
        self.assertIn(("ui", "📞 收到来电：来自 +8613812345678", "normal"), calls)
        self.assertIn(("status", "🔔 响铃中：+8613812345678", "blue"), calls)
        self.assertIn(("popup", "+8613812345678"), calls)
        self.assertNotIn(("hangup",), calls)

    def test_apply_call_decision_blocks_and_stops_processing(self):
        calls = []
        decision = CallLineDecision(
            state=CallState(),
            push_message="blocked",
            blocked_number="+8613812345678",
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
            ("ui", "🚫 防骚扰拦截：拒接 +8613812345678 (黑名单)", "warning"),
            calls,
        )
        self.assertIn(("hangup",), calls)
        self.assertNotIn(("popup", "+8613812345678"), calls)

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


if __name__ == "__main__":
    unittest.main()
