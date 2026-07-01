import unittest

from sms_core.call_events import (
    CallState,
    call_push_message,
    connected_call_number,
    handle_call_line,
    handle_clip_line,
    handle_hangup_line,
    is_call_active,
    refresh_ring_timeout,
    ring_timeout_expired,
)


class CallEventTests(unittest.TestCase):
    def test_handle_clip_line_allows_new_call(self):
        decision = handle_clip_line(
            '+CLIP: "+8613123123123",129',
            last_clip_num="",
            last_clip_time=0.0,
            now=10.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
        )

        self.assertEqual(decision.caller_num, "+8613123123123")
        self.assertFalse(decision.blocked)
        self.assertTrue(decision.new_clip)
        self.assertEqual(decision.last_clip_num, "+8613123123123")
        self.assertEqual(decision.last_clip_time, 10.0)
        self.assertEqual(decision.ring_timeout_target, 22.0)

    def test_handle_clip_line_blocks_blacklisted_call(self):
        decision = handle_clip_line(
            '+CLIP: "+8613123123123",129',
            last_clip_num="",
            last_clip_time=0.0,
            now=10.0,
            filter_mode="Blacklist",
            whitelist=[],
            blacklist=["13123123123"],
        )

        self.assertTrue(decision.blocked)
        self.assertTrue(decision.new_clip)
        self.assertEqual(decision.ring_timeout_target, 0.0)
        self.assertTrue(call_push_message(decision.caller_num, decision.blocked, decision.block_reason))

    def test_refresh_and_expire_ring_timeout(self):
        self.assertEqual(refresh_ring_timeout("[I]-[ril] RING", 0.0, 10.0), 22.0)
        self.assertEqual(refresh_ring_timeout("[I]-[ril] RING", -1.0, 10.0), -1.0)
        self.assertEqual(refresh_ring_timeout("noise", 5.0, 10.0), 5.0)
        self.assertTrue(ring_timeout_expired(5.0, 6.0))
        self.assertFalse(ring_timeout_expired(5.0, 4.0))

    def test_handle_hangup_line_debounces_notifications(self):
        decision = handle_hangup_line(
            "NO CARRIER",
            ring_timeout_target=20.0,
            current_dial_num="",
            popup_active=False,
            last_clip_num="10086",
            last_hangup_time=5.0,
            now=10.0,
        )

        self.assertTrue(decision.matched)
        self.assertTrue(decision.should_notify)
        self.assertEqual(decision.ring_timeout_target, 0.0)
        self.assertEqual(decision.current_dial_num, "")
        self.assertEqual(decision.last_clip_num, "")
        self.assertEqual(decision.last_hangup_time, 10.0)

        debounced = handle_hangup_line(
            "NO CARRIER",
            ring_timeout_target=20.0,
            current_dial_num="",
            popup_active=False,
            last_clip_num="10086",
            last_hangup_time=9.0,
            now=10.0,
        )
        self.assertTrue(debounced.matched)
        self.assertFalse(debounced.should_notify)
        self.assertEqual(debounced.last_clip_num, "10086")

    def test_connected_call_number_and_active_state(self):
        self.assertEqual(connected_call_number(' +CIEV: "CALL",1', "10086"), "10086")
        self.assertEqual(connected_call_number(' +CIEV: "CALL",1', ""), "")
        self.assertTrue(is_call_active(0.0, "10086", False))
        self.assertTrue(is_call_active(0.0, "", True))
        self.assertFalse(is_call_active(0.0, "", False))

    def test_handle_call_line_reports_incoming_call(self):
        decision = handle_call_line(
            '+CLIP: "+8613123123123",129',
            CallState(),
            now=10.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=False,
        )

        self.assertEqual(decision.incoming_number, "+8613123123123")
        self.assertEqual(decision.show_popup_number, "+8613123123123")
        self.assertEqual(decision.push_message, "收到来电：来自 +8613123123123")
        self.assertEqual(decision.state.last_clip_num, "+8613123123123")
        self.assertEqual(decision.state.ring_timeout_target, 22.0)
        self.assertFalse(decision.stop_processing)

    def test_handle_call_line_ignores_websocket_json_clip_payload(self):
        decision = handle_call_line(
            '[I]-[websocket] json: {"message":"+CLIP: \\"+8613123123123\\",129"}',
            CallState(),
            now=10.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=False,
        )

        self.assertEqual(decision.incoming_number, "")
        self.assertEqual(decision.push_message, "")
        self.assertEqual(decision.state.ring_timeout_target, 0.0)

    def test_handle_call_line_blocks_blacklisted_call(self):
        decision = handle_call_line(
            '+CLIP: "+8613123123123",129',
            CallState(),
            now=10.0,
            filter_mode="Blacklist",
            whitelist=[],
            blacklist=["13123123123"],
            popup_active=False,
        )

        self.assertEqual(decision.blocked_number, "+8613123123123")
        self.assertTrue(decision.block_reason)
        self.assertTrue(decision.stop_processing)
        self.assertEqual(decision.state.ring_timeout_target, 0.0)

    def test_handle_call_line_reports_hangup_and_connected(self):
        hangup = handle_call_line(
            "NO CARRIER",
            CallState(ring_timeout_target=20.0, last_clip_num="10086", last_hangup_time=5.0),
            now=10.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=False,
        )
        self.assertTrue(hangup.hangup_notify)
        self.assertEqual(hangup.state.ring_timeout_target, 0.0)
        self.assertEqual(hangup.state.last_clip_num, "")

        connected = handle_call_line(
            ' +CIEV: "CALL",1',
            CallState(ring_timeout_target=-1.0, current_dial_num="10086"),
            now=20.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=True,
        )
        self.assertEqual(connected.connected_number, "10086")
        self.assertEqual(connected.state.ring_timeout_target, 0.0)


if __name__ == "__main__":
    unittest.main()
