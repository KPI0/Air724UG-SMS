import unittest

from sms_core.call_events import (
    CallState,
    call_push_message,
    call_end_message,
    connected_call_number,
    handle_call_line,
    handle_clip_line,
    handle_hangup_line,
    is_call_active,
    line_confirms_call_presence,
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

    def test_call_presence_lines_take_priority_over_local_ring_timeout(self):
        self.assertTrue(line_confirms_call_presence("RING"))
        self.assertTrue(line_confirms_call_presence('+CLIP: "10086",129'))
        self.assertTrue(line_confirms_call_presence('+CIEV: "CALL",1'))
        self.assertFalse(line_confirms_call_presence("NO CARRIER"))

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
        self.assertEqual(debounced.last_clip_num, "")

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

    def test_handle_call_line_reports_corrupted_clip_as_unknown_call(self):
        decision = handle_call_line(
            '+CLIP: "(invalid)",129,,,,0',
            CallState(),
            now=10.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=False,
        )

        self.assertEqual(decision.incoming_number, "未知号码")
        self.assertEqual(decision.show_popup_number, "未知号码")
        self.assertEqual(decision.state.last_clip_num, "未知号码")
        self.assertEqual(decision.state.ring_timeout_target, 22.0)

    def test_handle_call_line_blocks_unknown_caller_in_whitelist_mode(self):
        decision = handle_call_line(
            '+CLIP: "",129',
            CallState(),
            now=10.0,
            filter_mode="Whitelist",
            whitelist=["10086", "未知号码"],
            blacklist=[],
            popup_active=False,
        )

        self.assertEqual(decision.blocked_number, "未知号码")
        self.assertEqual(decision.block_reason, "不在白名单")
        self.assertTrue(decision.stop_processing)
        self.assertEqual(decision.state.ring_timeout_target, 0.0)

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
        self.assertTrue(hangup.call_ended)
        self.assertEqual(hangup.end_direction, "incoming")
        self.assertEqual(hangup.end_number, "10086")
        self.assertEqual(hangup.end_phase, "ended")
        self.assertEqual(hangup.end_reason, "NO CARRIER")
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

        incoming_connected = handle_call_line(
            ' +CIEV: "CALL",1',
            CallState(ring_timeout_target=20.0, last_clip_num="10010"),
            now=20.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=True,
        )
        self.assertEqual(incoming_connected.incoming_connected_number, "10010")
        self.assertEqual(incoming_connected.state.ring_timeout_target, -1.0)

    def test_duplicate_clip_does_not_reset_connected_call_timeout(self):
        decision = handle_call_line(
            '+CLIP: "10010",129',
            CallState(
                ring_timeout_target=-1.0,
                last_clip_num="10010",
                last_clip_time=0.0,
                call_session_id="incoming:connected",
            ),
            now=30.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=True,
        )

        self.assertEqual(decision.incoming_number, "")
        self.assertEqual(decision.state.ring_timeout_target, -1.0)
        self.assertEqual(decision.state.call_session_id, "incoming:connected")

    def test_outgoing_hangup_is_marked_separately_with_reason(self):
        decision = handle_call_line(
            '+CIEV: "CALL",0',
            CallState(current_dial_num="10086"),
            now=20.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=False,
        )

        self.assertTrue(decision.call_ended)
        self.assertTrue(decision.outgoing_call_ended)
        self.assertEqual(decision.end_direction, "outgoing")
        self.assertEqual(decision.end_number, "10086")
        self.assertEqual(decision.end_phase, "ended")
        self.assertEqual(decision.end_reason, "CALL=0")
        self.assertEqual(decision.end_message, "📞 对方已挂断")
        self.assertEqual(call_end_message("BUSY"), "📞 对方忙线")
        self.assertEqual(call_end_message("NO ANSWER"), "📞 对方未接听")

    def test_explicit_modem_dial_error_finishes_outgoing_context(self):
        decision = handle_call_line(
            "+CME ERROR: 3",
            CallState(current_dial_num="15923240141"),
            now=20.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=False,
        )

        self.assertTrue(decision.call_ended)
        self.assertTrue(decision.outgoing_call_ended)
        self.assertEqual(decision.end_direction, "outgoing")
        self.assertEqual(decision.end_phase, "failed")
        self.assertEqual(decision.end_reason, "+CME ERROR: 3")
        self.assertEqual(decision.state.current_dial_num, "")

    def test_new_call_resets_hangup_debounce_generation(self):
        incoming = handle_call_line(
            '+CLIP: "10010",129',
            CallState(last_hangup_time=10.0),
            now=11.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=False,
        )
        ended = handle_call_line(
            "NO CARRIER",
            incoming.state,
            now=12.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=True,
        )

        self.assertEqual(incoming.state.last_hangup_time, 0.0)
        self.assertTrue(ended.call_ended)
        self.assertTrue(ended.hangup_notify)
        self.assertEqual(ended.state.last_clip_num, "")

    def test_call_session_id_survives_connect_and_is_carried_by_terminal(self):
        incoming = handle_call_line(
            '+CLIP: "10010",129',
            CallState(),
            now=11.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=False,
        )
        session_id = incoming.state.call_session_id
        self.assertTrue(session_id.startswith("incoming:"))
        self.assertNotIn("10010", session_id)

        connected = handle_call_line(
            '+CIEV: "CALL",1',
            incoming.state,
            now=12.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=True,
        )
        self.assertEqual(connected.call_session_id, session_id)
        self.assertEqual(connected.state.call_session_id, session_id)

        ended = handle_call_line(
            "NO CARRIER",
            connected.state,
            now=13.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=True,
        )
        self.assertEqual(ended.call_session_id, session_id)
        self.assertEqual(ended.state.call_session_id, "")

    def test_same_number_redial_gets_a_new_call_session_id(self):
        first = handle_call_line(
            '+CLIP: "10010",129',
            CallState(),
            now=20.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=False,
        )
        ended = handle_call_line(
            "NO CARRIER",
            first.state,
            now=21.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=True,
        )
        second = handle_call_line(
            '+CLIP: "10010",129',
            ended.state,
            now=22.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=False,
        )
        self.assertNotEqual(first.state.call_session_id, second.state.call_session_id)

    def test_debounced_hangup_still_marks_call_ended(self):
        ended = handle_call_line(
            "NO CARRIER",
            CallState(
                ring_timeout_target=20.0,
                last_clip_num="10010",
                last_hangup_time=9.0,
            ),
            now=10.0,
            filter_mode="Disabled",
            whitelist=[],
            blacklist=[],
            popup_active=False,
        )

        self.assertTrue(ended.call_ended)
        self.assertFalse(ended.hangup_notify)
        self.assertEqual(ended.state.ring_timeout_target, 0.0)
        self.assertEqual(ended.state.last_clip_num, "")


if __name__ == "__main__":
    unittest.main()
