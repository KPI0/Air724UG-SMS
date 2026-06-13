import unittest

from sms_core.serial_parsers import (
    parse_temperature,
    parse_cesq_rsrp,
    parse_operator_message,
    parse_sim_status_message,
    parse_attach_status_message,
    parse_cfun_status_message,
    parse_lte_cell_messages,
    parse_serial_debug_insights,
    parse_clip_number,
    evaluate_call_filter,
    is_new_clip,
    is_ring_line,
    is_hangup_event,
    is_call_connected_event,
    is_sms_collection_boundary,
)


class SerialParsersTests(unittest.TestCase):
    def test_parse_temperature_extracts_value(self):
        """Should extract temperature from +RFTEMPERATURE response."""
        result = parse_temperature("+RFTEMPERATURE: 35")
        self.assertEqual(result, "35")

    def test_parse_temperature_returns_none_for_non_match(self):
        """Should return None when line doesn't contain temperature."""
        self.assertIsNone(parse_temperature("OK"))
        self.assertIsNone(parse_temperature(""))

    def test_parse_cesq_rsrp_extracts_sixth_field(self):
        """Should extract RSRP (6th field) from +CESQ response."""
        result = parse_cesq_rsrp("+CESQ: 99,99,255,255,20,80")
        self.assertEqual(result, "80")

    def test_parse_cesq_rsrp_returns_none_for_insufficient_fields(self):
        """Should return None when less than 6 fields."""
        self.assertIsNone(parse_cesq_rsrp("+CESQ: 99,99,255"))

    def test_parse_operator_message_recognizes_china_mobile(self):
        """Should recognize China Mobile PLMN codes."""
        result = parse_operator_message('+COPS: 0,0,"46000"')
        self.assertIn("中国移动", result)
        self.assertIn("46000", result)

    def test_parse_operator_message_recognizes_china_unicom(self):
        """Should recognize China Unicom."""
        result = parse_operator_message('+COPS: 0,0,"46001"')
        self.assertIn("中国联通", result)

    def test_parse_operator_message_recognizes_china_telecom(self):
        """Should recognize China Telecom."""
        result = parse_operator_message('+COPS: 0,0,"46003"')
        self.assertIn("中国电信", result)

    def test_parse_operator_message_handles_unknown_plmn(self):
        """Should return unknown carrier for unrecognized PLMN."""
        result = parse_operator_message('+COPS: 0,0,"99999"')
        self.assertIn("未知运营商", result)

    def test_parse_sim_status_message_recognizes_ready(self):
        """Should recognize READY status."""
        result = parse_sim_status_message("+CPIN: READY")
        self.assertIn("SIM卡正常", result)

    def test_parse_sim_status_message_recognizes_pin_required(self):
        """Should recognize PIN required status."""
        result = parse_sim_status_message("+CPIN: SIM PIN")
        self.assertIn("等待输入PIN码", result)

    def test_parse_sim_status_message_recognizes_not_inserted(self):
        """Should recognize SIM not inserted."""
        result = parse_sim_status_message("+CPIN: NOT INSERTED")
        self.assertIn("未检测到SIM卡", result)

    def test_parse_attach_status_message_recognizes_attached(self):
        """Should recognize attached status (1)."""
        result = parse_attach_status_message("+CGATT: 1")
        self.assertIn("已附着", result)

    def test_parse_attach_status_message_recognizes_detached(self):
        """Should recognize detached status (0)."""
        result = parse_attach_status_message("+CGATT: 0")
        self.assertIn("未附着", result)

    def test_parse_cfun_status_message_recognizes_normal_mode(self):
        """Should recognize normal mode (1)."""
        result = parse_cfun_status_message("+CFUN: 1")
        self.assertIn("正常模式", result)

    def test_parse_cfun_status_message_recognizes_airplane_mode(self):
        """Should recognize airplane mode (0)."""
        result = parse_cfun_status_message("+CFUN: 0")
        self.assertIn("飞行模式", result)

    def test_parse_lte_cell_messages_extracts_cell_info(self):
        """Should extract MCC/MNC/TAC/CI from LTE cell info."""
        # Format: +EEMLTESVC: MCC,MNC,?,TAC,?,?,?,?,?,CI,...
        line = "+EEMLTESVC: 1120,1,0,12345,0,0,0,0,0,67890"
        messages = parse_lte_cell_messages(line)

        self.assertGreater(len(messages), 0)
        self.assertTrue(any("基站定位" in msg for msg in messages))
        self.assertTrue(any("460" in msg for msg in messages))  # MCC in hex

    def test_parse_lte_cell_messages_returns_empty_for_insufficient_fields(self):
        """Should return empty list when fields < 10."""
        messages = parse_lte_cell_messages("+EEMLTESVC: 1,2,3")
        self.assertEqual(messages, [])

    def test_parse_serial_debug_insights_aggregates_parsers(self):
        """Should run all parsers and aggregate results."""
        insights = parse_serial_debug_insights("+CPIN: READY")
        self.assertGreater(len(insights), 0)
        self.assertTrue(any("SIM卡" in msg for msg in insights))

    def test_parse_clip_number_extracts_caller_id(self):
        """Should extract caller ID from +CLIP."""
        number = parse_clip_number('+CLIP: "13800138000",129')
        self.assertEqual(number, "13800138000")

    def test_parse_clip_number_handles_international_format(self):
        """Should handle + prefix in number."""
        number = parse_clip_number('+CLIP: "+8613800138000",145')
        self.assertEqual(number, "+8613800138000")

    def test_parse_clip_number_returns_default_on_no_match(self):
        """Should return default when no CLIP found."""
        number = parse_clip_number("RING", default="未知号码")
        self.assertEqual(number, "未知号码")

    def test_evaluate_call_filter_whitelist_mode_allows_listed(self):
        """Whitelist mode: should allow numbers in whitelist."""
        blocked, reason = evaluate_call_filter(
            "13800138000",
            "Whitelist",
            ["13800138000", "13900139000"],
            []
        )
        self.assertFalse(blocked)

    def test_evaluate_call_filter_whitelist_mode_blocks_unlisted(self):
        """Whitelist mode: should block numbers not in whitelist."""
        blocked, reason = evaluate_call_filter(
            "13700137000",
            "Whitelist",
            ["13800138000"],
            []
        )
        self.assertTrue(blocked)
        self.assertIn("白名单", reason)

    def test_evaluate_call_filter_blacklist_mode_blocks_listed(self):
        """Blacklist mode: should block numbers in blacklist."""
        blocked, reason = evaluate_call_filter(
            "13800138000",
            "Blacklist",
            [],
            ["13800138000"]
        )
        self.assertTrue(blocked)
        self.assertIn("黑名单", reason)

    def test_evaluate_call_filter_blacklist_mode_allows_unlisted(self):
        """Blacklist mode: should allow numbers not in blacklist."""
        blocked, reason = evaluate_call_filter(
            "13700137000",
            "Blacklist",
            [],
            ["13800138000"]
        )
        self.assertFalse(blocked)

    def test_evaluate_call_filter_disabled_mode_allows_all(self):
        """Disabled mode: should allow all numbers."""
        blocked, reason = evaluate_call_filter(
            "13800138000",
            "Disabled",
            [],
            ["13800138000"]
        )
        self.assertFalse(blocked)

    def test_evaluate_call_filter_normalizes_numbers(self):
        """Should normalize both caller and list numbers before comparing."""
        # +8613800138000 should match 13800138000 in whitelist
        blocked, _ = evaluate_call_filter(
            "+8613800138000",
            "Whitelist",
            ["13800138000"],
            []
        )
        self.assertFalse(blocked)

    def test_is_new_clip_detects_different_number(self):
        """Should return True for different caller number."""
        self.assertTrue(is_new_clip("13800138000", "13900139000", 100, 90))

    def test_is_new_clip_detects_timeout(self):
        """Should return True when time window exceeded."""
        # Same number but >4 seconds gap
        self.assertTrue(is_new_clip("13800138000", "13800138000", 100, 90, window_sec=4.0))

    def test_is_new_clip_returns_false_within_window(self):
        """Should return False for same number within time window."""
        self.assertFalse(is_new_clip("13800138000", "13800138000", 100, 98, window_sec=4.0))

    def test_is_ring_line_recognizes_ring(self):
        """Should recognize RING lines."""
        self.assertTrue(is_ring_line("RING"))
        self.assertTrue(is_ring_line("something RING"))

    def test_is_ring_line_rejects_non_ring(self):
        """Should return False for non-RING lines."""
        self.assertFalse(is_ring_line("OK"))
        self.assertFalse(is_ring_line("RINGING"))  # Ends with RING, but has suffix

    def test_is_hangup_event_recognizes_no_carrier(self):
        """Should recognize NO CARRIER as hangup."""
        self.assertTrue(is_hangup_event("NO CARRIER"))

    def test_is_hangup_event_recognizes_busy(self):
        """Should recognize BUSY as hangup."""
        self.assertTrue(is_hangup_event("BUSY"))

    def test_is_hangup_event_recognizes_ciev_call_0(self):
        """Should recognize +CIEV CALL,0 as hangup."""
        self.assertTrue(is_hangup_event('+CIEV: "CALL",0'))

    def test_is_call_connected_event_recognizes_connect(self):
        """Should recognize CONNECT."""
        self.assertTrue(is_call_connected_event("CONNECT"))

    def test_is_call_connected_event_recognizes_ciev_call_1(self):
        """Should recognize +CIEV CALL,1 as connected."""
        self.assertTrue(is_call_connected_event('+CIEV: "CALL",1'))

    def test_is_sms_collection_boundary_recognizes_log_prefixes(self):
        """Should recognize log prefixes as boundaries."""
        self.assertTrue(is_sms_collection_boundary("[I]- Info"))
        self.assertTrue(is_sms_collection_boundary("[W]- Warning"))
        self.assertTrue(is_sms_collection_boundary("[E]- Error"))

    def test_is_sms_collection_boundary_recognizes_at_commands(self):
        """Should recognize AT commands as boundaries."""
        self.assertTrue(is_sms_collection_boundary("AT+CMGS="))
        self.assertTrue(is_sms_collection_boundary("+CMTI: "))

    def test_is_sms_collection_boundary_recognizes_debug_marker(self):
        """Should recognize >>> as boundary."""
        self.assertTrue(is_sms_collection_boundary(">>> Debug"))


if __name__ == "__main__":
    unittest.main()
