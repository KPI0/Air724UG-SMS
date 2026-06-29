import unittest

from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.serial_sms_pdu_cache import SmsPduCorrectionCache


def _swap_number_digits(number):
    digits = str(number or "").lstrip("+")
    if len(digits) % 2:
        digits += "F"
    return "".join(digits[i + 1] + digits[i] for i in range(0, len(digits), 2))


def _incoming_ucs2_pdu(sender, message, *, reference=0x2A, total=1, index=1, timestamp_hex="62608211510523"):
    sender_text = str(sender or "")
    number_type = "91" if sender_text.startswith("+") else "81"
    sender_digits = sender_text.lstrip("+")
    first_octet = "40" if total > 1 else "00"
    payload = str(message or "").encode("utf-16-be")
    user_data = payload
    if total > 1:
        user_data = bytes((0x05, 0x00, 0x03, reference & 0xFF, total & 0xFF, index & 0xFF)) + payload
    return (
        "00"
        + first_octet
        + f"{len(sender_digits):02X}"
        + number_type
        + _swap_number_digits(sender_digits)
        + "00"
        + "08"
        + timestamp_hex
        + f"{len(user_data):02X}"
        + user_data.hex().upper()
    )


def parse_callback_head(text):
    sender, body = parse_sms_callback_head(text)
    if sender:
        return sender, body
    parts = str(text or "").split(" ", 1)
    return (parts[0], parts[1] if len(parts) > 1 else "")


class SmsPduCorrectionCacheTests(unittest.TestCase):
    def test_correct_callback_text_consumes_matching_cache_entry(self):
        cache = SmsPduCorrectionCache()
        cache._complete_by_key[("10086", "")] = [("中国电信温馨提醒:尊享来电识别", 10.0)]

        corrected = cache.correct_callback_text(
            "10086 中国电信温馨提醒:尊享来电�",
            parse_callback_head,
            11.0,
        )

        self.assertEqual(corrected, "10086 中国电信温馨提醒:尊享来电识别")
        self.assertEqual(cache.correct_callback_text(
            "10086 中国电信温馨提醒:尊享来电�",
            parse_callback_head,
            12.0,
        ), "10086 中国电信温馨提醒:尊享来电�")

    def test_correct_callback_text_rejects_stale_same_sender_mismatch(self):
        cache = SmsPduCorrectionCache()
        cache._complete_by_key[("10086", "")] = [("上一条短信完整正文", 10.0)]

        corrected = cache.correct_callback_text(
            "10086 新短信内容�",
            parse_callback_head,
            11.0,
        )

        self.assertEqual(corrected, "10086 新短信内容�")
        self.assertIn(("10086", ""), cache._complete_by_key)

    def test_correct_callback_text_requires_matching_timestamp_when_present(self):
        cache = SmsPduCorrectionCache()
        cache._complete_by_key[("10086", "26/06/28,11:15:50+32")] = [
            ("【中国电信】验证码342089，3分钟内有效。", 10.0),
        ]

        corrected = cache.correct_callback_text(
            "10086 26/06/28,11:16:10+32 【中国电信】验证码�",
            parse_callback_head,
            11.0,
        )

        self.assertEqual(corrected, "10086 26/06/28,11:16:10+32 【中国电信】验证码�")

    def test_correct_callback_text_normalizes_china_country_code(self):
        cache = SmsPduCorrectionCache()
        cache._complete_by_key[("13812345678", "26/06/28,11:15:50+32")] = [
            ("中国电信温馨提醒:尊享来电识别", 10.0),
        ]

        corrected = cache.correct_callback_text(
            "+8613812345678 26/06/28,11:15:50+32 中国电信温馨提醒:尊享来电�",
            parse_callback_head,
            11.0,
        )

        self.assertEqual(
            corrected,
            "+8613812345678 26/06/28,11:15:50+32 中国电信温馨提醒:尊享来电识别",
        )

    def test_correct_callback_text_selects_matching_candidate_for_same_timestamp(self):
        cache = SmsPduCorrectionCache()
        cache._complete_by_key[("10086", "26/06/28,11:15:50+32")] = [
            ("【中国电信】验证码342089，3分钟内有效。", 10.0),
            ("【中国电信】账单提醒，您本月消费20元。", 10.0),
        ]

        corrected = cache.correct_callback_text(
            "10086 26/06/28,11:15:50+32 【中国电信】账单提醒，您本月�",
            parse_callback_head,
            11.0,
        )

        self.assertEqual(
            corrected,
            "10086 26/06/28,11:15:50+32 【中国电信】账单提醒，您本月消费20元。",
        )
        self.assertEqual(
            cache._complete_by_key[("10086", "26/06/28,11:15:50+32")],
            [("【中国电信】验证码342089，3分钟内有效。", 10.0)],
        )

    def test_correct_callback_text_assembles_multipart_with_single_timestamp_anchor(self):
        cache = SmsPduCorrectionCache()
        body = "中国电信温馨提醒:尊享来电识别【号码百事通】"
        part1 = body[:12]
        part2 = body[12:]
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0)
        cache.observe_line(_incoming_ucs2_pdu("10086", part1, total=2, index=1), 10.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=2 true OK +CMGR: 0,,80", 11.0)
        cache.observe_line(_incoming_ucs2_pdu(
            "10086",
            part2,
            total=2,
            index=2,
        ), 11.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 11.2)

        corrected = cache.correct_callback_text(
            "10086 26/06/28,11:15:50+32 " + body[:18] + "\ufffd",
            parse_callback_head,
            12.0,
        )

        self.assertEqual(corrected, "10086 26/06/28,11:15:50+32 " + body)

    def test_correct_callback_text_rejects_multipart_with_near_timestamp_anchor_collision(self):
        cache = SmsPduCorrectionCache()
        body = "中国电信温馨提醒:尊享来电识别【号码百事通】"
        part1 = body[:12]
        part2 = body[12:]
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0)
        cache.observe_line(_incoming_ucs2_pdu("10086", part1, total=2, index=1), 10.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=2 true OK +CMGR: 0,,80", 11.0)
        cache.observe_line(_incoming_ucs2_pdu(
            "10086",
            part2,
            total=2,
            index=2,
            timestamp_hex="62608211511523",
        ), 11.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 11.2)
        corrupted = "10086 26/06/28,11:15:50+32 " + body[:18] + "\ufffd"

        corrected = cache.correct_callback_text(corrupted, parse_callback_head, 12.0)

        self.assertEqual(corrected, corrupted)

    def test_correct_callback_text_rejects_multipart_with_far_timestamps(self):
        cache = SmsPduCorrectionCache(multipart_timestamp_tolerance=5.0)
        body = "中国电信温馨提醒:尊享来电识别【号码百事通】"
        part1 = body[:12]
        part2 = body[12:]
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0)
        cache.observe_line(_incoming_ucs2_pdu("10086", part1, total=2, index=1), 10.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=2 true OK +CMGR: 0,,80", 11.0)
        cache.observe_line(_incoming_ucs2_pdu(
            "10086",
            part2,
            total=2,
            index=2,
            timestamp_hex="62608211610223",
        ), 11.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 11.2)
        corrupted = "10086 26/06/28,11:15:50+32 " + body[:18] + "\ufffd"

        corrected = cache.correct_callback_text(corrupted, parse_callback_head, 12.0)

        self.assertEqual(corrected, corrupted)

    def test_correct_callback_text_rejects_multipart_ref_collision_with_timestamp_gap(self):
        cache = SmsPduCorrectionCache(multipart_timestamp_tolerance=5.0)
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0)
        cache.observe_line(_incoming_ucs2_pdu(
            "10086",
            "A_PREFIX_",
            reference=0x2A,
            total=2,
            index=1,
        ), 10.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=2 true OK +CMGR: 0,,80", 11.0)
        cache.observe_line(_incoming_ucs2_pdu(
            "10086",
            "B_SUFFIX",
            reference=0x2A,
            total=2,
            index=2,
            timestamp_hex="62608211512523",
        ), 11.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 11.2)
        corrupted = "10086 26/06/28,11:15:50+32 A_PREFIX_\ufffd"

        corrected = cache.correct_callback_text(corrupted, parse_callback_head, 12.0)

        self.assertEqual(corrected, corrupted)

    def test_concat_info_for_callback_matches_exact_part(self):
        cache = SmsPduCorrectionCache()
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0)
        cache.observe_line(_incoming_ucs2_pdu("10086", "AA", total=2, index=1), 10.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)

        part = cache.concat_part_for_callback(
            "10086 26/06/28,11:15:50+32 AA",
            parse_callback_head,
            10.3,
        )

        info = part.concat_info
        self.assertEqual(part.body, "AA")
        self.assertEqual(part.timestamp, "26/06/28,11:15:50+32")
        self.assertEqual((info.reference, info.total, info.index), (0x2A, 2, 1))

    def test_concat_info_does_not_match_already_merged_body(self):
        cache = SmsPduCorrectionCache()
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0)
        cache.observe_line(_incoming_ucs2_pdu("10086", "AA", total=2, index=1), 10.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)

        info = cache.concat_info_for_callback(
            "10086 26/06/28,11:15:50+32 AABB",
            parse_callback_head,
            10.3,
        )

        self.assertIsNone(info)

    def test_concat_part_falls_back_to_unique_pdu_sender(self):
        cache = SmsPduCorrectionCache()
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0)
        cache.observe_line(_incoming_ucs2_pdu("10086", "AA", total=2, index=1), 10.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)

        part = cache.concat_part_for_callback(
            "AA",
            parse_sms_callback_head,
            10.3,
        )

        self.assertEqual(part.sender, "10086")
        self.assertEqual(part.body, "AA")
        self.assertEqual(part.timestamp, "26/06/28,11:15:50+32")
        self.assertEqual((part.concat_info.reference, part.concat_info.total, part.concat_info.index), (0x2A, 2, 1))

    def test_concat_part_matches_unique_near_timestamp(self):
        cache = SmsPduCorrectionCache(multipart_timestamp_tolerance=5.0)
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0)
        cache.observe_line(_incoming_ucs2_pdu(
            "10086",
            "AA",
            total=2,
            index=2,
            timestamp_hex="62608211511523",
        ), 10.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)

        part = cache.concat_part_for_callback(
            "10086 26/06/28,11:15:50+32 AA",
            parse_callback_head,
            10.3,
        )

        self.assertEqual(part.sender, "10086")
        self.assertEqual(part.body, "AA")
        self.assertEqual(part.timestamp, "26/06/28,11:15:51+32")
        self.assertEqual((part.concat_info.reference, part.concat_info.total, part.concat_info.index), (0x2A, 2, 2))

    def test_concat_part_dedupes_identical_near_timestamp_matches(self):
        cache = SmsPduCorrectionCache(multipart_timestamp_tolerance=5.0)
        for offset in (0.0, 1.0):
            cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0 + offset)
            cache.observe_line(_incoming_ucs2_pdu(
                "10086",
                "AA",
                total=2,
                index=2,
                timestamp_hex="62608211511523",
            ), 10.1 + offset)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + offset)

        part = cache.concat_part_for_callback(
            "10086 26/06/28,11:15:50+32 AA",
            parse_callback_head,
            12.0,
        )

        self.assertIsNotNone(part)
        self.assertEqual(part.sender, "10086")
        self.assertEqual(part.body, "AA")
        self.assertEqual(part.timestamp, "26/06/28,11:15:51+32")
        self.assertEqual((part.concat_info.reference, part.concat_info.total, part.concat_info.index), (0x2A, 2, 2))

    def test_concat_part_fallback_requires_unique_match(self):
        cache = SmsPduCorrectionCache()
        for index, sender in enumerate(("10086", "10010"), start=1):
            cache.observe_line(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,80", 10.0 + index)
            cache.observe_line(_incoming_ucs2_pdu(sender, "AA", total=2, index=1), 10.1 + index)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + index)

        self.assertIsNone(cache.concat_part_for_callback("AA", parse_sms_callback_head, 13.0))

    def test_multipart_cache_limit_evicts_oldest_entry(self):
        logs = []
        cache = SmsPduCorrectionCache(max_multipart_entries=1, max_segment_entries=10)
        for index, reference in enumerate((0x2A, 0x2B), start=1):
            cache.observe_line(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,80", 10.0 + index, log=logs.append)
            cache.observe_line(_incoming_ucs2_pdu("10086", f"P{index}", reference=reference, total=2, index=1), 10.1 + index, log=logs.append)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + index, log=logs.append)

        self.assertEqual(len(cache._multipart), 1)
        self.assertFalse(any(key[3] == 0x2A for key in cache._multipart))
        self.assertTrue(any("PDU CACHE EVICT" in item and "Cache=MULTIPART" in item for item in logs))

    def test_segment_cache_limit_evicts_oldest_entry(self):
        logs = []
        cache = SmsPduCorrectionCache(max_segment_entries=2, max_multipart_entries=10)
        for index, reference in enumerate((0x2A, 0x2B, 0x2C), start=1):
            cache.observe_line(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,80", 10.0 + index, log=logs.append)
            cache.observe_line(_incoming_ucs2_pdu("10086", f"P{index}", reference=reference, total=2, index=1), 10.1 + index, log=logs.append)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + index, log=logs.append)

        segment_count = sum(len(items) for items in cache._segments_by_key.values())
        self.assertEqual(segment_count, 2)
        self.assertTrue(any("PDU CACHE EVICT" in item and "Cache=SEGMENT" in item for item in logs))

    def test_complete_cache_limit_evicts_oldest_entry(self):
        logs = []
        cache = SmsPduCorrectionCache(max_complete_entries=1)
        for index, sender in enumerate(("10086", "10010"), start=1):
            cache.observe_line(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,80", 10.0 + index, log=logs.append)
            cache.observe_line(_incoming_ucs2_pdu(sender, f"BODY{index}"), 10.1 + index, log=logs.append)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + index, log=logs.append)

        complete_count = sum(len(items) for items in cache._complete_by_key.values())
        self.assertEqual(complete_count, 1)
        self.assertNotIn(("10086", "26/06/28,11:15:50+32"), cache._complete_by_key)
        self.assertTrue(any("PDU CACHE EVICT" in item and "Cache=COMPLETE" in item for item in logs))


if __name__ == "__main__":
    unittest.main()
