import unittest

from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.serial_sms_pdu_cache import SmsPduCorrectionCache


def _swap_number_digits(number):
    digits = str(number or "").lstrip("+")
    if len(digits) % 2:
        digits += "F"
    return "".join(digits[i + 1] + digits[i] for i in range(0, len(digits), 2))


def _pack_gsm7_ascii_sender(sender):
    septets = [ord(char) for char in str(sender or "")]
    if any(value < 0 or value > 0x7F for value in septets):
        raise ValueError("test sender must use GSM7-compatible ASCII")
    packed = sum(value << (7 * index) for index, value in enumerate(septets))
    octet_count = (len(septets) * 7 + 7) // 8
    return packed.to_bytes(octet_count, "little")


def _incoming_alphanumeric_ucs2_pdu(
    sender,
    message,
    *,
    timestamp_hex="62608211510523",
    reference=0x2A,
    total=1,
    index=1,
):
    sender_text = str(sender or "")
    sender_data = _pack_gsm7_ascii_sender(sender_text)
    sender_length = (len(sender_text) * 7 + 3) // 4
    payload = str(message or "").encode("utf-16-be")
    first_octet = "40" if total > 1 else "00"
    user_data = payload
    if total > 1:
        user_data = bytes((
            0x05,
            0x00,
            0x03,
            reference & 0xFF,
            total & 0xFF,
            index & 0xFF,
        )) + payload
    return (
        "00"
        + first_octet
        + f"{sender_length:02X}"
        + "D0"
        + sender_data.hex().upper()
        + "00"
        + "08"
        + timestamp_hex
        + f"{len(user_data):02X}"
        + user_data.hex().upper()
    )


def _incoming_ucs2_pdu(
    sender,
    message,
    *,
    reference=0x2A,
    total=1,
    index=1,
    timestamp_hex="62608211510523",
    number_type=None,
):
    sender_text = str(sender or "")
    number_type = number_type or ("91" if sender_text.startswith("+") else "81")
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
    def test_correct_callback_text_is_compatibility_alias_for_single_pdu_correction(self):
        cache = SmsPduCorrectionCache()
        cache._complete_by_key[("10086", "")] = [("中国电信温馨提醒:尊享来电识别", 10.0)]

        corrected = cache.correct_callback_text(
            "10086 中国电信温馨提醒:尊享来电�",
            parse_callback_head,
            11.0,
        )

        self.assertEqual(corrected, "10086 中国电信温馨提醒:尊享来电识别")
        self.assertEqual(cache.last_corrected_message(), None)

    def test_correct_callback_text_consumes_matching_cache_entry(self):
        cache = SmsPduCorrectionCache()
        cache._complete_by_key[("10086", "")] = [("中国电信温馨提醒:尊享来电识别", 10.0)]

        corrected = cache.correct_single_pdu_callback_text(
            "10086 中国电信温馨提醒:尊享来电�",
            parse_callback_head,
            11.0,
        )

        self.assertEqual(corrected, "10086 中国电信温馨提醒:尊享来电识别")
        self.assertEqual(cache.correct_single_pdu_callback_text(
            "10086 中国电信温馨提醒:尊享来电�",
            parse_callback_head,
            12.0,
        ), "10086 中国电信温馨提醒:尊享来电�")

    def test_exact_callback_consumes_matching_cache_entry_without_rewriting_text(self):
        cache = SmsPduCorrectionCache()
        key = ("10086", "26/06/28,11:15:50+32")
        body = "ACCOUNT-NOTICE-LONG-COMMON-PREFIX-OLD-AMOUNT-100-END"
        callback = f"10086 {key[1]} {body}"
        cache._complete_by_key[key] = [(body, 10.0)]

        self.assertEqual(
            cache.correct_single_pdu_callback_text(callback, parse_callback_head, 11.0),
            callback,
        )
        self.assertNotIn(key, cache._complete_by_key)

    def test_ambiguous_prefix_candidates_preserve_callback_text(self):
        cache = SmsPduCorrectionCache()
        key = ("10086", "26/06/28,11:15:50+32")
        prefix = "ACCOUNT-NOTICE-LONG-COMMON-PREFIX-"
        callback = f"10086 {key[1]} {prefix}"
        cache._complete_by_key[key] = [
            (prefix + "OLD-AMOUNT-100-END", 10.0),
            (prefix + "NEW-AMOUNT-200-END", 10.1),
        ]

        self.assertEqual(
            cache.correct_single_pdu_callback_text(callback, parse_callback_head, 11.0),
            callback,
        )
        self.assertEqual(len(cache._complete_by_key[key]), 2)

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
        cache._complete_by_key[("13123123123", "26/06/28,11:15:50+32")] = [
            ("中国电信温馨提醒:尊享来电识别", 10.0),
        ]

        corrected = cache.correct_callback_text(
            "+8613123123123 26/06/28,11:15:50+32 中国电信温馨提醒:尊享来电�",
            parse_callback_head,
            11.0,
        )

        self.assertEqual(
            corrected,
            "+8613123123123 26/06/28,11:15:50+32 中国电信温馨提醒:尊享来电识别",
        )

    def test_alphanumeric_sender_starting_with_86_is_not_country_code_normalized(self):
        cache = SmsPduCorrectionCache()
        timestamp = "26/06/28,11:15:50+32"
        body = "品牌短信正文"
        cache._complete_by_key[("86Brand", timestamp)] = [(body, 10.0)]

        corrected = cache.correct_single_pdu_callback_text(
            "86Brand " + timestamp + " " + body,
            parse_callback_head,
            10.3,
        )

        self.assertEqual(corrected, "86Brand " + timestamp + " " + body)
        self.assertNotIn(("86Brand", timestamp), cache._complete_by_key)

    def test_spaced_alphanumeric_sender_uses_timestamp_boundary(self):
        cache = SmsPduCorrectionCache()
        timestamp = "26/06/28,11:15:50+32"
        body = "品牌活动温馨提醒内容完整"
        pdu = _incoming_alphanumeric_ucs2_pdu("Bank Alert", body)
        cache.observe_line(
            f"[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,{len(pdu) // 2 - 1}",
            10.0,
        )
        cache.observe_line(pdu, 10.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)

        corrected = cache.correct_single_pdu_callback_text(
            "Bank Alert " + timestamp + " 品牌活动温馨提醒�",
            parse_sms_callback_head,
            10.3,
        )

        self.assertEqual(corrected, "Bank Alert " + timestamp + " " + body)

    def test_alphanumeric_numeric_text_keeps_exact_cache_identity(self):
        cache = SmsPduCorrectionCache()
        timestamp = "26/06/28,11:15:50+32"
        alpha_body = "ALPHA-MESSAGE-LONG-CONTENT"
        numeric_body = "NUMERIC-MESSAGE-LONG-CONTENT"
        alpha_pdu = _incoming_alphanumeric_ucs2_pdu("86123", alpha_body)
        numeric_pdu = _incoming_ucs2_pdu("123", numeric_body)
        for index, pdu in enumerate((alpha_pdu, numeric_pdu), start=1):
            cache.observe_line(
                f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,{len(pdu) // 2 - 1}",
                10.0 + index,
            )
            cache.observe_line(pdu, 10.1 + index)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + index)

        self.assertIn(("86123", timestamp), cache._complete_by_key)
        self.assertIn(("123", timestamp), cache._complete_by_key)
        corrected = cache.correct_single_pdu_callback_text(
            "86123 " + timestamp + " ALPHA-MESSAGE-LONG-�",
            parse_sms_callback_head,
            13.5,
        )

        self.assertEqual(corrected, "86123 " + timestamp + " " + alpha_body)
        self.assertIn(("123", timestamp), cache._complete_by_key)

    def test_alphanumeric_numeric_text_selects_its_own_concat_segment(self):
        cache = SmsPduCorrectionCache()
        timestamp = "26/06/28,11:15:50+32"
        body = "COMMON-CONCAT-SEGMENT-BODY"
        numeric_pdu = _incoming_ucs2_pdu("123", body, total=2, index=1)
        alpha_pdu = _incoming_alphanumeric_ucs2_pdu("86123", body, total=2, index=1)
        for index, pdu in enumerate((numeric_pdu, alpha_pdu), start=1):
            cache.observe_line(
                f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,{len(pdu) // 2 - 1}",
                20.0 + index,
            )
            cache.observe_line(pdu, 20.1 + index)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 20.2 + index)

        part = cache.concat_part_for_callback(
            "86123 " + timestamp + " " + body,
            parse_sms_callback_head,
            23.5,
        )

        self.assertIsNotNone(part)
        self.assertEqual(part.sender, "86123")
        self.assertTrue(part.sender_is_alphanumeric)

    def test_correct_callback_text_matches_non_numeric_sender_and_surrogate_pair(self):
        cache = SmsPduCorrectionCache()
        sender = "3A968BD3FABBFC8C52"
        body = "品牌活动温馨提醒：请点击👉观看直播"
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0)
        cache.observe_line(
            _incoming_ucs2_pdu(sender, body, number_type="D0"),
            10.1,
        )
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)

        corrected = cache.correct_single_pdu_callback_text(
            sender + " 26/06/28,11:15:50+32 品牌活动温馨提醒：请点击������观看直播",
            parse_sms_callback_head,
            10.3,
        )

        self.assertEqual(
            corrected,
            "#SamsungHK 26/06/28,11:15:50+32 " + body,
        )

    def test_legacy_sender_correction_matches_old_firmware_86_filtered_alias(self):
        cache = SmsPduCorrectionCache()
        body = "品牌短信正文"
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,22", 10.0)
        cache.observe_line(
            "00000DD068B1501E7693010008626082115105230C"
            + body.encode("utf-16-be").hex().upper(),
            10.1,
        )
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)

        corrected = cache.correct_single_pdu_callback_text(
            "1B05E1673910 26/06/28,11:15:50+32 " + body,
            parse_sms_callback_head,
            10.3,
        )

        self.assertEqual(
            corrected,
            "hbBrand 26/06/28,11:15:50+32 " + body,
        )

    def test_legacy_sender_correction_consumes_only_matching_cached_message(self):
        cache = SmsPduCorrectionCache()
        timestamp = "26/06/28,11:15:50+32"
        cache._complete_by_key[("#SamsungHK", timestamp)] = [
            ("第一条品牌短信", 10.0, "#SamsungHK", True, "3A968BD3FABBFC8C52"),
            ("第二条品牌短信", 10.0, "#SamsungHK", True, "3A968BD3FABBFC8C52"),
        ]

        corrected = cache.correct_single_pdu_callback_text(
            "3A968BD3FABBFC8C52 " + timestamp + " 第一条品牌短信",
            parse_callback_head,
            10.3,
        )

        self.assertTrue(corrected.startswith("#SamsungHK " + timestamp))
        self.assertEqual(
            [item[0] for item in cache._complete_by_key[("#SamsungHK", timestamp)]],
            ["第二条品牌短信"],
        )

    def test_legacy_sender_correction_requires_exact_pdu_derived_alias(self):
        cache = SmsPduCorrectionCache()
        timestamp = "26/06/28,11:15:50+32"
        cache._complete_by_key[("#SamsungHK", timestamp)] = [
            ("品牌短信正文", 10.0, "#SamsungHK", True, "3A968BD3FABBFC8C52"),
        ]

        unchanged = cache.correct_single_pdu_callback_text(
            "10086 " + timestamp + " 品牌短信正文",
            parse_callback_head,
            10.3,
        )

        self.assertEqual(unchanged, "10086 " + timestamp + " 品牌短信正文")

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

    def test_multipart_pdu_cache_exposes_segments_without_assembling_body(self):
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

        callback = "10086 26/06/28,11:15:50+32 " + body[:18] + "\ufffd"
        corrected = cache.correct_callback_text(
            callback,
            parse_callback_head,
            12.0,
        )
        part = cache.concat_part_for_callback(callback, parse_callback_head, 12.0)
        segments = cache.segments_for_concat_part(part, 12.0)

        self.assertEqual(corrected, callback)
        self.assertEqual([segment.body for segment in segments], [part1, part2])
        self.assertEqual(cache._complete_by_key, {})

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

    def test_concat_part_matches_merged_callback_beyond_first_segment(self):
        cache = SmsPduCorrectionCache()
        part1 = "FIRST-PDU-SEGMENT-LONG-ANCHOR-"
        part2 = "SECOND-PDU-SEGMENT-CONTENT"
        for index, part in enumerate((part1, part2), start=1):
            cache.observe_line(
                f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,80",
                10.0 + index,
            )
            cache.observe_line(
                _incoming_ucs2_pdu("10086", part, total=2, index=index),
                10.1 + index,
            )
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + index)

        matched = cache.concat_part_for_callback(
            "10086 26/06/28,11:15:50+32 " + part1 + part2,
            parse_sms_callback_head,
            13.0,
        )

        self.assertIsNotNone(matched)
        self.assertEqual(matched.body, part1)
        self.assertEqual(matched.concat_info.index, 1)

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

    def test_concat_part_matches_unique_far_timestamp_for_same_sender(self):
        cache = SmsPduCorrectionCache(multipart_timestamp_tolerance=5.0)
        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0)
        cache.observe_line(_incoming_ucs2_pdu(
            "10001",
            "【充值成功提醒】尊敬的客户：您已充值0.01元，充值后可用余额",
            reference=0xEE,
            total=4,
            index=1,
            timestamp_hex="62702251611323",
        ), 10.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)

        part = cache.concat_part_for_callback(
            "10001 26/07/22,15:16:19+32 【充值成功提醒】尊敬的客户：您已充值0.01元，充值",
            parse_callback_head,
            10.3,
        )

        self.assertIsNotNone(part)
        self.assertEqual(part.sender, "10001")
        self.assertIn("充值后", part.body)
        self.assertEqual(part.match_kind, "sender_fallback")
        self.assertEqual((part.concat_info.reference, part.concat_info.total, part.concat_info.index), (0xEE, 4, 1))

    def test_concat_part_far_timestamp_fallback_requires_unique_sender_match(self):
        cache = SmsPduCorrectionCache(multipart_timestamp_tolerance=5.0)
        prefix = "同一发送方重复正文前缀足够长"
        for offset, reference in enumerate((0x31, 0x32)):
            cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0 + offset)
            cache.observe_line(_incoming_ucs2_pdu(
                "10001",
                prefix + str(reference),
                reference=reference,
                total=2,
                index=1,
                timestamp_hex="62702251611323",
            ), 10.1 + offset)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + offset)

        part = cache.concat_part_for_callback(
            "10001 26/07/22,15:16:19+32 " + prefix,
            parse_callback_head,
            12.1,
        )

        self.assertIsNone(part)

    def test_concat_part_fallback_requires_unique_match(self):
        cache = SmsPduCorrectionCache()
        for index, sender in enumerate(("10086", "10010"), start=1):
            cache.observe_line(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,80", 10.0 + index)
            cache.observe_line(_incoming_ucs2_pdu(sender, "AA", total=2, index=1), 10.1 + index)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + index)

        self.assertIsNone(cache.concat_part_for_callback("AA", parse_sms_callback_head, 13.0))

    def test_segment_cache_limit_evicts_oldest_entry(self):
        logs = []
        cache = SmsPduCorrectionCache(max_segment_entries=2)
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

    def test_reset_clears_collection_and_all_cached_messages(self):
        cache = SmsPduCorrectionCache()
        cache._collecting = True
        cache._pdu_lines.append("0011")
        cache._complete_by_key[("10086", "time")] = [object()]
        cache._segments_by_key[("10086", 8, 42, 2)] = [object()]
        cache._last_corrected_message = object()

        cache.reset()

        self.assertFalse(cache._collecting)
        self.assertEqual(cache._pdu_lines, [])
        self.assertEqual(cache._complete_by_key, {})
        self.assertEqual(cache._segments_by_key, {})
        self.assertIsNone(cache._last_corrected_message)


if __name__ == "__main__":
    unittest.main()
