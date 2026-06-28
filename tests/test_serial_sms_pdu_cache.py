import unittest

from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.serial_sms_pdu_cache import SmsPduCorrectionCache


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


if __name__ == "__main__":
    unittest.main()
