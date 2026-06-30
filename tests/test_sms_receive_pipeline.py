import unittest

from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.long_sms_assembler import LongSmsAssembler
from sms_core.serial_sms import CollectedSmsCallback
from sms_core.serial_sms_pdu_cache import SmsPduCorrectionCache
from sms_core.sms_pdu import ConcatSmsInfo
from sms_core.sms_receive_pipeline import SmsReceivePipeline


def parse_head(text):
    parts = str(text or "").split(" ", 1)
    return (parts[0], parts[1] if len(parts) > 1 else "")


def _swap_number_digits(number):
    digits = str(number or "").lstrip("+")
    if len(digits) % 2:
        digits += "F"
    return "".join(digits[i + 1] + digits[i] for i in range(0, len(digits), 2))


def _incoming_ucs2_pdu(
    sender,
    message,
    *,
    reference=0x2A,
    total=1,
    index=1,
    timestamp_hex="62608211510523",
):
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


class NoopCorrectionCache:
    def observe_line(self, line, now):
        pass

    def correct_callback_text(self, text, parse_callback_head, now):
        return text

    def concat_part_for_callback(self, text, parse_callback_head, now):
        return None


class CorrectingCache(NoopCorrectionCache):
    def correct_callback_text(self, text, parse_callback_head, now):
        return "+10086 corrected body"


class CachedPart:
    sender = "+10086"
    body = "part1"
    concat_info = ConcatSmsInfo(0x2A, 2, 1)
    timestamp = "26/06/28,11:15:50+32"


class ConcatCache(NoopCorrectionCache):
    def concat_part_for_callback(self, text, parse_callback_head, now):
        return CachedPart()


class FailingSemanticCache(NoopCorrectionCache):
    def correct_callback_text(self, *_args):
        raise AssertionError("pipeline must not correct SMS content")

    def concat_part_for_callback(self, *_args):
        raise AssertionError("pipeline must not inspect concat metadata")


class CapturingAssembler:
    def __init__(self):
        self.calls = []

    def add_collected(self, collected, correction_cache=None, now=None, log=None):
        self.calls.append((collected, correction_cache, now, log))
        return "forwarded"

    def reset(self):
        pass


class SmsReceivePipelineTests(unittest.TestCase):
    def test_pipeline_add_collected_is_blind_forwarder(self):
        assembler = CapturingAssembler()
        cache = FailingSemanticCache()
        pipeline = SmsReceivePipeline(parse_head, cache, assembler)
        collected = CollectedSmsCallback("+10086 first", ["second"])
        log = object()

        result = pipeline.add_collected(collected, now=1.0, log=log)

        self.assertEqual(result, "forwarded")
        self.assertEqual(assembler.calls, [(collected, cache, 1.0, log)])

    def test_normal_callback_keeps_raw_continuation_lines(self):
        pipeline = SmsReceivePipeline(
            parse_head,
            NoopCorrectionCache(),
            LongSmsAssembler(parse_head),
        )

        pending = pipeline.add_collected(
            CollectedSmsCallback("+10086 first", ["second", "third"]),
            now=1.0,
        )

        self.assertEqual(pending.callback_head, "+10086 first")
        self.assertEqual(pending.full_msg, "first\nsecond\nthird")
        self.assertEqual(pending.display_lines, ["📩 收到短信：", "+10086 first", "second", "third"])

    def test_corrected_callback_ignores_raw_continuation_lines(self):
        pipeline = SmsReceivePipeline(
            parse_head,
            CorrectingCache(),
            LongSmsAssembler(parse_head),
        )

        pending = pipeline.add_collected(
            CollectedSmsCallback("+10086 broken", ["pollution"]),
            now=1.0,
        )

        self.assertEqual(pending.callback_head, "+10086 corrected body")
        self.assertEqual(pending.full_msg, "corrected body")
        self.assertEqual(pending.display_lines, ["📩 收到短信：", "+10086 corrected body"])

    def test_concat_part_uses_pdu_body_and_ignores_raw_continuation_lines(self):
        pipeline = SmsReceivePipeline(
            parse_head,
            ConcatCache(),
            LongSmsAssembler(parse_head),
        )

        result = pipeline.add_collected(
            CollectedSmsCallback("+10086 noisy", ["pollution"]),
            now=1.0,
        )

        self.assertIsNone(result)
        key = ("+10086", 8, 0x2A, 2, 0)
        entry = pipeline.long_sms_assembler._pending[key]
        self.assertEqual(entry["parts"], {1: "part1"})

    def test_corrected_multipart_pdu_keeps_trace_metadata(self):
        cache = SmsPduCorrectionCache()
        body = "PDU trace part one and part two complete body"
        part1 = body[:22]
        part2 = body[22:]
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
        pipeline = SmsReceivePipeline(
            parse_sms_callback_head,
            cache,
            LongSmsAssembler(parse_sms_callback_head),
        )

        result = pipeline.add_collected(
            CollectedSmsCallback(
                "10086 26/06/28,11:15:50+32 " + body[:24] + "\ufffd",
                ["pollution"],
            ),
            now=12.0,
        )

        self.assertEqual(result.full_msg, body)
        self.assertEqual(result.concat_reference, 0x2A)
        self.assertEqual(result.concat_reference_bits, 8)
        self.assertEqual(result.concat_total, 2)
        self.assertRegex(result.message_trace_id, r"^[0-9a-f]{12}$")
        self.assertEqual(pipeline.long_sms_assembler._pending, {})

    def test_reference_reuse_after_session_window_does_not_mix_full_pipeline_parts(self):
        cache = SmsPduCorrectionCache()
        ref = 0x77
        timestamp_a = "26/06/28,11:15:50+32"
        timestamp_b = "26/06/28,11:15:51+32"
        timestamp_b_hex = "62608211511523"
        pipeline = SmsReceivePipeline(
            parse_sms_callback_head,
            cache,
            LongSmsAssembler(parse_sms_callback_head),
        )

        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,80", 10.0)
        cache.observe_line(_incoming_ucs2_pdu(
            "10086",
            "A1-",
            reference=ref,
            total=2,
            index=1,
            timestamp_hex="62608211510523",
        ), 10.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2)
        self.assertIsNone(pipeline.add_collected(
            CollectedSmsCallback("10086 " + timestamp_a + " A1-", []),
            now=20.0,
        ))

        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=2 true OK +CMGR: 0,,80", 80.0)
        cache.observe_line(_incoming_ucs2_pdu(
            "10086",
            "B2",
            reference=ref,
            total=2,
            index=2,
            timestamp_hex=timestamp_b_hex,
        ), 80.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 80.2)
        self.assertIsNone(pipeline.add_collected(
            CollectedSmsCallback("10086 " + timestamp_b + " B2", []),
            now=81.0,
        ))
        self.assertEqual(len(pipeline.long_sms_assembler._pending), 2)

        cache.observe_line("[I]-[lib_sms rsp] +CMGR AT+CMGR=3 true OK +CMGR: 0,,80", 83.0)
        cache.observe_line(_incoming_ucs2_pdu(
            "10086",
            "B1-",
            reference=ref,
            total=2,
            index=1,
            timestamp_hex=timestamp_b_hex,
        ), 83.1)
        cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 83.2)
        complete_b = pipeline.add_collected(
            CollectedSmsCallback("10086 " + timestamp_b + " B1-", []),
            now=86.0,
        )

        self.assertEqual(complete_b.full_msg, "B1-B2")
        self.assertEqual(len(pipeline.long_sms_assembler._pending), 1)

    def test_merged_lua_callback_wins_over_matching_first_pdu_part(self):
        cache = SmsPduCorrectionCache()
        body = (
            "【充值成功提醒】尊敬的客户：您已充值0.01元，充值后可用通用余额为352.62元。"
            "短信回复数字1025可查看余额及费用详情。更多查询及缴费方式："
            "1、拨打查费专线1000111或微信关注湖南电信服务号。"
            "2、登录中国电信APP。"
            "3、点击查询链接，如对费用有疑问，可进入在线人工客服咨询。"
            "温馨提醒：缴费复机后如仍不能上网可重启终端。【中国电信】"
        )
        parts = [body[:70], body[70:140], body[140:210], body[210:]]
        timestamp_hexes = [
            "62608212033323",
            "62608212037323",
            "62608212033423",
            "62608212030423",
        ]
        for index, (part, timestamp_hex) in enumerate(zip(parts, timestamp_hexes), start=1):
            cache.observe_line(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,156", 10.0 + index)
            cache.observe_line(_incoming_ucs2_pdu(
                "10001",
                part,
                reference=0xB1,
                total=4,
                index=index,
                timestamp_hex=timestamp_hex,
            ), 10.1 + index)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + index)
        pipeline = SmsReceivePipeline(
            parse_sms_callback_head,
            cache,
            LongSmsAssembler(parse_sms_callback_head),
        )

        result = pipeline.add_collected(
            CollectedSmsCallback(
                "10001 26/06/28,21:30:33+32 " + parts[0],
                parts[1:],
            ),
            now=20.0,
        )

        self.assertEqual(result.full_msg, body)
        self.assertEqual(result.concat_reference, 0xB1)
        self.assertEqual(result.concat_total, 4)
        self.assertRegex(result.message_trace_id, r"^[0-9a-f]{12}$")
        self.assertEqual(pipeline.long_sms_assembler._pending, {})

    def test_merged_lua_callback_with_newlines_wins_over_matching_first_pdu_part(self):
        cache = SmsPduCorrectionCache()
        body = (
            "截止到2026年06月29日 12时33分，您的费用情况如下：\n"
            "1、本月已产生话费5.20元，通用话费余额：352.66元，本月已充值8.16元\n"
            "2、套餐使用情况：\n"
            "通用流量：本月总量400MB（其中上月结转200MB），已使用5.38MB，剩余394.62MB\n"
            "3、更多查询方式：\n"
            "1）关注“湖南电信”微信公众号\n"
            "2）登录官方链接 t.hn.189.cn/RVVRFzuu\n"
            "3）拨打查费专线1000111\n"
            "4）更多流量回复9"
        )
        parts = [
            body[:70],
            body[70:140],
            body[140:210],
            body[210:],
        ]
        for index, part in enumerate(parts, start=1):
            cache.observe_line(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,156", 10.0 + index)
            cache.observe_line(_incoming_ucs2_pdu(
                "10001",
                part,
                reference=0x50,
                total=4,
                index=index,
                timestamp_hex="62609221334523",
            ), 10.1 + index)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + index)
        pipeline = SmsReceivePipeline(
            parse_sms_callback_head,
            cache,
            LongSmsAssembler(parse_sms_callback_head),
        )

        result = pipeline.add_collected(
            CollectedSmsCallback(
                "10001 26/06/29,12:33:52+32 " + body.split("\n", 1)[0],
                body.split("\n")[1:],
            ),
            now=20.0,
        )

        self.assertIsNotNone(result)
        self.assertEqual(result.full_msg, body)
        self.assertIn("截止到2026年06月29日", result.full_msg)
        self.assertIn("情况如下：\n1、本月", result.full_msg)
        self.assertIn("4）更多流量回复9", result.full_msg)
        self.assertEqual(result.concat_reference, 0x50)
        self.assertEqual(result.concat_total, 4)
        self.assertEqual(pipeline.long_sms_assembler._pending, {})

    def test_merged_lua_callback_uses_pdu_segments_when_concat_metadata_matches(self):
        cache = SmsPduCorrectionCache()
        body = "PART1-body-PART2-body-PART3-body-PART4-body"
        parts = [body[:10], body[10:20], body[20:30], body[30:]]
        for index, part in enumerate(parts, start=1):
            cache.observe_line(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,156", 10.0 + index)
            cache.observe_line(_incoming_ucs2_pdu(
                "10001",
                part,
                reference=0xB1,
                total=4,
                index=index,
                timestamp_hex="62608212033323",
            ), 10.1 + index)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + index)
        pipeline = SmsReceivePipeline(
            parse_sms_callback_head,
            cache,
            LongSmsAssembler(parse_sms_callback_head),
        )

        result = pipeline.add_collected(
            CollectedSmsCallback(
                "10001 26/06/28,21:30:33+32 " + parts[0],
                ["unrelated continuation"],
            ),
            now=20.0,
        )

        self.assertIsNotNone(result)
        self.assertEqual(result.full_msg, body)
        self.assertEqual(result.concat_reference, 0xB1)
        self.assertEqual(pipeline.long_sms_assembler._pending, {})

    def test_merged_lua_callback_from_10001_is_shown_when_parts_are_out_of_order(self):
        cache = SmsPduCorrectionCache()
        body = (
            "截止到2026年06月29日 16时16分，您的费用情况如下：\n"
            "1、本月已产生话费5.20元，通用话费余额：352.70元，本月已充值8.20元\n"
            "2、套餐使用情况：\n"
            "通用流量：本月总量400MB（其中上月结转200MB），已使用5.42MB，剩余394.58MB\n"
            "3、更多查询方式：\n"
            "1）关注“湖南电信”微信公众号\n"
            "2）登录官方链接 t.hn.189.cn/RVVRFzuu\n"
            "3）拨打查费专线1000111\n"
            "4）更多流量回复9"
        )
        parts = [
            "截止到2026年06月29日 16时16分，您的费用情况如下：\n"
            "1、本月已产生话费5.20元，通用话费余额：352.70元，本月已充值",
            "8.20元\n"
            "2、套餐使用情况：\n"
            "通用流量：本月总量400MB（其中上月结转200MB），已使用5.42MB，剩余394.58MB\n"
            "3",
            "、更多查询方式：\n"
            "1）关注“湖南电信”微信公众号\n"
            "2）登录官方链接 t.hn.189.cn/RVVRFzuu \n"
            "3）拨打查费专线10000",
            "111 \n"
            "4）更多流量回复9",
        ]
        timestamps = {
            4: "62609261617223",
            1: "62609261611223",
            3: "62609261614223",
            2: "62609261617223",
        }
        for index in (4, 1, 3, 2):
            cache.observe_line(f"[I]-[lib_sms rsp] +CMGR AT+CMGR=1 true OK +CMGR: 0,,156", 10.0 + index)
            cache.observe_line(_incoming_ucs2_pdu(
                "10001",
                parts[index - 1],
                reference=0x1B,
                total=4,
                index=index,
                timestamp_hex=timestamps[index],
            ), 10.1 + index)
            cache.observe_line("[I]-[TP-PID : ] 0 dcs:  8", 10.2 + index)
        pipeline = SmsReceivePipeline(
            parse_sms_callback_head,
            cache,
            LongSmsAssembler(parse_sms_callback_head),
        )

        callback_lines = body.split("\n")
        result = pipeline.add_collected(
            CollectedSmsCallback(
                "10001 26/06/29,16:16:17+32 " + callback_lines[0],
                callback_lines[1:],
            ),
            now=30.0,
        )

        self.assertIsNotNone(result)
        self.assertEqual(result.full_msg, "".join(parts))
        self.assertIn("4）更多流量回复9", result.full_msg)
        self.assertEqual(pipeline.long_sms_assembler._pending, {})


if __name__ == "__main__":
    unittest.main()
