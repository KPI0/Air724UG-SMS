import json
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tools.sms_replay import _incoming_ucs2_pdu, _wrap_hex, replay_lines


FIXTURE_PATH = ROOT / "tests" / "fixtures" / "sms_replay" / "cases.json"


def _load_fixture():
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def _calls_by_name(calls, name):
    return [args for kind, args, _kwargs in calls if kind == name]


class SmsSerialReplayFixtureTests(unittest.TestCase):
    def test_fixture_schema_is_append_only_friendly(self):
        data = _load_fixture()

        self.assertEqual(data.get("schema"), 1)
        names = [case.get("name") for case in data.get("cases", [])]
        self.assertEqual(len(names), len(set(names)))
        self.assertGreaterEqual(len(names), 1)
        for case in data["cases"]:
            with self.subTest(case=case.get("name")):
                self.assertIsInstance(case.get("lines"), list)
                self.assertGreater(len(case["lines"]), 0)
                self.assertIsInstance(case.get("expected_popups"), list)
                self.assertIsInstance(case.get("expected_cloud_bodies"), list)

    def test_replays_desensitized_air724ug_sms_logs(self):
        for case in _load_fixture()["cases"]:
            with self.subTest(case=case["name"]):
                calls = replay_lines(case["lines"])
                popups = [args[0] for args in _calls_by_name(calls, "sms_popup")]
                cloud_bodies = [args[1] for args in _calls_by_name(calls, "cloud_sms")]

                self.assertEqual(popups, case["expected_popups"])
                self.assertEqual(cloud_bodies, case["expected_cloud_bodies"])

                for forbidden in case.get("forbidden_popup_substrings", []):
                    self.assertFalse(
                        any(forbidden in popup for popup in popups),
                        f"{forbidden!r} leaked into replay popup output",
                    )

    def test_partial_far_pdu_replay_preserves_only_real_callback_newlines(self):
        body = (
            "【充值成功提醒】尊敬的客户：您已充值0.01元，充值后可用余额为347.35元。"
            "短信回复指定数字可查看余额及费用详情，更多查询及缴费方式如下：\n"
            "1、拨打查费专线或关注服务号查询余额。\n"
            "2、登录官方客户端查询账单和套餐余量。\n"
            "3、点击查询链接进入在线人工客服咨询。\n"
            "温馨提醒：缴费复机后如仍不能上网可重启终端。【运营商】"
        )
        parts = [body[index:index + 67] for index in range(0, len(body), 67)]
        self.assertGreaterEqual(len(parts), 3)
        lines = []
        for index in (1, 3):
            pdu = _incoming_ucs2_pdu(
                "10001",
                parts[index - 1],
                "26/07/22,15:16:31+32" if index == 1 else "26/07/22,15:16:28+32",
                reference=0xEE,
                total=len(parts),
                index=index,
            )
            lines.append(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,156")
            lines.extend(_wrap_hex(pdu))
            lines.append("[I]-[TP-PID : ] 0 dcs:  8")
        first_line, *real_continuations = body.split("\n")
        forced_break = first_line.index("充值后") + len("充值")
        lines.append(
            "[I]-[handler_sms.smsCallback] 10001 26/07/22,15:16:19+32 "
            + first_line[:forced_break]
        )
        lines.extend([first_line[forced_break:]] + real_continuations)
        lines.append("[I]-[ril.proatc] OK")

        calls = replay_lines(lines)
        popups = [args[0] for args in _calls_by_name(calls, "sms_popup")]
        cloud_bodies = [args[1] for args in _calls_by_name(calls, "cloud_sms")]

        self.assertEqual(popups, [body])
        self.assertEqual(cloud_bodies, [body])
        self.assertNotIn("充值\n后", popups[0])
        self.assertEqual(popups[0].count("\n"), 4)

    def test_non_numeric_sender_long_sms_recovers_surrogate_pair_from_pdu(self):
        sender = "3A968BD3FABBFC8C52"
        timestamp = "26/07/22,17:33:02+32"
        body = (
            "品牌活动提醒：今晚九点参加直播，了解全新设备和服务说明。\n"
            "第二行继续介绍活动安排、嘉宾信息和参与方式，方便验证长短信分段重排。\n"
            "👉观看直播，并在规定时间前完成登记，即可获得活动礼遇：https://example.invalid/demo"
        )
        parts = [body[:45], body[45:90], body[90:]]
        self.assertTrue(all(parts))

        lines = []
        for index, part in enumerate(parts, start=1):
            pdu = _incoming_ucs2_pdu(
                sender,
                part,
                timestamp,
                reference=0x60,
                total=3,
                index=index,
                number_type="D0",
            )
            lines.append(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,162")
            lines.extend(_wrap_hex(pdu))
            lines.append("[I]-[TP-PID : ] 0 dcs:  8")

        first_line, second_line, third_line = body.split("\n")
        lines.append(f"[I]-[handler_sms.smsCallback] {sender} {timestamp} {first_line}")
        lines.append(second_line)
        lines.append(third_line.replace("👉", "������"))
        lines.append("[I]-[ril.proatc] OK")

        calls = replay_lines(lines)
        popups = [args[0] for args in _calls_by_name(calls, "sms_popup")]
        cloud_bodies = [args[1] for args in _calls_by_name(calls, "cloud_sms")]

        self.assertEqual(popups, [body])
        self.assertEqual(cloud_bodies, [body])
        self.assertNotIn("�", popups[0])

    def test_serial_wrap_after_first_pdu_segment_is_not_a_message_newline(self):
        sender = "10001"
        timestamp = "26/07/22,18:00:00+32"
        body = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789" * 4
        parts = [body[:48], body[48:96], body[96:]]
        lines = []
        for index, part in enumerate(parts, start=1):
            pdu = _incoming_ucs2_pdu(
                sender,
                part,
                timestamp,
                reference=0x71,
                total=3,
                index=index,
            )
            lines.append(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,162")
            lines.extend(_wrap_hex(pdu))
            lines.append("[I]-[TP-PID : ] 0 dcs:  8")

        split_at = 70
        lines.extend([
            f"[I]-[handler_sms.smsCallback] {sender} {timestamp} {body[:split_at]}",
            body[split_at:],
            "[I]-[ril.proatc] OK",
        ])

        calls = replay_lines(lines)
        popups = [args[0] for args in _calls_by_name(calls, "sms_popup")]
        cloud_bodies = [args[1] for args in _calls_by_name(calls, "cloud_sms")]

        self.assertEqual(popups, [body])
        self.assertEqual(cloud_bodies, [body])
        self.assertNotIn("\n", popups[0])

    def test_reference_reuse_does_not_mix_old_segment_into_new_callback(self):
        sender = "10001"
        old_timestamp = "26/07/22,18:01:00+32"
        new_timestamp = "26/07/22,18:01:01+32"
        reference = 0x72
        new_part1 = "CURRENT-MESSAGE-COMMON-PREFIX-LONG-"
        old_part2 = "OLD-SUFFIX-AMOUNT-100"
        new_part2 = "NEW-SUFFIX-AMOUNT-200"
        new_body = new_part1 + new_part2
        lines = []
        for timestamp, part, index in (
            (old_timestamp, old_part2, 2),
            (new_timestamp, new_part1, 1),
        ):
            pdu = _incoming_ucs2_pdu(
                sender,
                part,
                timestamp,
                reference=reference,
                total=2,
                index=index,
            )
            lines.append(f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,162")
            lines.extend(_wrap_hex(pdu))
            lines.append("[I]-[TP-PID : ] 0 dcs:  8")
        lines.extend([
            f"[I]-[handler_sms.smsCallback] {sender} {new_timestamp} {new_body}",
            "[I]-[ril.proatc] OK",
        ])

        calls = replay_lines(lines)
        popups = [args[0] for args in _calls_by_name(calls, "sms_popup")]
        cloud_bodies = [args[1] for args in _calls_by_name(calls, "cloud_sms")]

        self.assertEqual(popups, [new_body])
        self.assertEqual(cloud_bodies, [new_body])
        self.assertNotIn(old_part2, popups[0])

    def test_replay_reuses_production_repeat_state_across_all_lines(self):
        sender = "10086"
        timestamp = "26/07/22,19:20:00+32"
        body = "DUPLICATE-MESSAGE-BODY"
        lines = []
        for index in (1, 2):
            pdu = _incoming_ucs2_pdu(sender, body, timestamp)
            lines.extend([
                f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,80",
                *_wrap_hex(pdu),
                "[I]-[TP-PID : ] 0 dcs:  8",
                f"[I]-[handler_sms.smsCallback] {sender} {timestamp} {body}",
                "[I]-[ril.proatc] OK",
            ])

        calls = replay_lines(lines)
        popups = [args[0] for args in _calls_by_name(calls, "sms_popup")]

        self.assertEqual(popups, [body])

    def test_same_timestamp_shorter_followup_is_not_rewritten_or_suppressed(self):
        sender = "10001"
        timestamp = "26/07/22,19:40:00+32"
        new_body = "ACCOUNT-NOTICE-LONG-COMMON-PREFIX"
        old_body = new_body + "-OLD-EXTRA-CONTENT"
        lines = []
        for index, body in enumerate((old_body, new_body), start=1):
            pdu = _incoming_ucs2_pdu(sender, body, timestamp)
            lines.extend([
                f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,80",
                *_wrap_hex(pdu),
                "[I]-[TP-PID : ] 0 dcs:  8",
                f"[I]-[handler_sms.smsCallback] {sender} {timestamp} {body}",
                "[I]-[ril.proatc] OK",
            ])

        calls = replay_lines(lines)
        popups = [args[0] for args in _calls_by_name(calls, "sms_popup")]
        cloud_bodies = [args[1] for args in _calls_by_name(calls, "cloud_sms")]

        self.assertEqual(popups, [old_body, new_body])
        self.assertEqual(cloud_bodies, [old_body, new_body])

    def test_interleaved_near_timestamp_concat_reuse_does_not_cross_messages(self):
        sender = "10001"
        reference = 0x73
        timestamp_a = "26/07/22,19:50:00+32"
        timestamp_b = "26/07/22,19:50:01+32"
        lines = []

        for slot, (body, index, timestamp) in enumerate((
            ("A1", 1, timestamp_a),
            ("B2", 2, timestamp_b),
            ("A2", 2, timestamp_a),
            ("B1", 1, timestamp_b),
        ), start=1):
            pdu = _incoming_ucs2_pdu(
                sender,
                body,
                timestamp,
                reference=reference,
                total=2,
                index=index,
            )
            lines.extend([
                f"[I]-[lib_sms rsp] +CMGR AT+CMGR={slot} true OK +CMGR: 0,,80",
                *_wrap_hex(pdu),
                "[I]-[TP-PID : ] 0 dcs:  8",
                f"[I]-[handler_sms.smsCallback] {sender} {timestamp} {body}",
                "[I]-[ril.proatc] OK",
            ])

        calls = replay_lines(lines)
        popups = [args[0] for args in _calls_by_name(calls, "sms_popup")]
        cloud_bodies = [args[1] for args in _calls_by_name(calls, "cloud_sms")]

        self.assertEqual(popups, ["A1A2", "B1B2"])
        self.assertEqual(cloud_bodies, popups)
        self.assertNotIn("A1B2", popups)
        self.assertNotIn("B1A2", popups)

    def test_mixed_lua_callbacks_are_deferred_until_pdu_timestamps_disambiguate(self):
        sender = "10001"
        reference = 0x74
        timestamp_a = "26/07/22,19:51:00+32"
        timestamp_b = "26/07/22,19:51:01+32"
        a1 = "MESSAGE-A-FIRST-SEGMENT-LONG-"
        a2 = "MESSAGE-A-SECOND-END"
        b1 = "MESSAGE-B-FIRST-SEGMENT-LONG-"
        b2 = "MESSAGE-B-SECOND-END"
        lines = []

        def append_pdu(slot, body, index, timestamp):
            pdu = _incoming_ucs2_pdu(
                sender,
                body,
                timestamp,
                reference=reference,
                total=2,
                index=index,
            )
            lines.extend([
                f"[I]-[lib_sms rsp] +CMGR AT+CMGR={slot} true OK +CMGR: 0,,80",
                *_wrap_hex(pdu),
                "[I]-[TP-PID : ] 0 dcs:  8",
            ])

        append_pdu(1, a1, 1, timestamp_a)
        append_pdu(2, b2, 2, timestamp_b)
        lines.extend([
            f"[I]-[handler_sms.smsCallback] {sender} {timestamp_a} {a1 + b2}",
            "[I]-[ril.proatc] OK",
        ])
        append_pdu(3, a2, 2, timestamp_a)
        append_pdu(4, b1, 1, timestamp_b)
        lines.extend([
            f"[I]-[handler_sms.smsCallback] {sender} {timestamp_a} {b1 + a2}",
            "[I]-[ril.proatc] OK",
        ])

        calls = replay_lines(lines)
        popups = [args[0] for args in _calls_by_name(calls, "sms_popup")]
        cloud_bodies = [args[1] for args in _calls_by_name(calls, "cloud_sms")]

        self.assertEqual(popups, [a1 + a2, b1 + b2])
        self.assertEqual(cloud_bodies, popups)
        self.assertNotIn(a1 + b2, popups)
        self.assertNotIn(b1 + a2, popups)

    def test_isolated_two_part_timestamp_variance_releases_after_ambiguity_grace(self):
        sender = "10001"
        reference = 0x75
        timestamp_1 = "26/07/22,19:52:00+32"
        timestamp_2 = "26/07/22,19:52:01+32"
        part1 = "LEGITIMATE-FIRST-SEGMENT-LONG-"
        part2 = "LEGITIMATE-SECOND-END"
        body = part1 + part2
        lines = []

        for slot, (part, index, timestamp) in enumerate((
            (part1, 1, timestamp_1),
            (part2, 2, timestamp_2),
        ), start=1):
            pdu = _incoming_ucs2_pdu(
                sender,
                part,
                timestamp,
                reference=reference,
                total=2,
                index=index,
            )
            lines.extend([
                f"[I]-[lib_sms rsp] +CMGR AT+CMGR={slot} true OK +CMGR: 0,,80",
                *_wrap_hex(pdu),
                "[I]-[TP-PID : ] 0 dcs:  8",
            ])
        lines.extend([
            f"[I]-[handler_sms.smsCallback] {sender} {timestamp_1} {body}",
            "[I]-[ril.proatc] OK",
        ])

        calls = replay_lines(lines)
        popups = [args[0] for args in _calls_by_name(calls, "sms_popup")]
        cloud_bodies = [args[1] for args in _calls_by_name(calls, "cloud_sms")]

        self.assertEqual(popups, [body])
        self.assertEqual(cloud_bodies, [body])


if __name__ == "__main__":
    unittest.main()
