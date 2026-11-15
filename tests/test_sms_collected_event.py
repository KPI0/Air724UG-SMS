import unittest
from types import SimpleNamespace

from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.long_sms_assembler import LongSmsAssembler
from sms_core.serial_sms import CollectedSmsCallback
from sms_core.serial_sms_pdu_cache import SmsPduCorrectionCache
from sms_core.sms_collected_event import pending_from_collected
from sms_core.sms_processing import parse_sms_callback_metadata, process_pending_sms
from sms_core.sms_receive_pipeline import SmsReceivePipeline
from tools.sms_replay import _incoming_ucs2_pdu, _wrap_hex


TIMESTAMP = "26/11/14,10:00:00+32"


def framed_collected(body, sender="10086"):
    return SimpleNamespace(callback_text=f"{sender} {TIMESTAMP} {body}", raw_lines=[], framed_body=body)


class FramedSmsCollectedEventTests(unittest.TestCase):
    def test_exact_body_keeps_whitespace_and_empty_body(self):
        for body in ("", "\r\n  \n", "  first\r\n\nlast  \n", "[I]-[ril.proatc] OK\n\0正文"):
            with self.subTest(body=repr(body)):
                pending = pending_from_collected(framed_collected(body), parse_sms_callback_head)
                self.assertEqual(pending.full_msg, body)
                self.assertTrue(pending.callback_head.startswith("10086 " + TIMESTAMP))

    def test_legacy_two_argument_callback_remains_compatible(self):
        collected = CollectedSmsCallback("10086 " + TIMESTAMP + " first ", ["second "])
        self.assertIsNone(getattr(collected, "framed_body", None))
        self.assertEqual(pending_from_collected(collected, parse_sms_callback_head).full_msg, "first\nsecond")

    def test_exact_body_reaches_push_cloud_popup_and_metadata(self):
        body = " \r\n[I]-[ril.proatc] IMEI: 882132456789012\nNO CARRIER\n末行  \n"
        pipeline = SmsReceivePipeline(parse_sms_callback_head, SmsPduCorrectionCache(),
                                      LongSmsAssembler(parse_sms_callback_head))
        pending = pipeline.add_collected(framed_collected(body, "Synthetic Sender"), now=10.0)
        pushes, clouds, popups, ui = [], [], [], []
        process_pending_sms(pending, [], False, "", "SYNTHETIC", {}, 3,
                            lambda message, variables=None: pushes.append((message, variables)),
                            lambda head, message, metadata=None: clouds.append((head, message, metadata)),
                            lambda *args: ui.append(args), lambda: None, popups.append,
                            lambda _item: None, lambda *_args: None)
        self.assertEqual(pushes[0][0], body)
        self.assertEqual(pushes[0][1]["sender"], "Synthetic Sender")
        self.assertEqual(clouds[0][1], body)
        self.assertEqual(clouds[0][2]["sms_time_identity"], "2026-11-14 10:00:00+32")
        self.assertEqual(popups, [body])
        self.assertIn(("[I]-[ril.proatc] IMEI: 882132456789012", "sms"), ui)

    def test_empty_body_keeps_original_no_push_or_popup_behavior(self):
        pending = pending_from_collected(framed_collected(""), parse_sms_callback_head)
        forwarded, system = [], []
        result = process_pending_sms(pending, [], False, "", "SYNTHETIC", {}, 3,
                                     lambda *args, **kwargs: forwarded.append(args),
                                     lambda *args, **kwargs: forwarded.append(args),
                                     lambda *args: forwarded.append(args), lambda: forwarded.append("alert"),
                                     lambda *args: forwarded.append(args), lambda _item: None,
                                     lambda *args: system.append(args))
        self.assertEqual(forwarded, [])
        self.assertEqual(result, "ignored")
        self.assertTrue(system)

    def test_single_pdu_correction_still_corrects_content(self):
        class CorrectingCache:
            def correct_single_pdu_callback_text(self, _text, _parse, _now):
                return "10086 " + TIMESTAMP + " recovered PDU body"

        pending = pending_from_collected(framed_collected("corrupted body�"), parse_sms_callback_head,
                                         correction_cache=CorrectingCache(), now=10.0)
        self.assertEqual(pending.full_msg, "recovered PDU body")

    def test_concat_pdus_keep_existing_assembler_and_trace_metadata(self):
        cache, assembler = SmsPduCorrectionCache(), LongSmsAssembler(parse_sms_callback_head)
        pipeline = SmsReceivePipeline(parse_sms_callback_head, cache, assembler)
        result = None
        for index, body in ((1, "FIRST-"), (2, "SECOND")):
            moment = 10.0 + index
            lines = [f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,80",
                     *_wrap_hex(_incoming_ucs2_pdu("10086", body, TIMESTAMP, reference=42, total=2, index=index)),
                     "[I]-[TP-PID : ] 0 dcs:  8"]
            for line in lines:
                cache.observe_line(line, moment)
            result = pipeline.add_collected(framed_collected(body), now=moment + 0.1)
            if index == 1:
                self.assertIsNone(result)
        self.assertEqual(result.full_msg, "FIRST-SECOND")
        self.assertEqual(result.concat_reference, 42)
        self.assertEqual(result.concat_total, 2)
        self.assertRegex(result.message_trace_id, r"^[0-9a-f]{12}$")
        self.assertEqual(assembler._pending, {})

    def test_framed_concat_whitespace_still_matches_exact_pdu_segments(self):
        cache, assembler = SmsPduCorrectionCache(), LongSmsAssembler(parse_sms_callback_head)
        pipeline = SmsReceivePipeline(parse_sms_callback_head, cache, assembler)
        parts = ("\r\n  FIRST-  ", " SECOND\n")
        result = None
        for index, body in enumerate(parts, 1):
            moment = 10.0 + index
            lines = [f"[I]-[lib_sms rsp] +CMGR AT+CMGR={index} true OK +CMGR: 0,,80",
                     *_wrap_hex(_incoming_ucs2_pdu("10086", body, TIMESTAMP, reference=43, total=2, index=index)),
                     "[I]-[TP-PID : ] 0 dcs:  8"]
            for line in lines:
                cache.observe_line(line, moment)
            result = pipeline.add_collected(framed_collected(body), now=moment + 0.1)
            if index == 1:
                self.assertIsNone(result, "first PDU segment must not be emitted as a whole SMS")
        self.assertEqual(result.full_msg, "".join(parts))
        self.assertEqual(result.concat_reference, 43)
        self.assertEqual(result.concat_total, 2)
        self.assertEqual(parse_sms_callback_metadata(result.callback_head)["sms_time_identity"],
                         "2026-11-14 10:00:00+32")

    def test_single_pdu_sender_correction_keeps_exact_framed_body(self):
        body = " \r\nexact content with whitespace  \n"

        class SenderCorrectingCache:
            def correct_single_pdu_callback_text(self, _text, _parse, _now):
                return "Synthetic Sender " + TIMESTAMP + " " + body

        pending = pending_from_collected(framed_collected(body, "legacy-alias"), parse_sms_callback_head,
                                         correction_cache=SenderCorrectingCache(), now=10.0)
        self.assertEqual(pending.full_msg, body)
        self.assertTrue(pending.callback_head.startswith("Synthetic Sender " + TIMESTAMP))


if __name__ == "__main__":
    unittest.main()
