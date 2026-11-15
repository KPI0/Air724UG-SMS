import unittest

import base64

from sms_core.serial_sms import (
    PendingSms,
    SMS_CALLBACK_PREFIX,
    SmsPendingCollector,
    flush_pending_sms,
    handle_sms_collector_line,
)
def parse_head(text):
    return ("+8613123123123", text.split(" ", 1)[1])


def framed_lines(body, sender="10086", timestamp="26/11/14,10:00:00+32", frame_id=1):
    payload = sender.encode("utf-8") + b"\0" + timestamp.encode("utf-8") + b"\0" + body.encode("utf-8")
    chunks = [payload[index:index + 24] for index in range(0, len(payload), 24)]
    return [f"@@SMS_V1_BEGIN|{frame_id}|{len(payload)}|{len(chunks)}",
            *[f"@@SMS_V1_CHUNK|{frame_id}|{index}|{base64.b64encode(chunk).decode('ascii')}"
              for index, chunk in enumerate(chunks, 1)],
            f"@@SMS_V1_END|{frame_id}|{len(chunks)}"]


class SerialSmsFramedCollectorTests(unittest.TestCase):
    def feed(self, collector, lines, collected, diagnostics=None):
        decisions = []
        for index, line in enumerate(lines):
            kwargs = {} if diagnostics is None else {"diagnostic_callback": diagnostics.append}
            decisions.append(handle_sms_collector_line(
                collector, line, 10.0 + index / 100,
                flush_callback=lambda: collected.append(collector.flush()), **kwargs))
        return decisions

    def test_complete_frame_delivers_only_after_end_and_preserves_exact_body(self):
        body = " \r\n[I]-[ril.proatc] IMEI: 882132456789012\r\nOK\nRING\n@@CALL_RECORD_BEGIN|body\n\0尾行  \n"
        lines = framed_lines(body)
        collector, collected = SmsPendingCollector(), []
        decisions = self.feed(collector, lines[:-1], collected)
        self.assertEqual(collected, [])
        self.assertTrue(all(item.action in ("start", "consumed") for item in decisions))
        end = self.feed(collector, lines[-1:], collected)[0]
        self.assertEqual(end.action, "consumed")
        self.assertTrue(end.flushed)
        self.assertEqual([item.framed_body for item in collected], [body])
        self.assertEqual(collected[0].raw_lines, [])
        self.assertFalse(collector.active)

    def test_empty_unicode_and_nul_body_without_business_parser(self):
        def forbidden(_text):
            raise AssertionError("collector must not parse SMS business content")

        for body in ("", "\n\n  ", "中文👉\r\n末行\0正文", "a" * 80):
            with self.subTest(body=repr(body)):
                collector, collected = SmsPendingCollector(forbidden), []
                self.feed(collector, framed_lines(body), collected)
                self.assertEqual([item.framed_body for item in collected], [body])

    def test_true_logs_urcs_and_recordings_pass_during_frame(self):
        collector, collected = SmsPendingCollector(), []
        lines = framed_lines("framed body")
        self.feed(collector, lines[:2], collected)
        real_lines = ["[I]-[ril.proatc] OK", "000000000000001", "RING", "+CMTI: \"SM\",1",
                      "ordinary log", "@@CALL_RECORD_BEGIN|real", "AT+CGSN"]
        decisions = self.feed(collector, real_lines, collected)
        self.assertEqual([item.action for item in decisions], ["pass"] * len(real_lines))
        self.assertEqual(collected, [])
        self.feed(collector, lines[2:], collected)
        self.assertEqual([item.framed_body for item in collected], ["framed body"])

    def test_legacy_body_keeps_complete_frame_text_and_real_boundary(self):
        collector, collected = SmsPendingCollector(), []
        collector.start("10086 legacy first", 10.0)
        lines = framed_lines("embedded text")
        decisions = self.feed(collector, lines, collected)
        self.assertEqual([item.action for item in decisions], ["consumed"] * len(lines))
        self.assertEqual(collected, [])
        boundary = self.feed(collector, ["[I]-[ril.proatc] OK"], collected)[0]
        self.assertEqual(boundary.action, "boundary")
        self.assertEqual(collected[0].raw_lines, lines)
        self.assertIsNone(getattr(collected[0], "framed_body", None))
        self.feed(collector, framed_lines("new body", frame_id=2), collected)
        self.assertEqual(collected[1].framed_body, "new body")

    def test_reserved_orphan_unknown_and_invalid_header_are_owned(self):
        lines = ["@@SMS_V1_CHUNK|1|1|YQ==", "@@SMS_V1_END|1|1", "@@SMS_V1_UNKNOWN|private",
                 "@@SMS_V1_BEGIN|0|3|1", "@@SMS_V1_BEGIN|1000000000|3|1",
                 "@@SMS_V1_BEGIN|1|0|0", "@@SMS_V1_BEGIN|1|131073|5462",
                 "@@SMS_V1_BEGIN|1|25|1", "@@SMS_V1_BEGIN|1|3|2",
                 "@@SMS_V1_BEGIN|x|3|1", "@@SMS_V1_BEGIN|1|3|1|extra"]
        for line in lines:
            with self.subTest(line=line):
                collector, collected, diagnostics = SmsPendingCollector(), [], []
                decision = self.feed(collector, [line], collected, diagnostics)[0]
                self.assertEqual(decision.action, "consumed")
                self.assertFalse(collector.active)
                self.assertEqual(collected, [])
                self.assertEqual(len(diagnostics), 1)
                self.assertNotIn("private", diagnostics[0])

    def test_invalid_chunks_and_end_drop_partial_then_recover(self):
        valid = framed_lines("original body")
        failures = [[valid[0], "@@SMS_V1_CHUNK|2|1|YQ=="], [valid[0], valid[2]],
                    [valid[0], valid[1], valid[1]], [valid[0], "@@SMS_V1_CHUNK|1|1|%%%"],
                    [valid[0], "@@SMS_V1_CHUNK|1|1|YQ=="],
                    [valid[0], "@@SMS_V1_CHUNK|1|1|" + "A" * 36],
                    [valid[0], "@@SMS_V1_CHUNK|1|1"],
                    [*valid[:-1], "@@SMS_V1_END|1|999"],
                    [*valid[:-1], "@@SMS_V1_END|2|2"],
                    [*valid[:-1], valid[-1] + "|extra"], [valid[0], valid[-1]]]
        for lines in failures:
            with self.subTest(lines=lines):
                collector, collected, diagnostics = SmsPendingCollector(), [], []
                self.feed(collector, lines, collected, diagnostics)
                self.assertEqual(collected, [])
                self.assertFalse(collector.active)
                self.assertTrue(diagnostics)
                self.feed(collector, framed_lines("recovered", frame_id=3), collected, diagnostics)
                self.assertEqual([item.framed_body for item in collected], ["recovered"])

    def test_noncanonical_base64_drops_before_payload_decode(self):
        collector, collected, diagnostics = SmsPendingCollector(), [], []
        self.feed(collector, ["@@SMS_V1_BEGIN|1|5|1", "@@SMS_V1_CHUNK|1|1|YQBiAGN="],
                  collected, diagnostics)
        self.assertEqual(collected, [])
        self.assertFalse(collector.active)
        self.assertTrue(diagnostics)

    def test_exact_transport_and_metadata_size_limits_are_accepted(self):
        sender, timestamp = "s" * 128, "t" * 64
        body = "x" * (131072 - len(sender) - len(timestamp) - 2)
        collector, collected = SmsPendingCollector(), []
        self.feed(collector, framed_lines(body, sender, timestamp, 999999999), collected)
        self.assertEqual([item.framed_body for item in collected], [body])
        self.assertEqual(collected[0].callback_text, f"{sender} {timestamp} {body}")
        self.assertFalse(collector.active)

    def test_invalid_utf8_and_field_lengths_drop_transport(self):
        payloads = [b"a\0b\0\xff", b"a\0missing-second-separator",
                    b"s" * 129 + b"\0t\0body", b"s\0" + b"t" * 65 + b"\0body"]
        for payload in payloads:
            with self.subTest(payload_length=len(payload)):
                chunks = [payload[index:index + 24] for index in range(0, len(payload), 24)]
                lines = [f"@@SMS_V1_BEGIN|1|{len(payload)}|{len(chunks)}",
                         *[f"@@SMS_V1_CHUNK|1|{index}|{base64.b64encode(chunk).decode('ascii')}"
                           for index, chunk in enumerate(chunks, 1)], f"@@SMS_V1_END|1|{len(chunks)}"]
                collector, collected, diagnostics = SmsPendingCollector(), [], []
                self.feed(collector, lines, collected, diagnostics)
                self.assertEqual(collected, [])
                self.assertFalse(collector.active)
                self.assertTrue(diagnostics)

    def test_new_begin_discards_partial_and_completed_frames_stay_independent(self):
        collector, collected, diagnostics = SmsPendingCollector(), [], []
        self.feed(collector, framed_lines("discarded")[:2], collected, diagnostics)
        self.feed(collector, framed_lines("first", frame_id=2), collected, diagnostics)
        self.feed(collector, framed_lines("second", frame_id=3), collected, diagnostics)
        self.assertEqual([item.framed_body for item in collected], ["first", "second"])
        self.assertTrue(diagnostics)

    def test_flush_and_reset_do_not_publish_partial_or_use_legacy_idle_timer(self):
        for operation in ("flush", "reset"):
            with self.subTest(operation=operation):
                collector, collected, diagnostics = SmsPendingCollector(), [], []
                self.feed(collector, framed_lines("partial")[:2], collected, diagnostics)
                self.assertFalse(collector.expired(1000.0))
                self.assertIsNone(getattr(collector, operation)())
                self.assertFalse(collector.active)
                self.feed(collector, framed_lines("complete", frame_id=2), collected, diagnostics)
                self.assertEqual([item.framed_body for item in collected], ["complete"])

    def test_embedded_prefix_in_ordinary_log_does_not_claim_ownership(self):
        collector, collected = SmsPendingCollector(), []
        decision = self.feed(collector, ["[I]-[probe] @@SMS_V1_BEGIN|1|3|1"], collected)[0]
        self.assertEqual(decision.action, "pass")
        self.assertFalse(collector.active)


class SerialSmsCollectorDecisionTests(unittest.TestCase):
    def test_handle_sms_collector_line_starts_new_callback(self):
        collector = SmsPendingCollector(parse_head)
        flushed = []

        decision = handle_sms_collector_line(
            collector,
            SMS_CALLBACK_PREFIX + " +8613123123123 first",
            now=10.0,
            flush_callback=lambda: flushed.append("flush"),
        )

        self.assertTrue(decision.started)
        self.assertEqual(decision.action, "start")
        self.assertTrue(decision.continue_read)
        self.assertFalse(decision.flushed)
        self.assertTrue(collector.active)
        self.assertEqual(flushed, [])

    def test_handle_sms_collector_line_ignores_websocket_json_payload(self):
        collector = SmsPendingCollector(parse_head)
        flushed = []

        decision = handle_sms_collector_line(
            collector,
            '[I]-[websocket] json: {"message":"[I]-[handler_sms.smsCallback] +8613123123123 first"}',
            now=10.0,
            flush_callback=lambda: flushed.append("flush"),
        )

        self.assertFalse(decision.started)
        self.assertEqual(decision.action, "pass")
        self.assertFalse(collector.active)
        self.assertEqual(flushed, [])

    def test_handle_sms_collector_line_flushes_previous_before_new_callback(self):
        collector = SmsPendingCollector(parse_head)
        collector.start("+8613123123123 first", now=10.0)
        flushed = []

        decision = handle_sms_collector_line(
            collector,
            SMS_CALLBACK_PREFIX + " +8613123123123 second",
            now=11.0,
            flush_callback=lambda: flushed.append("flush"),
        )

        self.assertTrue(decision.started)
        self.assertTrue(decision.flushed)
        self.assertEqual(flushed, ["flush"])
        self.assertEqual(collector.callback_head, "+8613123123123 second")

    def test_handle_sms_collector_line_consumes_fragment_and_boundary(self):
        collector = SmsPendingCollector(parse_head)
        collector.start("+8613123123123 first", now=10.0)
        flushed = []

        consumed = handle_sms_collector_line(
            collector,
            "fragment",
            now=10.1,
            flush_callback=lambda: flushed.append("flush"),
        )
        self.assertEqual(consumed.action, "consumed")
        self.assertTrue(consumed.continue_read)
        self.assertEqual(flushed, [])

        boundary = handle_sms_collector_line(
            collector,
            "[I]-[ril] next",
            now=10.2,
            flush_callback=lambda: flushed.append("flush"),
        )
        self.assertEqual(boundary.action, "boundary")
        self.assertFalse(boundary.continue_read)
        self.assertEqual(flushed, ["flush"])

    def test_handle_sms_collector_line_keeps_plus_digit_continuation(self):
        collector = SmsPendingCollector(parse_head)
        collector.start("+8613123123123 first", now=10.0)
        flushed = []

        consumed = handle_sms_collector_line(
            collector,
            "+100.00 元到账",
            now=10.1,
            flush_callback=lambda: flushed.append("flush"),
        )

        self.assertEqual(consumed.action, "consumed")
        self.assertEqual(collector.raw_lines, ["+100.00 元到账"])
        self.assertEqual(flushed, [])

        boundary = handle_sms_collector_line(
            collector,
            "+CMTI: \"SM\",1",
            now=10.2,
            flush_callback=lambda: flushed.append("flush"),
        )

        self.assertEqual(boundary.action, "boundary")
        self.assertEqual(flushed, ["flush"])

    def test_handle_sms_collector_line_treats_star_urc_as_boundary(self):
        collector = SmsPendingCollector(parse_head)
        collector.start("+8613123123123 first", now=10.0)
        flushed = []

        boundary = handle_sms_collector_line(
            collector,
            "*CGEV: DEACT,6",
            now=10.1,
            flush_callback=lambda: flushed.append("flush"),
        )

        self.assertEqual(boundary.action, "boundary")
        self.assertFalse(boundary.continue_read)
        self.assertEqual(collector.raw_lines, [])
        self.assertEqual(flushed, ["flush"])

    def test_handle_sms_collector_line_keeps_unrecognized_star_content(self):
        collector = SmsPendingCollector(parse_head)
        collector.start("+8613123123123 first", now=10.0)
        flushed = []

        lines = ["*NOTICE=hello", "*CGEV", "*CGEV=body", "*cgev: body"]
        for index, line in enumerate(lines, start=1):
            consumed = handle_sms_collector_line(
                collector,
                line,
                now=10.0 + index / 10,
                flush_callback=lambda: flushed.append("flush"),
            )
            self.assertEqual(consumed.action, "consumed")
            self.assertTrue(consumed.continue_read)

        self.assertEqual(collector.raw_lines, lines)
        self.assertEqual(flushed, [])

    def test_collector_flush_returns_raw_callback_output(self):
        collector = SmsPendingCollector(parse_head)
        collector.start("+8613123123123 first", now=10.0)

        self.assertEqual(collector.consume_line("follow line", now=10.1), "consumed")
        collected = collector.flush()

        self.assertEqual(collected.callback_text, "+8613123123123 first")
        self.assertEqual(collected.raw_lines, ["follow line"])

    def test_collector_does_not_parse_sms_content_while_buffering(self):
        def fail_parse(_text):
            raise AssertionError("collector must not parse SMS content")

        collector = SmsPendingCollector(fail_parse)

        self.assertTrue(collector.start("+8613123123123 first", now=10.0))
        self.assertEqual(collector.consume_line("follow line", now=10.1), "consumed")
        collected = collector.flush()

        self.assertEqual(collected.callback_text, "+8613123123123 first")
        self.assertEqual(collected.raw_lines, ["follow line"])

    def test_collector_does_not_cap_raw_lines(self):
        collector = SmsPendingCollector(parse_head)
        collector.start("+8613123123123 first", now=10.0)

        self.assertEqual(collector.consume_line("line1", now=10.1), "consumed")
        self.assertEqual(collector.consume_line("line2", now=10.2), "consumed")
        self.assertEqual(collector.consume_line("line3", now=10.3), "consumed")
        collected = collector.flush()

        self.assertEqual(collected.raw_lines, ["line1", "line2", "line3"])

    def test_flush_pending_sms_delegates_to_sms_processing(self):
        collector = SmsPendingCollector(parse_head)
        collector.start("+8613123123123 hello", now=10.0)
        calls = []

        result = flush_pending_sms(
            collector,
            keywords=["hello"],
            log_unmatched_sms=False,
            log_dir=".",
            log_prefix="COM5",
            ignore_repeat_state={},
            error_repeat_limit=3,
            enqueue_push=lambda msg: calls.append(("push", msg)),
            send_cloud_sms_event=lambda head, msg: calls.append(("cloud", head, msg)),
            port_ui=lambda text, level: calls.append(("ui", text, level)),
            play_alert=lambda: calls.append(("alert",)),
            show_sms_popup=lambda msg: calls.append(("popup", msg)),
            file_log=lambda item: calls.append(("file", item)),
            system_ui=lambda text, level: calls.append(("system", text, level)),
        )

        self.assertEqual(result, "shown")
        self.assertFalse(collector.active)
        self.assertIn(("push", "hello"), calls)
        self.assertIn(("cloud", "+8613123123123 hello", "hello"), calls)
        self.assertIn(("popup", "hello"), calls)
        self.assertIn(("alert",), calls)

    def test_flush_pending_sms_processes_multiple_completed_messages(self):
        collector = SmsPendingCollector(parse_head)
        collector.start("+8613123123123 trigger", now=10.0)
        calls = []

        class MultiAssembler:
            def add_collected(self, *_args, **_kwargs):
                return [
                    PendingSms("+8613123123123 first", "first", []),
                    PendingSms("+8613123123123 second", "second", []),
                ]

        result = flush_pending_sms(
            collector,
            keywords=[],
            log_unmatched_sms=False,
            log_dir=".",
            log_prefix="COM5",
            ignore_repeat_state={},
            error_repeat_limit=3,
            enqueue_push=lambda msg: calls.append(("push", msg)),
            send_cloud_sms_event=lambda head, msg: calls.append(("cloud", head, msg)),
            port_ui=lambda text, level: calls.append(("ui", text, level)),
            play_alert=lambda: calls.append(("alert",)),
            show_sms_popup=lambda msg: calls.append(("popup", msg)),
            file_log=lambda item: calls.append(("file", item)),
            system_ui=lambda text, level: calls.append(("system", text, level)),
            assembler=MultiAssembler(),
        )

        self.assertEqual(result, "shown")
        self.assertIn(("push", "first"), calls)
        self.assertIn(("push", "second"), calls)


if __name__ == "__main__":
    unittest.main()
