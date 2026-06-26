import unittest

from sms_core.serial_sms import (
    SMS_CALLBACK_PREFIX,
    SmsPendingCollector,
    flush_pending_sms,
    handle_sms_collector_line,
)


def parse_head(text):
    return ("+8613812345678", text.split(" ", 1)[1])


class SerialSmsCollectorDecisionTests(unittest.TestCase):
    def test_handle_sms_collector_line_starts_new_callback(self):
        collector = SmsPendingCollector(parse_head)
        flushed = []

        decision = handle_sms_collector_line(
            collector,
            SMS_CALLBACK_PREFIX + " +8613812345678 first",
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
            '[I]-[websocket] json: {"message":"[I]-[handler_sms.smsCallback] +8613812345678 first"}',
            now=10.0,
            flush_callback=lambda: flushed.append("flush"),
        )

        self.assertFalse(decision.started)
        self.assertEqual(decision.action, "pass")
        self.assertFalse(collector.active)
        self.assertEqual(flushed, [])

    def test_handle_sms_collector_line_flushes_previous_before_new_callback(self):
        collector = SmsPendingCollector(parse_head)
        collector.start("+8613812345678 first", now=10.0)
        flushed = []

        decision = handle_sms_collector_line(
            collector,
            SMS_CALLBACK_PREFIX + " +8613812345678 second",
            now=11.0,
            flush_callback=lambda: flushed.append("flush"),
        )

        self.assertTrue(decision.started)
        self.assertTrue(decision.flushed)
        self.assertEqual(flushed, ["flush"])
        self.assertEqual(collector.callback_head, "+8613812345678 second")

    def test_handle_sms_collector_line_consumes_fragment_and_boundary(self):
        collector = SmsPendingCollector(parse_head, max_follow_lines=3)
        collector.start("+8613812345678 first", now=10.0)
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

    def test_flush_pending_sms_delegates_to_sms_processing(self):
        collector = SmsPendingCollector(parse_head)
        collector.start("+8613812345678 hello", now=10.0)
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
        self.assertIn(("cloud", "+8613812345678 hello", "hello"), calls)
        self.assertIn(("popup", "hello"), calls)
        self.assertIn(("alert",), calls)


if __name__ == "__main__":
    unittest.main()
