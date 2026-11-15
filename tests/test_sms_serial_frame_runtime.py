"""SMS transport must preserve content without executing it as serial telemetry."""
import base64
import unittest
from dataclasses import replace

from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.cloud_state_namespace_runtime import request_cloud_device_imei_namespace_runtime
from sms_core.serial_runtime import SerialLineDecoder, SerialRuntimeState, handle_serial_runtime_line
from tests import test_cloud_imei_response_classification as imei_fixtures
from tests.test_serial_initialization import InitializationHarness
from tests.test_serial_runtime import runtime_callbacks, runtime_config
from tests.test_serial_startup_406_compatibility import configure, wait_for

IMEI = imei_fixtures.IMEI


def sms_frames(body, sender="10086", timestamp="26/11/14,10:00:00+32", frame_id=1):
    payload = (sender + "\0" + timestamp + "\0" + body).encode("utf-8")
    chunks = [payload[index:index + 24] for index in range(0, len(payload), 24)]
    return [f"@@SMS_V1_BEGIN|{frame_id}|{len(payload)}|{len(chunks)}"] + [
        f"@@SMS_V1_CHUNK|{frame_id}|{index}|" + base64.b64encode(chunk).decode("ascii")
        for index, chunk in enumerate(chunks, 1)
    ] + [f"@@SMS_V1_END|{frame_id}|{len(chunks)}"]


class SmsSerialFrameRuntimeTests(unittest.TestCase):
    def test_encoded_log_body_preserves_raw_query_window_and_real_imei(self):
        for body_line in ("[I]-[ril.proatc] IMEI: 882132456789012",
                          "[I]-[ril.proatc] 882132456789012"):
            with self.subTest(body_line=body_line):
                fixture = imei_fixtures.CloudImeiResponseClassificationTests("runTest")
                namespace, clock, _, calls, handle = fixture.make_runtime()
                namespace["cloud_imei_query_deadline"] = 0.0
                namespace.pop("_cloud_imei_query_owner", None)
                self.assertTrue(request_cloud_device_imei_namespace_runtime(namespace)[0])
                wait_for(lambda: not namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(),
                         "raw query worker did not finish")
                body = "SMS-BEGIN\n" + body_line + "\nOK\nRING\nNO CARRIER"
                for line in sms_frames(body):
                    handle(line)
                    fixture.assert_window_open(namespace)
                self.assertEqual(namespace["serial_obj"].writes, [b"AT+CGSN\r\n"])
                handle("[I]-[ril.proatc] " + IMEI)
                fixture.assert_identity(namespace)
                clock.value = 102.0
                handle("")
                self.assertIn(("cloud_sms", ("10086 26/11/14,10:00:00+32 " + body, body)), calls)
                self.assertIn(("sms_popup", (body,)), calls)
                self.assertFalse(any(call[0] == "call_popup" for call in calls))

    def test_production_reader_keeps_startup_and_framed_sms_ownership(self):
        harness = configure(InitializationHarness(hold="AT+CGSN"))
        sms_events = []
        harness.ns["_cloud_send_sms_event"] = lambda *args: sms_events.append(args)
        with harness:
            fixture = imei_fixtures.CloudImeiResponseClassificationTests("runTest")
            fixture.assert_startup_window(harness)
            body = "中文\r\n[I]-[ril.proatc] IMEI: 882132456789012\n尾行"
            harness.emit("\r\n".join(sms_frames(body)) + "\r\n[I]-[probe] FRAME-END\r\n")
            wait_for(lambda: "[I]-[probe] FRAME-END" in harness.lines, "reader did not finish frame")
            fixture.assert_window_open(harness.ns)
            self.assertEqual(harness.errors, [])
            harness.emit(IMEI + "\r\n")
            wait_for(lambda: harness.ns["CLOUD_DEVICE_IMEI"] == IMEI, "real identity not captured")
            harness.clock = 102.0
            wait_for(lambda: bool(sms_events), "complete SMS not delivered")
            self.assertEqual(sms_events, [("10086 26/11/14,10:00:00+32 " + body, body)])

    def test_byte_splits_preserve_crlf_unicode_and_literal_frame_text(self):
        body = " \r\n中文\x00\u2028@@SMS_V1_END|1|2\n[I]-[ril.proatc] 882132456789012\t "
        raw = ("\r\n".join(sms_frames(body)) + "\r\n").encode("ascii")
        for split in (1, 7, 23, 127, len(raw) - 1):
            with self.subTest(split=split):
                decoder, calls = SerialLineDecoder(), []
                state = SerialRuntimeState.create(parse_sms_callback_head)
                callbacks = runtime_callbacks(calls)
                for chunk in (raw[:split], raw[split:]):
                    for line in decoder.feed(chunk):
                        handle_serial_runtime_line(state, line, 100.0, "TEST", False,
                                                   runtime_config(), callbacks, {})
                handle_serial_runtime_line(state, "", 102.0, "TEST", False, runtime_config(), callbacks, {})
                self.assertIn(("sms_popup", (body,)), calls)
                self.assertFalse(any(call[0] == "imei" for call in calls))

    def test_real_terminals_and_recording_lines_pass_between_sms_chunks(self):
        calls, send_lines, recording_lines = [], [], []
        state = SerialRuntimeState.create(parse_sms_callback_head)
        callbacks = replace(runtime_callbacks(calls),
                            observe_sms_send_line=send_lines.append,
                            consume_call_recording_line=lambda line: recording_lines.append(line) or
                            line.startswith("@@CALL_RECORD"))
        frames = sms_frames("正文\nERROR\nRING")
        for line in frames[:2] + ["OK", "@@CALL_RECORD_CHUNK|a|1|YWJj", "ERROR"] + frames[2:]:
            handle_serial_runtime_line(state, line, 100.0, "TEST", False, runtime_config(), callbacks, {})
        handle_serial_runtime_line(state, "", 102.0, "TEST", False, runtime_config(), callbacks, {})
        self.assertIn("OK", send_lines)
        self.assertIn("ERROR", send_lines)
        self.assertFalse(any(line.startswith("@@SMS_V1_") for line in send_lines))
        self.assertIn("@@CALL_RECORD_CHUNK|a|1|YWJj", recording_lines)
        self.assertIn(("sms_popup", ("正文\nERROR\nRING",)), calls)

    def test_partial_frame_stop_never_publishes_sms(self):
        harness = configure(InitializationHarness(hold="AT+CGSN"))
        events = []
        harness.ns["_cloud_send_sms_event"] = lambda *args: events.append(args)
        with harness:
            harness.emit("\r\n".join(sms_frames("partial body")[:-1]) +
                         "\r\n[I]-[probe] PARTIAL-END\r\n")
            wait_for(lambda: "[I]-[probe] PARTIAL-END" in harness.lines, "partial frame not read")
        self.assertEqual(events, [])
        self.assertEqual(harness.errors, [])

    def test_malformed_frame_has_diagnostic_and_allows_next_sms(self):
        calls = []
        state = SerialRuntimeState.create(parse_sms_callback_head)
        callbacks = runtime_callbacks(calls)
        for line in ["@@SMS_V1_BEGIN|1|48|2", "@@SMS_V1_CHUNK|1|1|invalid!"] + sms_frames("恢复", frame_id=2):
            handle_serial_runtime_line(state, line, 100.0, "TEST", False, runtime_config(), callbacks, {})
        handle_serial_runtime_line(state, "", 102.0, "TEST", False, runtime_config(), callbacks, {})
        self.assertTrue(any(call[0] == "debug" and "FRAME DROP" in call[1][0] for call in calls))
        self.assertIn(("sms_popup", ("恢复",)), calls)


if __name__ == "__main__":
    unittest.main()
