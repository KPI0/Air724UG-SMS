"""IMEI responses must not consume phone numbers or SMS-owned content."""
import re
import threading
import unittest
from dataclasses import replace
from types import SimpleNamespace

from sms_core import serial_sender
from sms_core.cloud_protocol import normalize_imei, parse_sms_callback_head
from sms_core.cloud_state_namespace_runtime import (
    maybe_capture_cloud_device_imei_namespace_runtime,
    request_cloud_device_imei_namespace_runtime,
    set_cloud_device_imei_namespace_runtime,
)
from sms_core.serial_runtime import SerialRuntimeState, handle_serial_runtime_line
from sms_core.threading_runtime import WorkerThreadRegistry
from tests import test_cloud_imei_terminal_isolation as query_fixtures
from tests.test_serial_initialization import InitializationHarness
from tests.test_serial_runtime import runtime_callbacks, runtime_config
from tests.test_serial_startup_406_compatibility import configure, wait_for


IMEI = "000000000000001"
PHONE_NUMBERS = ("+88213245678901", "+882132456789012")
STARTUP_BYTES = [b"AT+CLIP=1\r\n", b"AT+CGSN\r\n", b"AT+CNUM\r\n"]


class CloudImeiResponseClassificationTests(unittest.TestCase):
    def make_runtime(self):
        clock = SimpleNamespace(value=100.0)
        logs, calls = [], []
        namespace = {
            "serial_lock": threading.RLock(), "serial_obj": query_fixtures.QueryModem(),
            "serial_connection_generation": 1, "serial_running": True,
            "serial_stop_event": threading.Event(), "TK_SHUTDOWN": threading.Event(),
            "CLOUD_DEVICE_IMEI": "", "cloud_imei_verified": False,
            "cloud_imei_query_deadline": 106.0, "_cloud_imei_query_owner": object(),
            "IMEI_REGEX": re.compile(r"\b(\d{14,17})\b"), "_normalize_imei": normalize_imei,
            "SERIAL_COMMAND_THREAD_REGISTRY": WorkerThreadRegistry(),
            "write_serial_command_result": serial_sender.write_serial_command_result,
            "_cloud_log": logs.append, "_push_serial_debug": lambda line: None,
            "_notify_cloud_identity_changed": lambda: None,
            "time": SimpleNamespace(monotonic=lambda: clock.value),
        }
        namespace["_set_cloud_device_imei"] = lambda value, **kwargs: (
            set_cloud_device_imei_namespace_runtime(namespace, value, **kwargs))
        callbacks = replace(
            runtime_callbacks(calls),
            capture_cloud_device_imei=lambda line: (
                maybe_capture_cloud_device_imei_namespace_runtime(namespace, line)),
            set_local_number=lambda value: namespace.update(LOCAL_NUMBER=value),
        )
        state = SerialRuntimeState.create(parse_sms_callback_head)

        def handle(line):
            return handle_serial_runtime_line(
                state, line, clock.value, "TEST", False, runtime_config(), callbacks, {},
            )

        return namespace, clock, state, calls, handle

    def assert_window_open(self, namespace, owner=None):
        self.assertEqual(namespace["CLOUD_DEVICE_IMEI"], "")
        self.assertFalse(namespace["cloud_imei_verified"])
        self.assertEqual(namespace["cloud_imei_query_deadline"], 106.0)
        if owner is not None:
            self.assertIs(namespace["_cloud_imei_query_owner"], owner)

    def assert_identity(self, namespace, expected=IMEI):
        self.assertEqual(namespace["CLOUD_DEVICE_IMEI"], expected)
        self.assertTrue(namespace["cloud_imei_verified"])
        self.assertEqual(namespace["cloud_imei_query_deadline"], 0.0)
        self.assertNotIn("_cloud_imei_query_owner", namespace)

    def drain_reader(self, harness, lines=""):
        sentinel = "[I]-[probe] CLASSIFICATION-END"
        harness.emit(lines + sentinel + "\r\n")
        wait_for(lambda: sentinel in harness.lines, "production reader did not process the input")
        harness.lines.clear()
        self.assertEqual(harness.errors, [])

    def assert_startup_window(self, harness):
        self.drain_reader(harness)
        self.assertEqual(harness.modem.writes, STARTUP_BYTES)
        self.assert_window_open(harness.ns)

    def test_production_reader_preserves_long_local_numbers_and_later_imei(self):
        for phone in PHONE_NUMBERS:
            for prefix in ("", "[I]-[ril.proatc] "):
                with self.subTest(phone=phone, prefix=prefix), configure(
                        InitializationHarness(hold="AT+CGSN")) as harness:
                    self.assert_startup_window(harness)
                    self.drain_reader(harness, prefix + '+CNUM: "","' + phone + '",145\r\n')
                    self.assertEqual(harness.ns["LOCAL_NUMBER"], phone)
                    self.assert_window_open(harness.ns)
                    self.drain_reader(harness, "[I]-[ril.proatc] " + IMEI + "\r\n")
                    self.assert_identity(harness.ns)
                    self.assertEqual(harness.ns["LOCAL_NUMBER"], phone)

    def test_production_reader_preserves_long_incoming_calls_and_later_imei(self):
        for phone in PHONE_NUMBERS:
            for prefix in ("", "[I]-[ril.proatc] "):
                with self.subTest(phone=phone, prefix=prefix):
                    harness = configure(InitializationHarness(hold="AT+CGSN"))
                    popups, events = [], []
                    harness.ns["show_call_popup"] = lambda number, **kwargs: popups.append(number)
                    harness.ns["_cloud_send_call_event"] = lambda number, message, **kwargs: events.append(number)
                    with harness:
                        self.assert_startup_window(harness)
                        self.drain_reader(harness, 'RING\r\n' + prefix + '+CLIP: "' + phone + '",145\r\n')
                        self.assertEqual(popups, [phone])
                        self.assertEqual(events, [phone])
                        self.assert_window_open(harness.ns)
                        self.drain_reader(harness, IMEI + "\r\n")
                        self.assert_identity(harness.ns)
                        self.assertEqual(popups, [phone])

    def test_production_reader_accepts_existing_complete_imei_formats(self):
        formats = [("0" * (length - 1) + "1", "0" * (length - 1) + "1")
                   for length in range(14, 18)]
        formats += [("[I]-[ril.proatc] " + IMEI, IMEI), ("IMEI: " + IMEI, IMEI),
                    ("[W]-[ril.proatc] iMeI: " + IMEI, IMEI)]
        for line, expected in formats:
            with self.subTest(line=line), configure(InitializationHarness(hold="AT+CGSN")) as harness:
                self.assert_startup_window(harness)
                self.drain_reader(harness, line + "\r\n")
                self.assert_identity(harness.ns, expected)
                self.assertEqual(harness.modem.writes, STARTUP_BYTES)

    def test_runtime_handler_preserves_cnum_clip_processing_and_query_owner(self):
        for phone in PHONE_NUMBERS:
            for command in ("CNUM", "CLIP"):
                with self.subTest(phone=phone, command=command):
                    namespace, _, state, calls, handle = self.make_runtime()
                    owner = namespace["_cloud_imei_query_owner"]
                    line = ('+CNUM: "","' + phone + '",145' if command == "CNUM"
                            else '[I]-[ril.proatc] +CLIP: "' + phone + '",145')
                    handle(line)
                    if command == "CNUM":
                        self.assertEqual(namespace["LOCAL_NUMBER"], phone)
                    else:
                        self.assertEqual(state.call_state.last_clip_num, phone)
                        self.assertIn(("call_popup", (phone,)), calls)
                    self.assert_window_open(namespace, owner)
                    handle(IMEI)
                    self.assert_identity(namespace)

    def test_runtime_handler_accepts_standard_log_bodies_without_source_whitelist(self):
        for prefix in ("", "[I]-[ril.proatc] ", "[W]-[ril.response] ", "[E]-[another.source] "):
            for label in ("", "IMEI: ", "imei: "):
                with self.subTest(prefix=prefix, label=label):
                    namespace, _, _, _, handle = self.make_runtime()
                    handle(prefix + label + IMEI)
                    self.assert_identity(namespace)

    def test_runtime_handler_ignores_other_reports_and_numeric_business_logs(self):
        lines = (
            '+CNUM: "","+882132456789012",145', '+CLIP: "+882132456789012",145',
            '+CLCC: 1,1,4,0,0,"+882132456789012",145', '+CMTI: "SM",882132456789012',
            "+CME ERROR: 882132456789012", "[I]-[ril.proatc] +CNUM: \"\",\"+882132456789012\",145",
            "[I]-[business] phone=" + IMEI, "[W]-[probe] received " + IMEI,
            "[E]-[probe] " + IMEI + " pending", "status IMEI: " + IMEI,
            "IMEI: " + IMEI + " extra", "+" + IMEI, '"' + IMEI + '"',
        )
        for line in lines:
            with self.subTest(line=line):
                namespace, _, _, _, handle = self.make_runtime()
                owner = namespace["_cloud_imei_query_owner"]
                handle(line)
                self.assert_window_open(namespace, owner)
                handle(IMEI)
                self.assert_identity(namespace)

    def test_runtime_handler_ignores_blank_echo_terminal_and_wrong_lengths(self):
        for line in ("", "OK", "ERROR", "AT+CGSN", "[I]-[ril.proatc] OK",
                     "0" * 13, "0" * 18):
            with self.subTest(line=line):
                namespace, _, _, _, handle = self.make_runtime()
                owner = namespace["_cloud_imei_query_owner"]
                handle(line)
                self.assert_window_open(namespace, owner)

    def test_runtime_handler_keeps_sms_digit_body_owned_until_its_boundary(self):
        namespace, _, state, calls, handle = self.make_runtime()
        owner = namespace["_cloud_imei_query_owner"]
        head = "10086 26/11/14,10:00:00+32 SMS-BEGIN"
        handle("[I]-[handler_sms.smsCallback] " + head)
        handle(IMEI)
        self.assertTrue(state.sms_collector.active)
        self.assert_window_open(namespace, owner)
        handle("[I]-[probe] SMS-END")
        self.assertFalse(state.sms_collector.active)
        self.assertIn(("cloud_sms", (head, "SMS-BEGIN\n" + IMEI)), calls)
        self.assert_window_open(namespace, owner)
        handle(IMEI)
        self.assert_identity(namespace)

    def test_runtime_handler_keeps_inactive_and_six_second_expiry_boundaries(self):
        for deadline, now, captures in ((0.0, 100.0, False), (106.0, 106.0, True),
                                       (106.0, 106.001, False)):
            with self.subTest(deadline=deadline, now=now):
                namespace, clock, _, _, handle = self.make_runtime()
                namespace["cloud_imei_query_deadline"] = deadline
                if not deadline:
                    namespace.pop("_cloud_imei_query_owner")
                clock.value = now
                handle(IMEI)
                if captures:
                    self.assert_identity(namespace)
                else:
                    self.assertEqual(namespace["CLOUD_DEVICE_IMEI"], "")
                    self.assertFalse(namespace["cloud_imei_verified"])
                    self.assertEqual(namespace["cloud_imei_query_deadline"], 0.0)
                    self.assertNotIn("_cloud_imei_query_owner", namespace)
                    handle('+CNUM: "","+882132456789012",145')
                    self.assertEqual(namespace["LOCAL_NUMBER"], PHONE_NUMBERS[1])
                    self.assertFalse(namespace["cloud_imei_verified"])

    def wait_query(self, namespace):
        wait_for(lambda: not namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(),
                 "raw IMEI worker did not stop")

    def test_failed_old_query_preserves_new_owner_and_later_classified_response(self):
        namespace, _, _, _, handle = self.make_runtime()
        newer_owner = object()

        def replace_before_failure(serial_obj, command, **kwargs):
            namespace["_cloud_imei_query_owner"] = newer_owner
            namespace["cloud_imei_query_deadline"] = 106.0
            return serial_sender.SerialCommandResult(False, "synthetic old query failure")

        namespace["write_serial_command_result"] = replace_before_failure
        self.assertTrue(request_cloud_device_imei_namespace_runtime(namespace)[0])
        self.wait_query(namespace)
        self.assert_window_open(namespace, newer_owner)
        handle('+CNUM: "","+882132456789012",145')
        self.assert_window_open(namespace, newer_owner)
        handle("IMEI: " + IMEI)
        self.assert_identity(namespace)

    def test_old_query_connection_and_stop_guards_do_not_create_or_redirect_windows(self):
        for change in ("serial_obj", "serial_connection_generation", "serial_stop_event",
                       "TK_SHUTDOWN", "serial_running", "is_exiting"):
            with self.subTest(change=change):
                namespace, _, _, _, _ = self.make_runtime()
                original = namespace["serial_obj"]
                replacement = query_fixtures.QueryModem()
                namespace["cloud_imei_query_deadline"] = 0.0
                namespace.pop("_cloud_imei_query_owner")
                try:
                    with serial_sender.DEFAULT_SERIAL_WRITE_LOCK:
                        self.assertTrue(request_cloud_device_imei_namespace_runtime(namespace)[0])
                        wait_for(lambda: bool(namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot()),
                                 "raw IMEI worker did not start")
                        if change in ("serial_stop_event", "TK_SHUTDOWN"):
                            namespace[change].set()
                        else:
                            namespace[change] = {"serial_obj": replacement,
                                "serial_connection_generation": 2, "serial_running": False,
                                "is_exiting": True}[change]
                finally:
                    self.wait_query(namespace)
                self.assertEqual(original.writes, [])
                self.assertEqual(replacement.writes, [])
                self.assertEqual(namespace["cloud_imei_query_deadline"], 0.0)
                self.assertNotIn("_cloud_imei_query_owner", namespace)


if __name__ == "__main__":
    unittest.main()
