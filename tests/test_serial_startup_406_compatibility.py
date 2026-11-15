"""Keep 4.0.6 startup writes independent of replies and retain real transactions."""
import threading
import time
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from sms_core import serial_sender
from sms_core.cloud_state_namespace_runtime import request_cloud_device_imei_namespace_runtime
from sms_core.serial_io_runtime import read_serial_line_safely_runtime
from sms_ui.app_infrastructure_namespace_runtime import safe_close_serial_namespace_runtime
from tests.test_serial_initialization import InitializationHarness


STARTUP_COMMANDS = ("AT+CLIP=1", "AT+CGSN", "AT+CNUM")
STARTUP_BYTES = [b"AT+CLIP=1\r\n", b"AT+CGSN\r\n", b"AT+CNUM\r\n"]


def wait_for(predicate, message):
    deadline = time.monotonic() + 3
    while time.monotonic() < deadline:
        if predicate():
            return
        threading.Event().wait(0.002)
    raise AssertionError(message)


def configure(harness):
    def queued_bytes(_modem):
        with harness.reads.mutex:
            return sum(len(value) for value in harness.reads.queue if isinstance(value, bytes))

    type(harness.modem).in_waiting = property(queued_bytes)
    harness.read_requests = []

    def read(size):
        harness.read_requests.append(size)
        with harness.reads.not_empty:
            if not harness.reads.queue:
                harness.reads.not_empty.wait(0.02)
            if not harness.reads.queue:
                return b""
            raw = harness.reads.queue.popleft()
            if len(raw) > size:
                harness.reads.queue.appendleft(raw[size:])
            return raw[:size]

    harness.modem.read = read
    return harness


class StartupHarness(InitializationHarness):
    def __init__(self, *, suppress="all", startup_response="OK", **kwargs):
        self.suppressed_command = suppress
        self.startup_response = startup_response
        super().__init__(hold="unused", **kwargs)

    def terminal(self, command, *, response="OK", extra=""):
        if command in STARTUP_COMMANDS:
            if self.suppressed_command in ("all", command):
                return
            response = self.startup_response
        return super().terminal(command, response=response, extra=extra)


class SerialStartup406CompatibilityTests(unittest.TestCase):
    def assert_started(self, harness):
        self.assertTrue(harness.wrote["AT+CNUM"].wait(0.3), "startup waited for an earlier reply")
        self.assertEqual(harness.modem.writes, STARTUP_BYTES)
        self.assertIsNone(harness.ns.get("_serial_initialization_session"))
        self.assertIsNone(getattr(harness.modem, "_serial_initialization_session", None))
        self.assertIsNone(harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active)
        self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))

    def confirmed_ath(self, harness, *, timeout=1.0):
        results = []

        def submit():
            results.append(serial_sender.write_serial_command_confirmed_locked(
                harness.ns["serial_lock"], lambda: harness.ns["serial_obj"], "ATH",
                response_coordinator=harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"],
                response_timeout=timeout,
            ))

        worker = threading.Thread(target=submit)
        worker.start()
        self.addCleanup(worker.join, 2)
        wait_for(lambda: b"ATH\r\n" in harness.modem.writes, "independent ATH was not submitted")
        return worker, results

    def test_open_submits_three_commands_before_connected_callback_without_a_waiter(self):
        with configure(StartupHarness()) as harness:
            self.assert_started(harness)
            self.assertTrue(harness.ready.is_set())
            self.assertEqual(harness.ns["cloud_imei_query_deadline"], harness.clock + 6)
            self.assertEqual(harness.errors, [])

    def test_missing_each_or_all_startup_replies_do_not_pause_raw_commands_or_ota(self):
        for command in (*STARTUP_COMMANDS, "all"):
            with self.subTest(missing=command), configure(StartupHarness(suppress=command)) as harness:
                self.assert_started(harness)
                harness.clock += 60
                harness.emit("AFTER-OLD-STARTUP-DEADLINE\r\n")
                wait_for(lambda: "AFTER-OLD-STARTUP-DEADLINE" in harness.lines, "reader stopped without ACK")
                self.assertTrue(serial_sender.write_serial_command_result(harness.modem, "AT").ok)
                owner = object()
                try:
                    self.assertTrue(serial_sender.reserve_serial_for_ota(harness.modem, owner))
                finally:
                    serial_sender.release_serial_from_ota(harness.modem, owner)
                self.assertEqual(harness.modem.writes, STARTUP_BYTES + [b"AT\r\n"])
                self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))
                self.assertEqual(harness.errors, [])

    def test_late_startup_logs_are_read_without_retransmission_or_a_new_gate(self):
        with configure(StartupHarness()) as harness:
            self.assert_started(harness)
            harness.clock += 300
            harness.emit("[I]-[ril.proatc] OK\r\n[I]-[ril.proatc] ERROR\r\nLATE-STARTUP-END\r\n")
            wait_for(lambda: "LATE-STARTUP-END" in harness.lines, "late startup logs were lost")
            self.assertEqual(harness.modem.writes, STARTUP_BYTES)
            self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))
            self.assertIsNone(harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active)
            self.assertEqual(harness.errors, [])

    def test_explicit_startup_error_replies_do_not_block_following_writes(self):
        for response in ("ERROR", "+CME ERROR: 10", "+CMS ERROR: 500"):
            with self.subTest(response=response), configure(StartupHarness(
                    suppress=None, startup_response=response)) as harness:
                self.assert_started(harness)
                wait_for(lambda: "[I]-[ril.proatc] " + response in harness.lines, "startup error was not logged")
                self.assertTrue(serial_sender.write_serial_command_result(harness.modem, "AT").ok)
                self.assertEqual(harness.modem.writes, STARTUP_BYTES + [b"AT\r\n"])
                self.assertEqual(harness.errors, [])

    def test_short_write_write_error_and_flush_error_keep_406_best_effort_startup(self):
        private = "synthetic-only-private-startup-error"
        for stage in ("partial", "write-error", "flush-error"):
            with self.subTest(stage=stage), configure(StartupHarness(stage=stage, error_message=private)) as harness:
                self.assert_started(harness)
                harness.emit("READ-AFTER-STARTUP-WRITE-FAILURE\r\n")
                wait_for(lambda: "READ-AFTER-STARTUP-WRITE-FAILURE" in harness.lines, "reader stopped after startup failure")
                self.assertNotIn(private, "\n".join(harness.messages + harness.lines))
                self.assertEqual(harness.errors, [])

    def test_fragmented_terminal_and_multichunk_logs_keep_all_bytes(self):
        with configure(StartupHarness()) as harness:
            self.assert_started(harness)
            harness.emit("[I]-[ril.pro")
            harness.emit("atc] OK\r")
            harness.emit("\n" + "X" * 4090 + "\r\nMULTICHUNK-END\r\n")
            wait_for(lambda: "MULTICHUNK-END" in harness.lines, "fragmented input did not finish")
            self.assertIn("[I]-[ril.proatc] OK", harness.lines)
            self.assertIn("X" * 4090, harness.lines)
            self.assertTrue(all(0 < size <= 4096 for size in harness.read_requests))
            self.assertEqual(harness.modem.writes, STARTUP_BYTES)
            self.assertEqual(harness.errors, [])

    def test_busy_log_callbacks_and_later_input_do_not_create_startup_deadlines(self):
        with configure(StartupHarness()) as harness:
            self.assert_started(harness)
            entered, release = threading.Event(), threading.Event()
            self.addCleanup(release.set)
            observed = []

            def pause(line):
                observed.append(line)
                if line == "BUSY-LOG":
                    entered.set()
                    if not release.wait(2):
                        raise AssertionError("log gate was not released")

            harness.on_line = pause
            harness.emit("BUSY-LOG\r\n")
            self.assertTrue(entered.wait(1))
            harness.clock += 300
            harness.emit("LATE-LOG\r\n" * 500 + "[I]-[ril.proatc] OK\r\nBUSY-END\r\n")
            release.set()
            wait_for(lambda: "BUSY-END" in harness.lines, "busy logs or late input were lost")
            self.assertEqual(observed.count("LATE-LOG"), 500)
            self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))
            self.assertEqual(harness.modem.writes, STARTUP_BYTES)
            self.assertEqual(harness.errors, [])

    def test_oversize_fake_os_return_preserves_sms_and_logs(self):
        harness = configure(StartupHarness())
        sms = []
        harness.ns["_cloud_send_sms_event"] = lambda head, body: sms.append(body)
        with harness:
            self.assert_started(harness)
            entered, release = threading.Event(), threading.Event()
            self.addCleanup(release.set)

            def pause(line):
                if line == "OVERSIZE-READ-GATE":
                    entered.set()
                    if not release.wait(2):
                        raise AssertionError("oversize read gate was not released")

            harness.on_line = pause
            harness.emit("OVERSIZE-READ-GATE\r\n")
            self.assertTrue(entered.wait(1))
            original_read = harness.modem.read

            def read(size):
                raw = original_read(size)
                if raw == b"PREFIX\r\n":
                    return raw + (
                        b"[I]-[handler_sms.smsCallback] 10000 26/11/14,10:00:00+32 SMS-EXTRA\r\n"
                        b"[I]-[probe] AFTER-SMS\r\nOVERSIZE-END\r\n"
                    )
                return raw

            harness.modem.read = read
            harness.emit("PREFIX\r\n")
            release.set()
            wait_for(lambda: "OVERSIZE-END" in harness.lines, "extra OS-return bytes were discarded")
            self.assertEqual(sms, ["SMS-EXTRA"])
            self.assertIn("PREFIX", harness.lines)
            self.assertEqual(harness.modem.writes, STARTUP_BYTES)
            self.assertEqual(harness.errors, [])

    def test_default_read_lock_fallback_remains_usable_without_startup_owner(self):
        with configure(StartupHarness()) as harness:
            self.assert_started(harness)
            harness.ns["serial_read_lock"] = None
            harness.emit("FALLBACK-READ-END\r\n")
            wait_for(lambda: "FALLBACK-READ-END" in harness.lines, "fallback read lease deadlocked")
            self.assertTrue(serial_sender.write_serial_command_result(harness.modem, "AT").ok)
            self.assertEqual(harness.errors, [])

    def test_real_sms_call_config_and_ota_unknown_markers_are_not_cleared_by_logs(self):
        for operation in ("SMS/PDU", "call", "device_config", "OTA"):
            with self.subTest(operation=operation), configure(StartupHarness()) as harness:
                self.assert_started(harness)
                marker = operation + " synthetic unknown result"
                harness.modem._sms_recovery_error = marker
                harness.clock += 300
                harness.emit("OK\r\n[I]-[ril.proatc] OK\r\nUNKNOWN-MARKER-END\r\n")
                wait_for(lambda: "UNKNOWN-MARKER-END" in harness.lines, "isolated transaction stopped logs")
                self.assertEqual(harness.modem._sms_recovery_error, marker)
                self.assertFalse(serial_sender.write_serial_command_result(harness.modem, "AT").ok)
                self.assertFalse(serial_sender.reserve_serial_for_ota(harness.modem, object()))
                self.assertEqual(harness.modem.writes, STARTUP_BYTES)
                self.assertEqual(harness.errors, [])

    def test_independent_confirmed_ath_still_requires_its_own_terminal(self):
        for terminal in ("OK", "ERROR"):
            with self.subTest(terminal=terminal), configure(StartupHarness()) as harness:
                self.assert_started(harness)
                worker, results = self.confirmed_ath(harness)
                harness.emit("ATH-STILL-WAITING\r\n")
                wait_for(lambda: "ATH-STILL-WAITING" in harness.lines, "reader did not process ATH gate")
                self.assertEqual(results, [])
                harness.emit("[I]-[ril.proatc] " + terminal + "\r\n")
                worker.join(1)
                self.assertFalse(worker.is_alive())
                self.assertEqual(results[0].ok, terminal == "OK")
                self.assertEqual(harness.modem.writes, STARTUP_BYTES + [b"ATH\r\n"])
                self.assertEqual(harness.errors, [])

    def test_old_connection_terminal_cannot_complete_current_confirmed_command(self):
        with configure(StartupHarness()) as harness:
            self.assert_started(harness)
            worker, results = self.confirmed_ath(harness)
            coordinator = harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]
            waiter = coordinator._active
            coordinator.observe_line("[I]-[ril.proatc] OK", connection=SimpleNamespace(is_open=False))
            self.assertFalse(waiter.done())
            self.assertEqual(results, [])
            harness.emit("OK\r\n")
            worker.join(1)
            self.assertFalse(worker.is_alive())
            self.assertTrue(results[0].ok)
            self.assertEqual(harness.errors, [])

    def test_cold_identity_missing_startup_cgsn_keeps_existing_independent_reread(self):
        with configure(StartupHarness(suppress="AT+CGSN")) as harness:
            self.assert_started(harness)
            harness.clock += 30
            harness.emit("000000000000001\r\nIDENTITY-EXPIRED-END\r\n")
            wait_for(lambda: "IDENTITY-EXPIRED-END" in harness.lines, "expired identity log was lost")
            self.assertEqual(harness.ns["_cloud_runtime_imei"](), "")
            harness.suppressed_command = None
            harness.ns["write_serial_command_result"] = serial_sender.write_serial_command_result
            request_cloud_device_imei_namespace_runtime(harness.ns)
            wait_for(lambda: harness.ns["cloud_imei_verified"]
                     and not harness.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(),
                     "independent IMEI query did not confirm the device")
            self.assertEqual(harness.ns["_cloud_runtime_imei"](), "000000000000001")
            self.assertEqual(harness.modem.writes, STARTUP_BYTES + [b"AT+CGSN\r\n"])
            self.assertEqual(harness.errors, [])

    def test_gui_close_stops_reader_and_keeps_already_submitted_startup_bytes(self):
        with configure(StartupHarness()) as harness:
            self.assert_started(harness)
            safe_close_serial_namespace_runtime(harness.ns)
            harness.stop.set()
            harness.thread.join(1)
            self.assertFalse(harness.thread.is_alive())
            self.assertFalse(harness.modem.is_open)
            self.assertIsNone(harness.ns["serial_obj"])
            self.assertEqual(harness.modem.writes, STARTUP_BYTES)
            self.assertIsNone(harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active)
            self.assertEqual(harness.errors, [])

    def test_real_input_query_failure_is_reported_as_a_read_failure(self):
        with configure(StartupHarness()) as harness:
            self.assert_started(harness)

            class FailedInput:
                is_open = True

                @property
                def in_waiting(self):
                    raise OSError("synthetic input query failure")

            with self.assertRaisesRegex(OSError, "synthetic input query failure"):
                read_serial_line_safely_runtime(threading.RLock(), lambda: FailedInput(), OSError)
            self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))
            self.assertEqual(harness.errors, [])

    def test_connected_callback_failure_does_not_leak_transaction_or_repeat_startup(self):
        with configure(StartupHarness(connected_error=True)) as harness:
            harness.thread.join(1)
            self.assertFalse(harness.thread.is_alive())
            self.assertEqual(harness.modem.writes, STARTUP_BYTES)
            self.assertIsNone(harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active)
            acquired = []

            def probe():
                available = serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK.acquire(timeout=0.2)
                acquired.append(available)
                if available:
                    serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK.release()

            worker = threading.Thread(target=probe)
            worker.start()
            worker.join(1)
            self.assertFalse(worker.is_alive())
            self.assertEqual(acquired, [True])

    def test_ota_reservation_renewal_cannot_extend_a_real_unknown_transaction(self):
        with configure(StartupHarness()) as harness:
            self.assert_started(harness)
            owner = object()
            expiry = time.monotonic() + 60
            reservation = (harness.modem, owner, expiry)
            with patch.dict(serial_sender._OTA_RESERVATIONS, {id(harness.modem): reservation}):
                harness.modem._sms_recovery_error = "synthetic unknown OTA result"
                self.assertFalse(serial_sender.reserve_serial_for_ota(harness.modem, owner))
                self.assertEqual(serial_sender._OTA_RESERVATIONS[id(harness.modem)], reservation)
            self.assertEqual(harness.modem.writes, STARTUP_BYTES)
            self.assertEqual(harness.errors, [])

    def test_blacklisted_call_hangup_does_not_wait_for_missing_startup_replies(self):
        with configure(StartupHarness()) as harness:
            self.assert_started(harness)
            harness.ns.update(CALL_FILTER_MODE="Blacklist", CALL_BLACKLIST=["10086"])
            harness.emit('RING\r\n+CLIP: "10086",129\r\n')
            wait_for(lambda: b"ATH\r\n" in harness.modem.writes, "blacklist hangup waited for startup ACK")
            self.assertEqual(harness.modem.writes.count(b"ATH\r\n"), 1)
            harness.emit("[I]-[ril.proatc] OK\r\n")
            wait_for(lambda: not harness.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(),
                     "real auto-hangup terminal did not release its worker")
            self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))
            self.assertEqual(harness.errors, [])

    def test_missing_startup_replies_keep_sms_logs_and_recording_lines(self):
        harness = configure(StartupHarness())
        sms = []
        harness.ns["_cloud_send_sms_event"] = lambda head, body: sms.append(body)
        with harness:
            self.assert_started(harness)
            harness.clock += 300
            harness.emit(
                "[I]-[handler_sms.smsCallback] 10000 26/11/14,10:00:00+32 STARTUP-SMS\r\n"
                "[I]-[probe] AFTER-SMS\r\nRING\r\n"
                "@@CALL_RECORD_META|virtual|incoming\r\nSTARTUP-LOG-END\r\n"
            )
            wait_for(lambda: "STARTUP-LOG-END" in harness.lines, "startup without ACK stopped the reader")
            self.assertEqual(sms, ["STARTUP-SMS"])
            self.assertIn("RING", harness.recording_lines)
            self.assertIn("@@CALL_RECORD_META|virtual|incoming", harness.recording_lines)
            self.assertEqual(harness.modem.writes, STARTUP_BYTES)
            self.assertEqual(harness.errors, [])


if __name__ == "__main__":
    unittest.main()
