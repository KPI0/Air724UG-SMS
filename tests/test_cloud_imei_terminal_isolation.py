"""System IMEI reads own their terminal before UI call/SMS transactions."""
import queue
import threading
import time
import unittest

from sms_app.cloud_namespace_bindings import install_cloud_namespace_bindings
from sms_core import serial_sender
from sms_core.cloud_serial_log_runtime import CloudSerialLogDrainState
from sms_core.sms_send_namespace_runtime import send_manual_sms_namespace_runtime
from sms_ui.app_infrastructure_namespace_runtime import safe_close_serial_namespace_runtime
from tests.test_serial_initialization import InitializationHarness


class MissingImeiHarness(InitializationHarness):
    def terminal(self, command, *, response="OK", extra=""):
        if command == "AT+CGSN":
            self.emit("[I]-[ril.proatc] " + response + "\r\n" + extra)
        elif command != "AT+CMGF=0":
            super().terminal(command, response=response, extra=extra)


class CloudImeiTerminalIsolationTests(unittest.TestCase):
    def wait_for(self, predicate, *, timeout=3):
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if predicate():
                return
            threading.Event().wait(0.005)
        self.fail("Synthetic reader/query condition was not reached")

    def bind_query(self, harness):
        self.wait_for(lambda: len(harness.modem.writes) == 3
            and not harness.ns["_serial_initialization_session"].pending)
        self.assertFalse(harness.ns["cloud_imei_verified"])
        harness.hold = "AT+CGSN"
        self.install_query(harness)

    def install_query(self, harness):
        harness.ns["write_serial_command_result"] = serial_sender.write_serial_command_result
        harness.ns["CLOUD_SERIAL_LOG_Q"] = queue.Queue()
        harness.ns["CLOUD_SERIAL_LOG_DRAIN_STATE"] = CloudSerialLogDrainState()
        install_cloud_namespace_bindings(harness.ns)
        harness.ns["_cloud_log"] = lambda *args, **kwargs: None
        harness.ns["_cloud_now_ts"] = lambda: 1234567890

    def test_explicit_query_queued_during_startup_waits_for_its_own_terminal(self):
        with InitializationHarness(hold="AT+CNUM") as harness:
            self.assertTrue(harness.wrote["AT+CNUM"].wait(1))
            self.assertTrue(harness.ns["cloud_imei_verified"])
            self.install_query(harness)
            harness.hold = "AT+CGSN"
            self.assertTrue(harness.ns["request_cloud_device_imei"]()[0])
            self.assertEqual(len(harness.modem.writes), 3)
            harness.terminal("AT+CNUM")
            self.wait_for(lambda: len(harness.modem.writes) == 4)
            waiter = harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active
            self.assertEqual(waiter.command, "AT+CGSN")
            self.assertFalse(waiter.done())
            harness.terminal("AT+CGSN")
            self.assert_workers_stopped(harness)
            self.assertEqual(harness.ns["CLOUD_DEVICE_IMEI"], "000000000000001")
            self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))

    def start_query(self, harness):
        self.assertTrue(harness.ns["request_cloud_device_imei"]()[0])
        self.wait_for(lambda: len(harness.modem.writes) == 4)
        waiter = harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active
        self.assertIsNotNone(waiter)
        self.assertEqual(waiter.command, "AT+CGSN")
        return waiter

    def assert_workers_stopped(self, harness):
        self.wait_for(lambda: not harness.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(), timeout=8)
        self.assertIsNone(harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active)

    def test_query_terminal_cannot_confirm_a_queued_ui_dial(self):
        for terminal in ("OK", "ERROR"):
            with self.subTest(terminal=terminal), MissingImeiHarness(hold="unused") as harness:
                self.bind_query(harness)
                query_waiter = self.start_query(harness)
                harness.fixture.actions["dial"]("10086")
                self.assertFalse(harness.wrote["ATD10086;"].wait(0.04))
                self.assertFalse(query_waiter.done())
                harness.emit("[I]-[ril.proatc] " + terminal + "\r\n")
                self.assertTrue(harness.wrote["ATD10086;"].wait(1))
                dial_waiter = harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active
                self.assertIsNotNone(dial_waiter)
                self.assertEqual(dial_waiter.command, "ATD10086;")
                self.assertFalse(dial_waiter.done())
                self.assertEqual(harness.ns["cloud_imei_query_deadline"], 0)
                self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))
                harness.emit("[I]-[ril.proatc] OK\r\n")
                self.assert_workers_stopped(harness)

    def test_query_terminal_cannot_confirm_the_next_manual_sms_cmgf(self):
        with MissingImeiHarness(hold="unused") as harness:
            self.bind_query(harness)
            self.start_query(harness)
            worker = send_manual_sms_namespace_runtime(harness.ns, "10086", "synthetic-only")
            try:
                self.assertFalse(worker is None)
                self.assertEqual(len(harness.modem.writes), 4)
                harness.emit("[I]-[ril.proatc] OK\r\n")
                self.wait_for(lambda: b"AT+CMGF=0\r\n" in harness.modem.writes)
                self.assertFalse(any(p.startswith(b"AT+CMGS=") for p in harness.modem.writes))
                waiter = harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active
                self.assertEqual(waiter.command, "AT+CMGF=0")
                self.assertFalse(waiter.done())
                harness.emit("[I]-[ril.proatc] OK\r\n")
                self.wait_for(lambda: any(p.startswith(b"AT+CMGS=") for p in harness.modem.writes))
                harness.emit("[I]-[ril.proatc] >\r\n")
                self.wait_for(lambda: any(p.endswith(b"\x1a") for p in harness.modem.writes))
                harness.emit("[I]-[ril.proatc] +CMGS: 1\r\n[I]-[ril.proatc] OK\r\n")
                worker.join(2)
                self.assertFalse(worker.is_alive())
                self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))
                self.assert_workers_stopped(harness)
            finally:
                harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"].cancel_active("synthetic test cleanup")
                serial_sender.DEFAULT_SMS_PDU_SEND_COORDINATOR.cancel_active("synthetic test cleanup")
                if worker is not None:
                    worker.join(2)

    def test_identity_capture_and_window_end_follow_the_real_reader(self):
        with MissingImeiHarness(hold="unused") as harness:
            self.bind_query(harness)
            self.start_query(harness)
            self.assertEqual(harness.ns["cloud_imei_query_deadline"], harness.clock + 6)
            harness.emit("[I]-[ril.proatc] 000000000000001\r\n[I]-[ril.proatc] OK\r\n")
            self.assert_workers_stopped(harness)
            self.assertTrue(harness.ns["cloud_imei_verified"])
            self.assertEqual(harness.ns["CLOUD_DEVICE_IMEI"], "000000000000001")
            self.assertEqual(harness.ns["cloud_imei_query_deadline"], 0)
            self.assertNotIn("_cloud_imei_query_owner", harness.ns)

    def test_query_wait_releases_write_lock_while_retaining_transaction(self):
        with MissingImeiHarness(hold="unused") as harness:
            self.bind_query(harness)
            self.start_query(harness)
            self.assertTrue(serial_sender.DEFAULT_SERIAL_WRITE_LOCK.acquire(timeout=0.2))
            serial_sender.DEFAULT_SERIAL_WRITE_LOCK.release()
            self.assertFalse(serial_sender.reserve_serial_for_ota(harness.modem, object()))
            harness.emit("READING-REMAINS-ACTIVE\r\n[I]-[ril.proatc] OK\r\n")
            self.assert_workers_stopped(harness)
            self.assertIn("READING-REMAINS-ACTIVE", harness.lines)

    def test_trusted_reader_terminal_during_flush_failure_remains_confirmed(self):
        for terminal in ("OK", "ERROR"):
            with self.subTest(terminal=terminal), MissingImeiHarness(hold="unused") as harness:
                self.bind_query(harness)
                reading = threading.Event()

                def read(_size):
                    reading.set()
                    try:
                        return harness.reads.get(timeout=1)
                    except queue.Empty:
                        return b""

                harness.modem.read = read
                self.assertTrue(reading.wait(1))

                def flush():
                    waiter = harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active
                    harness.emit("[I]-[ril.proatc] " + terminal + "\r\n")
                    if not waiter.wait(1).line:
                        raise OSError("Synthetic reader failed to return the terminal during flush")
                    raise OSError("Synthetic USB flush failure after terminal")

                harness.modem.flush = flush
                self.assertTrue(harness.ns["request_cloud_device_imei"]()[0])
                self.assert_workers_stopped(harness)
                self.assertEqual(len(harness.modem.writes), 4)
                self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))
                self.assertEqual(harness.ns["cloud_imei_query_deadline"], 0)
                self.assertNotIn("_cloud_imei_query_owner", harness.ns)

    def test_missing_terminal_isolates_original_serial_before_queued_dial(self):
        with MissingImeiHarness(hold="unused") as harness:
            self.bind_query(harness)
            self.start_query(harness)
            harness.fixture.actions["dial"]("10086")
            self.assert_workers_stopped(harness)
            self.assertEqual(len(harness.modem.writes), 4)
            self.assertIn("IMEI查询", serial_sender.serial_sms_recovery_error(harness.modem))
            self.assertTrue(harness.modem.is_open)
            harness.emit("READING-AFTER-TIMEOUT\r\n[I]-[ril.proatc] OK\r\n")
            self.wait_for(lambda: "READING-AFTER-TIMEOUT" in harness.lines)
            self.assertFalse(serial_sender.write_serial_command_result(harness.modem, "ATI").ok)
            self.assertEqual(harness.ns["cloud_imei_query_deadline"], 0)

    def test_stop_during_final_write_lock_wait_never_starts_query_window(self):
        with MissingImeiHarness(hold="unused") as harness:
            self.bind_query(harness)
            original_deadline = harness.ns["cloud_imei_query_deadline"]
            serial_sender.DEFAULT_SERIAL_WRITE_LOCK.acquire()
            try:
                self.assertTrue(harness.ns["request_cloud_device_imei"]()[0])
                self.wait_for(lambda: harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active is not None)
                self.assertEqual(harness.ns["cloud_imei_query_deadline"], original_deadline)
                harness.stop.set()
            finally:
                serial_sender.DEFAULT_SERIAL_WRITE_LOCK.release()
            self.assert_workers_stopped(harness)
            self.assertEqual(len(harness.modem.writes), 3)
            self.assertEqual(harness.ns["cloud_imei_query_deadline"], original_deadline)
            self.assertNotIn("_cloud_imei_query_owner", harness.ns)
            self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))

    def test_real_close_cancels_query_and_isolates_only_its_original_serial(self):
        with MissingImeiHarness(hold="unused") as harness:
            self.bind_query(harness)
            self.start_query(harness)
            harness.stop.set()
            safe_close_serial_namespace_runtime(harness.ns)
            self.assert_workers_stopped(harness)
            self.assertEqual(len(harness.modem.writes), 4)
            self.assertIn("IMEI查询", serial_sender.serial_sms_recovery_error(harness.modem))
            self.assertIsNone(harness.ns["serial_obj"])
            self.assertFalse(harness.modem.is_open)


if __name__ == "__main__":
    unittest.main()
