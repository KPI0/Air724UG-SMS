"""Automatic IMEI queries use raw writes without changing business confirmation."""
import re
import threading
import time
import unittest
from types import SimpleNamespace

from sms_core import serial_sender
from sms_core.cloud_imei_runtime import request_cloud_device_imei_worker
from sms_core.cloud_state_namespace_runtime import (
    cloud_channel_status_snapshot_namespace_runtime,
    maybe_capture_cloud_device_imei_namespace_runtime,
    request_cloud_device_imei_namespace_runtime,
)
from sms_core.threading_runtime import WorkerThreadRegistry
from sms_core.serial_io_runtime import read_serial_line_safely_runtime
from sms_core.serial_namespace_runtime import open_and_initialize_serial_namespace_runtime
from sms_ui.app_infrastructure_namespace_runtime import safe_close_serial_namespace_runtime


class QueryModem:
    def __init__(self):
        self.is_open = True
        self.writes = []
        self.wrote = threading.Event()
        self.write_error = None
        self.flush_error = None
        self.short_write = False

    def write(self, payload):
        if self.write_error:
            raise self.write_error
        self.writes.append(payload)
        self.wrote.set()
        return 1 if self.short_write else len(payload)

    def flush(self):
        if self.flush_error:
            raise self.flush_error

    def close(self):
        self.is_open = False


class CloudImeiTerminalIsolationTests(unittest.TestCase):
    def setUp(self):
        self.modem = QueryModem()
        self.clock = 100.0
        self.logs, self.workers, self.ota_owners = [], [], []
        self.coordinator = serial_sender.AtCommandResponseCoordinator()
        self.namespace = {
            "serial_lock": threading.Lock(), "serial_obj": self.modem,
            "serial_connection_generation": 1, "serial_running": True,
            "serial_stop_event": threading.Event(), "TK_SHUTDOWN": threading.Event(),
            "cloud_imei_query_deadline": 0.0, "cloud_imei_verified": False,
            "CLOUD_DEVICE_IMEI": "", "IMEI_REGEX": re.compile(r"\b(\d{14,17})\b"),
            "SERIAL_COMMAND_THREAD_REGISTRY": WorkerThreadRegistry(),
            "SERIAL_COMMAND_RESPONSE_COORDINATOR": self.coordinator,
            "write_serial_command_result": serial_sender.write_serial_command_result,
            "_cloud_log": self.logs.append, "_push_serial_debug": lambda line: None,
            "time": SimpleNamespace(monotonic=lambda: self.clock),
        }
        self.namespace["_set_cloud_device_imei"] = self.capture_identity

    def tearDown(self):
        self.namespace["serial_stop_event"].set()
        self.coordinator.cancel_active("synthetic test cleanup")
        for worker in self.namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot() + tuple(self.workers):
            worker.join(1)
            self.assertFalse(worker.is_alive())
        for modem, owner in self.ota_owners:
            serial_sender.release_serial_from_ota(modem, owner)

    def capture_identity(self, imei, *, source):
        self.assertEqual(source, "AT+CGSN")
        self.namespace["CLOUD_DEVICE_IMEI"] = imei
        self.namespace["cloud_imei_verified"] = True
        return True

    def wait_for(self, predicate, *, timeout=1):
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            if predicate():
                return
            threading.Event().wait(0.005)
        self.fail("Synthetic query condition was not reached")

    def start_query(self):
        self.assertTrue(request_cloud_device_imei_namespace_runtime(self.namespace)[0])

    def assert_query_stopped(self):
        self.wait_for(lambda: not self.namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(), timeout=0.3)

    def test_missing_terminal_does_not_wait_or_isolate(self):
        self.start_query()
        self.assertTrue(self.modem.wrote.wait(1))
        self.assert_query_stopped()
        self.assertEqual(self.modem.writes, [b"AT+CGSN\r\n"])
        self.assertIsNone(self.coordinator._active)
        self.assertFalse(serial_sender.serial_sms_recovery_error(self.modem))
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 106.0)
        self.assertTrue(serial_sender.write_serial_command_result(self.modem, "ATI").ok)

    def test_query_does_not_create_an_at_response_waiter(self):
        self.start_query()
        self.assertTrue(self.modem.wrote.wait(1))
        self.assertIsNone(self.coordinator._active)
        self.assert_query_stopped()

    def test_identity_data_remains_capturable_after_raw_worker_ends(self):
        self.start_query()
        self.assert_query_stopped()
        capture = lambda line: maybe_capture_cloud_device_imei_namespace_runtime(self.namespace, line)
        self.assertEqual(capture("[I]-[ril.proatc] OK"), "no_match")
        self.assertEqual(capture("AT+CGSN"), "ignored")
        self.assertEqual(capture("[I]-[ril.proatc] 000000000000001"), "captured")
        self.assertTrue(self.namespace["cloud_imei_verified"])
        self.assertEqual(self.namespace["CLOUD_DEVICE_IMEI"], "000000000000001")
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)
        self.assertNotIn("_cloud_imei_query_owner", self.namespace)
        self.assertEqual(capture("000000000000002"), "inactive")

    def test_expired_data_window_does_not_isolate_or_capture(self):
        self.start_query()
        self.assert_query_stopped()
        self.clock = 106.1
        self.assertEqual(maybe_capture_cloud_device_imei_namespace_runtime(
            self.namespace, "000000000000001"), "expired")
        self.assertFalse(self.namespace["cloud_imei_verified"])
        self.assertNotIn("_cloud_imei_query_owner", self.namespace)
        self.assertFalse(serial_sender.serial_sms_recovery_error(self.modem))

    def test_raw_write_failures_clear_only_the_query_window(self):
        for stage in ("write", "flush", "short"):
            with self.subTest(stage=stage):
                self.modem.write_error = OSError("synthetic write failure") if stage == "write" else None
                self.modem.flush_error = OSError("synthetic flush failure") if stage == "flush" else None
                self.modem.short_write = stage == "short"
                self.start_query()
                self.assert_query_stopped()
                self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)
                self.assertNotIn("_cloud_imei_query_owner", self.namespace)
                self.assertFalse(serial_sender.serial_sms_recovery_error(self.modem))
                self.assertTrue(any("读取IMEI失败" in log for log in self.logs))

    def test_existing_unknown_business_result_remains_isolated(self):
        self.modem._sms_recovery_error = "synthetic existing SMS unknown result"
        self.start_query()
        self.assert_query_stopped()
        self.assertEqual(self.modem.writes, [])
        self.assertEqual(self.modem._sms_recovery_error, "synthetic existing SMS unknown result")
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)
        _, status = cloud_channel_status_snapshot_namespace_runtime(self.namespace)
        self.assertFalse(status["control_available"])

    def test_stop_or_exit_while_waiting_for_final_write_lock_never_creates_window(self):
        for key in ("serial_stop_event", "TK_SHUTDOWN", "serial_running", "is_exiting"):
            with self.subTest(key=key):
                self.namespace["serial_stop_event"].clear()
                self.namespace["TK_SHUTDOWN"].clear()
                self.namespace["serial_running"] = True
                self.namespace["is_exiting"] = False
                with serial_sender.DEFAULT_SERIAL_WRITE_LOCK:
                    self.start_query()
                    self.wait_for(lambda: bool(self.namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot()))
                    self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)
                    if key.endswith("event") or key == "TK_SHUTDOWN":
                        self.namespace[key].set()
                    else:
                        self.namespace[key] = key == "is_exiting"
                self.assert_query_stopped()
                self.assertEqual(self.modem.writes, [])
                self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)
                self.assertNotIn("_cloud_imei_query_owner", self.namespace)
                self.assertFalse(serial_sender.serial_sms_recovery_error(self.modem))

    def test_connection_change_while_waiting_never_redirects_query(self):
        for change in ("serial", "generation"):
            with self.subTest(change=change):
                self.namespace["serial_obj"] = self.modem
                self.namespace["serial_connection_generation"] = 1
                replacement = QueryModem()
                with serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK:
                    self.start_query()
                    if change == "serial":
                        self.namespace["serial_obj"] = replacement
                    else:
                        self.namespace["serial_connection_generation"] = 2
                self.assert_query_stopped()
                self.assertEqual(self.modem.writes, [])
                self.assertEqual(replacement.writes, [])
                self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)

    def test_final_raw_write_guard_rechecks_original_connection(self):
        replacement = QueryModem()

        def change_before_writer(serial_obj, command, **kwargs):
            self.namespace["serial_obj"] = replacement
            return serial_sender.write_serial_command_result(serial_obj, command, **kwargs)

        self.namespace["write_serial_command_result"] = change_before_writer
        self.start_query()
        self.assert_query_stopped()
        self.assertEqual(self.modem.writes, [])
        self.assertEqual(replacement.writes, [])
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)

    def test_invalidated_original_connection_cannot_be_reused(self):
        forwarded = {}
        request_cloud_device_imei_namespace_runtime(
            self.namespace, request_runtime=lambda **kwargs: forwarded.update(kwargs))
        self.namespace["serial_connection_generation"] = 2
        self.assertIsNone(forwarded["get_serial"]())
        self.namespace["serial_connection_generation"] = 1
        self.assertIsNone(forwarded["get_serial"]())
        forwarded["set_query_deadline"](106.0)
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)

    def test_failed_old_query_cannot_clear_a_newer_equal_deadline(self):
        newer_owner = object()

        def newer_window_before_failure(serial_obj, command, **kwargs):
            self.namespace["_cloud_imei_query_owner"] = newer_owner
            self.namespace["cloud_imei_query_deadline"] = 106.0
            return serial_sender.SerialCommandResult(False, "synthetic failure after replacement")

        self.namespace["write_serial_command_result"] = newer_window_before_failure
        self.start_query()
        self.assert_query_stopped()
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 106.0)
        self.assertIs(self.namespace["_cloud_imei_query_owner"], newer_owner)

    def test_existing_ota_reservation_still_blocks_raw_query(self):
        owner = object()
        self.assertTrue(serial_sender.reserve_serial_for_ota(self.modem, owner))
        self.ota_owners.append((self.modem, owner))
        self.start_query()
        self.assert_query_stopped()
        self.assertEqual(self.modem.writes, [])
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)
        self.assertFalse(serial_sender.serial_sms_recovery_error(self.modem))

    def test_business_at_and_sms_mode_waiters_still_require_their_own_terminal(self):
        self.start_query()
        self.assert_query_stopped()
        for command in ("ATD10086;", "AT+CMGF=0"):
            with self.subTest(command=command):
                result = []
                worker = threading.Thread(target=lambda: result.append(
                    serial_sender.write_serial_command_confirmed_locked(
                        self.namespace["serial_lock"], lambda: self.modem, command,
                        response_coordinator=self.coordinator, response_timeout=1,
                        isolate_unconfirmed=True)))
                self.workers.append(worker)
                worker.start()
                payload = (command + "\r\n").encode("ascii")
                self.wait_for(lambda: payload in self.modem.writes)
                waiter = self.coordinator._active
                self.assertIsNotNone(waiter)
                self.assertEqual(waiter.command, command)
                self.assertFalse(waiter.done())
                self.assertTrue(worker.is_alive())
                self.coordinator.observe_line("OK", connection=self.modem)
                worker.join(1)
                self.assertFalse(worker.is_alive())
                self.assertTrue(result[0].ok)
                self.assertFalse(serial_sender.serial_sms_recovery_error(self.modem))

    def test_unknown_business_command_still_isolates_after_raw_query(self):
        self.start_query()
        self.assert_query_stopped()
        result = serial_sender.write_serial_command_confirmed_locked(
            self.namespace["serial_lock"], lambda: self.modem, "ATH",
            response_coordinator=self.coordinator, response_timeout=0.01,
            isolate_unconfirmed=True)
        self.assertFalse(result.ok)
        self.assertTrue(serial_sender.serial_sms_recovery_error(self.modem))
        self.assertFalse(serial_sender.write_serial_command_result(self.modem, "ATI").ok)

    def test_raw_query_window_does_not_make_open_channel_unavailable(self):
        self.start_query()
        self.assert_query_stopped()
        _, status = cloud_channel_status_snapshot_namespace_runtime(self.namespace)
        self.assertTrue(status["serial_connected"])
        self.assertTrue(status["control_available"])
        self.assertEqual(status["serial_connection_generation"], 1)

    def test_namespace_reply_during_flush_is_captured_without_reopening_window(self):
        processed = threading.Event()
        decisions, errors = [], []
        self.modem.in_waiting = 17

        def read(size):
            if not self.modem.wrote.wait(1):
                raise RuntimeError("synthetic query was not written")
            return b"000000000000001\r\n"

        def flush():
            if not processed.wait(0.5):
                raise RuntimeError("synthetic reply processing was blocked during flush")

        def reader():
            try:
                line = read_serial_line_safely_runtime(
                    self.namespace["serial_lock"], lambda: self.modem, RuntimeError,
                    read_lock=threading.Lock(),
                ).decode("ascii").strip()
                decisions.append(maybe_capture_cloud_device_imei_namespace_runtime(self.namespace, line))
            except Exception as exc:
                errors.append(type(exc).__name__)
            finally:
                processed.set()

        self.modem.read, self.modem.flush = read, flush
        worker = threading.Thread(target=reader)
        self.workers.append(worker)
        worker.start()
        self.start_query()
        self.assert_query_stopped()
        worker.join(1)
        self.assertFalse(worker.is_alive())
        self.assertEqual(errors, [])
        self.assertEqual(decisions, ["captured"])
        self.assertTrue(self.namespace["cloud_imei_verified"])
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 0.0)
        self.assertNotIn("_cloud_imei_query_owner", self.namespace)
        self.assertFalse(any("读取IMEI失败" in log for log in self.logs))

    def test_plain_connection_lock_and_channel_snapshot_do_not_invert_write_lock(self):
        held, attempted, proceed = threading.Event(), threading.Event(), threading.Event()
        errors, results = [], []

        class OrderedConnectionLock:
            def __init__(self):
                self.lock = threading.Lock()

            def __enter__(self):
                if threading.current_thread().name == "synthetic-imei-query":
                    attempted.set()
                if not self.lock.acquire(timeout=0.5):
                    raise RuntimeError("synthetic connection lock timed out")
                if threading.current_thread().name == "synthetic-channel-snapshot":
                    held.set()
                    proceed.wait(1)
                return self

            def __exit__(self, *args):
                self.lock.release()

        self.namespace["serial_lock"] = OrderedConnectionLock()
        forwarded = {}
        request_cloud_device_imei_namespace_runtime(
            self.namespace, request_runtime=lambda **kwargs: forwarded.update(kwargs))

        def channel():
            try:
                cloud_channel_status_snapshot_namespace_runtime(self.namespace)
            except Exception as exc:
                errors.append(type(exc).__name__)

        def query():
            keys = ("serial_lock", "get_serial", "write_command_result", "set_query_deadline",
                    "get_query_deadline", "cloud_log", "monotonic", "push_serial_debug")
            results.append(request_cloud_device_imei_worker(**{key: forwarded[key] for key in keys}))

        channel_worker = threading.Thread(target=channel, name="synthetic-channel-snapshot")
        query_worker = threading.Thread(target=query, name="synthetic-imei-query")
        self.workers.extend((channel_worker, query_worker))
        channel_worker.start()
        try:
            self.assertTrue(held.wait(1))
            query_worker.start()
            self.assertTrue(attempted.wait(1))
            available = serial_sender.DEFAULT_SERIAL_WRITE_LOCK.acquire(blocking=False)
            if available:
                serial_sender.DEFAULT_SERIAL_WRITE_LOCK.release()
            self.assertTrue(available, "IMEI must not own write while waiting for connection")
        finally:
            proceed.set()
            channel_worker.join(1)
            if query_worker.ident is not None:
                query_worker.join(1)
        self.assertFalse(channel_worker.is_alive())
        self.assertFalse(query_worker.is_alive())
        self.assertEqual(errors, [])
        self.assertEqual(results, [True])
        self.assertEqual(self.modem.writes, [b"AT+CGSN\r\n"])

    def test_failed_old_query_finishes_cleanup_before_real_reconnect_opens_new_window(self):
        release_flush, opened = threading.Event(), threading.Event()
        replacement = QueryModem()
        failures = []

        def fail_flush():
            if not release_flush.wait(1):
                raise RuntimeError("synthetic query flush was not released")
            raise OSError("synthetic old connection flush failure")

        def open_serial(*args, **kwargs):
            opened.set()
            return replacement

        self.modem.flush = fail_flush
        self.namespace.update(
            BAUD=115200, MODE="Manual", PORT="TEST",
            serial=SimpleNamespace(Serial=open_serial), lock_port_mutex=lambda port: None,
            unlock_port_mutex=lambda: None, serial_error_ui=lambda *args, **kwargs: None,
            set_status=lambda *args: None,
        )
        self.start_query()
        self.assertTrue(self.modem.wrote.wait(1))
        safe_close_serial_namespace_runtime(self.namespace)

        def reconnect():
            try:
                open_and_initialize_serial_namespace_runtime(self.namespace, "TEST")
            except Exception as exc:
                failures.append(type(exc).__name__)

        worker = threading.Thread(target=reconnect)
        self.workers.append(worker)
        worker.start()
        try:
            self.assertFalse(opened.wait(0.03))
            self.assertEqual(replacement.writes, [])
        finally:
            release_flush.set()
            worker.join(1)
        self.assert_query_stopped()
        self.assertFalse(worker.is_alive())
        self.assertEqual(failures, [])
        self.assertTrue(opened.is_set())
        self.assertEqual(replacement.writes, [b"AT+CLIP=1\r\n", b"AT+CGSN\r\n", b"AT+CNUM\r\n"])
        self.assertEqual(self.namespace["cloud_imei_query_deadline"], 106.0)
        self.assertNotIn("_cloud_imei_query_owner", self.namespace)


if __name__ == "__main__":
    unittest.main()
