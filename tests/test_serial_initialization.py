import queue
import re
import threading
import time
import unittest
from contextlib import ExitStack
from types import SimpleNamespace
from unittest.mock import Mock, patch

from sms_core import serial_sender
from sms_core.cloud_payloads import identity_payload
from sms_core.cloud_protocol import normalize_imei, parse_sms_callback_head
from sms_core.cloud_state_namespace_runtime import (
    cloud_identity_payload_namespace_runtime, cloud_runtime_imei_namespace_runtime,
    maybe_capture_cloud_device_imei_namespace_runtime, set_cloud_device_imei_namespace_runtime,
)
from sms_core.serial_app_runtime import run_serial_reader_namespace_runtime
from sms_core.serial_io_runtime import read_serial_line_safely_runtime
from sms_core.serial_namespace_runtime import (
    open_and_initialize_serial_namespace_runtime, schedule_delayed_connected_log_namespace_runtime,
)
from sms_ui.app_infrastructure_namespace_runtime import safe_close_serial_namespace_runtime
from sms_ui.call_popup_namespace_runtime import (
    get_serial_call_state_namespace_runtime, merge_serial_call_state_namespace_runtime,
    set_serial_call_state_namespace_runtime,
)
from tests import test_local_call_command_isolation as call_fixtures


class InitializationHarness:
    def __init__(self, *, hold="AT+CNUM", stage="normal", connected_error=False,
                 error_message="synthetic-private-startup-error"):
        self.fixture = call_fixtures.LocalDialButtonIsolationTests("runTest")
        self.fixture.setUp()
        self.ns = namespace = self.fixture.ns
        self.stop = threading.Event()
        self.ready = threading.Event()
        self.clock = 100.0
        self.errors, self.messages, self.lines, self.recording_lines = [], [], [], []
        self.statuses = []
        self.on_line = None
        self.connected_error = connected_error
        self.reads = queue.Queue()
        self.hold, self.stage = hold, stage
        self.wrote = {command: threading.Event() for command in ("AT+CLIP=1", "AT+CGSN", "AT+CNUM", "ATD10086;")}
        harness = self

        class Modem:
            is_open = True

            def __init__(self):
                self.writes = []

            @property
            def in_waiting(self):
                return 4096

            def write(self, payload):
                self.writes.append(payload)
                command = payload.decode("ascii").strip()
                if command in harness.wrote:
                    harness.wrote[command].set()
                if command.startswith("AT+"):
                    if harness.stage == "write-error":
                        raise OSError(error_message)
                    if harness.stage == "partial":
                        return 1
                    if command != harness.hold:
                        harness.terminal(command)
                return len(payload)

            def flush(self):
                if harness.stage == "flush-error":
                    raise OSError(error_message)

            def read(self, _size):
                try:
                    return harness.reads.get(timeout=0.02)
                except queue.Empty:
                    return b""

            def close(self):
                self.is_open = False

        self.modem = Modem()
        namespace.update(
            serial_obj=None, serial_read_lock=threading.Lock(), serial_stop_event=self.stop,
            serial_wakeup_event=threading.Event(), MODE="Manual", PORT="TEST", BAUD=115200,
            serial=SimpleNamespace(Serial=lambda *args, **kwargs: self.modem, SerialException=OSError),
            KEYWORDS=[], LOG_UNMATCHED_SMS=False, LOG_DIR="unused", LOG_PREFIX="TEST",
            ERROR_REPEAT_LIMIT=3, CALL_FILTER_MODE="Disabled", CALL_WHITELIST=[], CALL_BLACKLIST=[],
            FILE_LOG_Q=queue.Queue(), current_call_popup=None, _sms_ignore_repeat_state={},
            serial_running=True, RECONNECT_INTERVAL=1,
            CLOUD_DEVICE_IMEI="", cloud_imei_verified=False, cloud_imei_query_deadline=0.0,
            IMEI_REGEX=re.compile(r"\b(\d{14,17})\b"), _normalize_imei=normalize_imei,
            _cloud_identity_payload_core=identity_payload, APP_VERSION="4.0.2",
            _auto_connect_notice=SimpleNamespace(reset=lambda: None),
            _serial_error_notice=SimpleNamespace(clear=lambda: None),
            root=SimpleNamespace(after=lambda *args: None), ui_post=lambda callback: None,
            status_var=SimpleNamespace(get=lambda: "connecting"), APP_START_MONO=0.0, START_UI_DELAY=0.0,
            time=SimpleNamespace(monotonic=lambda: self.clock),
        )
        for name in (
            "enqueue_third_push", "_cloud_send_sms_event", "play_alert", "show_sms_popup", "system_ui",
            "set_temperature", "set_signal", "close_call_popup", "send_call_hangup_command",
            "show_call_popup", "serial_error_ui", "lock_port_mutex", "unlock_port_mutex",
            "_notify_cloud_identity_changed", "notify_cloud_channel_status", "_cloud_log", "set_status",
        ):
            namespace[name] = lambda *args, **kwargs: None
        namespace["port_ui"] = lambda text, *args: self.messages.append(text)
        namespace["set_status"] = lambda text, color: self.statuses.append((text, color))
        namespace["serial_error_ui"] = lambda text, *args, **kwargs: self.messages.append(text)
        namespace["_push_serial_debug"] = lambda text: self.lines.append(text)
        namespace["_cloud_send_serial_log"] = self.observe_line
        namespace["consume_call_recording_line"] = lambda line: self.recording_lines.append(line) or line.startswith("@@CALL_RECORD")
        namespace["_cloud_runtime_imei"] = lambda: cloud_runtime_imei_namespace_runtime(namespace)
        namespace["_cloud_identity_payload"] = lambda: cloud_identity_payload_namespace_runtime(namespace)
        namespace["_set_cloud_device_imei"] = lambda value, **kwargs: set_cloud_device_imei_namespace_runtime(namespace, value, **kwargs)
        namespace["_maybe_capture_cloud_device_imei"] = lambda line: maybe_capture_cloud_device_imei_namespace_runtime(namespace, line)
        namespace["set_serial_log_prefix"] = lambda value: namespace.update(LOG_PREFIX=value)
        namespace["get_serial_call_state"] = lambda: get_serial_call_state_namespace_runtime(namespace)
        namespace["set_serial_call_state"] = lambda *args: set_serial_call_state_namespace_runtime(namespace, *args)
        namespace["merge_serial_call_state"] = lambda *args: merge_serial_call_state_namespace_runtime(namespace, *args)
        namespace["resolve_serial_target_port"] = lambda: "TEST"
        namespace["open_and_initialize_serial"] = lambda port: open_and_initialize_serial_namespace_runtime(namespace, port)
        namespace["read_serial_line_safely"] = lambda: read_serial_line_safely_runtime(
            namespace["serial_lock"], lambda: namespace["serial_obj"], OSError,
            read_lock=namespace["serial_read_lock"],
        )
        namespace["safe_close_serial"] = lambda: safe_close_serial_namespace_runtime(namespace)
        namespace["try_manual_rebind_after_error"] = lambda error: False
        namespace["schedule_delayed_connected_log"] = lambda port, baud: schedule_delayed_connected_log_namespace_runtime(
            namespace, port, baud, start_runtime=self.connected,
        )

    def connected(self, *args, **kwargs):
        self.connected_kwargs = kwargs
        self.opened_session = self.ns.get("_serial_initialization_session")
        self.ready.set()
        if self.connected_error:
            self.stop.set()
            raise RuntimeError("synthetic connected callback failure")

    def terminal(self, command, *, response="OK", extra=""):
        body = "[I]-[ril.proatc] 000000000000001\r\n" if command == "AT+CGSN" else ""
        self.reads.put((body + "[I]-[ril.proatc] " + response + "\r\n" + extra).encode("ascii"))

    def emit(self, text):
        self.reads.put(text.encode("ascii"))

    def observe_line(self, text, *args):
        self.lines.append(text)
        if self.on_line is not None:
            self.on_line(text)

    def __enter__(self):
        self.stack = ExitStack()
        self.fixture.open_window(self.stack)

        def reader():
            try:
                run_serial_reader_namespace_runtime(
                    self.ns, parse_callback_head=parse_sms_callback_head,
                    apply_disconnect_effects=lambda *args: None,
                )
            except Exception as exc:
                self.errors.append(type(exc).__name__)

        self.thread = threading.Thread(target=reader)
        self.thread.start()
        if not self.ready.wait(2):
            self.__exit__(None, None, None)
            raise RuntimeError("formal reader did not establish the connection")
        return self

    def __exit__(self, *_args):
        self.stop.set()
        self.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"].cancel_active("synthetic stop")
        self.thread.join(2)
        for worker in self.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot():
            worker.join(2)
        self.stack.close()
        if self.thread.is_alive():
            raise RuntimeError("formal reader did not release its initialization owner")


class SerialInitializationTests(unittest.TestCase):
    def wait_for(self, predicate):
        deadline = time.monotonic() + 2
        while time.monotonic() < deadline:
            if predicate():
                return
            threading.Event().wait(0.005)
        self.fail("virtual initialization condition was not reached")

    def test_each_initialization_command_has_its_own_terminal_before_next_submission(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            self.assertTrue(harness.wrote["AT+CLIP=1"].wait(1))
            self.assertEqual(harness.modem.writes, [b"AT+CLIP=1\r\n"])
            harness.hold = "AT+CGSN"
            harness.clock += 4.0
            harness.terminal("AT+CLIP=1", extra="[I]-[ril.proatc] OK\r\n")
            self.assertTrue(harness.wrote["AT+CGSN"].wait(1))
            self.assertFalse(harness.wrote["AT+CNUM"].is_set())
            self.assertEqual(harness.ns["cloud_imei_query_deadline"], harness.clock + 6.0)
            harness.hold = "AT+CNUM"
            harness.terminal("AT+CGSN")
            self.assertTrue(harness.wrote["AT+CNUM"].wait(1))
            harness.terminal("AT+CNUM")
            self.wait_for(lambda: harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active is None)
            self.assertEqual(harness.errors, [])

    def test_pending_startup_terminal_cannot_confirm_queued_ui_voice(self):
        with InitializationHarness() as harness:
            self.assertTrue(harness.wrote["AT+CNUM"].wait(1))
            self.wait_for(lambda: harness.ns["cloud_imei_verified"])
            self.assertEqual(harness.ns["serial_connection_generation"], 2)
            harness.fixture.actions["dial"]("10086")
            self.assertFalse(harness.wrote["ATD10086;"].wait(0.03))
            harness.terminal("AT+CNUM")
            self.assertTrue(harness.wrote["ATD10086;"].wait(1))
            self.assertTrue(harness.ns["CLOUD_SMS_EVENT_Q"].empty())
            harness.terminal("ATD10086;")
            self.wait_for(lambda: harness.ns["CLOUD_SMS_EVENT_Q"].qsize() == 1)
            self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))

    def test_pending_initialization_keeps_reader_lines_and_rejects_reentrant_ota(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            results = []
            owner = object()

            def inspect(line, *args):
                harness.lines.append(line)
                if line == "PROBE":
                    results.append(serial_sender.reserve_serial_for_ota(harness.modem, owner))
                    results.append(serial_sender.write_serial_command_result(harness.modem, "ATH"))
                    serial_sender.release_serial_from_ota(harness.modem, owner)

            harness.on_line = inspect
            harness.emit("PROBE\r\nRING\r\nSMS-BODY\r\n@@CALL_RECORD_META|virtual|incoming\r\n")
            self.wait_for(lambda: any(line.startswith("@@CALL_RECORD_META") for line in harness.recording_lines))
            self.assertIn("RING", harness.recording_lines)
            self.assertIn("SMS-BODY", harness.recording_lines)
            self.assertEqual(results[0], False)
            self.assertFalse(results[1].ok)
            self.assertIn("初始化", results[1].error)
            self.assertNotIn(b"ATH\r\n", harness.modem.writes)

    def test_generation_change_terminates_original_initialization(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            self.assertTrue(harness.wrote["AT+CLIP=1"].wait(1))
            harness.ns["serial_connection_generation"] += 1
            harness.terminal("AT+CLIP=1")
            self.wait_for(lambda: bool(serial_sender.serial_sms_recovery_error(harness.modem)))
            self.assertFalse(harness.wrote["AT+CGSN"].is_set())
            self.assertFalse(harness.wrote["AT+CNUM"].is_set())

    def test_timeout_releases_owner_and_only_quarantines_original_connection(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            self.assertTrue(harness.wrote["AT+CLIP=1"].wait(1))
            harness.clock += 10.0
            self.wait_for(lambda: bool(serial_sender.serial_sms_recovery_error(harness.modem)))
            self.assertFalse(serial_sender.reserve_serial_for_ota(harness.modem, object()))
            acquired = []
            def probe():
                available = serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK.acquire(timeout=0.2)
                acquired.append(available)
                if available:
                    serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK.release()
            worker = threading.Thread(target=probe)
            worker.start()
            worker.join(1)
            self.assertEqual(acquired, [True])
            self.assertTrue(harness.modem.is_open)

    def test_shutdown_cancels_startup_waiter_and_releases_transaction(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            self.assertTrue(harness.wrote["AT+CLIP=1"].wait(1))
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
        self.assertEqual(acquired, [True])

    def assert_transaction_available(self):
        acquired = []
        def probe():
            available = serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK.acquire(timeout=0.2)
            acquired.append(available)
            if available:
                serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK.release()
        worker = threading.Thread(target=probe)
        worker.start()
        worker.join(1)
        self.assertEqual(acquired, [True])

    def test_sensitive_write_partial_and_flush_failures_release_owner_and_keep_reader(self):
        private = "CLOUD_DEVICE_SECRET=synthetic-only-secret;PIN=123456;wss://u:synthetic-only-secret@example.invalid"
        for stage in ("write-error", "partial", "flush-error"):
            with self.subTest(stage=stage):
                with InitializationHarness(hold="AT+CLIP=1", stage=stage, error_message=private) as harness:
                    self.wait_for(lambda: bool(serial_sender.serial_sms_recovery_error(harness.modem)))
                    self.assertFalse(harness.opened_session.pending)
                    self.assertFalse(harness.opened_session.lock_held)
                    self.assert_transaction_available()
                    self.assertFalse(serial_sender.reserve_serial_for_ota(harness.modem, object()))
                    harness.emit("STILL-READING\r\n")
                    self.wait_for(lambda: "STILL-READING" in harness.lines)
                    combined = "\n".join(harness.messages + harness.lines)
                    for value in ("synthetic-only-secret", "123456", "example.invalid"):
                        self.assertNotIn(value, combined)
                    self.assertFalse(harness.wrote["AT+CGSN"].is_set())

    def test_gui_close_cancels_owner_without_submitting_remaining_initialization(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            self.assertTrue(harness.wrote["AT+CLIP=1"].wait(1))
            session = harness.opened_session
            safe_close_serial_namespace_runtime(harness.ns)
            harness.stop.set()
            self.wait_for(lambda: not session.pending)
            self.assertFalse(session.lock_held)
            self.assertFalse(harness.wrote["AT+CGSN"].is_set())
            self.assertFalse(harness.modem.is_open)
            self.assert_transaction_available()

    def test_external_close_between_commands_is_latched_before_next_read(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            between, resume = threading.Event(), threading.Event()
            def pause(line):
                if line.endswith(" OK"):
                    between.set()
                    resume.wait(1)
            harness.on_line = pause
            harness.terminal("AT+CLIP=1")
            self.assertTrue(between.wait(1))
            session = harness.opened_session
            self.assertIsNone(session.waiter)
            session.close()
            resume.set()
            self.wait_for(lambda: not session.pending)
            self.assertFalse(harness.wrote["AT+CGSN"].is_set())
            self.assertFalse(session.lock_held)
            self.wait_for(lambda: bool(harness.statuses and "初始化未完成" in harness.statuses[-1][0]))
            self.assertEqual(harness.statuses[-1][1], "orange")
            self.assert_transaction_available()

    def test_waiter_finish_failure_does_not_repeat_startup_or_leak_owner(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            coordinator = harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]
            self.assertTrue(harness.wrote["AT+CLIP=1"].wait(1))
            waiter = harness.opened_session.waiter
            with patch.object(coordinator, "finish", side_effect=RuntimeError("synthetic-private-finish")):
                harness.terminal("AT+CLIP=1")
                self.wait_for(lambda: not harness.opened_session.pending)
            self.assertTrue(waiter.done())
            self.assertFalse(harness.opened_session.lock_held)
            self.assertEqual(harness.modem.writes, [b"AT+CLIP=1\r\n"])
            self.assert_transaction_available()
            self.assertNotIn("synthetic-private-finish", "\n".join(harness.messages))
            coordinator.finish(waiter)

    def test_property_cleanup_failure_still_releases_owner(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            modem_type = type(harness.modem)
            def get_session(modem):
                return modem.__dict__.get("_serial_initialization_session")
            def set_session(modem, value):
                if value is None:
                    raise OSError("synthetic-private-property")
                modem.__dict__["_serial_initialization_session"] = value
            with patch.object(modem_type, "_serial_initialization_session", property(get_session, set_session), create=True):
                harness.hold = ""
                harness.terminal("AT+CLIP=1")
                self.wait_for(lambda: not harness.opened_session.pending)
                self.assertFalse(harness.opened_session.lock_held)
                self.assert_transaction_available()
                self.assertFalse(serial_sender.serial_initialization_pending(harness.modem))
            self.assertNotIn("synthetic-private-property", "\n".join(harness.messages))

    def test_opened_session_is_released_if_connected_callback_fails_before_first_read(self):
        with InitializationHarness(connected_error=True) as harness:
            harness.thread.join(1)
            self.assertFalse(harness.opened_session.pending)
            self.assertFalse(harness.opened_session.lock_held)
            self.assertEqual(harness.modem.writes, [])
            self.assertIsNone(harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active)
            self.assert_transaction_available()

    def test_foreign_session_slot_does_not_receive_old_reader_cleanup_or_progress(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            original = harness.opened_session
            self.assertTrue(harness.wrote["AT+CLIP=1"].wait(1))
            foreign = SimpleNamespace(owner_thread=-1, pending=True, lock_held=False,
                                      close=Mock(), read_boundary=Mock(), observe_after_line=Mock())
            harness.ns["_serial_initialization_session"] = foreign
            harness.emit("FOREIGN-SLOT\r\n")
            self.wait_for(lambda: "FOREIGN-SLOT" in harness.lines)
            harness.stop.set()
            harness.thread.join(1)
            self.assertFalse(original.pending)
            self.assertFalse(original.lock_held)
            foreign.close.assert_not_called()
            foreign.read_boundary.assert_not_called()
            foreign.observe_after_line.assert_not_called()
            self.assertIs(harness.ns["_serial_initialization_session"], foreign)
            self.assert_transaction_available()

    def test_ota_renewal_is_rejected_for_pending_or_quarantined_original_serial(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            self.assertTrue(harness.wrote["AT+CLIP=1"].wait(1))
            owner = object()
            expiry = time.monotonic() + 60
            with patch.dict(serial_sender._OTA_RESERVATIONS, {id(harness.modem): (harness.modem, owner, expiry)}):
                self.assertFalse(serial_sender.reserve_serial_for_ota(harness.modem, owner))
                self.assertEqual(serial_sender._OTA_RESERVATIONS[id(harness.modem)][2], expiry)
                harness.clock += 10
                self.wait_for(lambda: not harness.opened_session.pending)
                self.assertFalse(serial_sender.reserve_serial_for_ota(harness.modem, owner))
                self.assertEqual(serial_sender._OTA_RESERVATIONS[id(harness.modem)][2], expiry)

    def test_old_initialization_cleanup_does_not_cancel_replacement_waiter_or_reservation(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            self.assertTrue(harness.wrote["AT+CLIP=1"].wait(1))
            original = harness.opened_session
            coordinator = harness.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]
            coordinator.cancel_active("original connection replaced")
            coordinator.finish(original.waiter)
            replacement = type(harness.modem)()
            harness.ns.update(serial_obj=replacement, serial_connection_generation=3)
            new_waiter = coordinator.begin(command="ATI", connection=replacement)
            new_waiter.mark_write_started()
            self.wait_for(lambda: not original.pending)
            self.assertIs(coordinator._active, new_waiter)
            self.assertFalse(new_waiter.done())
            self.assertFalse(serial_sender.serial_sms_recovery_error(replacement))
            owner = object()
            try:
                self.assertTrue(serial_sender.reserve_serial_for_ota(replacement, owner))
            finally:
                serial_sender.release_serial_from_ota(replacement, owner)
            self.assertFalse(harness.wrote["AT+CGSN"].is_set())

    def test_connected_status_remains_initializing_until_true_settlement_without_new_generation(self):
        with InitializationHarness(hold="AT+CNUM") as harness:
            self.assertTrue(harness.wrote["AT+CNUM"].wait(1))
            callback = harness.connected_kwargs["set_status"]
            callback("connected", "green")
            self.assertIn("初始化中", harness.statuses[-1][0])
            self.assertEqual(harness.statuses[-1][1], "orange")
            harness.terminal("AT+CNUM")
            self.wait_for(lambda: not harness.opened_session.pending)
            self.wait_for(lambda: bool(harness.statuses and harness.statuses[-1][1] == "green"))
            self.assertEqual(harness.ns["serial_connection_generation"], 2)

    def test_single_blocked_clip_is_deferred_by_real_reader_until_initialization_finishes(self):
        with InitializationHarness(hold="AT+CLIP=1") as harness:
            harness.ns.update(CALL_FILTER_MODE="Blacklist", CALL_BLACKLIST=["10086"])
            harness.emit('RING\r\n+CLIP: "10086",129\r\n')
            self.wait_for(lambda: len(harness.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot()) == 1)
            self.assertNotIn(b"ATH\r\n", harness.modem.writes)
            harness.hold = ""
            harness.terminal("AT+CLIP=1")
            self.wait_for(lambda: b"ATH\r\n" in harness.modem.writes)
            harness.terminal("ATH")
            self.wait_for(lambda: harness.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot() == ())
            self.assertEqual(harness.modem.writes.count(b"ATH\r\n"), 1)
            self.assertFalse(serial_sender.serial_sms_recovery_error(harness.modem))


if __name__ == "__main__":
    unittest.main()
