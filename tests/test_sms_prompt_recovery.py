"""CMGS cancellation must restore or isolate the original modem transaction."""
from collections import defaultdict
import queue
import threading
import unittest
from unittest.mock import patch

from sms_core import serial_sender
from sms_core.cloud_command_context import set_cloud_command_authorized
from sms_core.cloud_message_namespace_runtime import (
    apply_cloud_modem_health_namespace_runtime,
    send_cloud_serial_command_namespace_runtime,
)
from sms_core.cloud_modem_health import CloudModemHealthState
from sms_core.serial_app_runtime import run_serial_reader_namespace_runtime
from sms_core.serial_io_runtime import read_serial_line_safely_runtime
from tests.test_serial_transaction_boundaries import FinalWriteGate


class StatefulModem:
    def __init__(self):
        self.is_open = True
        self.phase = "command"
        self.writes = []
        self.observe = lambda _line: None
        self.on_prompt = lambda: None
        self.prompt = True
        self.escape_reply = "OK"
        self.command_reply = "OK"
        self.fail_escape = False
        self.fail_pdu = False
        self.short_pdu = False
        self.fail_cmgs = False
        self.escape_written = threading.Event()
        self.command_written = threading.Event()
        self.pdu_written = threading.Event()
        self.close_count = 0

    def write(self, payload):
        self.writes.append(payload)
        if payload == b"\x1b":
            self.escape_written.set()
            if self.fail_escape:
                raise OSError("synthetic ESC write failure")
            if self.escape_reply:
                self.finish_abort(self.escape_reply)
        elif self.phase == "pdu":
            if self.fail_pdu:
                raise OSError("synthetic partial PDU write failure")
            if self.short_pdu:
                return len(payload) - 1
            if payload.endswith(b"\x1a"):
                self.phase = "command"
                self.pdu_written.set()
                self.observe("+CMGS: 1")
                self.observe("OK")
        elif payload.startswith(b"AT+CMGS="):
            self.phase = "pdu"
            if self.fail_cmgs:
                raise OSError("synthetic partial CMGS write failure")
            if self.prompt:
                self.observe(">")
                self.on_prompt()
        else:
            self.command_written.set()
            if self.command_reply:
                self.observe(self.command_reply)
        return len(payload)

    def flush(self):
        pass

    def finish_abort(self, terminal="OK"):
        self.phase = "command"
        self.observe(terminal)

    @property
    def in_waiting(self):
        return 0

    def read(self, _size):
        return b""

    def close(self):
        self.close_count += 1
        self.is_open = False


class SmsPromptRecoveryTests(unittest.TestCase):
    def fixture(self, sms=None):
        at = serial_sender.AtCommandResponseCoordinator()
        sms = sms or serial_sender.SmsPduSendCoordinator()
        modem = StatefulModem()
        ns = defaultdict(lambda: lambda *_args, **_kwargs: None)
        ns.update({
            "serial_lock": threading.RLock(), "serial_obj": modem,
            "serial_connection_generation": 1,
            "_cloud_runtime_imei": lambda: "000000000000001",
            "CLOUD_DEVICE_SECRET": "synthetic-only",
            "CLOUD_WS_URL": "wss://example.invalid/ws/device",
            "CLOUD_CONTROL_ENABLED": True,
            "cloud_ws_conn": object(), "cloud_connected": True,
            "cloud_device_authorized": True,
            "SMS_SEND_COORDINATOR": sms,
            "SERIAL_COMMAND_RESPONSE_COORDINATOR": at,
            "CLOUD_SENSITIVE_COMMAND_PERMISSIONS": {"sms": True},
            "_cloud_log": lambda *_args, **_kwargs: None,
            "FILE_LOG_Q": queue.Queue(),
            "serial_stop_event": threading.Event(),
            "serial_wakeup_event": threading.Event(),
        })
        modem.observe = run_serial_reader_namespace_runtime(
            ns, parse_callback_head=None, apply_disconnect_effects=None,
            run_reader=lambda **kwargs: kwargs["callbacks"]["observe_sms_send_line"],
        )
        return ns, modem, at, sms

    def send(self, ns, modem, sms, *, locked=True, **kwargs):
        options = dict(response_coordinator=sms,
                       command_response_coordinator=ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"],
                       sleep_func=lambda _seconds: None,
                       prompt_timeout=0.02, segment_timeout=0.02, abort_timeout=0.02)
        options.update(kwargs)
        if locked:
            return serial_sender.write_text_sms_pdu_locked(
                ns["serial_lock"], lambda: ns["serial_obj"], "10086", "synthetic-only",
                **options,
            )
        return serial_sender.write_text_sms_pdu(modem, "10086", "synthetic-only", **options)

    def next_command(self, ns, at, command="ATI"):
        return serial_sender.write_serial_command_confirmed_locked(
            ns["serial_lock"], lambda: ns["serial_obj"], command,
            response_coordinator=at, response_timeout=0.03,
        )

    def test_cloud_revocation_at_prompt_escapes_and_next_local_at_succeeds(self):
        for reauthorize in (False, True):
            with self.subTest(reauthorize=reauthorize):
                ns, modem, at, sms = self.fixture()

                def revoke():
                    set_cloud_command_authorized(ns, False)
                    if reauthorize:
                        set_cloud_command_authorized(ns, True)

                modem.on_prompt = revoke
                result = send_cloud_serial_command_namespace_runtime(
                    ns, "CLOUD:SEND_SMS", {"sms_phone": "10086", "sms_message": "synthetic-only"},
                )
                self.assertFalse(result[0])
                self.assertEqual(modem.phase, "command")
                self.assertEqual(modem.writes[-1], b"\x1b")
                self.assertFalse(modem.pdu_written.is_set())
                self.assertTrue(self.next_command(ns, at).ok)
                self.assertIsNone(sms._active)

    def test_cancel_and_prompt_timeout_restore_both_sender_paths(self):
        for locked in (False, True):
            for prompt in (False, True):
                for terminal in ("OK", "+CMS ERROR: 500"):
                    with self.subTest(locked=locked, prompt=prompt, terminal=terminal):
                        ns, modem, at, sms = self.fixture()
                        modem.prompt = prompt
                        modem.escape_reply = terminal
                        modem.on_prompt = lambda: sms.cancel_active("synthetic cancellation")
                        self.assertFalse(self.send(ns, modem, sms, locked=locked))
                        self.assertEqual(modem.writes[-1], b"\x1b")
                        self.assertFalse(modem.pdu_written.is_set())
                        self.assertTrue(self.next_command(ns, at).ok)

    def test_partial_cmgs_and_pdu_writes_escape_before_releasing_transaction(self):
        for stage in ("cmgs", "pdu"):
            for locked in (False, True):
                with self.subTest(stage=stage, locked=locked):
                    ns, modem, at, sms = self.fixture()
                    setattr(modem, "short_pdu" if stage == "pdu" else "fail_cmgs", True)
                    self.assertFalse(self.send(ns, modem, sms, locked=locked))
                    self.assertEqual(modem.writes[-1], b"\x1b")
                    self.assertTrue(self.next_command(ns, at).ok)

    def test_normal_sms_and_modem_error_do_not_send_escape(self):
        for locked in (False, True):
            for terminal_error in (False, True):
                with self.subTest(locked=locked, terminal_error=terminal_error):
                    ns, modem, at, sms = self.fixture()
                    if terminal_error:
                        modem.on_prompt = lambda: modem.finish_abort("+CMS ERROR: 500")
                    self.assertEqual(self.send(ns, modem, sms, locked=locked), not terminal_error)
                    self.assertNotIn(b"\x1b", modem.writes)
                    self.assertTrue(self.next_command(ns, at).ok)

    def test_cancellation_while_waiting_for_pdu_write_lock_stops_body(self):
        for locked in (False, True):
            with self.subTest(locked=locked):
                ns, modem, at, sms = self.fixture()
                gate = FinalWriteGate(pause_write=3)
                results = []
                with patch.object(serial_sender, "DEFAULT_SERIAL_WRITE_LOCK", gate):
                    worker = threading.Thread(target=lambda: results.append(self.send(ns, modem, sms, locked=locked)))
                    worker.start()
                    try:
                        self.assertTrue(gate.reached.wait(1))
                        self.assertTrue(sms.cancel_active("cancel during final write lock wait"))
                    finally:
                        gate.resume.set()
                        worker.join(2)
                self.assertFalse(worker.is_alive())
                self.assertEqual(results, [False])
                self.assertFalse(modem.pdu_written.is_set())
                self.assertEqual(modem.writes[-1], b"\x1b")
                self.assertTrue(self.next_command(ns, at).ok)

    def test_delayed_abort_terminal_is_consumed_before_the_next_at_reply(self):
        ns, modem, at, sms = self.fixture()
        modem.on_prompt = lambda: sms.cancel_active("synthetic cancellation")
        modem.escape_reply = None
        results, next_results = [], []
        send_worker = threading.Thread(target=lambda: results.append(self.send(ns, modem, sms, abort_timeout=1)))
        send_worker.start()
        next_worker = None
        try:
            self.assertTrue(modem.escape_written.wait(1))
            modem.command_reply = None
            modem.command_written.clear()
            next_worker = threading.Thread(target=lambda: next_results.append(self.next_command(ns, at)))
            next_worker.start()
            self.assertFalse(modem.command_written.wait(0.03))

            # Pause this read after it wakes the old SMS sender. A new AT can
            # start before the old reader callback returns, but cannot reuse OK.
            original_observe = sms.observe_line

            def observe_and_pause(line, **kwargs):
                consumed = original_observe(line, **kwargs)
                if line == "OK":
                    self.assertTrue(modem.command_written.wait(1))
                return consumed

            with patch.object(sms, "observe_line", observe_and_pause):
                modem.finish_abort()
            self.assertFalse(at._active.done())
            modem.observe("ERROR")
        finally:
            send_worker.join(2)
            if next_worker is not None:
                next_worker.join(2)
        self.assertFalse(send_worker.is_alive())
        self.assertFalse(next_worker.is_alive())
        self.assertEqual(results, [False])
        self.assertEqual(len(next_results), 1)
        self.assertFalse(next_results[0].ok)
        self.assertEqual(next_results[0].error, "ERROR")

    def test_failed_or_unconfirmed_escape_is_bounded_and_blocks_all_modem_writers(self):
        for failed_write in (False, True):
            with self.subTest(failed_write=failed_write):
                ns, modem, at, sms = self.fixture()
                modem.on_prompt = lambda: sms.cancel_active("synthetic cancellation")
                modem.fail_escape = failed_write
                modem.escape_reply = None
                self.assertFalse(self.send(ns, modem, sms))
                offset = len(modem.writes)
                self.assertTrue(serial_sender.serial_sms_recovery_error(modem))
                self.assertTrue(modem.is_open)
                self.assertEqual(modem.close_count, 0)
                self.assertIsNone(sms._active)
                self.assertFalse(self.next_command(ns, at).ok)
                debug = []
                for command in ("ATI", "ATA", "ATH", "ATD*100#"):
                    self.assertFalse(serial_sender.write_serial_command_result(modem, command, push_debug=debug.append).ok)
                self.assertTrue(all("重启设备后重新连接" in line for line in debug))
                self.assertEqual(len(debug), 4)
                self.assertFalse(serial_sender.write_serial_command_sequence(modem, ["ATI"]))
                self.assertFalse(serial_sender.write_serial_command_sequence_locked(ns["serial_lock"], lambda: modem, ["ATI"]))
                self.assertFalse(self.send(ns, modem, sms, locked=False))
                self.assertFalse(self.send(ns, modem, sms, locked=True))
                self.assertEqual(len(modem.writes), offset)
                self.assertEqual(read_serial_line_safely_runtime(ns["serial_lock"], lambda: modem, OSError), b"")

                ns["CLOUD_MODEM_HEALTH"] = CloudModemHealthState(reconnect_threshold=1)
                with patch.dict(ns, {"safe_close_serial": lambda: self.fail("isolation triggered reconnect")}):
                    apply_cloud_modem_health_namespace_runtime(ns, (False, serial_sender.serial_sms_recovery_error(modem)))
                ns["serial_obj"] = StatefulModem()
                ns["serial_obj"].observe = lambda line: at.observe_line(line, connection=ns["serial_obj"])
                self.assertTrue(self.next_command(ns, at).ok)

    def test_replacement_or_generation_change_at_cleanup_lock_never_writes_escape(self):
        for replace in (False, True):
            with self.subTest(replace=replace):
                ns, modem, _at, sms = self.fixture()
                modem.on_prompt = lambda: set_cloud_command_authorized(ns, False)
                # Authorization rejects before entering the PDU write lock, so
                # the next final lock is the recovery ESC itself.
                gate = FinalWriteGate(pause_write=3)
                results = []
                replacement = StatefulModem()
                with patch.object(serial_sender, "DEFAULT_SERIAL_WRITE_LOCK", gate):
                    worker = threading.Thread(target=lambda: results.append(send_cloud_serial_command_namespace_runtime(
                        ns, "CLOUD:SEND_SMS", {"sms_phone": "10086", "sms_message": "synthetic-only"})))
                    worker.start()
                    try:
                        self.assertTrue(gate.reached.wait(1))
                        ns["serial_connection_generation"] += 1
                        if replace:
                            ns["serial_obj"] = replacement
                    finally:
                        gate.resume.set()
                        worker.join(2)
                self.assertFalse(worker.is_alive())
                self.assertFalse(results[0][0])
                self.assertNotIn(b"\x1b", modem.writes)
                self.assertEqual(replacement.writes, [])
                self.assertFalse(serial_sender.serial_sms_recovery_error(modem))
                self.assertFalse(serial_sender.serial_sms_recovery_error(replacement))
                self.assertIsNone(sms._active)

    def test_legacy_coordinator_signature_can_drain_cancellation(self):
        class LegacyCoordinator(serial_sender.SmsPduSendCoordinator):
            def begin_segment(self, label="", connection=None, cmgs_len=None):
                return super().begin_segment(label, connection, cmgs_len)

        ns, modem, at, sms = self.fixture(LegacyCoordinator())
        modem.on_prompt = lambda: sms.cancel_active("synthetic cancellation")
        self.assertFalse(self.send(ns, modem, sms))
        self.assertTrue(self.next_command(ns, at).ok)

    def test_post_pdu_timeout_waits_for_original_terminal_without_escape(self):
        ns, modem, at, sms = self.fixture()
        original_write = modem.write
        results = []

        def defer_pdu_result(payload):
            if payload.endswith(b"\x1a"):
                modem.writes.append(payload)
                modem.phase = "command"
                modem.pdu_written.set()
                return len(payload)
            return original_write(payload)

        modem.write = defer_pdu_result
        worker = threading.Thread(target=lambda: results.append(self.send(ns, modem, sms, abort_timeout=1)))
        worker.start()
        try:
            self.assertTrue(modem.pdu_written.wait(1))
            self.assertTrue(sms._active._event.wait(1))
            self.assertTrue(worker.is_alive())
            self.assertNotIn(b"\x1b", modem.writes)
            modem.observe("+CMGS: 1")
            modem.observe("OK")
        finally:
            worker.join(2)
        self.assertFalse(worker.is_alive())
        self.assertEqual(results, [False])
        self.assertFalse(serial_sender.serial_sms_recovery_error(modem))
        self.assertTrue(self.next_command(ns, at).ok)

    def test_untrusted_or_other_connection_terminal_cannot_release_abort(self):
        for other_connection in (False, True):
            with self.subTest(other_connection=other_connection):
                ns, modem, _at, sms = self.fixture()
                modem.on_prompt = lambda: sms.cancel_active("synthetic cancellation")
                modem.escape_reply = None
                original_write = modem.write

                def invalid_terminal(payload):
                    count = original_write(payload)
                    if payload == b"\x1b":
                        if other_connection:
                            sms.observe_line("OK", connection=object())
                        else:
                            modem.observe("[I]-[app.note] OK")
                    return count

                modem.write = invalid_terminal
                self.assertFalse(self.send(ns, modem, sms))
                self.assertTrue(serial_sender.serial_sms_recovery_error(modem))

    def test_cloud_error_includes_manual_recovery_action_when_escape_fails(self):
        ns, modem, _at, _sms = self.fixture()
        modem.on_prompt = lambda: set_cloud_command_authorized(ns, False)
        modem.fail_escape = True
        result = send_cloud_serial_command_namespace_runtime(
            ns, "CLOUD:SEND_SMS", {"sms_phone": "10086", "sms_message": "synthetic-only"},
        )
        self.assertFalse(result[0])
        self.assertIn("重启设备后重新连接", result[1])
        self.assertTrue(modem.is_open)

    def test_short_escape_write_is_not_treated_as_successful_cancellation(self):
        ns, modem, _at, sms = self.fixture()
        modem.on_prompt = lambda: sms.cancel_active("synthetic cancellation")
        original_write = modem.write
        modem.write = lambda payload: 0 if payload == b"\x1b" else original_write(payload)
        self.assertFalse(self.send(ns, modem, sms))
        self.assertTrue(serial_sender.serial_sms_recovery_error(modem))


if __name__ == "__main__":
    unittest.main()
