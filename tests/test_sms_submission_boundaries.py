"""CMGF acknowledgement and PDU submission are separate modem boundaries."""
import threading
import unittest
from unittest.mock import patch

from sms_core import serial_sender
from sms_core.sms_send_namespace_runtime import send_manual_sms_namespace_runtime
from tests import test_sms_prompt_recovery as recovery


class SmsSubmissionBoundaryTests(unittest.TestCase):
    def fixture(self):
        return recovery.SmsPromptRecoveryTests().fixture()

    def start_send(self, path, ns, modem, sms, outcomes, **options):
        options = dict(response_coordinator=sms, prompt_timeout=0.2,
                       segment_timeout=0.5, abort_timeout=0.2,
                       sleep_func=lambda _seconds: None) | options
        if path in ("async", "manual"):
            options["on_sent"] = lambda: outcomes.append(True)
            if path == "async":
                return serial_sender.send_text_sms_pdu_async(
                    ns["serial_lock"], lambda: modem, "10086", "synthetic-only", **options)
            ns["_cloud_identity_payload"] = lambda: {"imei": "000000000000001"}
            ns["CLOUD_CONTROL_ENABLED"] = False

            def actual_async(*args, **kwargs):
                kwargs.update(options)
                return serial_sender.send_text_sms_pdu_async(*args, **kwargs)

            return send_manual_sms_namespace_runtime(
                ns, "10086", "synthetic-only", send_runtime=actual_async)

        def send():
            if path == "locked":
                result = serial_sender.write_text_sms_pdu_locked(
                    ns["serial_lock"], lambda: modem, "10086", "synthetic-only", **options)
            else:
                result = serial_sender.write_text_sms_pdu(
                    modem, "10086", "synthetic-only", **options)
            outcomes.append(result)

        worker = threading.Thread(target=send)
        worker.start()
        return worker

    def test_cmgf_acknowledgement_precedes_cmgs_for_every_entry(self):
        for path in ("direct", "locked", "async", "manual"):
            with self.subTest(path=path):
                ns, modem, at, sms = self.fixture()
                modem.command_reply = None
                cmgf, cmgs = threading.Event(), threading.Event()
                original_write = modem.write

                def write(payload):
                    result = original_write(payload)
                    if payload == b"AT+CMGF=0\r\n":
                        cmgf.set()
                    elif payload.startswith(b"AT+CMGS="):
                        cmgs.set()
                    return result

                modem.write = write
                outcomes = []
                with patch.object(serial_sender, "DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR", at):
                    worker = self.start_send(path, ns, modem, sms, outcomes)
                    try:
                        self.assertTrue(cmgf.wait(1))
                        self.assertFalse(cmgs.wait(0.03), "CMGS submitted before CMGF terminal")
                        modem.observe("OK")
                    finally:
                        modem.observe("OK")
                        worker.join(2)
                self.assertFalse(worker.is_alive())
                self.assertEqual(outcomes, [True])
                self.assertTrue(cmgs.is_set())

    def test_cmgf_error_prevents_cmgs_for_every_entry(self):
        for path in ("direct", "locked", "async", "manual"):
            with self.subTest(path=path):
                ns, modem, at, sms = self.fixture()
                modem.command_reply = "ERROR"
                outcomes = []
                with patch.object(serial_sender, "DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR", at):
                    worker = self.start_send(path, ns, modem, sms, outcomes)
                    worker.join(2)
                self.assertFalse(worker.is_alive())
                self.assertEqual(outcomes, [] if path in ("async", "manual") else [False])
                self.assertEqual(modem.writes, [b"AT+CMGF=0\r\n"])
                self.assertFalse(serial_sender.serial_sms_recovery_error(modem))

    def test_unconfirmed_cmgf_isolates_original_serial(self):
        for path in ("direct", "locked", "async", "manual"):
            for failure in ("timeout", "cancel", "write_error"):
                with self.subTest(path=path, failure=failure):
                    ns, modem, at, sms = self.fixture()
                    modem.command_reply = None
                    original_write = modem.write

                    def write(payload):
                        result = original_write(payload)
                        if payload == b"AT+CMGF=0\r\n":
                            if failure == "cancel":
                                at.cancel_active("synthetic CMGF cancellation")
                            elif failure == "write_error":
                                raise OSError("synthetic CMGF write failure")
                        return result

                    modem.write = write
                    outcomes = []
                    with patch.object(serial_sender, "DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR", at):
                        worker = self.start_send(path, ns, modem, sms, outcomes, segment_timeout=0.02)
                        worker.join(1)
                    self.assertFalse(worker.is_alive())
                    self.assertEqual(outcomes, [] if path in ("async", "manual") else [False])
                    self.assertEqual(modem.writes, [b"AT+CMGF=0\r\n"])
                    self.assertTrue(serial_sender.serial_sms_recovery_error(modem))
                    self.assertIsNone(at._active)
                    modem.observe("OK")
                    self.assertFalse(serial_sender.write_serial_command_result(modem, "ATI").ok)

    def test_trusted_cmgf_terminal_before_flush_failure_preserves_all_entries(self):
        for path in ("direct", "locked", "async", "manual"):
            for terminal in ("OK", "ERROR"):
                with self.subTest(path=path, terminal=terminal):
                    ns, modem, at, sms = self.fixture()
                    modem.command_reply = terminal
                    faults = []

                    def flush():
                        if modem.writes[-1] == b"AT+CMGF=0\r\n" and not faults:
                            self.assertTrue(at._active.wait(0).line)
                            faults.append(True)
                            raise OSError("synthetic flush failure after trusted CMGF terminal")

                    modem.flush = flush
                    outcomes = []
                    worker = self.start_send(path, ns, modem, sms, outcomes,
                                             command_response_coordinator=at)
                    worker.join(2)
                    self.assertFalse(worker.is_alive())
                    expected = [True] if terminal == "OK" else ([] if path in ("async", "manual") else [False])
                    self.assertEqual(outcomes, expected)
                    self.assertEqual(faults, [True])
                    self.assertEqual(any(payload.startswith(b"AT+CMGS=") for payload in modem.writes), terminal == "OK")
                    self.assertFalse(serial_sender.serial_sms_recovery_error(modem))
                    self.assertIsNone(at._active)
                    self.assertIsNone(sms._active)
                    self.assertNotIn(b"\x1b", modem.writes)

    def test_confirmed_cmgf_flush_failure_cannot_restore_revoked_cloud_send(self):
        from sms_core.cloud_command_context import set_cloud_command_authorized
        from sms_core.cloud_message_namespace_runtime import send_cloud_serial_command_namespace_runtime

        for reauthorize in (False, True):
            with self.subTest(reauthorize=reauthorize):
                ns, modem, at, sms = self.fixture()

                def flush():
                    self.assertTrue(at._active.wait(0).line)
                    set_cloud_command_authorized(ns, False)
                    if reauthorize:
                        set_cloud_command_authorized(ns, True)
                    raise OSError("synthetic flush failure after confirmed CMGF and authorization change")

                modem.flush = flush
                result = send_cloud_serial_command_namespace_runtime(
                    ns, "CLOUD:SEND_SMS", {"sms_phone": "10086", "sms_message": "synthetic-only"},
                )
                self.assertFalse(result[0])
                self.assertEqual(modem.writes, [b"AT+CMGF=0\r\n"])
                self.assertFalse(serial_sender.serial_sms_recovery_error(modem))
                self.assertIsNone(at._active)
                self.assertIsNone(sms._active)

    def test_confirmed_cmgf_flush_failure_cannot_move_manual_send_to_new_connection(self):
        for change in ("generation", "connection"):
            with self.subTest(change=change):
                ns, modem, at, sms = self.fixture()

                def flush():
                    self.assertTrue(at._active.wait(0).line)
                    if change == "generation":
                        ns["serial_connection_generation"] += 1
                    else:
                        ns["serial_obj"] = recovery.StatefulModem()
                    raise OSError("synthetic flush failure after confirmed CMGF and connection change")

                modem.flush = flush
                outcomes = []
                worker = self.start_send("manual", ns, modem, sms, outcomes,
                                         command_response_coordinator=at)
                worker.join(2)
                self.assertFalse(worker.is_alive())
                self.assertEqual(outcomes, [])
                self.assertEqual(modem.writes, [b"AT+CMGF=0\r\n"])
                if change == "connection":
                    self.assertEqual(ns["serial_obj"].writes, [])
                self.assertFalse(serial_sender.serial_sms_recovery_error(modem))
                self.assertIsNone(at._active)
                self.assertIsNone(sms._active)

    def test_unknown_pdu_write_submission_does_not_send_escape(self):
        for path in ("direct", "locked"):
            ns, modem, at, sms = self.fixture()
            original_write = modem.write

            def write(payload):
                if payload.endswith(b"\x1a"):
                    modem.writes.append(payload)
                    raise OSError("synthetic write exception with unknown submitted length")
                return original_write(payload)

            modem.write = write
            outcomes = []
            with patch.object(serial_sender, "DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR", at):
                worker = self.start_send(path, ns, modem, sms, outcomes, abort_timeout=0.02)
                worker.join(1)
            self.assertFalse(worker.is_alive())
            self.assertEqual(outcomes, [False])
            self.assertNotIn(b"\x1b", modem.writes)
            self.assertTrue(serial_sender.serial_sms_recovery_error(modem))

    def test_submitted_pdu_flush_failure_only_drains_original_cmgs(self):
        for path in ("direct", "locked"):
            for terminal in ("confirmed", "missing"):
                with self.subTest(path=path, terminal=terminal):
                    ns, modem, at, sms = self.fixture()
                    original_write = modem.write
                    pdu_written = threading.Event()
                    flush_failed = threading.Event()

                    def write(payload):
                        if payload.endswith(b"\x1a"):
                            modem.writes.append(payload)
                            modem.phase = "command"
                            pdu_written.set()
                            return len(payload)
                        return original_write(payload)

                    def flush():
                        if pdu_written.is_set():
                            flush_failed.set()
                            raise OSError("synthetic flush failure after complete PDU write")

                    modem.write, modem.flush = write, flush
                    modem.escape_reply = None
                    outcomes = []
                    with patch.object(serial_sender, "DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR", at):
                        worker = self.start_send(path, ns, modem, sms, outcomes)
                        try:
                            self.assertTrue(flush_failed.wait(1))
                            self.assertFalse(modem.escape_written.wait(0.03), "ESC after Ctrl-Z submission")
                            modem.observe("OK")
                            self.assertIsNotNone(sms._active)
                            self.assertFalse(sms._active.terminal_received())
                            if terminal == "confirmed":
                                modem.observe("+CMGS: 7")
                                modem.observe("OK")
                        finally:
                            worker.join(2)
                    self.assertFalse(worker.is_alive())
                    self.assertEqual(outcomes, [False])
                    self.assertEqual(bool(serial_sender.serial_sms_recovery_error(modem)),
                                     terminal == "missing")
                    self.assertNotIn(b"\x1b", modem.writes)


if __name__ == "__main__":
    unittest.main(verbosity=2)
