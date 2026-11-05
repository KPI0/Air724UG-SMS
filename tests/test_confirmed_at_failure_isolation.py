import threading
import unittest

from sms_core import serial_sender
from tests.test_serial_transaction_boundaries import SyntheticModem


class OpaqueError(Exception):
    def __str__(self):
        raise AssertionError('Diagnostic must not format original error')

    __repr__ = __str__


class ConfirmedAtFailureIsolationTests(unittest.TestCase):
    def send_sms(self, modem, at, sms, *, locked, push_debug):
        options = dict(push_debug=push_debug, response_coordinator=sms,
            command_response_coordinator=at, segment_timeout=0, prompt_timeout=0, abort_timeout=0)
        if locked:
            return serial_sender.write_text_sms_pdu_locked(threading.RLock(), lambda: modem,
                '10086', 'synthetic body', **options)
        return serial_sender.write_text_sms_pdu(modem, '10086', 'synthetic body', **options)

    def test_cmgs_diagnostic_failure_still_escapes_original_prompt(self):
        for locked in (False, True):
            with self.subTest(locked=locked):
                at, sms = serial_sender.AtCommandResponseCoordinator(), serial_sender.SmsPduSendCoordinator()
                modem = SyntheticModem(at, sms)
                def fail(message):
                    if '发送: AT+CMGS=' in message:
                        raise OpaqueError()
                self.assertFalse(self.send_sms(modem, at, sms, locked=locked, push_debug=fail))
                self.assertEqual(len(modem.writes), 3)
                self.assertEqual(modem.writes[-1], b'\x1b')
                self.assertFalse(serial_sender.serial_sms_recovery_error(modem))
                self.assertIsNone(sms._active)
                self.assertTrue(serial_sender.write_serial_command_confirmed_locked(threading.RLock(),
                    lambda: modem, 'ATI', response_coordinator=at, response_timeout=0).ok)

    def test_pdu_diagnostic_failure_never_escapes_committed_sms(self):
        for locked in (False, True):
            for terminal in (False, True):
                with self.subTest(locked=locked, terminal=terminal):
                    at, sms = serial_sender.AtCommandResponseCoordinator(), serial_sender.SmsPduSendCoordinator()
                    class Modem(SyntheticModem):
                        def write(self, payload):
                            if payload.endswith(b'\x1a') and not terminal:
                                self.writes.append(payload)
                                return len(payload)
                            return super().write(payload)
                    modem = Modem(at, sms)
                    def fail(message):
                        if '发送 PDU 正文' in message:
                            raise OpaqueError()
                    self.assertFalse(self.send_sms(modem, at, sms, locked=locked, push_debug=fail))
                    self.assertEqual(len(modem.writes), 3)
                    self.assertTrue(modem.writes[-1].endswith(b'\x1a'))
                    self.assertNotIn(b'\x1b', modem.writes)
                    self.assertEqual(bool(serial_sender.serial_sms_recovery_error(modem)), not terminal)
                    self.assertIsNone(sms._active)
                    self.assertEqual(serial_sender.write_serial_command_confirmed_locked(threading.RLock(),
                        lambda: modem, 'ATI', response_coordinator=at, response_timeout=0).ok, terminal)

    def test_debug_failure_after_write_preserves_terminal_or_isolates_original(self):
        for terminal in ('', 'OK', 'ERROR'):
            with self.subTest(terminal=terminal):
                at = serial_sender.AtCommandResponseCoordinator()
                modem = SyntheticModem(at, serial_sender.SmsPduSendCoordinator(), command_reply=terminal)
                def fail(_message):
                    raise OpaqueError()
                result = serial_sender.write_serial_command_confirmed_locked(threading.RLock(), lambda: modem,
                    'ATA', response_coordinator=at, response_timeout=0, isolate_unconfirmed=True, push_debug=fail)
                self.assertEqual(result.ok, terminal == 'OK')
                self.assertEqual(bool(serial_sender.serial_sms_recovery_error(modem)), terminal == '')
                self.assertIsNone(at._active)
                modem.command_reply = 'OK'
                followup = serial_sender.write_serial_command_confirmed_locked(threading.RLock(), lambda: modem,
                    'ATI', response_coordinator=at, response_timeout=0, isolate_unconfirmed=True)
                self.assertEqual(followup.ok, terminal != '')
                self.assertEqual(len(modem.writes), 1 if terminal == '' else 2)

    def test_response_callback_runs_after_original_unconfirmed_isolation(self):
        for terminal in ('', 'OK'):
            with self.subTest(terminal=terminal):
                at = serial_sender.AtCommandResponseCoordinator()
                modem = SyntheticModem(at, serial_sender.SmsPduSendCoordinator(), command_reply=terminal)
                replacement = SyntheticModem(at, serial_sender.SmsPduSendCoordinator())
                current, observed = [modem], []
                def fail(_response):
                    observed.append(bool(serial_sender.serial_sms_recovery_error(modem)))
                    current[0] = replacement
                    raise OpaqueError()
                with self.assertRaises(OpaqueError):
                    serial_sender.write_serial_command_confirmed_locked(threading.RLock(), lambda: current[0],
                        'ATA', response_coordinator=at, response_timeout=0, isolate_unconfirmed=True,
                        on_modem_response=fail)
                self.assertEqual(observed, [terminal == ''])
                self.assertEqual(bool(serial_sender.serial_sms_recovery_error(modem)), terminal == '')
                self.assertFalse(serial_sender.serial_sms_recovery_error(replacement))
                self.assertIsNone(at._active)

    def test_unconfirmed_cmgf_debug_failure_cannot_submit_cmgs(self):
        at, sms = serial_sender.AtCommandResponseCoordinator(), serial_sender.SmsPduSendCoordinator()
        modem = SyntheticModem(at, sms, command_reply='')
        def fail(_message):
            raise OpaqueError()
        result = serial_sender.write_text_sms_pdu_locked(threading.RLock(), lambda: modem, '10086',
            'synthetic body', push_debug=fail, response_coordinator=sms, command_response_coordinator=at,
            segment_timeout=0, prompt_timeout=0, abort_timeout=0)
        self.assertFalse(result)
        self.assertEqual(modem.writes, [b'AT+CMGF=0\r\n'])
        self.assertTrue(serial_sender.serial_sms_recovery_error(modem))
        self.assertIsNone(at._active)
