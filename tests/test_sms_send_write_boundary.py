"""SMS prompt/error collection starts at the physical CMGS write boundary."""
import threading
import unittest
from unittest.mock import patch

from sms_core import serial_sender
from tests.test_serial_transaction_boundaries import FinalWriteGate


class SmsWriteBoundaryTests(unittest.TestCase):
    def run_case(self, locked, previous_reply, current_prompt, *, fail_cmgs=False):
        at = serial_sender.AtCommandResponseCoordinator()
        sms = serial_sender.SmsPduSendCoordinator()
        writes = []

        class Modem:
            is_open = True

            def write(self, payload):
                if fail_cmgs and payload.startswith(b'AT+CMGS='):
                    raise OSError('synthetic write failure')
                writes.append(payload)
                if payload == b'AT+CMGF=0\r\n':
                    at.observe_line('OK', connection=self)
                elif payload.startswith(b'AT+CMGS=') and current_prompt:
                    sms.observe_line('>', connection=self)
                elif payload.endswith(b'\x1a'):
                    sms.observe_line('+CMGS: 7', connection=self)
                    sms.observe_line('OK', connection=self)
                return len(payload)

            def flush(self):
                pass

        modem = Modem()
        gate = FinalWriteGate(pause_write=2)
        results = []
        errors = []

        def run():
            try:
                kwargs = dict(response_coordinator=sms, prompt_timeout=0.03,
                              segment_timeout=0.03, sleep_func=lambda _seconds: None)
                if locked:
                    result = serial_sender.write_text_sms_pdu_locked(
                        threading.RLock(), lambda: modem, '10086', 'synthetic-only',
                        command_response_coordinator=at, **kwargs)
                else:
                    result = serial_sender.write_text_sms_pdu(
                        modem, '10086', 'synthetic-only', **kwargs)
                results.append(result)
            except Exception as exc:
                errors.append(type(exc).__name__)

        with patch.object(serial_sender, 'DEFAULT_SERIAL_WRITE_LOCK', gate):
            worker = threading.Thread(target=run)
            worker.start()
            try:
                self.assertTrue(gate.reached.wait(2))
                self.assertEqual(writes, [b'AT+CMGF=0\r\n'])
                if previous_reply == 'cancel':
                    sms.cancel_active('synthetic cancellation')
                elif previous_reply:
                    sms.observe_line(previous_reply, connection=modem)
            finally:
                gate.resume.set()
                worker.join(2)
        self.assertFalse(worker.is_alive())
        self.assertEqual(errors, [])
        self.assertEqual(len(results), 1)
        self.assertIsNone(sms._active)
        return results[0], writes

    def test_no_current_prompt_does_not_send_body_positive_control(self):
        for locked in (False, True):
            with self.subTest(locked=locked):
                result, writes = self.run_case(locked, '', False)
                self.assertFalse(result)
                self.assertFalse(any(payload.endswith(b'\x1a') for payload in writes))

    def test_current_prompt_allows_body_positive_control(self):
        for locked in (False, True):
            with self.subTest(locked=locked):
                result, writes = self.run_case(locked, '', True)
                self.assertTrue(result)
                self.assertTrue(any(payload.endswith(b'\x1a') for payload in writes))

    def test_previous_prompt_before_cmgs_cannot_allow_body(self):
        for locked in (False, True):
            with self.subTest(locked=locked):
                _result, writes = self.run_case(locked, '>', False)
                self.assertFalse(any(payload.endswith(b'\x1a') for payload in writes),
                                 'PDU sent even though current CMGS never emitted a prompt')

    def test_previous_error_before_cmgs_cannot_fail_current_sms(self):
        for locked in (False, True):
            with self.subTest(locked=locked):
                result, _writes = self.run_case(locked, '+CMS ERROR: 500', True)
                self.assertTrue(result, 'Pre-write old error aborted current valid SMS')

    def test_cancelled_waiter_does_not_submit_cmgs_or_pdu(self):
        for locked in (False, True):
            with self.subTest(locked=locked):
                result, writes = self.run_case(locked, 'cancel', True)
                self.assertFalse(result)
                self.assertEqual(writes, [b'AT+CMGF=0\r\n'])

    def test_failed_cmgs_releases_waiter_and_transaction_for_next_sms(self):
        for locked in (False, True):
            with self.subTest(locked=locked):
                result, writes = self.run_case(locked, '', True, fail_cmgs=True)
                self.assertFalse(result)
                self.assertEqual(writes, [b'AT+CMGF=0\r\n'])
                next_result, _writes = self.run_case(locked, '', True)
                self.assertTrue(next_result)

    def test_direct_waiter_and_coordinator_keep_immediate_observation(self):
        for waiter in (serial_sender.SmsPduSendWaiter(),
                       serial_sender.SmsPduSendCoordinator().begin_segment()):
            with self.subTest(waiter=type(waiter).__name__):
                waiter.observe_line('>')
                self.assertTrue(waiter.wait_prompt(0).ok)
                waiter.observe_line('+CMS ERROR: 500')
                self.assertFalse(waiter.wait(0).ok)

    def test_legacy_coordinator_full_signature_keeps_connection_metadata(self):
        calls = []
        connection = object()

        class LegacyCoordinator:
            def begin_segment(self, label='', connection=None, cmgs_len=None):
                calls.append((label, connection, cmgs_len))
                return serial_sender.SmsPduSendWaiter(label, connection, cmgs_len)

        waiter, result = serial_sender._begin_sms_segment_response(
            LegacyCoordinator(), '(1/1)', connection=connection, cmgs_len=12)
        self.assertTrue(result.ok)
        self.assertEqual(calls, [('(1/1)', connection, 12)])
        waiter.observe_line('>')
        self.assertTrue(waiter.wait_prompt(0).ok)

    def test_legacy_no_argument_coordinator_is_supported(self):
        class LegacyCoordinator:
            def begin_segment(self):
                return serial_sender.SmsPduSendWaiter()

        waiter, result = serial_sender._begin_sms_segment_response(LegacyCoordinator(), '')
        self.assertTrue(result.ok)
        waiter.observe_line('>')
        self.assertTrue(waiter.wait_prompt(0).ok)


if __name__ == '__main__':
    unittest.main(verbosity=2)
