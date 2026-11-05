import threading
import unittest
from unittest.mock import patch

from sms_core import call_send_namespace_runtime, serial_sender
from sms_core.call_effects import apply_call_decision
from sms_core.call_events import CallState, handle_call_line
from sms_core.call_command_context import observe_call_command_line
from tests.test_local_outgoing_calls import namespace_fixture


class AutoCallRejectionTests(unittest.TestCase):
    def setUp(self):
        self.namespace = namespace_fixture()
        self.state = CallState()
        self.workers, self.notices, self.pushes, self.events = [], [], [], []

    def start(self, name, target, **kwargs):
        self.workers.append((name, target, kwargs))

    def build(self, **kwargs):
        return call_send_namespace_runtime.build_auto_call_hangup_callback(
            self.namespace, start_worker=self.start, **kwargs,
        )

    def line(self, text, now=10.0):
        decision = handle_call_line(text, self.state, now, 'Blacklist', [], ['10086'], False)
        self.state = decision.state
        return decision

    def apply(self, callback, decision):
        return apply_call_decision(
            decision, 'TEST', callback,
            lambda message, **kwargs: self.pushes.append((message, kwargs)),
            lambda message, level: self.notices.append((message, level)),
            lambda *_args: None, lambda *_args, **_kwargs: None, lambda: None,
            send_cloud_call_event=lambda *args, **kwargs: self.events.append((args, kwargs)),
        )

    def test_pending_callback_does_not_claim_rejection_before_completion(self):
        completions = []

        def pending_callback(*, call_session_id, on_result):
            completions.append(on_result)
            return None

        pending_callback.observe_decision = lambda _decision: None
        self.apply(pending_callback, self.line('+CLIP: "10086",129'))
        self.assertEqual((self.pushes, self.events, self.notices), ([], [], []))
        completions[0](serial_sender.SerialCommandResult(False, 'synthetic failure'))
        self.assertEqual((self.pushes, self.events), ([], []))
        self.assertEqual(len(self.notices), 1)
        self.assertIn('未完成', self.notices[0][0])

    def test_blocked_clip_duplicates_keep_independent_call_id_until_terminal(self):
        first = self.line('+CLIP: "10086",129', 10)
        repeated = self.line('+CLIP: "10086",129', 20)
        self.assertEqual(repeated.call_session_id, first.call_session_id)
        ended = self.line('NO CARRIER', 21)
        self.assertEqual(ended.state.call_session_id, '')
        self.assertFalse(ended.call_ended)
        new_call = self.line('+CLIP: "10086",129', 22)
        self.assertNotEqual(new_call.call_session_id, first.call_session_id)

    def test_single_blocked_clip_waits_for_worker_and_confirms_real_ath(self):
        callback = self.build()
        self.apply(callback, self.line('+CLIP: "10086",129'))
        self.assertEqual(self.namespace['serial_obj'].writes, [])
        self.assertEqual((self.pushes, self.events, self.notices), ([], [], []))
        self.assertEqual(len(self.workers), 1)
        self.assertIs(self.workers[0][2]['thread_registry'],
                      self.namespace['SERIAL_COMMAND_THREAD_REGISTRY'])
        self.workers[0][1]()
        self.assertEqual(self.namespace['serial_obj'].writes, [b'ATH\r\n'])
        self.assertEqual(len(self.events), 1)
        self.assertTrue(self.events[0][1]['blocked'])
        self.assertEqual(self.events[0][1]['call_session_id'], self.state.call_session_id)
        self.assertEqual(len(self.pushes), 1)
        self.assertIsNone(self.namespace['SERIAL_COMMAND_RESPONSE_COORDINATOR']._active)

    def test_duplicate_blocked_clips_do_not_enqueue_or_notify_twice(self):
        callback = self.build()
        for now in (10, 12, 25):
            self.apply(callback, self.line('+CLIP: "10086",129', now))
        self.assertEqual(len(self.workers), 1)
        self.workers[0][1]()
        self.apply(callback, self.line('+CLIP: "10086",129', 26))
        self.assertEqual(self.namespace['serial_obj'].writes, [b'ATH\r\n'])
        self.assertEqual((len(self.events), len(self.pushes)), (1, 1))

    def test_terminal_next_call_and_changed_connection_cancel_queued_rejection(self):
        for change in ('terminal', 'terminal_epoch', 'next_clip', 'serial', 'generation', 'shutdown'):
            with self.subTest(change=change):
                self.setUp()
                callback = self.build()
                modem = self.namespace['serial_obj']
                self.apply(callback, self.line('+CLIP: "10086",129'))
                if change == 'terminal':
                    self.apply(callback, self.line('NO CARRIER', 11))
                elif change == 'terminal_epoch':
                    observe_call_command_line(self.namespace, 'NO CARRIER')
                elif change == 'next_clip':
                    self.apply(callback, self.line('+CLIP: "10010",129', 11))
                elif change == 'serial':
                    self.namespace['serial_obj'] = namespace_fixture()['serial_obj']
                elif change == 'generation':
                    self.namespace['serial_connection_generation'] += 1
                else:
                    self.namespace['TK_SHUTDOWN'] = threading.Event()
                    self.namespace['TK_SHUTDOWN'].set()
                self.workers[0][1]()
                self.assertEqual(modem.writes, [])
                self.assertFalse(any(kwargs.get('blocked') for _args, kwargs in self.events))
                self.assertFalse(any('拒接' in message for message, _level in self.notices))
                self.assertIsNone(self.namespace['SERIAL_COMMAND_RESPONSE_COORDINATOR']._active)

    def test_new_call_after_modem_response_cancels_queued_ui_notification(self):
        ui = []
        self.namespace['run_on_ui_thread'] = lambda callback, _post: ui.append(callback)
        self.namespace['ui_post'] = lambda *_args: None
        callback = self.build()
        self.apply(callback, self.line('+CLIP: "10086",129'))
        self.workers[0][1]()
        self.assertEqual(len(ui), 1)
        self.apply(callback, self.line('+CLIP: "10010",129', 11))
        ui[0]()
        self.assertEqual(self.namespace['serial_obj'].writes, [b'ATH\r\n'])
        self.assertFalse(any(kwargs.get('blocked') for _args, kwargs in self.events))
        self.assertFalse(any('拒接' in message for message, _level in self.notices))

    def test_failed_modem_result_does_not_report_blocked_success(self):
        callback = self.build()
        self.namespace['serial_obj'].response = 'ERROR'
        self.apply(callback, self.line('+CLIP: "10086",129'))
        self.workers[0][1]()
        self.assertEqual((self.pushes, self.events), ([], []))
        self.assertEqual(len(self.notices), 1)
        self.assertIn('未完成', self.notices[0][0])

    def test_initialized_rejection_uses_confirmed_worker_and_legacy_callback_stays_synchronous(self):
        callback = self.build()
        self.apply(callback, self.line('+CLIP: "10086",129'))
        self.assertEqual(self.namespace['serial_obj'].writes, [])
        self.assertEqual((self.pushes, self.events), ([], []))
        self.assertEqual(len(self.workers), 1)
        self.workers[0][1]()
        self.assertEqual(self.namespace['serial_obj'].writes, [b'ATH\r\n'])
        self.assertEqual(len(self.events), 1)
        legacy_calls = []
        self.apply(lambda: legacy_calls.append('ATH'), self.line('+CLIP: "10086",129', 11))
        self.assertEqual(legacy_calls, ['ATH'])

    def test_initialized_rejection_without_modem_ack_does_not_report_success(self):
        modem = self.namespace['serial_obj']
        modem.response = None
        callback = self.build(response_timeout=0.001)
        self.apply(callback, self.line('+CLIP: "10086",129'))
        for _name, worker, _kwargs in self.workers:
            worker()
        self.assertEqual((self.pushes, self.events), ([], []))
        self.assertEqual(modem.writes, [b'ATH\r\n'])
        self.assertTrue(serial_sender.serial_sms_recovery_error(modem))
        modem.response = 'OK'
        next_result = serial_sender.write_serial_command_confirmed_locked(
            self.namespace['serial_lock'], lambda: modem, 'AT',
            response_coordinator=self.namespace['SERIAL_COMMAND_RESPONSE_COORDINATOR'],
            response_timeout=0.001,
        )
        self.assertFalse(next_result.ok)
        self.assertEqual(modem.writes, [b'ATH\r\n'])
        self.assertEqual(len(self.notices), 1)
        self.assertIn('未完成', self.notices[0][0])
        self.assertIsNone(self.namespace['SERIAL_COMMAND_RESPONSE_COORDINATOR']._active)

    def test_initialized_rejection_with_modem_error_does_not_report_success(self):
        modem = self.namespace['serial_obj']
        modem.response = 'ERROR'
        callback = self.build(response_timeout=0.001)
        self.apply(callback, self.line('+CLIP: "10086",129'))
        for _name, worker, _kwargs in self.workers:
            worker()
        self.assertEqual((self.pushes, self.events), ([], []))
        self.assertEqual(modem.writes, [b'ATH\r\n'])
        self.assertFalse(serial_sender.serial_sms_recovery_error(modem))
        self.assertEqual(len(self.notices), 1)
        self.assertIn('未完成', self.notices[0][0])
        self.assertIsNone(self.namespace['SERIAL_COMMAND_RESPONSE_COORDINATOR']._active)

    def test_initialized_rejection_rechecks_new_call_at_final_write(self):
        for change in ('next_clip', 'terminal', 'terminal_epoch', 'serial', 'generation', 'shutdown'):
            with self.subTest(change=change):
                self.setUp()
                callback = self.build(response_timeout=0.001)
                modem = self.namespace['serial_obj']
                coordinator = self.namespace['SERIAL_COMMAND_RESPONSE_COORDINATOR']
                original_begin = coordinator.begin

                def changed_context_after_waiter(**kwargs):
                    waiter = original_begin(**kwargs)
                    if change == 'next_clip':
                        self.apply(callback, self.line('+CLIP: "10010",129', 11))
                    elif change == 'terminal':
                        self.apply(callback, self.line('NO CARRIER', 11))
                    elif change == 'terminal_epoch':
                        observe_call_command_line(self.namespace, 'NO CARRIER')
                    elif change == 'serial':
                        self.namespace['serial_obj'] = namespace_fixture()['serial_obj']
                    elif change == 'generation':
                        self.namespace['serial_connection_generation'] += 1
                    else:
                        self.namespace['TK_SHUTDOWN'] = threading.Event()
                        self.namespace['TK_SHUTDOWN'].set()
                    return waiter

                with patch.object(coordinator, 'begin', side_effect=changed_context_after_waiter):
                    self.apply(callback, self.line('+CLIP: "10086",129'))
                    self.workers[0][1]()
                self.assertEqual(modem.writes, [])
                self.assertFalse(any(kwargs.get('blocked') for _args, kwargs in self.events))
                self.assertIsNone(coordinator._active)

    def test_worker_start_failure_is_explicit_without_success_notification(self):
        def reject_start(*_args, **_kwargs):
            raise RuntimeError('synthetic worker start failure')

        callback = call_send_namespace_runtime.build_auto_call_hangup_callback(
            self.namespace, start_worker=reject_start,
        )
        self.apply(callback, self.line('+CLIP: "10086",129'))
        self.assertEqual((self.pushes, self.events), ([], []))
        self.assertEqual(len(self.notices), 1)
        self.assertIn('未完成', self.notices[0][0])

    def test_start_failure_releases_job_for_later_same_session_clip(self):
        attempts = []

        def start(name, target, **kwargs):
            attempts.append(name)
            if len(attempts) == 1:
                raise RuntimeError('synthetic start rejection')
            self.workers.append((name, target, kwargs))

        callback = call_send_namespace_runtime.build_auto_call_hangup_callback(
            self.namespace, start_worker=start,
        )
        self.apply(callback, self.line('+CLIP: "10086",129'))
        self.apply(callback, self.line('+CLIP: "10086",129', 11))
        self.assertEqual(len(self.workers), 1)
        self.workers[0][1]()
        self.assertEqual(self.namespace['serial_obj'].writes, [b'ATH\r\n'])
        self.assertEqual(len(self.events), 1)
        self.assertTrue(self.events[0][1]['blocked'])

    def test_start_failure_does_not_release_a_reentrant_new_call_job(self):
        attempts = []

        def start(name, target, **kwargs):
            attempts.append(name)
            if len(attempts) == 1:
                self.apply(callback, self.line('NO CARRIER', 11))
                self.apply(callback, self.line('+CLIP: "10086",129', 12))
                raise RuntimeError('synthetic old start rejection')
            self.workers.append((name, target, kwargs))

        callback = call_send_namespace_runtime.build_auto_call_hangup_callback(
            self.namespace, start_worker=start,
        )
        self.apply(callback, self.line('+CLIP: "10086",129'))
        session_id = self.state.call_session_id
        self.apply(callback, self.line('+CLIP: "10086",129', 13))
        self.assertEqual(len(self.workers), 1)
        self.workers[0][1]()
        self.assertEqual(self.namespace['serial_obj'].writes, [b'ATH\r\n'])
        self.assertEqual([kwargs['call_session_id'] for _args, kwargs in self.events], [session_id])

    def test_confirmed_rejection_survives_original_terminal_before_ui_notice(self):
        for terminal_timing in ('before_post', 'after_post'):
            with self.subTest(terminal_timing=terminal_timing):
                self.setUp()
                ui = []
                self.namespace['run_on_ui_thread'] = lambda callback, _post: ui.append(callback)
                self.namespace['ui_post'] = lambda *_args: None
                callback = self.build()
                self.apply(callback, self.line('+CLIP: "10086",129'))
                original_session_id = self.state.call_session_id
                if terminal_timing == 'before_post':
                    self.namespace['serial_obj'].after_write = lambda: self.apply(
                        callback, self.line('NO CARRIER', 11))
                self.workers[0][1]()
                if terminal_timing == 'after_post':
                    self.apply(callback, self.line('NO CARRIER', 11))
                self.assertEqual(len(ui), 1)
                ui[0]()
                self.assertEqual(self.namespace['serial_obj'].writes, [b'ATH\r\n'])
                self.assertEqual(len(self.events), 1)
                self.assertTrue(self.events[0][1]['blocked'])
                self.assertEqual(self.events[0][1]['call_session_id'], original_session_id)
                self.assertEqual(len(self.pushes), 1)
                self.assertEqual(len(self.notices), 1)

    def test_registered_worker_waits_for_original_transaction_lock_and_releases_registry(self):
        reached = threading.Event()
        threads = []

        def start(name, target, **kwargs):
            def run():
                reached.set()
                target()

            thread = serial_sender.start_registered_serial_worker(name, run, **kwargs)
            threads.append(thread)
            return thread

        callback = call_send_namespace_runtime.build_auto_call_hangup_callback(
            self.namespace, start_worker=start,
        )
        try:
            with serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK:
                self.apply(callback, self.line('+CLIP: "10086",129'))
                self.assertTrue(reached.wait(2))
                self.assertEqual(self.namespace['serial_obj'].writes, [])
            threads[0].join(3)
            self.assertFalse(threads[0].is_alive())
            self.assertEqual(self.namespace['serial_obj'].writes, [b'ATH\r\n'])
            self.assertEqual(self.namespace['SERIAL_COMMAND_THREAD_REGISTRY'].snapshot(), ())
        finally:
            for thread in threads:
                thread.join(3)

    def test_new_ring_after_original_terminal_cancels_confirmed_rejection_notice(self):
        ui = []
        self.namespace['run_on_ui_thread'] = lambda callback, _post: ui.append(callback)
        self.namespace['ui_post'] = lambda *_args: None
        callback = self.build()
        self.apply(callback, self.line('+CLIP: "10086",129'))
        self.workers[0][1]()
        self.apply(callback, self.line('NO CARRIER', 11))
        self.apply(callback, self.line('RING', 12))
        self.assertEqual(len(ui), 1)
        ui[0]()
        self.assertEqual(self.namespace['serial_obj'].writes, [b'ATH\r\n'])
        self.assertEqual((self.pushes, self.events, self.notices), ([], [], []))


if __name__ == '__main__':
    unittest.main()
