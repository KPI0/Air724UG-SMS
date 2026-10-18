from dataclasses import replace
from functools import partial
import threading
import unittest

from sms_app.serial_namespace_bindings import install_serial_namespace_bindings
from sms_core.serial_app_runtime import run_serial_app_from_wiring, run_serial_app_runtime
from sms_core.serial_runtime import run_serial_runtime_thread
from tests import test_call_popup_command_isolation as popup_fixture
from tests.test_serial_runtime import parse_head, runtime_callbacks, runtime_config


class SerialCallStateMergeTests(unittest.TestCase):
    def setUp(self):
        self.fixture = popup_fixture.CallPopupCommandIsolationTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.tearDown)
        self.namespace = self.fixture.namespace
        self.namespace['current_dial_num'] = ''
        install_serial_namespace_bindings(self.namespace)
        self.now = 100.0
        self.popup = None
        self.interrupt = lambda _line: None
        self.after_show = lambda: None

    def run_lines(self, scenario):
        ns = self.namespace

        def show(_number, call_session_id=''):
            self.popup = self.fixture.start_popup(call_session_id)
            self.after_show()

        callbacks = replace(runtime_callbacks([]),
            start_incoming_call=ns['start_incoming_call_session'],
            start_serial_incoming_call=lambda *args: ns['start_serial_incoming_call'](*args),
            call_state_guard=ns['serial_call_state_guard'],
            finish_incoming_call=ns['finish_incoming_call_session'],
            reset_incoming_call=ns['reset_incoming_call_session'],
            show_call_popup=show, close_call_popup=ns['close_call_popup'],
            consume_remote_incoming_end=lambda session_id: ns['consume_remote_incoming_end'](session_id),
            push_serial_debug=lambda line: self.interrupt(line),
        ).__dict__
        callbacks = dict(callbacks, schedule_connected_log=lambda *_: None, serial_error_ui=lambda *_: None)

        def loop(**kwargs):
            handle = kwargs['handle_line']
            handle('RING')
            handle('+CLIP: "10086",129')
            self.assertEqual(ns['ring_timeout_target'], 112.0)
            scenario(handle)

        return run_serial_app_from_wiring(
            parse_callback_head=parse_head,
            values={key: (lambda value=value: value) for key, value in runtime_config().__dict__.items()},
            callbacks=callbacks,
            state_access={
                'get_call_state': ns['get_serial_call_state'],
                'set_call_state': ns['set_serial_call_state'],
                'merge_call_state': ns['merge_serial_call_state'],
                'popup_active': lambda: ns['current_call_popup'] is not None,
                'ignore_repeat_state': {}, 'serial_running': lambda: True,
                'port': lambda: 'TEST', 'baud': lambda: 115200, 'set_log_prefix': lambda *_: None,
            },
            io_callbacks={
                'resolve_target_port': lambda: 'TEST', 'open_and_initialize_serial': lambda *_: None,
                'read_serial_line': lambda: b'', 'safe_close_serial': lambda: None,
            },
            reconnect_callbacks={
                'stop_requested': lambda: False, 'interval': lambda: 1,
                'wakeup_wait': lambda **_: None, 'wakeup_clear': lambda: None,
                'try_manual_rebind_after_error': lambda *_: False,
            },
            apply_disconnect_effects=lambda *_: None,
            run_app=partial(run_serial_app_runtime, run_runtime_thread=partial(
                run_serial_runtime_thread, run_loop=loop, clock=lambda: self.now)),
        )

    def answer_during(self, line):
        def interrupt(current_line):
            if current_line == line:
                self.interrupt = lambda _: None
                self.fixture.click(self.popup, 'answer')
                self.fixture.settle()
                self.assertEqual(self.namespace['ring_timeout_target'], -1.0)
        self.interrupt = interrupt

    def test_answer_during_serial_line_survives_writeback_and_idle_timeout(self):
        def scenario(handle):
            self.answer_during('OK')
            self.now = 101.0
            handle('OK')
            self.assertEqual(self.namespace['ring_timeout_target'], -1.0)
            self.now = 114.0
            handle('')
            self.assertTrue(self.popup.exists)
            self.assertIsNotNone(self.fixture.tracker.snapshot().session_token)
        state = self.run_lines(scenario)
        self.assertEqual(state.call_state.ring_timeout_target, -1.0)
        self.assertEqual(self.fixture.serial.writes, [b'ATA\r\n'])

    def test_modem_terminal_still_wins_when_answer_finishes_in_same_line(self):
        def scenario(handle):
            self.answer_during('NO CARRIER')
            self.now = 101.0
            handle('NO CARRIER')
            self.assertEqual(self.namespace['ring_timeout_target'], 0.0)
            self.assertFalse(self.popup.exists)
            self.assertIsNone(self.fixture.tracker.snapshot().session_token)
        state = self.run_lines(scenario)
        self.assertEqual(state.call_state.ring_timeout_target, 0.0)
        self.assertEqual(state.call_state.call_session_id, '')

    def test_unrelated_serial_line_keeps_concurrent_dial_number(self):
        def scenario(handle):
            self.interrupt = lambda line: self.namespace.__setitem__('current_dial_num', '10010') if line == 'OK' else None
            handle('OK')
            self.assertEqual(self.namespace['current_dial_num'], '10010')
        state = self.run_lines(scenario)
        self.assertEqual(state.call_state.current_dial_num, '10010')

    def test_normal_ring_refresh_and_terminal_still_commit(self):
        def scenario(handle):
            self.now = 110.0
            handle('RING')
            self.assertEqual(self.namespace['ring_timeout_target'], 122.0)
            handle('NO CARRIER')
            self.assertEqual(self.namespace['ring_timeout_target'], 0.0)
            self.assertFalse(self.popup.exists)
        self.run_lines(scenario)

    def test_new_call_commit_preserves_answer_received_before_line_returns(self):
        def scenario(handle):
            handle('NO CARRIER')
            self.now = 101.0

            def answer_new():
                self.after_show = lambda: None
                self.fixture.click(self.popup, 'answer')
                self.fixture.settle()

            self.after_show = answer_new
            handle('+CLIP: "10086",129')
            self.assertEqual(self.namespace['ring_timeout_target'], -1.0)
            self.now = 120.0
            handle('')
            self.assertTrue(self.popup.exists)
        state = self.run_lines(scenario)
        self.assertEqual(state.call_state.ring_timeout_target, -1.0)

    def test_new_call_commit_preserves_terminal_when_timer_returns_to_old_value(self):
        def scenario(handle):
            ns = self.namespace
            ns['show_missed_call_popup'] = lambda _missed: None
            handle('NO CARRIER')
            self.assertEqual(ns['ring_timeout_target'], 0.0)
            self.now = 101.0

            def finish_new():
                self.after_show = lambda: None
                local_id = ns['call_popup_active_session_id']
                ns['register_peer_incoming_call']('peer-new', '10086', local_id)
                ns['finish_remote_incoming_call'](
                    'ended', session_id='peer-new', caller_num='10086')

            self.after_show = finish_new
            handle('+CLIP: "10086",129')
            self.assertIsNone(self.fixture.tracker.snapshot().session_token)
            self.assertEqual(ns['ring_timeout_target'], 0.0)
            self.assertFalse(self.popup.exists)
        state = self.run_lines(scenario)
        self.assertEqual(state.call_state.ring_timeout_target, 0.0)

    def test_peer_marker_consumption_and_new_lifetime_share_one_state_lock(self):
        def scenario(handle):
            ns = self.namespace
            ns['show_missed_call_popup'] = lambda _missed: None
            old_id = ns['call_popup_active_session_id']
            ns['register_peer_incoming_call']('peer-old', '10086', old_id)
            attempt = threading.Event()
            observed = {}
            worker = None

            class ObservedLock:
                def __init__(self):
                    self.lock = threading.RLock()

                def __enter__(self):
                    if threading.current_thread() is worker:
                        acquired = self.lock.acquire(blocking=False)
                        observed['blocked'] = not acquired
                        attempt.set()
                        if not acquired:
                            self.lock.acquire()
                    else:
                        self.lock.acquire()
                    return self

                def __exit__(self, *_args):
                    self.lock.release()

            ns['_call_peer_state_lock'] = ObservedLock()
            consume = ns['consume_remote_incoming_end']
            start = ns['start_serial_incoming_call']

            def observe_after_lock():
                with ns['serial_call_state_guard']():
                    observed['token'] = self.fixture.tracker.snapshot().session_token

            def consume_and_compete(session_id):
                nonlocal worker
                consumed = consume(session_id)
                if consumed:
                    worker = threading.Thread(target=observe_after_lock, daemon=True)
                    worker.start()
                    self.assertTrue(attempt.wait(3))
                    self.assertTrue(observed['blocked'])
                return consumed

            def start_and_verify(*args):
                result = start(*args)
                if worker is not None:
                    self.assertNotIn('token', observed)
                return result

            def finish_old(line):
                if line.startswith('+CLIP:'):
                    self.interrupt = lambda _: None
                    ns['finish_remote_incoming_call'](
                        'ended', session_id='peer-old', caller_num='10086')

            ns['consume_remote_incoming_end'] = consume_and_compete
            ns['start_serial_incoming_call'] = start_and_verify
            self.interrupt = finish_old
            self.now = 101.0
            try:
                handle('+CLIP: "10086",129')
            finally:
                if worker is not None:
                    worker.join(3)
                    self.assertFalse(worker.is_alive(), 'peer state lock deadlocked')
            self.assertIsNotNone(observed.get('token'))
            self.assertIs(observed['token'], self.fixture.tracker.snapshot().session_token)

        self.run_lines(scenario)

    def test_merge_does_not_change_concurrent_outgoing_state(self):
        ns = self.namespace
        ns['ring_timeout_target'] = -1.0
        ns['current_dial_num'] = '10010'
        merged = ns['merge_serial_call_state']((112.0, ''), 0.0, '')
        self.assertEqual(merged, (-1.0, '10010'))

    def test_peer_terminal_during_redial_clip_does_not_swallow_new_call(self):
        def scenario(handle):
            ns = self.namespace
            ns['show_missed_call_popup'] = lambda _missed: None
            old_id = ns['call_popup_active_session_id']
            old_token = self.fixture.tracker.snapshot().session_token
            old_popup = self.popup
            ns['register_peer_incoming_call']('peer-old', '10086', old_id)
            self.now = 100.5
            handle('RING')

            def finish_old(line):
                if line.startswith('+CLIP:'):
                    self.interrupt = lambda _: None
                    ns['finish_remote_incoming_call'](
                        'ended', session_id='peer-old', caller_num='10086')

            self.interrupt = finish_old
            self.now = 101.0
            handle('+CLIP: "10086",129')
            self.assertIsNotNone(self.fixture.tracker.snapshot().session_token)
            self.assertIsNot(self.fixture.tracker.snapshot().session_token, old_token)
            self.assertIsNot(self.popup, old_popup)
            self.assertTrue(self.popup.exists)
            self.assertNotEqual(ns['call_popup_active_session_id'], old_id)
            self.assertEqual(ns['ring_timeout_target'], 113.0)
            handle('')
            self.assertTrue(self.popup.exists)
            self.assertEqual(ns['ring_timeout_target'], 113.0)

        self.run_lines(scenario)


if __name__ == '__main__':
    unittest.main()
