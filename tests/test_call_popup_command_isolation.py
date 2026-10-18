import threading
import unittest
from datetime import datetime
from functools import partial
from types import SimpleNamespace
from unittest.mock import patch

from sms_core import serial_sender
from sms_core.call_session import IncomingCallSessionTracker
from sms_ui.call_popup_namespace_runtime import (
    close_call_popup_namespace_runtime,
    finish_incoming_call_session_namespace_runtime,
    reset_incoming_call_session_namespace_runtime,
    show_call_popup_namespace_runtime,
    start_incoming_call_session_namespace_runtime,
)
from sms_ui.call_popup_runtime import show_call_popup_app_runtime, show_call_popup_runtime


class ObservedLock:
    """A real lock with an event proving the sender reached its wait point."""

    def __init__(self):
        self.lock = threading.RLock()
        self.owner = threading.current_thread()
        self.waiting = threading.Event()

    def __enter__(self):
        if threading.current_thread() is not self.owner:
            self.waiting.set()
        self.lock.acquire()
        return self

    def __exit__(self, *_args):
        self.lock.release()


class FakeSerial:
    is_open = True

    def __init__(self):
        self.writes = []
        self.on_write = lambda: None

    def write(self, payload):
        self.on_write()
        self.writes.append(payload)

    def flush(self):
        pass


class CallPopupCommandIsolationTests(unittest.TestCase):
    def setUp(self):
        self.serial = FakeSerial()
        self.tracker = IncomingCallSessionTracker(now_func=lambda: datetime(2026, 1, 1))
        self.threads, self.errors, self.results, self.messages, self.ui_tasks = [], [], [], [], []
        self.ui_thread = threading.current_thread()
        self.namespace = {
            'INCOMING_CALL_SESSION': self.tracker,
            'serial_obj': self.serial, 'serial_lock': ObservedLock(),
            'serial_connection_generation': 1, 'root': None,
            'current_call_popup': None, 'current_dial_popup': None,
            'center_window': lambda *_: None,
            'port_ui': lambda *args: self.messages.append(args),
            'set_status': lambda *_: None,
            'run_on_ui_thread': self.on_ui,
            'ui_post': self.ui_tasks.append,
            'log_file_only': self.errors.append,
            'ring_timeout_target': 123.0,
            'close_call_popup': lambda: close_call_popup_namespace_runtime(self.namespace),
        }

    def tearDown(self):
        for thread in self.threads:
            thread.join(3)
            self.assertFalse(thread.is_alive(), 'sender deadlocked')

    def on_ui(self, callback, _post):
        if threading.current_thread() is self.ui_thread:
            return callback()
        self.ui_tasks.append(callback)

    def start_popup(self, session_id='local-old'):
        start_incoming_call_session_namespace_runtime(self.namespace, '10086')

        def open_popup(_parent, _caller, _center, answer, hangup, *_close):
            popup = SimpleNamespace(answer=answer, hangup=hangup, exists=True)
            popup.winfo_exists = lambda: popup.exists
            popup.destroy = lambda: setattr(popup, 'exists', False)
            return popup

        def send(*args, **kwargs):
            callback = kwargs['on_result']

            def result(value):
                self.results.append(value)
                callback(value)

            kwargs['on_result'] = result
            self.threads.append(serial_sender.send_command_with_result_async(*args, **kwargs))

        return show_call_popup_namespace_runtime(
            self.namespace, '10086', session_id,
            show_app_runtime=partial(show_call_popup_app_runtime,
                show_runtime=partial(show_call_popup_runtime, open_popup=open_popup, send_command_async=send)),
        )

    def click(self, popup, action):
        if action == 'answer':
            popup.answer(lambda: None, lambda: None)
        else:
            popup.hangup(lambda: None)

    def settle(self):
        for thread in self.threads:
            thread.join(3)
            self.assertFalse(thread.is_alive(), 'sender deadlocked')
        while self.ui_tasks:
            self.ui_tasks.pop(0)()
        self.assertEqual(self.errors, [])

    def test_queued_actions_revalidate_after_each_sender_lock(self):
        for lock_name in ('serial', 'transaction', 'write'):
            for action in ('answer', 'hangup'):
                for change in ('redial', 'different_number', 'ended', 'reset', 'replace_serial', 'generation', 'shutdown'):
                    with self.subTest(lock=lock_name, action=action, change=change):
                        self.setUp()
                        blocker = ObservedLock()
                        if lock_name == 'serial':
                            self.namespace['serial_lock'] = blocker
                            target = 'DEFAULT_SERIAL_TRANSACTION_LOCK'
                            substitute = serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK
                        else:
                            target = 'DEFAULT_SERIAL_' + lock_name.upper() + '_LOCK'
                            substitute = blocker
                        popup = self.start_popup()
                        with patch.object(serial_sender, target, substitute):
                            with blocker:
                                self.click(popup, action)
                                self.assertTrue(blocker.waiting.wait(3))
                                if change in ('redial', 'different_number', 'ended'):
                                    finish_incoming_call_session_namespace_runtime(self.namespace)
                                    if change != 'ended':
                                        start_incoming_call_session_namespace_runtime(
                                            self.namespace, '10086' if change == 'redial' else '10010')
                                elif change == 'reset':
                                    reset_incoming_call_session_namespace_runtime(self.namespace)
                                elif change == 'replace_serial':
                                    self.namespace['serial_obj'] = FakeSerial()
                                elif change == 'generation':
                                    self.namespace['serial_connection_generation'] += 1
                                else:
                                    self.namespace['serial_stop_event'] = threading.Event()
                                    self.namespace['serial_stop_event'].set()
                            self.settle()
                        self.assertEqual(self.serial.writes, [])
                        self.assertEqual(self.namespace['serial_obj'].writes, [])
                        self.assertEqual(len(self.results), 1)
                        self.assertFalse(self.results[0].ok)
                        self.assertEqual(self.namespace['ring_timeout_target'], 123.0)
                        if change in ('redial', 'different_number'):
                            self.assertFalse(self.tracker.snapshot().handled)
                            self.assertEqual(self.messages, [])

    def test_lingering_old_buttons_cannot_mark_or_control_new_call(self):
        for action in ('answer', 'hangup'):
            with self.subTest(action=action):
                self.setUp()
                popup = self.start_popup()
                finish_incoming_call_session_namespace_runtime(self.namespace)
                start_incoming_call_session_namespace_runtime(self.namespace, '10086')
                self.click(popup, action)
                self.assertEqual(self.threads, [])
                self.assertFalse(self.tracker.snapshot().handled)
                self.assertEqual(self.serial.writes, [])

    def test_hiding_current_popup_does_not_cancel_queued_action(self):
        for action, command in (('answer', b'ATA\r\n'), ('hangup', b'ATH\r\n')):
            with self.subTest(action=action):
                self.setUp()
                popup = self.start_popup()
                blocker = self.namespace['serial_lock']
                with blocker:
                    self.click(popup, action)
                    self.assertTrue(blocker.waiting.wait(3))
                    close_call_popup_namespace_runtime(self.namespace)
                self.settle()
                self.assertEqual(self.serial.writes, [command])
                self.assertTrue(self.results[0].ok)
                self.assertTrue(self.tracker.snapshot().handled)

    def test_call_transition_waits_until_command_bytes_are_committed(self):
        popup = self.start_popup()
        writing, release_write, transition_started, transition_finished = (
            threading.Event() for _ in range(4))

        def hold_write():
            writing.set()
            if not release_write.wait(3):
                raise RuntimeError('write was not released')

        def redial():
            transition_started.set()
            finish_incoming_call_session_namespace_runtime(self.namespace)
            start_incoming_call_session_namespace_runtime(self.namespace, '10086')
            transition_finished.set()

        self.serial.on_write = hold_write
        self.click(popup, 'hangup')
        transition = threading.Thread(target=redial)
        try:
            self.assertTrue(writing.wait(3))
            transition.start()
            self.assertTrue(transition_started.wait(3))
            self.assertFalse(transition_finished.wait(0.05))
        finally:
            release_write.set()
            if transition.ident is not None:
                transition.join(3)
        self.assertTrue(transition_finished.is_set())
        self.settle()
        self.assertEqual(self.serial.writes, [b'ATH\r\n'])
        self.assertFalse(self.tracker.snapshot().handled)


if __name__ == '__main__':
    unittest.main()
