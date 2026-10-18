import unittest
from functools import partial
from datetime import datetime
from types import SimpleNamespace

from sms_core.call_effects import apply_call_hangup_result
from sms_core.call_session import IncomingCallSessionTracker
from sms_ui.call_popup_namespace_runtime import (
    _get_peer_call_binding,
    close_call_popup_namespace_runtime,
    finish_incoming_call_session_namespace_runtime,
    finish_remote_incoming_call_namespace_runtime,
    mark_peer_incoming_call_connected_namespace_runtime,
    register_peer_incoming_call_namespace_runtime,
    reset_incoming_call_session_namespace_runtime,
    show_call_popup_namespace_runtime,
)
from sms_ui.call_popup_runtime import show_call_popup_app_runtime, show_call_popup_runtime


class CallPopupHandoverTests(unittest.TestCase):
    def setUp(self):
        # Coarse Windows clocks can give consecutive calls the same time.
        self.tracker = IncomingCallSessionTracker(now_func=lambda: datetime(2026, 1, 1))
        self.missed = []
        self.popups = []
        self.namespace = {
            "INCOMING_CALL_SESSION": self.tracker,
            "current_call_popup": None, "current_dial_popup": None,
            "root": None, "serial_lock": None, "serial_obj": None,
            "center_window": lambda *_: None,
            "port_ui": lambda *_: None, "set_status": lambda *_: None,
            "run_on_ui_thread": lambda callback, _post: callback(),
            "ui_post": lambda callback: callback(),
            "show_missed_call_popup": self.missed.append,
            "mark_incoming_call_handled": self.tracker.mark_handled,
            "close_call_popup": lambda: close_call_popup_namespace_runtime(self.namespace),
            "_call_peer_monotonic": lambda: 100.0,
        }

    def start(self, session_id, caller="10086"):
        self.tracker.start(caller)

        def show(**kwargs):
            popup = SimpleNamespace(
                _call_popup_caller_num=caller, exists=True,
                _call_popup_mark_connected=self.tracker.mark_handled,
                close=kwargs["close_popup"],
            )
            popup.winfo_exists = lambda: popup.exists
            popup.destroy = lambda: setattr(popup, "exists", False)
            self.popups.append(popup)
            kwargs["set_popup"](popup)
            return popup

        return show_call_popup_namespace_runtime(
            self.namespace, caller, session_id, show_app_runtime=show)

    def test_late_serial_finish_preserves_next_unbound_peer(self):
        self.start("local-old")
        register_peer_incoming_call_namespace_runtime(self.namespace, "peer-old", "10086")
        finish_remote_incoming_call_namespace_runtime(
            self.namespace, session_id="peer-old", caller_num="10086")
        register_peer_incoming_call_namespace_runtime(self.namespace, "peer-new", "10086")
        finish_incoming_call_session_namespace_runtime(self.namespace)
        close_call_popup_namespace_runtime(self.namespace)
        popup = self.start("local-new")
        self.assertTrue(register_peer_incoming_call_namespace_runtime(
            self.namespace, "peer-new", "10086"))
        self.assertTrue(mark_peer_incoming_call_connected_namespace_runtime(
            self.namespace, "peer-new", "10086"))
        self.assertTrue(self.tracker.snapshot().handled)
        self.assertIs(self.namespace["current_call_popup"], popup)

    def test_next_peer_does_not_replace_current_binding_on_connected(self):
        self.start("local-old")
        register_peer_incoming_call_namespace_runtime(self.namespace, "peer-old", "10086")
        register_peer_incoming_call_namespace_runtime(self.namespace, "peer-new", "10086")
        self.assertTrue(mark_peer_incoming_call_connected_namespace_runtime(
            self.namespace, "peer-old", "10086"))
        self.assertEqual(_get_peer_call_binding(self.namespace)["peer_session_id"], "peer-old")
        finish_incoming_call_session_namespace_runtime(self.namespace)
        close_call_popup_namespace_runtime(self.namespace)
        self.start("local-new")
        self.assertEqual(_get_peer_call_binding(self.namespace)["peer_session_id"], "peer-new")
        self.assertFalse(self.tracker.snapshot().handled)

    def test_reset_discards_unbound_frame_without_retiring_its_id(self):
        register_peer_incoming_call_namespace_runtime(self.namespace, "peer-new", "10086")
        reset_incoming_call_session_namespace_runtime(self.namespace)
        self.assertIsNone(_get_peer_call_binding(self.namespace))
        self.assertTrue(register_peer_incoming_call_namespace_runtime(
            self.namespace, "peer-new", "10086"))

    def test_expired_next_peer_is_not_bound_to_later_call(self):
        self.start("local-old")
        register_peer_incoming_call_namespace_runtime(self.namespace, "peer-old", "10086")
        register_peer_incoming_call_namespace_runtime(self.namespace, "peer-new", "10086")
        self.namespace["_call_peer_monotonic"] = lambda: 200.0
        finish_incoming_call_session_namespace_runtime(self.namespace)
        close_call_popup_namespace_runtime(self.namespace)
        self.start("local-new")
        self.assertIsNone(_get_peer_call_binding(self.namespace))

    def test_next_peer_announced_before_old_finish_is_preserved(self):
        for remote_finish in (False, True):
            with self.subTest(remote_finish=remote_finish):
                self.setUp()
                self.start("local-old")
                register_peer_incoming_call_namespace_runtime(self.namespace, "peer-old", "10086")
                register_peer_incoming_call_namespace_runtime(self.namespace, "peer-new", "10086")
                if remote_finish:
                    finish_remote_incoming_call_namespace_runtime(
                        self.namespace, session_id="peer-old", caller_num="10086")
                else:
                    finish_incoming_call_session_namespace_runtime(self.namespace)
                    close_call_popup_namespace_runtime(self.namespace)
                self.start("local-new")
                binding = _get_peer_call_binding(self.namespace)
                self.assertIsNotNone(binding)
                self.assertEqual(binding["peer_session_id"], "peer-new")
                self.assertEqual(binding["local_session_id"], "local-new")
                self.assertFalse(register_peer_incoming_call_namespace_runtime(
                    self.namespace, "peer-old", "10086"))
                self.assertTrue(finish_remote_incoming_call_namespace_runtime(
                    self.namespace, session_id="peer-new", caller_num="10086"))
                self.assertIsNone(self.namespace["current_call_popup"])

    def test_late_hangup_result_cannot_close_new_incoming_popup(self):
        old_popup = self.start("local-old")
        finish_incoming_call_session_namespace_runtime(self.namespace)
        close_call_popup_namespace_runtime(self.namespace)
        new_popup = self.start("local-new")
        apply_call_hangup_result(
            SimpleNamespace(ok=True), lambda: None, lambda *_: None,
            lambda callback: callback(), old_popup.close,
        )
        self.assertIs(self.namespace["current_call_popup"], new_popup)
        self.assertTrue(new_popup.exists)
        self.assertEqual(self.namespace["call_popup_active_session_id"], "local-new")

    def test_current_popup_close_still_works_with_or_without_session_id(self):
        for session_id in ("local-1", ""):
            with self.subTest(session_id=session_id):
                popup = self.start(session_id)
                popup.close()
                self.assertFalse(popup.exists)
                self.assertIsNone(self.namespace["current_call_popup"])

    def test_queued_button_close_checks_generation_when_callback_runs(self):
        for session_id in ("local-old", ""):
            with self.subTest(session_id=session_id):
                self.setUp()
                old_popup = self.start(session_id)
                queued = []
                self.namespace["run_on_ui_thread"] = lambda callback, _post: queued.append(callback)
                old_popup.close()
                new_popup = self.start("local-new")
                queued.pop(0)()
                self.assertIs(self.namespace["current_call_popup"], new_popup)
                self.assertTrue(new_popup.exists)
                self.assertEqual(self.namespace["call_popup_active_session_id"], "local-new")

    def test_queued_close_only_destroys_the_captured_popup(self):
        old_popup = self.start("local-old")
        queued = []
        self.namespace["run_on_ui_thread"] = lambda callback, _post: queued.append(callback)
        close_call_popup_namespace_runtime(self.namespace)
        self.namespace["current_call_popup"] = None
        new_popup = self.start("local-new")
        queued.pop(0)()
        self.assertIs(self.namespace["current_call_popup"], new_popup)
        self.assertTrue(new_popup.exists)
        self.assertFalse(old_popup.exists)

    def open_during(self, session_id, during_open=lambda: None, *, send_command_async=None):
        def open_popup(*args):
            popup = SimpleNamespace(_call_popup_caller_num='10086', exists=True)
            popup.answer = args[3]
            popup.hangup = args[4]
            popup.winfo_exists = lambda: popup.exists
            popup.destroy = lambda: setattr(popup, 'exists', False)
            self.popups.append(popup)
            during_open()
            return popup

        options = {'open_popup': open_popup}
        if send_command_async is not None:
            options['send_command_async'] = send_command_async
        return show_call_popup_namespace_runtime(
            self.namespace, '10086', session_id,
            show_app_runtime=partial(show_call_popup_app_runtime,
                show_runtime=partial(show_call_popup_runtime, **options)))

    def test_late_answer_result_does_not_disable_new_calls_ring_timeout(self):
        results, statuses = [], []
        self.tracker.start('10086')
        self.namespace['set_status'] = lambda *args: statuses.append(args)
        old = self.open_during('local-old', send_command_async=lambda *a, **k: results.append(k['on_result']))
        old.answer(lambda: None, lambda: None)
        finish_incoming_call_session_namespace_runtime(self.namespace)
        close_call_popup_namespace_runtime(self.namespace)
        new = self.start('local-new')
        self.namespace['ring_timeout_target'] = 123.0
        results[0](SimpleNamespace(ok=True))
        self.assertEqual(self.namespace['ring_timeout_target'], 123.0)
        self.assertEqual(statuses, [])
        self.assertIs(self.namespace['current_call_popup'], new)

    def test_queued_answer_result_rechecks_call_when_ui_runs(self):
        results, queued = [], []
        self.tracker.start('10086')
        old = self.open_during('local-old', send_command_async=lambda *a, **k: results.append(k['on_result']))
        old.answer(lambda: None, lambda: None)
        self.namespace['run_on_ui_thread'] = lambda cb, _post: queued.append(cb)
        self.namespace['ui_post'] = queued.append
        self.namespace['ring_timeout_target'] = 456.0
        results[0](SimpleNamespace(ok=True))
        self.assertEqual(self.namespace['ring_timeout_target'], 456.0)
        self.namespace['run_on_ui_thread'] = lambda cb, _post: cb()
        finish_incoming_call_session_namespace_runtime(self.namespace)
        close_call_popup_namespace_runtime(self.namespace)
        self.start('local-new')
        self.namespace['ring_timeout_target'] = 123.0
        for callback in queued:
            callback()
        self.assertEqual(self.namespace['ring_timeout_target'], 123.0)

    def test_answer_result_still_applies_after_hiding_current_popup(self):
        results = []
        self.tracker.start('10086')
        popup = self.open_during('local-1', send_command_async=lambda *a, **k: results.append(k['on_result']))
        popup.answer(lambda: None, lambda: None)
        close_call_popup_namespace_runtime(self.namespace)
        self.namespace['ring_timeout_target'] = 123.0
        results[0](SimpleNamespace(ok=True))
        self.assertEqual(self.namespace['ring_timeout_target'], -1.0)
        self.assertEqual(self.tracker.snapshot().caller_num, '10086')

    def test_late_failed_actions_do_not_show_errors_for_new_call(self):
        for action in ('answer', 'hangup'):
            with self.subTest(action=action):
                self.setUp()
                results, messages, restored = [], [], []
                self.tracker.start('10086')
                self.namespace['port_ui'] = lambda *a: messages.append(a)
                popup = self.open_during('local-old', send_command_async=lambda *a, **k: results.append(k['on_result']))
                if action == 'answer':
                    popup.answer(lambda: None, lambda: restored.append(True))
                else:
                    popup.hangup(lambda: restored.append(True))
                finish_incoming_call_session_namespace_runtime(self.namespace)
                close_call_popup_namespace_runtime(self.namespace)
                self.start('local-new')
                results[0](SimpleNamespace(ok=False, error='serial disconnected'))
                self.assertEqual(messages, [])
                self.assertEqual(restored, [])

    def test_current_failed_actions_still_restore_controls(self):
        for action in ('answer', 'hangup'):
            with self.subTest(action=action):
                self.setUp()
                results, messages, restored = [], [], []
                self.tracker.start('10086')
                self.namespace['port_ui'] = lambda *a: messages.append(a)
                popup = self.open_during('local-1', send_command_async=lambda *a, **k: results.append(k['on_result']))
                if action == 'answer':
                    popup.answer(lambda: None, lambda: restored.append(True))
                else:
                    popup.hangup(lambda: restored.append(True))
                results[0](SimpleNamespace(ok=False, error='serial disconnected'))
                self.assertEqual(restored, [True])
                self.assertEqual(len(messages), 1)
                self.assertIs(self.namespace['current_call_popup'], popup)

    def test_answer_result_after_reset_is_ignored(self):
        results = []
        self.tracker.start('10086')
        popup = self.open_during('local-1', send_command_async=lambda *a, **k: results.append(k['on_result']))
        popup.answer(lambda: None, lambda: None)
        reset_incoming_call_session_namespace_runtime(self.namespace)
        self.namespace['ring_timeout_target'] = 0.0
        results[0](SimpleNamespace(ok=True))
        self.assertEqual(self.namespace['ring_timeout_target'], 0.0)

    def test_legacy_no_id_result_still_isolates_same_number_redial(self):
        results = []
        self.tracker.start('10086')
        popup = self.open_during('', send_command_async=lambda *a, **k: results.append(k['on_result']))
        popup.answer(lambda: None, lambda: None)
        finish_incoming_call_session_namespace_runtime(self.namespace)
        close_call_popup_namespace_runtime(self.namespace)
        self.start('')
        self.namespace['ring_timeout_target'] = 123.0
        results[0](SimpleNamespace(ok=True))
        self.assertEqual(self.namespace['ring_timeout_target'], 123.0)

    def test_call_ending_during_window_creation_does_not_leave_orphan(self):
        for session_id in ('local-old', ''):
            with self.subTest(session_id=session_id):
                popup = self.open_during(session_id, lambda: close_call_popup_namespace_runtime(self.namespace))
                self.assertFalse(popup.exists)
                self.assertIsNone(self.namespace['current_call_popup'])

    def test_new_generation_during_old_creation_does_not_inherit_old_window(self):
        def redial():
            close_call_popup_namespace_runtime(self.namespace)
            self.start('local-new')

        old_popup = self.open_during('local-old', redial)
        self.assertFalse(old_popup.exists)
        self.assertIs(self.namespace['current_call_popup'], self.popups[-1])
        self.assertEqual(self.namespace['current_call_popup']._call_popup_session_id, 'local-new')

    def test_new_call_replaces_lingering_window_before_delayed_close(self):
        old_popup = self.start('local-old')
        queued = []
        self.namespace['run_on_ui_thread'] = lambda cb, _post: queued.append(cb)
        close_call_popup_namespace_runtime(self.namespace)
        self.namespace['run_on_ui_thread'] = lambda cb, _post: cb()
        new_popup = self.open_during('local-new')
        self.assertIsNot(new_popup, old_popup)
        self.assertFalse(old_popup.exists)
        queued.pop(0)()
        self.assertIs(self.namespace['current_call_popup'], new_popup)
        self.assertTrue(new_popup.exists)


if __name__ == "__main__":
    unittest.main()
