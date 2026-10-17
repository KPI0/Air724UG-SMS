import unittest
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


class CallPopupHandoverTests(unittest.TestCase):
    def setUp(self):
        self.tracker = IncomingCallSessionTracker()
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


if __name__ == "__main__":
    unittest.main()
