import threading
import unittest
from unittest.mock import patch

from sms_core.call_session import IncomingCallSessionTracker
from sms_ui import call_popup_namespace_runtime as runtime


class CallStateWritebackTests(unittest.TestCase):
    def replay_timer_write_during_transition(self, *, redial):
        tracker = IncomingCallSessionTracker()
        snapshot_read = threading.Event()
        resume_result = threading.Event()
        threads = {}
        failures = []

        class ObservedRLock:
            def __init__(self):
                self.lock = threading.RLock()

            def __enter__(self):
                if not self.lock.acquire(blocking=False):
                    if threading.current_thread() is threads.get("transition"):
                        # If the write is atomic, release it before the real
                        # serial transition takes the same lock.
                        resume_result.set()
                    self.lock.acquire()
                return self

            def __exit__(self, *_args):
                self.lock.release()

        namespace = {
            "INCOMING_CALL_SESSION": tracker,
            "_call_peer_state_lock": ObservedRLock(),
            "current_call_popup": None, "current_dial_popup": None,
            "serial_obj": None, "serial_lock": threading.RLock(), "root": None,
            "center_window": lambda *_: None, "port_ui": lambda *_: None,
            "set_status": lambda *_: None, "ui_post": lambda callback: callback(),
            "run_on_ui_thread": lambda callback, _post: callback(),
            "ring_timeout_target": 112.0, "current_dial_num": "",
        }
        runtime.start_incoming_call_session_namespace_runtime(namespace, "10086")
        original_token = tracker.snapshot().session_token
        popup_callbacks = {}
        runtime.show_call_popup_namespace_runtime(
            namespace, "10086", "local-old",
            show_app_runtime=lambda **kwargs: popup_callbacks.update(kwargs),
        )
        original_snapshot = runtime._call_snapshot

        def pause_after_snapshot(context):
            snapshot = original_snapshot(context)
            if threading.current_thread() is threads.get("result"):
                snapshot_read.set()
                if not resume_result.wait(3):
                    raise AssertionError("serial transition did not release the result")
            return snapshot

        def write_old_answer_result():
            try:
                popup_callbacks["set_ring_timeout"](-1.0)
            except Exception as exc:
                failures.append(exc)

        def change_call():
            try:
                if redial:
                    runtime.finish_incoming_call_session_namespace_runtime(namespace)
                    runtime.start_incoming_call_session_namespace_runtime(namespace, "10086")
                else:
                    runtime.reset_incoming_call_session_namespace_runtime(namespace)
                runtime.set_serial_call_state_namespace_runtime(
                    namespace, 123.0 if redial else 0.0, "")
            except Exception as exc:
                failures.append(exc)
            finally:
                resume_result.set()

        threads["result"] = threading.Thread(target=write_old_answer_result, daemon=True)
        threads["transition"] = threading.Thread(target=change_call, daemon=True)
        with patch.object(runtime, "_call_snapshot", side_effect=pause_after_snapshot):
            threads["result"].start()
            try:
                self.assertTrue(snapshot_read.wait(3))
                threads["transition"].start()
                threads["transition"].join(3)
            finally:
                resume_result.set()
                threads["result"].join(3)
            self.assertFalse(threads["result"].is_alive(), "result writer deadlocked")
            self.assertFalse(threads["transition"].is_alive(), "serial transition deadlocked")
        self.assertEqual(failures, [])
        self.assertEqual(namespace["ring_timeout_target"], 123.0 if redial else 0.0)
        if redial:
            self.assertIsNot(tracker.snapshot().session_token, original_token)
            self.assertEqual(tracker.snapshot().caller_num, "10086")
        else:
            self.assertIsNone(tracker.snapshot().session_token)

    def test_old_answer_timer_write_cannot_cross_same_number_redial(self):
        self.replay_timer_write_during_transition(redial=True)

    def test_old_answer_timer_write_cannot_cross_reset(self):
        self.replay_timer_write_during_transition(redial=False)


if __name__ == "__main__":
    unittest.main()
