import tkinter as tk
import unittest
from types import SimpleNamespace

from sms_core.file_log_runtime import FileLogErrorState
from sms_ui.app_restart_runtime import start_restart_software_app_runtime
from sms_ui.app_shutdown_runtime import start_cleanup_and_exit_app_runtime
from sms_ui.ui_log_namespace_runtime import notify_file_log_errors_namespace_runtime


class OpaqueNotificationError(Exception):
    def __str__(self):
        raise AssertionError("Untrusted notification text must not be formatted")

    __repr__ = __str__


class FakeRoot:
    def __init__(self):
        self.pending = {}
        self.next_id = 0

    def after(self, _delay, callback):
        self.next_id += 1
        token = str(self.next_id)
        self.pending[token] = callback
        return token

    def after_cancel(self, token):
        self.pending.pop(token, None)

    def pump(self, limit=3):
        for _ in range(limit):
            if not self.pending:
                break
            token = next(iter(self.pending))
            self.pending.pop(token)()


class ShutdownNotificationFailureTests(unittest.TestCase):
    def setUp(self):
        self.root = FakeRoot()
        self.state = {"exiting": False}
        self.events = []
        self.errors = []
        self.logs = []
        self.notice_exception = tk.TclError("Synthetic private notification detail")
        self.file_log_errors = FileLogErrorState(monotonic=lambda: 0.0)
        self.file_log_errors.report("Synthetic pending log failure")
        self.messagebox = SimpleNamespace(
            askyesno=lambda *_args, **_kwargs: True,
            showwarning=self.showwarning,
            showerror=lambda title, message, **_kwargs: self.errors.append((title, message)),
        )
        self.namespace = {
            "FILE_LOG_ERROR_STATE": self.file_log_errors,
            "messagebox": self.messagebox,
            "root": self.root,
        }
        events = self.events

        class Progress:
            def __init__(self, _root):
                events.append("progress_opened")

            def close(self):
                events.append("progress_closed")

        self.progress_factory = Progress

    def showwarning(self, *_args, **_kwargs):
        self.events.append("warning_attempted")
        raise self.notice_exception

    def options(self):
        return {
            "root": self.root,
            "messagebox": self.messagebox,
            "is_exiting": self.state["exiting"],
            "set_exiting": lambda value: self.state.update(exiting=value),
            "progress_factory": self.progress_factory,
            "start_worker": lambda _name, callback, **_kwargs: callback(),
            "notify_log_errors": lambda **kwargs: notify_file_log_errors_namespace_runtime(
                self.namespace, **kwargs,
            ),
            "log_error": self.logs.append,
        }

    def assert_finish_delivered(self, error_type):
        self.assertIn("warning_attempted", self.events)
        self.assertIn("progress_closed", self.events)
        self.assertFalse(self.root.pending, "A completed shutdown must stop its UI pump")
        self.assertEqual(len(self.logs), 1)
        self.assertIn(error_type, self.logs[0])
        self.assertNotIn("Synthetic private notification detail", self.logs[0])
        self.assertNotIn("UI task failed", self.logs[0])

    def test_log_warning_tcl_error_does_not_drop_successful_exit(self):
        result = start_cleanup_and_exit_app_runtime(
            **self.options(),
            cleanup_runtime=lambda **_kwargs: "exited",
            destroy_root=lambda: self.events.append("root_destroyed"),
        )
        self.assertEqual(result, "started")
        self.root.pump()
        self.assert_finish_delivered("TclError")
        self.assertIn("root_destroyed", self.events)
        self.assertFalse(self.errors)

    def test_log_warning_tcl_error_keeps_failed_cleanup_retryable(self):
        result = start_cleanup_and_exit_app_runtime(
            **self.options(),
            cleanup_runtime=lambda **_kwargs: "worker_wait_failed",
            destroy_root=lambda: self.events.append("root_destroyed"),
        )
        self.assertEqual(result, "started")
        self.root.pump()
        self.assert_finish_delivered("TclError")
        self.assertFalse(self.state["exiting"])
        self.assertNotIn("root_destroyed", self.events)
        self.assertEqual(len(self.errors), 1)
        self.assertEqual(self.errors[0][0], "退出未完成")
        self.assertIn("后台清理未能完成", self.errors[0][1])

    def test_shared_restart_completion_survives_log_warning_tcl_error(self):
        result = start_restart_software_app_runtime(
            **self.options(),
            restart_runtime=lambda **_kwargs: SimpleNamespace(status="exited"),
            exit_process=lambda code: self.events.append(("process_exit", code)),
            argv=["synthetic-client"],
            current_pid=1,
        )
        self.assertEqual(result, "started")
        self.root.pump()
        self.assert_finish_delivered("TclError")
        self.assertIn(("process_exit", 0), self.events)
        self.assertFalse(self.errors)

    def test_opaque_notification_error_is_not_formatted_and_exit_completes(self):
        self.notice_exception = OpaqueNotificationError("Synthetic private notification detail")
        result = start_cleanup_and_exit_app_runtime(
            **self.options(),
            cleanup_runtime=lambda **_kwargs: "exited",
            destroy_root=lambda: self.events.append("root_destroyed"),
        )
        self.assertEqual(result, "started")
        self.root.pump()
        self.assert_finish_delivered("OpaqueNotificationError")
        self.assertIn("root_destroyed", self.events)
        self.assertFalse(self.errors)


if __name__ == "__main__":
    unittest.main()
