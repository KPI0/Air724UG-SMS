import unittest

from sms_ui.app_shutdown_runtime import cleanup_and_exit_app_runtime


class FakeRoot:
    def __init__(self):
        self.destroy_calls = 0

    def destroy(self):
        self.destroy_calls += 1


class FakeMessageBox:
    def __init__(self, answer=True):
        self.answer = answer
        self.ask_calls = []

    def askyesno(self, title, message, parent=None):
        self.ask_calls.append((title, message, parent))
        return self.answer


class AppShutdownRuntimeTests(unittest.TestCase):
    def test_cleanup_and_exit_app_runtime_wires_runtime_dependencies(self):
        calls = []
        root = FakeRoot()
        messagebox = FakeMessageBox(answer=True)

        def cleanup_runtime(**kwargs):
            calls.append(kwargs)
            self.assertTrue(kwargs["confirm_exit"]())
            kwargs["destroy_root"]()
            return "exited"

        result = cleanup_and_exit_app_runtime(
            root=root,
            messagebox=messagebox,
            is_exiting=False,
            set_exiting=lambda value: calls.append(("exiting", value)),
            set_serial_running=lambda value: calls.append(("serial", value)),
            shutdown_events=("tk",),
            worker_stop_events=("file", "third"),
            tts_stop_event="tts",
            safe_set_events=lambda *events: calls.append(("events", events)),
            stop_cloud_control=lambda **kwargs: calls.append(("cloud", kwargs)),
            safe_close_serial=lambda: calls.append(("close",)),
            stop_tray_icon=lambda **kwargs: calls.append(("tray", kwargs)),
            flush_log_queue=lambda queue: calls.append(("flush", queue)),
            file_log_queue="queue",
            file_log_thread="file_thread",
            file_log_stop_event="file_stop",
            worker_threads=("producer",),
            deferred_worker_stop_events=("third",),
            deferred_worker_threads=("third-thread",),
            deferred_worker_queues=("third-queue",),
            log_error="logger",
            cleanup_runtime=cleanup_runtime,
        )

        self.assertEqual(result, "exited")
        kwargs = calls[0]
        self.assertFalse(kwargs["is_exiting"])
        self.assertEqual(kwargs["shutdown_events"], ("tk",))
        self.assertEqual(kwargs["worker_stop_events"], ("file", "third"))
        self.assertEqual(kwargs["tts_stop_event"], "tts")
        self.assertEqual(kwargs["file_log_queue"], "queue")
        self.assertEqual(kwargs["file_log_thread"], "file_thread")
        self.assertEqual(kwargs["file_log_stop_event"], "file_stop")
        self.assertEqual(kwargs["worker_threads"], ("producer",))
        self.assertEqual(kwargs["deferred_worker_stop_events"], ("third",))
        self.assertEqual(kwargs["deferred_worker_threads"], ("third-thread",))
        self.assertEqual(kwargs["deferred_worker_queues"], ("third-queue",))
        self.assertEqual(kwargs["log_error"], "logger")
        self.assertEqual(messagebox.ask_calls[0][2], root)
        self.assertEqual(root.destroy_calls, 1)

    def test_cleanup_and_exit_app_runtime_accepts_destroy_override(self):
        calls = []
        root = FakeRoot()

        def cleanup_runtime(**kwargs):
            calls.append(kwargs)
            kwargs["destroy_root"]()
            return "ok"

        result = cleanup_and_exit_app_runtime(
            root=root,
            messagebox=FakeMessageBox(answer=True),
            is_exiting=False,
            set_exiting=lambda value: None,
            set_serial_running=lambda value: None,
            shutdown_events=(),
            worker_stop_events=(),
            tts_stop_event=None,
            safe_set_events=lambda *events: None,
            stop_cloud_control=lambda **kwargs: None,
            safe_close_serial=lambda: None,
            stop_tray_icon=lambda **kwargs: None,
            flush_log_queue=lambda queue: None,
            file_log_queue=None,
            destroy_root=lambda: calls.append("destroy"),
            cleanup_runtime=cleanup_runtime,
        )

        self.assertEqual(result, "ok")
        self.assertEqual(calls[1], "destroy")
        self.assertEqual(root.destroy_calls, 0)


if __name__ == "__main__":
    unittest.main()
