from pathlib import Path
import queue
import tempfile
import threading
import tkinter as tk
from types import SimpleNamespace
import unittest

from sms_core.app_shutdown import cleanup_and_exit_runtime, flush_log_queue, safe_set_events, wait_for_pending_cloud_events
from sms_core.third_push_runtime import third_push_worker_runtime
from sms_ui.app_shutdown_runtime import start_cleanup_and_exit_app_runtime
from sms_ui.shutdown_progress import ShutdownProgress


class ShutdownOrderTests(unittest.TestCase):
    def test_receiver_finishes_before_cloud_stops_and_final_push_and_logs_are_drained(self):
        serial_stop = threading.Event()
        push_stop = threading.Event()
        pending = queue.Queue()
        logs = queue.Queue()
        received = threading.Event()
        order = []

        def receiver():
            self.assertTrue(serial_stop.wait(2))
            pending.put({"message": "last-sms", "channels": ["custom_post"]})
            logs.put(("unused", "last-sms"))
            order.append("received")
            received.set()

        serial_thread = threading.Thread(target=receiver)
        serial_thread.start()
        push_thread = threading.Thread(target=third_push_worker_runtime, kwargs={
            "stop_event": push_stop, "push_queue": pending,
            "send_channel_func": lambda *a: order.append("pushed") or (True, "ok"),
            "system_ui": lambda *a: None, "show_result": lambda *a: None, "poll_timeout": 0.01,
        })
        push_thread.start()

        def stop_cloud(**kwargs):
            self.assertTrue(received.is_set())
            order.append("cloud-stopped")

        result = cleanup_and_exit_runtime(
            is_exiting=False, confirm_exit=lambda: True, set_exiting=lambda _v: None,
            set_serial_running=lambda _v: None, shutdown_events=(), worker_stop_events=(serial_stop,),
            tts_stop_event=None, pre_cloud_worker_threads=lambda: (serial_thread,),
            before_stop_cloud=lambda: order.append("cloud-flushed"), stop_cloud_control=stop_cloud,
            safe_close_serial=lambda: None, stop_tray_icon=lambda **k: None, worker_threads=(serial_thread,),
            deferred_worker_stop_events=(push_stop,), deferred_worker_threads=(push_thread,),
            deferred_worker_queues=(pending,), file_log_queue=logs,
            flush_log_queue=lambda q: order.append(q.get_nowait()[1]), destroy_root=lambda: order.append("destroyed"),
        )
        self.assertEqual(result, "exited")
        self.assertLess(order.index("received"), order.index("cloud-flushed"))
        self.assertLess(order.index("cloud-flushed"), order.index("cloud-stopped"))
        self.assertLess(order.index("pushed"), order.index("last-sms"))
        self.assertEqual(order[-1], "destroyed")
        self.assertFalse(serial_thread.is_alive())
        self.assertFalse(push_thread.is_alive())
        self.assertEqual(pending.unfinished_tasks, 0)

    def test_cloud_ack_wait_is_bounded_and_never_discards_unconfirmed_events(self):
        pending = queue.Queue()
        pending.put({"type": "sms"})
        notices = []
        self.assertFalse(wait_for_pending_cloud_events(pending, lambda: False, log_error=notices.append))
        self.assertFalse(wait_for_pending_cloud_events(pending, lambda: True, timeout=0.01, log_error=notices.append))
        self.assertEqual(pending.qsize(), 1)
        self.assertEqual(pending.unfinished_tasks, 1)
        self.assertEqual(len(notices), 2)
        item = pending.get_nowait()
        self.assertEqual(item["type"], "sms")
        timer = threading.Timer(0.02, pending.task_done)
        timer.start()
        try:
            self.assertTrue(wait_for_pending_cloud_events(pending, lambda: True, timeout=1))
        finally:
            timer.join()


class ShutdownResponsivenessTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(str(exc))
        self.root.withdraw()
        self.state = {"exiting": False}
        self.ui_calls = []
        self.errors = []
        self.owner = threading.get_ident()
        self.messagebox = SimpleNamespace(
            askyesno=lambda *a, **k: True,
            showerror=lambda *a, **k: self.errors.append((threading.get_ident(), a)),
        )

    def tearDown(self):
        self.root.destroy()

    def options(self):
        test = self

        class Progress(ShutdownProgress):
            def __init__(self, root):
                test.ui_calls.append(threading.get_ident())
                super().__init__(root)

            def update(self, message):
                test.ui_calls.append(threading.get_ident())
                super().update(message)

            def close(self):
                test.ui_calls.append(threading.get_ident())
                super().close()

        return dict(
            root=self.root, messagebox=self.messagebox, is_exiting=self.state["exiting"],
            set_exiting=lambda value: self.state.update(exiting=value),
            set_serial_running=lambda _v: None, shutdown_events=(threading.Event(),), worker_stop_events=(),
            tts_stop_event=threading.Event(), safe_set_events=safe_set_events,
            stop_cloud_control=lambda **k: None, safe_close_serial=lambda: None,
            stop_tray_icon=lambda **k: None, flush_log_queue=flush_log_queue,
            file_log_queue=queue.Queue(), progress_factory=Progress,
            destroy_root=lambda: (self.ui_calls.append(threading.get_ident()), self.root.quit()),
        )

    def test_slow_push_shutdown_keeps_tk_responsive_and_flushes_logs(self):
        release = threading.Event()
        pending = queue.Queue()
        stop = threading.Event()
        sent = []
        heartbeat_count = []
        heartbeat_timer = None
        for text in ("first", "second", "third"):
            pending.put({"message": text, "channels": ["custom_post"]})

        def send(_channel, body, _settings):
            if not release.wait(2):
                return False, "UI did not release the sender"
            sent.append(body)
            return True, "ok"

        worker = threading.Thread(target=third_push_worker_runtime, kwargs={
            "stop_event": stop, "push_queue": pending, "send_channel_func": send,
            "system_ui": lambda *a: None, "show_result": lambda *a: None,
            "format_message_func": lambda message, *a, **k: message,
        }, daemon=True)
        worker.start()

        def heartbeat():
            nonlocal heartbeat_timer
            heartbeat_count.append(True)
            heartbeat_timer = self.root.after(10, heartbeat)

        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "sms.txt"
            options = self.options()
            options["file_log_queue"].put((str(path), "last-sms\n"))
            options.update(deferred_worker_stop_events=(stop,), deferred_worker_threads=(worker,),
                           deferred_worker_queues=(pending,))
            self.assertEqual(start_cleanup_and_exit_app_runtime(**options), "started")
            self.assertEqual(start_cleanup_and_exit_app_runtime(**self.options()), "already_exiting")
            self.root.after(0, heartbeat)
            self.root.after(80, release.set)
            watchdog = self.root.after(3000, self.root.quit)
            try:
                self.root.mainloop()
            finally:
                release.set()
                stop.set()
                worker.join(2)
                self.root.after_cancel(watchdog)
                if heartbeat_timer is not None:
                    self.root.after_cancel(heartbeat_timer)
            self.assertEqual(sent, ["first", "second", "third"])
            self.assertEqual(path.read_text(), "last-sms\n")
        self.assertGreaterEqual(len(heartbeat_count), 2)
        self.assertTrue(self.ui_calls)
        self.assertEqual(set(self.ui_calls), {self.owner})
        self.assertEqual(self.errors, [])
        self.assertEqual(pending.unfinished_tasks, 0)

    def test_cancel_does_not_start_shutdown(self):
        self.messagebox.askyesno = lambda *a, **k: False
        self.assertEqual(start_cleanup_and_exit_app_runtime(**self.options()), "cancelled")
        self.assertFalse(self.state["exiting"])
        self.assertEqual(self.ui_calls, [])

    def test_thread_start_failure_leaves_receiver_running_and_allows_retry(self):
        options = self.options()
        shutdown = options["shutdown_events"][0]
        options["start_worker"] = lambda *a, **k: (_ for _ in ()).throw(RuntimeError("synthetic"))
        self.assertEqual(start_cleanup_and_exit_app_runtime(**options), "start_failed")
        self.assertFalse(shutdown.is_set())
        self.assertFalse(self.state["exiting"])
        self.assertEqual(self.errors[0][0], self.owner)

    def test_cleanup_failure_is_reported_on_ui_and_keeps_window_for_retry(self):
        options = self.options()
        options["cleanup_runtime"] = lambda **k: "worker_wait_failed"
        self.messagebox.showerror = lambda *a, **k: (self.errors.append((threading.get_ident(), a)), self.root.quit())
        self.assertEqual(start_cleanup_and_exit_app_runtime(**options), "started")
        watchdog = self.root.after(3000, self.root.quit)
        try:
            self.root.mainloop()
        finally:
            self.root.after_cancel(watchdog)
        self.assertEqual(self.errors[0][0], self.owner)
        self.assertFalse(self.state["exiting"])
        self.assertTrue(self.root.winfo_exists())


if __name__ == "__main__":
    unittest.main()
