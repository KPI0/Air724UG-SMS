import os
import queue
import tempfile
import unittest

from sms_core.app_shutdown import cleanup_and_exit_runtime, drain_log_queue, flush_log_queue, safe_set_events


class FakeEvent:
    def __init__(self, fail=False):
        self.fail = fail
        self.set_count = 0

    def set(self):
        if self.fail:
            raise RuntimeError("set failed")
        self.set_count += 1


class AppShutdownTests(unittest.TestCase):
    def test_safe_set_events_ignores_none_and_failed_events(self):
        ok = FakeEvent()
        failed = FakeEvent(fail=True)

        self.assertEqual(safe_set_events(ok, None, failed), 1)
        self.assertEqual(ok.set_count, 1)

    def test_safe_set_events_logs_failed_events(self):
        logs = []

        self.assertEqual(safe_set_events(FakeEvent(fail=True), log_error=logs.append), 0)

        self.assertTrue(any("set failed" in message for message in logs))

    def test_drain_log_queue_groups_lines_by_path(self):
        log_queue = queue.Queue()
        log_queue.put(("a.log", "a1"))
        log_queue.put(("b.log", "b1"))
        log_queue.put(("a.log", "a2"))

        self.assertEqual(
            drain_log_queue(log_queue),
            {"a.log": ["a1", "a2"], "b.log": ["b1"]},
        )
        self.assertTrue(log_queue.empty())

    def test_flush_log_queue_writes_pending_lines(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = os.path.join(tmp, "sms.log")
            log_queue = queue.Queue()
            log_queue.put((path, "line1\n"))
            log_queue.put((path, "line2\n"))

            self.assertEqual(flush_log_queue(log_queue), 2)

            with open(path, encoding="utf-8") as file:
                self.assertEqual(file.read(), "line1\nline2\n")

    def test_flush_log_queue_logs_write_failures(self):
        logs = []
        log_queue = queue.Queue()
        log_queue.put(("Z:\\missing\\sms.log", "line\n"))

        self.assertEqual(flush_log_queue(log_queue, log_error=logs.append), 0)
        self.assertTrue(any("sms.log" in message for message in logs))

    def test_cleanup_and_exit_runtime_skips_when_already_exiting(self):
        calls = []

        result = cleanup_and_exit_runtime(
            is_exiting=True,
            confirm_exit=lambda: calls.append(("confirm",)) or True,
            set_exiting=lambda value: calls.append(("exiting", value)),
            set_serial_running=lambda value: calls.append(("serial", value)),
            shutdown_events=(),
            worker_stop_events=(),
            tts_stop_event=None,
            stop_cloud_control=lambda **kwargs: calls.append(("cloud", kwargs)),
            safe_close_serial=lambda: calls.append(("close",)),
            stop_tray_icon=lambda **kwargs: calls.append(("tray", kwargs)),
            file_log_queue=queue.Queue(),
            destroy_root=lambda: calls.append(("destroy",)),
        )

        self.assertEqual(result, "already_exiting")
        self.assertEqual(calls, [])

    def test_cleanup_and_exit_runtime_skips_when_confirmation_cancelled(self):
        calls = []

        result = cleanup_and_exit_runtime(
            is_exiting=False,
            confirm_exit=lambda: False,
            set_exiting=lambda value: calls.append(("exiting", value)),
            set_serial_running=lambda value: calls.append(("serial", value)),
            shutdown_events=(),
            worker_stop_events=(),
            tts_stop_event=None,
            stop_cloud_control=lambda **kwargs: calls.append(("cloud", kwargs)),
            safe_close_serial=lambda: calls.append(("close",)),
            stop_tray_icon=lambda **kwargs: calls.append(("tray", kwargs)),
            file_log_queue=queue.Queue(),
            destroy_root=lambda: calls.append(("destroy",)),
        )

        self.assertEqual(result, "cancelled")
        self.assertEqual(calls, [])

    def test_cleanup_and_exit_runtime_runs_shutdown_steps(self):
        calls = []
        shutdown_event = FakeEvent()
        worker_event = FakeEvent()
        tts_event = FakeEvent()
        file_event = FakeEvent()

        result = cleanup_and_exit_runtime(
            is_exiting=False,
            confirm_exit=lambda: True,
            set_exiting=lambda value: calls.append(("exiting", value)),
            set_serial_running=lambda value: calls.append(("serial", value)),
            shutdown_events=(shutdown_event,),
            worker_stop_events=(worker_event,),
            tts_stop_event=tts_event,
            stop_cloud_control=lambda **kwargs: calls.append(("cloud", kwargs)),
            safe_close_serial=lambda: calls.append(("close",)),
            stop_tray_icon=lambda **kwargs: calls.append(("tray", kwargs)),
            file_log_queue=queue.Queue(),
            destroy_root=lambda: calls.append(("destroy",)),
            safe_set_events=lambda *events: calls.append(("events", events)),
            flush_log_queue=lambda log_queue: calls.append(("flush", log_queue)),
            file_log_thread="file_thread",
            file_log_stop_event=file_event,
            worker_threads=("producer_thread",),
            wait_worker_threads=lambda threads: calls.append(("wait_workers", threads)),
            wait_file_log_worker=lambda thread: calls.append(("wait_file", thread)),
        )

        self.assertEqual(result, "exited")
        self.assertEqual(calls, [
            ("exiting", True),
            ("events", (shutdown_event,)),
            ("serial", False),
            ("events", (worker_event, tts_event)),
            ("cloud", {"update_status": False}),
            ("close",),
            ("tray", {"wait_after": 0.25}),
            ("wait_workers", ("producer_thread",)),
            ("events", (file_event,)),
            ("wait_file", "file_thread"),
            ("flush", calls[10][1]),
            ("destroy",),
        ])

    def test_cleanup_resolves_worker_snapshot_after_stopping_producers(self):
        calls = []

        result = cleanup_and_exit_runtime(
            is_exiting=False,
            confirm_exit=lambda: True,
            set_exiting=lambda value: calls.append(("exiting", value)),
            set_serial_running=lambda value: calls.append(("serial", value)),
            shutdown_events=(),
            worker_stop_events=(),
            tts_stop_event=None,
            stop_cloud_control=lambda **kwargs: calls.append(("cloud", kwargs)),
            safe_close_serial=lambda: calls.append(("close",)),
            stop_tray_icon=lambda **kwargs: calls.append(("tray", kwargs)),
            file_log_queue="queue",
            destroy_root=lambda: calls.append(("destroy",)),
            worker_threads=lambda: calls.append(("snapshot",)) or ("late-worker",),
            wait_worker_threads=lambda threads: calls.append(("wait", threads)),
            wait_file_log_worker=lambda thread: None,
            flush_log_queue=lambda log_queue: None,
        )

        self.assertEqual(result, "exited")
        self.assertLess(calls.index(("close",)), calls.index(("snapshot",)))
        self.assertEqual(calls[calls.index(("snapshot",)) + 1], ("wait", ("late-worker",)))

    def test_cleanup_aborts_when_worker_snapshot_raises(self):
        calls = []
        file_event = FakeEvent()

        result = cleanup_and_exit_runtime(
            is_exiting=False,
            confirm_exit=lambda: True,
            set_exiting=lambda value: calls.append(("exiting", value)),
            set_serial_running=lambda value: calls.append(("serial", value)),
            shutdown_events=(),
            worker_stop_events=(),
            tts_stop_event=None,
            stop_cloud_control=lambda **kwargs: calls.append(("cloud", kwargs)),
            safe_close_serial=lambda: calls.append(("close",)),
            stop_tray_icon=lambda **kwargs: calls.append(("tray", kwargs)),
            file_log_queue="queue",
            destroy_root=lambda: calls.append(("destroy",)),
            file_log_thread="file_thread",
            file_log_stop_event=file_event,
            worker_threads=lambda: (_ for _ in ()).throw(
                RuntimeError("snapshot failed")
            ),
            wait_worker_threads=lambda threads: calls.append(("wait_workers", threads)),
            wait_file_log_worker=lambda thread: calls.append(("wait_file", thread)),
            flush_log_queue=lambda log_queue: calls.append(("flush", log_queue)),
            log_error=lambda message: calls.append(("log", message)),
        )

        self.assertEqual(result, "worker_wait_failed")
        self.assertTrue(
            any("snapshot failed" in item[1] for item in calls if item[0] == "log")
        )
        self.assertEqual(file_event.set_count, 0)
        self.assertFalse(
            any(
                item[0] in ("wait_workers", "wait_file", "flush", "destroy")
                for item in calls
            )
        )

    def test_cleanup_does_not_stop_file_logger_or_exit_when_producer_is_running(self):
        calls = []
        file_event = FakeEvent()

        result = cleanup_and_exit_runtime(
            is_exiting=False,
            confirm_exit=lambda: True,
            set_exiting=lambda value: calls.append(("exiting", value)),
            set_serial_running=lambda value: calls.append(("serial", value)),
            shutdown_events=(), worker_stop_events=(), tts_stop_event=None,
            stop_cloud_control=lambda **kwargs: None,
            safe_close_serial=lambda: None,
            stop_tray_icon=lambda **kwargs: None,
            file_log_queue="queue",
            destroy_root=lambda: calls.append(("destroy",)),
            file_log_thread="file_thread",
            file_log_stop_event=file_event,
            worker_threads=("serial-command",),
            wait_worker_threads=lambda threads: False,
            wait_file_log_worker=lambda thread: calls.append(("wait_file", thread)),
            flush_log_queue=lambda log_queue: calls.append(("flush", log_queue)),
        )

        self.assertEqual(result, "worker_wait_failed")
        self.assertEqual(file_event.set_count, 0)
        self.assertNotIn(("wait_file", "file_thread"), calls)
        self.assertNotIn(("flush", "queue"), calls)
        self.assertNotIn(("destroy",), calls)

    def test_cleanup_does_not_continue_when_producer_wait_raises(self):
        calls = []

        result = cleanup_and_exit_runtime(
            is_exiting=False,
            confirm_exit=lambda: True,
            set_exiting=lambda value: calls.append(("exiting", value)),
            set_serial_running=lambda value: calls.append(("serial", value)),
            shutdown_events=(), worker_stop_events=(), tts_stop_event=None,
            stop_cloud_control=lambda **kwargs: None,
            safe_close_serial=lambda: None,
            stop_tray_icon=lambda **kwargs: None,
            file_log_queue="queue",
            destroy_root=lambda: calls.append(("destroy",)),
            worker_threads=("serial-command",),
            wait_worker_threads=lambda threads: (_ for _ in ()).throw(RuntimeError("join failed")),
            wait_file_log_worker=lambda thread: calls.append(("wait_file", thread)),
            flush_log_queue=lambda log_queue: calls.append(("flush", log_queue)),
            log_error=lambda message: calls.append(("log", message)),
        )

        self.assertEqual(result, "worker_wait_failed")
        self.assertFalse(any(item[0] in ("wait_file", "flush", "destroy") for item in calls))
        self.assertTrue(any("join failed" in item[1] for item in calls if item[0] == "log"))

    def test_cleanup_does_not_flush_or_destroy_when_file_logger_is_still_running(self):
        calls = []

        result = cleanup_and_exit_runtime(
            is_exiting=False,
            confirm_exit=lambda: True,
            set_exiting=lambda value: calls.append(("exiting", value)),
            set_serial_running=lambda value: calls.append(("serial", value)),
            shutdown_events=(),
            worker_stop_events=(),
            tts_stop_event=None,
            stop_cloud_control=lambda **kwargs: calls.append(("cloud", kwargs)),
            safe_close_serial=lambda: calls.append(("close",)),
            stop_tray_icon=lambda **kwargs: calls.append(("tray", kwargs)),
            file_log_queue="queue",
            destroy_root=lambda: calls.append(("destroy",)),
            file_log_thread="file_thread",
            file_log_stop_event=FakeEvent(),
            worker_threads=(),
            wait_worker_threads=lambda threads: True,
            wait_file_log_worker=lambda thread: False,
            flush_log_queue=lambda log_queue: calls.append(("flush", log_queue)),
        )

        self.assertEqual(result, "file_log_wait_failed")
        self.assertNotIn(("flush", "queue"), calls)
        self.assertNotIn(("destroy",), calls)

    def test_cleanup_does_not_flush_or_destroy_when_file_logger_wait_raises(self):
        calls = []

        result = cleanup_and_exit_runtime(
            is_exiting=False,
            confirm_exit=lambda: True,
            set_exiting=lambda value: calls.append(("exiting", value)),
            set_serial_running=lambda value: calls.append(("serial", value)),
            shutdown_events=(), worker_stop_events=(), tts_stop_event=None,
            stop_cloud_control=lambda **kwargs: None,
            safe_close_serial=lambda: None,
            stop_tray_icon=lambda **kwargs: None,
            file_log_queue="queue",
            destroy_root=lambda: calls.append(("destroy",)),
            file_log_thread="file_thread",
            file_log_stop_event=FakeEvent(),
            worker_threads=(),
            wait_worker_threads=lambda threads: True,
            wait_file_log_worker=lambda thread: (_ for _ in ()).throw(RuntimeError("join failed")),
            flush_log_queue=lambda log_queue: calls.append(("flush", log_queue)),
            log_error=lambda message: calls.append(("log", message)),
        )

        self.assertEqual(result, "file_log_wait_failed")
        self.assertNotIn(("flush", "queue"), calls)
        self.assertNotIn(("destroy",), calls)
        self.assertTrue(any("join failed" in item[1] for item in calls if item[0] == "log"))

    def test_cleanup_and_exit_runtime_logs_and_continues_after_cleanup_errors(self):
        logs = []
        calls = []

        result = cleanup_and_exit_runtime(
            is_exiting=False,
            confirm_exit=lambda: True,
            set_exiting=lambda value: calls.append(("exiting", value)),
            set_serial_running=lambda value: calls.append(("serial", value)),
            shutdown_events=(FakeEvent(fail=True),),
            worker_stop_events=(),
            tts_stop_event=None,
            stop_cloud_control=lambda **kwargs: (_ for _ in ()).throw(RuntimeError("cloud")),
            safe_close_serial=lambda: (_ for _ in ()).throw(RuntimeError("serial")),
            stop_tray_icon=lambda **kwargs: (_ for _ in ()).throw(RuntimeError("tray")),
            file_log_queue=queue.Queue(),
            destroy_root=lambda: (_ for _ in ()).throw(RuntimeError("destroy")),
            log_error=logs.append,
        )

        self.assertEqual(result, "exited")
        self.assertIn(("exiting", True), calls)
        self.assertIn(("serial", False), calls)
        self.assertTrue(any("set failed" in message for message in logs))
        self.assertTrue(any("cloud" in message for message in logs))
        self.assertTrue(any("serial" in message for message in logs))
        self.assertTrue(any("tray" in message for message in logs))
        self.assertTrue(any("destroy" in message for message in logs))


if __name__ == "__main__":
    unittest.main()
