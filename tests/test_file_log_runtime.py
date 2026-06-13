import queue
import unittest

from sms_core.file_log_runtime import (
    drain_available_log_lines,
    run_file_log_worker,
    start_file_log_worker,
    write_log_batches,
)


class FakeStopEvent:
    def __init__(self, stop_after=1):
        self.calls = 0
        self.stop_after = stop_after

    def is_set(self):
        self.calls += 1
        return self.calls > self.stop_after


class FakeFile:
    def __init__(self, writes):
        self.writes = writes

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def writelines(self, lines):
        self.writes.append(list(lines))


class FileLogRuntimeTests(unittest.TestCase):
    def test_drain_available_log_lines_groups_by_path(self):
        log_queue = queue.Queue()
        log_queue.put_nowait(("a.log", "a2\n"))
        log_queue.put_nowait(("b.log", "b1\n"))

        batches = drain_available_log_lines(log_queue, ("a.log", "a1\n"))

        self.assertEqual(batches, {
            "a.log": ["a1\n", "a2\n"],
            "b.log": ["b1\n"],
        })
        self.assertTrue(log_queue.empty())

    def test_write_log_batches_writes_each_path_once(self):
        opened = []
        writes = []

        def open_file(path, mode, encoding):
            opened.append((path, mode, encoding))
            return FakeFile(writes)

        count = write_log_batches(
            {"a.log": ["a1\n", "a2\n"], "b.log": ["b1\n"]},
            open_file=open_file,
        )

        self.assertEqual(count, 3)
        self.assertEqual(opened, [
            ("a.log", "a", "utf-8"),
            ("b.log", "a", "utf-8"),
        ])
        self.assertEqual(writes, [["a1\n", "a2\n"], ["b1\n"]])

    def test_run_file_log_worker_drains_and_writes_batch(self):
        log_queue = queue.Queue()
        log_queue.put_nowait(("a.log", "a1\n"))
        stop_event = FakeStopEvent(stop_after=1)
        writes = []

        run_file_log_worker(
            log_queue=log_queue,
            stop_event=stop_event,
            poll_timeout=0,
            write_batches=lambda batches: writes.append(batches),
        )

        self.assertEqual(writes, [{"a.log": ["a1\n"]}])

    def test_start_file_log_worker_starts_daemon_thread(self):
        calls = []

        class FakeThread:
            def __init__(self, target, daemon):
                calls.append(("thread", daemon))
                self.target = target

            def start(self):
                calls.append(("start",))

        thread = start_file_log_worker(
            log_queue="queue",
            stop_event="stop",
            thread_factory=FakeThread,
        )

        self.assertIsInstance(thread, FakeThread)
        self.assertEqual(calls, [("thread", True), ("start",)])


if __name__ == "__main__":
    unittest.main()


class FileLogWorkerResilienceTests(unittest.TestCase):
    def test_malformed_queue_item_does_not_kill_worker(self):
        # A bad item (not a (path, line) tuple) must not crash the only
        # thread that flushes logs to disk. It should be reported and skipped.
        log_queue = queue.Queue()
        log_queue.put_nowait("not-a-tuple")
        stop_event = FakeStopEvent(stop_after=1)
        errors = []

        run_file_log_worker(
            log_queue=log_queue,
            stop_event=stop_event,
            poll_timeout=0,
            on_error=errors.append,
        )

        self.assertEqual(len(errors), 1)
        self.assertIn("file_log_worker", errors[0])

    def test_good_item_after_bad_item_is_still_written(self):
        # Worker survives a malformed item and keeps processing later items.
        log_queue = queue.Queue()
        log_queue.put_nowait("bad")
        log_queue.put_nowait(("a.log", "a1\n"))
        stop_event = FakeStopEvent(stop_after=2)
        writes = []
        errors = []

        run_file_log_worker(
            log_queue=log_queue,
            stop_event=stop_event,
            poll_timeout=0,
            write_batches=lambda batches: writes.append(batches),
            on_error=errors.append,
        )

        self.assertEqual(len(errors), 1)
        self.assertEqual(writes, [{"a.log": ["a1\n"]}])

    def test_write_log_batches_reports_failing_path_and_continues(self):
        # A failing path (e.g. disk full) must be reported, not silently
        # swallowed, and must not abort writes to the remaining paths.
        errors = []
        writes = []

        def open_file(path, mode, encoding):
            if path == "bad.log":
                raise OSError("disk full")
            return FakeFile(writes)

        count = write_log_batches(
            {"bad.log": ["x\n"], "good.log": ["y\n"]},
            open_file=open_file,
            on_error=errors.append,
        )

        # good.log still written; bad.log reported.
        self.assertEqual(count, 1)
        self.assertEqual(writes, [["y\n"]])
        self.assertEqual(len(errors), 1)
        self.assertIn("bad.log", errors[0])
        self.assertIn("disk full", errors[0])

    def test_worker_forwards_on_error_to_default_write_batches(self):
        # The worker must forward its on_error into write_log_batches so that
        # per-path write failures surface through the same channel.
        errors = []
        log_queue = queue.Queue()
        log_queue.put_nowait(("bad.log", "x\n"))
        stop_event = FakeStopEvent(stop_after=1)

        def open_file(path, mode, encoding):
            raise OSError("denied")

        run_file_log_worker(
            log_queue=log_queue,
            stop_event=stop_event,
            poll_timeout=0,
            write_batches=lambda batches, on_error=None: write_log_batches(
                batches, open_file=open_file, on_error=on_error
            ),
            on_error=errors.append,
        )

        self.assertEqual(len(errors), 1)
        self.assertIn("bad.log", errors[0])

    def test_start_file_log_worker_passes_log_error_to_guard(self):
        # If the worker thread dies, the daemon guard must be able to log why.
        captured = []

        def boom_factory(target, daemon, name=None):
            class _T:
                def __init__(self):
                    self.target = target

                def start(self):
                    # Simulate the thread body running and raising.
                    self.target()

            return _T()

        # Force run_file_log_worker to raise immediately by giving a stop_event
        # whose is_set raises; the guard should capture and log it.
        class ExplodingStop:
            def is_set(self):
                raise RuntimeError("stop check failed")

        start_file_log_worker(
            log_queue=queue.Queue(),
            stop_event=ExplodingStop(),
            thread_factory=boom_factory,
            log_error=captured.append,
        )

        self.assertEqual(len(captured), 1)
        self.assertIn("file_log_worker", captured[0])

    def test_queue_read_failure_is_reported_and_backed_off(self):
        class BrokenQueue:
            def get(self, timeout):
                raise RuntimeError("queue failed")

        errors = []
        sleeps = []

        run_file_log_worker(
            log_queue=BrokenQueue(),
            stop_event=FakeStopEvent(stop_after=1),
            poll_timeout=0,
            queue_error_sleep=0.25,
            on_error=errors.append,
            sleep=sleeps.append,
        )

        self.assertEqual(len(errors), 1)
        self.assertIn("queue failed", errors[0])
        self.assertEqual(sleeps, [0.25])
