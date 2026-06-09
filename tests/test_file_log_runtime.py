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
