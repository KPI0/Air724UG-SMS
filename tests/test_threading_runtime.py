import unittest
import importlib.util
from pathlib import Path


MODULE_PATH = Path(__file__).resolve().parents[1] / "sms" / "sms_core" / "threading_runtime.py"
SPEC = importlib.util.spec_from_file_location("target_threading_runtime", MODULE_PATH)
threading_runtime = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(threading_runtime)
start_daemon_thread = threading_runtime.start_daemon_thread
start_registered_daemon_thread = threading_runtime.start_registered_daemon_thread
wait_for_worker_threads = threading_runtime.wait_for_worker_threads
WorkerThreadRegistry = threading_runtime.WorkerThreadRegistry
SingleFlightTaskState = threading_runtime.SingleFlightTaskState


class FakeThread:
    def __init__(self, target, daemon, name=None):
        self.target = target
        self.daemon = daemon
        self.name = name
        self.started = False

    def start(self):
        self.started = True
        self.target()


class ThreadingRuntimeTests(unittest.TestCase):
    def test_worker_thread_registry_snapshots_and_removes_threads(self):
        registry = WorkerThreadRegistry()
        first = object()
        second = object()

        self.assertTrue(registry.register(first))
        self.assertTrue(registry.register(second))
        self.assertEqual(set(registry.snapshot()), {first, second})
        self.assertTrue(registry.unregister(first))
        self.assertFalse(registry.unregister(first))
        self.assertEqual(registry.snapshot(), (second,))

    def test_single_flight_task_state_allows_only_one_active_task(self):
        state = SingleFlightTaskState()

        self.assertTrue(state.try_acquire())
        self.assertTrue(state.is_active())
        self.assertFalse(state.try_acquire())
        self.assertTrue(state.release())
        self.assertFalse(state.is_active())
        self.assertFalse(state.release())
        self.assertTrue(state.try_acquire())

    def test_wait_for_worker_threads_joins_each_unique_thread(self):
        calls = []

        class Worker:
            name = "worker"

            def join(self, timeout):
                calls.append(("join", timeout))

            def is_alive(self):
                return False

        worker = Worker()
        times = iter((10.0, 10.0))
        self.assertTrue(
            wait_for_worker_threads(
                (worker, worker, None),
                timeout=2,
                monotonic=lambda: next(times),
                current_thread=lambda: object(),
            )
        )
        self.assertEqual(calls, [("join", 2.0)])

    def test_wait_for_worker_threads_blocks_without_production_timeout(self):
        calls = []

        class Worker:
            name = "worker"

            def join(self, timeout=None):
                calls.append(("join", timeout))

            def is_alive(self):
                return False

        self.assertTrue(
            wait_for_worker_threads(
                (Worker(),),
                current_thread=lambda: object(),
            )
        )
        self.assertEqual(calls, [("join", None)])

    def test_start_daemon_thread_logs_target_exception(self):
        logs = []

        def boom():
            raise RuntimeError("boom")

        thread = start_daemon_thread(
            "worker",
            boom,
            log_error=logs.append,
            thread_factory=FakeThread,
        )

        self.assertTrue(thread.started)
        self.assertTrue(thread.daemon)
        self.assertEqual(thread.name, "worker")
        self.assertEqual(len(logs), 1)
        self.assertIn("worker", logs[0])
        self.assertIn("boom", logs[0])
        self.assertIn("Traceback", logs[0])

    def test_start_daemon_thread_calls_before_start_before_starting(self):
        calls = []

        class OrderedThread(FakeThread):
            def start(self):
                calls.append("start")
                super().start()

        thread = start_daemon_thread(
            "ordered",
            lambda: calls.append("target"),
            before_start=lambda thread: calls.append(("before", thread)),
            thread_factory=OrderedThread,
        )

        self.assertEqual(calls, [("before", thread), "start", "target"])

    def test_registered_daemon_thread_is_visible_until_target_finishes(self):
        registry = WorkerThreadRegistry()
        seen = []

        thread = start_registered_daemon_thread(
            "registered",
            lambda: seen.extend(registry.snapshot()),
            thread_registry=registry,
            thread_factory=FakeThread,
        )

        self.assertEqual(seen, [thread])
        self.assertEqual(registry.snapshot(), ())

    def test_registered_daemon_thread_unregisters_when_start_fails(self):
        registry = WorkerThreadRegistry()

        class BrokenThread(FakeThread):
            def start(self):
                raise RuntimeError("start failed")

        with self.assertRaisesRegex(RuntimeError, "start failed"):
            start_registered_daemon_thread(
                "broken",
                lambda: None,
                thread_registry=registry,
                thread_factory=BrokenThread,
            )

        self.assertEqual(registry.snapshot(), ())


if __name__ == "__main__":
    unittest.main()
