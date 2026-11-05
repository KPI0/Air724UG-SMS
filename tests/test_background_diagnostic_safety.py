import queue
import tempfile
import threading
import unittest
from pathlib import Path

from sms_core.threading_runtime import start_registered_daemon_thread, WorkerThreadRegistry
from sms_ui.app_shutdown_runtime import start_shutdown_task_runtime


class OpaqueError(Exception):
    def __str__(self):
        raise AssertionError('Exception text must not be formatted')

    __repr__ = __str__


class BackgroundDiagnosticSafetyTests(unittest.TestCase):
    def run_worker(self, target):
        logs = []
        unhandled = []
        registry = WorkerThreadRegistry()
        original_hook = threading.excepthook
        threading.excepthook = unhandled.append
        try:
            thread = start_registered_daemon_thread('synthetic-worker', target,
                thread_registry=registry, log_error=logs.append)
            thread.join(5)
            self.assertFalse(thread.is_alive())
            self.assertEqual(registry.snapshot(), ())
            self.assertEqual(unhandled, [])
        finally:
            threading.excepthook = original_hook
        return logs

    def test_real_file_failure_omits_private_path_and_keeps_stack_location(self):
        with tempfile.TemporaryDirectory() as directory:
            private = Path(directory) / 'password=synthetic-private-path'
            def read_missing_file():
                private.read_bytes()
            logs = self.run_worker(read_missing_file)
        self.assertEqual(len(logs), 1)
        self.assertIn('FileNotFoundError', logs[0])
        self.assertIn('read_missing_file', logs[0])
        self.assertNotIn(str(private), logs[0])
        self.assertNotIn('synthetic-private-path', logs[0])

    def test_opaque_worker_error_does_not_escape_guard(self):
        def fail():
            raise OpaqueError()
        logs = self.run_worker(fail)
        self.assertEqual(len(logs), 1)
        self.assertIn('OpaqueError', logs[0])

    def test_shutdown_failure_always_delivers_completion_to_ui(self):
        callbacks = queue.Queue()
        workers = []
        exiting = []
        errors = []
        progress_calls = []
        logs = []
        class Root:
            def after(self, _delay, callback):
                callbacks.put(callback)
                return 'synthetic-after'
        class Dialogs:
            def askyesno(self, *_args, **_kwargs):
                return True
            def showerror(self, *_args, **_kwargs):
                errors.append(True)
        class Progress:
            def __init__(self, _root):
                pass
            def update(self, _message):
                pass
            def close(self):
                progress_calls.append('closed')
        def fail(_report):
            raise OpaqueError()
        result = start_shutdown_task_runtime(root=Root(), messagebox=Dialogs(),
            is_exiting=False, set_exiting=exiting.append, run_task=fail,
            on_success=lambda: self.fail('Failed cleanup cannot close the root'),
            confirmation='synthetic confirmation',
            start_worker=lambda _name, worker, **_kwargs: workers.append(worker),
            progress_factory=Progress, log_error=logs.append)
        self.assertEqual(result, 'started')
        workers[0]()
        callbacks.get_nowait()()
        self.assertEqual(exiting, [True, False])
        self.assertEqual(progress_calls, ['closed'])
        self.assertEqual(errors, [True])
        self.assertEqual(logs, ['Background shutdown failed: OpaqueError'])
