import configparser
import queue
import tempfile
import threading
import unittest
from pathlib import Path

from sms_core.app_shutdown import cleanup_and_exit_runtime, flush_log_queue, safe_set_events
from sms_core.config_runtime import ConfigInitializationError, initialize_config_runtime, read_startup_config_values
from sms_core.file_log_runtime import FileLogErrorState, run_file_log_worker, write_log_batches
from sms_ui.thread_runtime import schedule_delayed_ui_runtime, ui_pump_runtime


class OpaqueError(Exception):
    def __str__(self):
        raise AssertionError('Untrusted exception text must not be formatted')

    __repr__ = __str__


class SupportFailureIsolationTests(unittest.TestCase):
    def test_ui_callback_error_keeps_later_tasks_and_next_pump(self):
        tasks = queue.Queue()
        calls, logs = [], []
        def fail():
            raise OpaqueError()
        tasks.put((fail, (), {}))
        tasks.put((lambda: calls.append('next'), (), {}))
        class Root:
            def after(self, delay, callback):
                calls.append(('after', delay, callback))
        callback = lambda: None
        self.assertEqual(ui_pump_runtime(tasks, Root(), lambda: True, callback, log_error=logs.append), 2)
        self.assertEqual(calls, ['next', ('after', 30, callback)])
        self.assertEqual(tasks.unfinished_tasks, 0)
        self.assertEqual(logs, ['UI task failed: OpaqueError'])

    def test_delayed_ui_schedule_failure_runs_callback_once(self):
        calls, logs = [], []
        def reject(_delay, _callback):
            raise OpaqueError()
        result = schedule_delayed_ui_runtime(lambda: calls.append('ran') or 'done',
            app_start_mono=0, start_ui_delay=2, monotonic=lambda: 1, root_after=reject,
            run_on_ui_thread=lambda callback, _post: callback(), ui_post=None, log_error=logs.append)
        self.assertEqual(result, 'done')
        self.assertEqual(calls, ['ran'])
        self.assertEqual(logs, ['Schedule delayed UI callback failed: OpaqueError'])

    def test_shutdown_event_failure_does_not_skip_next_event(self):
        logs = []
        class BrokenEvent:
            def set(self):
                raise OpaqueError()
        healthy = threading.Event()
        self.assertEqual(safe_set_events(BrokenEvent(), healthy, log_error=logs.append), 1)
        self.assertTrue(healthy.is_set())
        self.assertEqual(logs, ['Set shutdown event failed: OpaqueError'])

    def cleanup_options(self, calls):
        return dict(is_exiting=False, confirm_exit=lambda: True, set_exiting=lambda _value: None,
            set_serial_running=lambda _value: None, shutdown_events=(), worker_stop_events=(),
            tts_stop_event=None, stop_cloud_control=lambda **_kwargs: None,
            safe_close_serial=lambda: calls.append('serial'), stop_tray_icon=lambda **_kwargs: calls.append('tray'),
            file_log_queue=queue.Queue(), destroy_root=lambda: calls.append('destroy'),
            wait_worker_threads=lambda _threads: True, wait_file_log_worker=lambda _thread: True,
            flush_log_queue=lambda _queue: calls.append('flush'))

    def test_shutdown_cloud_failure_still_closes_other_resources(self):
        calls, logs = [], []
        def fail(**_kwargs):
            raise OpaqueError()
        options = self.cleanup_options(calls)
        options['stop_cloud_control'] = fail
        self.assertEqual(cleanup_and_exit_runtime(**options, log_error=logs.append), 'exited')
        self.assertEqual(calls, ['serial', 'tray', 'flush', 'destroy'])
        self.assertEqual(logs, ['Stop cloud control during shutdown failed: OpaqueError'])

    def test_shutdown_worker_wait_failure_keeps_logger_and_root_alive(self):
        calls, logs = [], []
        stop = threading.Event()
        def fail(_threads):
            raise OpaqueError()
        options = self.cleanup_options(calls)
        options.update(wait_worker_threads=fail, file_log_stop_event=stop)
        self.assertEqual(cleanup_and_exit_runtime(**options, log_error=logs.append), 'worker_wait_failed')
        self.assertFalse(stop.is_set())
        self.assertEqual(calls, ['serial', 'tray'])
        self.assertEqual(logs, ['Wait for producer worker threads raised: OpaqueError'])

    def test_file_worker_keeps_valid_records_in_malformed_batch(self):
        for malformed, first in [('bad', False), ('bad', True), (([], 'bad\n'), False), (('unused.log', object()), False)]:
            with self.subTest(malformed=type(malformed).__name__, first=first), tempfile.TemporaryDirectory() as directory:
                path = Path(directory, 'valid.log')
                tasks, stop, logs = queue.Queue(), threading.Event(), []
                entries = [(path, 'one\n'), malformed, (path, 'two\n')]
                if first:
                    entries = [malformed, (path, 'one\n'), (path, 'two\n')]
                for entry in entries:
                    tasks.put(entry)
                def write(batches, on_error=None):
                    try:
                        return write_log_batches(batches, on_error=on_error)
                    finally:
                        stop.set()
                run_file_log_worker(log_queue=tasks, stop_event=stop, write_batches=write, on_error=logs.append)
                self.assertEqual(path.read_text(encoding='utf-8'), 'one\ntwo\n')
                self.assertEqual(tasks.unfinished_tasks, 0)
                self.assertTrue(tasks.empty())
                self.assertEqual(len(logs), 1)

    def test_file_write_error_reports_independent_notice_and_continues(self):
        notice, logs = FileLogErrorState(), []
        with tempfile.TemporaryDirectory() as directory:
            good = Path(directory, 'valid.log')
            def open_file(path, mode, encoding):
                if path == 'synthetic-broken':
                    raise OpaqueError()
                return open(path, mode, encoding=encoding)
            def error(message):
                logs.append(message)
                notice.report(message)
            self.assertEqual(write_log_batches({'synthetic-broken':['lost\n'], good:['kept\n']},
                open_file=open_file, on_error=error), 1)
            self.assertEqual(good.read_text(encoding='utf-8'), 'kept\n')
            self.assertIsNotNone(notice.take_notice(force=True))
            self.assertEqual(len(logs), 1)
            self.assertIn('OpaqueError', logs[0])

    def test_shutdown_flush_skips_only_malformed_records(self):
        with tempfile.TemporaryDirectory() as directory:
            path, tasks, logs = Path(directory, 'valid.log'), queue.Queue(), []
            for entry in [(path, 'one\n'), ([], 'bad\n'), (path, 'two\n')]:
                tasks.put(entry)
            self.assertEqual(flush_log_queue(tasks, log_error=logs.append), 2)
            self.assertEqual(path.read_text(encoding='utf-8'), 'one\ntwo\n')
            self.assertEqual(tasks.unfinished_tasks, 0)
            self.assertEqual(len(logs), 1)

    def test_config_backup_error_preserves_original_and_expected_error(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory, 'config.ini')
            original = b'not-an-ini-header\npassword=synthetic-private-value\n'
            path.write_bytes(original)
            config, saves, logs = configparser.ConfigParser(interpolation=None), [], []
            def reject_backup(*_args):
                raise OpaqueError()
            with self.assertRaises(ConfigInitializationError):
                initialize_config_runtime(config=config, config_file=path,
                    defaults_by_section={'ui':{'voice_enabled':'1'}}, save_config=lambda: saves.append(True),
                    backup_file=reject_backup, log_error=logs.append)
            self.assertEqual(path.read_bytes(), original)
            self.assertEqual(config.sections(), [])
            self.assertEqual(saves, [])
            self.assertIn('OpaqueError', '\n'.join(logs))

    def test_invalid_typed_config_uses_default_without_echoing_value(self):
        private = 'password=synthetic-private-value'
        config = configparser.ConfigParser(interpolation=None)
        config['ui'] = {'voice_enabled':private}
        logs = []
        result = read_startup_config_values(config, default_voice_text='safe text', log_error=logs.append)
        self.assertTrue(result.voice_enabled)
        self.assertNotIn(private, '\n'.join(logs))
        self.assertIn('ValueError', '\n'.join(logs))
