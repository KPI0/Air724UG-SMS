import queue
from pathlib import Path
import tempfile
import threading
import tkinter as tk
from types import SimpleNamespace
import unittest

from sms_core.app_launch import restart_software_runtime
from sms_core.app_shutdown import safe_set_events
from sms_core.third_push_runtime import third_push_worker_runtime
from sms_ui.app_restart_runtime import start_restart_software_app_runtime
from sms_ui.config_sync_runtime import ConfigFileWatchState
from sms_ui.config_sync_namespace_runtime import start_config_file_watch_namespace_runtime
from sms_ui.firmware_update_window import open_firmware_update_window
from sms_ui.maintenance_runtime import AutoLogCleanupState, run_auto_log_cleanup_tick_runtime
from sms_core.threading_runtime import WorkerThreadRegistry, start_daemon_thread
from tests import test_app_launch


class RestartOrderTests(unittest.TestCase):
    def test_receiver_and_cloud_ack_finish_before_cloud_close(self):
        calls = []
        options = test_app_launch.AppLaunchRestartTests()._runtime_kwargs(calls)
        options.update(
            pre_cloud_worker_threads=lambda: ('receiver',),
            before_stop_cloud=lambda: calls.append(('ack',)),
        )
        result = restart_software_runtime(**options)
        self.assertEqual(result.status, 'exited')
        receiver = next(i for i, item in enumerate(calls) if item[:2] == ('wait_workers', ('receiver',)))
        ack = calls.index(('ack',))
        cloud = next(i for i, item in enumerate(calls) if item[0] == 'cloud')
        self.assertLess(receiver, ack)
        self.assertLess(ack, cloud)

    def test_receiver_failure_cancels_helper_without_closing_cloud_or_exiting(self):
        calls = []
        options = test_app_launch.AppLaunchRestartTests()._runtime_kwargs(calls)
        options.update(pre_cloud_worker_threads=('receiver',), wait_worker_threads=lambda *a, **k: False)
        result = restart_software_runtime(**options)
        self.assertEqual(result.status, 'worker_wait_failed')
        self.assertIn(('cancel', 'helper-process'), calls)
        self.assertFalse(any(item[0] in ('cloud', 'flush', 'release', 'exit') for item in calls))


class RestartUiIntegrationTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(str(exc))
        self.root.withdraw()
        self.owner = threading.get_ident()
        self.state = {'exiting': False, 'running': True}
        self.shutdown = threading.Event()
        self.calls, self.errors = [], []
        self.messagebox = SimpleNamespace(
            askyesno=lambda *a, **k: True,
            showerror=lambda *a, **k: (self.errors.append((threading.get_ident(), a)), self.root.quit()),
        )

    def tearDown(self):
        self.root.destroy()

    def runtime(self, **kwargs):
        return restart_software_runtime(
            **kwargs, prepare_launch=lambda *a: (['synthetic-helper'], '.'),
            launch_process=lambda *a, **k: self.calls.append(('launch', threading.get_ident())) or object(),
            cancel_launch=lambda helper: self.calls.append(('cancel',)) or True,
            clean_env=lambda: {},
        )

    def options(self):
        return dict(
            root=self.root, messagebox=self.messagebox, is_exiting=self.state['exiting'],
            set_exiting=lambda value: self.state.update(exiting=value),
            log_error=lambda message: self.calls.append(('log', message)),
            set_serial_running=lambda value: self.state.update(running=value),
            autostart_flag='--autostart', restart_helper_flag='--restart-helper',
            system_ui=lambda *a: None, stop_tray_icon=lambda **k: None,
            safe_set_events=safe_set_events, stop_events=(self.shutdown,),
            stop_cloud_control=lambda **k: self.calls.append(('cloud',)), safe_close_serial=lambda: None,
            app_mutex=None, release_mutex=lambda *a: None,
            flush_log_queue=lambda q: self.calls.append(('flush',)), file_log_queue=queue.Queue(),
            exit_process=lambda code: (self.calls.append(('exit', threading.get_ident())), self.root.quit()),
            argv=['synthetic'], current_pid=123, restart_runtime=self.runtime,
        )

    def run_ui(self):
        watchdog = self.root.after(3000, self.root.quit)
        try:
            self.root.mainloop()
        finally:
            self.root.after_cancel(watchdog)

    def test_restart_keeps_tk_responsive_until_pushes_and_logs_finish(self):
        release, stop = threading.Event(), threading.Event()
        pending = queue.Queue()
        for message in ('first', 'last'):
            pending.put({'message': message, 'channels': ['custom_post']})
        sent = []

        def send(channel, body, settings):
            if not release.wait(2):
                raise AssertionError('Tk did not release the sender')
            sent.append(body)
            return True, 'ok'

        worker = threading.Thread(target=third_push_worker_runtime, kwargs=dict(
            stop_event=stop, push_queue=pending, send_channel_func=send,
            system_ui=lambda *a: None, show_result=lambda *a: None,
            format_message_func=lambda body, *a, **k: body,
        ), daemon=True)
        worker.start()
        options = self.options()
        options.update(deferred_stop_events=(stop,), deferred_worker_threads=(worker,), deferred_worker_queues=(pending,))
        ticks = []
        timers = []
        try:
            self.assertEqual(start_restart_software_app_runtime(**options), 'started')
            self.assertEqual(start_restart_software_app_runtime(**self.options()), 'already_exiting')
            timers = [self.root.after(delay, lambda: ticks.append(True)) for delay in (10, 20, 30)]
            timers.append(self.root.after(80, release.set))
            self.run_ui()
        finally:
            release.set()
            stop.set()
            worker.join(2)
            for timer in timers:
                self.root.after_cancel(timer)
        self.assertEqual(sent, ['first', 'last'])
        self.assertEqual(len(ticks), 3)
        self.assertEqual(self.calls[-2:], [('flush',), ('exit', self.owner)])
        self.assertNotEqual(self.calls[0][1], self.owner)
        self.assertEqual(self.errors, [])
        self.assertEqual(pending.unfinished_tasks, 0)

    def test_launch_failure_keeps_receiver_active_and_reports_on_ui(self):
        def runtime(**kwargs):
            return restart_software_runtime(
                **kwargs, prepare_launch=lambda *a: (['synthetic'], '.'),
                launch_process=lambda *a, **k: (_ for _ in ()).throw(OSError('synthetic')),
            )
        options = self.options()
        options['restart_runtime'] = runtime
        self.assertEqual(start_restart_software_app_runtime(**options), 'started')
        self.run_ui()
        self.assertTrue(self.state['running'])
        self.assertFalse(self.state['exiting'])
        self.assertFalse(self.shutdown.is_set())
        self.assertEqual(self.errors[0][0], self.owner)
        self.assertIn('当前软件将继续运行', self.errors[0][1][1])

    def test_worker_failure_cancels_helper_and_allows_another_exit_attempt(self):
        options = self.options()
        options['pre_cloud_worker_threads'] = lambda: (_ for _ in ()).throw(RuntimeError('synthetic'))
        self.assertEqual(start_restart_software_app_runtime(**options), 'started')
        self.run_ui()
        self.assertIn(('cancel',), self.calls)
        self.assertFalse(self.state['exiting'])
        self.assertEqual(self.errors[0][0], self.owner)
        self.assertFalse(any(item[0] == 'exit' for item in self.calls))

    def _check_cleanup_survives_launch_failure(self, already_running):
        cleanup_state = AutoLogCleanupState()
        registry = WorkerThreadRegistry()
        release_helper, release_cleanup = threading.Event(), threading.Event()
        shutdown_threads, notices, timers = [], [], []

        def cleanup(days):
            if already_running and not release_cleanup.wait(2):
                raise RuntimeError('Test did not release cleanup')
            return 1

        def completed(*args):
            notices.append((threading.get_ident(), args))
            if not self.state['exiting']:
                self.root.quit()

        def tick():
            run_auto_log_cleanup_tick_runtime(
                state=cleanup_state, is_enabled=lambda: True, retention_days=lambda: 7,
                interval_hours=lambda: 1, cleanup_old_logs=cleanup, system_ui=completed,
                tk_alive=lambda: not self.shutdown.is_set(), root_after=self.root.after,
                tick_callback=tick, ui_post=lambda fn: False, thread_registry=registry,
                is_stopping=lambda: self.state['exiting'] or self.shutdown.is_set(),
            )

        def launch(*args, **kwargs):
            if not release_helper.wait(2):
                raise RuntimeError('Test did not release restart helper')
            raise OSError('synthetic helper failure')

        def runtime(**kwargs):
            return restart_software_runtime(
                **kwargs, prepare_launch=lambda *args: (['synthetic'], '.'),
                launch_process=launch, clean_env=lambda: {},
            )

        def start_worker(*args, **kwargs):
            thread = start_daemon_thread(*args, **kwargs)
            shutdown_threads.append(thread)
            return thread

        self.messagebox.showerror = lambda *a, **k: self.errors.append((threading.get_ident(), a))
        options = self.options()
        options.update(restart_runtime=runtime, start_worker=start_worker)
        try:
            if already_running:
                tick()
            self.assertEqual(start_restart_software_app_runtime(**options), 'started')
            if not already_running:
                timers.append(self.root.after(10, tick))
            timers.append(self.root.after(30, release_cleanup.set))
            timers.append(self.root.after(220, release_helper.set))
            self.run_ui()
            self.assertTrue(self.state['running'])
            self.assertFalse(self.state['exiting'])
            self.assertFalse(self.shutdown.is_set())
            self.assertEqual(len(self.errors), 1)
            self.assertEqual(len(notices), 1)
            self.assertEqual(notices[0][0], self.owner)
            self.assertFalse(cleanup_state.running)
            self.assertIsNotNone(cleanup_state.after_id)
            self.assertFalse(registry.snapshot())
        finally:
            release_helper.set()
            release_cleanup.set()
            for thread in tuple(shutdown_threads) + registry.snapshot():
                thread.join(3)
            for timer in timers + [cleanup_state.after_id]:
                if timer is not None:
                    self.root.after_cancel(timer)

    def test_launch_failure_resumes_cleanup_that_was_already_running(self):
        self._check_cleanup_survives_launch_failure(True)

    def test_launch_failure_resumes_cleanup_due_during_restart_preparation(self):
        self._check_cleanup_survives_launch_failure(False)

    def test_launch_failure_resumes_config_and_existing_firmware_window(self):
        release = threading.Event()
        observations, changes, workers = [], [], []
        state = dict(device='synthetic', current_version='1', target_version='2', filename='',
                     busy=False, progress=0, message='initial', can_start=False,
                     can_cancel=False, config_changed=False)

        def snapshot():
            observations.append(state['message'])
            return dict(state)

        def launch(*args, **kwargs):
            if not release.wait(2):
                raise RuntimeError('Test did not release helper')
            raise OSError('synthetic launch failure')

        def runtime(**kwargs):
            return restart_software_runtime(**kwargs, prepare_launch=lambda *a: (['synthetic'], '.'),
                                           launch_process=launch, clean_env=lambda: {})

        def start_worker(*args, **kwargs):
            worker = start_daemon_thread(*args, **kwargs)
            workers.append(worker)
            return worker

        with tempfile.TemporaryDirectory() as folder:
            config = Path(folder) / 'config.ini'
            config.write_text('[test]\nvalue=initial\n', encoding='utf-8')
            namespace = dict(root=self.root, is_exiting=False, TK_SHUTDOWN=self.shutdown,
                CONFIG_FILE=str(config), CONFIG_FILE_WATCH_STATE=ConfigFileWatchState(),
                tk_alive=lambda: not self.shutdown.is_set(),
                reload_shared_ui_config=lambda: changes.append(config.read_text(encoding='utf-8')),
                local_firmware_updater=SimpleNamespace(snapshot=snapshot, prepare=lambda *a: None,
                                                      start=lambda: None, cancel=lambda: None))
            win = open_firmware_update_window(namespace)
            win.withdraw()
            start_config_file_watch_namespace_runtime(namespace, interval_ms=50)

            def set_exiting(value):
                namespace['is_exiting'] = value
                self.state['exiting'] = value

            self.messagebox.showerror = lambda *a, **k: self.errors.append(a)
            options = self.options()
            options.update(set_exiting=set_exiting, restart_runtime=runtime, start_worker=start_worker)

            def resumed():
                state['message'] = 'recovered'
                config.write_text('[test]\nvalue=recovered\n', encoding='utf-8')

            timers = [self.root.after(20, lambda: config.write_text('[test]\nvalue=before\n', encoding='utf-8')),
                      self.root.after(150, lambda: start_restart_software_app_runtime(**options)),
                      self.root.after(400, release.set), self.root.after(550, resumed),
                      self.root.after(950, self.root.quit)]
            try:
                self.run_ui()
                self.assertEqual(len(self.errors), 1)
                self.assertFalse(namespace['is_exiting'])
                self.assertFalse(self.shutdown.is_set())
                self.assertEqual(changes, ['[test]\nvalue=before\n', '[test]\nvalue=recovered\n'])
                self.assertIn('recovered', observations)
                self.assertIsNotNone(namespace['CONFIG_FILE_WATCH_STATE'].after_id)
                self.assertIs(namespace['firmware_update_window'], win)
            finally:
                release.set()
                for worker in workers:
                    worker.join(3)
                for timer in timers + [namespace['CONFIG_FILE_WATCH_STATE'].after_id]:
                    if timer is not None:
                        self.root.after_cancel(timer)
