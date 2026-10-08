import queue
import threading
import time
import unittest
from pathlib import Path
from unittest.mock import patch

from sms_core import config_runtime
from sms_core.cloud_runtime import read_cloud_control_settings
from sms_core.serial_namespace_runtime import try_rebind_manual_port_namespace_runtime
from sms_core.third_push_config import read_third_push_settings
from sms_ui.app_infrastructure_namespace_bindings import install_app_infrastructure_namespace_bindings
from sms_ui.cloud_control_app_runtime import save_cloud_control_setting_runtime
from sms_ui.config_save_runtime import run_config_transaction_for_ui
from sms_ui.desktop_shortcut_runtime import save_desktop_shortcut_name_runtime
from sms_ui.maintenance_runtime import apply_log_cleanup_runtime, save_update_proxy_config
from sms_ui.settings_runtime import (
    save_cloud_sensitive_commands_config, save_keywords_config, save_ui_config_values,
)
from sms_ui.third_push_app_runtime import save_third_push_setting_runtime
from sms_ui.thread_runtime import ui_pump_runtime
from sms_ui.ui_log_namespace_runtime import ui_post_namespace_runtime
from tests import test_serial_config_concurrency as serial_quality


class ConfigTransactionDispatchTests(unittest.TestCase):
    def setUp(self):
        self.fixture = serial_quality.SerialConfigConcurrencyTests()
        self.fixture.setUp()
        self.addCleanup(self.fixture.doCleanups)
        self.addCleanup(self.fixture.tearDown)

    def namespace(self):
        ns = self.fixture.namespace()
        tasks = queue.Queue()
        ns.update(ui_post=lambda callback: tasks.put_nowait(callback),
                  TK_SHUTDOWN=threading.Event(), serial_stop_event=threading.Event())
        install_app_infrastructure_namespace_bindings(ns)
        read_third_push_settings(ns['config'])
        self.assertTrue(ns['safe_save_config']())
        return ns, tasks

    def start_rebind(self, ns):
        results = []
        worker = threading.Thread(target=lambda: results.append(try_rebind_manual_port_namespace_runtime(ns)))
        worker.start()
        self.addCleanup(worker.join, 3)
        self.addCleanup(ns['serial_stop_event'].set)
        return worker, results

    def settle(self, worker, results, expected):
        worker.join(3)
        self.assertFalse(worker.is_alive())
        self.assertEqual(results, [expected])
        self.assertEqual(self.fixture.fixture.errors, [])
        root = self.fixture.root
        self.assertFalse(root.tk.call('after', 'info'))
        self.assertFalse(root.tk.globalgetvar('background_errors'))

    def apply_setting(self, kind, ns):
        config, save = ns['config'], ns['safe_save_config']
        if kind == 'voice':
            return save_ui_config_values(config, {'voice_enabled': '0'}, save)
        if kind == 'font':
            return save_ui_config_values(config, {'sms_font_size': '20', 'sms_font_color': '#123456'}, save)
        if kind == 'keywords':
            return save_keywords_config(config, ['synthetic'], save)
        if kind == 'security':
            return save_cloud_sensitive_commands_config(config, {}, save)
        if kind == 'third_push':
            return save_third_push_setting_runtime(config=config, save_config=save,
                current_settings=lambda: read_third_push_settings(config), apply_settings=lambda _: None,
                sms_enabled=False) is not None
        if kind == 'cloud':
            return save_cloud_control_setting_runtime(config=config, save_config=save,
                current_settings=lambda: read_cloud_control_settings(config), apply_settings=lambda _: None,
                system_ui=lambda *_: None, reconnect_interval=17) is not None
        if kind == 'cleanup':
            return apply_log_cleanup_runtime(19, config=config, save_config=save,
                set_cleanup_state=lambda *_: None, schedule_cleanup=lambda **_: None,
                system_ui=lambda *_: None, interval_hours=12)
        if kind == 'proxy':
            save_update_proxy_config(config, 'https://example.invalid/', 'https://example.invalid/', save)
            return True
        return save_desktop_shortcut_name_runtime(config, 'Synthetic client', save)

    def test_all_settings_and_rebind_outcomes_preserve_committed_values(self):
        kinds = ('voice', 'font', 'keywords', 'security', 'third_push', 'cloud', 'cleanup', 'proxy', 'shortcut')
        for kind in kinds:
            for before_commit in (True, False):
                for ui_fails in (False, True):
                    for rebind_fails in (False, True):
                        with self.subTest(kind=kind, before=before_commit, ui_fails=ui_fails, rebind_fails=rebind_fails):
                            ns, tasks = self.namespace()
                            stage = ['rebind']
                            mutations, writes = [], []
                            old_restore = config_runtime.restore_config_runtime

                            def restore(config, snapshot):
                                if config is ns['config']:
                                    mutations.append(threading.current_thread() is threading.main_thread())
                                return old_restore(config, snapshot)

                            def save(**options):
                                fail = rebind_fails if stage[0] == 'rebind' else ui_fails

                                def replace(source, target):
                                    writes.append(threading.current_thread() is threading.main_thread())
                                    if fail:
                                        raise PermissionError('synthetic write failure')
                                    Path(source).replace(target)

                                return config_runtime.safe_save_config_runtime(**options, replace_file=replace)

                            def apply():
                                stage[0] = 'ui'
                                original = config_runtime.snapshot_config_runtime(ns['config'])
                                try:
                                    saved = self.apply_setting(kind, ns)
                                except RuntimeError:
                                    saved = False
                                self.assertEqual(saved, not ui_fails)
                                if ui_fails:
                                    self.assertEqual(config_runtime.snapshot_config_runtime(ns['config']), original)
                                else:
                                    self.assertNotEqual(config_runtime.snapshot_config_runtime(ns['config']), original)
                                stage[0] = 'rebind'

                            with patch.object(config_runtime, 'restore_config_runtime', new=restore), patch(
                                'sms_ui.app_infrastructure_namespace_runtime.safe_save_config_runtime', new=save
                            ):
                                worker, results = self.start_rebind(ns)
                                callback = tasks.get(timeout=2)
                                if before_commit:
                                    apply()
                                ui_values = config_runtime.snapshot_config_runtime(ns['config'])
                                callback()
                                if before_commit:
                                    for section, values in ui_values.items():
                                        if section != 'serial':
                                            self.assertEqual(dict(ns['config'][section]), values)
                                else:
                                    apply()
                            self.settle(worker, results, not rebind_fails)
                            if not (ui_fails and rebind_fails):
                                self.assertTrue(mutations)
                            self.assertTrue(all(mutations), 'Shared configuration changed outside UI thread')
                            self.assertTrue(writes)
                            self.assertFalse(any(writes), 'File replacement blocked the UI thread')
                            self.assertEqual(config_runtime.load_config_snapshot(ns['CONFIG_FILE']),
                                             config_runtime.snapshot_config_runtime(ns['config']))
                            self.assertEqual(ns['PORT'], 'COM901' if rebind_fails else 'COM902')
                            self.assertEqual(ns['config']['serial']['port'], ns['PORT'])
                            self.assertEqual(ns['config']['serial']['unknown'], 'keep')

    def test_queued_scan_rechecks_ui_choice_before_commit(self):
        for mode, port in (('Manual', 'COM903'), ('Auto', '')):
            with self.subTest(mode=mode):
                ns, tasks = self.namespace()
                worker, results = self.start_rebind(ns)
                callback = tasks.get(timeout=2)
                ns['safe_close_serial'] = lambda: None
                self.assertTrue(self.fixture.apply(ns, mode, port, 57600))
                callback()
                self.settle(worker, results, False)
                self.fixture.assert_consistent(ns, mode, port, 57600)

    def test_busy_save_rejects_queued_rebind_without_touching_draft(self):
        ns, tasks = self.namespace()
        worker, results = self.start_rebind(ns)
        callback = tasks.get(timeout=2)
        invoked = []

        def during_write(**options):
            self.fixture.root.after(1, lambda: (invoked.append(True), callback()))

            def replace(source, target):
                time.sleep(0.10)
                Path(source).replace(target)

            return config_runtime.safe_save_config_runtime(**options, replace_file=replace)

        with patch('sms_ui.app_infrastructure_namespace_runtime.safe_save_config_runtime', new=during_write):
            self.assertTrue(self.apply_setting('voice', ns))
        self.assertEqual(invoked, [True])
        self.settle(worker, results, False)
        self.assertEqual(ns['config']['ui']['voice_enabled'], '0')
        worker, results = self.start_rebind(ns)
        tasks.get(timeout=2)()
        self.settle(worker, results, True)
        self.assertEqual(config_runtime.load_config_snapshot(ns['CONFIG_FILE'])['ui']['voice_enabled'], '0')

    def test_cancelled_or_timed_out_queue_cannot_commit_later(self):
        for reason in ('shutdown', 'serial_stop', 'timeout'):
            with self.subTest(reason=reason):
                ns, tasks = self.namespace()
                if reason == 'timeout':
                    ns['run_config_transaction'] = lambda callback: run_config_transaction_for_ui(
                        ns, callback, queue_timeout=0.02)
                original = config_runtime.load_config_snapshot(ns['CONFIG_FILE'])
                worker, results = self.start_rebind(ns)
                callback = tasks.get(timeout=2)
                if reason != 'timeout':
                    ns['TK_SHUTDOWN' if reason == 'shutdown' else 'serial_stop_event'].set()
                self.settle(worker, results, False)
                callback()
                self.assertEqual(config_runtime.snapshot_config_runtime(ns['config']), original)
                self.assertEqual(config_runtime.load_config_snapshot(ns['CONFIG_FILE']), original)
                self.assertEqual(ns['PORT'], 'COM901')

    def test_full_ui_queue_does_not_run_transaction_on_worker(self):
        ns, _ = self.namespace()
        ns['ui_post'] = lambda _: False
        worker, results = self.start_rebind(ns)
        self.settle(worker, results, False)
        self.assertEqual(ns['PORT'], 'COM901')

    def test_shutdown_during_started_write_waits_for_commit_and_keeps_tk_responsive(self):
        ns, tasks = self.namespace()
        worker, results = self.start_rebind(ns)
        callback = tasks.get(timeout=2)
        observed = []

        def stop():
            ns['TK_SHUTDOWN'].set()
            observed.append(worker.is_alive())

        def save(**options):
            self.fixture.root.after(20, stop)

            def replace(source, target):
                time.sleep(0.15)
                Path(source).replace(target)

            return config_runtime.safe_save_config_runtime(**options, replace_file=replace)

        with patch('sms_ui.app_infrastructure_namespace_runtime.safe_save_config_runtime', new=save):
            callback()
        self.settle(worker, results, True)
        self.assertEqual(observed, [True])
        self.assertEqual(config_runtime.load_config_snapshot(ns['CONFIG_FILE'])['serial']['port'], 'COM902')
        self.assertEqual(ns['PORT'], 'COM902')

    def test_production_ui_queue_pumps_rebind_and_following_tasks(self):
        ns, _ = self.namespace()
        ns['UI_TASK_QUEUE'] = queue.Queue()
        queued = threading.Event()

        def post(callback):
            result = ui_post_namespace_runtime(ns, callback)
            queued.set()
            return result

        ns['ui_post'] = post
        worker, results = self.start_rebind(ns)
        self.assertTrue(queued.wait(2))
        following = []
        ui_post_namespace_runtime(ns, lambda: following.append(ns['PORT']))
        processed = ui_pump_runtime(ns['UI_TASK_QUEUE'], self.fixture.root,
            lambda: False, lambda: None)
        self.assertEqual(processed, 2)
        self.assertEqual(following, ['COM902'])
        self.assertEqual(ns['UI_TASK_QUEUE'].unfinished_tasks, 0)
        self.settle(worker, results, True)

    def test_failed_transaction_releases_waiter_even_if_logging_fails(self):
        ns, tasks = self.namespace()
        results = []

        def fail(*_):
            raise RuntimeError('synthetic callback failure')

        ns['log_file_only'] = fail
        worker = threading.Thread(target=lambda: results.append(run_config_transaction_for_ui(ns, fail)))
        worker.start()
        self.addCleanup(worker.join, 3)
        self.addCleanup(ns['serial_stop_event'].set)
        tasks.get(timeout=2)()
        self.settle(worker, results, False)


if __name__ == '__main__':
    unittest.main()
