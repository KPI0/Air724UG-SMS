import asyncio
import json
import os
from pathlib import Path
import queue
import smtplib
import sys
import tempfile
import threading
import unittest
from unittest.mock import AsyncMock, patch

from sms_core import third_push_sender as sender
from sms_core.cloud_auth import secret_match_result
from sms_core.cloud_message_runtime import cloud_session_revoke_proof, send_cloud_register_runtime
from sms_ui.maintenance_runtime import (
    AutoLogCleanupState, run_auto_log_cleanup_tick_runtime, schedule_auto_log_cleanup_runtime,
)
from sms_core.file_log_runtime import FileLogErrorState, FileLogQueue, start_file_log_worker
from sms_core.app_shutdown import flush_log_queue
from sms_core.threading_runtime import WorkerThreadRegistry
from sms_ui.ui_log_namespace_bindings import install_ui_log_namespace_bindings
from sms_ui.ui_log_namespace_runtime import notify_file_log_errors_namespace_runtime
from sms_ui.app_lifecycle_namespace_runtime import _flush_file_logs, _registered_worker_threads
from sms_ui.app_shutdown_runtime import start_shutdown_task_runtime


class RecipientResultTests(unittest.TestCase):
    def test_smtp_checks_actual_stdlib_recipient_results(self):
        settings = dict(email_smtp_host='smtp.example.test', email_smtp_port='465',
                        email_encryption='ssl', email_username='sender@example.test',
                        email_password='test-password', email_from_address='sender@example.test',
                        email_to_addresses='one@example.test;two@example.test')
        for refused_count in (0, 1, 2):
            with self.subTest(refused_count=refused_count):
                class MemorySMTP(smtplib.SMTP):
                    def __init__(self, *args, **kwargs):
                        super().__init__()
                        self.does_esmtp = False
                        self.data_calls = 0
                        self.rcpt_calls = 0

                    def ehlo_or_helo_if_needed(self):
                        pass

                    def login(self, *args):
                        return 235, b'OK'

                    def mail(self, *args):
                        return 250, b'OK'

                    def rcpt(self, address, *args):
                        self.rcpt_calls += 1
                        return (550, b'private remote detail') if self.rcpt_calls <= refused_count else (250, b'OK')

                    def data(self, message):
                        self.data_calls += 1
                        return 250, b'OK'

                    def _rset(self):
                        pass

                    def quit(self):
                        pass

                smtp = MemorySMTP()
                with patch.object(sender.smtplib, 'SMTP_SSL', return_value=smtp):
                    ok, info = sender.send_email('test content', settings)
                self.assertEqual(ok, refused_count == 0)
                self.assertEqual(smtp.data_calls, int(refused_count < 2))
                self.assertEqual(smtp.rcpt_calls, 2)
                if refused_count:
                    self.assertIn(f'拒收 {refused_count}', info)
                for private in ('one@', 'two@', 'private remote detail', 'test-password'):
                    self.assertNotIn(private, info)

    def test_wxpusher_checks_every_recipient(self):
        for codes in ((1000, 1000), (1000, 1001), (1001, 1001)):
            with self.subTest(codes=codes):
                body = json.dumps({'code': 1000, 'data': [
                    {'uid': 'UID_one', 'code': codes[0], 'status': 'private detail'},
                    {'uid': 'UID_two', 'code': codes[1]},
                ]})
                with patch.object(sender, 'http_request', return_value=(True, 200, body)) as request:
                    ok, info = sender.send_wxpusher('text', {
                        'wxpusher_app_token': 'AT_test', 'wxpusher_uids': 'UID_one UID_two'})
                self.assertEqual(ok, all(code == 1000 for code in codes))
                self.assertEqual(request.call_count, 1)
                if not ok:
                    self.assertIn(f'失败 {codes.count(1001)}', info)
                self.assertNotIn('UID_', info)
                self.assertNotIn('private detail', info)

    def test_wxpusher_does_not_confirm_missing_or_malformed_results(self):
        for results in (None, [], {}, [None], [{'uid': 'UID_one'}]):
            with self.subTest(results=results):
                ok, _ = sender.api_ok('wxpusher', True, 200, json.dumps({'code': 1000, 'data': results}))
                self.assertFalse(ok)

    def test_wxpusher_does_not_confirm_incomplete_or_duplicate_recipient_list(self):
        for uids in (['UID_one'], ['UID_one', 'UID_one'], ['UID_one', 'UID_other']):
            with self.subTest(uids=uids):
                body = json.dumps({'code': 1000, 'data': [{'uid': uid, 'code': 1000} for uid in uids]})
                with patch.object(sender, 'http_request', return_value=(True, 200, body)):
                    ok, _ = sender.send_wxpusher('text', {'wxpusher_app_token': 'AT_test', 'wxpusher_uids': 'UID_one UID_two'})
                self.assertFalse(ok)


class UnicodeCloudPasswordTests(unittest.TestCase):
    def test_auth_supports_unicode_and_preserves_exact_comparison(self):
        for expected, incoming, matched in (
            ('中文密码', ' 中文密码 ', True), ('café🔐', 'café🔐', True),
            ('中文密码', '中文密碼', False), ('ascii', '中文', False),
            ('中文', 'ascii', False), ('é', 'e\u0301', False),
            ('ascii', '\ud800', False),
        ):
            with self.subTest(expected=expected, incoming=incoming):
                ok, reason = secret_match_result({'secret': incoming}, expected)
                self.assertEqual(ok, matched)
                self.assertNotIn(expected, reason)

    def test_register_unicode_password_rotation_uses_previous_hmac_proof(self):
        for old, new in (('ascii', '中文'), ('中文', 'café🔐'), ('中文', '中文'), (' 中文 ', '中文'), ('', '中文')):
            with self.subTest(old=old, new=new):
                ws = AsyncMock()
                result = asyncio.run(send_cloud_register_runtime(
                    ws, auto_upload=True, build_payload=lambda *_: {'secret': new},
                    timestamp=lambda: 0, identity_payload=lambda: {}, secret=new,
                    serial_port='', serial_baud=115200, serial_mode='Auto',
                    runtime_imei=lambda: '123456789012345', log=lambda *a, **k: None,
                    previous_session_secret=old,
                ))
                self.assertEqual(result, 'sent')
                payload = json.loads(ws.send.call_args.args[0])
                self.assertEqual(payload.get('previous_session_proof'),
                                 cloud_session_revoke_proof(old, '123456789012345')
                                 if old.strip() and old.strip() != new.strip() else None)


class CleanupBackgroundTests(unittest.TestCase):
    def setUp(self):
        self.state = AutoLogCleanupState()
        self.registry = WorkerThreadRegistry()
        self.threads = []
        self.timers = {}
        self.next_timer = 0
        self.messages = []
        self.cleaned = []
        self.stopping = False
        self.tk_alive = True
        self.enabled = True
        self.days = 7

        owner = self
        class DeferredThread:
            def __init__(self, target, **kwargs):
                self.target = target
                owner.threads.append(self)
            def start(self):
                pass
        self.thread_factory = DeferredThread

    def after(self, delay, callback):
        self.next_timer += 1
        self.timers[self.next_timer] = (delay, callback)
        return self.next_timer

    def tick(self):
        return run_auto_log_cleanup_tick_runtime(
            state=self.state, is_enabled=lambda: self.enabled, retention_days=lambda: self.days,
            interval_hours=lambda: 1, cleanup_old_logs=lambda days: self.cleaned.append(days) or 3,
            system_ui=lambda *args: self.messages.append(args), tk_alive=lambda: self.tk_alive,
            root_after=self.after, tick_callback=self.tick, ui_post=lambda fn: False,
            is_stopping=lambda: self.stopping, thread_registry=self.registry,
            thread_factory=self.thread_factory,
        )

    def schedule(self, restart=True):
        return schedule_auto_log_cleanup_runtime(
            state=self.state, restart=restart, first_delay_sec=5,
            is_enabled=lambda: self.enabled, tk_alive=lambda: self.tk_alive, root_after=self.after,
            root_after_cancel=self.timers.pop, tick_callback=self.tick,
            ui_post=lambda fn: False, is_stopping=lambda: self.stopping,
        )

    def poll(self):
        _, callback = self.timers.pop(self.state.after_id)
        callback()

    def test_repeated_ticks_and_settings_changes_keep_one_worker_and_one_timer(self):
        self.tick()
        self.tick()
        self.days = 30
        self.schedule()
        self.schedule()
        self.assertEqual(len(self.threads), 1)
        self.assertEqual(len(self.timers), 1)
        self.assertEqual(self.registry.snapshot(), tuple(self.threads))
        self.threads[0].target()
        self.assertEqual(self.cleaned, [7])
        self.assertFalse(self.registry.snapshot())
        self.assertFalse(self.messages)
        self.poll()
        self.assertEqual(len(self.messages), 1)
        self.assertEqual(len(self.timers), 1)
        self.assertEqual(self.timers[self.state.after_id][0], 5000)
        self.poll()
        self.threads[1].target()
        self.poll()
        self.assertEqual(self.cleaned, [7, 30])

    def test_completion_survives_full_general_ui_queue(self):
        self.tick()
        self.threads[0].target()
        self.poll()
        self.assertFalse(self.state.running)
        self.assertEqual(len(self.messages), 1)
        self.assertEqual(self.timers[self.state.after_id][0], 3600000)

    def test_disable_while_running_does_not_reschedule(self):
        self.tick()
        self.enabled = False
        self.schedule()
        self.threads[0].target()
        self.poll()
        self.assertFalse(self.timers)
        self.assertIsNone(self.state.after_id)

    def test_shutdown_skips_unstarted_work_and_ui_completion(self):
        self.tick()
        self.stopping = True
        self.tk_alive = False
        self.threads[0].target()
        self.poll()
        self.assertFalse(self.cleaned)
        self.assertFalse(self.messages)
        self.assertFalse(self.timers)
        self.assertFalse(self.registry.snapshot())

    def test_stopping_does_not_start_or_schedule_cleanup(self):
        self.stopping = True
        self.tk_alive = False
        self.tick()
        self.schedule()
        self.assertFalse(self.threads)
        self.assertFalse(self.timers)

    def test_restart_preparation_defers_tick_until_launch_failure_recovers(self):
        self.stopping = True
        self.tick()
        self.assertFalse(self.threads)
        self.assertEqual(len(self.timers), 1)
        self.poll()
        self.assertFalse(self.threads)
        self.assertEqual(len(self.timers), 1)
        self.stopping = False
        self.poll()
        self.threads[0].target()
        self.poll()
        self.assertEqual(self.cleaned, [7])
        self.assertEqual(len(self.messages), 1)
        self.assertEqual(self.timers[self.state.after_id][0], 3600000)

    def test_restart_preparation_retains_completed_result_until_recovered(self):
        self.tick()
        self.threads[0].target()
        self.stopping = True
        self.poll()
        self.assertEqual(len(self.timers), 1)
        self.assertEqual(self.state.results.qsize(), 1)
        self.assertFalse(self.messages)
        self.stopping = False
        self.poll()
        self.assertFalse(self.state.running)
        self.assertEqual(self.state.results.qsize(), 0)
        self.assertEqual(len(self.messages), 1)
        self.assertEqual(self.timers[self.state.after_id][0], 3600000)

    def test_restart_preparation_poll_stops_when_shutdown_is_committed(self):
        self.stopping = True
        self.tick()
        self.assertEqual(len(self.timers), 1)
        self.tk_alive = False
        self.poll()
        self.assertFalse(self.timers)
        self.assertFalse(self.threads)

    def test_schedule_without_restart_preserves_existing_timer(self):
        self.schedule()
        timer = self.state.after_id
        self.schedule(restart=False)
        self.assertEqual(self.state.after_id, timer)
        self.assertEqual(len(self.timers), 1)

    def test_maintenance_registry_is_included_in_shutdown_snapshot(self):
        self.tick()
        self.assertEqual(_registered_worker_threads({'MAINTENANCE_THREAD_REGISTRY': self.registry}), tuple(self.threads))

    def test_slow_cleanup_returns_without_blocking_ui_caller(self):
        entered, release = threading.Event(), threading.Event()
        finished = threading.Event()
        caller_thread = threading.get_ident()
        cleanup_threads = []
        def cleanup(days):
            cleanup_threads.append(threading.get_ident())
            entered.set()
            release.wait(2)
            finished.set()
            return 1
        timers = []
        try:
            run_auto_log_cleanup_tick_runtime(
                state=AutoLogCleanupState(), is_enabled=lambda: True, retention_days=lambda: 7,
                interval_hours=lambda: 1, cleanup_old_logs=cleanup, system_ui=lambda *a: None,
                tk_alive=lambda: True, root_after=lambda ms, fn: timers.append((ms, fn)) or len(timers),
                tick_callback=lambda: None, ui_post=lambda fn: None,
            )
            self.assertTrue(entered.wait(1))
            self.assertNotEqual(cleanup_threads, [caller_thread])
            self.assertFalse(finished.is_set())
        finally:
            release.set()
            finished.wait(3)


class FileLogNotificationTests(unittest.TestCase):
    def setUp(self):
        self.now = 0
        self.state = FileLogErrorState(monotonic=lambda: self.now)
        self.shown = []
        self.has_text = False
        self.namespace = {
            'FILE_LOG_ERROR_STATE': self.state,
            'main_text_available': lambda: self.has_text,
            'ui_only': self.shown.append,
            'root': object(),
            'messagebox': type('Messages', (), {'showwarning': lambda _, *a, **k: self.shown.append(a)})(),
        }

    def test_startup_errors_are_retained_until_ui_ready(self):
        self.state.report('private data')
        notify_file_log_errors_namespace_runtime(self.namespace)
        self.assertFalse(self.shown)
        self.has_text = True
        notify_file_log_errors_namespace_runtime(self.namespace)
        self.assertEqual(len(self.shown), 1)
        self.assertIn('日志保存失败', self.shown[0])
        self.assertNotIn('private data', self.shown[0])

    def test_repeated_errors_are_coalesced_rate_limited_and_finally_forced(self):
        self.state.report('first')
        self.assertIsNotNone(self.state.take_notice())
        for _ in range(10000):
            self.state.report('again')
        self.assertIsNone(self.state.take_notice())
        self.now = 59
        self.assertIsNone(self.state.take_notice())
        self.now = 60
        self.assertIsNotNone(self.state.take_notice())
        self.state.report('last flush')
        self.namespace['is_exiting'] = True
        notify_file_log_errors_namespace_runtime(self.namespace)
        self.assertFalse(self.shown)
        notify_file_log_errors_namespace_runtime(self.namespace, force=True)
        self.assertEqual(len(self.shown), 1)
        self.assertIsNone(self.state.take_notice(force=True))

    def test_queue_full_reports_without_recursive_file_logging(self):
        logs = FileLogQueue(1, log_error=self.state.report)
        logs.put_nowait(('test.log', 'first'))
        with self.assertRaises(queue.Full):
            logs.put_nowait(('test.log', 'second'))
        self.assertIsNotNone(self.state.take_notice())
        self.assertEqual(logs.qsize(), 1)

    def test_normal_ui_pump_reports_despite_full_ui_queue(self):
        ns = dict(self.namespace, UI_TASK_QUEUE=queue.Queue(1), TK_SHUTDOWN=threading.Event())
        install_ui_log_namespace_bindings(ns)
        ns.update(main_text_available=lambda: True, ui_only=self.shown.append,
                  tk_alive=lambda: False, log_file_only=lambda message: self.fail('recursive logging'))
        ns['UI_TASK_QUEUE'].put_nowait((lambda: None, (), {}))
        self.state.report('error')
        ns['ui_pump']()
        self.assertEqual(len(self.shown), 1)

    def test_shutdown_flush_reports_independently_and_balances_queue(self):
        with tempfile.TemporaryDirectory() as temp:
            logs = queue.Queue()
            logs.put_nowait((str(Path(temp) / 'missing' / 'log.txt'), 'line'))
            flush = _flush_file_logs(dict(self.namespace, flush_log_queue=flush_log_queue))
            with patch.object(sys, 'stderr', None):
                self.assertEqual(flush(logs, log_error=lambda msg: self.fail('recursive logging')), 0)
            self.assertEqual(logs.unfinished_tasks, 0)
            self.assertIsNotNone(self.state.take_notice(force=True))

    def test_windowed_worker_reports_failed_batch_then_keeps_writing(self):
        with tempfile.TemporaryDirectory() as temp:
            logs = queue.Queue()
            stop = threading.Event()
            worker = start_file_log_worker(log_queue=logs, stop_event=stop, log_error=self.state.report)
            try:
                with patch.object(sys, 'stderr', None):
                    logs.put_nowait((str(Path(temp) / 'missing' / 'log.txt'), 'lost\n'))
                    logs.put_nowait((str(Path(temp) / 'good.txt'), 'retained\n'))
                    with logs.all_tasks_done:
                        self.assertTrue(logs.all_tasks_done.wait_for(lambda: logs.unfinished_tasks == 0, timeout=3))
                    self.assertIsNotNone(self.state.take_notice())
                    self.assertEqual((Path(temp) / 'good.txt').read_text(), 'retained\n')
            finally:
                stop.set()
                worker.join(3)
            self.assertFalse(worker.is_alive())

    @unittest.skipUnless(os.name == 'nt', 'Requires Windows file sharing locks')
    def test_windows_file_lock_is_reported_and_later_writes_recover(self):
        import ctypes
        from ctypes import wintypes
        create_file = ctypes.windll.kernel32.CreateFileW
        create_file.argtypes = [wintypes.LPCWSTR, wintypes.DWORD, wintypes.DWORD,
                               ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD, wintypes.HANDLE]
        create_file.restype = wintypes.HANDLE
        close_handle = ctypes.windll.kernel32.CloseHandle
        close_handle.argtypes = [wintypes.HANDLE]
        with tempfile.TemporaryDirectory() as temp:
            path = Path(temp) / 'locked.txt'
            path.write_text('', encoding='utf-8')
            handle = create_file(str(path), 0x80000000, 1, None, 3, 0x80, None)
            self.assertNotEqual(handle, wintypes.HANDLE(-1).value)
            logs, stop = queue.Queue(), threading.Event()
            worker = start_file_log_worker(log_queue=logs, stop_event=stop, log_error=self.state.report)
            try:
                try:
                    with patch.object(sys, 'stderr', None):
                        logs.put_nowait((str(path), 'blocked\n'))
                        with logs.all_tasks_done:
                            self.assertTrue(logs.all_tasks_done.wait_for(lambda: logs.unfinished_tasks == 0, timeout=3))
                        self.assertIsNotNone(self.state.take_notice())
                finally:
                    close_handle(handle)
                logs.put_nowait((str(path), 'recovered\n'))
                with logs.all_tasks_done:
                    self.assertTrue(logs.all_tasks_done.wait_for(lambda: logs.unfinished_tasks == 0, timeout=3))
                self.assertEqual(path.read_text(), 'recovered\n')
            finally:
                stop.set()
                worker.join(3)
            self.assertFalse(worker.is_alive())

    def test_final_notification_precedes_exit_and_restart_success(self):
        for action in ('退出', '重启'):
            with self.subTest(action=action):
                calls, timers = [], []
                class Root:
                    def after(self, ms, fn):
                        timers.append(fn)
                        return len(timers)
                class Progress:
                    def __init__(self, root):
                        pass
                    def close(self):
                        calls.append('close')
                    def update(self, message):
                        pass
                messages = type('Messages', (), {'askyesno': lambda *a, **k: True})()
                result = start_shutdown_task_runtime(
                    root=Root(), messagebox=messages, is_exiting=False, set_exiting=lambda v: None,
                    run_task=lambda report: 'exited', on_success=lambda: calls.append('success'),
                    confirmation='test', action=action, start_worker=lambda name, fn, **kw: fn(),
                    progress_factory=Progress, notify_log_errors=lambda **kw: calls.append(('notice', kw)),
                )
                self.assertEqual(result, 'started')
                timers.pop(0)()
                self.assertEqual(calls, [('notice', {'force': True}), 'close', 'success'])


if __name__ == '__main__':
    unittest.main()
