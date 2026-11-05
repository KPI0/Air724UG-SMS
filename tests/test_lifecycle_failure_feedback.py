import os
from pathlib import Path
import tempfile
import threading
import unittest
from unittest.mock import patch

from sms_core.threading_runtime import WorkerThreadRegistry, start_registered_daemon_thread
from sms_ui.maintenance_namespace_bindings import install_maintenance_namespace_bindings
from sms_ui.maintenance_runtime import AutoLogCleanupState
from sms_ui import maintenance_runtime
from sms_ui.maintenance_namespace_runtime import open_update_proxy_dialog_namespace_runtime
from tests import test_settings_window_sync as window_fixtures


class ImmediateThread:
    def __init__(self, target, daemon, name=None):
        self.target = target

    def start(self):
        self.target()


class LogCleanupFailureFeedbackTests(unittest.TestCase):
    def run_cleanup(self, directory):
        messages, timers = [], []
        namespace = {
            "LOG_DIR": str(directory), "AUTO_LOG_CLEANUP": True,
            "LOG_RETENTION_DAYS": 30, "AUTO_CLEANUP_INTERVAL_HOURS": 6,
            "AUTO_LOG_CLEANUP_STATE": AutoLogCleanupState(),
            "TK_SHUTDOWN": threading.Event(), "is_exiting": False,
            "tk_alive": lambda: True, "ui_post": lambda callback: callback(),
            "system_ui": lambda message, tag="normal": messages.append(message),
            "clear_window": lambda: None,
            "root": type("Timers", (), {"after": lambda self, delay, callback:
                    timers.append((delay, callback)) or "next"})(),
        }
        install_maintenance_namespace_bindings(namespace)

        def start_worker(name, target, **options):
            options["thread_factory"] = ImmediateThread
            return start_registered_daemon_thread(name, target, **options)

        with patch.object(maintenance_runtime, "start_registered_daemon_thread", start_worker):
            namespace["_auto_log_cleanup_tick"]()
        self.assertFalse(namespace["AUTO_LOG_CLEANUP_STATE"].running)
        self.assertEqual(namespace["AUTO_LOG_CLEANUP_STATE"].results.unfinished_tasks, 0)
        self.assertEqual(timers[0][0], 6 * 3600 * 1000)
        return messages

    def test_partial_delete_failure_is_reported_and_other_logs_still_deleted(self):
        with tempfile.TemporaryDirectory() as directory:
            paths = [Path(directory, f"sms_{name}_2000-01-01.txt") for name in ("first", "locked", "last")]
            for path in paths:
                path.write_text("synthetic log", encoding="utf-8")
            original_remove = os.remove

            def remove(path):
                if Path(path) == paths[1]:
                    raise PermissionError("synthetic-sensitive-marker")
                original_remove(path)

            with patch("sms_core.log_cleanup.os.listdir", return_value=[path.name for path in paths]), \
                    patch("sms_core.log_cleanup.os.remove", side_effect=remove):
                messages = self.run_cleanup(directory)
            self.assertFalse(paths[0].exists())
            self.assertTrue(paths[1].exists())
            self.assertFalse(paths[2].exists())
            self.assertEqual(len(messages), 1)
            self.assertIn("失败", messages[0])
            self.assertIn("2", messages[0])
            self.assertIn("1", messages[0])
            self.assertNotIn("synthetic-sensitive-marker", messages[0])

    def test_file_metadata_failure_is_reported_without_exception_detail(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory, "sms_undated.txt")
            path.write_text("synthetic log", encoding="utf-8")
            with patch("sms_core.log_cleanup.os.path.getmtime", side_effect=PermissionError("synthetic-sensitive-marker")):
                messages = self.run_cleanup(directory)
            self.assertTrue(path.exists())
            self.assertIn("失败", messages[0])
            self.assertNotIn("synthetic-sensitive-marker", messages[0])

    def test_log_removed_by_another_instance_is_a_successful_cleanup(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory, "sms_shared_2000-01-01.txt")
            path.write_text("synthetic log", encoding="utf-8")
            original_remove = os.remove

            def remove(path):
                original_remove(path)
                raise FileNotFoundError("already removed by peer")

            with patch("sms_core.log_cleanup.os.remove", side_effect=remove):
                messages = self.run_cleanup(directory)
            self.assertFalse(path.exists())
            self.assertNotIn("失败", messages[0])


class UpdateProxyWorkerStartFeedbackTests(unittest.TestCase):
    setUp = window_fixtures.SettingsWindowSyncTests.setUp
    tearDown = window_fixtures.SettingsWindowSyncTests.tearDown
    make_namespace = window_fixtures.SettingsWindowSyncTests.make_namespace
    open_window = window_fixtures.SettingsWindowSyncTests.open_window
    button = window_fixtures.SettingsWindowSyncTests.button

    def check_rejected_worker(self, factory):
        registry = self.second["UPDATE_THREAD_REGISTRY"] = WorkerThreadRegistry()
        win = self.open_window(open_update_proxy_dialog_namespace_runtime)
        button = self.button(win, "测试连接")
        callback_errors = []
        self.root.report_callback_exception = lambda *args: callback_errors.append(args)

        def start_worker(name, target, **options):
            options["thread_factory"] = factory
            return start_registered_daemon_thread(name, target, **options)

        with patch.object(maintenance_runtime, "start_registered_daemon_thread", start_worker), \
                patch("sms_ui.update_proxy_dialog.messagebox.showerror") as show_error:
            button.invoke()
            self.root.update()
        self.assertEqual(callback_errors, [])
        self.assertEqual(registry.snapshot(), ())
        self.assertEqual((button.cget("state"), button.cget("text")), ("normal", "测试连接"))
        show_error.assert_called_once()
        self.assertNotIn("synthetic-sensitive-marker", str(show_error.call_args))
        win.destroy()

    def test_proxy_thread_construction_failure_restores_test_button(self):
        def factory(**_options):
            raise RuntimeError("synthetic-sensitive-marker")

        self.check_rejected_worker(factory)

    def test_proxy_thread_start_failure_restores_test_button(self):
        class RejectedStart(ImmediateThread):
            def start(self):
                raise RuntimeError("synthetic-sensitive-marker")

        self.check_rejected_worker(RejectedStart)


if __name__ == "__main__":
    unittest.main()
