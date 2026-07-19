import unittest

from sms_core.app_launch import (
    build_restart_helper_command,
    decode_restart_args,
    filtered_restart_args,
    prepare_restart_helper_launch,
    restart_software_runtime,
)


class AppLaunchRestartTests(unittest.TestCase):
    def test_filtered_restart_args_removes_restart_only_flags(self):
        self.assertEqual(
            filtered_restart_args(
                ["--autostart", "--debug", "--restart-helper", "--keep"],
                ("--autostart", "--restart-helper"),
            ),
            ["--debug", "--keep"],
        )

    def test_build_restart_helper_command_encodes_remaining_args(self):
        command = build_restart_helper_command(
            "pythonw.exe",
            "sms.pyw",
            "--restart-helper",
            1234,
            ["--debug", "中文"],
        )

        self.assertEqual(command[:4], ["pythonw.exe", "sms.pyw", "--restart-helper", "1234"])
        self.assertEqual(decode_restart_args(command[4]), ["--debug", "中文"])

    def test_prepare_restart_helper_launch_uses_launch_target_and_filters_flags(self):
        command, workdir = prepare_restart_helper_launch(
            ["sms.pyw", "--autostart", "--debug", "--restart-helper", "--keep"],
            "--autostart",
            "--restart-helper",
            4321,
            launch_target_func=lambda: ("pythonw.exe", "sms.pyw", "E:/sms"),
        )

        self.assertEqual(workdir, "E:/sms")
        self.assertEqual(command[:4], ["pythonw.exe", "sms.pyw", "--restart-helper", "4321"])
        self.assertEqual(decode_restart_args(command[4]), ["--debug", "--keep"])

    def _runtime_kwargs(self, calls, **overrides):
        values = {
            "is_exiting": False,
            "confirm_restart": lambda: True,
            "argv": ["sms.pyw"],
            "autostart_flag": "--autostart",
            "restart_helper_flag": "--restart-helper",
            "current_pid": 123,
            "log_error": lambda msg: calls.append(("log_error", msg)),
            "show_launch_error": lambda error: calls.append(("show_error", str(error))),
            "set_exiting": lambda value: calls.append(("exiting", value)),
            "system_ui": lambda *args: calls.append(("system", args)),
            "stop_tray_icon": lambda **kwargs: calls.append(("tray", kwargs)),
            "set_serial_running": lambda value: calls.append(("serial", value)),
            "safe_set_events": lambda *events: calls.append(("events", events)),
            "stop_events": ("third", "serial", "wakeup"),
            "stop_cloud_control": lambda **kwargs: calls.append(("cloud", kwargs)),
            "safe_close_serial": lambda: calls.append(("close_serial",)),
            "app_mutex": "mutex",
            "release_mutex": lambda handle: calls.append(("release", handle)),
            "flush_log_queue": lambda queue: calls.append(("flush", queue)),
            "file_log_queue": "queue",
            "file_log_thread": "file_thread",
            "file_log_stop_event": "file_stop",
            "worker_threads": ("producer",),
            "wait_worker_threads": lambda threads, **kwargs: calls.append(("wait_workers", threads, kwargs)),
            "wait_file_log_worker": lambda thread, **kwargs: calls.append(("wait_file", thread, kwargs)),
            "exit_process": lambda code: calls.append(("exit", code)),
            "prepare_launch": lambda argv, autostart, restart, pid: (["helper"], "E:/sms"),
            "launch_process": lambda command, env, cwd: calls.append(("launch", command, env, cwd)),
            "clean_env": lambda: {"clean": "1"},
        }
        values.update(overrides)
        return values

    def test_restart_runtime_cancel_and_already_exiting_do_not_cleanup(self):
        calls = []
        result = restart_software_runtime(**self._runtime_kwargs(calls, is_exiting=True))
        self.assertEqual(result.status, "already_exiting")
        self.assertEqual(calls, [])

        calls = []
        result = restart_software_runtime(**self._runtime_kwargs(calls, confirm_restart=lambda: False))
        self.assertEqual(result.status, "cancelled")
        self.assertEqual(calls, [])

    def test_restart_runtime_launch_failure_keeps_current_app_running(self):
        calls = []

        def fail_launch(command, env, cwd):
            raise RuntimeError("boom")

        result = restart_software_runtime(**self._runtime_kwargs(calls, launch_process=fail_launch))

        self.assertEqual(result.status, "launch_failed")
        self.assertTrue(any(item[0] == "log_error" for item in calls))
        self.assertIn(("show_error", "boom"), calls)
        self.assertFalse(any(item[0] == "exit" for item in calls))

    def test_restart_runtime_success_starts_helper_then_cleans_up_and_exits(self):
        calls = []
        result = restart_software_runtime(**self._runtime_kwargs(calls))

        self.assertEqual(result.status, "exited")
        self.assertEqual(calls[0], ("launch", ["helper"], {"clean": "1"}, "E:/sms"))
        self.assertIn(("exiting", True), calls)
        self.assertIn(("serial", False), calls)
        self.assertIn(("events", ("third", "serial", "wakeup")), calls)
        self.assertIn(("cloud", {"update_status": False}), calls)
        self.assertIn(("close_serial",), calls)
        producer_waits = [item for item in calls if item[0] == "wait_workers"]
        self.assertEqual(len(producer_waits), 1)
        self.assertEqual(producer_waits[0][1], ("producer",))
        self.assertIn(("events", ("file_stop",)), calls)
        self.assertIn(("release", "mutex"), calls)
        wait_calls = [item for item in calls if item[0] == "wait_file"]
        self.assertEqual(len(wait_calls), 1)
        self.assertEqual(wait_calls[0][1], "file_thread")
        self.assertIn("log_error", wait_calls[0][2])
        self.assertIn(("flush", "queue"), calls)
        self.assertEqual(calls[-1], ("exit", 0))


if __name__ == "__main__":
    unittest.main()
