import os
import unittest

from sms_ui.app_instance_runtime import (
    APP_MUTEX_NAME,
    app_dir_mutex_name,
    check_single_instance_app_runtime,
    claim_instance_number_app_runtime,
    format_instance_window_title,
    instance_mutex_name,
)


class AppInstanceRuntimeTests(unittest.TestCase):
    def test_format_instance_window_title_keeps_first_instance_plain(self):
        self.assertEqual(format_instance_window_title("短信监听系统", 1), "短信监听系统")
        self.assertEqual(format_instance_window_title("短信监听系统", 2), "短信监听系统 2")
        self.assertEqual(format_instance_window_title("串口调试", 4), "串口调试 4")

    def test_claim_instance_number_uses_first_available_slot(self):
        app_dir = os.path.join(os.getcwd(), "sms-client")
        calls = []

        def acquire(name):
            calls.append(("acquire", name))
            if name == instance_mutex_name(app_dir, 1):
                return "occupied-1", 183
            if name == instance_mutex_name(app_dir, 2):
                return "occupied-2", 5
            return "instance-3", 0

        result = claim_instance_number_app_runtime(
            app_dir=app_dir,
            acquire_mutex=acquire,
            close_handle=lambda handle: calls.append(("close", handle)),
            existing_error=lambda code: code in (183, 5),
        )

        self.assertEqual(result, (3, "instance-3"))
        self.assertEqual(
            calls,
            [
                ("acquire", instance_mutex_name(app_dir, 1)),
                ("close", "occupied-1"),
                ("acquire", instance_mutex_name(app_dir, 2)),
                ("close", "occupied-2"),
                ("acquire", instance_mutex_name(app_dir, 3)),
            ],
        )

    def test_claim_instance_number_falls_back_to_plain_title(self):
        logs = []
        result = claim_instance_number_app_runtime(
            max_instance_number=2,
            acquire_mutex=lambda _name: (None, 8),
            existing_error=lambda _code: False,
            log_error=logs.append,
        )

        self.assertEqual(result, (1, None))
        self.assertTrue(any("using unnumbered title" in message for message in logs))

    def test_check_multi_instance_retains_presence_mutex(self):
        calls = []

        result = check_single_instance_app_runtime(
            allow_multi_instance=True,
            window_title="title",
            acquire_mutex=lambda name: calls.append(("acquire", name)) or ("mutex", 183),
        )

        self.assertEqual(result, "mutex")
        self.assertEqual(calls, [("acquire", APP_MUTEX_NAME)])

    def test_check_multi_instance_logs_presence_mutex_failure(self):
        logs = []
        calls = []

        result = check_single_instance_app_runtime(
            allow_multi_instance=True,
            window_title="title",
            acquire_mutex=lambda _name: (None, 8),
            log_error=logs.append,
            show_message=lambda *args: calls.append(("message", args)),
            exit_process=lambda code: calls.append(("exit", code)),
        )

        self.assertIsNone(result)
        self.assertTrue(any("presence mutex failed" in message for message in logs))
        self.assertEqual(calls[0][0], "message")
        self.assertEqual(calls[1], ("exit", 1))

    def test_each_multi_instance_retains_presence_handle_that_blocks_single_startup(self):
        responses = iter([
            ("multi-1", 0),
            ("multi-2", 183),
            ("single-probe", 183),
        ])
        closed = []
        exits = []

        first = check_single_instance_app_runtime(
            allow_multi_instance=True,
            window_title="title",
            acquire_mutex=lambda _name: next(responses),
        )
        second = check_single_instance_app_runtime(
            allow_multi_instance=True,
            window_title="title",
            acquire_mutex=lambda _name: next(responses),
        )
        single = check_single_instance_app_runtime(
            allow_multi_instance=False,
            window_title="title",
            acquire_mutex=lambda _name: next(responses),
            close_handle=closed.append,
            existing_error=lambda code: code == 183,
            focus_window=lambda _title: True,
            exit_process=exits.append,
        )

        self.assertEqual((first, second), ("multi-1", "multi-2"))
        self.assertIsNone(single)
        self.assertEqual(closed, ["single-probe"])
        self.assertEqual(exits, [0])

    def test_check_single_instance_returns_acquired_mutex(self):
        calls = []

        result = check_single_instance_app_runtime(
            allow_multi_instance=False,
            window_title="title",
            acquire_mutex=lambda name: calls.append(("acquire", name)) or ("mutex", 0),
            existing_error=lambda code: False,
        )

        self.assertEqual(result, "mutex")
        self.assertEqual(calls, [("acquire", APP_MUTEX_NAME)])

    def test_app_dir_mutex_name_is_stable_per_directory(self):
        base_dir = os.path.join(os.getcwd(), "sms-client")
        first = app_dir_mutex_name(base_dir)
        same = app_dir_mutex_name(os.path.join(base_dir, "."))
        other = app_dir_mutex_name(os.path.join(os.getcwd(), "sms-client-2"))

        self.assertEqual(first, same)
        self.assertNotEqual(first, other)
        self.assertTrue(first.startswith(APP_MUTEX_NAME + "_"))

    def test_check_single_instance_uses_app_dir_mutex_name(self):
        calls = []
        app_dir = os.path.join(os.getcwd(), "sms-client")

        result = check_single_instance_app_runtime(
            allow_multi_instance=False,
            window_title="title",
            app_dir=app_dir,
            acquire_mutex=lambda name: calls.append(("acquire", name)) or ("mutex", 0),
            existing_error=lambda code: False,
        )

        self.assertEqual(result, "mutex")
        self.assertEqual(calls, [("acquire", app_dir_mutex_name(app_dir))])

    def test_check_single_instance_exits_existing_instance_and_focuses_window(self):
        calls = []

        result = check_single_instance_app_runtime(
            allow_multi_instance=False,
            window_title="title",
            acquire_mutex=lambda name: ("mutex", 183),
            close_handle=lambda handle: calls.append(("close", handle)),
            existing_error=lambda code: code == 183,
            focus_window=lambda title: calls.append(("focus", title)) or True,
            show_message=lambda *args: calls.append(("message", args)),
            exit_process=lambda code: calls.append(("exit", code)),
        )

        self.assertIsNone(result)
        self.assertEqual(calls, [("close", "mutex"), ("focus", "title"), ("exit", 0)])

    def test_check_single_instance_logs_close_handle_failure(self):
        calls = []

        result = check_single_instance_app_runtime(
            allow_multi_instance=False,
            window_title="title",
            acquire_mutex=lambda name: ("mutex", 183),
            close_handle=lambda handle: (_ for _ in ()).throw(RuntimeError("close failed")),
            existing_error=lambda code: code == 183,
            focus_window=lambda title: calls.append(("focus", title)) or True,
            exit_process=lambda code: calls.append(("exit", code)),
            log_error=lambda message: calls.append(("log", message)),
        )

        self.assertIsNone(result)
        self.assertEqual(calls[0][0], "log")
        self.assertIn("close failed", calls[0][1])
        self.assertEqual(calls[1:], [("focus", "title"), ("exit", 0)])

    def test_check_single_instance_shows_message_when_existing_window_not_found(self):
        calls = []

        check_single_instance_app_runtime(
            allow_multi_instance=False,
            window_title="title",
            acquire_mutex=lambda name: ("mutex", 5),
            close_handle=lambda handle: calls.append(("close", handle)),
            existing_error=lambda code: code == 5,
            focus_window=lambda title: False,
            show_message=lambda *args: calls.append(("message", args)),
            exit_process=lambda code: calls.append(("exit", code)),
        )

        self.assertEqual(calls[0], ("close", "mutex"))
        self.assertEqual(calls[1][0], "message")
        self.assertEqual(calls[2], ("exit", 0))

    def test_check_single_instance_exits_when_mutex_creation_fails(self):
        calls = []

        result = check_single_instance_app_runtime(
            allow_multi_instance=False,
            window_title="title",
            acquire_mutex=lambda name: (None, 8),
            existing_error=lambda code: False,
            show_message=lambda *args: calls.append(("message", args)),
            exit_process=lambda code: calls.append(("exit", code)),
        )

        self.assertIsNone(result)
        self.assertEqual(calls[0][0], "message")
        self.assertEqual(calls[1], ("exit", 1))


if __name__ == "__main__":
    unittest.main()
