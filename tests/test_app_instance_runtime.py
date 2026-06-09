import unittest

from sms_ui.app_instance_runtime import APP_MUTEX_NAME, check_single_instance_app_runtime


class AppInstanceRuntimeTests(unittest.TestCase):
    def test_check_single_instance_skips_mutex_when_multi_instance_allowed(self):
        calls = []

        result = check_single_instance_app_runtime(
            allow_multi_instance=True,
            window_title="title",
            acquire_mutex=lambda name: calls.append(("acquire", name)) or ("mutex", 0),
        )

        self.assertIsNone(result)
        self.assertEqual(calls, [])

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
