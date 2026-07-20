import unittest
from unittest.mock import patch

from sms_ui.update_proxy_dialog import _show_test_error, _show_test_result


class FakeWidget:
    def __init__(self, exists=True):
        self.exists = exists
        self.config_calls = []

    def winfo_exists(self):
        return self.exists

    def config(self, **kwargs):
        self.config_calls.append(kwargs)


class UpdateProxyDialogTests(unittest.TestCase):
    def test_result_and_error_use_live_dialog_as_parent(self):
        win = FakeWidget()
        button = FakeWidget()

        with patch("sms_ui.update_proxy_dialog.messagebox.showinfo") as show_info, patch(
            "sms_ui.update_proxy_dialog.messagebox.showerror"
        ) as show_error:
            self.assertTrue(_show_test_result(win, button, "connected"))
            self.assertTrue(_show_test_error(win, button, "failed"))

        self.assertEqual(
            button.config_calls,
            [
                {"state": "normal", "text": "测试连接"},
                {"state": "normal", "text": "测试连接"},
            ],
        )
        show_info.assert_called_once_with("测试结果", "connected", parent=win)
        show_error.assert_called_once_with("测试失败", "failed", parent=win)

    def test_callbacks_are_discarded_after_dialog_closes(self):
        win = FakeWidget(exists=False)
        button = FakeWidget()

        with patch("sms_ui.update_proxy_dialog.messagebox.showinfo") as show_info, patch(
            "sms_ui.update_proxy_dialog.messagebox.showerror"
        ) as show_error:
            self.assertFalse(_show_test_result(win, button, "connected"))
            self.assertFalse(_show_test_error(win, button, "failed"))

        self.assertEqual(button.config_calls, [])
        show_info.assert_not_called()
        show_error.assert_not_called()

    def test_callbacks_are_discarded_when_test_button_is_destroyed(self):
        win = FakeWidget()
        button = FakeWidget(exists=False)

        with patch("sms_ui.update_proxy_dialog.messagebox.showinfo") as show_info:
            self.assertFalse(_show_test_result(win, button, "connected"))

        show_info.assert_not_called()


if __name__ == "__main__":
    unittest.main()
