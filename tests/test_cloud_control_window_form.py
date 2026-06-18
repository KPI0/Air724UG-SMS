import unittest
from unittest.mock import patch

from sms_ui.cloud_control_window_form import CloudControlFormController


class FakeBoolVar:
    def __init__(self, value):
        self.value = bool(value)

    def get(self):
        return self.value

    def set(self, value):
        self.value = bool(value)


class CloudControlWindowFormTests(unittest.TestCase):
    def make_form(self, initial_value, calls):
        form = CloudControlFormController.__new__(CloudControlFormController)
        form.win = "window"
        form.auto_upload_var = FakeBoolVar(initial_value)
        form.on_auto_upload_changed = lambda value, win: calls.append((value, win))
        return form

    def test_auto_upload_toggle_confirms_before_public_device(self):
        calls = []
        form = self.make_form(True, calls)

        with patch("sms_ui.cloud_control_window_form.confirm_public_device", return_value=True) as confirm:
            form._auto_upload_toggled()

        confirm.assert_called_once_with("window")
        self.assertEqual(calls, [(True, "window")])
        self.assertTrue(form.auto_upload_var.get())

    def test_auto_upload_toggle_cancel_reverts_checkbox(self):
        calls = []
        form = self.make_form(True, calls)

        with patch("sms_ui.cloud_control_window_form.confirm_public_device", return_value=False) as confirm:
            form._auto_upload_toggled()

        confirm.assert_called_once_with("window")
        self.assertEqual(calls, [])
        self.assertFalse(form.auto_upload_var.get())

    def test_auto_upload_toggle_off_does_not_confirm(self):
        calls = []
        form = self.make_form(False, calls)

        with patch("sms_ui.cloud_control_window_form.confirm_public_device") as confirm:
            form._auto_upload_toggled()

        confirm.assert_not_called()
        self.assertEqual(calls, [(False, "window")])


if __name__ == "__main__":
    unittest.main()
