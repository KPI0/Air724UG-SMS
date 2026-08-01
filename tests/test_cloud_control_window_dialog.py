import unittest
from types import SimpleNamespace
from unittest.mock import patch

from sms_ui.cloud_control_window import open_cloud_control_window_dialog


class FakeWindow:
    def __init__(self, parent):
        self.parent = parent
        self.transient_parent = None

    def withdraw(self):
        pass

    def title(self, _value):
        pass

    def minsize(self, *_args):
        pass

    def resizable(self, *_args):
        pass

    def transient(self, parent):
        self.transient_parent = parent

    def protocol(self, *_args):
        pass

    def bind(self, *_args):
        pass

    def update_idletasks(self):
        pass

    def deiconify(self):
        pass

    def lift(self):
        pass

    def focus_force(self):
        pass


class FakeFrame:
    def __init__(self, *_args, **_kwargs):
        pass

    def pack(self, **_kwargs):
        pass

    def grid_columnconfigure(self, *_args, **_kwargs):
        pass


class CloudControlWindowDialogTests(unittest.TestCase):
    def test_cloud_window_does_not_force_itself_above_main_window(self):
        fake_form = SimpleNamespace(sync_from_state=lambda: None)

        with patch(
            "sms_ui.cloud_control_window.tk.Toplevel",
            side_effect=FakeWindow,
        ), patch(
            "sms_ui.cloud_control_window.ttk.Frame",
            FakeFrame,
        ), patch(
            "sms_ui.cloud_control_window.CloudControlFormController",
            return_value=fake_form,
        ), patch(
            "sms_ui.cloud_control_window.build_cloud_action_buttons",
        ):
            win = open_cloud_control_window_dialog(
                "main_window",
                lambda: {},
                "status_var",
                lambda *_args: None,
                lambda *_args: None,
                lambda *_args: None,
                lambda *_args: None,
                lambda *_args: None,
                lambda *_args: None,
            )

        self.assertIsNone(win.transient_parent)


if __name__ == "__main__":
    unittest.main()
