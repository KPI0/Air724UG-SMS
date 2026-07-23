import unittest
from unittest.mock import patch

import sms_ui.serial_debug_dialog_helpers as helpers


class FakeWindow:
    def __init__(self, events):
        self.events = events

    def withdraw(self):
        self.events.append("withdraw")

    def update_idletasks(self):
        self.events.append("update")

    def deiconify(self):
        self.events.append("deiconify")


class FakeFocusWidget:
    def __init__(self, events):
        self.events = events

    def focus_set(self):
        self.events.append("focus")


class SerialDebugDialogHelperTests(unittest.TestCase):
    def test_create_debug_dialog_hides_window_immediately(self):
        events = []
        parent = object()
        window = FakeWindow(events)

        with patch.object(helpers.tk, "Toplevel", return_value=window) as toplevel:
            result = helpers.create_debug_dialog(parent)

        self.assertIs(result, window)
        toplevel.assert_called_once_with(parent)
        self.assertEqual(events, ["withdraw"])

    def test_show_centered_debug_dialog_positions_before_showing_and_focusing(self):
        events = []
        parent = object()
        window = FakeWindow(events)
        focus_widget = FakeFocusWidget(events)

        def center_window(actual_window, actual_parent):
            self.assertIs(actual_window, window)
            self.assertIs(actual_parent, parent)
            events.append("center")

        helpers.show_centered_debug_dialog(window, center_window, parent, focus_widget)

        self.assertEqual(events, ["update", "center", "deiconify", "focus"])


if __name__ == "__main__":
    unittest.main()
