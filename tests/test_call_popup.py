import unittest
from unittest.mock import patch

from sms_ui.call_popup import open_call_popup


class FakeWidget:
    def __init__(self, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs

    def pack(self, *args, **kwargs):
        return None

    def config(self, *args, **kwargs):
        return None

    def pack_forget(self):
        return None

    def winfo_exists(self):
        return True


class FakeWindow(FakeWidget):
    def __init__(self, events):
        super().__init__()
        self.events = events

    def withdraw(self):
        self.events.append("withdraw")

    def title(self, value):
        return None

    def minsize(self, width, height):
        return None

    def resizable(self, width, height):
        return None

    def attributes(self, *args):
        return None

    def protocol(self, name, callback):
        return None

    def deiconify(self):
        self.events.append("deiconify")

    def lift(self):
        self.events.append("lift")


class CallPopupTests(unittest.TestCase):
    def test_popup_is_centered_while_hidden_before_it_is_shown(self):
        events = []
        parent = object()
        window = FakeWindow(events)

        def center_window(actual_window, actual_parent):
            self.assertIs(actual_window, window)
            self.assertIs(actual_parent, parent)
            events.append("center")

        with patch("sms_ui.call_popup.tk.Toplevel", return_value=window), \
                patch("sms_ui.call_popup.ttk.Frame", FakeWidget), \
                patch("sms_ui.call_popup.tk.Label", FakeWidget), \
                patch("sms_ui.call_popup.ttk.Button", FakeWidget):
            result = open_call_popup(
                parent,
                "10086",
                center_window,
                lambda *_args: None,
                lambda *_args: None,
                lambda: None,
                lambda: None,
            )

        self.assertIs(result, window)
        self.assertEqual(events, ["withdraw", "center", "deiconify", "lift"])


if __name__ == "__main__":
    unittest.main()
