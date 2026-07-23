import unittest
from unittest.mock import patch

import sms_ui.serial_debug_finder as finder_module


class FakeVariable:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value

    def trace_add(self, _mode, _callback):
        return "trace-id"


class FakeText:
    def tag_config(self, *_args, **_kwargs):
        return None

    def tag_raise(self, *_args, **_kwargs):
        return None


class FakeWidget:
    def __init__(self, events, name="widget"):
        self.events = events
        self.name = name

    def pack(self, *_args, **_kwargs):
        return None

    def grid(self, *_args, **_kwargs):
        return None

    def bind(self, *_args, **_kwargs):
        return None

    def focus_set(self):
        self.events.append("focus")


class FakeWindow(FakeWidget):
    def attributes(self, name, value):
        self.events.append(("attributes", name, value))

    def withdraw(self):
        self.events.append("withdraw")

    def title(self, _title):
        return None

    def resizable(self, _width, _height):
        return None

    def transient(self, _parent):
        return None

    def protocol(self, *_args, **_kwargs):
        return None

    def update_idletasks(self):
        self.events.append("update")

    def deiconify(self):
        self.events.append("deiconify")

    def lift(self):
        self.events.append("lift")


class SerialDebugFinderTests(unittest.TestCase):
    def test_open_centers_hidden_window_before_showing(self):
        events = []
        parent = object()
        window = FakeWindow(events, "window")
        entry = FakeWidget(events, "entry")

        def widget_factory(*_args, **_kwargs):
            return FakeWidget(events)

        def center_window(actual_window, actual_parent):
            self.assertIs(actual_window, window)
            self.assertIs(actual_parent, parent)
            events.append("center")

        with (
            patch.object(finder_module.tk, "StringVar", FakeVariable),
            patch.object(finder_module.tk, "Toplevel", return_value=window) as toplevel,
            patch.object(finder_module.ttk, "Frame", side_effect=widget_factory),
            patch.object(finder_module.ttk, "Label", side_effect=widget_factory),
            patch.object(finder_module.ttk, "Entry", return_value=entry),
            patch.object(finder_module.ttk, "Button", side_effect=widget_factory),
        ):
            finder = finder_module.SerialDebugFinder(parent, FakeText(), center_window)
            finder.open()

        toplevel.assert_called_once_with(parent)
        self.assertEqual(events, [
            "withdraw",
            ("attributes", "-alpha", 0.0),
            "update",
            "center",
            "deiconify",
            "update",
            "center",
            ("attributes", "-alpha", 1.0),
            "lift",
            "focus",
        ])

    def test_open_falls_back_when_window_alpha_is_unavailable(self):
        events = []
        parent = object()
        window = FakeWindow(events, "window")
        entry = FakeWidget(events, "entry")

        def reject_alpha(_name, _value):
            events.append("alpha-unavailable")
            raise finder_module.tk.TclError("unsupported")

        window.attributes = reject_alpha

        def widget_factory(*_args, **_kwargs):
            return FakeWidget(events)

        def center_window(_actual_window, _actual_parent):
            events.append("center")

        with (
            patch.object(finder_module.tk, "StringVar", FakeVariable),
            patch.object(finder_module.tk, "Toplevel", return_value=window),
            patch.object(finder_module.ttk, "Frame", side_effect=widget_factory),
            patch.object(finder_module.ttk, "Label", side_effect=widget_factory),
            patch.object(finder_module.ttk, "Entry", return_value=entry),
            patch.object(finder_module.ttk, "Button", side_effect=widget_factory),
        ):
            finder = finder_module.SerialDebugFinder(parent, FakeText(), center_window)
            finder.open()

        self.assertEqual(events, [
            "withdraw",
            "alpha-unavailable",
            "update",
            "center",
            "deiconify",
            "lift",
            "focus",
        ])


if __name__ == "__main__":
    unittest.main()
