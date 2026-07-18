import unittest
from unittest.mock import patch

from sms_ui import serial_debug_panel


class FakeWidget:
    def __init__(self, _parent=None, **kwargs):
        self.kwargs = kwargs
        self.pack_calls = []
        self.grid_calls = []
        self.config_calls = []

    def pack(self, **kwargs):
        self.pack_calls.append(kwargs)

    def grid(self, **kwargs):
        self.grid_calls.append(kwargs)

    def grid_rowconfigure(self, *_args, **_kwargs):
        pass

    def grid_columnconfigure(self, *_args, **_kwargs):
        pass

    def config(self, **kwargs):
        self.config_calls.append(kwargs)

    configure = config

    def bind(self, *_args, **_kwargs):
        pass

    def bind_all(self, *_args, **_kwargs):
        pass

    def unbind_all(self, *_args, **_kwargs):
        pass

    def create_window(self, *_args, **_kwargs):
        return "window"

    def itemconfig(self, *_args, **_kwargs):
        pass

    def bbox(self, *_args, **_kwargs):
        return (0, 0, 100, 100)

    def yview(self, *_args, **_kwargs):
        pass

    def xview(self, *_args, **_kwargs):
        pass

    def set(self, *_args, **_kwargs):
        pass


class SerialDebugPanelLayoutTests(unittest.TestCase):
    def test_debug_text_has_horizontal_scrollbar(self):
        scrollbars = []

        def make_scrollbar(parent=None, **kwargs):
            widget = FakeWidget(parent, **kwargs)
            scrollbars.append(widget)
            return widget

        with (
            patch.object(serial_debug_panel.ttk, "Frame", FakeWidget),
            patch.object(serial_debug_panel.ttk, "LabelFrame", FakeWidget),
            patch.object(serial_debug_panel.ttk, "Scrollbar", side_effect=make_scrollbar),
            patch.object(serial_debug_panel.ttk, "Button", FakeWidget),
            patch.object(serial_debug_panel.tk, "Canvas", FakeWidget),
            patch.object(serial_debug_panel.tk, "Text", FakeWidget),
            patch.object(serial_debug_panel, "COMMON_SERIAL_COMMANDS", []),
        ):
            serial_text, _quick_panel, _quick_scroll_frame = (
                serial_debug_panel.create_serial_debug_body(FakeWidget(), lambda _cmd: None)
            )

        horizontal = next(
            scrollbar
            for scrollbar in scrollbars
            if scrollbar.kwargs.get("orient") == "horizontal"
        )
        self.assertEqual(serial_text.kwargs["wrap"], "none")
        self.assertEqual(serial_text.kwargs["xscrollcommand"], horizontal.set)
        self.assertIn({"side": "bottom", "fill": "x"}, horizontal.pack_calls)
        self.assertTrue(
            any(call.get("command") == serial_text.xview for call in horizontal.config_calls)
        )


if __name__ == "__main__":
    unittest.main()
