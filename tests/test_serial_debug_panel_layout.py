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
        self.bind_calls = getattr(self, "bind_calls", [])
        self.bind_calls.append((_args, _kwargs))

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
        bound_events = [call[0][0] for call in serial_text.bind_calls]
        self.assertIn("<Control-a>", bound_events)
        self.assertIn("<Control-c>", bound_events)

    def test_clipboard_text_escapes_embedded_control_characters(self):
        self.assertEqual(
            serial_debug_panel.format_serial_debug_clipboard_text(
                ">>> 发送: AT\\r\\n\nUSBS\x10bad\x00tail\tvalue"
            ),
            ">>> 发送: AT\\r\\n\nUSBS\\x10bad\\x00tail\tvalue",
        )

    def test_copy_selection_preserves_text_after_embedded_nul(self):
        class ClipboardText:
            def __init__(self):
                self.clipboard = ""
                self.updated = False

            def get(self, start, end):
                self.assert_indexes = (start, end)
                return "before\x00after"

            def clipboard_clear(self):
                self.clipboard = ""

            def clipboard_append(self, value):
                self.clipboard += value

            def update_idletasks(self):
                self.updated = True

        text = ClipboardText()

        result = serial_debug_panel.copy_serial_debug_selection(text)

        self.assertEqual(result, "break")
        self.assertEqual(text.assert_indexes, ("sel.first", "sel.last"))
        self.assertEqual(text.clipboard, "before\\x00after")
        self.assertTrue(text.updated)

    def test_information_center_action_follows_own_number_action(self):
        buttons = []

        def make_button(parent=None, **kwargs):
            widget = FakeWidget(parent, **kwargs)
            buttons.append(widget)
            return widget

        quick_button = FakeWidget()
        with patch.object(serial_debug_panel.ttk, "Button", side_effect=make_button):
            serial_debug_panel.create_serial_debug_quick_actions(
                "parent",
                "enabled_var",
                FakeWidget(),
                FakeWidget(),
                quick_button,
                lambda _command: None,
                lambda _commands: None,
                lambda *_args: None,
                lambda *_args: None,
                lambda *_args: None,
                lambda *_args: None,
                lambda *_args: None,
            )

        labels = [button.kwargs.get("text") for button in buttons]
        own_number_index = labels.index("修改本机号码 ☎")
        self.assertEqual(
            labels[own_number_index + 1],
            "修改信息中心号码 ✉️",
        )

    def test_manual_operator_action_follows_operator_scan(self):
        buttons = []
        actions = []

        def make_button(parent=None, **kwargs):
            widget = FakeWidget(parent, **kwargs)
            buttons.append(widget)
            return widget

        with (
            patch.object(serial_debug_panel.ttk, "Frame", FakeWidget),
            patch.object(serial_debug_panel.ttk, "LabelFrame", FakeWidget),
            patch.object(serial_debug_panel.ttk, "Scrollbar", FakeWidget),
            patch.object(serial_debug_panel.ttk, "Button", side_effect=make_button),
            patch.object(serial_debug_panel.tk, "Canvas", FakeWidget),
            patch.object(serial_debug_panel.tk, "Text", FakeWidget),
            patch.object(
                serial_debug_panel,
                "COMMON_SERIAL_COMMANDS",
                [
                    ("AT+COPS=?", "查询附近可用运营商"),
                    ("AT+COPS=0", "自动选择运营商"),
                    ("AT+CPIN?", "查看 PIN 码锁状态"),
                ],
            ),
        ):
            serial_debug_panel.create_serial_debug_body(
                FakeWidget(),
                lambda command: actions.append(command),
                lambda: actions.append("manual"),
            )

        labels = [button.kwargs.get("text") for button in buttons]
        scan_index = labels.index("AT+COPS=?  (查询附近可用运营商)")
        self.assertEqual(
            labels[scan_index + 1],
            'AT+COPS=1,2,"PLMN"  (手动切换运营商)',
        )
        self.assertEqual(
            labels[scan_index + 2],
            "AT+COPS=0  (自动选择运营商)",
        )
        buttons[scan_index + 1].kwargs["command"]()
        buttons[scan_index + 2].kwargs["command"]()
        self.assertEqual(actions, ["manual", "AT+COPS=0"])


if __name__ == "__main__":
    unittest.main()
