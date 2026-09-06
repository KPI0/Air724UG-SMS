import unittest
from unittest.mock import patch

from sms_ui.serial_debug_window import open_serial_debug_window_dialog


class FakeVar:
    def __init__(self, value=None):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value

    def trace_add(self, *_args):
        return None


class FakeWidget:
    def __init__(self, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs

    def pack(self, *args, **kwargs):
        return None

    def grid(self, *args, **kwargs):
        return None

    def config(self, *args, **kwargs):
        self.kwargs.update(kwargs)

    def bind(self, *args, **kwargs):
        return None

    def delete(self, *args, **kwargs):
        return None


class FakeWindow(FakeWidget):
    def __init__(self):
        super().__init__()
        self.protocols = {}

    def winfo_exists(self):
        return True

    def withdraw(self):
        return None

    def title(self, *_args):
        return None

    def geometry(self, *_args):
        return None

    def minsize(self, *_args):
        return None

    def lift(self):
        return None

    def focus_force(self):
        return None

    def protocol(self, name, callback):
        self.protocols[name] = callback

    def bind(self, *args, **kwargs):
        return None

    def update_idletasks(self):
        return None

    def deiconify(self):
        return None


class FakePauseController:
    def __init__(self, *_args):
        return None

    def refresh(self):
        return None

    def toggle(self):
        return None


class FakeFinder:
    def __init__(self, *_args):
        return None


class SerialDebugDialCallbackTests(unittest.TestCase):
    def test_popup_hangup_callback_accepts_popup_close_callback(self):
        window = FakeWindow()
        serial_text = FakeWidget()
        captured = {}
        sent = []
        status = []
        dial_state = []

        def capture_quick_actions(*args):
            captured["close_active_dial_call"] = args[10]

        with (
            patch("sms_ui.serial_debug_window.tk.Toplevel", return_value=window),
            patch("sms_ui.serial_debug_window.tk.BooleanVar", side_effect=FakeVar),
            patch("sms_ui.serial_debug_window.tk.StringVar", side_effect=FakeVar),
            patch("sms_ui.serial_debug_window.ttk.Frame", FakeWidget),
            patch("sms_ui.serial_debug_window.ttk.Checkbutton", FakeWidget),
            patch("sms_ui.serial_debug_window.ttk.Label", FakeWidget),
            patch("sms_ui.serial_debug_window.ttk.Entry", FakeWidget),
            patch("sms_ui.serial_debug_window.ttk.Button", FakeWidget),
            patch(
                "sms_ui.serial_debug_window.create_serial_debug_body",
                return_value=(serial_text, FakeWidget(), FakeWidget()),
            ),
            patch(
                "sms_ui.serial_debug_window.create_serial_debug_quick_actions",
                side_effect=capture_quick_actions,
            ),
            patch("sms_ui.serial_debug_window.SerialDebugPauseController", FakePauseController),
            patch("sms_ui.serial_debug_window.SerialDebugFinder", FakeFinder),
            patch("sms_ui.serial_debug_window.start_serial_debug_append_loop"),
            patch(
                "sms_ui.serial_debug_window.send_command_async",
                side_effect=lambda *args, **kwargs: sent.append((args, kwargs)),
            ),
        ):
            open_serial_debug_window_dialog(
                "parent",
                None,
                None,
                True,
                lambda: 0,
                "queue",
                "lock",
                lambda: "serial",
                lambda *_args: None,
                lambda *_args: None,
                lambda *_args: status.append(_args),
                lambda port: f"connected:{port}",
                lambda: "COM5",
                lambda value: dial_state.append(value),
                lambda *_args: None,
                lambda *_args: None,
                lambda: None,
                "center",
                get_dial_popup=lambda: object(),
                set_dial_popup=lambda *_args: None,
            )

            captured["close_active_dial_call"]()

        self.assertEqual(len(sent), 1)
        self.assertEqual(sent[0][0][2], "ATH")
        self.assertTrue(status)
        self.assertEqual(dial_state, [""])


if __name__ == "__main__":
    unittest.main()
