import unittest
from unittest.mock import patch

from sms_ui.serial_debug_sms_call_dialogs import format_sms_pdu_counter, open_dial_dialog


class FakeVar:
    def __init__(self, value="15923240141"):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeWidget:
    def __init__(self, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs

    def pack(self, *args, **kwargs):
        return None

    def bind(self, *args, **kwargs):
        return None


class FakeDialogWindow:
    def __init__(self):
        self.destroyed = False
        self.grab_released = False

    def title(self, *_args):
        return None

    def resizable(self, *_args):
        return None

    def transient(self, *_args):
        return None

    def grab_set(self):
        return None

    def grab_release(self):
        self.grab_released = True

    def destroy(self):
        self.destroyed = True

    def bind(self, *_args, **_kwargs):
        return None

    def protocol(self, *_args, **_kwargs):
        return None


class SerialDebugSmsCallDialogTests(unittest.TestCase):
    def test_format_sms_pdu_counter_allows_segmented_long_messages(self):
        text, too_long = format_sms_pdu_counter("A" * 100)

        self.assertEqual(text, "100 字 | UCS2 200 字节 | 2/255 段")
        self.assertFalse(too_long)

    def test_format_sms_pdu_counter_counts_emoji_by_ucs2_bytes(self):
        text, too_long = format_sms_pdu_counter("😀" * 36)

        self.assertEqual(text, "36 字 | UCS2 144 字节 | 2/255 段")
        self.assertFalse(too_long)

    def test_format_sms_pdu_counter_marks_over_segment_limit(self):
        text, too_long = format_sms_pdu_counter("A" * 17086)

        self.assertEqual(text, "17086 字 | UCS2 34172 字节 | 256/255 段")
        self.assertTrue(too_long)

    def test_dial_submission_releases_modal_grab_before_showing_call_status(self):
        window = FakeDialogWindow()
        enabled = FakeVar(True)
        dialed = []
        buttons = []

        class FakeButton(FakeWidget):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                buttons.append(self)

        with patch(
            "sms_ui.serial_debug_sms_call_dialogs.create_debug_dialog",
            return_value=window,
        ), patch(
            "sms_ui.serial_debug_sms_call_dialogs.tk.StringVar",
            return_value=FakeVar("15923240141"),
        ), patch(
            "sms_ui.serial_debug_sms_call_dialogs.ttk.Frame",
            FakeWidget,
        ), patch(
            "sms_ui.serial_debug_sms_call_dialogs.ttk.Label",
            FakeWidget,
        ), patch(
            "sms_ui.serial_debug_sms_call_dialogs.ttk.Entry",
            FakeWidget,
        ), patch(
            "sms_ui.serial_debug_sms_call_dialogs.tk.Label",
            FakeWidget,
        ), patch(
            "sms_ui.serial_debug_sms_call_dialogs.ttk.Button",
            FakeButton,
        ), patch(
            "sms_ui.serial_debug_sms_call_dialogs.show_centered_debug_dialog",
        ):
            open_dial_dialog(
                "parent",
                enabled,
                dialed.append,
                lambda: None,
                lambda: None,
                lambda *_args: None,
            )

            dial_button = next(button for button in buttons if "拨号" in button.kwargs["text"])
            dial_button.kwargs["command"]()

        self.assertEqual(dialed, ["15923240141"])
        self.assertTrue(window.grab_released)
        self.assertTrue(window.destroyed)


if __name__ == "__main__":
    unittest.main()
