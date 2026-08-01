import unittest
from unittest.mock import MagicMock, patch

import sms_ui.serial_debug_identity_dialogs as dialogs


class EnabledVar:
    def get(self):
        return True


class SerialDebugIdentityDialogTests(unittest.TestCase):
    def _open_information_center_dialog(self, value):
        parent = object()
        window = MagicMock()
        string_var = MagicMock()
        string_var.get.return_value = value
        widget = MagicMock()
        captured = {}

        def capture_finish(_win, _center, _parent, _entry, submit):
            captured["submit"] = submit

        patches = (
            patch.object(dialogs, "create_debug_dialog", return_value=window),
            patch.object(dialogs.tk, "StringVar", return_value=string_var),
            patch.object(dialogs.tk, "Label", return_value=widget),
            patch.object(dialogs.ttk, "Frame", return_value=widget),
            patch.object(dialogs.ttk, "Label", return_value=widget),
            patch.object(dialogs.ttk, "Entry", return_value=widget),
            patch.object(dialogs, "finish_debug_dialog", side_effect=capture_finish),
        )
        for item in patches:
            item.start()
            self.addCleanup(item.stop)

        quick_send = MagicMock()
        dialogs.open_modify_information_center_dialog(
            parent,
            EnabledVar(),
            quick_send,
            lambda *_args: None,
        )
        return window, quick_send, captured["submit"]

    def test_invalid_information_center_number_stays_open_and_does_not_send(self):
        window, quick_send, submit = self._open_information_center_dialog(
            "+8613800100500\r\nAT+RESET"
        )

        with patch.object(dialogs.messagebox, "showerror") as show_error:
            submit()

        show_error.assert_called_once()
        self.assertIn("7-15 位数字", show_error.call_args.args[1])
        window.destroy.assert_not_called()
        quick_send.assert_not_called()

    def test_valid_information_center_number_is_normalized_before_send(self):
        window, quick_send, submit = self._open_information_center_dialog(
            " +86 (1380) 010-0500 "
        )

        submit()

        window.destroy.assert_called_once()
        quick_send.assert_called_once_with('AT+CSCA="+8613800100500",145')


if __name__ == "__main__":
    unittest.main()
