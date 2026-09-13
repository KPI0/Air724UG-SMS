"""Execute the real desktop SMS editor submit path using synthetic widgets."""
from contextlib import ExitStack
import unittest
from unittest.mock import patch

from sms_core.sms_pdu import measure_text_sms_pdus
from sms_ui import serial_debug_sms_call_dialogs as dialogs


class SmsEditorTextTests(unittest.TestCase):
    def submit(self, body, phone="  10086  "):
        buttons, sent = [], []
        class Var:
            def __init__(self, value=phone, **kwargs): self.value = value
            def get(self): return self.value
            def set(self, value): self.value = value
        class Widget:
            destroyed = False
            def __init__(self, *args, **kwargs): self.kwargs = kwargs
            def pack(self, *args, **kwargs): pass
            def bind(self, *args, **kwargs): pass
            def config(self, *args, **kwargs): pass
            def get(self, *args): return body
            def title(self, *args): pass
            def resizable(self, *args): pass
            def transient(self, *args): pass
            def grab_set(self, *args): pass
            def destroy(self): self.destroyed = True
        class Button(Widget):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                buttons.append(self)
        window = Widget()
        with ExitStack() as stack:
            stack.enter_context(patch.object(dialogs, "create_debug_dialog", return_value=window))
            stack.enter_context(patch.object(dialogs, "show_centered_debug_dialog"))
            stack.enter_context(patch.object(dialogs, "ensure_debug_enabled", return_value=True))
            stack.enter_context(patch.object(dialogs.tk, "StringVar", Var))
            for provider, name in ((dialogs.tk, "Text"), (dialogs.tk, "Label"), (dialogs.ttk, "Frame"), (dialogs.ttk, "Label"), (dialogs.ttk, "Entry")):
                stack.enter_context(patch.object(provider, name, Widget))
            stack.enter_context(patch.object(dialogs.ttk, "Button", Button))
            error = stack.enter_context(patch.object(dialogs.messagebox, "showerror"))
            measure = stack.enter_context(patch.object(dialogs, "measure_text_sms_pdus", wraps=measure_text_sms_pdus))
            dialogs.open_send_sms_dialog(None, None, lambda *args: sent.append(args), lambda *_: None)
            next(button for button in buttons if button.kwargs["text"] == "发送指令").kwargs["command"]()
        return sent, error, measure, window

    def test_send_preserves_spaces_newlines_and_uses_same_text_for_measurement(self):
        for body in ("alpha\nbeta", "  alpha\nbeta  ", "\nalpha\n", "\t正文\t", " " + "A" * 69 + " "):
            with self.subTest(body=body):
                sent, error, measure, window = self.submit(body)
                self.assertEqual(sent, [("10086", body)])
                measure.assert_called_once_with(body)
                error.assert_not_called()
                self.assertTrue(window.destroyed)

    def test_blank_body_or_number_is_rejected_without_closing_editor(self):
        for phone, body in (("10086", ""), ("10086", " \t\n\u3000"), ("   ", "content")):
            with self.subTest(phone=phone, body=body):
                sent, error, measure, window = self.submit(body, phone)
                self.assertEqual(sent, [])
                error.assert_called_once()
                measure.assert_not_called()
                self.assertFalse(window.destroyed)

    def test_segment_limit_counts_edge_whitespace(self):
        body = " " + "A" * 17085
        sent, error, measure, window = self.submit(body)
        self.assertEqual(sent, [])
        error.assert_called_once()
        measure.assert_called_once_with(body)
        self.assertFalse(window.destroyed)
