import unittest
from unittest.mock import patch

from sms_ui.call_popup import open_call_popup, open_dial_call_popup


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


class TimerWindow(FakeWidget):
    def __init__(self, events):
        super().__init__()
        self.events = events
        self.callbacks = {}
        self.cancelled = []
        self.next_after_id = 0

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
        self.protocol_callback = callback

    def deiconify(self):
        self.events.append("deiconify")

    def lift(self):
        self.events.append("lift")

    def after(self, delay, callback):
        self.next_after_id += 1
        self.callbacks[self.next_after_id] = callback
        return self.next_after_id

    def after_cancel(self, after_id):
        self.cancelled.append(after_id)
        self.callbacks.pop(after_id, None)


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
    def test_dial_popup_shows_duration_and_hangup_callback(self):
        events = []
        buttons = []
        labels = []
        window = TimerWindow(events)
        clock = iter((100.0, 101.2))
        callbacks = {}

        class FakeLabel(FakeWidget):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self.configured = []
                self.pack_calls = []
                labels.append(self)

            def config(self, *args, **kwargs):
                self.configured.append(kwargs)

            def pack(self, *args, **kwargs):
                self.pack_calls.append((args, kwargs))

        class FakeButton(FakeWidget):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                buttons.append(self)

        with patch("sms_ui.call_popup.time.monotonic", side_effect=lambda: next(clock)), \
                patch("sms_ui.call_popup.tk.Toplevel", return_value=window), \
                patch("sms_ui.call_popup.ttk.Frame", FakeWidget), \
                patch("sms_ui.call_popup.tk.Label", FakeLabel), \
                patch("sms_ui.call_popup.ttk.Button", FakeButton):
            result = open_dial_call_popup(
                object(),
                "15923240141",
                lambda *_args: None,
                lambda: (callbacks.__setitem__("hangup", True), window._call_popup_cleanup()),
                lambda: events.append("closed"),
            )

            self.assertIs(result, window)
            self.assertEqual(labels[2].kwargs["text"], "00:00")
            self.assertEqual(labels[2].pack_calls, [])
            window._call_popup_mark_connected()
            self.assertEqual(labels[0].configured[0]["text"], "📞 通话中...")
            self.assertEqual(labels[2].configured[0]["text"], "00:00")
            self.assertEqual(labels[2].configured[1]["text"], "00:01")
            self.assertEqual(len(labels[2].pack_calls), 1)
            self.assertEqual(len(buttons), 1)
            self.assertIn("挂断", buttons[0].kwargs["text"])
            buttons[0].kwargs["command"]()

        self.assertTrue(callbacks.get("hangup", False))
        self.assertTrue(window.cancelled)

    def test_dial_popup_hangup_uses_no_argument_callback(self):
        events = []
        buttons = []
        window = FakeWindow(events)

        class FakeButton(FakeWidget):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                buttons.append(self)

        with patch("sms_ui.call_popup.tk.Toplevel", return_value=window), \
                patch("sms_ui.call_popup.ttk.Frame", FakeWidget), \
                patch("sms_ui.call_popup.tk.Label", FakeWidget), \
                patch("sms_ui.call_popup.ttk.Button", FakeButton):
            open_dial_call_popup(
                object(),
                "15923240141",
                lambda *_args: None,
                lambda: events.append("hangup"),
                lambda: events.append("closed"),
            )

        events.clear()
        buttons[0].kwargs["command"]()
        self.assertEqual(events, ["hangup"])

    def test_dial_popup_shows_terminal_status_and_can_be_closed(self):
        events = []
        buttons = []
        window = TimerWindow(events)

        class FakeButton(FakeWidget):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                buttons.append(self)

            def config(self, *args, **kwargs):
                self.kwargs.update(kwargs)

        class FakeLabel(FakeWidget):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self.configured = []

            def config(self, *args, **kwargs):
                self.configured.append(kwargs)

        with patch("sms_ui.call_popup.tk.Toplevel", return_value=window), \
                patch("sms_ui.call_popup.ttk.Frame", FakeWidget), \
                patch("sms_ui.call_popup.tk.Label", FakeLabel), \
                patch("sms_ui.call_popup.ttk.Button", FakeButton):
            open_dial_call_popup(
                object(),
                "15923240141",
                lambda *_args: None,
                lambda: events.append("hangup"),
                lambda: events.append("closed"),
            )
            window._call_popup_mark_ended("📞 对方已挂断")

        events.clear()
        self.assertEqual(buttons[0].kwargs["text"], "关闭")
        self.assertEqual(buttons[0].kwargs["state"], "normal")
        self.assertIn("command", buttons[0].kwargs)
        buttons[0].kwargs["command"]()
        self.assertEqual(events, ["closed"])

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

    def test_third_action_is_labeled_as_hide(self):
        events = []
        buttons = []
        window = FakeWindow(events)

        class FakeButton(FakeWidget):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                buttons.append(self)

        with patch("sms_ui.call_popup.tk.Toplevel", return_value=window), \
                patch("sms_ui.call_popup.ttk.Frame", FakeWidget), \
                patch("sms_ui.call_popup.tk.Label", FakeWidget), \
                patch("sms_ui.call_popup.ttk.Button", FakeButton):
            open_call_popup(
                object(),
                "10086",
                lambda *_args: None,
                lambda *_args: None,
                lambda *_args: None,
                lambda: None,
                lambda: None,
            )

        self.assertEqual(buttons[2].kwargs["text"], "隐藏")

    def test_duration_starts_when_answer_is_connected_and_refreshes(self):
        events = []
        buttons = []
        labels = []
        window = TimerWindow(events)
        clock = iter((100.0, 101.2))

        class FakeLabel(FakeWidget):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                self.configured = []
                labels.append(self)

            def config(self, *args, **kwargs):
                self.configured.append(kwargs)

        class FakeButton(FakeWidget):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                buttons.append(self)

        with patch("sms_ui.call_popup.time.monotonic", side_effect=lambda: next(clock)), \
                patch("sms_ui.call_popup.tk.Toplevel", return_value=window), \
                patch("sms_ui.call_popup.ttk.Frame", FakeWidget), \
                patch("sms_ui.call_popup.tk.Label", FakeLabel), \
                patch("sms_ui.call_popup.ttk.Button", FakeButton):
            open_call_popup(
                object(),
                "10086",
                lambda *_args: None,
                lambda mark_connected, _restore: mark_connected(),
                lambda *_args: None,
                lambda: None,
                lambda: None,
            )
            buttons[0].kwargs["command"]()
        duration_label = labels[2]
        self.assertEqual(duration_label.configured[0]["text"], "00:00")
        self.assertEqual(duration_label.configured[1]["text"], "00:01")
        self.assertEqual(len(window.callbacks), 1)

    def test_duration_timer_is_cancelled_when_popup_closes(self):
        events = []
        buttons = []
        window = TimerWindow(events)

        class FakeButton(FakeWidget):
            def __init__(self, *args, **kwargs):
                super().__init__(*args, **kwargs)
                buttons.append(self)

        with patch("sms_ui.call_popup.tk.Toplevel", return_value=window), \
                patch("sms_ui.call_popup.ttk.Frame", FakeWidget), \
                patch("sms_ui.call_popup.tk.Label", FakeWidget), \
                patch("sms_ui.call_popup.ttk.Button", FakeButton), \
                patch("sms_ui.call_popup.time.monotonic", return_value=100.0):
            open_call_popup(
                object(),
                "10086",
                lambda *_args: None,
                lambda mark_connected, _restore: mark_connected(),
                lambda *_args: None,
                lambda: None,
                lambda: None,
            )
            buttons[0].kwargs["command"]()
        after_id = next(iter(window.callbacks))
        window._call_popup_cleanup()
        self.assertEqual(window.cancelled, [after_id])


if __name__ == "__main__":
    unittest.main()
