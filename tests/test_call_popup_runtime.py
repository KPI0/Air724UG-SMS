import unittest
from types import SimpleNamespace

from sms_ui.call_popup_runtime import (
    close_call_popup_app_runtime,
    close_call_popup_runtime,
    popup_exists,
    show_call_popup_app_runtime,
    show_call_popup_runtime,
)


class FakePopup:
    def __init__(self, exists=True):
        self.exists = exists
        self.destroyed = False

    def winfo_exists(self):
        return self.exists

    def destroy(self):
        self.destroyed = True
        self.exists = False


class BrokenDestroyPopup(FakePopup):
    def destroy(self):
        raise RuntimeError("destroy failed")


class CallPopupRuntimeTests(unittest.TestCase):
    def test_popup_exists_handles_missing_and_broken_windows(self):
        class BrokenPopup:
            def winfo_exists(self):
                raise RuntimeError("window already gone")

        self.assertFalse(popup_exists(None))
        self.assertFalse(popup_exists(BrokenPopup()))
        self.assertTrue(popup_exists(FakePopup()))

    def test_popup_exists_logs_broken_windows(self):
        class BrokenPopup:
            def winfo_exists(self):
                raise RuntimeError("window already gone")

        logs = []

        self.assertFalse(popup_exists(BrokenPopup(), log_error=logs.append))
        self.assertEqual(len(logs), 1)
        self.assertIn("window already gone", logs[0])

    def test_close_call_popup_runtime_destroys_existing_popup_and_clears_state(self):
        popup = FakePopup()
        values = []

        close_call_popup_runtime(popup, values.append)

        self.assertTrue(popup.destroyed)
        self.assertEqual(values, [None])

    def test_close_call_popup_runtime_logs_destroy_failure_and_clears_state(self):
        popup = BrokenDestroyPopup()
        values = []
        logs = []

        close_call_popup_runtime(popup, values.append, log_error=logs.append)

        self.assertEqual(values, [None])
        self.assertEqual(len(logs), 1)
        self.assertIn("destroy failed", logs[0])

    def test_close_call_popup_app_runtime_posts_close_to_ui_thread(self):
        calls = []

        result = close_call_popup_app_runtime(
            get_popup=lambda: "popup",
            set_popup=lambda value: calls.append(("set", value)),
            run_on_ui_thread=lambda callback, ui_post: calls.append(("run", ui_post)) or callback(),
            ui_post="post",
            close_runtime=lambda popup, set_popup: calls.append(("close", popup)) or set_popup(None),
        )

        self.assertIsNone(result)
        self.assertEqual(calls, [("run", "post"), ("close", "popup"), ("set", None)])

    def test_close_call_popup_app_runtime_forwards_log_error(self):
        calls = []

        result = close_call_popup_app_runtime(
            get_popup=lambda: "popup",
            set_popup=lambda value: calls.append(("set", value)),
            run_on_ui_thread=lambda callback, ui_post: callback(),
            ui_post="post",
            log_error=lambda message: ("log", message),
            close_runtime=lambda popup, set_popup, **kwargs: calls.append(("close", popup, kwargs["log_error"]("msg"))),
        )

        self.assertIsNone(result)
        self.assertEqual(calls, [("close", "popup", ("log", "msg"))])

    def test_show_call_popup_runtime_skips_when_popup_exists(self):
        current = FakePopup()
        opened = []

        result = show_call_popup_runtime(
            object(),
            "10086",
            current,
            lambda value: self.fail(f"unexpected state change: {value!r}"),
            lambda *_: None,
            object(),
            lambda: object(),
            lambda *_: None,
            lambda *_: None,
            lambda callback: callback(),
            lambda: None,
            lambda value: None,
            open_popup=lambda *args: opened.append(args),
        )

        self.assertIs(result, current)
        self.assertEqual(opened, [])

    def test_show_call_popup_runtime_opens_and_wires_answer_and_hangup(self):
        popup = FakePopup()
        opened = {}
        sent = []
        port_messages = []
        statuses = []
        posted = []
        ring_timeout = []
        closed = []
        handled = []

        def open_popup(parent, caller_num, center_window, on_answer, on_hangup, on_ignore, on_close):
            opened.update(
                parent=parent,
                caller_num=caller_num,
                on_answer=on_answer,
                on_hangup=on_hangup,
                on_ignore=on_ignore,
                on_close=on_close,
            )
            return popup

        def send_command_async(serial_lock, serial_getter, command, *, on_result):
            sent.append((serial_lock, serial_getter(), command))
            on_result(SimpleNamespace(ok=True, error=""))

        parent = object()
        serial_lock = object()
        serial_obj = object()
        state = []
        mark_connected = object()

        result = show_call_popup_runtime(
            parent,
            "10086",
            None,
            state.append,
            lambda *_: None,
            serial_lock,
            lambda: serial_obj,
            lambda *args: port_messages.append(args),
            lambda *args: statuses.append(args),
            posted.append,
            lambda: closed.append("closed"),
            ring_timeout.append,
            lambda: handled.append("handled"),
            open_popup=open_popup,
            send_command_async=send_command_async,
        )
        opened["on_answer"](mark_connected, lambda: None)
        opened["on_hangup"](lambda: None)
        opened["on_ignore"]()
        opened["on_close"]()

        self.assertIs(result, popup)
        self.assertEqual(state, [popup])
        self.assertEqual(opened["parent"], parent)
        self.assertEqual(opened["caller_num"], "10086")
        self.assertEqual(sent, [(serial_lock, serial_obj, "ATA"), (serial_lock, serial_obj, "ATH")])
        self.assertEqual(ring_timeout, [-1.0])
        self.assertEqual(posted, [mark_connected])
        self.assertEqual(closed, ["closed", "closed", "closed"])
        self.assertEqual(handled, ["handled", "handled"])
        self.assertTrue(port_messages)
        self.assertTrue(statuses)

    def test_show_call_popup_runtime_restores_buttons_on_command_failure(self):
        opened = {}
        posted = []
        handled = []

        def open_popup(parent, caller_num, center_window, on_answer, on_hangup, on_ignore, on_close):
            opened.update(on_answer=on_answer, on_hangup=on_hangup)
            return FakePopup()

        def send_command_async(serial_lock, serial_getter, command, *, on_result):
            on_result(SimpleNamespace(ok=False, error="denied"))

        show_call_popup_runtime(
            object(),
            "10086",
            None,
            lambda value: None,
            lambda *_: None,
            object(),
            lambda: object(),
            lambda *_: None,
            lambda *_: None,
            posted.append,
            lambda: None,
            lambda value: None,
            lambda: handled.append("handled"),
            open_popup=open_popup,
            send_command_async=send_command_async,
        )

        restore_answer = object()
        restore_hangup = object()
        opened["on_answer"](lambda: None, restore_answer)
        opened["on_hangup"](restore_hangup)

        self.assertEqual(posted, [restore_answer, restore_hangup])
        self.assertEqual(handled, ["handled", "handled"])

    def test_show_call_popup_app_runtime_posts_show_to_ui_thread(self):
        calls = []
        popup = object()

        result = show_call_popup_app_runtime(
            parent="root",
            caller_num="10086",
            get_popup=lambda: popup,
            set_popup=lambda value: calls.append(("set", value)),
            center_window="center",
            serial_lock="lock",
            get_serial=lambda: "serial",
            port_ui=lambda *_: None,
            set_status=lambda *_: None,
            ui_post="post",
            close_popup=lambda: None,
            set_ring_timeout=lambda value: None,
            run_on_ui_thread=lambda callback, ui_post: calls.append(("run", ui_post)) or callback(),
            show_runtime=lambda *args: calls.append(("show", args)) or "win",
        )

        self.assertEqual(result, "win")
        self.assertEqual(calls[0], ("run", "post"))
        self.assertEqual(calls[1][0], "show")
        self.assertEqual(calls[1][1][0], "root")
        self.assertEqual(calls[1][1][1], "10086")
        self.assertIs(calls[1][1][2], popup)

    def test_show_call_popup_app_runtime_rechecks_setting_on_ui_thread(self):
        queued = []
        enabled = [True]
        shown = []

        result = show_call_popup_app_runtime(
            parent="root",
            caller_num="10086",
            get_popup=lambda: None,
            set_popup=lambda value: None,
            center_window="center",
            serial_lock="lock",
            get_serial=lambda: "serial",
            port_ui=lambda *_: None,
            set_status=lambda *_: None,
            ui_post="post",
            close_popup=lambda: None,
            set_ring_timeout=lambda value: None,
            run_on_ui_thread=lambda callback, _ui_post: queued.append(callback),
            is_enabled=lambda: enabled[0],
            show_runtime=lambda *args: shown.append(args),
        )

        self.assertIsNone(result)
        enabled[0] = False
        self.assertEqual(queued[0](), "disabled")
        self.assertEqual(shown, [])


if __name__ == "__main__":
    unittest.main()
