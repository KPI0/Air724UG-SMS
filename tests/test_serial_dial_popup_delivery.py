import queue
import threading
import unittest
from contextlib import ExitStack
from types import SimpleNamespace
from unittest.mock import patch

from sms_core.serial_app_runtime import run_serial_reader_namespace_runtime
from sms_core.serial_reconnect import apply_serial_disconnect_effects
from sms_ui.call_popup_namespace_runtime import (
    finish_dial_popup_namespace_runtime,
    mark_dial_popup_connected_namespace_runtime,
)
from sms_ui.serial_debug_window import open_serial_debug_window_dialog
from tests.test_serial_debug_window import (
    FakeFinder, FakePauseController, FakeVar, FakeWidget, FakeWindow,
)


def reader_namespace(terminal):
    markers = []
    namespace = {name: lambda *_args, **_kwargs: None for name in (
        "enqueue_third_push", "_cloud_send_sms_event", "port_ui", "play_alert",
        "show_sms_popup", "system_ui", "_push_serial_debug", "_cloud_send_serial_log",
        "_maybe_capture_cloud_device_imei", "set_temperature", "set_signal", "set_status",
        "close_call_popup", "send_call_hangup_command", "show_call_popup",
        "schedule_delayed_connected_log", "serial_error_ui", "safe_close_serial",
    )}
    namespace.update(
        KEYWORDS=[], LOG_UNMATCHED_SMS=True, LOG_DIR="unused", LOG_PREFIX="TEST",
        ERROR_REPEAT_LIMIT=3, CALL_FILTER_MODE="Disabled", CALL_WHITELIST=[], CALL_BLACKLIST=[],
        FILE_LOG_Q=queue.Queue(), current_call_popup=None, _sms_ignore_repeat_state={},
        serial_running=True, PORT="TEST", BAUD=115200, RECONNECT_INTERVAL=0,
        serial_stop_event=threading.Event(), serial_wakeup_event=threading.Event(),
        serial_obj=object(), ring_timeout_target=0.0, current_dial_num="",
        ui_post=lambda callback: callback(), run_on_ui_thread=lambda callback, _post: callback(),
        current_dial_popup=SimpleNamespace(
            winfo_exists=lambda: True,
            _call_popup_mark_connected=lambda: markers.append("connected"),
            _call_popup_mark_ended=lambda _message: markers.append("ended"),
        ),
    )
    namespace["mark_dial_connected"] = lambda: mark_dial_popup_connected_namespace_runtime(namespace)
    namespace["finish_dial_call"] = lambda message: finish_dial_popup_namespace_runtime(namespace, message)
    namespace["get_serial_call_state"] = lambda: (namespace["ring_timeout_target"], namespace["current_dial_num"])
    namespace["set_serial_call_state"] = lambda ring, number: namespace.update(ring_timeout_target=ring, current_dial_num=number)
    namespace["set_serial_log_prefix"] = lambda value: namespace.update(LOG_PREFIX=value)
    namespace["resolve_serial_target_port"] = lambda: "TEST"
    namespace["try_manual_rebind_after_error"] = lambda _error: False
    namespace["open_and_initialize_serial"] = lambda _port: namespace.update(current_dial_num="10086")
    frames = iter((b"CONNECT\r\n", terminal.encode("ascii") + b"\r\n"))

    def read():
        result = next(frames)
        if result != b"CONNECT\r\n":
            namespace["serial_running"] = False
        return result

    namespace["read_serial_line_safely"] = read
    return namespace, markers


class SerialDialPopupDeliveryTests(unittest.TestCase):
    def test_formal_serial_reader_delivers_outgoing_terminals_to_popup(self):
        for terminal in ("NO CARRIER", "NO ANSWER", "BUSY", '+CIEV: "CALL",0', "+CME ERROR: 3"):
            with self.subTest(terminal=terminal):
                namespace, markers = reader_namespace(terminal)
                run_serial_reader_namespace_runtime(
                    namespace,
                    parse_callback_head=lambda value: ("", value),
                    apply_disconnect_effects=apply_serial_disconnect_effects,
                )
                self.assertEqual(markers, ["connected", "ended"])
                self.assertEqual(namespace["current_dial_num"], "")

    def test_queued_outgoing_markers_do_not_follow_a_replaced_popup(self):
        for callback in (mark_dial_popup_connected_namespace_runtime, finish_dial_popup_namespace_runtime):
            with self.subTest(callback=callback.__name__):
                namespace, old_markers = reader_namespace("NO CARRIER")
                queued = []
                namespace["run_on_ui_thread"] = lambda callback, _post: queued.append(callback)
                callback(namespace)
                new_markers = []
                namespace["current_dial_popup"] = SimpleNamespace(
                    winfo_exists=lambda: True,
                    _call_popup_mark_connected=lambda: new_markers.append("connected"),
                    _call_popup_mark_ended=lambda _message: new_markers.append("ended"),
                )
                queued.pop()()
                self.assertEqual(new_markers, [])
                self.assertEqual(old_markers, [])

    def test_queued_outgoing_markers_keep_the_original_live_popup(self):
        for callback, expected in ((mark_dial_popup_connected_namespace_runtime, "connected"),
                                   (finish_dial_popup_namespace_runtime, "ended")):
            with self.subTest(callback=callback.__name__):
                namespace, markers = reader_namespace("NO CARRIER")
                queued = []
                namespace["run_on_ui_thread"] = lambda callback, _post: queued.append(callback)
                callback(namespace)
                self.assertEqual(markers, [])
                queued.pop()()
                self.assertEqual(markers, [expected])

    def test_immediate_modem_replies_find_the_new_same_number_dial_popup(self):
        namespace, old_markers = reader_namespace("NO CARRIER")
        namespace["current_dial_popup"].destroy = lambda: None
        new_markers = []
        popup = SimpleNamespace(
            winfo_exists=lambda: True,
            _call_popup_mark_connected=lambda: new_markers.append("connected"),
            _call_popup_mark_ended=lambda _message: new_markers.append("ended"),
        )
        captured = {}

        def send(_command, **_kwargs):
            mark_dial_popup_connected_namespace_runtime(namespace)
            finish_dial_popup_namespace_runtime(namespace, "ended")

        with ExitStack() as stack:
            prefix = "sms_ui.serial_debug_window."
            for name, replacement in (
                ("tk.BooleanVar", FakeVar), ("tk.StringVar", FakeVar),
                ("ttk.Frame", FakeWidget), ("ttk.Checkbutton", FakeWidget),
                ("ttk.Label", FakeWidget), ("ttk.Entry", FakeWidget), ("ttk.Button", FakeWidget),
                ("SerialDebugPauseController", FakePauseController), ("SerialDebugFinder", FakeFinder),
            ):
                stack.enter_context(patch(prefix + name, replacement))
            stack.enter_context(patch(prefix + "tk.Toplevel", return_value=FakeWindow()))
            stack.enter_context(patch(prefix + "create_serial_debug_body", return_value=(FakeWidget(), FakeWidget(), FakeWidget())))
            stack.enter_context(patch(prefix + "create_serial_debug_quick_actions", side_effect=lambda *args: captured.update(dial=args[8])))
            stack.enter_context(patch(prefix + "open_dial_call_popup", return_value=popup))
            stack.enter_context(patch(prefix + "start_serial_debug_append_loop"))
            open_serial_debug_window_dialog(
                "parent", None, None, True, lambda: 0, queue.Queue(), threading.Lock(),
                lambda: None, lambda *_args: None, lambda *_args: None, lambda *_args: None,
                lambda port: str(port), lambda: "TEST", lambda number: namespace.update(current_dial_num=number),
                lambda *_args: None, lambda *_args: None, lambda: None, lambda *_args: None,
                get_dial_popup=lambda: namespace["current_dial_popup"],
                set_dial_popup=lambda value: namespace.update(current_dial_popup=value),
                send_local_command=send,
            )
            captured["dial"]("10086")
        self.assertEqual(old_markers, [])
        self.assertEqual(new_markers, ["connected", "ended"])


if __name__ == "__main__":
    unittest.main()
