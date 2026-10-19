import queue
import unittest
from contextlib import ExitStack
from functools import partial
from types import SimpleNamespace
from unittest.mock import patch

from sms_app.serial_namespace_bindings import install_serial_namespace_bindings
from sms_core.call_command_context import call_state_lock, observe_call_command_line
from sms_core.call_session import IncomingCallSessionTracker
from sms_core.call_send_namespace_runtime import send_local_serial_command_namespace_runtime
from sms_core.serial_app_runtime import (
    run_serial_app_from_wiring, run_serial_app_runtime, run_serial_reader_namespace_runtime,
)
from sms_core.serial_reconnect import apply_serial_disconnect_effects
from sms_core.serial_runtime import run_serial_runtime_thread
from sms_core.namespace_binding import make_namespace_runtime_binder
from sms_ui.app_ui_namespace_runtime import set_status_namespace_runtime
from sms_ui.call_popup_namespace_runtime import (
    finish_dial_popup_namespace_runtime, get_serial_call_state_namespace_runtime,
    mark_dial_popup_connected_namespace_runtime, merge_serial_call_state_namespace_runtime,
)
from sms_ui.serial_debug_namespace_runtime import set_serial_debug_state_namespace_runtime
from sms_ui.serial_debug_window import open_serial_debug_window_dialog
from tests.test_local_outgoing_calls import namespace_fixture
from tests.test_serial_debug_window import (
    FakeFinder, FakePauseController, FakeVar, FakeWidget, FakeWindow,
)
from tests.test_serial_dial_popup_delivery import reader_namespace


class SerialDialGenerationTests(unittest.TestCase):
    def setUp(self):
        self.namespace, _markers = reader_namespace("NO CARRIER")
        self.namespace.update(namespace_fixture("OK"))
        self.namespace["current_dial_popup"] = None
        self.ui = []
        self.workers = []
        self.popups = []
        self.pending = False
        self.state_hook = lambda *_args, **_kwargs: None
        self.port_hook = lambda *_args, **_kwargs: None
        self.line_hook = lambda *_args, **_kwargs: None
        self.status = []
        self.call_states = []
        self.raw_loop = None
        ns = self.namespace
        ns.update(
            run_on_ui_thread=lambda callback, _post: self.ui.append(callback),
            get_serial_call_state=partial(get_serial_call_state_namespace_runtime, ns),
            merge_serial_call_state=partial(merge_serial_call_state_namespace_runtime, ns),
            serial_call_state_guard=lambda: call_state_lock(ns),
            mark_dial_connected=partial(mark_dial_popup_connected_namespace_runtime, ns),
            finish_dial_call=partial(finish_dial_popup_namespace_runtime, ns),
            _cloud_send_call_state=lambda *args, **kwargs: (
                self.call_states.append((args, kwargs)), self.state_hook(*args, **kwargs)),
            port_ui=lambda *args, **kwargs: self.port_hook(*args, **kwargs),
            _push_serial_debug=lambda *args, **kwargs: self.line_hook(*args, **kwargs),
            tk_alive=lambda: True,
            status_var=SimpleNamespace(set=lambda value: self.status.append(value)),
            status_label=FakeWidget(),
            INCOMING_CALL_SESSION=IncomingCallSessionTracker(),
        )
        original = dict(ns)
        install_serial_namespace_bindings(ns)
        formal = {key: ns[key] for key in (
            "get_serial_call_state", "merge_serial_call_state", "serial_call_state_guard",
            "mark_dial_connected", "finish_dial_call",
            "start_incoming_call_session", "start_serial_incoming_call",
            "finish_incoming_call_session", "reset_incoming_call_session",
        )}
        ns.clear()
        ns.update(original)
        ns.update(formal)
        ns["set_status"] = make_namespace_runtime_binder(ns, globals())("set_status_namespace_runtime")

    def drain_ui(self):
        while self.ui:
            self.ui.pop(0)()

    def create_popup(self, *_args):
        popup = SimpleNamespace(markers=[], exists=True)
        popup.winfo_exists = lambda: popup.exists
        popup.destroy = lambda: setattr(popup, "exists", False)
        popup._call_popup_mark_connected = lambda: popup.markers.append("connected")
        popup._call_popup_mark_ended = lambda _message: popup.markers.append("ended")
        self.popups.append(popup)
        return popup

    def send_local(self, command, **kwargs):
        def start(_name, target, **_kwargs):
            if self.pending:
                self.workers.append(target)
                return object()
            return target()

        return send_local_serial_command_namespace_runtime(
            self.namespace, command, **kwargs, start_worker=start, response_timeout=0.01,
        )

    def run_scenario(self, scenario):
        ns = self.namespace
        captured = {}
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
            stack.enter_context(patch(prefix + "open_dial_call_popup", side_effect=self.create_popup))
            stack.enter_context(patch(prefix + "start_serial_debug_append_loop"))
            open_serial_debug_window_dialog(
                "parent", None, None, True, lambda: 0, queue.Queue(), ns["serial_lock"],
                lambda: ns["serial_obj"], lambda *_args: None, lambda *_args: None, ns["set_status"],
                str, lambda: "TEST", lambda number: set_serial_debug_state_namespace_runtime(ns, "current_dial_num", number),
                lambda *_args: None, lambda *_args: None, lambda: None, lambda *_args: None,
                get_dial_popup=lambda: ns["current_dial_popup"],
                set_dial_popup=lambda popup: set_serial_debug_state_namespace_runtime(ns, "dial_popup", popup),
                send_local_command=self.send_local,
            )

            def run_loop(**kwargs):
                if self.raw_loop is None:
                    scenario(kwargs["handle_line"], captured["dial"])
                else:
                    self.raw_loop(kwargs, captured["dial"])

            return run_serial_reader_namespace_runtime(
                ns, parse_callback_head=lambda value: ("", value),
                apply_disconnect_effects=apply_serial_disconnect_effects,
                run_reader=partial(run_serial_app_from_wiring, run_app=partial(
                    run_serial_app_runtime, run_runtime_thread=partial(
                        run_serial_runtime_thread, run_loop=run_loop, clock=lambda: 100.0))),
            )

    def test_old_terminal_decision_does_not_end_or_clear_same_number_redial(self):
        def scenario(handle, dial):
            dial("10086")
            handle("CONNECT")
            self.drain_ui()
            self.state_hook = lambda *_args, **_kwargs: dial("10086")
            handle("NO CARRIER")
            self.state_hook = lambda *_args, **_kwargs: None
            self.drain_ui()
            self.assertEqual(self.namespace["current_dial_num"], "10086")
            self.assertEqual(self.popups[-1].markers, [])
            self.assertIn("呼叫中", self.status[-1])
            handle("CONNECT")
            self.drain_ui()
            self.assertEqual(self.popups[-1].markers, ["connected"])
            handle("NO CARRIER")
            self.drain_ui()
            self.assertEqual(self.popups[-1].markers, ["connected", "ended"])
            self.assertEqual(self.namespace["current_dial_num"], "")
        self.run_scenario(scenario)

    def test_old_connect_decision_does_not_connect_same_number_redial(self):
        def scenario(handle, dial):
            dial("10086")
            def redial(message, *_args):
                if "对方已接听" in message:
                    self.port_hook = lambda *_args: None
                    dial("10086")
            self.port_hook = redial
            handle("CONNECT")
            self.drain_ui()
            self.assertEqual(self.popups[-1].markers, [])
            self.assertIn("呼叫中", self.status[-1])
            self.assertEqual(self.namespace["current_dial_num"], "10086")
            handle("CONNECT")
            self.drain_ui()
            self.assertEqual(self.popups[-1].markers, ["connected"])
        self.run_scenario(scenario)

    def test_redial_before_old_terminal_decision_preserves_new_call(self):
        def scenario(handle, dial):
            dial("10086")
            def redial(line, *_args):
                if line == "NO CARRIER":
                    self.line_hook = lambda *_args: None
                    dial("10086")
            self.line_hook = redial
            handle("NO CARRIER")
            self.drain_ui()
            self.assertEqual(self.namespace["current_dial_num"], "10086")
            self.assertEqual(self.popups[-1].markers, [])
            handle("CONNECT")
            self.drain_ui()
            self.assertEqual(self.popups[-1].markers, ["connected"])
        self.run_scenario(scenario)

    def test_same_request_receives_valid_responses_after_queued_dial_commits(self):
        def scenario(handle, dial):
            self.pending = True
            dial("10086")
            handle("OK")
            self.assertEqual(self.namespace["current_dial_num"], "10086")
            self.assertEqual(self.namespace["serial_obj"].writes, [])
            self.workers.pop()()
            handle("CONNECT")
            self.drain_ui()
            self.assertEqual(self.popups[-1].markers, ["connected"])
            handle("NO ANSWER")
            self.drain_ui()
            self.assertEqual(self.popups[-1].markers, ["connected", "ended"])
        self.run_scenario(scenario)

    def test_submission_after_snapshot_cannot_be_revoked_by_pre_submission_terminal(self):
        def scenario(handle, dial):
            self.pending = True
            dial("10086")

            def commit_before_observe(namespace, line, **kwargs):
                self.workers.pop()()
                observe_call_command_line(namespace, line, **kwargs)

            with patch("sms_core.serial_app_runtime.observe_call_command_line", side_effect=commit_before_observe):
                handle("NO CARRIER")
            self.assertIsNotNone(self.namespace.get("_cloud_dial_lifetime"))
            self.assertEqual(self.namespace["current_dial_num"], "10086")
            self.drain_ui()
            self.assertEqual(self.popups[-1].markers, [])
            handle("CONNECT")
            self.drain_ui()
            self.assertEqual(self.popups[-1].markers, ["connected"])
        self.run_scenario(scenario)

    def test_submission_during_old_line_processing_keeps_current_call_state(self):
        def scenario(handle, dial):
            self.pending = True
            dial("10086")
            def submit(line, *_args):
                if line == "OK":
                    self.line_hook = lambda *_args: None
                    self.workers.pop()()
            self.line_hook = submit
            handle("OK")
            self.assertEqual(self.namespace["current_dial_num"], "10086")
            handle("CONNECT")
            self.drain_ui()
            self.assertEqual(self.popups[-1].markers, ["connected"])
        self.run_scenario(scenario)

    def test_pre_submission_terminal_cannot_mark_pending_request(self):
        def scenario(handle, dial):
            self.pending = True
            dial("10086")
            handle("NO CARRIER")
            self.drain_ui()
            self.assertEqual(self.popups[-1].markers, [])
            self.assertEqual(self.namespace["current_dial_num"], "10086")
            self.workers.pop()()
            self.drain_ui()
            self.assertEqual(self.namespace["serial_obj"].writes, [])
            self.assertEqual(self.namespace["current_dial_num"], "")
            self.assertEqual(self.popups[-1].markers, ["ended"])
        self.run_scenario(scenario)

    def test_queued_dial_failure_updates_its_window_and_clears_its_state(self):
        def scenario(_handle, dial):
            self.pending = True
            dial("10086")
            self.namespace["serial_connection_generation"] += 1
            self.workers.pop()()
            self.drain_ui()
            self.assertEqual(self.namespace["serial_obj"].writes, [])
            self.assertEqual(self.namespace["current_dial_num"], "")
            self.assertEqual(self.popups[-1].markers, ["ended"])
        self.run_scenario(scenario)

    def test_old_queued_failure_cannot_clear_or_end_new_pending_request(self):
        def scenario(_handle, dial):
            self.pending = True
            dial("10086")
            dial("10086")
            self.workers.pop(0)()
            self.drain_ui()
            self.assertEqual(self.namespace["current_dial_num"], "10086")
            self.assertEqual(self.popups[-1].markers, [])
            self.workers.pop(0)()
            self.assertEqual(self.namespace["serial_obj"].writes, [b"ATD10086;\r\n"])
        self.run_scenario(scenario)

    def test_modem_rejection_and_write_failure_finish_current_popup(self):
        for response in ("ERROR", "+CME ERROR: 3", "write_failure"):
            with self.subTest(response=response):
                self.setUp()
                def scenario(_handle, dial):
                    serial = self.namespace["serial_obj"]
                    serial.response = response
                    if response == "write_failure":
                        serial.write = lambda _payload: (_ for _ in ()).throw(OSError("synthetic write failure"))
                    dial("10086")
                    self.drain_ui()
                    self.assertEqual(self.namespace["current_dial_num"], "")
                    self.assertEqual(self.popups[-1].markers, ["ended"])
                self.run_scenario(scenario)

    def test_hidden_submitted_popup_keeps_valid_callbacks(self):
        def scenario(handle, dial):
            dial("10086")
            popup = self.popups[-1]
            popup.winfo_viewable = lambda: False
            handle("CONNECT")
            handle("NO CARRIER")
            self.drain_ui()
            self.assertEqual(popup.markers, ["connected", "ended"])
        self.run_scenario(scenario)

    def test_old_same_number_snapshot_cannot_clear_redial_state(self):
        def scenario(_handle, dial):
            dial("10086")
            previous = self.namespace["get_serial_call_state"]()
            dial("10086")
            self.namespace["ring_timeout_target"] = -1.0
            merged = self.namespace["merge_serial_call_state"](previous, 0.0, "")
            self.assertEqual(tuple(merged), (-1.0, "10086"))
        self.run_scenario(scenario)

    def test_read_terminal_keeps_original_snapshot_across_config_loading(self):
        def raw_loop(kwargs, dial):
            ns = self.namespace
            frames = iter((b"NO CARRIER\r\n", b"CONNECT\r\n", b""))
            read_old = False
            def read():
                nonlocal read_old
                raw = next(frames)
                if raw == b"NO CARRIER\r\n":
                    read_old = True
                if not raw:
                    ns["serial_running"] = False
                return raw

            ns["read_serial_line_safely"] = read
            ns["open_and_initialize_serial"] = lambda _port: dial("10086")
            from sms_core import serial_app_runtime
            original_config = serial_app_runtime.build_serial_runtime_config
            fired = False

            def config(settings):
                nonlocal fired
                if read_old and not fired:
                    fired = True
                    dial("10086")
                return original_config(settings)

            with patch.object(serial_app_runtime, "build_serial_runtime_config", side_effect=config):
                run_serial_reader_namespace_runtime(
                    ns, parse_callback_head=lambda value: ("", value),
                    apply_disconnect_effects=apply_serial_disconnect_effects,
                )
            self.drain_ui()
            self.assertTrue(fired)
            self.assertEqual(self.namespace["current_dial_num"], "10086")
            self.assertEqual(self.popups[-1].markers, ["connected"])

        self.raw_loop = raw_loop
        self.run_scenario(lambda *_args: None)

    def run_read_chunks(self, chunks, *, before_read=None):
        def raw_loop(_kwargs, dial):
            ns = self.namespace
            frames = iter(enumerate(chunks))

            def read():
                index, raw = next(frames)
                if before_read is not None:
                    before_read(index, dial)
                if not raw:
                    ns["serial_running"] = False
                return raw

            ns["read_serial_line_safely"] = read
            ns["open_and_initialize_serial"] = lambda _port: dial("10086")
            run_serial_reader_namespace_runtime(
                ns, parse_callback_head=lambda value: ("", value),
                apply_disconnect_effects=apply_serial_disconnect_effects,
            )
            self.drain_ui()

        self.raw_loop = raw_loop
        self.run_scenario(lambda *_args: None)

    def test_one_chunk_terminal_then_ok_does_not_restore_old_dial_number(self):
        self.run_read_chunks((b"NO CARRIER\r\nOK\r\n", b""))
        self.assertEqual(self.namespace["current_dial_num"], "")
        self.assertEqual(self.popups[-1].markers, ["ended"])

    def test_one_chunk_connect_terminal_and_late_connect_stays_ended(self):
        self.run_read_chunks((b"CONNECT\r\nNO CARRIER\r\nCONNECT\r\n", b""))
        self.assertEqual(self.namespace["current_dial_num"], "")
        self.assertEqual(self.popups[-1].markers, ["connected", "ended"])

    def test_one_chunk_outgoing_end_new_incoming_and_terminal_stay_ordered(self):
        self.run_read_chunks((b'NO CARRIER\r\n+CLIP: "10010",129\r\nCONNECT\r\nNO CARRIER\r\n', b""))
        directions = [kwargs["direction"] for _args, kwargs in self.call_states]
        self.assertEqual(directions, ["outgoing", "incoming", "incoming"])
        self.assertEqual(self.namespace["current_dial_num"], "")
        self.assertEqual(self.namespace["ring_timeout_target"], 0.0)
        self.assertIsNone(self.namespace["INCOMING_CALL_SESSION"].snapshot().session_token)

    def test_incoming_after_terminal_does_not_suppress_old_popup_cleanup(self):
        self.run_read_chunks((b'CONNECT\r\nNO CARRIER\r\n+CLIP: "10010",129\r\n', b""))
        self.assertIn("ended", self.popups[0].markers)
        self.assertIn("响铃中", self.status[-1])
        self.assertEqual(self.namespace["INCOMING_CALL_SESSION"].snapshot().caller_num, "10010")

    def test_partial_old_terminal_keeps_old_source_after_next_chunk_redial(self):
        self.run_read_chunks(
            (b"NO CARR", b"IER\r\nCONNECT\r\n", b""),
            before_read=lambda index, dial: dial("10086") if index == 1 else None,
        )
        self.assertEqual(self.namespace["current_dial_num"], "10086")
        self.assertEqual(self.popups[-1].markers, ["connected"])


if __name__ == "__main__":
    unittest.main()
