import threading
import unittest
from contextlib import ExitStack
from types import SimpleNamespace
from unittest.mock import patch

from sms_core import serial_sender
from sms_core.call_command_context import call_state_lock, observe_call_command_line
from sms_core.call_session import IncomingCallSessionTracker
from sms_core.call_send_namespace_runtime import send_local_serial_command_namespace_runtime
from sms_ui.serial_debug_namespace_runtime import open_serial_debug_window_namespace_runtime
from tests.test_local_outgoing_calls import namespace_fixture
from tests.test_serial_debug_window import FakeFinder, FakePauseController, FakeVar, FakeWidget, FakeWindow


class Gate:
    def __init__(self, pause=1):
        self.pause = pause
        self.count = 0
        self.reached = threading.Event()
        self.resume = threading.Event()
        self.lock = threading.RLock()

    def __enter__(self):
        self.count += 1
        if self.count == self.pause or (self.pause == "worker" and threading.current_thread() is not threading.main_thread()):
            self.reached.set()
            if not self.resume.wait(3):
                raise RuntimeError("Call test gate was not released")
        return self.lock.__enter__()

    def __exit__(self, *args):
        return self.lock.__exit__(*args)


class LocalCallCommandIsolationTests(unittest.TestCase):
    def setUp(self):
        self.ns = namespace_fixture()
        self.tracker = IncomingCallSessionTracker()
        self.ns["INCOMING_CALL_SESSION"] = self.tracker
        self.ns["ring_timeout_target"] = 0.0
        self.tracker.start("synthetic-old")

    def replace_call(self):
        with call_state_lock(self.ns):
            observe_call_command_line(self.ns, "NO CARRIER")
            self.tracker.finish()
            self.tracker.start("synthetic-new")

    def paused(self, command, boundary, change):
        gate = Gate("worker" if boundary == "connection" else 3 if boundary == "final_connection" else 1)
        modem = self.ns["serial_obj"]
        with ExitStack() as stack:
            if boundary in ("connection", "final_connection"):
                self.ns["serial_lock"] = gate
            else:
                stack.enter_context(patch.object(serial_sender,
                    "DEFAULT_SERIAL_TRANSACTION_LOCK" if boundary == "transaction" else "DEFAULT_SERIAL_WRITE_LOCK", gate))
            worker = send_local_serial_command_namespace_runtime(self.ns, command, response_timeout=0.01)
            try:
                self.assertTrue(gate.reached.wait(2))
                change()
            finally:
                gate.resume.set()
                worker.join(3)
        self.assertFalse(worker.is_alive())
        self.assertEqual(self.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(), ())
        self.assertIsNone(self.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"]._active)
        return modem

    def test_old_answer_and_hangup_are_rejected_after_every_send_lock_wait(self):
        for command in ("ATA", "ATH"):
            for boundary in ("transaction", "connection", "final_connection", "write"):
                with self.subTest(command=command, boundary=boundary):
                    self.setUp()
                    modem = self.paused(command, boundary, self.replace_call)
                    self.assertEqual(modem.writes, [])
                    self.assertFalse(self.tracker.snapshot().handled)

    def test_initial_connection_lock_wait_never_rebinds_old_call_or_serial(self):
        for command in ("ATA", "ATH"):
            for change_kind in ("call", "serial", "generation"):
                with self.subTest(command=command, change_kind=change_kind):
                    self.setUp()
                    gate = Gate()
                    self.ns["serial_lock"] = gate
                    original = self.ns["serial_obj"]
                    replacement = namespace_fixture()["serial_obj"]
                    results = []
                    entry = threading.Thread(target=lambda: results.append(
                        send_local_serial_command_namespace_runtime(self.ns, command, response_timeout=0.01)))
                    entry.start()
                    try:
                        self.assertTrue(gate.reached.wait(2))
                        if change_kind == "call":
                            self.replace_call()
                        elif change_kind == "serial":
                            self.ns["serial_obj"] = replacement
                        else:
                            self.ns["serial_connection_generation"] += 1
                    finally:
                        gate.resume.set()
                        entry.join(3)
                    self.assertFalse(entry.is_alive())
                    for worker in results:
                        if worker is not None:
                            worker.join(3)
                            self.assertFalse(worker.is_alive())
                    self.assertEqual(original.writes, [])
                    self.assertEqual(replacement.writes, [])
                    self.assertEqual(self.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(), ())

    def test_serial_replacement_and_same_object_reconnection_cancel_old_controls(self):
        for command in ("ATA", "ATH"):
            for field in ("serial_obj", "serial_connection_generation"):
                with self.subTest(command=command, field=field):
                    self.setUp()
                    replacement = namespace_fixture()["serial_obj"]
                    modem = self.paused(command, "write", lambda: self.ns.update({
                        field: replacement if field == "serial_obj" else 2}))
                    self.assertEqual(modem.writes, [])
                    self.assertEqual(replacement.writes, [])

    def test_current_controls_preserve_exact_editor_bytes_without_call_history(self):
        for command, append, expected in ((" ata;\t", True, b" ata;\t\r\n"),
                                          ("ATH\r", False, b"ATH\r")):
            with self.subTest(command=command):
                self.setUp()
                worker = send_local_serial_command_namespace_runtime(self.ns, command, append_crlf=append)
                worker.join(2)
                self.assertFalse(worker.is_alive())
                self.assertEqual(self.ns["serial_obj"].writes, [expected])
                self.assertTrue(self.ns["CLOUD_SMS_EVENT_Q"].empty())

    def test_unterminated_controls_still_use_raw_editor_sender(self):
        sent = []
        for command in ("ATA", "ATH"):
            send_local_serial_command_namespace_runtime(self.ns, command, append_crlf=False,
                send_raw=lambda *args, **kwargs: sent.append((args[2], kwargs["append_crlf"])))
        self.assertEqual(sent, [("ATA", False), ("ATH", False)])

    def test_completed_outgoing_request_does_not_block_current_incoming_hangup(self):
        from sms_core.call_command_context import set_local_dial_number
        set_local_dial_number(self.ns, "10086")
        worker = send_local_serial_command_namespace_runtime(self.ns, "ATD10086;")
        worker.join(2)
        self.replace_call()
        self.ns["current_dial_num"] = ""
        self.ns["serial_obj"].writes.clear()
        worker = send_local_serial_command_namespace_runtime(self.ns, "ATH")
        worker.join(2)
        self.assertFalse(worker.is_alive())
        self.assertEqual(self.ns["serial_obj"].writes, [b"ATH\r\n"])


class LocalDialButtonIsolationTests(unittest.TestCase):
    def setUp(self):
        self.ns = namespace_fixture()
        self.ns["ring_timeout_target"] = 0.0
        self.ns["current_dial_popup"] = None
        self.actions, self.popups, self.workers = {}, [], []

    def open_window(self, stack):
        def popup_factory(*args):
            popup = SimpleNamespace(exists=True, hangup=args[3], close=args[4])
            popup.winfo_exists = lambda: popup.exists
            popup.destroy = lambda: setattr(popup, "exists", False)
            self.popups.append(popup)
            return popup

        overrides = {
            "tk.Toplevel": lambda *_: FakeWindow(), "tk.BooleanVar": FakeVar, "tk.StringVar": FakeVar,
            "ttk.Frame": FakeWidget, "ttk.Checkbutton": FakeWidget, "ttk.Label": FakeWidget,
            "ttk.Entry": FakeWidget, "ttk.Button": FakeWidget,
            "create_serial_debug_body": lambda *_: (FakeWidget(), FakeWidget(), FakeWidget()),
            "create_serial_debug_quick_actions": lambda *args: self.actions.update(dial=args[8], hangup=args[9]),
            "SerialDebugPauseController": FakePauseController, "SerialDebugFinder": FakeFinder,
            "start_serial_debug_append_loop": lambda *_: None, "open_dial_call_popup": popup_factory,
        }
        for name, value in overrides.items():
            stack.enter_context(patch("sms_ui.serial_debug_window." + name, value))
        open_serial_debug_window_namespace_runtime(self.ns)

    def join_workers(self):
        for worker in self.workers + list(self.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot()):
            worker.join(3)
            self.assertFalse(worker.is_alive())

    def test_old_popup_hangup_and_delayed_close_do_not_touch_same_number_redial(self):
        with ExitStack() as stack:
            self.open_window(stack)
            self.actions["dial"]("10086")
            self.join_workers()
            old = self.popups[-1]
            self.actions["dial"]("10086")
            self.join_workers()
            new = self.popups[-1]
            self.ns["serial_obj"].writes.clear()
            generation = self.ns["_dial_state_generation"]
            old.hangup()
            old.close()
            self.join_workers()
            self.assertEqual(self.ns["serial_obj"].writes, [])
            self.assertIs(self.ns["_dial_state_generation"], generation)
            self.assertEqual(self.ns["current_dial_num"], "10086")
            self.assertIs(self.ns["current_dial_popup"], new)
            self.assertTrue(new.exists)

    def test_popup_and_shortcut_hangup_immediately_cancel_unsubmitted_dial(self):
        for popup_button in (False, True):
            with self.subTest(popup_button=popup_button):
                self.setUp()
                with ExitStack() as stack:
                    self.open_window(stack)
                    with serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK:
                        self.actions["dial"]("10086")
                        self.workers.extend(self.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot())
                        request = self.ns["_local_dial_request"]
                        (self.popups[-1].hangup if popup_button else self.actions["hangup"])()
                        self.workers.extend(self.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot())
                    self.join_workers()
                    self.assertTrue(request.finished)
                    self.assertEqual(self.ns["current_dial_num"], "")
                    self.assertEqual(self.ns["serial_obj"].writes, [])
                    self.assertTrue(self.ns["CLOUD_SMS_EVENT_Q"].empty())

    def test_submitted_popup_hangup_remains_bound_during_queue_wait(self):
        with ExitStack() as stack:
            self.open_window(stack)
            self.actions["dial"]("10086")
            self.join_workers()
            popup = self.popups[-1]
            modem = self.ns["serial_obj"]
            modem.writes.clear()
            gate = Gate()
            stack.enter_context(patch.object(serial_sender, "DEFAULT_SERIAL_TRANSACTION_LOCK", gate))
            popup.hangup()
            self.workers.extend(self.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot())
            try:
                self.assertTrue(gate.reached.wait(2))
                self.assertEqual(self.ns["current_dial_num"], "10086")
                with call_state_lock(self.ns):
                    observe_call_command_line(self.ns, "NO CARRIER")
                    self.ns["INCOMING_CALL_SESSION"] = IncomingCallSessionTracker()
                    self.ns["INCOMING_CALL_SESSION"].start("synthetic-new")
            finally:
                gate.resume.set()
                self.join_workers()
            self.assertEqual(modem.writes, [])

    def test_current_submitted_hangup_writes_once_and_closes_original_popup(self):
        with ExitStack() as stack:
            self.open_window(stack)
            self.actions["dial"]("10086")
            self.join_workers()
            popup = self.popups[-1]
            self.ns["serial_obj"].writes.clear()
            popup.hangup()
            self.join_workers()
            self.assertEqual(self.ns["serial_obj"].writes, [b"ATH\r\n"])
            self.assertEqual(self.ns["current_dial_num"], "")
            self.assertIsNone(self.ns["current_dial_popup"])
            self.assertFalse(popup.exists)

    def test_existing_popup_does_not_hang_up_a_reconnected_serial(self):
        for same_serial in (False, True):
            with self.subTest(same_serial=same_serial):
                self.setUp()
                with ExitStack() as stack:
                    self.open_window(stack)
                    self.actions["dial"]("10086")
                    self.join_workers()
                    old_modem = self.ns["serial_obj"]
                    old_modem.writes.clear()
                    self.ns["serial_connection_generation"] += 1
                    if not same_serial:
                        self.ns["serial_obj"] = namespace_fixture()["serial_obj"]
                    self.popups[-1].hangup()
                    self.join_workers()
                    self.assertEqual(old_modem.writes, [])
                    self.assertEqual(self.ns["serial_obj"].writes, [])

    def test_delayed_hangup_ui_result_does_not_clear_or_restyle_new_call(self):
        with ExitStack() as stack:
            self.open_window(stack)
            self.actions["dial"]("10086")
            self.join_workers()
            popup = self.popups[-1]
            ui, statuses = [], []
            self.ns["run_on_ui_thread"] = lambda callback, _post: ui.append(callback)
            self.ns["set_status"] = lambda *args: statuses.append(args)
            popup.hangup()
            self.join_workers()
            self.assertEqual(len(ui), 1)
            self.actions["dial"]("10086")
            self.join_workers()
            new = self.popups[-1]
            statuses.clear()
            ui.pop(0)()
            self.assertIs(self.ns["current_dial_popup"], new)
            self.assertEqual(self.ns["current_dial_num"], "10086")
            self.assertTrue(new.exists)
            self.assertEqual(statuses, [])


if __name__ == "__main__":
    unittest.main()
