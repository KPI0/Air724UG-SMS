"""Connected status must belong to a ready serial and the latest UI request."""
from contextlib import ExitStack
from dataclasses import replace
from functools import partial
import queue
import threading
import time
from types import SimpleNamespace
import unittest
from unittest.mock import patch

from sms_core.call_command_context import SerialCallStateSnapshot, set_local_dial_number
from sms_core.call_session import IncomingCallSessionTracker
from sms_core.connected_log_runtime import can_restore_connected_status
from sms_core.serial_app_runtime import (
    run_serial_app_from_wiring, run_serial_app_runtime, run_serial_reader_namespace_runtime,
)
from sms_core.serial_runtime import run_serial_runtime_thread
from sms_ui.app_infrastructure_namespace_runtime import safe_close_serial_namespace_runtime
from sms_ui.app_ui_namespace_bindings import install_app_ui_namespace_bindings
from sms_ui.app_ui_namespace_runtime import (
    restore_connected_status_namespace_runtime, set_status_namespace_runtime,
)
from sms_ui import call_popup_namespace_runtime as popup_runtime
from sms_ui.thread_runtime import run_on_ui_thread, ui_post_runtime
from tests import test_local_call_command_isolation as local_fixture
from tests import test_peer_call_sync as peer_fixture
from tests import test_peer_terminal_ui_queue as peer_queue_fixture
from tests import test_serial_initialization as initialization_fixture


class StatusValue:
    def __init__(self):
        self.text = "initial"
        self.color = "black"

    def set(self, value):
        self.text = value

    def get(self):
        return self.text

    def config(self, **kwargs):
        self.color = kwargs["fg"]


def install_status(namespace, tasks):
    value = StatusValue()
    namespace.update(
        status_var=value, status_label=value, tk_alive=lambda: True,
        run_on_ui_thread=run_on_ui_thread,
        ui_post=lambda callback: ui_post_runtime(tasks, callback),
    )
    namespace["set_status"] = partial(set_status_namespace_runtime, namespace)
    namespace["restore_connected_status"] = partial(restore_connected_status_namespace_runtime, namespace)
    return value


def drain(tasks):
    while not tasks.empty():
        callback, args, kwargs = tasks.get_nowait()
        callback(*args, **kwargs)
        tasks.task_done()


class ConnectedStatusGenerationTests(unittest.TestCase):
    def peer_fixture(self):
        fixture = peer_fixture.PeerCallSyncTests("runTest")
        fixture.setUp()
        namespace = fixture.namespace
        replay = peer_queue_fixture.PeerTerminalReplay(
            namespace, show_popup=lambda _number, call_session_id="": fixture.show_popup(call_session_id),
        )
        serial = SimpleNamespace(is_open=True)
        serial.close = lambda: setattr(serial, "is_open", False)
        namespace.update(serial_obj=serial, serial_lock=threading.RLock(), serial_running=True,
                         serial_connection_generation=1, unlock_port_mutex=lambda: None)
        value = install_status(namespace, replay.tasks)
        replay.callbacks = replace(
            replay.callbacks, set_status=namespace["set_status"],
            restore_connected_status=namespace["restore_connected_status"],
        )
        return fixture, replay, namespace, value, serial

    def change_connection(self, namespace, serial, change):
        if change == "closed":
            safe_close_serial_namespace_runtime(namespace)
        elif change == "reconnected":
            safe_close_serial_namespace_runtime(namespace)
            namespace["serial_obj"] = SimpleNamespace(is_open=True)
            namespace["serial_connection_generation"] += 1
        elif change == "same_object_generation":
            namespace["serial_connection_generation"] += 1
        elif change == "pending":
            serial._serial_initialization_session = SimpleNamespace(pending=True, quarantined=False)
        elif change == "quarantined":
            serial._sms_recovery_error = "synthetic unresolved transaction"
        elif change == "initialization_error":
            namespace["_serial_initialization_session"] = SimpleNamespace(serial=serial, has_command_errors=True)
        elif change == "stopped":
            namespace["serial_running"] = False
        elif change == "exiting":
            namespace["is_exiting"] = True
        elif change in ("serial_stop_event", "TK_SHUTDOWN"):
            event = threading.Event()
            event.set()
            namespace[change] = event

    def test_older_main_status_cannot_replace_newer_status_or_repeat(self):
        _fixture, replay, namespace, value, _serial = self.peer_fixture()
        replay.background(lambda: namespace["set_status"]("old connected status", "green"))
        callback, args, kwargs = replay.tasks.get_nowait()
        replay.tasks.task_done()
        namespace["set_status"]("new serial state", "orange")
        callback(*args, **kwargs)
        callback(*args, **kwargs)
        self.assertEqual((value.text, value.color), ("new serial state", "orange"))

    def test_invalid_call_snapshot_cannot_displace_a_valid_queued_status(self):
        _fixture, replay, namespace, value, _serial = self.peer_fixture()
        old = SerialCallStateSnapshot(namespace)
        set_local_dial_number(namespace, "10086")
        current = SerialCallStateSnapshot(namespace)
        replay.background(lambda: namespace["set_status"]("current call", "blue", call_state_snapshot=current))
        token = namespace["_status_ui_token"]
        self.assertFalse(namespace["set_status"]("stale call", "green", call_state_snapshot=old))
        self.assertIs(namespace["_status_ui_token"], token)
        replay.drain()
        self.assertEqual((value.text, value.color), ("current call", "blue"))

        _fixture, replay, namespace, value, _serial = self.peer_fixture()
        validated, resume, errors = threading.Event(), threading.Event(), []
        threads = {}

        class ObservedRLock:
            def __init__(self):
                self.lock = threading.RLock()

            def __enter__(self):
                if not self.lock.acquire(blocking=False):
                    if threading.current_thread() is threads.get("new"):
                        resume.set()
                    self.lock.acquire()
                return self

            def __exit__(self, *_args):
                self.lock.release()

        namespace["_call_peer_state_lock"] = ObservedRLock()
        old = SerialCallStateSnapshot(namespace)
        original_current = old.is_current

        def paused_validation():
            result = original_current()
            validated.set()
            if not resume.wait(2):
                raise AssertionError("new call did not reach the validation boundary")
            return result

        def old_status():
            try:
                namespace["set_status"]("old call", "blue", call_state_snapshot=old)
            except Exception as exc:
                errors.append(exc)

        def new_status():
            try:
                set_local_dial_number(namespace, "10086")
                current = SerialCallStateSnapshot(namespace)
                namespace["set_status"]("current call", "blue", call_state_snapshot=current)
            except Exception as exc:
                errors.append(exc)
            finally:
                resume.set()

        threads["old"] = threading.Thread(target=old_status)
        threads["new"] = threading.Thread(target=new_status)
        with patch.object(old, "is_current", side_effect=paused_validation):
            threads["old"].start()
            try:
                self.assertTrue(validated.wait(2))
                threads["new"].start()
                threads["new"].join(3)
            finally:
                resume.set()
                threads["old"].join(3)
            self.assertFalse(threads["old"].is_alive())
            self.assertFalse(threads["new"].is_alive())
            replay.drain()
        self.assertEqual(errors, [])
        self.assertEqual((value.text, value.color), ("current call", "blue"))

    def test_restore_preserves_a_more_recent_status_on_the_same_serial(self):
        _fixture, replay, namespace, value, _serial = self.peer_fixture()
        replay.background(lambda: namespace["restore_connected_status"]("TEST"))
        namespace["set_status"]("new serial state", "orange")
        replay.drain()
        self.assertEqual((value.text, value.color), ("new serial state", "orange"))

    def test_restore_requires_original_ready_connection_and_no_new_call(self):
        changes = ("closed", "reconnected", "same_object_generation", "pending", "quarantined",
                   "initialization_error", "stopped", "exiting", "serial_stop_event", "TK_SHUTDOWN",
                   "new_incoming", "new_outgoing", "new_ring")
        for change in changes:
            with self.subTest(change=change):
                fixture, replay, namespace, value, serial = self.peer_fixture()
                replay.background(lambda: namespace["restore_connected_status"]("OLD_PORT"))
                if change == "new_incoming":
                    fixture.tracker.start("10086")
                elif change == "new_outgoing":
                    set_local_dial_number(namespace, "10086")
                elif change == "new_ring":
                    namespace["_call_command_pending_incoming"] = True
                else:
                    self.change_connection(namespace, serial, change)
                self.assertFalse(can_restore_connected_status(namespace, serial, 1))
                replay.drain()
                self.assertEqual((value.text, value.color), ("initial", "black"))

    def test_current_ready_restore_and_namespace_binding_work(self):
        _fixture, replay, namespace, value, _serial = self.peer_fixture()
        install_app_ui_namespace_bindings(namespace)
        replay.background(lambda: namespace["restore_connected_status"]("TEST"))
        replay.drain()
        self.assertEqual((value.text, value.color), ("🟢 已连接：TEST", "green"))

    def test_peer_terminal_closes_original_window_without_restoring_changed_serial(self):
        changes = ("closed", "reconnected", "same_object_generation", "pending", "quarantined",
                   "initialization_error", "stopped", "exiting", "serial_stop_event", "TK_SHUTDOWN")
        for change in changes:
            with self.subTest(change=change):
                _fixture, replay, namespace, value, serial = self.peer_fixture()
                old_id = replay.start_old()
                replay.drain()
                popup = namespace["current_call_popup"]
                replay.cloud("incoming", old_id)
                replay.cloud("ended")
                self.change_connection(namespace, serial, change)
                before = (value.text, value.color)
                replay.drain()
                self.assertEqual((value.text, value.color), before)
                self.assertFalse(popup.exists)
                self.assertIsNone(namespace["current_call_popup"])
                self.assertEqual(len(replay.missed), 1)

    def test_peer_terminal_restores_current_serial_and_preserves_newer_status(self):
        for newer_status in (False, True):
            with self.subTest(newer_status=newer_status):
                _fixture, replay, namespace, value, _serial = self.peer_fixture()
                old_id = replay.start_old()
                replay.drain()
                replay.cloud("incoming", old_id)
                replay.cloud("ended")
                if newer_status:
                    namespace["set_status"]("new serial state", "orange")
                replay.drain()
                expected = ("new serial state", "orange") if newer_status else ("🟢 已连接：TEST", "green")
                self.assertEqual((value.text, value.color), expected)
                self.assertEqual(len(replay.missed), 1)

    def test_serial_terminal_and_timeout_restore_only_after_original_call_finishes(self):
        for terminal in ("NO CARRIER", "timeout"):
            for close_first in (False, True):
                with self.subTest(terminal=terminal, close_first=close_first):
                    _fixture, replay, namespace, value, _serial = self.peer_fixture()
                    replay.start_old()
                    replay.drain()
                    replay.serial("[I]-[ril.proatc] NO CARRIER" if terminal != "timeout" else "", 114.0)
                    if close_first:
                        safe_close_serial_namespace_runtime(namespace)
                        namespace["set_status"]("new serial state", "orange")
                    replay.drain()
                    expected = ("new serial state", "orange") if close_first else ("🟢 已连接：TEST", "green")
                    self.assertEqual((value.text, value.color), expected)
                    self.assertIsNone(namespace["current_call_popup"])
                    self.assertEqual(len(replay.missed), 1)

    def test_local_confirmed_hangup_keeps_window_cleanup_separate_from_serial_status(self):
        for change in ("current", "closed", "reconnected", "pending", "quarantined", "stopped", "exiting", "new_status"):
            with self.subTest(change=change):
                fixture = local_fixture.LocalDialButtonIsolationTests("runTest")
                fixture.setUp()
                namespace = fixture.ns
                tasks = queue.Queue()
                value = install_status(namespace, tasks)
                namespace["unlock_port_mutex"] = lambda: None
                with ExitStack() as stack:
                    fixture.open_window(stack)
                    fixture.actions["dial"]("10086")
                    fixture.join_workers()
                    drain(tasks)
                    popup, serial = fixture.popups[-1], namespace["serial_obj"]
                    serial.close = lambda: setattr(serial, "is_open", False)
                    serial.writes.clear()
                    popup.hangup()
                    fixture.join_workers()
                    self.assertEqual(serial.writes, [b"ATH\r\n"])
                    self.assertEqual(tasks.qsize(), 1)
                    if change != "current":
                        self.change_connection(namespace, serial, change)
                        namespace["set_status"]("new serial state", "orange")
                    drain(tasks)
                    expected = ("new serial state", "orange") if change != "current" else ("🟢 已连接：TEST", "green")
                    self.assertEqual((value.text, value.color), expected)
                    self.assertFalse(popup.exists)
                    self.assertIsNone(namespace["current_dial_popup"])

    def test_real_reader_uses_restore_callback_for_terminal_and_timeout(self):
        for terminal in ("NO CARRIER", "timeout"):
            with self.subTest(terminal=terminal):
                harness = initialization_fixture.InitializationHarness(hold="NONE")
                namespace = harness.ns
                tasks = queue.Queue()
                value = install_status(namespace, tasks)
                tracker = IncomingCallSessionTracker()
                namespace.update(
                    INCOMING_CALL_SESSION=tracker,
                    start_serial_incoming_call=partial(popup_runtime.start_serial_incoming_call_namespace_runtime, namespace),
                    finish_incoming_call_session=partial(popup_runtime.finish_incoming_call_session_namespace_runtime, namespace),
                    reset_incoming_call_session=partial(popup_runtime.reset_incoming_call_session_namespace_runtime, namespace),
                )
                clocked_reader = partial(
                    run_serial_reader_namespace_runtime,
                    run_reader=partial(run_serial_app_from_wiring, run_app=partial(
                        run_serial_app_runtime, run_runtime_thread=partial(
                            run_serial_runtime_thread, clock=lambda: harness.clock))),
                )
                with patch.object(initialization_fixture, "run_serial_reader_namespace_runtime", clocked_reader), harness:
                    self.wait_for(lambda: not harness.opened_session.pending)
                    drain(tasks)
                    harness.emit('[I]-[ril.proatc] +CLIP: "10086",129\r\n')
                    self.wait_for(lambda: tracker.snapshot().session_token is not None)
                    drain(tasks)
                    if terminal == "timeout":
                        harness.clock = 114.0
                    else:
                        harness.emit("[I]-[ril.proatc] NO CARRIER\r\n")
                    self.wait_for(lambda: tracker.snapshot().session_token is None and namespace["ring_timeout_target"] == 0)
                    self.wait_for(lambda: not tasks.empty())
                    drain(tasks)
                    self.assertEqual((value.text, value.color), ("🟢 已连接：TEST", "green"))
                    self.assertEqual(harness.errors, [])

    def wait_for(self, predicate):
        deadline = time.monotonic() + 2
        while time.monotonic() < deadline:
            if predicate():
                return
            threading.Event().wait(0.005)
        self.fail("virtual status condition was not reached")


if __name__ == "__main__":
    unittest.main()
