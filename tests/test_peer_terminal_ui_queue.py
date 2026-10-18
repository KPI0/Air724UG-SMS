"""A received peer terminal must release serial call state before Tk can run."""

import queue
import threading
import unittest
from dataclasses import replace

from sms_core.cloud_message_runtime import handle_device_call_state_runtime
from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.serial_runtime import SerialRuntimeState, handle_serial_runtime_line
from sms_ui import call_popup_namespace_runtime as popup_runtime
from sms_ui.thread_runtime import run_on_ui_thread, ui_post_runtime
from tests import test_call_popup_visibility as visibility_fixture
from tests import test_peer_call_sync as peer_fixture
from tests.test_serial_runtime import runtime_callbacks, runtime_config


class PeerTerminalReplay:
    def __init__(self, namespace, show_popup=None):
        self.namespace = namespace
        self.tasks = queue.Queue()
        self.errors = []
        self.calls = []
        self.missed = []
        self.state = SerialRuntimeState.create(parse_sms_callback_head)
        namespace.update(
            run_on_ui_thread=run_on_ui_thread,
            ui_post=lambda callback: ui_post_runtime(self.tasks, callback),
            PORT="TEST", ring_timeout_target=0.0, current_dial_num="",
        )
        show_missed = namespace.get("show_missed_call_popup")

        def remember_missed(missed_call):
            self.missed.append(missed_call)
            if callable(show_missed):
                show_missed(missed_call)

        namespace["show_missed_call_popup"] = remember_missed
        self.callbacks = replace(
            runtime_callbacks(self.calls),
            start_incoming_call=lambda number: popup_runtime.start_incoming_call_session_namespace_runtime(namespace, number),
            finish_incoming_call=lambda: popup_runtime.finish_incoming_call_session_namespace_runtime(namespace),
            close_call_popup=lambda: popup_runtime.close_call_popup_namespace_runtime(namespace),
            show_call_popup=show_popup or (lambda number, call_session_id="":
                popup_runtime.show_call_popup_namespace_runtime(namespace, number, call_session_id)),
            show_missed_call_popup=remember_missed,
            consume_remote_incoming_end=lambda session: popup_runtime.consume_remote_incoming_end_namespace_runtime(namespace, session),
        )

    def background(self, callback):
        def run():
            try:
                callback()
            except BaseException as exc:
                self.errors.append(exc)
        worker = threading.Thread(target=run)
        worker.start()
        worker.join(3)
        if worker.is_alive():
            raise AssertionError("Replay worker did not finish")
        if self.errors:
            raise self.errors.pop(0)

    def serial(self, line, now):
        def dispatch():
            ns = self.namespace
            self.state.sync_call_state(ns["ring_timeout_target"], ns["current_dial_num"])
            handle_serial_runtime_line(
                self.state, line, now, "TEST", ns.get("current_call_popup") is not None,
                runtime_config(), self.callbacks, {},
            )
            ns["ring_timeout_target"] = self.state.call_state.ring_timeout_target
            ns["current_dial_num"] = self.state.call_state.current_dial_num
        self.background(dispatch)

    def cloud(self, phase, local_id="", peer_id="peer-old"):
        payload = {
            "type": "device_call_state", "imei": "000000000000001",
            "direction": "incoming", "phase": phase, "phone": "10086",
            "call_session_id": peer_id,
        }
        if phase == "incoming":
            payload.update(local_call_session_id=local_id,
                           call_session_ids={"desktop": local_id, "firmware": peer_id})
        ns = self.namespace
        self.background(lambda: handle_device_call_state_runtime(
            payload, runtime_imei=lambda: "000000000000001",
            register_incoming=lambda *args: popup_runtime.register_peer_incoming_call_namespace_runtime(ns, *args),
            mark_peer_call_connected=lambda *args: popup_runtime.mark_peer_incoming_call_connected_namespace_runtime(ns, *args),
            finish_call=lambda *args, **kwargs: popup_runtime.finish_remote_incoming_call_namespace_runtime(ns, *args, **kwargs),
        ))

    def drain(self):
        while not self.tasks.empty():
            callback, args, kwargs = self.tasks.get_nowait()
            callback(*args, **kwargs)
            self.tasks.task_done()

    def start_old(self):
        self.serial("RING", 100.0)
        self.serial('+CLIP: "10086",129', 100.1)
        return self.state.call_state.call_session_id

    def redial(self):
        self.serial("RING", 100.5)
        self.serial('+CLIP: "10086",129', 101.0)
        return self.state.call_state.call_session_id


class PeerTerminalUiQueueTests(unittest.TestCase):
    def setUp(self):
        self.fixture = peer_fixture.PeerCallSyncTests()
        self.fixture.setUp()
        self.replay = PeerTerminalReplay(
            self.fixture.namespace,
            show_popup=lambda _number, call_session_id="": self.fixture.show_popup(call_session_id),
        )
        self.namespace = self.fixture.namespace

    def test_received_terminal_releases_call_before_ui_queue_drains(self):
        old_id = self.replay.start_old()
        self.replay.cloud("incoming", old_id)
        self.replay.cloud("ended")
        self.assertIsNone(self.fixture.tracker.snapshot().session_token)
        self.assertIn(old_id, self.namespace["call_popup_pending_serial_ends"])
        new_id = self.replay.redial()
        self.assertNotEqual(new_id, old_id)
        new_popup = self.namespace["current_call_popup"]
        self.replay.drain()
        self.assertIs(self.namespace["current_call_popup"], new_popup)
        self.assertTrue(new_popup.exists)
        self.assertEqual(self.namespace["call_popup_active_session_id"], new_id)
        self.assertGreater(self.namespace["ring_timeout_target"], 101)
        self.assertEqual(len(self.replay.missed), 1)
        self.assertEqual(self.fixture.statuses, [])

    def test_repeated_queued_terminal_does_not_finish_redial_or_repeat_notice(self):
        old_id = self.replay.start_old()
        self.replay.cloud("incoming", old_id)
        self.replay.cloud("ended")
        self.replay.cloud("ended")
        self.assertEqual(self.replay.tasks.qsize(), 1)
        new_id = self.replay.redial()
        self.replay.cloud("ended")
        self.replay.drain()
        self.assertNotEqual(new_id, old_id)
        self.assertEqual(self.fixture.tracker.snapshot().caller_num, "10086")
        self.assertEqual(len(self.replay.missed), 1)

    def test_hidden_call_and_terminal_before_binding_still_release_serial(self):
        old_id = self.replay.start_old()
        popup_runtime.close_call_popup_namespace_runtime(self.namespace)
        self.replay.cloud("ended")
        self.replay.cloud("incoming", old_id)
        new_id = self.replay.redial()
        self.replay.drain()
        self.assertNotEqual(new_id, old_id)
        self.assertEqual(self.namespace["call_popup_active_session_id"], new_id)
        self.assertTrue(self.namespace["current_call_popup"].exists)
        self.assertEqual(len(self.replay.missed), 1)

    def test_delayed_cleanup_preserves_new_outgoing_state(self):
        old_id = self.replay.start_old()
        self.replay.cloud("incoming", old_id)
        self.replay.cloud("ended")
        self.namespace["current_dial_num"] = "10010"
        self.namespace["ring_timeout_target"] = -1.0
        self.replay.drain()
        self.assertEqual(self.namespace["current_dial_num"], "10010")
        self.assertEqual(self.namespace["ring_timeout_target"], -1.0)
        self.assertEqual(self.fixture.statuses, [])
        self.assertEqual(len(self.replay.missed), 1)

    def test_accepted_connected_before_queued_terminal_is_not_a_missed_call(self):
        old_id = self.replay.start_old()
        self.replay.cloud("incoming", old_id)
        self.replay.cloud("connected")
        self.replay.cloud("ended")
        new_id = self.replay.redial()
        self.replay.drain()
        self.assertNotEqual(new_id, old_id)
        self.assertFalse(self.fixture.tracker.snapshot().handled)
        self.assertGreater(self.namespace["ring_timeout_target"], 101)
        self.assertEqual(self.replay.missed, [])

    def test_repeated_unnumbered_terminal_without_live_call_is_not_queued_twice(self):
        self.replay.start_old()
        self.replay.cloud("ended", peer_id="")
        self.replay.cloud("ended", peer_id="")
        self.assertEqual(self.replay.tasks.qsize(), 1)
        self.replay.drain()
        self.assertEqual(len(self.replay.missed), 1)


class PeerTerminalTkQueueTests(unittest.TestCase):
    def test_real_tk_redial_survives_old_cleanup_before_or_after_new_show(self):
        for create_initial in (False, True):
            for new_show_first in (False, True):
                with self.subTest(create_initial=create_initial, new_show_first=new_show_first):
                    fixture = visibility_fixture.CallPopupVisibilityTests()
                    fixture.setUp()
                    try:
                        ns = fixture.namespace
                        ns["show_missed_call_popup"] = lambda missed: popup_runtime.show_missed_call_popup_namespace_runtime(ns, missed)
                        replay = PeerTerminalReplay(ns)
                        old_id = replay.start_old()
                        if create_initial:
                            replay.drain()
                            fixture.root.update()
                        replay.cloud("incoming", old_id)
                        replay.cloud("ended")
                        new_id = replay.redial()
                        self.assertNotEqual(new_id, old_id)
                        if new_show_first:
                            pending = []
                            while not replay.tasks.empty():
                                pending.append(replay.tasks.get_nowait())
                                replay.tasks.task_done()
                            callback, args, kwargs = pending.pop()
                            callback(*args, **kwargs)
                            for callback, args, kwargs in pending:
                                callback(*args, **kwargs)
                        else:
                            replay.drain()
                        fixture.root.update()
                        incoming = ns["current_call_popup"]
                        missed = ns["current_missed_call_popup"]
                        self.assertEqual(incoming._call_popup_session_id, new_id)
                        self.assertTrue(incoming.winfo_exists())
                        self.assertEqual(len(replay.missed), 1)
                        fixture.assert_incoming_above_missed(incoming, missed)
                    finally:
                        fixture.tearDown()
