import threading
import unittest
from dataclasses import replace
from datetime import datetime
from types import SimpleNamespace
from unittest.mock import patch

from sms_core.call_session import IncomingCallSessionTracker
from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.cloud_message_runtime import handle_device_call_state_runtime
from sms_core.serial_runtime import SerialRuntimeState, handle_serial_runtime_line
from sms_ui import call_popup_namespace_runtime as popup_runtime
from sms_ui.call_popup_namespace_runtime import (
    PEER_CALL_BINDING_TTL_SECONDS,
    _get_peer_call_binding,
    close_call_popup_namespace_runtime,
    finish_incoming_call_session_namespace_runtime,
    finish_remote_incoming_call_namespace_runtime,
    mark_peer_incoming_call_connected_namespace_runtime,
    register_peer_incoming_call_namespace_runtime,
    show_call_popup_namespace_runtime,
)
from tests.test_serial_runtime import runtime_callbacks, runtime_config


class PeerCallSyncTests(unittest.TestCase):
    def setUp(self):
        self.now = [100.0]
        self.tracker = IncomingCallSessionTracker(now_func=lambda: datetime(2026, 1, 1))
        self.missed = []
        self.statuses = []
        self.namespace = {
            "INCOMING_CALL_SESSION": self.tracker,
            "current_call_popup": None, "current_dial_popup": None,
            "root": None, "serial_lock": None, "serial_obj": None,
            "center_window": lambda *_: None, "port_ui": lambda *_: None,
            "set_status": lambda *args: self.statuses.append(args),
            "PORT": "COM5", "ring_timeout_target": 0.0,
            "run_on_ui_thread": lambda callback, _post: callback(),
            "ui_post": lambda callback: callback(),
            "show_missed_call_popup": self.missed.append,
            "mark_incoming_call_handled": self.tracker.mark_handled,
            "close_call_popup": lambda: close_call_popup_namespace_runtime(self.namespace),
            "_call_peer_monotonic": lambda: self.now[0],
        }

    def start(self, local_id):
        self.assertTrue(self.tracker.start("10086"))
        self.namespace["ring_timeout_target"] = 123.0
        return self.show_popup(local_id)

    def show_popup(self, local_id):
        def show(**kwargs):
            popup = SimpleNamespace(
                _call_popup_caller_num="10086", exists=True,
                _call_popup_mark_connected=self.tracker.mark_handled,
                close=kwargs["close_popup"],
            )
            popup.winfo_exists = lambda: popup.exists
            popup.destroy = lambda: setattr(popup, "exists", False)
            kwargs["set_popup"](popup)
            return popup

        return show_call_popup_namespace_runtime(
            self.namespace, "10086", local_id, show_app_runtime=show)

    def register(self, peer_id, local_id):
        return register_peer_incoming_call_namespace_runtime(
            self.namespace, peer_id, "10086", local_id)

    def connected(self, peer_id):
        return mark_peer_incoming_call_connected_namespace_runtime(
            self.namespace, peer_id, "10086")

    def terminal(self, peer_id):
        return finish_remote_incoming_call_namespace_runtime(
            self.namespace, "ended", "NO CARRIER", session_id=peer_id, caller_num="10086")

    def finish_local(self):
        finish_incoming_call_session_namespace_runtime(self.namespace)
        close_call_popup_namespace_runtime(self.namespace)

    def assert_finished(self, popup):
        self.assertIsNone(self.namespace["current_call_popup"])
        self.assertFalse(popup.exists)
        self.assertEqual(self.tracker.snapshot().caller_num, "")
        self.assertIsNone(self.tracker.snapshot().session_token)

    def test_late_old_incoming_cannot_replace_pending_new_correlation(self):
        self.assertTrue(self.register("peer-new", "local-new"))
        self.assertTrue(self.register("peer-old", "local-old"))

        popup = self.start("local-new")

        binding = _get_peer_call_binding(self.namespace)
        self.assertIsNotNone(binding)
        self.assertEqual(binding["peer_session_id"], "peer-new")
        self.assertEqual(binding["local_session_id"], "local-new")
        self.assertTrue(self.terminal("peer-new"))
        self.assert_finished(popup)

    def test_binding_commit_cannot_overwrite_redial_from_another_thread(self):
        binding_ready = threading.Event()
        resume_binding = threading.Event()
        failures = []
        redial = {}

        class ObservedRLock:
            def __init__(self):
                self.lock = threading.RLock()

            def __enter__(self):
                if not self.lock.acquire(blocking=False):
                    if threading.current_thread() is redial.get("thread"):
                        # Let the in-flight binding finish when the real lock
                        # serializes the redial, without a timing-based sleep.
                        resume_binding.set()
                    self.lock.acquire()
                return self

            def __exit__(self, *_args):
                self.lock.release()

        self.namespace["_call_peer_state_lock"] = ObservedRLock()
        self.start("local-old")
        self.assertTrue(self.register("peer-new", "local-new"))
        original_set = popup_runtime._set_peer_call_binding

        def pause_old_binding(namespace, peer_id, *args):
            if peer_id == "peer-old":
                binding_ready.set()
                if not resume_binding.wait(3):
                    raise AssertionError("redial did not release the binding")
            return original_set(namespace, peer_id, *args)

        def bind_old_call():
            try:
                self.assertTrue(self.register("peer-old", "local-old"))
            except Exception as exc:
                failures.append(exc)

        def start_redial():
            try:
                self.finish_local()
                popup_runtime.start_incoming_call_session_namespace_runtime(
                    self.namespace, "10086")
                redial["popup"] = self.show_popup("local-new")
            except Exception as exc:
                failures.append(exc)
            finally:
                resume_binding.set()

        binding_thread = threading.Thread(target=bind_old_call, daemon=True)
        redial["thread"] = threading.Thread(target=start_redial, daemon=True)
        with patch.object(popup_runtime, "_set_peer_call_binding", side_effect=pause_old_binding):
            binding_thread.start()
            try:
                self.assertTrue(binding_ready.wait(3))
                redial["thread"].start()
                redial["thread"].join(3)
            finally:
                resume_binding.set()
                binding_thread.join(3)
            self.assertFalse(binding_thread.is_alive())
            self.assertFalse(redial["thread"].is_alive())
        self.assertEqual(failures, [])
        self.assertEqual(_get_peer_call_binding(self.namespace), {
            "peer_session_id": "peer-new", "local_session_id": "local-new",
            "caller_num": "10086",
        })
        self.assertTrue(self.terminal("peer-new"))
        self.assert_finished(redial["popup"])

    def test_late_old_terminal_cannot_replace_pending_new_terminal(self):
        self.assertTrue(self.terminal("peer-new"))
        self.assertTrue(self.register("peer-new", "local-new"))
        self.assertTrue(self.terminal("peer-old"))

        popup = self.start("local-new")

        self.assert_finished(popup)
        self.assertEqual(len(self.missed), 1)

    def test_late_old_connected_cannot_replace_pending_new_connected(self):
        self.assertTrue(self.connected("peer-new"))
        self.assertTrue(self.register("peer-new", "local-new"))
        self.assertTrue(self.connected("peer-old"))

        popup = self.start("local-new")

        self.assertIs(self.namespace["current_call_popup"], popup)
        self.assertTrue(popup.exists)
        self.assertTrue(self.tracker.snapshot().handled)
        self.assertEqual(_get_peer_call_binding(self.namespace)["peer_session_id"], "peer-new")

    def test_next_terminal_waits_while_previous_call_is_bound(self):
        old_popup = self.start("local-old")
        self.assertTrue(self.register("peer-old", "local-old"))
        self.assertTrue(self.register("peer-new", "local-new"))
        self.terminal("peer-new")
        self.assertIs(self.namespace["current_call_popup"], old_popup)
        self.assertTrue(old_popup.exists)
        self.assertEqual(self.tracker.snapshot().caller_num, "10086")

        self.finish_local()
        popup = self.start("local-new")

        self.assert_finished(popup)
        self.assertEqual(len(self.missed), 1)

    def test_current_correlation_arriving_after_hide_can_finish_call(self):
        popup = self.start("local-1")
        token = self.tracker.snapshot().session_token
        popup.close()
        self.assertIsNone(self.namespace["current_call_popup"])
        self.assertIs(self.tracker.snapshot().session_token, token)

        self.assertTrue(self.register("peer-1", "local-1"))
        self.assertTrue(self.terminal("peer-1"))

        self.assert_finished(popup)
        self.assertEqual(len(self.missed), 1)

    def test_connected_hidden_call_is_handled_before_its_terminal(self):
        popup = self.start("local-1")
        self.assertTrue(self.register("peer-1", "local-1"))
        popup.close()

        self.assertTrue(self.connected("peer-1"))
        self.assertTrue(self.tracker.snapshot().handled)
        self.assertEqual(self.namespace["ring_timeout_target"], -1.0)
        self.assertTrue(self.terminal("peer-1"))

        self.assert_finished(popup)
        self.assertEqual(self.missed, [])

    def test_peer_terminal_clears_ring_timer_and_restores_connected_status(self):
        popup = self.start("local-1")
        self.assertTrue(self.register("peer-1", "local-1"))

        self.assertTrue(self.terminal("peer-1"))

        self.assert_finished(popup)
        self.assertEqual(self.namespace["ring_timeout_target"], 0.0)
        self.assertEqual(self.statuses[-1:], [("🟢 已连接：COM5", "green")])

    def test_no_id_connected_without_active_call_cannot_change_runtime_state(self):
        for stale_popup in (False, True):
            with self.subTest(stale_popup=stale_popup):
                self.setUp()
                if stale_popup:
                    self.namespace["current_call_popup"] = SimpleNamespace(
                        _call_popup_caller_num="10086", winfo_exists=lambda: True,
                        _call_popup_mark_connected=lambda: self.fail("stale popup marked connected"),
                    )
                before = self.tracker.snapshot()
                self.namespace["ring_timeout_target"] = 321.0

                self.assertFalse(popup_runtime.mark_call_popup_connected_namespace_runtime(self.namespace))
                self.assertFalse(self.connected(""))

                self.assertEqual(self.tracker.snapshot(), before)
                self.assertEqual(self.namespace["ring_timeout_target"], 321.0)
                self.assertEqual(self.statuses, [])

    def test_peer_state_preserves_active_outgoing_timer_status_and_window(self):
        for phase in ("connected", "terminal"):
            with self.subTest(phase=phase):
                self.setUp()
                self.start("local-old")
                self.assertTrue(self.register("peer-old", "local-old"))
                dial_popup = SimpleNamespace(exists=True)
                self.namespace.update(
                    current_dial_num="10010", current_dial_popup=dial_popup,
                    ring_timeout_target=321.0,
                )
                self.statuses.append(("📞 拨号中：10010", "blue"))

                self.assertTrue(getattr(self, phase)("peer-old"))

                self.assertEqual(self.namespace["ring_timeout_target"], 321.0)
                self.assertEqual(self.namespace["current_dial_num"], "10010")
                self.assertEqual(self.statuses, [("📞 拨号中：10010", "blue")])
                self.assertIs(self.namespace["current_dial_popup"], dial_popup)
                self.assertTrue(dial_popup.exists)

    def test_pending_peer_caches_are_bounded_and_expire(self):
        for kind in ("incoming", "connected", "terminal"):
            with self.subTest(kind=kind):
                self.setUp()

                def receive(peer_id):
                    if kind == "incoming":
                        return self.register(peer_id, "local-" + peer_id)
                    return getattr(self, kind)(peer_id)

                for number in range(105):
                    self.assertTrue(receive(f"peer-{number}"))
                key = f"call_popup_pending_peer_{kind}"
                entries = self.namespace[key]
                self.assertEqual(len(entries), 100)
                self.assertEqual(set(entries), {f"peer-{number}" for number in range(5, 105)})

                self.now[0] += PEER_CALL_BINDING_TTL_SECONDS
                self.assertTrue(receive("peer-fresh"))
                self.assertEqual(set(self.namespace[key]), {"peer-fresh"})

    def test_expired_binding_does_not_clear_new_pending_messages(self):
        self.assertTrue(self.register("peer-new", "local-new"))
        self.assertTrue(self.connected("peer-new"))
        self.assertTrue(self.terminal("peer-new"))
        self.namespace.update(
            call_popup_peer_session_id="peer-old", call_popup_peer_caller_num="10086",
            call_popup_peer_local_session_id="local-old", call_popup_peer_expires_at=99.0,
        )

        self.assertIsNone(_get_peer_call_binding(self.namespace))
        popup = self.start("local-new")

        self.assert_finished(popup)
        self.assertEqual(self.missed, [])

    def test_finished_old_call_preserves_unassociated_new_terminal(self):
        old_popup = self.start("local-old")
        self.assertTrue(self.register("peer-old", "local-old"))
        self.terminal("peer-new")

        self.assertTrue(self.terminal("peer-old"))
        self.assert_finished(old_popup)
        self.missed.clear()
        self.assertTrue(self.register("peer-new", "local-new"))
        popup = self.start("local-new")

        self.assert_finished(popup)
        self.assertEqual(len(self.missed), 1)

    def test_legacy_single_pending_entry_survives_new_peer_arrival(self):
        self.namespace["call_popup_pending_peer_terminal"] = {
            "phase": "ended", "reason": "NO CARRIER", "session_id": "peer-new",
            "caller_num": "10086", "expires_at": self.now[0] + PEER_CALL_BINDING_TTL_SECONDS,
        }
        self.assertTrue(self.terminal("peer-old"))
        self.assertTrue(self.register("peer-new", "local-new"))

        popup = self.start("local-new")

        self.assert_finished(popup)
        self.assertEqual(len(self.missed), 1)

    def serial_callbacks(self, calls):
        return replace(
            runtime_callbacks(calls),
            start_incoming_call=self.tracker.start,
            finish_incoming_call=lambda: finish_incoming_call_session_namespace_runtime(self.namespace),
            close_call_popup=lambda: close_call_popup_namespace_runtime(self.namespace),
            show_call_popup=lambda _number, call_session_id="": self.show_popup(call_session_id),
            consume_remote_incoming_end=lambda local_id: popup_runtime.consume_remote_incoming_end_namespace_runtime(
                self.namespace, local_id),
        )

    def serial_line(self, state, callbacks, line, now):
        return handle_serial_runtime_line(
            state, line, now, "COM5", self.namespace["current_call_popup"] is not None,
            runtime_config(), callbacks, {})

    def test_remote_end_allows_serial_ring_and_same_number_clip_to_start_new_call(self):
        for remote_finish in (False, True):
            for phases in (("incoming", "ended"), ("ended", "incoming")):
                with self.subTest(remote_finish=remote_finish, phases=phases):
                    self.setUp()
                    state = SerialRuntimeState.create(parse_sms_callback_head)
                    callbacks = self.serial_callbacks([])
                    self.serial_line(state, callbacks, "RING", 100.0)
                    self.serial_line(state, callbacks, '+CLIP: "10086",129', 100.1)
                    old_id = state.call_state.call_session_id
                    old_popup = self.namespace["current_call_popup"]
                    self.assertTrue(old_id)
                    self.now[0] = 100.2
                    if remote_finish:
                        self.assertTrue(self.register("peer-old", old_id))
                        self.assertTrue(self.terminal("peer-old"))
                    else:
                        self.serial_line(state, callbacks, "NO CARRIER", 100.2)
                    self.assert_finished(old_popup)

                    self.serial_line(state, callbacks, "RING", 100.5)
                    self.serial_line(state, callbacks, '+CLIP: "10086",129', 101.0)
                    new_id = state.call_state.call_session_id
                    new_popup = self.namespace["current_call_popup"]
                    new_token = self.tracker.snapshot().session_token
                    self.assertTrue(new_id)
                    self.assertNotEqual(new_id, old_id)
                    self.assertEqual(state.call_state.call_session_sequence, 2)
                    self.assertIsNot(new_popup, old_popup)

                    self.now[0] = 101.1
                    for phase in phases:
                        self.assertTrue(handle_device_call_state_runtime(
                            {
                                "type": "device_call_state", "direction": "incoming",
                                "phase": phase, "phone": "10086", "call_session_id": "peer-old",
                                "local_call_session_id": old_id,
                                "call_session_ids": {"desktop": old_id, "firmware": "peer-old"},
                            },
                            runtime_imei=lambda: "000000000000001",
                            register_incoming=lambda peer, _caller, local: self.register(peer, local),
                            finish_call=lambda *args, **kwargs: finish_remote_incoming_call_namespace_runtime(
                                self.namespace, *args, **kwargs),
                        ))

                    self.assertEqual(state.call_state.call_session_id, new_id)
                    self.assertEqual(state.call_state.last_clip_num, "10086")
                    self.assertEqual(self.tracker.snapshot().caller_num, "10086")
                    self.assertIs(self.tracker.snapshot().session_token, new_token)
                    self.assertFalse(self.tracker.snapshot().handled)
                    self.assertIs(self.namespace["current_call_popup"], new_popup)
                    self.assertEqual(new_popup._call_popup_session_id, new_id)
                    self.assertTrue(new_popup.exists)

    def test_serial_consumes_remote_end_once_and_clears_only_incoming_fields(self):
        popup = self.start("local-old")
        self.assertTrue(self.register("peer-old", "local-old"))
        self.assertTrue(self.terminal("peer-old"))
        self.assert_finished(popup)
        state = SerialRuntimeState.create(parse_sms_callback_head)
        state.call_state = replace(
            state.call_state, call_session_id="local-old", ring_timeout_target=200.0,
            last_clip_num="10086", last_clip_time=99.0, last_hangup_time=90.0,
            call_session_sequence=8, pending_incoming_connected_until=110.0,
            incoming_connected_session_id="local-old", outgoing_connected_session_id="outgoing:earlier",
        )
        callbacks = self.serial_callbacks([])

        self.serial_line(state, callbacks, "OK", 101.0)

        self.assertEqual(state.call_state.ring_timeout_target, 0.0)
        self.assertEqual(state.call_state.call_session_id, "")
        self.assertEqual(state.call_state.last_clip_num, "")
        self.assertEqual(state.call_state.last_clip_time, 0.0)
        self.assertEqual(state.call_state.pending_incoming_connected_until, 0.0)
        self.assertEqual(state.call_state.incoming_connected_session_id, "")
        self.assertEqual(state.call_state.call_session_sequence, 8)
        self.assertEqual(state.call_state.last_hangup_time, 90.0)
        self.assertEqual(state.call_state.outgoing_connected_session_id, "outgoing:earlier")
        self.assertFalse(popup_runtime.consume_remote_incoming_end_namespace_runtime(
            self.namespace, "local-old"))

    def test_old_remote_end_does_not_clear_new_incoming_or_outgoing_serial_state(self):
        for outgoing in (False, True):
            with self.subTest(outgoing=outgoing):
                self.setUp()
                self.start("local-old")
                self.assertTrue(self.register("peer-old", "local-old"))
                self.assertTrue(self.terminal("peer-old"))
                state = SerialRuntimeState.create(parse_sms_callback_head)
                state.call_state = replace(
                    state.call_state,
                    call_session_id="outgoing:new" if outgoing else "local-new",
                    ring_timeout_target=200.0, last_clip_num="10086", last_clip_time=99.0,
                    call_session_sequence=8, current_dial_num="10010" if outgoing else "",
                    pending_incoming_connected_until=110.0,
                    incoming_connected_session_id="local-new", outgoing_connected_session_id="outgoing:new",
                )
                expected = replace(state.call_state)

                self.serial_line(state, self.serial_callbacks([]), "OK", 101.0)

                self.assertEqual(state.call_state, expected)
                self.assertTrue(popup_runtime.consume_remote_incoming_end_namespace_runtime(
                    self.namespace, "local-old"))

    def test_remote_end_with_matching_id_does_not_reset_active_outgoing_call(self):
        state = SerialRuntimeState.create(parse_sms_callback_head)
        state.call_state = replace(
            state.call_state, call_session_id="outgoing:current", current_dial_num="10010",
            ring_timeout_target=200.0, call_session_sequence=9,
        )
        callbacks = replace(
            runtime_callbacks([]),
            consume_remote_incoming_end=lambda _session_id: self.fail("outgoing call consumed incoming end"),
        )
        expected = replace(state.call_state)

        self.serial_line(state, callbacks, "OK", 101.0)

        self.assertEqual(state.call_state, expected)

    def test_serial_end_markers_are_bounded_and_expire_without_consuming_new_marker(self):
        for number in range(105):
            self.start(f"local-{number}")
            self.assertTrue(self.register(f"peer-{number}", f"local-{number}"))
            self.assertTrue(self.terminal(f"peer-{number}"))
        markers = self.namespace["call_popup_pending_serial_ends"]
        self.assertEqual(len(markers), 100)
        self.assertEqual(set(markers), {f"local-{number}" for number in range(5, 105)})

        self.now[0] += 120.0
        self.start("local-new")
        self.assertTrue(self.register("peer-new", "local-new"))
        self.assertTrue(self.terminal("peer-new"))

        self.assertFalse(popup_runtime.consume_remote_incoming_end_namespace_runtime(
            self.namespace, "local-104"))
        self.assertEqual(set(self.namespace["call_popup_pending_serial_ends"]), {"local-new"})
        self.assertTrue(popup_runtime.consume_remote_incoming_end_namespace_runtime(
            self.namespace, "local-new"))


if __name__ == "__main__":
    unittest.main()
