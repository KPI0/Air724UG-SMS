import unittest

from types import SimpleNamespace
from datetime import datetime
import threading
import itertools
import gc
import tkinter as tk

from sms_core.call_session import IncomingCallSessionTracker, MissedCall
from sms_core.cloud_message_runtime import handle_device_call_state_runtime
from sms_ui.thread_runtime import run_on_ui_thread

from sms_ui.call_popup_namespace_runtime import (
    RETIRED_PEER_CALL_MAX_ENTRIES,
    RETIRED_PEER_CALL_TTL_SECONDS,
    close_call_popup_namespace_runtime,
    close_missed_call_popup_namespace_runtime,
    close_phone_popups_namespace_runtime,
    finish_incoming_call_session_namespace_runtime,
    finish_remote_incoming_call_namespace_runtime,
    get_serial_call_state_namespace_runtime,
    mark_call_popup_connected_namespace_runtime,
    mark_dial_popup_connected_namespace_runtime,
    finish_dial_popup_namespace_runtime,
    mark_incoming_call_handled_namespace_runtime,
    mark_peer_incoming_call_connected_namespace_runtime,
    register_peer_incoming_call_namespace_runtime,
    _get_peer_call_binding,
    reset_incoming_call_session_namespace_runtime,
    set_call_popup_namespace_runtime,
    set_dial_popup_namespace_runtime,
    set_missed_call_popup_namespace_runtime,
    set_serial_call_state_namespace_runtime,
    show_call_popup_namespace_runtime,
    show_missed_call_popup_namespace_runtime,
    start_incoming_call_session_namespace_runtime,
)


class CallPopupNamespaceRuntimeTests(unittest.TestCase):
    def test_incoming_peer_requires_explicit_local_session_without_number_guessing(self):
        for active in (False, True):
            with self.subTest(active=active):
                namespace = self.base_namespace()
                tracker = IncomingCallSessionTracker()
                tracker.start("10086")
                namespace.update(
                    INCOMING_CALL_SESSION=tracker,
                    call_popup_active_session_id="local-1" if active else "",
                    current_call_popup=SimpleNamespace(_call_popup_caller_num="10086"),
                )
                if active:
                    self.announce_local_session(namespace, "local-1")
                for local_id in (None, "", "   "):
                    self.assertFalse(register_peer_incoming_call_namespace_runtime(
                        namespace, "peer-unknown", "10086", local_id))
                    self.assertIsNone(_get_peer_call_binding(namespace))
                self.assertFalse(register_peer_incoming_call_namespace_runtime(
                    namespace, "peer-unknown", "10086"))
                self.assertIsNone(_get_peer_call_binding(namespace))
                self.assertTrue(register_peer_incoming_call_namespace_runtime(
                    namespace, "peer-1", "10086", "local-1"))
                if not active:
                    popup = namespace["current_call_popup"]
                    self.announce_local_session(namespace, "local-unrelated")
                    set_call_popup_namespace_runtime(namespace, popup)
                    self.assertIsNone(_get_peer_call_binding(namespace))
                    self.announce_local_session(namespace, "local-1")
                    set_call_popup_namespace_runtime(namespace, popup)
                binding = _get_peer_call_binding(namespace)
                self.assertIsNotNone(binding)
                self.assertEqual(binding["peer_session_id"], "peer-1")
                self.assertEqual(binding["local_session_id"], "local-1")

    def test_existing_peer_correlation_cannot_be_retargeted_to_same_number_redial(self):
        namespace = self.base_namespace()
        namespace.update(
            call_popup_active_session_id="local-old",
            current_call_popup=SimpleNamespace(_call_popup_caller_num="10086"),
        )
        self.assertTrue(register_peer_incoming_call_namespace_runtime(
            namespace, "peer-old", "10086", "local-old"))
        self.assertFalse(register_peer_incoming_call_namespace_runtime(
            namespace, "peer-old", "10086", "local-new"))
        binding = _get_peer_call_binding(namespace)
        self.assertEqual(binding["peer_session_id"], "peer-old")
        self.assertEqual(binding["local_session_id"], "local-old")

    def test_replayed_finished_peer_call_cannot_remove_next_ringing_popup(self):
        for order in itertools.permutations(("incoming", "connected", "ended")):
            with self.subTest(order=order):
                namespace = self.base_namespace()
                tracker = IncomingCallSessionTracker()
                namespace.update(INCOMING_CALL_SESSION=tracker, current_call_popup=None)
                missed, destroyed = [], []
                namespace["show_missed_call_popup"] = missed.append

                def show(**kwargs):
                    popup = SimpleNamespace(
                        _call_popup_caller_num="10086", winfo_exists=lambda: True,
                        destroy=lambda: destroyed.append(popup),
                        _call_popup_mark_connected=lambda: tracker.mark_handled(),
                    )
                    kwargs["set_popup"](popup)
                    return popup

                for generation in range(1, 6):
                    tracker.start("10086")
                    popup = show_call_popup_namespace_runtime(
                        namespace, "10086", f"local-{generation}", show_app_runtime=show)
                    if generation > 1:
                        for phase in order:
                            handle_device_call_state_runtime(
                                {"type": "device_call_state", "direction": "incoming",
                                 "phone": "10086", "phase": phase,
                                 "call_session_id": f"peer-{generation - 1}",
                                 "local_call_session_id": f"local-{generation - 1}",
                                 "call_session_ids": {
                                     "desktop": f"local-{generation - 1}",
                                     "firmware": f"peer-{generation - 1}",
                                 }},
                                runtime_imei=lambda: "000000000000001",
                                register_incoming=lambda *args: register_peer_incoming_call_namespace_runtime(namespace, *args),
                                mark_peer_call_connected=lambda *args: mark_peer_incoming_call_connected_namespace_runtime(namespace, *args),
                                finish_call=lambda *args, **kwargs: finish_remote_incoming_call_namespace_runtime(namespace, *args, **kwargs),
                            )
                        self.assertIs(namespace["current_call_popup"], popup)
                        self.assertEqual(namespace["call_popup_active_session_id"], f"local-{generation}")
                        self.assertEqual(tracker.snapshot().caller_num, "10086")
                        self.assertFalse(tracker.snapshot().handled)
                        self.assertEqual(missed, [])
                    self.assertTrue(register_peer_incoming_call_namespace_runtime(
                        namespace, f"peer-{generation}", "10086", f"local-{generation}"))
                    finish_incoming_call_session_namespace_runtime(namespace)
                    close_call_popup_namespace_runtime(namespace)
                self.assertEqual(len(destroyed), 5)

    def test_remote_finish_and_reset_do_not_revive_finished_peer(self):
        for finish_remote in (True, False):
            with self.subTest(finish_remote=finish_remote):
                namespace = self.base_namespace()
                tracker = IncomingCallSessionTracker()
                tracker.start("10086")
                namespace.update(
                    INCOMING_CALL_SESSION=tracker,
                    current_call_popup=None,
                )
                self.announce_local_session(namespace, "local-old")
                self.assertTrue(register_peer_incoming_call_namespace_runtime(
                    namespace, "peer-old", "10086", "local-old"))
                if finish_remote:
                    self.assertTrue(finish_remote_incoming_call_namespace_runtime(
                        namespace, session_id="peer-old", caller_num="10086"))
                else:
                    reset_incoming_call_session_namespace_runtime(namespace)

                tracker.start("10086")
                self.announce_local_session(namespace, "local-new")
                self.assertFalse(register_peer_incoming_call_namespace_runtime(
                    namespace, "peer-old", "10086", "local-old"))
                self.assertFalse(mark_peer_incoming_call_connected_namespace_runtime(
                    namespace, "peer-old", "10086"))
                self.assertFalse(finish_remote_incoming_call_namespace_runtime(
                    namespace, session_id="peer-old", caller_num="10086"))
                self.assertEqual(tracker.snapshot().caller_num, "10086")
                self.assertFalse(tracker.snapshot().handled)
                self.assertTrue(register_peer_incoming_call_namespace_runtime(
                    namespace, "peer-new", "10086", "local-new"))
                self.assertTrue(mark_peer_incoming_call_connected_namespace_runtime(
                    namespace, "peer-new", "10086"))
                connected = []
                popup = SimpleNamespace(
                    _call_popup_caller_num="10086",
                    _call_popup_mark_connected=lambda: connected.append(True),
                    winfo_exists=lambda: True,
                    destroy=lambda: None,
                )
                namespace["mark_incoming_call_handled"] = tracker.mark_handled
                set_call_popup_namespace_runtime(namespace, popup)
                self.assertEqual(connected, [True])
                self.assertTrue(tracker.snapshot().handled)
                missed = []
                namespace["show_missed_call_popup"] = missed.append
                self.assertTrue(finish_remote_incoming_call_namespace_runtime(
                    namespace, session_id="peer-new", caller_num="10086"))
                self.assertIsNone(namespace["current_call_popup"])
                self.assertEqual(missed, [])

    def test_retired_peer_protection_starts_when_long_call_finishes(self):
        namespace = self.base_namespace()
        now = [100.0]
        namespace["_call_peer_monotonic"] = lambda: now[0]
        self.assertTrue(register_peer_incoming_call_namespace_runtime(
            namespace, "peer-old", "10086", "local-old"))
        namespace["call_popup_active_session_id"] = "local-old"
        set_call_popup_namespace_runtime(namespace, SimpleNamespace(
            _call_popup_caller_num="10086"))
        now[0] += 3600
        finish_incoming_call_session_namespace_runtime(namespace)
        namespace.update(current_call_popup=None, call_popup_active_session_id="")
        now[0] += RETIRED_PEER_CALL_TTL_SECONDS - 0.1
        self.assertFalse(register_peer_incoming_call_namespace_runtime(
            namespace, "peer-old", "10086", "local-old"))
        now[0] += 0.1
        self.assertTrue(register_peer_incoming_call_namespace_runtime(
            namespace, "peer-old", "10086", "local-old"))

    def test_retired_peer_cache_is_bounded_and_namespace_local(self):
        namespace = self.base_namespace()
        tracker = IncomingCallSessionTracker()
        namespace["INCOMING_CALL_SESSION"] = tracker
        for generation in range(RETIRED_PEER_CALL_MAX_ENTRIES + 1):
            tracker.start("10086")
            self.announce_local_session(namespace, f"local-{generation}")
            self.assertTrue(register_peer_incoming_call_namespace_runtime(
                namespace, f"peer-{generation}", "10086", f"local-{generation}"))
            finish_incoming_call_session_namespace_runtime(namespace)
        self.assertEqual(len(namespace["call_popup_retired_peer_sessions"]),
                         RETIRED_PEER_CALL_MAX_ENTRIES)
        self.assertFalse(register_peer_incoming_call_namespace_runtime(
            namespace, "peer-1", "10086", "local-1"))
        tracker.start("10086")
        self.announce_local_session(namespace, "local-reused")
        self.assertTrue(register_peer_incoming_call_namespace_runtime(
            namespace, "peer-0", "10086", "local-reused"))
        self.assertTrue(register_peer_incoming_call_namespace_runtime(
            self.base_namespace(), "peer-1", "10086", "local-1"))
        namespace["_call_peer_monotonic"] = lambda: 1000.0
        finish_incoming_call_session_namespace_runtime(namespace)
        self.assertEqual(list(namespace["call_popup_retired_peer_sessions"]), ["peer-0"])

    def test_expiring_unbound_peer_does_not_retire_a_live_generation(self):
        namespace = self.base_namespace()
        self.assertTrue(register_peer_incoming_call_namespace_runtime(
            namespace, "peer-1", "10086", "local-1"))
        namespace["_call_peer_monotonic"] = lambda: 200.0
        self.assertIsNone(_get_peer_call_binding(namespace))
        self.assertTrue(register_peer_incoming_call_namespace_runtime(
            namespace, "peer-1", "10086", "local-1"))

    def test_retired_peer_replay_preserves_pending_new_terminal(self):
        namespace = self.base_namespace()
        tracker = IncomingCallSessionTracker()
        namespace.update(INCOMING_CALL_SESSION=tracker, current_call_popup=None)
        tracker.start("10086")
        self.announce_local_session(namespace, "local-old")
        register_peer_incoming_call_namespace_runtime(namespace, "peer-old", "10086", "local-old")
        finish_incoming_call_session_namespace_runtime(namespace)
        namespace["call_popup_active_session_id"] = ""
        self.assertTrue(finish_remote_incoming_call_namespace_runtime(
            namespace, session_id="peer-new", caller_num="10086"))
        self.assertFalse(finish_remote_incoming_call_namespace_runtime(
            namespace, session_id="peer-old", caller_num="10086"))
        self.assertFalse(register_peer_incoming_call_namespace_runtime(
            namespace, "peer-old", "10086", "local-old"))
        self.assertTrue(register_peer_incoming_call_namespace_runtime(
            namespace, "peer-new", "10086", "local-new"))
        tracker.start("10086")
        self.announce_local_session(namespace, "local-new")
        missed = []
        namespace["show_missed_call_popup"] = missed.append
        set_call_popup_namespace_runtime(namespace, SimpleNamespace(
            _call_popup_caller_num="10086",
            winfo_exists=lambda: True,
            destroy=lambda: None,
        ))
        self.assertIsNone(namespace["current_call_popup"])
        self.assertEqual(len(missed), 1)

    def test_real_popups_coexist_during_repeated_calls_and_peer_replays(self):
        try:
            root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk display unavailable: {exc}")
        root.withdraw()
        errors = []
        root.report_callback_exception = lambda *args: errors.append(args)
        tracker = IncomingCallSessionTracker()
        namespace = self.base_namespace()

        def cleanup():
            root.destroy()
            namespace.clear()
            gc.collect()

        self.addCleanup(cleanup)

        def position_popup(window, _parent):
            window.attributes("-alpha", 0)
            window.geometry("+100+100")

        namespace.update(
            root=root, INCOMING_CALL_SESSION=tracker,
            current_call_popup=None, current_missed_call_popup=None,
            center_window=position_popup, run_on_ui_thread=run_on_ui_thread,
            serial_lock=threading.Lock(), serial_obj=None, ui_post=lambda cb: cb(),
            show_window=lambda: None,
            close_call_popup=lambda: close_call_popup_namespace_runtime(namespace),
            show_missed_call_popup=lambda call: show_missed_call_popup_namespace_runtime(namespace, call),
        )
        show_missed_call_popup_namespace_runtime(namespace, MissedCall("10086", datetime(2026, 1, 1)))
        missed_popup = namespace["current_missed_call_popup"]
        for generation in range(1, 6):
            tracker.start("10086")
            popup = show_call_popup_namespace_runtime(namespace, "10086", f"local-{generation}")
            if generation > 1:
                old_peer = f"peer-{generation - 1}"
                register_peer_incoming_call_namespace_runtime(namespace, old_peer, "10086", f"local-{generation - 1}")
                mark_peer_incoming_call_connected_namespace_runtime(namespace, old_peer, "10086")
                finish_remote_incoming_call_namespace_runtime(namespace, session_id=old_peer, caller_num="10086")
            root.update()
            self.assertIs(namespace["current_call_popup"], popup)
            self.assertIs(namespace["current_missed_call_popup"], missed_popup)
            self.assertTrue(popup.winfo_viewable())
            self.assertTrue(missed_popup.winfo_viewable())
            self.assertIsNone(root.grab_current())
            buttons = popup.winfo_children()[0].winfo_children()[-1].winfo_children()
            self.assertEqual([button.cget("text") for button in buttons], ["✅ 接听", "❌ 挂断", "隐藏"])
            self.assertTrue(all(button.winfo_ismapped() for button in buttons))
            self.assertEqual(missed_popup.missed_call_popup_count, generation)
            self.assertTrue(register_peer_incoming_call_namespace_runtime(namespace, f"peer-{generation}", "10086", f"local-{generation}"))
            missed = finish_incoming_call_session_namespace_runtime(namespace)
            close_call_popup_namespace_runtime(namespace)
            show_missed_call_popup_namespace_runtime(namespace, missed)
            root.update()
            self.assertFalse(popup.winfo_exists())
        close_missed_call_popup_namespace_runtime(namespace)
        root.update()
        self.assertEqual(errors, [])

    def queued_terminal_case(self, *, local_session=True, show_popup=True, peer_session=""):
        now = [datetime(2026, 1, 1)]
        tracker = IncomingCallSessionTracker(now_func=lambda: now[0])
        tracker.start("10086")
        queued, destroyed, missed = [], [], []

        def make_popup(session):
            return SimpleNamespace(
                _call_popup_session_id=session,
                _call_popup_caller_num="10086",
                winfo_exists=lambda: True,
                destroy=lambda: destroyed.append(session),
            )

        session = "incoming:desktop:old" if local_session else ""
        namespace = self.base_namespace()
        namespace.update({
            "INCOMING_CALL_SESSION": tracker,
            "current_call_popup": make_popup(session) if show_popup else None,
            "call_popup_active_session_id": session,
            "ui_post": queued.append,
            "run_on_ui_thread": run_on_ui_thread,
            "show_missed_call_popup": missed.append,
        })
        self.announce_local_session(namespace, session)
        if peer_session:
            register_peer_incoming_call_namespace_runtime(namespace, peer_session, "10086", session)

        def receive():
            handle_device_call_state_runtime(
                {"type": "device_call_state", "direction": "incoming", "phase": "ended",
                 "phone": "10086", "call_session_id": peer_session},
                runtime_imei=lambda: "123456789012345",
                finish_call=lambda *args, **kwargs: finish_remote_incoming_call_namespace_runtime(
                    namespace, *args, **kwargs),
            )

        worker = threading.Thread(target=receive)
        worker.start()
        worker.join(timeout=3)
        self.assertFalse(worker.is_alive())
        self.assertEqual(len(queued), 1)
        return namespace, now, queued, destroyed, missed, make_popup

    def test_queued_legacy_terminal_preserves_same_number_redial(self):
        for local_session, show_popup in ((True, True), (False, True), (False, False)):
            with self.subTest(local_session=local_session, show_popup=show_popup):
                namespace, now, queued, destroyed, missed, make_popup = self.queued_terminal_case(
                    local_session=local_session, show_popup=show_popup)
                tracker = namespace["INCOMING_CALL_SESSION"]
                tracker.finish()
                # A new call must stay isolated even when the wall clock has
                # not advanced between the old finish and the new start.
                tracker.start("10086")
                new_id = "incoming:desktop:new" if local_session else ""
                new_popup = make_popup(new_id) if show_popup else None
                namespace.update(current_call_popup=new_popup, call_popup_active_session_id=new_id)
                self.assertTrue(queued.pop()())
                self.assertIs(namespace["current_call_popup"], new_popup)
                self.assertEqual(namespace["call_popup_active_session_id"], new_id)
                self.assertEqual(tracker.snapshot().started_at, now[0])
                self.assertEqual(destroyed, ["incoming:desktop:old" if local_session else ""] if show_popup else [])
                self.assertEqual(len(missed), 1)

    def test_queued_terminal_invalidates_late_show_and_finishes_captured_popup(self):
        for peer, show_popup in (("", True), ("", False), ("incoming:firmware:old", True)):
            with self.subTest(peer=peer, show_popup=show_popup):
                namespace, _now, queued, destroyed, missed, make_popup = self.queued_terminal_case(
                    show_popup=show_popup, peer_session=peer)
                self.assertIsNone(namespace["INCOMING_CALL_SESSION"].snapshot().session_token)
                if not show_popup:
                    self.assertEqual(show_call_popup_namespace_runtime(
                        namespace, "10086", "incoming:desktop:old"), "disabled")
                self.assertTrue(queued.pop()())
                self.assertIsNone(namespace["current_call_popup"])
                self.assertEqual(namespace["INCOMING_CALL_SESSION"].snapshot().caller_num, "")
                self.assertEqual(destroyed, ["incoming:desktop:old"] if show_popup else [])
                self.assertEqual(len(missed), 1)

    def test_queued_peer_terminal_preserves_replaced_local_session(self):
        namespace, _now, queued, destroyed, missed, make_popup = self.queued_terminal_case(
            peer_session="incoming:firmware:old")
        new_popup = make_popup("incoming:desktop:new")
        namespace["INCOMING_CALL_SESSION"].start("10086")
        namespace.update(current_call_popup=new_popup, call_popup_active_session_id="incoming:desktop:new")
        self.announce_local_session(namespace, "incoming:desktop:new")
        self.assertTrue(queued.pop()())
        self.assertIs(namespace["current_call_popup"], new_popup)
        self.assertEqual(namespace["call_popup_active_session_id"], "incoming:desktop:new")
        self.assertEqual(namespace["INCOMING_CALL_SESSION"].snapshot().caller_num, "10086")
        self.assertEqual(destroyed, ["incoming:desktop:old"])
        self.assertEqual(len(missed), 1)

    def announce_local_session(self, namespace, local_session_id):
        show_call_popup_namespace_runtime(
            namespace,
            namespace["INCOMING_CALL_SESSION"].snapshot().caller_num,
            local_session_id,
            show_app_runtime=lambda **_kwargs: None,
        )

    def base_namespace(self):
        tracker = SimpleNamespace(
            start=lambda caller_num: ("start", caller_num),
            mark_handled=lambda: "handled",
            finish=lambda: "finished",
            reset=lambda: "reset",
        )
        return {
            "root": "root",
            "current_call_popup": "popup",
            "current_dial_popup": None,
            "current_missed_call_popup": "missed_popup",
            "INCOMING_CALL_SESSION": tracker,
            "CALL_POPUP_ENABLED": True,
            "ring_timeout_target": 3.0,
            "current_dial_num": "10086",
            "center_window": "center",
            "serial_lock": "lock",
            "serial_obj": "serial",
            "port_ui": lambda *args: ("port", args),
            "set_status": lambda *args: ("status", args),
            "ui_post": "post",
            "close_call_popup": lambda: "closed",
            "close_missed_call_popup": lambda: "missed_closed",
            "show_window": "show_window",
            "run_on_ui_thread": lambda callback, ui_post: callback(),
            "log_file_only": lambda message: ("log", message),
            "_call_peer_monotonic": lambda: 100.0,
        }

    def test_set_call_popup_namespace_runtime_updates_window(self):
        namespace = self.base_namespace()

        result = set_call_popup_namespace_runtime(namespace, "next")

        self.assertIsNone(result)
        self.assertEqual(namespace["current_call_popup"], "next")

    def test_set_dial_popup_namespace_runtime_updates_window(self):
        namespace = self.base_namespace()

        result = set_dial_popup_namespace_runtime(namespace, "dial_popup")

        self.assertIsNone(result)
        self.assertEqual(namespace["current_dial_popup"], "dial_popup")

    def test_mark_dial_popup_connected_runs_marker_on_ui_thread(self):
        namespace = self.base_namespace()
        calls = []
        popup = SimpleNamespace(
            winfo_exists=lambda: True,
            _call_popup_mark_connected=lambda: calls.append("connected"),
        )
        namespace["current_dial_popup"] = popup

        result = mark_dial_popup_connected_namespace_runtime(namespace)

        self.assertTrue(result)
        self.assertEqual(calls, ["connected"])

    def test_mark_call_popup_connected_runs_marker_on_ui_thread(self):
        namespace = self.base_namespace()
        calls = []
        popup = SimpleNamespace(
            winfo_exists=lambda: True,
            _call_popup_mark_connected=lambda: calls.append("connected"),
        )
        namespace["current_call_popup"] = popup
        self.assertTrue(
            register_peer_incoming_call_namespace_runtime(
                namespace,
                "incoming:firmware:1",
                "15923240141",
                "incoming:desktop:1",
            )
        )

        result = mark_call_popup_connected_namespace_runtime(namespace)

        self.assertTrue(result)
        self.assertEqual(calls, ["connected"])

    def test_finish_remote_incoming_call_closes_popup_without_hangup(self):
        class Tracker:
            def __init__(self):
                self.finished = 0

            def snapshot(self):
                return SimpleNamespace(caller_num="15923240141")

            def finish(self):
                self.finished += 1
                return None

        namespace = self.base_namespace()
        tracker = Tracker()
        namespace["INCOMING_CALL_SESSION"] = tracker
        namespace["call_popup_active_session_id"] = "incoming:desktop:1"
        namespace["ui_post"] = "post"
        closed = []
        namespace["close_call_popup"] = lambda: closed.append("closed")
        popup = SimpleNamespace(
            _call_popup_session_id="incoming:desktop:1",
            _call_popup_caller_num="15923240141",
        )
        namespace["current_call_popup"] = popup
        self.assertTrue(
            register_peer_incoming_call_namespace_runtime(
                namespace,
                "incoming:firmware:1",
                "15923240141",
                "incoming:desktop:1",
            )
        )

        result = finish_remote_incoming_call_namespace_runtime(
            namespace,
            "no_answer",
            "NO ANSWER",
            session_id="incoming:firmware:1",
            caller_num="15923240141",
        )

        self.assertTrue(result)
        self.assertEqual(tracker.finished, 1)
        self.assertEqual(namespace["current_call_popup"], None)
        self.assertEqual(closed, [])
        self.assertEqual(namespace["call_popup_active_session_id"], "")

    def test_peer_generation_binds_to_local_generation_before_connected(self):
        namespace = self.base_namespace()
        namespace["call_popup_active_session_id"] = "incoming:desktop:1"
        calls = []
        namespace["current_call_popup"] = SimpleNamespace(
            winfo_exists=lambda: True,
            _call_popup_session_id="incoming:desktop:1",
            _call_popup_caller_num="15923240141",
            _call_popup_mark_connected=lambda: calls.append("connected"),
        )

        self.assertTrue(
            register_peer_incoming_call_namespace_runtime(
                namespace,
                "incoming:firmware:1",
                "15923240141",
                "incoming:desktop:1",
            )
        )
        self.assertTrue(
            mark_peer_incoming_call_connected_namespace_runtime(
                namespace,
                "incoming:firmware:1",
                "15923240141",
            )
        )

        self.assertEqual(calls, ["connected"])
        self.assertEqual(
            namespace["call_popup_peer_local_session_id"],
            "incoming:desktop:1",
        )

    def test_old_peer_terminal_cannot_close_same_number_redial(self):
        class Tracker:
            def __init__(self):
                self.finished = 0

            def snapshot(self):
                return SimpleNamespace(caller_num="15923240141")

            def finish(self):
                self.finished += 1
                return None

        namespace = self.base_namespace()
        tracker = Tracker()
        namespace["INCOMING_CALL_SESSION"] = tracker
        namespace["call_popup_active_session_id"] = "incoming:desktop:new"
        namespace["current_call_popup"] = SimpleNamespace(
            _call_popup_session_id="incoming:desktop:new",
            _call_popup_caller_num="15923240141",
        )
        self.assertTrue(
            register_peer_incoming_call_namespace_runtime(
                namespace,
                "incoming:firmware:new",
                "15923240141",
                "incoming:desktop:new",
            )
        )

        result = finish_remote_incoming_call_namespace_runtime(
            namespace,
            "ended",
            "NO CARRIER",
            session_id="incoming:firmware:old",
            caller_num="15923240141",
        )

        self.assertTrue(result)
        self.assertEqual(tracker.finished, 0)
        self.assertEqual(
            namespace["call_popup_pending_peer_terminal"]["incoming:firmware:old"]["session_id"],
            "incoming:firmware:old",
        )
        self.assertEqual(
            _get_peer_call_binding(namespace)["peer_session_id"],
            "incoming:firmware:new",
        )
        self.assertEqual(
            namespace["call_popup_active_session_id"],
            "incoming:desktop:new",
        )

    def test_bound_peer_generation_survives_pending_ttl_until_call_finishes(self):
        namespace = self.base_namespace()
        now = [100.0]
        namespace["_call_peer_monotonic"] = lambda: now[0]
        namespace["call_popup_active_session_id"] = "incoming:desktop:1"
        namespace["current_call_popup"] = SimpleNamespace(
            _call_popup_session_id="incoming:desktop:1",
            _call_popup_caller_num="15923240141",
        )
        self.assertTrue(
            register_peer_incoming_call_namespace_runtime(
                namespace,
                "incoming:firmware:1",
                "15923240141",
                "incoming:desktop:1",
            )
        )
        namespace["call_popup_peer_local_session_id"] = "incoming:desktop:1"

        now[0] = 1000.0

        self.assertIsNotNone(_get_peer_call_binding(namespace))

    def test_terminal_before_incoming_is_consumed_after_local_popup_binds(self):
        class Tracker:
            def __init__(self):
                self.caller_num = ""
                self.finished = 0

            def snapshot(self):
                return SimpleNamespace(caller_num=self.caller_num)

            def finish(self):
                self.finished += 1
                self.caller_num = ""
                return None

        namespace = self.base_namespace()
        tracker = Tracker()
        namespace["INCOMING_CALL_SESSION"] = tracker
        namespace["current_call_popup"] = None
        namespace["call_popup_active_session_id"] = ""

        self.assertTrue(
            finish_remote_incoming_call_namespace_runtime(
                namespace,
                "no_answer",
                "NO ANSWER",
                session_id="incoming:firmware:1",
                caller_num="15923240141",
            )
        )
        self.assertTrue(
            register_peer_incoming_call_namespace_runtime(
                namespace,
                "incoming:firmware:1",
                "15923240141",
                "incoming:desktop:1",
            )
        )
        tracker.caller_num = "15923240141"
        namespace["call_popup_active_session_id"] = "incoming:desktop:1"
        popup = SimpleNamespace(
            _call_popup_session_id="incoming:desktop:1",
            _call_popup_caller_num="15923240141",
            winfo_exists=lambda: True,
            destroy=lambda: None,
        )

        set_call_popup_namespace_runtime(namespace, popup)

        self.assertEqual(tracker.finished, 1)
        self.assertIsNone(namespace["current_call_popup"])

    def test_connected_before_incoming_is_consumed_after_local_popup_binds(self):
        namespace = self.base_namespace()
        namespace["current_call_popup"] = None
        namespace["call_popup_active_session_id"] = ""
        calls = []

        self.assertTrue(
            mark_peer_incoming_call_connected_namespace_runtime(
                namespace,
                "incoming:firmware:1",
                "15923240141",
            )
        )
        self.assertTrue(
            register_peer_incoming_call_namespace_runtime(
                namespace,
                "incoming:firmware:1",
                "15923240141",
                "incoming:desktop:1",
            )
        )
        namespace["call_popup_active_session_id"] = "incoming:desktop:1"
        popup = SimpleNamespace(
            _call_popup_session_id="incoming:desktop:1",
            _call_popup_caller_num="15923240141",
            _call_popup_mark_connected=lambda: calls.append("connected"),
            winfo_exists=lambda: True,
        )

        set_call_popup_namespace_runtime(namespace, popup)

        self.assertEqual(calls, ["connected"])

    def test_mark_call_popup_connected_rejects_unknown_generation_with_matching_caller(self):
        namespace = self.base_namespace()
        namespace["call_popup_active_session_id"] = "incoming:desktop:1"
        calls = []
        popup = SimpleNamespace(
            winfo_exists=lambda: True,
            _call_popup_session_id="incoming:desktop:1",
            _call_popup_caller_num="15923240141",
            _call_popup_mark_connected=lambda: calls.append("connected"),
        )
        namespace["current_call_popup"] = popup

        result = mark_call_popup_connected_namespace_runtime(
            namespace,
            "incoming:firmware:1",
            "15923240141",
        )

        self.assertFalse(result)
        self.assertEqual(calls, [])

    def test_mark_call_popup_connected_is_deferred_until_popup_is_registered(self):
        namespace = self.base_namespace()
        namespace["current_call_popup"] = None
        namespace["call_popup_active_session_id"] = "incoming:1"
        namespace["call_popup_pending_connected_session_id"] = ""
        calls = []

        result = mark_call_popup_connected_namespace_runtime(namespace, "incoming:1")

        self.assertTrue(result)
        self.assertEqual(namespace["call_popup_pending_connected_session_id"], "incoming:1")

        popup = SimpleNamespace(
            winfo_exists=lambda: True,
            _call_popup_mark_connected=lambda: calls.append("connected"),
        )
        set_call_popup_namespace_runtime(namespace, popup)

        self.assertEqual(calls, ["connected"])
        self.assertEqual(namespace["call_popup_pending_connected_session_id"], "")

    def test_late_connected_state_cannot_update_a_new_or_closed_session(self):
        namespace = self.base_namespace()
        namespace["current_call_popup"] = None
        namespace["call_popup_active_session_id"] = "incoming:new"
        namespace["call_popup_pending_connected_session_id"] = ""

        self.assertFalse(
            mark_call_popup_connected_namespace_runtime(namespace, "incoming:old")
        )
        self.assertEqual(namespace["call_popup_pending_connected_session_id"], "")

        namespace["call_popup_active_session_id"] = ""
        self.assertFalse(
            mark_call_popup_connected_namespace_runtime(namespace, "incoming:new")
        )
        self.assertEqual(namespace["call_popup_pending_connected_session_id"], "")

    def test_finish_dial_popup_runs_terminal_marker_on_ui_thread(self):
        namespace = self.base_namespace()
        calls = []
        popup = SimpleNamespace(
            winfo_exists=lambda: True,
            _call_popup_mark_ended=lambda message: calls.append(message),
        )
        namespace["current_dial_popup"] = popup

        result = finish_dial_popup_namespace_runtime(namespace, "📞 对方已挂断")

        self.assertTrue(result)
        self.assertEqual(calls, ["📞 对方已挂断"])

    def test_close_call_popup_namespace_runtime_forwards_state_callbacks(self):
        namespace = self.base_namespace()
        calls = []

        result = close_call_popup_namespace_runtime(
            namespace,
            close_app_runtime=lambda **kwargs: calls.append(kwargs) or "closed",
        )

        self.assertEqual(result, "closed")
        forwarded = calls[0]
        self.assertEqual(forwarded["get_popup"](), "popup")
        forwarded["set_popup"](None)
        self.assertIsNone(namespace["current_call_popup"])
        self.assertEqual(forwarded["ui_post"], "post")
        self.assertEqual(forwarded["log_error"]("close log"), ("log", "close log"))

    def test_close_call_popup_namespace_runtime_also_closes_dial_popup(self):
        namespace = self.base_namespace()
        namespace["current_dial_popup"] = "dial_popup"
        calls = []

        result = close_call_popup_namespace_runtime(
            namespace,
            close_app_runtime=lambda **kwargs: calls.append(kwargs) or "closed",
        )

        self.assertEqual(result, "closed")
        self.assertEqual(len(calls), 2)
        self.assertEqual(calls[1]["get_popup"](), "dial_popup")
        calls[1]["set_popup"](None)
        self.assertIsNone(namespace["current_dial_popup"])

    def test_close_missed_call_popup_namespace_runtime_forwards_state_callbacks(self):
        namespace = self.base_namespace()
        calls = []

        result = close_missed_call_popup_namespace_runtime(
            namespace,
            close_app_runtime=lambda **kwargs: calls.append(kwargs) or "closed",
        )

        self.assertEqual(result, "closed")
        forwarded = calls[0]
        self.assertEqual(forwarded["get_popup"](), "missed_popup")
        forwarded["set_popup"](None)
        self.assertIsNone(namespace["current_missed_call_popup"])

    def test_close_phone_popups_namespace_runtime_closes_both_windows(self):
        namespace = self.base_namespace()
        calls = []
        namespace["close_call_popup"] = lambda: calls.append("incoming")
        namespace["close_missed_call_popup"] = lambda: calls.append("missed")

        close_phone_popups_namespace_runtime(namespace)

        self.assertEqual(calls, ["incoming", "missed"])

    def test_close_phone_popups_continues_when_one_window_close_fails(self):
        namespace = self.base_namespace()
        calls = []

        def fail_close():
            raise RuntimeError("failed")

        namespace["close_call_popup"] = fail_close
        namespace["close_missed_call_popup"] = lambda: calls.append("missed")
        namespace["log_file_only"] = calls.append

        close_phone_popups_namespace_runtime(namespace)

        self.assertEqual(calls[-1], "missed")
        self.assertTrue(any("close_call_popup" in str(item) for item in calls))

    def test_show_call_popup_namespace_runtime_forwards_dependencies_and_sets_ring_timeout(self):
        namespace = self.base_namespace()
        calls = []

        result = show_call_popup_namespace_runtime(
            namespace,
            "13123123123",
            show_app_runtime=lambda **kwargs: calls.append(kwargs) or "shown",
        )

        self.assertEqual(result, "shown")
        forwarded = calls[0]
        self.assertEqual(forwarded["parent"], "root")
        self.assertEqual(forwarded["caller_num"], "13123123123")
        self.assertEqual(forwarded["get_popup"](), "popup")
        forwarded["set_popup"]("next_popup")
        self.assertEqual(namespace["current_call_popup"], "next_popup")
        self.assertEqual(forwarded["get_serial"](), "serial")
        forwarded["set_ring_timeout"](-1.0)
        self.assertEqual(namespace["ring_timeout_target"], -1.0)
        self.assertEqual(forwarded["log_error"]("show log"), ("log", "show log"))
        self.assertEqual(forwarded["mark_call_handled"](), "handled")
        self.assertTrue(forwarded["is_enabled"]())
        namespace["CALL_POPUP_ENABLED"] = False
        self.assertFalse(forwarded["is_enabled"]())

    def test_missed_popup_namespace_runtime_forwards_singleton_state(self):
        namespace = self.base_namespace()
        calls = []
        missed_call = object()

        result = show_missed_call_popup_namespace_runtime(
            namespace,
            missed_call,
            show_app_runtime=lambda **kwargs: calls.append(kwargs) or "shown",
        )

        self.assertEqual(result, "shown")
        forwarded = calls[0]
        self.assertIs(forwarded["missed_call"], missed_call)
        self.assertEqual(forwarded["get_popup"](), "missed_popup")
        forwarded["set_popup"]("next")
        self.assertEqual(namespace["current_missed_call_popup"], "next")
        self.assertEqual(forwarded["show_window"], "show_window")
        self.assertTrue(forwarded["is_enabled"]())

    def test_call_session_namespace_runtime_forwards_tracker_actions(self):
        namespace = self.base_namespace()

        self.assertEqual(
            start_incoming_call_session_namespace_runtime(namespace, "10086"),
            ("start", "10086"),
        )
        self.assertEqual(mark_incoming_call_handled_namespace_runtime(namespace), "handled")
        self.assertEqual(finish_incoming_call_session_namespace_runtime(namespace), "finished")
        self.assertEqual(reset_incoming_call_session_namespace_runtime(namespace), "reset")
        self.assertIsNone(set_missed_call_popup_namespace_runtime(namespace, None))
        self.assertIsNone(namespace["current_missed_call_popup"])

    def test_phone_popup_setting_disables_incoming_and_missed_windows(self):
        namespace = self.base_namespace()
        namespace["CALL_POPUP_ENABLED"] = False

        self.assertEqual(
            show_call_popup_namespace_runtime(
                namespace,
                "10086",
                show_app_runtime=lambda **_kwargs: self.fail("incoming popup opened"),
            ),
            "disabled",
        )
        self.assertEqual(
            show_missed_call_popup_namespace_runtime(
                namespace,
                object(),
                show_app_runtime=lambda **_kwargs: self.fail("missed popup opened"),
            ),
            "disabled",
        )

    def test_serial_call_state_namespace_runtime_reads_and_writes_state(self):
        namespace = self.base_namespace()

        self.assertEqual(get_serial_call_state_namespace_runtime(namespace), (3.0, "10086"))

        result = set_serial_call_state_namespace_runtime(namespace, 0.0, "")

        self.assertIsNone(result)
        self.assertEqual(namespace["ring_timeout_target"], 0.0)
        self.assertEqual(namespace["current_dial_num"], "")


if __name__ == "__main__":
    unittest.main()
