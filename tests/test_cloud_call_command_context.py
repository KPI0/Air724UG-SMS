import asyncio
import json
import threading
import unittest
from contextlib import ExitStack
from dataclasses import replace
from types import SimpleNamespace
from unittest.mock import patch

from sms_core import serial_sender
from sms_app.cloud_namespace_bindings import install_cloud_namespace_bindings
from sms_core.call_command_context import CallCommandContext, observe_call_command_line
from sms_core.cloud_command_context import set_cloud_command_authorized
from sms_core.cloud_payloads import build_register_payload, build_channel_status_payload, build_status_payload
from sms_core.cloud_message_namespace_runtime import (
    run_registered_cloud_serial_command_namespace_runtime,
    send_cloud_serial_command_namespace_runtime,
)
from sms_core.threading_runtime import WorkerThreadRegistry
from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.serial_app_runtime import build_serial_runtime_callbacks
from sms_core.serial_runtime import SerialRuntimeState, handle_serial_runtime_line
from sms_ui import call_popup_namespace_runtime as popup_runtime
from tests import test_peer_call_sync
from tests.test_serial_transaction_boundaries import FinalWriteGate, SyntheticModem
from tests.test_serial_runtime import runtime_callbacks, runtime_config
from tests.test_serial_app_runtime import app_callbacks


def incoming_meta(session_id):
    return {"call_control_guard": 1, "call_direction": "incoming",
            "call_session_id": session_id, "call_phone": "10086"}


def outgoing_meta(dial_id):
    return {"call_control_guard": 1, "call_direction": "outgoing", "call_dial_id": dial_id}


class CloudCallCommandContextTests(unittest.TestCase):
    def setUp(self):
        self.peer = test_peer_call_sync.PeerCallSyncTests()
        self.peer.setUp()
        self.namespace = self.peer.namespace
        self.at = serial_sender.AtCommandResponseCoordinator()
        self.modem = SyntheticModem(self.at, serial_sender.SmsPduSendCoordinator())
        self.namespace.update({
            "serial_lock": threading.RLock(), "serial_obj": self.modem,
            "serial_connection_generation": 1, "_cloud_runtime_imei": lambda: "000000000000001",
            "cloud_ws_conn": object(), "cloud_device_authorized": True, "cloud_connected": True,
            "CLOUD_CONTROL_ENABLED": True, "SERIAL_COMMAND_RESPONSE_COORDINATOR": self.at,
            "CLOUD_SENSITIVE_COMMAND_PERMISSIONS": True,
            "SERIAL_COMMAND_THREAD_REGISTRY": WorkerThreadRegistry(),
            "_cloud_log": lambda *args, **kwargs: None, "current_dial_num": "",
            "mark_call_connected": lambda *args, **kwargs:
                popup_runtime.mark_call_popup_connected_namespace_runtime(self.namespace, *args, **kwargs),
        })

    def send(self, command, metadata=None):
        return send_cloud_serial_command_namespace_runtime(self.namespace, command, metadata)

    def redial(self):
        self.peer.finish_local()
        self.peer.start("desktop-new")

    def paused(self, command, metadata, change):
        gate = FinalWriteGate()
        results = []
        with patch.object(serial_sender, "DEFAULT_SERIAL_WRITE_LOCK", gate):
            worker = threading.Thread(target=lambda: results.append(self.send(command, metadata)))
            worker.start()
            try:
                self.assertTrue(gate.reached.wait(3))
                change()
            finally:
                gate.resume.set()
                worker.join(3)
        self.assertFalse(worker.is_alive())
        self.assertEqual(len(results), 1)
        self.assertIsNone(self.at._active)
        return results[0]

    def test_old_incoming_command_cannot_cross_final_write_wait(self):
        for command in ("ATA", "ATH"):
            with self.subTest(command=command):
                self.setUp()
                self.peer.start("desktop-old")
                result = self.paused(command, incoming_meta("desktop-old"), self.redial)
                self.assertFalse(result[0])
                self.assertEqual(self.modem.writes, [])
                self.assertFalse(self.peer.tracker.snapshot().handled)
                self.assertEqual(self.namespace["ring_timeout_target"], 123.0)

    def test_unknown_numbered_call_never_falls_back_to_same_phone(self):
        self.peer.start("desktop-current")
        self.assertFalse(self.send("ATA", incoming_meta("unknown"))[0])
        self.assertEqual(self.modem.writes, [])
        self.assertFalse(self.peer.tracker.snapshot().handled)

    def test_missing_guard_identity_is_not_raw_at(self):
        self.peer.start("desktop-current")
        for metadata in ({"call_control_guard": 1}, incoming_meta(""), outgoing_meta("")):
            with self.subTest(metadata=metadata):
                self.assertFalse(self.send("ATH", metadata)[0])
        self.assertEqual(self.modem.writes, [])

    def test_old_outgoing_hangup_cannot_reach_new_incoming(self):
        self.assertTrue(self.send("ATD10086;", outgoing_meta("dial-old"))[0])
        self.modem.writes.clear()
        result = self.paused("ATH", outgoing_meta("dial-old"), lambda: self.peer.start("desktop-new"))
        self.assertFalse(result[0])
        self.assertEqual(self.modem.writes, [])

    def test_queued_outgoing_dial_cannot_replace_new_incoming(self):
        result = self.paused("ATD10086;", outgoing_meta("dial-old"), lambda: self.peer.start("desktop-new"))
        self.assertFalse(result[0])
        self.assertEqual(self.modem.writes, [])

    def test_current_incoming_and_outgoing_controls_still_write(self):
        self.peer.start("desktop-current")
        self.assertTrue(self.send("ATA", incoming_meta("desktop-current"))[0])
        self.assertTrue(self.peer.tracker.snapshot().handled)
        self.assertTrue(self.send("ATH", incoming_meta("desktop-current"))[0])
        self.peer.finish_local()
        self.assertTrue(self.send("ATD10086;", outgoing_meta("dial-current"))[0])
        self.assertTrue(self.send("ATH", outgoing_meta("dial-current"))[0])
        self.assertEqual(self.modem.writes, [b"ATA\r\n", b"ATH\r\n", b"ATD10086;\r\n", b"ATH\r\n"])

    def test_raw_at_keeps_current_behavior_and_numeric_dial_replaces_old_identity(self):
        self.assertTrue(self.send("ATA")[0])
        self.assertTrue(self.send("ATH")[0])
        self.assertTrue(self.send("ATD10086;", outgoing_meta("dial-old"))[0])
        self.assertTrue(self.send("ATD10086;")[0])
        self.modem.writes.clear()
        self.assertFalse(self.send("ATH", outgoing_meta("dial-old"))[0])
        self.assertEqual(self.modem.writes, [])

    def test_raw_queued_answer_is_also_bound_to_original_lifetime(self):
        self.peer.start("desktop-old")
        result = self.paused("ATA", None, self.redial)
        self.assertFalse(result[0])
        self.assertEqual(self.modem.writes, [])

    def test_connected_ui_callback_does_not_mark_same_number_redial(self):
        self.peer.start("desktop-old")
        queued = []
        self.namespace["run_on_ui_thread"] = lambda callback, _post: queued.append(callback)
        self.assertTrue(self.send("ATA", incoming_meta("desktop-old"))[0])
        self.assertEqual(len(queued), 1)
        self.namespace["run_on_ui_thread"] = lambda callback, _post: callback()
        self.redial()
        queued[0]()
        self.assertFalse(self.peer.tracker.snapshot().handled)
        self.assertEqual(self.namespace["ring_timeout_target"], 123.0)

    def test_hidden_popup_keeps_current_answer_valid(self):
        self.peer.start("desktop-current")
        popup_runtime.close_call_popup_namespace_runtime(self.namespace)
        self.assertTrue(self.send("ATA", incoming_meta("desktop-current"))[0])
        self.assertTrue(self.peer.tracker.snapshot().handled)

    def test_serial_terminals_and_ring_revoke_outgoing_before_ui_effects(self):
        for line in ("NO CARRIER", "BUSY", "NO ANSWER", '+CIEV: "CALL",0',
                     "RING", '+CLIP: "10086",129'):
            with self.subTest(line=line):
                self.setUp()
                self.assertTrue(self.send("ATD10086;", outgoing_meta("dial-old"))[0])
                self.modem.writes.clear()
                observe_call_command_line(self.namespace, line)
                self.assertFalse(self.send("ATH", outgoing_meta("dial-old"))[0])
                self.assertEqual(self.modem.writes, [])

    def test_unrelated_at_errors_do_not_revoke_active_outgoing_control(self):
        self.assertTrue(self.send("ATD10086;", outgoing_meta("dial-current"))[0])
        for line in ("+CME ERROR: 3", "+CMS ERROR: 500", "ERROR"):
            observe_call_command_line(self.namespace, line)
        self.modem.writes.clear()
        self.assertTrue(self.send("ATH", outgoing_meta("dial-current"))[0])
        self.assertEqual(self.modem.writes, [b"ATH\r\n"])

    def test_serial_reader_retires_call_before_response_and_debug_callbacks(self):
        self.assertTrue(self.send("ATD10086;", outgoing_meta("dial-old"))[0])
        self.modem.writes.clear()
        state = SerialRuntimeState.create(parse_sms_callback_head)
        state.call_state.current_dial_num = "10086"
        observed = []

        def assert_retired(line):
            observed.append(line)
            self.assertFalse(self.send("ATH", outgoing_meta("dial-old"))[0])

        callback_map = app_callbacks([])
        callback_map["observe_call_command_line"] = lambda line: observe_call_command_line(self.namespace, line)
        wired = build_serial_runtime_callbacks(callback_map)
        callbacks = replace(runtime_callbacks([]),
            observe_call_command_line=wired.observe_call_command_line,
            observe_sms_send_line=assert_retired, push_serial_debug=assert_retired)
        handle_serial_runtime_line(state, "NO CARRIER", 100.0, "TEST", False,
                                   runtime_config(), callbacks, {})
        self.assertTrue(observed)
        self.assertEqual(self.modem.writes, [])

    def test_repeated_ring_and_clip_do_not_revoke_current_incoming(self):
        self.peer.start("desktop-current")
        observe_call_command_line(self.namespace, "RING")
        observe_call_command_line(self.namespace, '+CLIP: "10086",129')
        self.assertTrue(self.send("ATA", incoming_meta("desktop-current"))[0])

    def test_non_voice_dial_does_not_create_outgoing_context(self):
        self.assertTrue(self.send("ATD*100#;")[0])
        self.assertIsNone(self.namespace.get("_cloud_dial_lifetime"))
        self.assertEqual(self.namespace["current_dial_num"], "")

    def test_capability_is_advertised_on_login_and_both_status_messages(self):
        payloads = (
            build_register_payload(True, 1, {}, "synthetic", "TEST", 115200, "Manual"),
            build_channel_status_payload(1, {}, True, True, 1),
            build_status_payload(1, {}, True, True, "TEST", 115200, "Manual"),
        )
        for payload in payloads:
            self.assertIs(type(payload.get("call_control_guard")), int)
            self.assertEqual(payload["call_control_guard"], 1)

    def test_incomplete_and_wrongly_typed_metadata_never_falls_back_to_raw(self):
        self.peer.start("desktop-current")
        for metadata in (
            {"call_session_id": "desktop-current"}, {"call_phone": "10086"},
            {**incoming_meta("desktop-current"), "call_control_guard": True},
            {**incoming_meta("desktop-current"), "call_control_guard": "1"},
            {**incoming_meta("desktop-current"), "call_direction": "outgoing"},
            {**incoming_meta("desktop-current"), "call_dial_id": "foreign"},
        ):
            with self.subTest(metadata=metadata):
                self.assertFalse(self.send("ATA", metadata)[0])
        self.assertEqual(self.modem.writes, [])

    def test_old_answer_response_after_redial_does_not_mark_new_call(self):
        self.peer.start("desktop-old")
        self.assert_delayed_result_preserves_new_call("ATA", incoming_meta("desktop-old"))

    def test_failed_dial_response_does_not_clear_new_call_state(self):
        self.modem.command_reply = "ERROR"
        self.assert_delayed_result_preserves_new_call("ATD10086;", outgoing_meta("dial-old"))

    def assert_delayed_result_preserves_new_call(self, command, metadata):
        reached, resume = threading.Event(), threading.Event()
        original_wait = serial_sender.AtCommandResponseWaiter.wait
        results = []

        def delayed_wait(waiter, timeout):
            reached.set()
            if not resume.wait(3):
                raise RuntimeError("Missing response release")
            return original_wait(waiter, timeout)

        with patch.object(serial_sender.AtCommandResponseWaiter, "wait", delayed_wait):
            worker = threading.Thread(target=lambda: results.append(self.send(command, metadata)))
            worker.start()
            try:
                self.assertTrue(reached.wait(3))
                if command == "ATA":
                    self.peer.finish_local()
                else:
                    observe_call_command_line(self.namespace, "RING")
                popup_runtime.start_serial_incoming_call_namespace_runtime(
                    self.namespace, "10086", "desktop-new", 456.0)
            finally:
                resume.set()
                worker.join(3)
        self.assertFalse(worker.is_alive())
        self.assertFalse(results[0][0])
        self.assertFalse(self.peer.tracker.snapshot().handled)
        self.assertEqual(self.namespace["ring_timeout_target"], 456.0)
        self.assertEqual(self.namespace["current_dial_num"], "")

    def test_lifetime_lock_remains_held_through_actual_serial_write(self):
        self.peer.start("desktop-current")
        write_started, resume_write, transition_started, transition_finished = (
            threading.Event() for _ in range(4))
        original_write = self.modem.write
        results = []

        def blocking_write(payload):
            write_started.set()
            if not resume_write.wait(3):
                raise RuntimeError("Missing serial release")
            return original_write(payload)

        def transition():
            transition_started.set()
            self.peer.finish_local()
            transition_finished.set()

        self.modem.write = blocking_write
        sender = threading.Thread(target=lambda: results.append(self.send("ATH", incoming_meta("desktop-current"))))
        changer = threading.Thread(target=transition)
        sender.start()
        try:
            self.assertTrue(write_started.wait(3))
            changer.start()
            self.assertTrue(transition_started.wait(3))
            self.assertFalse(transition_finished.wait(0.03))
        finally:
            resume_write.set()
            sender.join(3)
            changer.join(3)
        self.assertFalse(sender.is_alive())
        self.assertFalse(changer.is_alive())
        self.assertEqual(self.modem.writes, [b"ATH\r\n"])

    def test_dial_and_its_hangup_can_wait_behind_the_same_at_transaction(self):
        self.assert_pending_dial_pair("unchanged")

    def test_pending_dial_and_hangup_are_both_cancelled_by_new_incoming(self):
        self.assert_pending_dial_pair("incoming")

    def test_pending_dial_is_released_when_authorization_changes(self):
        self.assert_pending_dial_pair("authorization")

    def test_pending_dial_is_released_after_executor_exception(self):
        def fail(*args, **kwargs):
            raise RuntimeError("synthetic executor error")
        with self.assertRaises(RuntimeError):
            send_cloud_serial_command_namespace_runtime(
                self.namespace, "ATD10086;", outgoing_meta("failed-dial"), send_runtime=fail)
        self.assertIsNone(self.namespace.get("_cloud_pending_dial"))
        self.assertTrue(self.send("ATD10086;", outgoing_meta("next-dial"))[0])

    def test_permission_rejection_releases_pending_dial(self):
        self.namespace["CLOUD_SENSITIVE_COMMAND_PERMISSIONS"] = False
        self.assertFalse(self.send("ATD10086;", outgoing_meta("denied-dial"))[0])
        self.assertIsNone(self.namespace.get("_cloud_pending_dial"))
        self.assertEqual(self.modem.writes, [])
        self.namespace["CLOUD_SENSITIVE_COMMAND_PERMISSIONS"] = True
        self.assertTrue(self.send("ATD10086;", outgoing_meta("next-dial"))[0])

    def test_authorization_regrant_cannot_revive_queued_call_commands(self):
        for command, direction in (("ATA", "incoming"), ("ATH", "incoming"),
                                   ("ATD10086;", "outgoing"), ("ATH", "outgoing")):
            for revoke in (False, True):
                with self.subTest(command=command, direction=direction, revoke=revoke):
                    self.setUp()
                    metadata = incoming_meta("desktop-current") if direction == "incoming" else outgoing_meta("dial-current")
                    if direction == "incoming":
                        self.peer.start("desktop-current")
                    elif command == "ATH":
                        self.assertTrue(self.send("ATD10086;", metadata)[0])
                        self.modem.writes.clear()
                    original_token = self.peer.tracker.snapshot().session_token
                    original_dial = self.namespace.get("_cloud_dial_lifetime")

                    def update_authorization():
                        if revoke:
                            set_cloud_command_authorized(self.namespace, False)
                        set_cloud_command_authorized(self.namespace, True)

                    result = self.paused(command, metadata, update_authorization)
                    self.assertEqual(result[0], not revoke, result)
                    self.assertEqual(self.modem.writes, [] if revoke else [(command + "\r\n").encode("ascii")])
                    self.assertIs(self.peer.tracker.snapshot().session_token, original_token)
                    if command != "ATD10086;":
                        self.assertIs(self.namespace.get("_cloud_dial_lifetime"), original_dial)
                    if revoke:
                        self.assertIsNone(self.namespace.get("_cloud_pending_dial"))
                        self.assertTrue(self.send(command, metadata)[0])

    def assert_pending_dial_pair(self, change):
        lock = serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK
        dial_queued, hangup_queued = threading.Event(), threading.Event()
        original_wait = CallCommandContext.wait_for_dial
        results = {}

        class TransactionGate:
            def __enter__(self):
                if threading.current_thread().name == "test-pending-dial":
                    dial_queued.set()
                return lock.__enter__()

            def __exit__(self, *args):
                return lock.__exit__(*args)

        def wait_for_dial(context, timeout):
            if context.pending_dial is not None and not context.dial_number:
                hangup_queued.set()
            return original_wait(context, timeout)

        def send(name, command):
            results[name] = self.send(command, outgoing_meta("dial-pending"))

        dial = threading.Thread(target=lambda: send("dial", "ATD10086;"), name="test-pending-dial")
        hangup = threading.Thread(target=lambda: send("hangup", "ATH"))
        lock.acquire()
        with patch.object(serial_sender, "DEFAULT_SERIAL_TRANSACTION_LOCK", TransactionGate()), \
                patch.object(CallCommandContext, "wait_for_dial", wait_for_dial):
            dial.start()
            try:
                self.assertTrue(dial_queued.wait(3))
                self.assertEqual(self.modem.writes, [])
                self.assertEqual(self.namespace["current_dial_num"], "")
                self.assertIsNone(self.namespace.get("_cloud_dial_lifetime"))
                hangup.start()
                self.assertTrue(hangup_queued.wait(3))
                if change == "incoming":
                    self.peer.start("desktop-new")
                elif change == "authorization":
                    self.namespace["cloud_device_authorized"] = False
            finally:
                lock.release()
                dial.join(3)
                if hangup.ident is not None:
                    hangup.join(3)
        self.assertFalse(dial.is_alive())
        self.assertFalse(hangup.is_alive())
        self.assertIsNone(self.namespace.get("_cloud_pending_dial"))
        if change == "unchanged":
            self.assertTrue(results["dial"][0], results)
            self.assertTrue(results["hangup"][0], results)
            self.assertEqual(self.modem.writes, [b"ATD10086;\r\n", b"ATH\r\n"])
        else:
            self.assertFalse(results["dial"][0], results)
            self.assertFalse(results["hangup"][0], results)
            self.assertEqual(self.modem.writes, [])


class CloudCallReceiveQueueTests(unittest.IsolatedAsyncioTestCase):
    async def exercise(self, command, direction, pause, *, payload=None, replace_call=True,
                       authorization_change=None):
        fixture = CloudCallCommandContextTests()
        fixture.setUp()
        ns, peer, modem, at = fixture.namespace, fixture.peer, fixture.modem, fixture.at
        if direction == "incoming":
            peer.start("desktop-old")
            metadata = payload if payload is not None else incoming_meta("desktop-old")
        else:
            metadata = payload if payload is not None else outgoing_meta("dial-old")
            if command == "ATH":
                self.assertTrue(fixture.send("ATD10086;", metadata)[0])
                modem.writes.clear()
        loop = asyncio.get_running_loop()
        reached = asyncio.Event()
        release_async = asyncio.Event()
        release_worker = threading.Event()
        workers, replies = [], []
        loop_thread = threading.current_thread()
        ns.update({
            "cloud_imei_verified": True, "CLOUD_DEVICE_IMEI": "000000000000001",
            "CLOUD_DEVICE_SECRET": "synthetic-only", "CLOUD_WS_URL": "wss://example.invalid/ws/device",
            "SERIAL_COMMAND_THREAD_REGISTRY": WorkerThreadRegistry(),
            "show_window": lambda: None, "hide_window": lambda: None,
            "set_cloud_auth_status_from_ack": lambda _data: None,
            "set_cloud_status": lambda *args: None, "_normalize_imei": lambda value: value,
        })
        install_cloud_namespace_bindings(ns)

        async def reply(_ws, payload):
            replies.append(payload)
            if pause == "started" and payload.get("status") == "started":
                reached.set()
                await release_async.wait()

        async def replay(_ws, _data, mark_seen=True):
            if pause == "replay" and not mark_seen:
                reached.set()
                await release_async.wait()
            return True

        def wait_worker():
            loop.call_soon_threadsafe(reached.set)
            if not release_worker.wait(4):
                raise RuntimeError("Missing queue release")

        def thread_factory(*args, target, **kwargs):
            def run():
                if pause == "worker":
                    wait_worker()
                target()
            thread = threading.Thread(*args, target=run, **kwargs)
            workers.append(thread)
            return thread

        class QueueLock:
            def __init__(self, lock, condition=lambda: True):
                self.lock, self.condition = lock, condition
                self.paused = False

            def __enter__(self):
                if (not self.paused and threading.current_thread() is not loop_thread
                        and self.condition()):
                    self.paused = True
                    wait_worker()
                return self.lock.__enter__()

            def __exit__(self, *args):
                return self.lock.__exit__(*args)

        ns.update({
            "_cloud_reply": reply, "_cloud_check_replay_window": replay,
            "_cloud_log": lambda *args, **kwargs: None,
            "_cloud_auth_matches": lambda data: data.get("secret") == "synthetic-only",
            "_cloud_send_status_payload": lambda: {},
            "_handle_cloud_call_recording_message": lambda *args, **kwargs: False,
            "threading": SimpleNamespace(Thread=thread_factory),
        })
        with ExitStack() as patches:
            if pause == "transaction":
                patches.enter_context(patch.object(serial_sender, "DEFAULT_SERIAL_TRANSACTION_LOCK",
                                                   QueueLock(serial_sender.DEFAULT_SERIAL_TRANSACTION_LOCK)))
            elif pause == "serial":
                ns["serial_lock"] = QueueLock(ns["serial_lock"], lambda: at._active is not None)
            elif pause == "write":
                patches.enter_context(patch.object(serial_sender, "DEFAULT_SERIAL_WRITE_LOCK",
                                                   QueueLock(serial_sender.DEFAULT_SERIAL_WRITE_LOCK)))
            request = payload if payload is not None else {
                "type": "cmd", "command": command, "task_id": "synthetic-call-task",
                "target_imei": ns["CLOUD_DEVICE_IMEI"], "secret": "synthetic-only", **metadata,
            }
            handling = asyncio.create_task(ns["_handle_cloud_message"](ns["cloud_ws_conn"], json.dumps(request)))
            try:
                await asyncio.wait_for(reached.wait(), 3)
                if authorization_change is not None:
                    original_token = peer.tracker.snapshot().session_token
                    if authorization_change == "reauthorize":
                        await ns["_handle_cloud_message"](ns["cloud_ws_conn"], json.dumps({
                            "type": "device_login_ack", "imei": ns["CLOUD_DEVICE_IMEI"],
                            "ok": False, "auth_status": "failed",
                        }))
                    await ns["_handle_cloud_message"](ns["cloud_ws_conn"], json.dumps({
                        "type": "device_login_ack", "imei": ns["CLOUD_DEVICE_IMEI"],
                        "ok": True, "auth_status": "authorized",
                    }))
                    self.assertIs(peer.tracker.snapshot().session_token, original_token)
                elif replace_call:
                    if direction == "incoming":
                        peer.finish_local()
                    else:
                        observe_call_command_line(ns, "RING")
                    peer.start("desktop-new")
            finally:
                release_async.set()
                release_worker.set()
                await asyncio.wait_for(handling, 5)
                for worker in workers:
                    await asyncio.to_thread(worker.join, 2)
                    self.assertFalse(worker.is_alive())
        result = [payload for payload in replies if payload.get("type") == "send_at_result"]
        self.assertEqual(len(result), 1)
        if replace_call or authorization_change == "reauthorize":
            self.assertFalse(result[0]["ok"], result[0])
            self.assertEqual(modem.writes, [])
            self.assertFalse(peer.tracker.snapshot().handled)
            if direction == "incoming" or authorization_change is None:
                self.assertEqual(ns["ring_timeout_target"], 123.0)
        else:
            self.assertTrue(result[0]["ok"], result[0])
            self.assertEqual(modem.writes, [(command + "\r\n").encode("ascii")])
        self.assertIsNone(at._active)
        self.assertEqual(ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(), ())

    async def test_pending_dial_is_released_if_worker_registration_or_execution_fails(self):
        for stage in ("registration", "execution"):
            with self.subTest(stage=stage):
                fixture = CloudCallCommandContextTests()
                fixture.setUp()

                def fail(*args, **kwargs):
                    raise RuntimeError("synthetic worker failure")

                fixture.namespace["_cloud_send_serial_command"] = fail
                starter = fail if stage == "registration" else lambda _name, target, **kwargs: target()
                result = await run_registered_cloud_serial_command_namespace_runtime(
                    fixture.namespace, asyncio.get_running_loop(), "ATD10086;", outgoing_meta("failed-dial"),
                    start_worker=starter)
                self.assertFalse(result[0])
                self.assertIsNone(fixture.namespace.get("_cloud_pending_dial"))
                self.assertTrue(fixture.send("ATD10086;", outgoing_meta("next-dial"))[0])

    async def test_actual_login_ack_regrant_cancels_old_commands_but_repeated_authorized_ack_does_not(self):
        for command, direction in (("ATA", "incoming"), ("ATH", "incoming"),
                                   ("ATD10086;", "outgoing"), ("ATH", "outgoing")):
            for pause in ("replay", "started", "worker", "transaction", "serial", "write"):
                for change in ("reauthorize", "repeat_authorized"):
                    with self.subTest(command=command, direction=direction, pause=pause, change=change):
                        await self.exercise(command, direction, pause, replace_call=False,
                                            authorization_change=change)

    async def test_incoming_answer_and_hangup_across_all_admission_and_write_queues(self):
        for command in ("ATA", "ATH"):
            for pause in ("replay", "started", "worker", "transaction", "serial", "write"):
                with self.subTest(command=command, pause=pause):
                    await self.exercise(command, "incoming", pause)

    async def test_outgoing_dial_and_hangup_across_all_admission_and_write_queues(self):
        for command in ("ATD10086;", "ATH"):
            for pause in ("replay", "started", "worker", "transaction", "serial", "write"):
                with self.subTest(command=command, pause=pause):
                    await self.exercise(command, "outgoing", pause)


if __name__ == "__main__":
    unittest.main()
