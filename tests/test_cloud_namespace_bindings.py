import asyncio
import json
import queue
import threading
import unittest
from dataclasses import replace
from functools import partial
from types import SimpleNamespace
from unittest.mock import patch

import sms_app.cloud_namespace_bindings as bindings
from sms_core.call_session import IncomingCallSessionTracker
from sms_core.call_command_context import call_state_lock
from sms_core.cloud_payloads import build_call_event_payload, build_serial_log_payload
from sms_core.cloud_command_context import set_cloud_command_authorized
from sms_core.cloud_message_namespace_runtime import send_cloud_serial_command_namespace_runtime
from sms_core.serial_sender import AtCommandResponseCoordinator, SerialCommandResult
from sms_core.cloud_serial_log_runtime import CloudSerialLogDrainState
from sms_core.cloud_protocol import parse_sms_callback_head
from sms_core.cloud_sms_event_runtime import (
    CloudSmsEventDrainState, handle_cloud_sms_event_ack, interrupt_cloud_sms_event_drain,
)
from sms_ui.call_popup_namespace_runtime import (
    finish_incoming_call_session_namespace_runtime,
    mark_call_popup_connected_namespace_runtime,
    set_call_popup_namespace_runtime,
    show_call_popup_namespace_runtime,
    start_serial_incoming_call_namespace_runtime,
)
from sms_ui.call_popup_runtime import show_call_popup_app_runtime, show_call_popup_runtime
from sms_core.serial_runtime import SerialRuntimeState, handle_serial_runtime_line
from tests.test_serial_runtime import runtime_callbacks, runtime_config


class CloudNamespaceBindingsTests(unittest.TestCase):
    def make_namespace(self):
        calls = []
        return {
            "calls": calls,
            "cloud_ws_loop": "loop",
            "cloud_ws_conn": "ws",
            "cloud_connected": True,
            "CLOUD_AUTO_UPLOAD": True,
            "_cloud_identity_payload": lambda: {"imei": "861"},
            "_cloud_log": lambda *args, **kwargs: calls.append(("log", args, kwargs)),
        }

    def test_install_cloud_namespace_bindings_registers_expected_names(self):
        namespace = self.make_namespace()

        result = bindings.install_cloud_namespace_bindings(namespace)

        self.assertIs(result, namespace)
        for name in (
            "_cloud_runtime_imei",
            "_cloud_send_register",
            "_cloud_send_session_revoke",
            "_cloud_send_unregister",
            "_cloud_schedule_unregister",
            "_cloud_unregister_then_close",
            "save_cloud_control_setting",
            "_cloud_send_payload",
            "_cloud_send_call_event",
            "_cloud_send_call_state",
            "_cloud_reply",
            "_handle_cloud_message",
            "start_cloud_control",
            "stop_cloud_control",
            "restart_cloud_control",
            "refresh_cloud_control_settings_from_config",
            "open_cloud_control_window",
        ):
            self.assertIn(name, namespace)
            self.assertTrue(callable(namespace[name]))

    def test_sync_bindings_forward_namespace_and_arguments(self):
        namespace = self.make_namespace()
        bindings.install_cloud_namespace_bindings(namespace)

        with patch.object(bindings, "cloud_runtime_imei_namespace_runtime", return_value="861") as runtime_imei, \
                patch.object(bindings, "save_cloud_control_setting_namespace_runtime", return_value="saved") as save_setting, \
                patch.object(bindings, "cloud_log_namespace_runtime", return_value="logged") as cloud_log, \
                patch.object(bindings, "send_cloud_serial_log_namespace_runtime", return_value="serial_log") as serial_log, \
                patch.object(bindings, "refresh_cloud_control_settings_namespace_runtime", return_value=True) as refresh_settings, \
                patch.object(bindings, "start_cloud_control_namespace_runtime", return_value=True) as start_control:
            self.assertEqual(namespace["_cloud_runtime_imei"](), "861")
            self.assertEqual(namespace["save_cloud_control_setting"](enabled=True, url="ws://host"), "saved")
            self.assertEqual(namespace["_cloud_log"]("hello", show_main=True), "logged")
            self.assertEqual(namespace["_cloud_send_serial_log"]("line"), "serial_log")
            self.assertTrue(namespace["refresh_cloud_control_settings_from_config"]())
            self.assertTrue(namespace["start_cloud_control"](show_errors=True))

        runtime_imei.assert_called_once_with(namespace)
        self.assertEqual(save_setting.call_args.args, (namespace,))
        self.assertTrue(save_setting.call_args.kwargs["enabled"])
        self.assertEqual(save_setting.call_args.kwargs["url"], "ws://host")
        cloud_log.assert_called_once_with(namespace, "hello", show_main=True)
        serial_log.assert_called_once_with(namespace, "line")
        refresh_settings.assert_called_once_with(namespace)
        start_control.assert_called_once_with(namespace, show_errors=True)

    def test_schedule_and_close_bindings_forward_connection_context(self):
        namespace = self.make_namespace()
        bindings.install_cloud_namespace_bindings(namespace)

        with patch.object(bindings, "schedule_cloud_unregister_runtime", return_value="scheduled") as schedule_runtime:
            self.assertEqual(namespace["_cloud_schedule_unregister"]("hidden"), "scheduled")

        kwargs = schedule_runtime.call_args.kwargs
        self.assertEqual(kwargs["reason"], "hidden")
        self.assertEqual(kwargs["get_loop"](), "loop")
        self.assertEqual(kwargs["get_ws"](), "ws")
        self.assertTrue(kwargs["is_connected"]())
        self.assertIs(kwargs["send_unregister"], namespace["_cloud_send_unregister"])

        async def close_runtime(ws, **kwargs):
            return ("closed", ws, kwargs)

        with patch.object(bindings, "unregister_then_close_cloud_connection_runtime", side_effect=close_runtime):
            result = asyncio.run(namespace["_cloud_unregister_then_close"]("ws2", reason="disconnect"))

        self.assertEqual(result[0:2], ("closed", "ws2"))
        self.assertTrue(result[2]["auto_upload"])
        self.assertIs(result[2]["send_unregister"], namespace["_cloud_send_unregister"])
        self.assertIs(result[2]["send_session_revoke"], namespace["_cloud_send_session_revoke"])

    def test_async_bindings_forward_to_underlying_runtimes(self):
        namespace = self.make_namespace()
        bindings.install_cloud_namespace_bindings(namespace)

        async def wait_runtime(ns, ws, timeout):
            return ("ack", ns is namespace, ws, timeout)

        async def payload_runtime(ws, payload):
            return ("payload", ws, payload)

        async def revoke_runtime(ns, ws, reason):
            return ("revoke", ns is namespace, ws, reason)

        async def reply_runtime(ws, payload, **kwargs):
            return ("reply", ws, payload, kwargs["identity_payload"]())

        with patch.object(bindings, "wait_cloud_login_ack_namespace_runtime", side_effect=wait_runtime), \
                patch.object(bindings, "cloud_identity_payload_namespace_runtime", return_value={"imei": "861"}), \
                patch.object(bindings, "send_cloud_session_revoke_namespace_runtime", side_effect=revoke_runtime), \
                patch.object(bindings, "send_cloud_payload_runtime", side_effect=payload_runtime), \
                patch.object(bindings, "reply_cloud_payload_runtime", side_effect=reply_runtime):
            ack = asyncio.run(namespace["_cloud_wait_login_ack"]("ws", timeout=3.0))
            revoke = asyncio.run(
                namespace["_cloud_send_session_revoke"]("ws", reason="password_changed")
            )
            payload = asyncio.run(namespace["_cloud_send_payload"]("ws", {"type": "x"}))
            reply = asyncio.run(namespace["_cloud_reply"]("ws", {"ok": True}))

        self.assertEqual(ack, ("ack", True, "ws", 3.0))
        self.assertEqual(revoke, ("revoke", True, "ws", "password_changed"))
        self.assertEqual(payload, ("payload", "ws", {"type": "x"}))
        self.assertEqual(reply, ("reply", "ws", {"ok": True}, {"imei": "861"}))


class CloudCallHistoryDeliveryTests(unittest.IsolatedAsyncioTestCase):
    def setUp(self):
        self.now = 100.0
        self.logs = []
        self.namespace = {
            "CLOUD_CONTROL_ENABLED": True, "cloud_connected": True,
            "cloud_device_authorized": True, "cloud_stop_event": threading.Event(),
            "cloud_ws_loop": None, "cloud_ws_conn": None,
            "CLOUD_SMS_EVENT_Q": queue.Queue(maxsize=10),
            "CLOUD_SMS_EVENT_DRAIN_STATE": CloudSmsEventDrainState(),
            "CLOUD_SMS_EVENT_DRAIN_BATCH": 100,
            "INCOMING_CALL_SESSION": IncomingCallSessionTracker(),
            "current_dial_num": "", "ring_timeout_target": 0.0,
            "time": SimpleNamespace(monotonic=lambda: self.now),
            "log_file_only": self.logs.append,
            "_cloud_build_call_event_payload": build_call_event_payload,
        }
        bindings.install_cloud_namespace_bindings(self.namespace)
        self.namespace["_cloud_runtime_imei"] = lambda: "synthetic-device"
        self.namespace["_cloud_identity_payload"] = lambda: {"imei": "synthetic-device"}
        self.namespace["_cloud_now_ts"] = lambda: -1000000000
        self.writes = []
        self.ws = self.make_socket()
        self.namespace["cloud_ws_conn"] = self.ws

    def make_socket(self, *, auto_ack=True):
        namespace, writes = self.namespace, self.writes

        class Socket:
            async def send(self, raw):
                lock = namespace.get("_call_peer_state_lock")
                if lock is not None and hasattr(lock, "_is_owned"):
                    self_test.assertFalse(lock._is_owned())
                value = json.loads(raw)
                writes.append((self, value))
                if (self.auto_ack and value.get("ack_required") is True
                        and namespace["CLOUD_SMS_EVENT_DRAIN_STATE"].pending_ack is not None):
                    self_test.assertTrue(handle_cloud_sms_event_ack(self, {
                        "type": "device_event_ack", "ok": True,
                        "imei": value["imei"], "source_event_id": value["source_event_id"],
                    }, state=namespace["CLOUD_SMS_EVENT_DRAIN_STATE"]))

        self_test = self
        result = Socket()
        result.auto_ack = auto_ack
        return result

    def start_call(self, session_id="call-a"):
        self.assertTrue(start_serial_incoming_call_namespace_runtime(
            self.namespace, "synthetic-caller", session_id, self.now + 12.0,
        ))

    def payload(self, session_id="call-a"):
        return build_call_event_payload("synthetic-caller", "synthetic incoming call", -1000000000,
                                        {"imei": "synthetic-device"}, call_session_id=session_id)

    async def drain(self):
        return await self.namespace["_cloud_drain_sms_event_queue"](self.namespace["cloud_ws_conn"])

    async def wait_for_writes(self, count):
        async def wait():
            while len(self.writes) < count:
                await asyncio.sleep(0)
        await asyncio.wait_for(wait(), 1.0)

    async def test_offline_call_finished_before_same_number_redial_is_history_only(self):
        self.start_call()
        self.namespace["cloud_connected"] = False
        self.namespace["cloud_device_authorized"] = False
        self.assertEqual(self.namespace["_cloud_send_call_event"](
            "synthetic-caller", "synthetic incoming call", call_session_id="call-a",
        ), "queued_offline")
        original = self.namespace["CLOUD_SMS_EVENT_Q"].queue[0]
        before = dict(original)
        finish_incoming_call_session_namespace_runtime(self.namespace)
        self.start_call("call-b")
        current = self.payload("call-b")
        self.namespace["CLOUD_SMS_EVENT_Q"].put_nowait(current)
        self.namespace["cloud_connected"] = self.namespace["cloud_device_authorized"] = True
        await self.drain()
        self.assertEqual([value["history_only"] for _, value in self.writes], [True, False])
        self.assertEqual(original, before)
        self.assertNotIn("history_only", original)
        self.assertTrue(self.namespace["CLOUD_SMS_EVENT_Q"].empty())
        self.assertEqual(self.namespace["CLOUD_SMS_EVENT_Q"].unfinished_tasks, 0)

    async def test_same_ringing_call_survives_reconnect_without_event_age_guessing(self):
        self.start_call()
        payload = self.payload()
        payload["timestamp"] = 9999999999
        self.ws = self.make_socket()
        self.namespace["cloud_ws_conn"] = self.ws
        self.namespace["CLOUD_SMS_EVENT_Q"].put_nowait(payload)
        await self.drain()
        self.assertIs(self.writes[0][1]["history_only"], False)
        self.assertEqual(self.writes[0][1]["timestamp"], payload["timestamp"])
        self.assertNotIn("history_only", payload)

    async def test_real_serial_first_clip_commits_lifetime_before_cloud_delivery(self):
        self.namespace["cloud_ws_loop"] = asyncio.get_running_loop()
        state = SerialRuntimeState.create(parse_sms_callback_head)

        def send_event(caller, message, **kwargs):
            self.assertEqual(self.namespace["call_popup_local_session"][1], kwargs["call_session_id"])
            self.assertGreater(self.namespace["ring_timeout_target"], self.now)
            return self.namespace["_cloud_send_call_event"](caller, message, **kwargs)

        callbacks = replace(runtime_callbacks([]),
            call_state_guard=lambda: call_state_lock(self.namespace),
            start_serial_incoming_call=lambda caller, session_id, timeout:
                start_serial_incoming_call_namespace_runtime(self.namespace, caller, session_id, timeout),
            finish_incoming_call=lambda: finish_incoming_call_session_namespace_runtime(self.namespace),
            send_cloud_call_event=send_event)
        for line in ('RING', '+CLIP: "10086",129'):
            handle_serial_runtime_line(state, line, self.now, "synthetic-port", False,
                                       runtime_config(), callbacks, {})
            self.now += 1.0
        await self.wait_for_writes(1)
        self.assertIs(self.writes[0][1]["history_only"], False)
        self.assertEqual(self.writes[0][1]["call_session_id"], state.call_state.call_session_id)
        self.assertTrue(self.namespace["CLOUD_SMS_EVENT_Q"].empty())

    async def retry_unacknowledged(self, *, replace_socket, finish_call):
        self.start_call()
        payload = self.payload()
        before = dict(payload)
        self.namespace["CLOUD_SMS_EVENT_Q"].put_nowait(payload)
        self.ws.auto_ack = False
        task = asyncio.create_task(self.drain())
        try:
            await self.wait_for_writes(1)
            self.assertIs(self.writes[0][1]["history_only"], False)
            self.namespace["cloud_device_authorized"] = False
            if finish_call:
                finish_incoming_call_session_namespace_runtime(self.namespace)
            interrupt_cloud_sms_event_drain(self.namespace["CLOUD_SMS_EVENT_DRAIN_STATE"])
            await asyncio.gather(task, return_exceptions=True)
            self.assertIs(self.namespace["CLOUD_SMS_EVENT_Q"].queue[0], payload)
            old_ws = self.ws
            if replace_socket:
                self.ws = self.make_socket()
                self.namespace["cloud_ws_conn"] = self.ws
            self.ws.auto_ack = True
            self.namespace["cloud_device_authorized"] = True
            await self.drain()
            self.assertEqual(self.writes[1][1]["history_only"], finish_call)
            self.assertEqual(self.writes[0][1]["source_event_id"], self.writes[1][1]["source_event_id"])
            self.assertEqual(payload, before)
            self.assertEqual(self.namespace["CLOUD_SMS_EVENT_Q"].unfinished_tasks, 0)
            if replace_socket:
                self.assertIsNot(self.writes[1][0], old_ws)
        finally:
            task.cancel()
            await asyncio.gather(task, return_exceptions=True)

    async def test_unacknowledged_call_replayed_after_disconnect_and_finish_is_history_only(self):
        await self.retry_unacknowledged(replace_socket=True, finish_call=True)

    async def test_same_socket_reauthorization_can_notify_same_still_ringing_call(self):
        await self.retry_unacknowledged(replace_socket=False, finish_call=False)

    async def test_call_state_is_checked_when_send_coroutine_runs(self):
        self.start_call()
        payload = self.payload()
        send = self.namespace["_cloud_send_payload"](self.ws, payload)
        finish_incoming_call_session_namespace_runtime(self.namespace)
        self.assertEqual(await send, "sent")
        self.assertIs(self.writes[0][1]["history_only"], True)
        self.assertNotIn("history_only", payload)

    async def test_call_state_is_rechecked_after_peer_lock_wait_without_awaiting_under_lock(self):
        self.start_call()
        payload = self.payload()
        entered, acquired = threading.Event(), threading.Event()
        lock = threading.Lock()

        class ObservedLock:
            def __enter__(self):
                entered.set()
                lock.acquire()

            def __exit__(self, *_args):
                lock.release()

        def finish_while_locked():
            with lock:
                acquired.set()
                if entered.wait(1.0):
                    self.namespace["INCOMING_CALL_SESSION"].finish()
                    self.namespace["call_popup_local_session"] = None

        self.namespace["_call_peer_state_lock"] = ObservedLock()
        worker = threading.Thread(target=finish_while_locked)
        worker.start()
        self.assertTrue(await asyncio.to_thread(acquired.wait, 1.0))
        try:
            self.assertEqual(await self.namespace["_cloud_send_payload"](self.ws, payload), "sent")
        finally:
            await asyncio.to_thread(worker.join, 1.0)
        self.assertFalse(worker.is_alive())
        self.assertIs(self.writes[0][1]["history_only"], True)

    async def test_answered_expired_or_invalid_ringing_state_is_history_only(self):
        for target in (0.0, -1.0, self.now, True, "112", float("nan"), float("inf")):
            with self.subTest(target=target):
                self.start_call()
                self.namespace["ring_timeout_target"] = target
                await self.namespace["_cloud_send_payload"](self.ws, self.payload())
                self.assertIs(self.writes[-1][1]["history_only"], True)
                finish_incoming_call_session_namespace_runtime(self.namespace)
        self.start_call()
        self.namespace["INCOMING_CALL_SESSION"].mark_handled()
        self.namespace["INCOMING_CALL_SESSION"].confirm_handled(
            self.namespace["INCOMING_CALL_SESSION"].snapshot().session_token)
        await self.namespace["_cloud_send_payload"](self.ws, self.payload())
        self.assertIs(self.writes[-1][1]["history_only"], True)

    async def test_unconfirmed_lifetime_is_history_only_but_unrelated_payloads_are_unchanged(self):
        self.start_call()
        for state in (None, (object(), "call-a"), self.namespace["call_popup_local_session"]):
            with self.subTest(state=state):
                self.namespace["call_popup_local_session"] = state
                payload = self.payload("unknown-call")
                await self.namespace["_cloud_send_payload"](self.ws, payload)
                self.assertIs(self.writes[-1][1]["history_only"], True)
        for payload in ({"type": "sms_event", "message": "synthetic"},
                        {"type": "call_event", "phase": "outgoing"},
                        {"type": "call_state", "phase": "connected"}):
            await self.namespace["_cloud_send_payload"](self.ws, payload)
            self.assertEqual(self.writes[-1][1], payload)
            self.assertNotIn("history_only", self.writes[-1][1])

    async def test_unavailable_or_failed_tracker_keeps_history_delivery_and_safe_error(self):
        payload = self.payload()
        await self.namespace["_cloud_send_payload"](self.ws, payload)
        self.assertIs(self.writes[-1][1]["history_only"], True)

        def fail():
            raise RuntimeError("synthetic private diagnostic must not be logged")

        self.namespace["INCOMING_CALL_SESSION"] = SimpleNamespace(snapshot=fail)
        self.assertEqual(await self.namespace["_cloud_send_payload"](self.ws, payload), "sent")
        self.assertIs(self.writes[-1][1]["history_only"], True)
        self.assertTrue(any("RuntimeError" in message for message in self.logs))
        self.assertFalse(any("synthetic private diagnostic" in message for message in self.logs))

    def setup_log_queue(self):
        self.namespace.update({
            "cloud_ws_loop": asyncio.get_running_loop(),
            "CLOUD_SERIAL_LOG_Q": queue.Queue(maxsize=10),
            "CLOUD_SERIAL_LOG_DRAIN_STATE": CloudSerialLogDrainState(drain_scheduled=True),
            "CLOUD_SERIAL_LOG_DRAIN_BATCH": 100,
            "_cloud_build_serial_log_payload": build_serial_log_payload,
            "PORT": "synthetic-port", "BAUD": 115200,
        })

    async def test_old_call_logs_surviving_same_socket_reauthorization_are_history_only(self):
        self.setup_log_queue()
        self.start_call()
        self.assertEqual(self.namespace["_cloud_send_serial_log"]('+CLIP: "10086",129'), "queued")
        original = self.namespace["CLOUD_SERIAL_LOG_Q"].queue[0]
        before = dict(original)
        set_cloud_command_authorized(self.namespace, False)
        interrupt_cloud_sms_event_drain(self.namespace["CLOUD_SMS_EVENT_DRAIN_STATE"])
        await self.namespace["_cloud_drain_serial_log_queue"](self.ws)
        self.assertFalse(self.writes)
        self.assertIs(self.namespace["CLOUD_SERIAL_LOG_Q"].queue[0], original)
        finish_incoming_call_session_namespace_runtime(self.namespace)
        set_cloud_command_authorized(self.namespace, True)
        self.namespace["CLOUD_SERIAL_LOG_DRAIN_STATE"].drain_scheduled = True
        self.assertEqual(self.namespace["_cloud_send_serial_log"]("synthetic new log"), "queued")
        await self.namespace["_cloud_drain_serial_log_queue"](self.ws)
        self.assertTrue(all(value.get("history_only") is True for _, value in self.writes))
        self.assertEqual(original, before)
        self.assertNotIn("history_only", original)
        self.assertEqual(self.writes[0][1]["raw"], before["raw"])
        self.assertEqual(self.namespace["CLOUD_SERIAL_LOG_Q"].unfinished_tasks, 0)

    async def test_ended_call_log_backlog_is_history_only_without_muting_new_structured_call(self):
        self.setup_log_queue()
        self.start_call()
        for line in ("RING", '+CLIP: "10086",129', "[I]-[handler_call.incoming] synthetic call"):
            self.assertEqual(self.namespace["_cloud_send_serial_log"](line), "queued")
        original = list(self.namespace["CLOUD_SERIAL_LOG_Q"].queue)
        before = [dict(payload) for payload in original]
        finish_incoming_call_session_namespace_runtime(self.namespace)
        self.start_call("call-b")
        await self.namespace["_cloud_drain_serial_log_queue"](self.ws)
        self.assertTrue(all(value.get("history_only") is True for _, value in self.writes))
        self.assertEqual(original, before)
        current = self.payload("call-b")
        self.namespace["CLOUD_SMS_EVENT_Q"].put_nowait(current)
        await self.drain()
        self.assertIs(self.writes[-1][1]["history_only"], False)
        self.assertEqual([value["raw"] for _, value in self.writes[:-1]],
                         [payload["raw"] for payload in before])
        self.assertEqual(self.namespace["CLOUD_SERIAL_LOG_Q"].unfinished_tasks, 0)

    def local_control_popup(self):
        self.start_call()
        namespace = self.namespace
        self.control_results, self.control_ui, self.control_restored = [], [], []
        self.defer_control_ui = False

        def on_ui(callback, _post):
            if self.defer_control_ui:
                self.control_ui.append(callback)
                return None
            return callback()

        def open_popup(_parent, caller, _center, answer, hangup, *_close):
            return SimpleNamespace(answer=answer, hangup=hangup,
                                   winfo_exists=lambda: True, _call_popup_caller_num=caller)

        def send(_lock, _getter, command, **kwargs):
            self.control_results.append((command, kwargs["on_result"]))

        namespace.update({
            "root": None, "serial_obj": object(), "serial_lock": threading.RLock(),
            "serial_connection_generation": 1, "current_call_popup": None,
            "center_window": lambda *_: None, "port_ui": lambda *_: None,
            "set_status": lambda *_: None, "run_on_ui_thread": on_ui,
            "ui_post": self.control_ui.append, "close_call_popup": lambda: None,
        })
        return show_call_popup_namespace_runtime(namespace, "synthetic-caller", "call-a",
            show_app_runtime=partial(show_call_popup_app_runtime,
                show_runtime=partial(show_call_popup_runtime, open_popup=open_popup,
                                     send_command_async=send)))

    def click_local_control(self, popup, action):
        restore = lambda: self.control_restored.append(action)
        if action == "answer":
            popup.answer(lambda: None, restore)
        else:
            popup.hangup(restore)

    async def test_failed_local_answer_and_hangup_keep_live_notification_and_handled_semantics(self):
        for action in ("answer", "hangup"):
            with self.subTest(action=action):
                self.setUp()
                popup = self.local_control_popup()
                token = self.namespace["INCOMING_CALL_SESSION"].snapshot().session_token
                self.click_local_control(popup, action)
                self.control_results[0][1](SerialCommandResult(False, "synthetic write failure"))
                while self.control_ui:
                    self.control_ui.pop(0)()
                current = self.namespace["INCOMING_CALL_SESSION"].snapshot()
                self.assertIs(current.session_token, token)
                self.assertTrue(current.handled)
                self.assertFalse(current.confirmed_handled)
                self.assertEqual(self.control_restored, [action])
                await self.namespace["_cloud_send_payload"](self.ws, self.payload())
                self.assertIs(self.writes[-1][1]["history_only"], False)
                self.assertIsNone(self.namespace["INCOMING_CALL_SESSION"].finish())

    async def test_successful_local_control_is_confirmed_before_ui_and_late_failure_cannot_clear_it(self):
        for action in ("answer", "hangup"):
            with self.subTest(action=action):
                self.setUp()
                popup = self.local_control_popup()
                self.click_local_control(popup, action)
                self.defer_control_ui = True
                callback = self.control_results[0][1]
                callback(SerialCommandResult(True))
                self.assertTrue(self.control_ui)
                self.assertGreater(self.namespace["ring_timeout_target"], self.now)
                self.assertTrue(self.namespace["INCOMING_CALL_SESSION"].snapshot().confirmed_handled)
                callback(SerialCommandResult(False, "synthetic late failure"))
                self.assertTrue(self.namespace["INCOMING_CALL_SESSION"].snapshot().confirmed_handled)
                await self.namespace["_cloud_send_payload"](self.ws, self.payload())
                self.assertIs(self.writes[-1][1]["history_only"], True)

    async def test_old_local_success_does_not_confirm_replacement_call_or_serial(self):
        for change in ("call", "serial", "generation"):
            with self.subTest(change=change):
                self.setUp()
                popup = self.local_control_popup()
                self.click_local_control(popup, "answer")
                if change == "call":
                    finish_incoming_call_session_namespace_runtime(self.namespace)
                    self.start_call()
                elif change == "serial":
                    self.namespace["serial_obj"] = object()
                else:
                    self.namespace["serial_connection_generation"] += 1
                self.control_results[0][1](SerialCommandResult(True))
                self.assertFalse(self.namespace["INCOMING_CALL_SESSION"].snapshot().confirmed_handled)
                self.assertFalse(self.control_ui)

    async def test_trusted_connected_confirmation_is_pure_state_before_ui(self):
        self.start_call()
        queued = []
        self.namespace.update({"current_call_popup": None,
            "run_on_ui_thread": lambda callback, _post: queued.append(callback),
            "ui_post": queued.append})
        mark_call_popup_connected_namespace_runtime(self.namespace, "call-a", "synthetic-caller")
        self.assertTrue(queued)
        self.assertGreater(self.namespace["ring_timeout_target"], self.now)
        self.assertTrue(self.namespace["INCOMING_CALL_SESSION"].snapshot().confirmed_handled)
        await self.namespace["_cloud_send_payload"](self.ws, self.payload())
        self.assertIs(self.writes[-1][1]["history_only"], True)
        self.assertIsNone(finish_incoming_call_session_namespace_runtime(self.namespace))

    async def test_old_pending_same_id_and_popup_refresh_do_not_confirm_new_token(self):
        self.start_call()
        previous = self.namespace["INCOMING_CALL_SESSION"].snapshot().session_token
        finish_incoming_call_session_namespace_runtime(self.namespace)
        self.start_call()
        current = self.namespace["INCOMING_CALL_SESSION"].snapshot().session_token
        self.assertIsNot(previous, current)
        self.namespace["call_popup_pending_connected_session_id"] = "call-a"
        self.namespace["mark_incoming_call_handled"] = self.namespace["INCOMING_CALL_SESSION"].mark_handled
        window = SimpleNamespace(winfo_exists=lambda: True, _call_popup_mark_connected=lambda: None)
        set_call_popup_namespace_runtime(self.namespace, window, call_session_id="call-a")
        self.assertFalse(self.namespace["INCOMING_CALL_SESSION"].snapshot().confirmed_handled)
        await self.namespace["_cloud_send_payload"](self.ws, self.payload())
        self.assertIs(self.writes[-1][1]["history_only"], False)

    def cloud_control_modem(self, response="OK", after_write=None):
        self.start_call()
        namespace = self.namespace
        coordinator = AtCommandResponseCoordinator()

        class Modem:
            is_open = True

            def __init__(self):
                self.writes = []

            def write(self, data):
                self.writes.append(data)
                if after_write is not None:
                    after_write()
                coordinator.observe_line(response, connection=self)
                return len(data)

            def flush(self):
                pass

        modem = Modem()
        queued = []
        namespace.update({"serial_lock": threading.RLock(), "serial_obj": modem,
            "serial_connection_generation": 1, "CLOUD_DEVICE_SECRET": "synthetic-only",
            "CLOUD_WS_URL": "wss://synthetic.invalid/", "CLOUD_SENSITIVE_COMMAND_PERMISSIONS": True,
            "SERIAL_COMMAND_RESPONSE_COORDINATOR": coordinator, "current_call_popup": None,
            "_cloud_log": lambda *_args, **_kwargs: None,
            "run_on_ui_thread": lambda callback, _post: queued.append(callback), "ui_post": queued.append,
            "mark_call_connected": lambda *args, **kwargs: mark_call_popup_connected_namespace_runtime(
                namespace, *args, **kwargs)})
        return modem

    async def test_raw_and_protected_cloud_controls_confirm_only_real_modem_success_before_ui(self):
        for command in ("ATA", "ATH"):
            for protected in (False, True):
                for response in ("OK", "ERROR"):
                    with self.subTest(command=command, protected=protected, response=response):
                        self.setUp()
                        modem = self.cloud_control_modem(response)
                        meta = ({"call_control_guard": 1, "call_direction": "incoming",
                                 "call_session_id": "call-a", "call_phone": "synthetic-caller"}
                                if protected else None)
                        result = send_cloud_serial_command_namespace_runtime(self.namespace, command, meta)
                        self.assertEqual(result[0], response == "OK")
                        self.assertEqual(modem.writes, [(command + "\r\n").encode()])
                        self.assertGreater(self.namespace["ring_timeout_target"], self.now)
                        self.assertEqual(self.namespace["INCOMING_CALL_SESSION"].snapshot().confirmed_handled,
                                         response == "OK")
                        await self.namespace["_cloud_send_payload"](self.ws, self.payload())
                        self.assertEqual(self.writes[-1][1]["history_only"], response == "OK")

    async def test_written_cloud_success_remains_confirmed_after_auth_revoke_and_reauthorize(self):
        for command in ("ATA", "ATH"):
            with self.subTest(command=command):
                self.setUp()
                modem = self.cloud_control_modem(after_write=lambda: set_cloud_command_authorized(
                    self.namespace, False))
                send_cloud_serial_command_namespace_runtime(self.namespace, command)
                self.assertEqual(modem.writes, [(command + "\r\n").encode()])
                self.assertTrue(self.namespace["INCOMING_CALL_SESSION"].snapshot().confirmed_handled)
                set_cloud_command_authorized(self.namespace, True)
                await self.namespace["_cloud_send_payload"](self.ws, self.payload())
                self.assertIs(self.writes[-1][1]["history_only"], True)

    async def test_old_cloud_success_or_success_without_write_does_not_confirm_current_call(self):
        for change in ("call", "serial", "generation", "preauth", "unwritten"):
            with self.subTest(change=change):
                self.setUp()

                def after_write():
                    if change == "call":
                        finish_incoming_call_session_namespace_runtime(self.namespace)
                        self.start_call()
                    elif change == "serial":
                        self.namespace["serial_obj"] = object()
                    elif change == "generation":
                        self.namespace["serial_connection_generation"] += 1

                modem = self.cloud_control_modem(after_write=after_write)
                if change == "preauth":
                    set_cloud_command_authorized(self.namespace, False)
                    result = send_cloud_serial_command_namespace_runtime(self.namespace, "ATH")
                    self.assertFalse(result[0])
                    self.assertFalse(modem.writes)
                elif change == "unwritten":
                    result = send_cloud_serial_command_namespace_runtime(self.namespace, "ATH",
                        send_runtime=lambda *_args, **_kwargs: (True, "synthetic unsubmitted result"))
                    self.assertTrue(result[0])
                    self.assertFalse(modem.writes)
                else:
                    send_cloud_serial_command_namespace_runtime(self.namespace, "ATH")
                    self.assertTrue(modem.writes)
                self.assertFalse(self.namespace["INCOMING_CALL_SESSION"].snapshot().confirmed_handled)


if __name__ == "__main__":
    unittest.main()
