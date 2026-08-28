import asyncio
import json
import unittest

from sms_core.cloud_command_security import (
    CLOUD_SEND_SMS_TRANSACTION_COMMAND,
    CLOUD_SET_OWN_NUMBER_TRANSACTION_COMMAND,
)
from sms_core.cloud_message_runtime import (
    cloud_session_revoke_proof,
    cloud_status_payload_runtime,
    handle_cloud_message_runtime,
    send_cloud_call_event_runtime,
    send_cloud_register_runtime,
    send_cloud_session_revoke_runtime,
    send_cloud_serial_command_runtime,
    send_cloud_sms_event_runtime,
    send_cloud_unregister_runtime,
)


def run(coro):
    return asyncio.run(coro)


class FakeLoop:
    def __init__(self, running=True):
        self.running = running

    def is_running(self):
        return self.running


class DummyLock:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


class FakeSerial:
    def __init__(self, is_open=True):
        self.is_open = is_open


class SimpleResult:
    def __init__(self, ok, error=""):
        self.ok = ok
        self.error = error


class FakeWs:
    def __init__(self, *, fail=False):
        self.fail = fail
        self.sent = []

    async def send(self, payload):
        if self.fail:
            raise RuntimeError("send failed")
        self.sent.append(payload)


class CloudMessageRuntimeTests(unittest.TestCase):
    async def _handle(self, message, *, authorized=False, auth_ok=True, replay_ok=True):
        state = {"authorized": authorized}
        calls = []
        replies = []
        replay_checks = []

        async def reply(payload):
            replies.append(payload)

        async def check_replay_window(data, mark_seen=True):
            replay_checks.append(mark_seen)
            return replay_ok

        async def send_serial_command(command, command_data=None):
            calls.append(("send", command, command_data))
            return True, f"sent {command}"

        await handle_cloud_message_runtime(
            message,
            is_authorized=lambda: state["authorized"],
            set_authorized=lambda value: state.__setitem__("authorized", bool(value)),
            reply=reply,
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            set_auth_status_from_ack=lambda data: calls.append(("ack", data)),
            set_cloud_status=lambda *args: calls.append(("status", args)),
            check_replay_window=check_replay_window,
            auth_matches=lambda data: auth_ok,
            time_text=lambda: "now",
            timestamp=lambda: 123,
            status_payload=lambda: {"type": "status"},
            send_serial_command=send_serial_command,
            show_window=lambda: calls.append(("show",)),
            hide_window=lambda: calls.append(("hide",)),
        )
        return state, calls, replies, replay_checks

    def test_send_session_revoke_uses_identity_without_secret(self):
        ws = FakeWs()
        logs = []

        result = run(send_cloud_session_revoke_runtime(
            ws,
            reason="password_changed",
            build_payload=lambda reason, timestamp, identity: {
                "type": "device_session_revoke",
                "reason": reason,
                "timestamp": timestamp,
                **identity,
            },
            timestamp=lambda: 123,
            identity_payload=lambda: {"imei": "861"},
            log=logs.append,
        ))

        self.assertEqual(result, "sent")
        self.assertEqual(json.loads(ws.sent[0]), {
            "type": "device_session_revoke",
            "reason": "password_changed",
            "timestamp": 123,
            "imei": "861",
        })
        self.assertEqual(logs, [])

    def test_authorized_ack_updates_state_without_reply(self):
        state, calls, replies, replay_checks = run(self._handle(json.dumps({
            "type": "device_login_ack",
            "auth_status": "ok",
            "message": "ok",
        })))

        self.assertTrue(state["authorized"])
        self.assertEqual(replies, [])
        self.assertEqual(replay_checks, [])
        self.assertTrue(any(item[0] == "ack" for item in calls))
        self.assertTrue(any(item[0] == "log" and item[2].get("show_main") for item in calls))

    def test_failed_auth_ack_returns_disconnect_signal(self):
        state = {"authorized": True}
        calls = []

        async def reply(_payload):
            raise AssertionError("auth ack must not send a command reply")

        result = run(handle_cloud_message_runtime(
            json.dumps({
                "type": "device_login_ack",
                "auth_status": "auth_failed",
                "message": "bad secret",
            }),
            is_authorized=lambda: state["authorized"],
            set_authorized=lambda value: state.__setitem__("authorized", bool(value)),
            reply=reply,
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            set_auth_status_from_ack=lambda data: calls.append(("ack", data)),
            set_cloud_status=lambda *args: calls.append(("status", args)),
            check_replay_window=lambda *_args, **_kwargs: None,
            auth_matches=lambda _data: True,
            time_text=lambda: "now",
            timestamp=lambda: 123,
            status_payload=lambda: {},
            send_serial_command=lambda *_args: None,
            show_window=lambda: None,
            hide_window=lambda: None,
        ))

        self.assertEqual(result, "auth_failed")
        self.assertFalse(state["authorized"])
        self.assertTrue(any(item[0] == "status" and item[1][0] == "🌐 授权失败" for item in calls))

    def test_waiting_auth_ack_keeps_connection_without_authorizing(self):
        state = {"authorized": True}
        calls = []

        async def reply(_payload):
            raise AssertionError("auth ack must not send a command reply")

        result = run(handle_cloud_message_runtime(
            json.dumps({
                "type": "device_login_ack",
                "auth_status": "waiting",
                "message": "waiting for binding",
            }),
            is_authorized=lambda: state["authorized"],
            set_authorized=lambda value: state.__setitem__("authorized", bool(value)),
            reply=reply,
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            set_auth_status_from_ack=lambda data: calls.append(("ack", data)),
            set_cloud_status=lambda *args: calls.append(("status", args)),
            check_replay_window=lambda *_args, **_kwargs: None,
            auth_matches=lambda _data: True,
            time_text=lambda: "now",
            timestamp=lambda: 123,
            status_payload=lambda: {},
            send_serial_command=lambda *_args: None,
            show_window=lambda: None,
            hide_window=lambda: None,
        ))

        self.assertEqual(result, "waiting")
        self.assertFalse(state["authorized"])
        self.assertTrue(any(item[0] == "status" and item[1][0] == "🌐 等待授权" for item in calls))

    def test_unauthorized_command_replies_with_task_ids(self):
        state, calls, replies, replay_checks = run(self._handle(json.dumps({
            "cmd": "ATI",
            "task_id": "task-1",
        })))

        self.assertFalse(state["authorized"])
        self.assertEqual(replay_checks, [])
        self.assertEqual(replies[0]["type"], "auth_failed")
        self.assertEqual(replies[0]["task_id"], "task-1")
        self.assertEqual(replies[0]["command_task_id"], "task-1")

    def test_auth_failure_replies_before_marking_replay_seen(self):
        state, calls, replies, replay_checks = run(self._handle(
            json.dumps({"cmd": "ATI", "task_id": "task-1"}),
            authorized=True,
            auth_ok=False,
        ))

        self.assertTrue(state["authorized"])
        self.assertEqual(replay_checks, [False])
        self.assertEqual(replies[0]["type"], "auth_failed")
        self.assertTrue(any(item[0] == "status" for item in calls))

    def test_authorized_send_at_dispatches_and_replies(self):
        state, calls, replies, replay_checks = run(self._handle(
            json.dumps({"cmd": "ATI", "task_id": "task-1"}),
            authorized=True,
            auth_ok=True,
        ))

        self.assertTrue(state["authorized"])
        self.assertEqual(replay_checks, [False, True])
        self.assertIn(("send", "ATI", {"cmd": "ATI", "task_id": "task-1"}), calls)
        self.assertEqual(replies[0]["type"], "command_task_status")
        self.assertEqual(replies[0]["status"], "started")
        self.assertEqual(replies[0]["task_id"], "task-1")
        self.assertEqual(replies[1]["type"], "send_at_result")
        self.assertEqual(replies[0]["task_id"], "task-1")
        self.assertTrue(replies[0]["ok"])

    def test_replay_rejection_stops_before_auth_and_dispatch(self):
        state, calls, replies, replay_checks = run(self._handle(
            json.dumps({"cmd": "ATI", "task_id": "task-1"}),
            authorized=True,
            replay_ok=False,
        ))

        self.assertTrue(state["authorized"])
        self.assertEqual(replay_checks, [False])
        self.assertEqual(replies, [])
        self.assertNotIn(("send", "ATI"), calls)

    def test_handle_cloud_message_hides_suppressed_pdu_in_receive_log(self):
        pdu = "0011000D916831...\x1a"
        state, calls, replies, replay_checks = run(self._handle(
            json.dumps({
                "cmd": pdu,
                "command": pdu,
                "data": pdu,
                "sms_log": "suppress",
                "secret": "device-secret",
                "task_id": "task-1",
            }),
            authorized=True,
        ))

        receive_logs = [
            item[1][0]
            for item in calls
            if item[0] == "log" and item[1] and str(item[1][0]).startswith("收到：")
        ]

        self.assertTrue(state["authorized"])
        self.assertTrue(receive_logs)
        self.assertFalse(any(pdu in item for item in receive_logs))
        self.assertFalse(any("device-secret" in item for item in receive_logs))
        self.assertTrue(any("已隐藏" in item for item in receive_logs))

    def test_handle_cloud_message_hides_sms_summary_metadata_in_receive_log(self):
        phone = "+8613123123123"
        message = "验证码 1234"
        state, calls, replies, replay_checks = run(self._handle(
            json.dumps({
                "cmd": "AT+CMGF=0",
                "command_kind": "send_sms",
                "sms_log": "summary",
                "sms_phone": phone,
                "sms_message": message,
                "task_id": "task-1",
            }),
            authorized=True,
        ))

        receive_logs = [
            item[1][0]
            for item in calls
            if item[0] == "log" and item[1] and str(item[1][0]).startswith("收到：")
        ]

        self.assertTrue(state["authorized"])
        self.assertTrue(receive_logs)
        self.assertFalse(any(phone in item for item in receive_logs))
        self.assertFalse(any(message in item for item in receive_logs))
        self.assertTrue(any("短信元数据" in item for item in receive_logs))

    def test_send_cloud_sms_event_runtime_skips_unavailable_states(self):
        base = {
            "authorized": True,
            "get_loop": lambda: FakeLoop(True),
            "get_ws": lambda: object(),
            "is_connected": lambda: True,
            "runtime_imei": lambda: "imei",
            "build_payload": lambda *_args: {"type": "sms_event"},
            "send_payload": lambda *_args: "coro",
            "timestamp": lambda: 123,
            "identity_payload": lambda: {"imei": "imei"},
            "run_coroutine_threadsafe": lambda *_args: None,
        }

        self.assertEqual(send_cloud_sms_event_runtime("head", "", **base), "empty")
        self.assertEqual(send_cloud_sms_event_runtime("head", "body", **{**base, "authorized": False}), "unauthorized")
        self.assertEqual(send_cloud_sms_event_runtime("head", "body", **{**base, "get_loop": lambda: None}), "not_connected")
        self.assertEqual(send_cloud_sms_event_runtime("head", "body", **{**base, "get_loop": lambda: FakeLoop(False)}), "not_connected")
        self.assertEqual(send_cloud_sms_event_runtime("head", "body", **{**base, "get_ws": lambda: None}), "not_connected")
        self.assertEqual(send_cloud_sms_event_runtime("head", "body", **{**base, "is_connected": lambda: False}), "not_connected")
        self.assertEqual(send_cloud_sms_event_runtime("head", "body", **{**base, "runtime_imei": lambda: ""}), "missing_imei")
        self.assertEqual(send_cloud_sms_event_runtime("head", "body", **{**base, "build_payload": lambda *_args: None}), "empty_payload")

    def test_send_cloud_sms_event_runtime_schedules_payload(self):
        calls = []
        loop = FakeLoop(True)
        ws = object()

        result = send_cloud_sms_event_runtime(
            "head",
            " body ",
            authorized=True,
            get_loop=lambda: loop,
            get_ws=lambda: ws,
            is_connected=lambda: True,
            runtime_imei=lambda: "imei",
            build_payload=lambda head, body, ts, identity: {
                "head": head,
                "body": body,
                "ts": ts,
                **identity,
            },
            send_payload=lambda next_ws, payload: ("send", next_ws, payload),
            timestamp=lambda: 123,
            identity_payload=lambda: {"imei": "imei"},
            run_coroutine_threadsafe=lambda coro, next_loop: calls.append((coro, next_loop)),
        )

        self.assertEqual(result, "scheduled")
        self.assertEqual(calls, [(("send", ws, {"head": "head", "body": "body", "ts": 123, "imei": "imei"}), loop)])

    def test_send_cloud_sms_event_runtime_passes_metadata_to_payload(self):
        calls = []
        loop = FakeLoop(True)
        ws = object()

        result = send_cloud_sms_event_runtime(
            "head",
            "body",
            {"message_trace_id": "abc123def456"},
            authorized=True,
            get_loop=lambda: loop,
            get_ws=lambda: ws,
            is_connected=lambda: True,
            runtime_imei=lambda: "imei",
            build_payload=lambda head, body, ts, identity, metadata=None: {
                "head": head,
                "body": body,
                "metadata": metadata,
                **identity,
            },
            send_payload=lambda next_ws, payload: ("send", next_ws, payload),
            timestamp=lambda: 123,
            identity_payload=lambda: {"imei": "imei"},
            run_coroutine_threadsafe=lambda coro, next_loop: calls.append((coro, next_loop)),
        )

        self.assertEqual(result, "scheduled")
        self.assertEqual(calls[0][0][2]["metadata"], {"message_trace_id": "abc123def456"})

    def test_send_cloud_sms_event_runtime_closes_coro_when_submit_fails(self):
        created = []

        async def send_payload(_ws, _payload):
            return None

        def build_send_coro(ws, payload):
            coro = send_payload(ws, payload)
            created.append(coro)
            return coro

        result = send_cloud_sms_event_runtime(
            "head",
            "body",
            authorized=True,
            get_loop=lambda: FakeLoop(True),
            get_ws=lambda: object(),
            is_connected=lambda: True,
            runtime_imei=lambda: "imei",
            build_payload=lambda *_args: {"type": "sms_event"},
            send_payload=build_send_coro,
            timestamp=lambda: 123,
            identity_payload=lambda: {"imei": "imei"},
            run_coroutine_threadsafe=lambda _coro, _loop: (_ for _ in ()).throw(
                RuntimeError("loop rejected")
            ),
        )

        self.assertEqual(result, "error")
        self.assertEqual(len(created), 1)
        self.assertIsNone(created[0].cr_frame)

    def test_send_cloud_call_event_runtime_uses_shared_event_queue_callback(self):
        calls = []
        loop = FakeLoop(True)
        ws = object()

        result = send_cloud_call_event_runtime(
            "10086",
            "incoming",
            blocked=True,
            block_reason="blacklist",
            authorized=False,
            get_loop=lambda: loop,
            get_ws=lambda: ws,
            is_connected=lambda: False,
            runtime_imei=lambda: "imei",
            build_payload=lambda caller, message, ts, identity, **metadata: {
                "caller": caller,
                "message": message,
                "ts": ts,
                **identity,
                **metadata,
            },
            send_payload=lambda *_args: None,
            timestamp=lambda: 123,
            identity_payload=lambda: {"imei": "imei"},
            run_coroutine_threadsafe=lambda *_args: None,
            enqueue_payload=lambda payload, next_loop, next_ws, can_send: (
                calls.append((payload, next_loop, next_ws, can_send)) or "queued_offline"
            ),
        )

        self.assertEqual(result, "queued_offline")
        self.assertEqual(calls[0][0]["caller"], "10086")
        self.assertTrue(calls[0][0]["blocked"])
        self.assertEqual(calls[0][0]["block_reason"], "blacklist")
        self.assertFalse(calls[0][3])

    def test_send_cloud_call_event_runtime_supports_legacy_payload_builder(self):
        calls = []
        result = send_cloud_call_event_runtime(
            "10086",
            "incoming",
            authorized=True,
            get_loop=lambda: FakeLoop(True),
            get_ws=lambda: object(),
            is_connected=lambda: True,
            runtime_imei=lambda: "imei",
            build_payload=lambda caller, message, ts, identity: {
                "caller": caller,
                "message": message,
                "ts": ts,
                **identity,
            },
            send_payload=lambda ws, payload: ("send", ws, payload),
            timestamp=lambda: 123,
            identity_payload=lambda: {"imei": "imei"},
            run_coroutine_threadsafe=lambda coro, loop: calls.append((coro, loop)),
        )

        self.assertEqual(result, "scheduled")
        self.assertEqual(calls[0][0][2]["caller"], "10086")

    def test_cloud_status_payload_runtime_reports_serial_connection(self):
        payload = cloud_status_payload_runtime(
            serial_lock=DummyLock(),
            get_serial=lambda: FakeSerial(is_open=True),
            build_payload=lambda ts, identity, cloud, serial, port, baud, mode: {
                "ts": ts,
                **identity,
                "cloud": cloud,
                "serial": serial,
                "port": port,
                "baud": baud,
                "mode": mode,
            },
            timestamp=lambda: 123,
            identity_payload=lambda: {"imei": "imei"},
            cloud_connected=True,
            serial_port="COM5",
            serial_baud=115200,
            serial_mode="Auto",
        )

        self.assertEqual(payload["ts"], 123)
        self.assertTrue(payload["cloud"])
        self.assertTrue(payload["serial"])
        self.assertEqual(payload["port"], "COM5")

    def test_cloud_status_payload_runtime_treats_serial_errors_as_disconnected(self):
        payload = cloud_status_payload_runtime(
            serial_lock=DummyLock(),
            get_serial=lambda: (_ for _ in ()).throw(RuntimeError("boom")),
            build_payload=lambda _ts, _identity, _cloud, serial, *_args: {"serial": serial},
            timestamp=lambda: 123,
            identity_payload=lambda: {},
            cloud_connected=False,
            serial_port="COM5",
            serial_baud=115200,
            serial_mode="Auto",
        )

        self.assertFalse(payload["serial"])

    def test_send_cloud_register_runtime_sends_public_payload(self):
        calls = []
        ws = FakeWs()

        result = run(send_cloud_register_runtime(
            ws,
            auto_upload=True,
            build_payload=lambda auto, ts, identity, secret, port, baud, mode: {
                "auto": auto,
                "ts": ts,
                **identity,
                "secret": secret,
                "port": port,
                "baud": baud,
                "mode": mode,
            },
            timestamp=lambda: 123,
            identity_payload=lambda: {"imei": "imei"},
            secret="secret",
            serial_port="COM5",
            serial_baud=115200,
            serial_mode="Auto",
            runtime_imei=lambda: "imei",
            log=lambda message, **kwargs: calls.append((message, kwargs)),
        ))

        self.assertEqual(result, "sent")
        self.assertEqual(json.loads(ws.sent[0])["imei"], "imei")
        self.assertIn("IMEI", calls[0][0])
        self.assertTrue(calls[0][1].get("show_main"))

    def test_send_cloud_register_runtime_adds_hmac_proof_without_old_secret(self):
        ws = FakeWs()

        result = run(send_cloud_register_runtime(
            ws,
            auto_upload=True,
            build_payload=lambda *_args: {
                "type": "device_login",
                "secret": "new-secret",
            },
            timestamp=lambda: 123,
            identity_payload=lambda: {"imei": "123456789012345"},
            secret="new-secret",
            serial_port="COM5",
            serial_baud=115200,
            serial_mode="Auto",
            runtime_imei=lambda: "123456789012345",
            log=lambda *_args, **_kwargs: None,
            previous_session_secret="old-secret",
        ))

        self.assertEqual(result, "sent")
        payload = json.loads(ws.sent[0])
        self.assertEqual(
            payload["previous_session_proof"],
            cloud_session_revoke_proof("old-secret", "123456789012345"),
        )
        self.assertNotEqual(payload["previous_session_proof"], "old-secret")
        self.assertNotIn("previous_secret", payload)

    def test_send_cloud_register_runtime_logs_hidden_mode(self):
        calls = []
        ws = FakeWs()

        result = run(send_cloud_register_runtime(
            ws,
            auto_upload=False,
            build_payload=lambda auto, *_args: {"auto": auto},
            timestamp=lambda: 123,
            identity_payload=lambda: {"imei": "imei"},
            secret="secret",
            serial_port="COM5",
            serial_baud=115200,
            serial_mode="Auto",
            runtime_imei=lambda: "imei",
            log=lambda message, **kwargs: calls.append((message, kwargs)),
        ))

        self.assertEqual(result, "sent")
        self.assertIn("隐身模式", calls[0][0])
        self.assertTrue(calls[0][1].get("show_main"))

    def test_send_cloud_register_runtime_logs_send_error(self):
        calls = []

        result = run(send_cloud_register_runtime(
            FakeWs(fail=True),
            auto_upload=True,
            build_payload=lambda *_args: {"type": "device_login"},
            timestamp=lambda: 123,
            identity_payload=lambda: {"imei": "imei"},
            secret="secret",
            serial_port="COM5",
            serial_baud=115200,
            serial_mode="Auto",
            runtime_imei=lambda: "imei",
            log=lambda message: calls.append(message),
        ))

        self.assertEqual(result, "error")
        self.assertIn("send failed", calls[0])

    def test_send_cloud_unregister_runtime_sends_payload(self):
        calls = []
        ws = FakeWs()

        result = run(send_cloud_unregister_runtime(
            ws,
            reason="disconnect",
            build_payload=lambda reason, ts, identity, secret, port, baud, mode: {
                "reason": reason,
                "ts": ts,
                **identity,
            },
            timestamp=lambda: 123,
            identity_payload=lambda: {"imei": "imei"},
            secret="secret",
            serial_port="COM5",
            serial_baud=115200,
            serial_mode="Auto",
            runtime_imei=lambda: "imei",
            log=lambda message: calls.append(message),
        ))

        self.assertEqual(result, "sent")
        self.assertEqual(json.loads(ws.sent[0])["reason"], "disconnect")
        self.assertIn("离线", calls[0])

    def test_send_cloud_serial_command_runtime_rejects_empty_command(self):
        ok, info = send_cloud_serial_command_runtime(
            " ",
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda *_args: None,
            push_serial_debug=lambda *_args: None,
            log=lambda *_args: None,
        )

        self.assertFalse(ok)
        self.assertIn("不能为空", info)

    def test_cloud_sms_transaction_executes_once_without_raw_marker_write(self):
        calls = []

        ok, info = send_cloud_serial_command_runtime(
            CLOUD_SEND_SMS_TRANSACTION_COMMAND,
            command_meta={
                "sms_phone": "+8613123123123",
                "sms_message": "hello",
            },
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda *_args: calls.append("raw_write") or SimpleResult(True),
            push_serial_debug=lambda *_args: None,
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
            allow_sensitive_commands={"sms": True},
            send_sms_transaction=lambda phone, message: (
                calls.append(("sms", phone, message)) or True
            ),
        )

        self.assertTrue(ok)
        self.assertEqual(info, "短信发送成功")
        self.assertIn(("sms", "+8613123123123", "hello"), calls)
        self.assertNotIn("raw_write", calls)

    def test_cloud_own_number_transaction_stops_on_transaction_failure(self):
        calls = []

        ok, info = send_cloud_serial_command_runtime(
            CLOUD_SET_OWN_NUMBER_TRANSACTION_COMMAND,
            command_meta={"own_number": "+8613123123123"},
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda *_args: calls.append("raw_write") or SimpleResult(True),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            allow_sensitive_commands={"phone_number": True},
            set_own_number_transaction=lambda phone: (
                calls.append(("own_number", phone))
                or SimpleResult(False, "AT+CPBS 执行失败")
            ),
        )

        self.assertFalse(ok)
        self.assertEqual(info, "AT+CPBS 执行失败")
        self.assertEqual(calls, [("own_number", "+8613123123123")])

    def test_cloud_transaction_markers_require_permission_and_valid_metadata(self):
        calls = []
        common = {
            "serial_lock": DummyLock(),
            "get_serial": lambda: object(),
            "write_command_result": lambda *_args: calls.append("raw_write") or SimpleResult(True),
            "push_serial_debug": lambda *_args: None,
            "log": lambda *_args, **_kwargs: None,
            "send_sms_transaction": lambda *_args: calls.append("sms") or True,
        }

        blocked, blocked_info = send_cloud_serial_command_runtime(
            CLOUD_SEND_SMS_TRANSACTION_COMMAND,
            command_meta={"sms_phone": "+8613123123123", "sms_message": "hello"},
            allow_sensitive_commands={"sms": False},
            **common,
        )
        invalid, invalid_info = send_cloud_serial_command_runtime(
            CLOUD_SEND_SMS_TRANSACTION_COMMAND,
            command_meta={"sms_phone": "bad", "sms_message": "hello"},
            allow_sensitive_commands={"sms": True},
            **common,
        )

        self.assertFalse(blocked)
        self.assertIn("安全设置", blocked_info)
        self.assertFalse(invalid)
        self.assertIn("格式", invalid_info)
        self.assertEqual(calls, [])

    def test_transaction_metadata_cannot_reclassify_regular_at_command(self):
        calls = []
        ok, _info = send_cloud_serial_command_runtime(
            "AT+RESET",
            command_meta={
                "command_kind": "modify_own_number",
                "own_number": "+8613123123123",
            },
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda _serial, command: (
                calls.append(("raw", command)) or SimpleResult(True)
            ),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            set_own_number_transaction=lambda *_args: calls.append("transaction"),
        )

        self.assertTrue(ok)
        self.assertEqual(calls, [("raw", "AT+RESET")])

    def test_send_cloud_serial_command_runtime_rejects_multiline_commands_even_when_all_permissions_enabled(self):
        commands = (
            "AT+CSQ\r\nAT+RESET",
            "AT+CSQ\nAT+CMGD=1",
            "AT+CSQ\rATD10086;",
            "AT+CSQ;AT+RESET",
            "001122\x1aAT+RESET",
        )
        for command in commands:
            with self.subTest(command=repr(command)):
                calls = []
                ok, info = send_cloud_serial_command_runtime(
                    command,
                    serial_lock=DummyLock(),
                    get_serial=lambda: object(),
                    write_command_result=lambda *_args: calls.append("write") or SimpleResult(True),
                    push_serial_debug=lambda message: calls.append(("debug", message)),
                    port_ui=lambda *args: calls.append(("port_ui", args)),
                    log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
                    allow_sensitive_commands=True,
                )

                self.assertFalse(ok)
                self.assertIn("每次只发送一条", info)
                self.assertNotIn("write", calls)
                self.assertFalse(any(command in str(call) for call in calls))

    def test_send_cloud_serial_command_runtime_rejects_unsupported_control_characters(self):
        rejected_codes = (
            *range(0x00, 0x09),
            0x0B,
            0x0C,
            *range(0x0E, 0x1A),
            *range(0x1B, 0x20),
            0x7F,
        )
        for code in rejected_codes:
            command = "AT+CSQ" + chr(code) + "AT+RESET"
            calls = []
            with self.subTest(code=hex(code)):
                ok, info = send_cloud_serial_command_runtime(
                    command,
                    serial_lock=DummyLock(),
                    get_serial=lambda: object(),
                    write_command_result=lambda *_args: calls.append("write") or SimpleResult(True),
                    push_serial_debug=lambda message: calls.append(("debug", message)),
                    port_ui=lambda *args: calls.append(("port_ui", args)),
                    log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
                    allow_sensitive_commands=True,
                )

                self.assertFalse(ok)
                self.assertIn("控制字符", info)
                self.assertNotIn("write", calls)
                self.assertFalse(any(command in str(call) for call in calls))

    def test_send_cloud_serial_command_runtime_allows_terminal_sms_ctrl_z(self):
        writes = []
        command = "001122\x1a"

        ok, _info = send_cloud_serial_command_runtime(
            command,
            command_meta={"command_kind": "send_sms", "sms_log": "suppress"},
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda _serial, value: (
                writes.append(value) or SimpleResult(True)
            ),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            allow_sensitive_commands={"sms": True},
        )

        self.assertTrue(ok)
        self.assertEqual(writes, [command])

    def test_send_cloud_serial_command_runtime_writes_without_port_log_for_generic_at(self):
        calls = []
        serial_obj = object()

        ok, info = send_cloud_serial_command_runtime(
            " ATI ",
            serial_lock=DummyLock(),
            get_serial=lambda: serial_obj,
            write_command_result=lambda next_serial, command: (
                calls.append(("write", next_serial, command)) or SimpleResult(True)
            ),
            push_serial_debug=lambda message: calls.append(("debug", message)),
            port_ui=lambda message, tag: calls.append(("port_ui", message, tag)),
            log=lambda message: calls.append(("log", message)),
        )

        self.assertTrue(ok)
        self.assertEqual(info, "执行成功：ATI")
        self.assertEqual(calls[0], ("write", serial_obj, "ATI"))
        self.assertIn("ATI", calls[1][1])
        self.assertEqual(calls[2][0], "log")
        self.assertFalse(any(call[0] == "port_ui" for call in calls))

    def test_send_cloud_serial_command_runtime_returns_write_failure(self):
        ok, info = send_cloud_serial_command_runtime(
            "ATI",
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda *_args: SimpleResult(False, "closed"),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args: None,
        )

        self.assertFalse(ok)
        self.assertEqual(info, "closed")

    def test_send_cloud_serial_command_runtime_logs_readable_sms_summary(self):
        calls = []

        ok, _info = send_cloud_serial_command_runtime(
            "AT+CMGF=0",
            command_meta={
                "sms_log": "summary",
                "sms_phone": "+8613123123123",
                "sms_message": "验证码 1234",
            },
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda *_args: SimpleResult(True),
            push_serial_debug=lambda *_args: None,
            port_ui=lambda message, tag: calls.append(("port_ui", message, tag)),
            log=lambda *_args: None,
            allow_sensitive_commands={"sms": True},
        )

        self.assertTrue(ok)
        self.assertEqual(calls, [
            ("port_ui", "云端发送短信至 +8613123123123：", "normal"),
            ("port_ui", "验证码 1234", "sms"),
        ])

    def test_send_cloud_serial_command_runtime_suppresses_sms_pdu_noise(self):
        calls = []
        pdu = "0011000D916831...\x1a"

        ok, info = send_cloud_serial_command_runtime(
            pdu,
            command_meta={"sms_log": "suppress"},
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda serial_obj, command: (
                calls.append(("write", command)) or SimpleResult(True)
            ),
            push_serial_debug=lambda message: calls.append(("debug", message)),
            port_ui=lambda message, tag: calls.append(("port_ui", message, tag)),
            log=lambda message: calls.append(("log", message)),
            allow_sensitive_commands={"sms": True},
        )

        self.assertTrue(ok)
        self.assertEqual(calls[0], ("write", pdu))
        self.assertFalse(any(call[0] == "port_ui" for call in calls))
        self.assertFalse(any(pdu in str(call) for call in calls[1:]))
        self.assertNotIn(pdu, info)
        self.assertIn("已隐藏", info)

    def test_send_cloud_serial_command_runtime_rejects_metadata_category_spoofing(self):
        writes = []
        common = {
            "serial_lock": DummyLock(),
            "get_serial": lambda: object(),
            "write_command_result": lambda _serial, command: (
                writes.append(command) or SimpleResult(True)
            ),
            "push_serial_debug": lambda *_args: None,
            "log": lambda *_args, **_kwargs: None,
            "allow_sensitive_commands": {"sms": True},
        }

        reset_ok, reset_info = send_cloud_serial_command_runtime(
            "AT+REBOOT",
            command_meta={"command_kind": "send_sms", "sms_log": "suppress"},
            **common,
        )
        pin_ok, pin_info = send_cloud_serial_command_runtime(
            'AT+CPIN="1234"',
            command_meta={"command_kind": "send_sms", "sms_log": "summary"},
            **common,
        )

        self.assertFalse(reset_ok)
        self.assertFalse(pin_ok)
        self.assertEqual(writes, [])
        self.assertIn("重置或关闭设备", reset_info)
        self.assertIn("PIN", pin_info)

    def test_send_cloud_serial_command_runtime_blocks_sensitive_commands_by_default(self):
        calls = []
        pin_command = 'AT+CPIN="1234"'

        ok, info = send_cloud_serial_command_runtime(
            pin_command,
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda *_args: calls.append("write") or SimpleResult(True),
            push_serial_debug=lambda message: calls.append(("debug", message)),
            log=lambda *args, **kwargs: calls.append(("log", args, kwargs)),
        )

        self.assertFalse(ok)
        self.assertNotIn(pin_command, info)
        self.assertIn("PIN", info)
        self.assertNotIn("write", calls)
        self.assertFalse(any(pin_command in str(call) for call in calls))

    def test_send_cloud_serial_command_runtime_applies_permissions_independently(self):
        writes = []
        common = {
            "serial_lock": DummyLock(),
            "get_serial": lambda: object(),
            "write_command_result": lambda _serial, command: (
                writes.append(command) or SimpleResult(True)
            ),
            "push_serial_debug": lambda *_args: None,
            "log": lambda *_args, **_kwargs: None,
            "allow_sensitive_commands": {"sms": True, "call": False},
        }

        sms_ok, _sms_info = send_cloud_serial_command_runtime(
            "AT+CMGS=23",
            **common,
        )
        call_ok, call_info = send_cloud_serial_command_runtime(
            "ATD10086;",
            **common,
        )

        self.assertTrue(sms_ok)
        self.assertFalse(call_ok)
        self.assertEqual(writes, ["AT+CMGS=23"])
        self.assertIn("电话", call_info)

    def test_send_cloud_serial_command_runtime_controls_cell_location_query(self):
        writes = []

        blocked, blocked_info = send_cloud_serial_command_runtime(
            "AT+EEMGINFO?",
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda _serial, command: (
                writes.append(command) or SimpleResult(True)
            ),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            allow_sensitive_commands={"cell_location": False},
        )
        allowed, _allowed_info = send_cloud_serial_command_runtime(
            "AT+EEMGINFO?",
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda _serial, command: (
                writes.append(command) or SimpleResult(True)
            ),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            allow_sensitive_commands={"cell_location": True},
        )

        self.assertFalse(blocked)
        self.assertIn("基站定位", blocked_info)
        self.assertTrue(allowed)
        self.assertEqual(writes, ["AT+EEMGINFO?"])

    def test_send_cloud_serial_command_runtime_controls_puk_separately_from_pin(self):
        writes = []
        puk_command = 'AT+CPIN="12345678","1234"'

        blocked, blocked_info = send_cloud_serial_command_runtime(
            puk_command,
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda _serial, command: (
                writes.append(command) or SimpleResult(True)
            ),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            allow_sensitive_commands={"pin": True, "puk": False},
        )
        allowed, _allowed_info = send_cloud_serial_command_runtime(
            puk_command,
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda _serial, command: (
                writes.append(command) or SimpleResult(True)
            ),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            allow_sensitive_commands={"pin": False, "puk": True},
        )

        self.assertFalse(blocked)
        self.assertIn("PUK", blocked_info)
        self.assertTrue(allowed)
        self.assertEqual(writes, [puk_command])

    def test_send_cloud_serial_command_runtime_includes_pin_lock_in_pin_operations(self):
        writes = []
        lock_command = 'AT+CLCK="SC",1,"1234"'

        blocked, blocked_info = send_cloud_serial_command_runtime(
            lock_command,
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda _serial, command: (
                writes.append(command) or SimpleResult(True)
            ),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            allow_sensitive_commands={"pin": False},
        )
        allowed, _allowed_info = send_cloud_serial_command_runtime(
            lock_command,
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda _serial, command: (
                writes.append(command) or SimpleResult(True)
            ),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            allow_sensitive_commands={"pin": True},
        )

        self.assertFalse(blocked)
        self.assertIn("PIN 码操作", blocked_info)
        self.assertTrue(allowed)
        self.assertEqual(writes, [lock_command])

    def test_send_cloud_serial_command_runtime_accepts_legacy_pin_lock_permission(self):
        writes = []
        lock_command = 'AT+CLCK="SC",1,"1234"'

        allowed, _info = send_cloud_serial_command_runtime(
            lock_command,
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda _serial, command: (
                writes.append(command) or SimpleResult(True)
            ),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            allow_sensitive_commands={"pin_lock": True},
        )

        self.assertTrue(allowed)
        self.assertEqual(writes, [lock_command])

    def test_send_cloud_serial_command_runtime_controls_new_security_groups(self):
        cases = (
            ("ussd", 'AT+CUSD=1,"*100#"', "USSD"),
            ("call_control", "AT+CCFC=0,3", "呼叫转移"),
            ("sms_center", 'AT+CSCA="+8613800100500"', "信息中心"),
            ("delete_data", "AT+CMGD=1", "删除设备数据"),
            ("device_power", "AT+REBOOT", "重置或关闭设备"),
        )

        for category, command, expected_reason in cases:
            with self.subTest(category=category):
                writes = []
                common = {
                    "serial_lock": DummyLock(),
                    "get_serial": lambda: object(),
                    "write_command_result": lambda _serial, value: (
                        writes.append(value) or SimpleResult(True)
                    ),
                    "push_serial_debug": lambda *_args: None,
                    "log": lambda *_args, **_kwargs: None,
                }

                blocked, blocked_info = send_cloud_serial_command_runtime(
                    command,
                    allow_sensitive_commands={category: False},
                    **common,
                )
                allowed, _allowed_info = send_cloud_serial_command_runtime(
                    command,
                    allow_sensitive_commands={category: True},
                    **common,
                )

                self.assertFalse(blocked)
                self.assertIn(expected_reason, blocked_info)
                self.assertTrue(allowed)
                self.assertEqual(writes, [command])

    def test_ussd_permission_does_not_allow_call_forwarding_mmi_code(self):
        command = "ATD**21*13800138000#;"

        blocked, blocked_info = send_cloud_serial_command_runtime(
            command,
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda *_args: SimpleResult(True),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            allow_sensitive_commands={"ussd": True, "call_control": False},
        )

        self.assertFalse(blocked)
        self.assertIn("呼叫转移", blocked_info)

    def test_pin_permission_does_not_allow_call_barring_password_change(self):
        command = 'AT+CPWD="AO","1234","5678"'

        blocked, blocked_info = send_cloud_serial_command_runtime(
            command,
            serial_lock=DummyLock(),
            get_serial=lambda: object(),
            write_command_result=lambda *_args: SimpleResult(True),
            push_serial_debug=lambda *_args: None,
            log=lambda *_args, **_kwargs: None,
            allow_sensitive_commands={"pin": True, "call_control": False},
        )

        self.assertFalse(blocked)
        self.assertIn("呼叫限制", blocked_info)


if __name__ == "__main__":
    unittest.main()
