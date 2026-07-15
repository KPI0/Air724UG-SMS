import asyncio
import json
import unittest

from sms_core.cloud_message_runtime import (
    cloud_status_payload_runtime,
    handle_cloud_message_runtime,
    send_cloud_register_runtime,
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
        self.assertEqual(replies[0]["type"], "send_at_result")
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
        pdu = "0011000D916831..."
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
        self.assertEqual(info, "已发送：ATI")
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
        )

        self.assertTrue(ok)
        self.assertEqual(calls, [
            ("port_ui", "云端发送短信至 +8613123123123：", "normal"),
            ("port_ui", "验证码 1234", "sms"),
        ])

    def test_send_cloud_serial_command_runtime_suppresses_sms_pdu_noise(self):
        calls = []
        pdu = "0011000D916831..."

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
        )

        self.assertTrue(ok)
        self.assertEqual(calls[0], ("write", pdu))
        self.assertFalse(any(call[0] == "port_ui" for call in calls))
        self.assertFalse(any(pdu in str(call) for call in calls[1:]))
        self.assertNotIn(pdu, info)
        self.assertIn("已隐藏", info)


if __name__ == "__main__":
    unittest.main()
