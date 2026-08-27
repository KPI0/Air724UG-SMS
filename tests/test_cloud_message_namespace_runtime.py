import asyncio
import queue
import threading
import unittest
from unittest.mock import patch

from sms_core.cloud_message_namespace_runtime import (
    apply_cloud_modem_health_namespace_runtime,
    drain_cloud_serial_log_queue_namespace_runtime,
    handle_cloud_message_namespace_runtime,
    reset_cloud_serial_log_state_namespace_runtime,
    schedule_cloud_serial_log_drain_namespace_runtime,
    send_cloud_call_event_namespace_runtime,
    send_cloud_serial_command_namespace_runtime,
    send_cloud_serial_log_namespace_runtime,
    send_cloud_sms_event_namespace_runtime,
)
from sms_core.serial_sender import AtCommandResponseCoordinator, SerialCommandResult
from sms_core.cloud_modem_health import CloudModemHealthState
from sms_core.threading_runtime import WorkerThreadRegistry


class FakeLoop:
    def __init__(self, running=True):
        self.running = running

    def is_running(self):
        return self.running


class CallbackSerial:
    def __init__(self, on_write=None):
        self.is_open = True
        self.on_write = on_write
        self.writes = []

    def write(self, payload):
        self.writes.append(payload)
        if self.on_write:
            self.on_write(payload)

    def flush(self):
        return None


class CloudMessageNamespaceRuntimeTests(unittest.TestCase):
    def base_namespace(self):
        return {
            "CLOUD_SERIAL_LOG_Q": queue.Queue(),
            "CLOUD_SERIAL_LOG_DRAIN_STATE": "drain_state",
            "CLOUD_SERIAL_LOG_DRAIN_BATCH": 25,
            "cloud_ws_conn": "ws",
            "cloud_connected": True,
            "cloud_device_authorized": True,
            "CLOUD_SENSITIVE_COMMAND_PERMISSIONS": {"sms": False},
            "cloud_ws_loop": FakeLoop(),
            "PORT": "COM5",
            "BAUD": 115200,
            "serial_lock": "lock",
            "serial_obj": "serial",
            "SERIAL_COMMAND_THREAD_REGISTRY": WorkerThreadRegistry(),
            "asyncio": asyncio,
            "_cloud_build_serial_log_payload": lambda text, ts, identity, port, baud: {
                "text": text,
                "ts": ts,
                **identity,
                "port": port,
                "baud": baud,
            },
            "_cloud_build_sms_event_payload": lambda head, body, ts, identity: {
                "head": head,
                "body": body,
                "ts": ts,
                **identity,
            },
            "_cloud_build_call_event_payload": lambda caller, message, ts, identity, **metadata: {
                "caller": caller,
                "message": message,
                "ts": ts,
                **identity,
                **metadata,
            },
            "_cloud_now_ts": lambda: 123,
            "_cloud_identity_payload": lambda: {"imei": "861"},
            "_cloud_runtime_imei": lambda: "861",
            "_schedule_cloud_serial_log_drain": lambda loop, ws: ("schedule", loop, ws),
            "_cloud_drain_serial_log_queue": lambda ws: ("drain", ws),
            "_cloud_send_payload": lambda ws, payload: ("send_payload", ws, payload),
            "_cloud_log": lambda message: ("log", message),
            "_push_serial_debug": lambda line: ("debug", line),
            "write_serial_command_result": lambda serial, command: ("write", serial, command),
            "set_cloud_auth_status_from_ack": lambda data: ("ack", data),
            "set_cloud_status": lambda *args: ("status", args),
            "_cloud_check_replay_window": lambda ws, data, mark_seen=True: ("replay", ws, data, mark_seen),
            "_cloud_auth_matches": lambda data: True,
            "_cloud_send_status_payload": lambda: {"status": "ok"},
            "_cloud_send_serial_command": lambda command, command_data=None: (
                "serial_command",
                command,
                command_data,
            ),
            "show_window": lambda: "show",
            "hide_window": lambda: "hide",
            "_cloud_reply": _async_reply,
        }

    def test_serial_log_state_namespace_runtimes_forward_queue_state(self):
        namespace = self.base_namespace()
        calls = []

        with patch(
            "sms_core.cloud_message_namespace_runtime.reset_cloud_serial_log_state",
            lambda log_queue, state: calls.append(("reset", log_queue, state)) or "reset",
        ), patch(
            "sms_core.cloud_message_namespace_runtime.drain_cloud_serial_log_queue",
            lambda ws, **kwargs: calls.append(("drain", ws, kwargs)) or _async_value("drained"),
        ), patch(
            "sms_core.cloud_message_namespace_runtime.schedule_cloud_serial_log_drain",
            lambda loop, ws, **kwargs: calls.append(("schedule", loop, ws, kwargs)) or "scheduled",
        ):
            self.assertEqual(reset_cloud_serial_log_state_namespace_runtime(namespace), "reset")
            self.assertEqual(
                asyncio.run(drain_cloud_serial_log_queue_namespace_runtime(namespace, "ws")),
                "drained",
            )
            self.assertEqual(schedule_cloud_serial_log_drain_namespace_runtime(namespace, "loop", "ws"), "scheduled")

        self.assertEqual(calls[0], ("reset", namespace["CLOUD_SERIAL_LOG_Q"], "drain_state"))
        self.assertEqual(calls[1][0:2], ("drain", "ws"))
        self.assertEqual(calls[1][2]["batch_size"], 25)
        self.assertTrue(calls[1][2]["is_current_connection"]("ws"))
        self.assertTrue(calls[1][2]["is_connected"]())
        self.assertEqual(calls[2][0:3], ("schedule", "loop", "ws"))
        self.assertEqual(calls[2][3]["state"], "drain_state")

    def test_send_cloud_serial_log_namespace_runtime_builds_payload_context(self):
        namespace = self.base_namespace()
        calls = []

        result = send_cloud_serial_log_namespace_runtime(
            namespace,
            "AT",
            send_runtime=lambda line, **kwargs: calls.append((line, kwargs)) or "queued",
        )

        self.assertEqual(result, "queued")
        line, forwarded = calls[0]
        self.assertEqual(line, "AT")
        self.assertTrue(forwarded["authorized"])
        self.assertEqual(forwarded["get_loop"](), namespace["cloud_ws_loop"])
        self.assertEqual(forwarded["get_ws"](), "ws")
        self.assertEqual(forwarded["runtime_imei"](), "861")
        self.assertEqual(forwarded["build_payload"]("AT")["port"], "COM5")
        self.assertEqual(forwarded["schedule_drain"]("loop", "ws"), ("schedule", "loop", "ws"))

    def test_send_cloud_sms_event_namespace_runtime_forwards_connection_context(self):
        namespace = self.base_namespace()
        calls = []

        result = send_cloud_sms_event_namespace_runtime(
            namespace,
            "head",
            "body",
            {"message_trace_id": "trace-1"},
            send_runtime=lambda head, body, **kwargs: calls.append((head, body, kwargs)) or "scheduled",
        )

        self.assertEqual(result, "scheduled")
        head, body, forwarded = calls[0]
        self.assertEqual((head, body), ("head", "body"))
        self.assertTrue(forwarded["authorized"])
        self.assertEqual(forwarded["metadata"], {"message_trace_id": "trace-1"})
        self.assertEqual(forwarded["timestamp"](), 123)
        self.assertEqual(forwarded["identity_payload"](), {"imei": "861"})
        self.assertEqual(forwarded["send_payload"]("ws", {"x": 1}), ("send_payload", "ws", {"x": 1}))

    def test_send_cloud_call_event_namespace_runtime_forwards_shared_queue_context(self):
        namespace = self.base_namespace()
        calls = []

        result = send_cloud_call_event_namespace_runtime(
            namespace,
            "10086",
            "incoming",
            blocked=True,
            block_reason="blacklist",
            send_runtime=lambda caller, message, **kwargs: (
                calls.append((caller, message, kwargs)) or "scheduled"
            ),
        )

        self.assertEqual(result, "scheduled")
        caller, message, forwarded = calls[0]
        self.assertEqual((caller, message), ("10086", "incoming"))
        self.assertTrue(forwarded["blocked"])
        self.assertEqual(forwarded["block_reason"], "blacklist")
        self.assertTrue(forwarded["authorized"])
        self.assertEqual(forwarded["timestamp"](), 123)
        self.assertEqual(forwarded["identity_payload"](), {"imei": "861"})

    def test_send_cloud_serial_command_namespace_runtime_forwards_serial_callbacks(self):
        namespace = self.base_namespace()
        calls = []

        with patch(
            "sms_core.cloud_message_namespace_runtime.write_serial_command_sequence_confirmed_locked",
            return_value=SerialCommandResult(True),
        ) as confirmed:
            result = send_cloud_serial_command_namespace_runtime(
                namespace,
                "ATI",
                send_runtime=lambda command, **kwargs: calls.append((command, kwargs)) or (True, "ok"),
            )

            self.assertEqual(result, (True, "ok"))
            command, forwarded = calls[0]
            self.assertEqual(command, "ATI")
            self.assertEqual(forwarded["serial_lock"], "lock")
            self.assertEqual(forwarded["get_serial"](), "serial")
            self.assertEqual(
                forwarded["write_command_result"]("serial", "ATI"),
                SerialCommandResult(True),
            )
            self.assertEqual(forwarded["allow_sensitive_commands"], {"sms": False})

        confirmed.assert_called_once()
        confirmed_args = confirmed.call_args.args
        self.assertEqual(confirmed_args[0], "lock")
        self.assertTrue(callable(confirmed_args[1]))
        self.assertEqual(confirmed_args[1](), "serial")
        self.assertEqual(confirmed_args[2], ("ATI",))
        self.assertEqual(confirmed.call_args.kwargs["response_timeout"], 10.0)

    def test_cloud_at_reports_success_only_after_modem_ok(self):
        namespace = self.base_namespace()
        coordinator = AtCommandResponseCoordinator()
        serial_obj = CallbackSerial(
            lambda _payload: coordinator.observe_line("[I]-[ril.proatc] OK")
        )
        namespace.update({
            "serial_lock": threading.RLock(),
            "serial_obj": serial_obj,
            "SERIAL_COMMAND_RESPONSE_COORDINATOR": coordinator,
            "write_serial_command_result": lambda *_args: self.fail(
                "云端 AT 不得回退到只确认串口写入的函数"
            ),
        })

        result = send_cloud_serial_command_namespace_runtime(namespace, "ATI")

        self.assertEqual(result, (True, "执行成功：ATI"))
        self.assertEqual(serial_obj.writes, [b"ATI\r\n"])

    def test_cloud_at_reports_modem_error_as_failure(self):
        namespace = self.base_namespace()
        coordinator = AtCommandResponseCoordinator()
        serial_obj = CallbackSerial(
            lambda _payload: coordinator.observe_line(
                "[I]-[ril.proatc] +CME ERROR: 3"
            )
        )
        namespace.update({
            "serial_lock": threading.RLock(),
            "serial_obj": serial_obj,
            "SERIAL_COMMAND_RESPONSE_COORDINATOR": coordinator,
        })

        ok, info = send_cloud_serial_command_namespace_runtime(namespace, "AT+RESET")

        self.assertFalse(ok)
        self.assertEqual(info, "+CME ERROR: 3")
        self.assertNotIn("AT+RESET", info)
        self.assertEqual(serial_obj.writes, [b"AT+RESET\r\n"])

    def test_cloud_at_reports_timeout_instead_of_write_success(self):
        namespace = self.base_namespace()
        serial_obj = CallbackSerial()
        namespace.update({
            "serial_lock": threading.RLock(),
            "serial_obj": serial_obj,
            "SERIAL_COMMAND_RESPONSE_COORDINATOR": AtCommandResponseCoordinator(),
        })

        with patch(
            "sms_core.cloud_message_namespace_runtime.AT_COMMAND_RESPONSE_DEFAULT_TIMEOUT",
            0,
        ):
            ok, info = send_cloud_serial_command_namespace_runtime(namespace, "ATI")

        self.assertFalse(ok)
        self.assertIn("超时", info)
        self.assertEqual(serial_obj.writes, [b"ATI\r\n"])

    def test_three_consecutive_timeouts_safely_close_and_wake_serial(self):
        namespace = self.base_namespace()
        calls = []
        namespace.update({
            "CLOUD_MODEM_HEALTH": CloudModemHealthState(reconnect_threshold=3),
            "safe_close_serial": lambda: calls.append("close"),
            "serial_wakeup_event": type(
                "WakeEvent",
                (),
                {"set": lambda _self: calls.append("wake")},
            )(),
            "_cloud_log": lambda message, **kwargs: calls.append((message, kwargs)),
        })

        first = apply_cloud_modem_health_namespace_runtime(
            namespace,
            (False, "等待 Modem 指令响应超时"),
        )
        second = apply_cloud_modem_health_namespace_runtime(
            namespace,
            (False, "等待 Modem 指令响应超时"),
        )
        third = apply_cloud_modem_health_namespace_runtime(
            namespace,
            (False, "等待 Modem 指令响应超时"),
        )

        self.assertFalse(first[2]["modem_unresponsive"])
        self.assertFalse(second[2]["modem_unresponsive"])
        self.assertTrue(third[2]["modem_unresponsive"])
        self.assertEqual(calls[-2:], ["close", "wake"])

    def test_handle_cloud_message_namespace_runtime_forwards_and_mutates_state(self):
        namespace = self.base_namespace()
        calls = []

        async def handle_runtime(message, **kwargs):
            calls.append((message, kwargs))
            kwargs["set_authorized"](False)
            await kwargs["reply"]({"type": "pong"})
            return await kwargs["send_serial_command"]("ATI")

        result = asyncio.run(handle_cloud_message_namespace_runtime(
            namespace,
            "ws",
            "message",
            handle_runtime=handle_runtime,
        ))

        self.assertEqual(result, ("serial_command", "ATI", None))
        self.assertFalse(namespace["cloud_device_authorized"])
        message, forwarded = calls[0]
        self.assertEqual(message, "message")
        self.assertFalse(forwarded["is_authorized"]())
        self.assertEqual(forwarded["status_payload"](), {"status": "ok"})
        self.assertEqual(forwarded["show_window"](), "show")
        self.assertEqual(forwarded["hide_window"](), "hide")
        self.assertEqual(namespace["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot(), ())


async def _async_value(value):
    return value


async def _async_reply(ws, payload):
    return "reply", ws, payload


if __name__ == "__main__":
    unittest.main()
