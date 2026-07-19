import asyncio
import queue
import unittest
from unittest.mock import patch

from sms_core.cloud_message_namespace_runtime import (
    drain_cloud_serial_log_queue_namespace_runtime,
    handle_cloud_message_namespace_runtime,
    reset_cloud_serial_log_state_namespace_runtime,
    schedule_cloud_serial_log_drain_namespace_runtime,
    send_cloud_serial_command_namespace_runtime,
    send_cloud_serial_log_namespace_runtime,
    send_cloud_sms_event_namespace_runtime,
)
from sms_core.threading_runtime import WorkerThreadRegistry


class FakeLoop:
    def __init__(self, running=True):
        self.running = running

    def is_running(self):
        return self.running


class CloudMessageNamespaceRuntimeTests(unittest.TestCase):
    def base_namespace(self):
        return {
            "CLOUD_SERIAL_LOG_Q": queue.Queue(),
            "CLOUD_SERIAL_LOG_DRAIN_STATE": "drain_state",
            "CLOUD_SERIAL_LOG_DRAIN_BATCH": 25,
            "cloud_ws_conn": "ws",
            "cloud_connected": True,
            "cloud_device_authorized": True,
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

    def test_send_cloud_serial_command_namespace_runtime_forwards_serial_callbacks(self):
        namespace = self.base_namespace()
        calls = []

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
        self.assertEqual(forwarded["write_command_result"]("serial", "ATI"), ("write", "serial", "ATI"))

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
