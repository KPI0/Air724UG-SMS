import asyncio
import unittest

from sms_core.cloud_ws_namespace_runtime import (
    cloud_thread_main_namespace_runtime,
    cloud_ws_main_namespace_runtime,
    send_cloud_register_namespace_runtime,
    send_cloud_unregister_namespace_runtime,
    wait_cloud_login_ack_namespace_runtime,
)


class FakeTime:
    @staticmethod
    def monotonic():
        return 10.0


class FakeWebsockets:
    @staticmethod
    def connect(*args, **kwargs):
        return ("connect", args, kwargs)


class CloudWsNamespaceRuntimeTests(unittest.TestCase):
    def base_namespace(self):
        return {
            "cloud_stop_event": "stop",
            "cloud_device_authorized": False,
            "cloud_ws_conn": None,
            "cloud_connected": False,
            "cloud_ws_loop": None,
            "cloud_ws_thread": "thread",
            "cloud_ws_lock": "lock",
            "CLOUD_AUTO_UPLOAD": True,
            "CLOUD_DEVICE_SECRET": "secret",
            "CLOUD_CONTROL_ENABLED": True,
            "PORT": "COM5",
            "BAUD": 115200,
            "MODE": "Manual",
            "websockets": FakeWebsockets,
            "time": FakeTime,
            "set_cloud_auth_status_from_ack": lambda data: ("ack", data),
            "_cloud_log": lambda *args, **kwargs: ("log", args, kwargs),
            "_cloud_safe_preview": lambda value: f"safe:{value}",
            "_cloud_build_register_payload": lambda *args: ("register_payload", args),
            "_cloud_build_unregister_payload": lambda *args: ("unregister_payload", args),
            "_cloud_now_ts": lambda: 123,
            "_cloud_identity_payload": lambda: {"imei": "861"},
            "_cloud_runtime_imei": lambda: "861",
            "request_cloud_device_imei": lambda: "imei",
            "set_cloud_status": lambda *args: ("status", args),
            "_reset_cloud_serial_log_state": lambda: "reset",
            "_cloud_send_register": lambda ws: ("register", ws),
            "_cloud_wait_login_ack": lambda ws: ("ack", ws),
            "_handle_cloud_message": lambda ws, message: ("message", ws, message),
            "_schedule_cloud_sms_event_drain": lambda: "schedule_sms",
            "_cloud_ws_main": lambda url, interval: ("main", url, interval),
        }

    def test_wait_cloud_login_ack_namespace_runtime_forwards_and_sets_authorized(self):
        namespace = self.base_namespace()
        calls = []

        async def wait_runtime(ws, **kwargs):
            calls.append((ws, kwargs))
            kwargs["set_authorized"](True)
            return "waited"

        result = asyncio.run(wait_cloud_login_ack_namespace_runtime(
            namespace,
            "ws",
            timeout=3.5,
            wait_runtime=wait_runtime,
        ))

        self.assertEqual(result, "waited")
        self.assertTrue(namespace["cloud_device_authorized"])
        forwarded = calls[0][1]
        self.assertEqual(forwarded["stop_event"], "stop")
        self.assertEqual(forwarded["timeout"], 3.5)
        self.assertEqual(forwarded["monotonic"](), 10.0)

    def test_send_register_and_unregister_namespace_runtimes_forward_payload_context(self):
        namespace = self.base_namespace()
        calls = []

        async def send_runtime(ws, **kwargs):
            calls.append((ws, kwargs))
            return "sent"

        register = asyncio.run(send_cloud_register_namespace_runtime(
            namespace,
            "ws",
            send_runtime=send_runtime,
        ))
        unregister = asyncio.run(send_cloud_unregister_namespace_runtime(
            namespace,
            "ws",
            reason="disconnect",
            send_runtime=send_runtime,
        ))

        self.assertEqual((register, unregister), ("sent", "sent"))
        self.assertTrue(calls[0][1]["auto_upload"])
        self.assertEqual(calls[0][1]["serial_port"], "COM5")
        self.assertEqual(calls[1][1]["reason"], "disconnect")
        self.assertEqual(calls[1][1]["runtime_imei"](), "861")

    def test_cloud_ws_main_namespace_runtime_forwards_and_mutates_connection_state(self):
        namespace = self.base_namespace()
        calls = []

        async def ws_main_runtime(url, reconnect_interval, **kwargs):
            calls.append((url, reconnect_interval, kwargs))
            kwargs["set_ws"]("ws")
            kwargs["set_connected"](1)
            kwargs["set_authorized"]("")
            return "main"

        result = asyncio.run(cloud_ws_main_namespace_runtime(
            namespace,
            "ws://host",
            7,
            ws_main_runtime=ws_main_runtime,
        ))

        self.assertEqual(result, "main")
        self.assertEqual(namespace["cloud_ws_conn"], "ws")
        self.assertTrue(namespace["cloud_connected"])
        self.assertFalse(namespace["cloud_device_authorized"])
        forwarded = calls[0][2]
        self.assertEqual(forwarded["connect"]("url"), ("connect", ("url",), {}))
        self.assertEqual(forwarded["monotonic"](), 10.0)
        self.assertTrue(forwarded["cloud_control_enabled"]())
        self.assertEqual(forwarded["schedule_pending_sms_events"](), "schedule_sms")
        namespace["CLOUD_CONTROL_ENABLED"] = False
        self.assertFalse(forwarded["cloud_control_enabled"]())

    def test_cloud_thread_main_namespace_runtime_forwards_thread_state(self):
        namespace = self.base_namespace()
        calls = []

        result = cloud_thread_main_namespace_runtime(
            namespace,
            "ws://host",
            5,
            thread_main_runtime=lambda url, interval, **kwargs: calls.append((url, interval, kwargs)) or "thread",
        )

        self.assertEqual(result, "thread")
        forwarded = calls[0][2]
        self.assertEqual(forwarded["lock"], "lock")
        self.assertEqual(forwarded["get_thread"](), "thread")
        forwarded["set_loop"]("loop")
        forwarded["set_thread"](None)
        self.assertEqual(namespace["cloud_ws_loop"], "loop")
        self.assertIsNone(namespace["cloud_ws_thread"])


if __name__ == "__main__":
    unittest.main()
