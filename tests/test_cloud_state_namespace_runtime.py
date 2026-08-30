import asyncio
import re
import unittest
from types import SimpleNamespace

from sms_core.cloud_state_namespace_runtime import (
    cloud_auth_matches_namespace_runtime,
    cloud_check_replay_window_namespace_runtime,
    cloud_identity_payload_namespace_runtime,
    cloud_runtime_imei_namespace_runtime,
    cloud_status_payload_namespace_runtime,
    maybe_capture_cloud_device_imei_namespace_runtime,
    notify_cloud_channel_status_namespace_runtime,
    notify_cloud_identity_changed_namespace_runtime,
    request_cloud_device_imei_namespace_runtime,
    set_cloud_device_imei_namespace_runtime,
)
from sms_core.threading_runtime import WorkerThreadRegistry


class FakeTime:
    @staticmethod
    def monotonic():
        return 10.0


class FakeThreading:
    class Thread:
        pass


class CloudStateNamespaceRuntimeTests(unittest.TestCase):
    def base_namespace(self):
        return {
            "APP_VERSION": "1.0.0",
            "CLOUD_DEVICE_IMEI": "IMEI: 123456789012345",
            "CLOUD_DEVICE_SECRET": "secret",
            "cloud_imei_verified": True,
            "cloud_ws_loop": "loop",
            "cloud_ws_conn": "ws",
            "cloud_connected": True,
            "cloud_imei_query_deadline": 20.0,
            "cloud_replay_seen": {},
            "CLOUD_REPLAY_WINDOW_SECONDS": 30,
            "CLOUD_REPLAY_CACHE_MAX": 100,
            "serial_lock": "lock",
            "serial_obj": "serial",
            "SERIAL_COMMAND_THREAD_REGISTRY": WorkerThreadRegistry(),
            "PORT": "COM5",
            "BAUD": 115200,
            "MODE": "Manual",
            "IMEI_REGEX": re.compile(r"\b(\d{14,17})\b"),
            "time": FakeTime,
            "threading": FakeThreading,
            "_normalize_imei": lambda value: re.sub(r"\D", "", str(value or "")),
            "_cloud_identity_payload_core": lambda imei, version: {"imei": imei, "version": version},
            "_cloud_runtime_imei": lambda: "123456789012345",
            "_cloud_send_register": lambda ws: ("register", ws),
            "_cloud_log": lambda message: None,
            "_notify_cloud_identity_changed": lambda: "notified",
            "write_serial_command_result": lambda serial, command: ("write", serial, command),
            "_push_serial_debug": lambda line: ("debug", line),
            "_set_cloud_device_imei": lambda imei, source="": ("imei", imei, source),
            "_cloud_auth_match_result": lambda data, imei, secret: (True, ""),
            "_cloud_build_status_payload": lambda *args: ("status_payload", args),
            "_cloud_now_ts": lambda: 123,
            "_cloud_identity_payload": lambda: {"imei": "123456789012345"},
            "_cloud_reply": lambda ws, payload: ("reply", ws, payload),
            "_cloud_check_replay_window_core": lambda *args, **kwargs: SimpleNamespace(ok=True, log_message="", payload=None),
        }

    def test_runtime_imei_and_identity_payload_read_namespace(self):
        namespace = self.base_namespace()

        self.assertEqual(cloud_runtime_imei_namespace_runtime(namespace), "123456789012345")
        self.assertEqual(
            cloud_identity_payload_namespace_runtime(namespace),
            {"imei": "123456789012345", "version": "1.0.0"},
        )

        namespace["cloud_imei_verified"] = False
        self.assertEqual(cloud_runtime_imei_namespace_runtime(namespace), "")

    def test_notify_cloud_identity_changed_namespace_runtime_forwards_callbacks(self):
        namespace = self.base_namespace()
        calls = []

        result = notify_cloud_identity_changed_namespace_runtime(
            namespace,
            notify_runtime=lambda **kwargs: calls.append(kwargs) or "notified",
        )

        self.assertEqual(result, "notified")
        forwarded = calls[0]
        self.assertEqual(forwarded["get_loop"](), "loop")
        self.assertEqual(forwarded["get_ws"](), "ws")
        self.assertTrue(forwarded["is_connected"]())
        self.assertEqual(forwarded["runtime_imei"](), "123456789012345")

    def test_set_cloud_device_imei_namespace_runtime_updates_state(self):
        namespace = self.base_namespace()
        calls = []

        result = set_cloud_device_imei_namespace_runtime(
            namespace,
            "123456789012999",
            source="test",
            set_imei_runtime=lambda imei, **kwargs: (
                calls.append((imei, kwargs)),
                kwargs["set_device_imei"]("123456789012999"),
                kwargs["set_verified"](True),
                "updated",
            )[-1],
        )

        self.assertEqual(result, "updated")
        self.assertEqual(namespace["CLOUD_DEVICE_IMEI"], "123456789012999")
        self.assertTrue(namespace["cloud_imei_verified"])
        self.assertEqual(calls[0][1]["source"], "test")

    def test_request_and_capture_imei_namespace_runtime_forward_state(self):
        namespace = self.base_namespace()
        calls = []

        request_result = request_cloud_device_imei_namespace_runtime(
            namespace,
            request_runtime=lambda **kwargs: calls.append(("request", kwargs)) or "requested",
        )
        capture_result = maybe_capture_cloud_device_imei_namespace_runtime(
            namespace,
            "IMEI: 123456789012345",
            capture_runtime=lambda line, **kwargs: (
                calls.append(("capture", line, kwargs)),
                kwargs["set_query_deadline"](0.0),
                "captured",
            )[-1],
        )

        self.assertEqual(request_result, "requested")
        self.assertEqual(capture_result, "captured")
        self.assertEqual(calls[0][1]["get_serial"](), "serial")
        self.assertIs(
            calls[0][1]["thread_registry"],
            namespace["SERIAL_COMMAND_THREAD_REGISTRY"],
        )
        self.assertEqual(calls[1][2]["query_deadline"], 20.0)
        self.assertEqual(namespace["cloud_imei_query_deadline"], 0.0)
        calls[0][1]["set_query_deadline"](16.0)
        self.assertEqual(namespace["cloud_imei_query_deadline"], 16.0)

    def test_auth_matches_logs_failures(self):
        namespace = self.base_namespace()
        logs = []
        namespace["_cloud_log"] = logs.append
        namespace["_cloud_auth_match_result"] = lambda data, imei, secret: (False, "bad auth")

        self.assertFalse(cloud_auth_matches_namespace_runtime(namespace, {"nonce": "n"}))
        self.assertEqual(logs, ["bad auth"])

    def test_status_payload_namespace_runtime_forwards_serial_state(self):
        namespace = self.base_namespace()
        calls = []

        result = cloud_status_payload_namespace_runtime(
            namespace,
            status_runtime=lambda **kwargs: calls.append(kwargs) or "status",
        )

        self.assertEqual(result, "status")
        forwarded = calls[0]
        self.assertEqual(forwarded["get_serial"](), "serial")
        self.assertEqual(forwarded["timestamp"](), 123)
        self.assertEqual(forwarded["identity_payload"](), {"imei": "123456789012345"})
        self.assertEqual(forwarded["serial_port"], "COM5")

    def test_notify_cloud_channel_status_schedules_only_once_per_state(self):
        namespace = self.base_namespace()
        namespace["cloud_device_authorized"] = True
        namespace["cloud_last_channel_status"] = None
        namespace["cloud_ws_loop"] = SimpleNamespace(is_running=lambda: True)
        calls = []
        namespace["_cloud_send_payload"] = lambda ws, payload: calls.append((ws, payload)) or ("send", payload)
        namespace["asyncio"] = SimpleNamespace(run_coroutine_threadsafe=lambda coro, loop: calls.append(("scheduled", coro, loop)))

        self.assertTrue(notify_cloud_channel_status_namespace_runtime(namespace, False))
        self.assertFalse(notify_cloud_channel_status_namespace_runtime(namespace, False))
        scheduled = [item for item in calls if item[0] == "scheduled"]
        self.assertEqual(len(scheduled), 1)
        self.assertEqual(scheduled[0][1][1]["type"], "device_channel_status")

    def test_check_replay_window_namespace_runtime_replies_on_rejection(self):
        namespace = self.base_namespace()
        calls = []
        namespace["_cloud_log"] = lambda message: calls.append(("log", message))

        async def reply(ws, payload):
            calls.append(("reply", ws, payload))

        namespace["_cloud_reply"] = reply
        namespace["_cloud_check_replay_window_core"] = lambda *args, **kwargs: SimpleNamespace(
            ok=False,
            log_message="replay",
            payload={"type": "error"},
        )

        result = asyncio.run(cloud_check_replay_window_namespace_runtime(
            namespace,
            "ws",
            {"nonce": "n"},
            mark_seen=False,
        ))

        self.assertFalse(result)
        self.assertIn(("log", "replay"), calls)
        self.assertIn(("reply", "ws", {"type": "error"}), calls)


if __name__ == "__main__":
    unittest.main()
