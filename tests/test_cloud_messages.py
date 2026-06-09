import unittest

from sms_core.cloud_messages import (
    CLOUD_AUTH_FAILED_MESSAGE,
    CLOUD_UNAUTHORIZED_MESSAGE,
    JSON_OBJECT_REQUIRED_MESSAGE,
    attach_cloud_task_ids,
    cloud_action_kind,
    cloud_auth_failed_payload,
    cloud_pong_payload,
    cloud_send_at_result_payload,
    cloud_unauthorized_payload,
    cloud_unknown_command_payload,
    cloud_window_result_payload,
    dispatch_cloud_action,
    is_cloud_auth_ack_type,
    parse_cloud_message,
)


class CloudMessageTests(unittest.TestCase):
    def test_parse_cloud_message_accepts_bytes_and_extracts_action(self):
        incoming, error = parse_cloud_message(
            b'{"cmd":"AT","task_id":"task-1","target_imei":"123","secret":"s"}'
        )

        self.assertIsNone(error)
        self.assertEqual(incoming.raw, '{"cmd":"AT","task_id":"task-1","target_imei":"123","secret":"s"}')
        self.assertEqual(incoming.task_id, "task-1")
        self.assertEqual(incoming.action, "cmd")

    def test_parse_cloud_message_rejects_invalid_json(self):
        incoming, error = parse_cloud_message("{bad json")

        self.assertIsNone(incoming.data)
        self.assertEqual(error["type"], "error")
        self.assertFalse(error["ok"])
        self.assertEqual(error["message"], JSON_OBJECT_REQUIRED_MESSAGE)

    def test_parse_cloud_message_rejects_non_object_json(self):
        incoming, error = parse_cloud_message('["not", "object"]')

        self.assertIsNone(incoming.data)
        self.assertEqual(error["message"], JSON_OBJECT_REQUIRED_MESSAGE)

    def test_attach_cloud_task_ids_preserves_existing_values(self):
        payload = attach_cloud_task_ids(
            {"type": "send_at_result", "task_id": "existing"},
            "task-2",
        )

        self.assertEqual(payload["task_id"], "existing")
        self.assertEqual(payload["command_task_id"], "task-2")

    def test_cloud_action_kind_groups_supported_actions(self):
        self.assertEqual(cloud_action_kind("heartbeat"), "ping")
        self.assertEqual(cloud_action_kind("get_status"), "status")
        self.assertEqual(cloud_action_kind("command"), "send_at")
        self.assertEqual(cloud_action_kind("show_window"), "show_window")
        self.assertEqual(cloud_action_kind("hide_window"), "hide_window")
        self.assertEqual(cloud_action_kind("nope"), "unknown")

    def test_cloud_payload_helpers(self):
        self.assertTrue(is_cloud_auth_ack_type("device_auth_result"))
        self.assertFalse(is_cloud_auth_ack_type("cmd"))

        self.assertEqual(cloud_unauthorized_payload()["message"], CLOUD_UNAUTHORIZED_MESSAGE)
        self.assertEqual(cloud_auth_failed_payload()["message"], CLOUD_AUTH_FAILED_MESSAGE)
        self.assertEqual(cloud_pong_payload("now", 123), {
            "type": "pong",
            "ok": True,
            "time": "now",
            "timestamp": 123,
        })
        self.assertEqual(cloud_send_at_result_payload(1, "ok"), {
            "type": "send_at_result",
            "ok": True,
            "message": "ok",
        })
        self.assertEqual(cloud_window_result_payload("show_window"), {
            "type": "show_window_result",
            "ok": True,
        })
        self.assertIn("未知云端指令", cloud_unknown_command_payload("")["message"])

    async def _dispatch(self, action, data=None, calls=None):
        calls = calls if calls is not None else []

        async def send_serial(command):
            calls.append(("send", command))
            return True, f"sent {command}"

        return await dispatch_cloud_action(
            action,
            data or {},
            time_text=lambda: "now",
            timestamp=lambda: 123,
            status_payload=lambda: {"type": "status"},
            send_serial_command=send_serial,
            show_window=lambda: calls.append(("show",)),
            hide_window=lambda: calls.append(("hide",)),
            log=lambda message: calls.append(("log", message)),
        )

    def test_dispatch_cloud_action_ping_status_and_unknown(self):
        import asyncio

        self.assertEqual(asyncio.run(self._dispatch("ping")), {
            "type": "pong",
            "ok": True,
            "time": "now",
            "timestamp": 123,
        })
        self.assertEqual(asyncio.run(self._dispatch("status")), {"type": "status"})
        self.assertEqual(asyncio.run(self._dispatch("missing"))["type"], "error")

    def test_dispatch_cloud_action_send_at_and_window(self):
        import asyncio

        calls = []
        payload = asyncio.run(self._dispatch("send_at", {"cmd": "ATI"}, calls))
        self.assertEqual(payload, {"type": "send_at_result", "ok": True, "message": "sent ATI"})
        self.assertIn(("log", "云端下发指令：ATI"), calls)
        self.assertIn(("send", "ATI"), calls)

        calls = []
        self.assertEqual(asyncio.run(self._dispatch("show_window", calls=calls)), {
            "type": "show_window_result",
            "ok": True,
        })
        self.assertEqual(calls, [("show",)])

        calls = []
        self.assertEqual(asyncio.run(self._dispatch("hide_window", calls=calls)), {
            "type": "hide_window_result",
            "ok": True,
        })
        self.assertEqual(calls, [("hide",)])


if __name__ == "__main__":
    unittest.main()
