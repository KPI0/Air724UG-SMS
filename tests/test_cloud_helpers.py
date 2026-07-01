import unittest

from sms_core.cloud_auth import (
    auth_match_result,
    command_action,
    command_text,
    secret_match_result,
    target_match_result,
)
from sms_core.cloud_payloads import (
    build_register_payload,
    build_serial_log_payload,
    build_sms_event_payload,
    build_status_payload,
    build_unregister_payload,
    identity_payload,
    truncate_serial_log_text,
)
from sms_core.cloud_protocol import auth_status_from_ack
from sms_core.cloud_security import check_replay_window, repeat_filter
from sms_core.status_text import format_connected_status, format_connecting_status


class CloudHelperTests(unittest.TestCase):
    def test_cloud_auth_status_from_ack(self):
        self.assertEqual(auth_status_from_ack({"auth_status": "ok"}), "authorized")
        self.assertEqual(auth_status_from_ack({"status": "auth_failed"}), "failed")
        self.assertEqual(auth_status_from_ack({"message": "密码错误"}), "failed")
        self.assertEqual(auth_status_from_ack({"status": "pending"}), "waiting")

    def test_cloud_command_auth_helpers(self):
        self.assertEqual(command_action({"cmd": "AT"}), "cmd")
        self.assertEqual(command_action({"action": "PING"}), "ping")
        self.assertEqual(command_text({"data": "ATI"}), "ATI")

        self.assertEqual(secret_match_result({"secret": "s3"}, "s3"), (True, ""))
        ok, reason = secret_match_result({"secret": "bad"}, "s3")
        self.assertFalse(ok)
        self.assertTrue(reason)

        self.assertEqual(
            target_match_result({"target_imei": ["000", "123456789012345"]}, "123456789012345"),
            (True, ""),
        )
        ok, reason = target_match_result({"target_imei": "000"}, "123456789012345")
        self.assertFalse(ok)
        self.assertTrue(reason)

        self.assertEqual(
            auth_match_result({"target_imei": "123456789012345", "password": "s3"}, "123456789012345", "s3"),
            (True, ""),
        )

    def test_cloud_payload_builders(self):
        identity = identity_payload("123456789012345", "3.6.6", device_name="host-a")
        self.assertEqual(identity["imei"], "123456789012345")
        self.assertEqual(identity["device_name"], "host-a")

        register = build_register_payload(True, 100, identity, "s3", "COM5", 115200, "Auto", "now")
        self.assertEqual(register["event"], "register")
        self.assertTrue(register["public"])
        self.assertEqual(register["serial_port"], "COM5")

        unregister = build_unregister_payload("hidden", 101, identity, "s3", "COM5", 115200, "Auto", "now")
        self.assertEqual(unregister["event"], "offline")
        self.assertFalse(unregister["online"])
        self.assertEqual(unregister["reason"], "hidden")

        self.assertEqual(truncate_serial_log_text("x" * 5, limit=3), "xxx...")
        serial_log = build_serial_log_payload("AT", 102, identity, "COM5", 115200, "now")
        self.assertEqual(serial_log["type"], "log")
        self.assertEqual(serial_log["raw"], "AT")
        self.assertIn("AT", serial_log["data"])

        sms = build_sms_event_payload(
            "+8613123123123 26/06/08,12:00:00+32 hello",
            "hello world",
            103,
            identity,
            "now",
        )
        self.assertEqual(sms["type"], "sms_event")
        self.assertEqual(sms["phone"], "+8613123123123")
        self.assertEqual(sms["content"], "hello world")
        self.assertIn("hello world", sms["message"])

        traced_sms = build_sms_event_payload(
            "+8613123123123 26/06/08,12:00:00+32 hello",
            "hello world",
            103,
            identity,
            "now",
            metadata={"message_trace_id": "abc123def456"},
        )
        self.assertEqual(traced_sms["message_trace_id"], "abc123def456")
        self.assertEqual(traced_sms["trace_id"], "abc123def456")

        status = build_status_payload(104, identity, True, False, "COM5", 115200, "Auto", "now")
        self.assertTrue(status["cloud_connected"])
        self.assertFalse(status["serial_connected"])
        self.assertEqual(status["serial_mode"], "Auto")

    def test_repeat_filter_limits_repeated_messages(self):
        state = {}
        self.assertEqual(repeat_filter(state, "same", 10, 3, now=1), "same")
        self.assertEqual(repeat_filter(state, "same", 10, 3, now=2), "same")
        self.assertEqual(repeat_filter(state, "same", 10, 3, now=3), "same（后续同类消息已忽略）")
        self.assertIsNone(repeat_filter(state, "same", 10, 3, now=4))
        self.assertEqual(repeat_filter(state, "same", 10, 3, now=20), "same")


    def test_check_replay_window_validates_timestamp_and_nonce(self):
        cache = {}

        missing = check_replay_window(
            {"task_id": "t1"},
            cache,
            now_ts=100,
            window_seconds=60,
            max_size=10,
        )
        self.assertFalse(missing.ok)
        self.assertEqual(missing.payload["task_id"], "t1")
        self.assertIn("timestamp", missing.payload["message"])

        expired = check_replay_window(
            {"timestamp": 1, "nonce": "n1"},
            cache,
            now_ts=100,
            window_seconds=60,
            max_size=10,
        )
        self.assertFalse(expired.ok)
        self.assertEqual(expired.payload["type"], "error")
        self.assertIn("timestamp=1", expired.log_message)

        first = check_replay_window(
            {"timestamp": 100, "nonce": "n1"},
            cache,
            now_ts=100,
            window_seconds=60,
            max_size=10,
        )
        self.assertTrue(first.ok)
        self.assertIn("nonce:n1", cache)

        duplicate = check_replay_window(
            {"timestamp": 100, "nonce": "n1"},
            cache,
            now_ts=101,
            window_seconds=60,
            max_size=10,
        )
        self.assertFalse(duplicate.ok)
        self.assertEqual(duplicate.payload["type"], "error")
        self.assertIn("timestamp=100", duplicate.log_message)

    def test_status_text_helpers(self):
        self.assertEqual(format_connected_status("COM5"), "🟢 已连接：COM5")
        self.assertEqual(format_connected_status(""), "🟢 已连接")
        self.assertEqual(format_connecting_status("COM5"), "🟡 连接中：COM5")
        self.assertEqual(format_connecting_status(None), "🟡 连接中")


if __name__ == "__main__":
    unittest.main()
