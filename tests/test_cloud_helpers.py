import unittest

from sms_core.cloud_auth import (
    auth_match_result,
    command_action,
    command_text,
    secret_match_result,
    target_match_result,
)
from sms_core.cloud_payloads import (
    build_call_event_payload,
    build_call_recording_status_payload,
    build_register_payload,
    build_session_revoke_payload,
    build_serial_log_payload,
    build_sms_event_payload,
    build_status_payload,
    build_channel_status_payload,
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
        self.assertEqual(
            auth_status_from_ack(
                {"ok": False, "status": "waiting", "auth_status": "waiting"}
            ),
            "waiting",
        )
        self.assertEqual(
            auth_status_from_ack(
                {"ok": False, "status": "failed", "auth_status": "failed"}
            ),
            "failed",
        )

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
        self.assertEqual(register["secret"], "s3")
        self.assertEqual(register["serial_port"], "COM5")
        self.assertEqual(register["channel_type"], "desktop")
        channel_status = build_channel_status_payload(
            106,
            identity,
            False,
            False,
            4,
            "now",
        )
        self.assertEqual(channel_status["type"], "device_channel_status")
        self.assertFalse(channel_status["control_available"])
        self.assertEqual(channel_status["serial_connection_generation"], 4)

        unregister = build_unregister_payload("hidden", 101, identity, "s3", "COM5", 115200, "Auto", "now")
        self.assertEqual(unregister["event"], "offline")
        self.assertFalse(unregister["online"])
        self.assertEqual(unregister["reason"], "hidden")
        self.assertNotIn("secret", unregister)
        self.assertNotIn("password", unregister)

        revoke = build_session_revoke_payload("password_changed", 102, identity, "now")
        self.assertEqual(revoke["type"], "device_session_revoke")
        self.assertEqual(revoke["imei"], "123456789012345")
        self.assertEqual(revoke["reason"], "password_changed")
        self.assertNotIn("secret", revoke)
        self.assertNotIn("password", revoke)

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

        timed_sms = build_sms_event_payload(
            "+8613123123123 26/06/08,12:00:00+32 hello",
            "hello world",
            103,
            identity,
            "now",
            metadata={
                "sms_time": "2026-06-08 12:00:00",
                "sms_time_identity": "2026-06-08 12:00:00+32",
            },
        )
        self.assertEqual(timed_sms["sms_time"], "2026-06-08 12:00:00")
        self.assertEqual(timed_sms["sms_time_identity"], "2026-06-08 12:00:00+32")

        call = build_call_event_payload(
            "+8613123123123",
            "收到来电：来自 +8613123123123（已拦截：黑名单）",
            104,
            identity,
            "now",
            blocked=True,
            block_reason="黑名单",
        )
        self.assertEqual(call["type"], "call_event")
        self.assertEqual(call["caller"], "+8613123123123")
        self.assertTrue(call["blocked"])
        self.assertEqual(call["block_reason"], "黑名单")

        recording_status = build_call_recording_status_payload(
            "uploading",
            "rec-123",
            "10086",
            1788048000,
            3200,
            4096,
            105,
            identity,
            "now",
        )
        self.assertEqual(recording_status["type"], "call_recording_status")
        self.assertEqual(recording_status["status"], "uploading")
        self.assertEqual(recording_status["recording_id"], "rec-123")
        self.assertNotIn("secret", recording_status)

        status = build_status_payload(105, identity, True, False, "COM5", 115200, "Auto", "now")
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
