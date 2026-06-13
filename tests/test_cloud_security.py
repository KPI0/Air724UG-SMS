import unittest
import time

from sms_core.cloud_security import (
    CloudReplayCheckResult,
    safe_preview,
    read_unix_timestamp,
    replay_key,
    replay_error_payload,
    check_replay_window,
    prune_replay_cache,
    repeat_filter,
)


class CloudSecurityTests(unittest.TestCase):
    def test_safe_preview_masks_secret_fields(self):
        """Sensitive fields should be masked in preview output."""
        payload = '{"device_secret": "abc123", "message": "hello", "token": "xyz"}'
        preview = safe_preview(payload)

        self.assertIn("***", preview)
        self.assertNotIn("abc123", preview)
        self.assertNotIn("xyz", preview)
        self.assertIn("hello", preview)

    def test_safe_preview_handles_non_json(self):
        """Non-JSON input should be returned as-is."""
        raw = "plain text message"
        preview = safe_preview(raw)
        self.assertEqual(preview, raw)

    def test_safe_preview_truncates_long_text(self):
        """Long previews should be truncated."""
        long_text = "x" * 600
        preview = safe_preview(long_text, limit=100)
        self.assertEqual(len(preview), 103)  # 100 + "..."
        self.assertTrue(preview.endswith("..."))

    def test_read_unix_timestamp_accepts_multiple_field_names(self):
        """Should read timestamp from various field names."""
        for field in ["timestamp", "ts", "unix_time", "time"]:
            ts, raw = read_unix_timestamp({field: "1234567890"})
            self.assertEqual(ts, 1234567890)
            self.assertEqual(raw, "1234567890")

    def test_read_unix_timestamp_handles_milliseconds(self):
        """Millisecond timestamps should be converted to seconds."""
        ts, raw = read_unix_timestamp({"timestamp": 1234567890123})
        self.assertEqual(ts, 1234567890)

    def test_read_unix_timestamp_returns_none_for_invalid(self):
        """Invalid timestamps should return None."""
        ts, raw = read_unix_timestamp({"timestamp": "not-a-number"})
        self.assertIsNone(ts)
        self.assertEqual(raw, "not-a-number")

    def test_read_unix_timestamp_returns_none_for_missing(self):
        """Missing timestamp should return None."""
        ts, raw = read_unix_timestamp({})
        self.assertIsNone(ts)
        self.assertIsNone(raw)

    def test_replay_key_uses_nonce_when_present(self):
        """Replay key should prefer nonce over fingerprint."""
        key = replay_key({"nonce": "abc123", "command": "reboot"}, 1234567890)
        self.assertEqual(key, "nonce:abc123")

    def test_replay_key_generates_fingerprint_without_nonce(self):
        """Without nonce, should generate fingerprint from payload."""
        key1 = replay_key({"command": "reboot", "type": "command"}, 1234567890)
        key2 = replay_key({"command": "reboot", "type": "command"}, 1234567890)
        key3 = replay_key({"command": "status", "type": "command"}, 1234567890)

        self.assertTrue(key1.startswith("fp:"))
        self.assertEqual(key1, key2)  # Same payload = same fingerprint
        self.assertNotEqual(key1, key3)  # Different payload = different fingerprint

    def test_replay_error_payload_includes_task_id(self):
        """Error payload should include task_id when provided."""
        payload = replay_error_payload("test error", "task123")
        self.assertEqual(payload["type"], "error")
        self.assertFalse(payload["ok"])
        self.assertEqual(payload["message"], "test error")
        self.assertEqual(payload["task_id"], "task123")
        self.assertEqual(payload["command_task_id"], "task123")

    def test_replay_error_payload_works_without_task_id(self):
        """Error payload should work without task_id."""
        payload = replay_error_payload("test error")
        self.assertEqual(payload["type"], "error")
        self.assertNotIn("task_id", payload)

    def test_check_replay_window_rejects_missing_timestamp(self):
        """Should reject payloads without timestamp."""
        result = check_replay_window({}, {}, 1234567890, 60, 1000)

        self.assertFalse(result.ok)
        self.assertIn("时间戳", result.log_message)
        self.assertEqual(result.payload["type"], "error")

    def test_check_replay_window_rejects_expired_timestamp(self):
        """Should reject payloads with expired timestamp."""
        now = 1234567890
        cache = {}

        # Timestamp 120 seconds old (window is 60 seconds)
        result = check_replay_window(
            {"timestamp": now - 120},
            cache,
            now,
            window_seconds=60,
            max_size=1000
        )

        self.assertFalse(result.ok)
        self.assertIn("时间戳超时", result.log_message)

    def test_check_replay_window_accepts_fresh_timestamp(self):
        """Should accept payloads within time window."""
        now = 1234567890
        cache = {}

        result = check_replay_window(
            {"timestamp": now - 30, "nonce": "unique1"},
            cache,
            now,
            window_seconds=60,
            max_size=1000
        )

        self.assertTrue(result.ok)

    def test_check_replay_window_detects_duplicate_nonce(self):
        """Should detect and reject duplicate nonce."""
        now = 1234567890
        cache = {}

        # First request
        result1 = check_replay_window(
            {"timestamp": now, "nonce": "abc123"},
            cache,
            now,
            window_seconds=60,
            max_size=1000
        )
        self.assertTrue(result1.ok)

        # Replay attack with same nonce
        result2 = check_replay_window(
            {"timestamp": now, "nonce": "abc123"},
            cache,
            now,
            window_seconds=60,
            max_size=1000
        )
        self.assertFalse(result2.ok)
        self.assertIn("重复", result2.log_message)

    def test_check_replay_window_detects_duplicate_fingerprint(self):
        """Should detect replay by command fingerprint."""
        now = 1234567890
        cache = {}

        payload = {"timestamp": now, "command": "reboot", "type": "command"}

        # First request
        result1 = check_replay_window(payload, cache, now, 60, 1000)
        self.assertTrue(result1.ok)

        # Replay attack
        result2 = check_replay_window(payload, cache, now, 60, 1000)
        self.assertFalse(result2.ok)

    def test_check_replay_window_respects_mark_seen_flag(self):
        """mark_seen=False should not record in cache."""
        now = 1234567890
        cache = {}

        payload = {"timestamp": now, "nonce": "test"}

        # Check without marking seen
        result1 = check_replay_window(payload, cache, now, 60, 1000, mark_seen=False)
        self.assertTrue(result1.ok)
        self.assertEqual(len(cache), 0)

        # Should still accept (not in cache)
        result2 = check_replay_window(payload, cache, now, 60, 1000, mark_seen=False)
        self.assertTrue(result2.ok)

    def test_prune_replay_cache_removes_expired(self):
        """Should remove expired entries from cache."""
        now = 1234567890
        cache = {
            "key1": now - 100,  # Expired (>60s ago)
            "key2": now - 30,   # Fresh
            "key3": now - 80,   # Expired
        }

        prune_replay_cache(cache, now, window_seconds=60, max_size=1000)

        self.assertNotIn("key1", cache)
        self.assertIn("key2", cache)
        self.assertNotIn("key3", cache)

    def test_prune_replay_cache_enforces_max_size(self):
        """Should enforce maximum cache size."""
        now = 1234567890
        cache = {
            f"key{i}": now - i for i in range(10)
        }

        prune_replay_cache(cache, now, window_seconds=1000, max_size=5)

        self.assertLessEqual(len(cache), 5)

    def test_prune_replay_cache_handles_errors_gracefully(self):
        """Should clear cache on errors."""
        cache = {"bad": "not-an-int"}
        prune_replay_cache(cache, 1234567890, 60, 1000)
        self.assertEqual(len(cache), 0)

    def test_repeat_filter_allows_first_messages(self):
        """Should pass through first occurrences."""
        state = {}

        msg1 = repeat_filter(state, "error A", reset_seconds=10, repeat_limit=3, now=100)
        msg2 = repeat_filter(state, "error A", reset_seconds=10, repeat_limit=3, now=101)

        self.assertEqual(msg1, "error A")
        self.assertEqual(msg2, "error A")

    def test_repeat_filter_marks_limit_message(self):
        """Should mark the limit-th message."""
        state = {}

        repeat_filter(state, "error A", reset_seconds=10, repeat_limit=3, now=100)
        repeat_filter(state, "error A", reset_seconds=10, repeat_limit=3, now=101)
        msg3 = repeat_filter(state, "error A", reset_seconds=10, repeat_limit=3, now=102)

        self.assertIn("后续同类消息已忽略", msg3)

    def test_repeat_filter_suppresses_after_limit(self):
        """Should suppress messages after limit."""
        state = {}

        for i in range(3):
            repeat_filter(state, "error A", reset_seconds=10, repeat_limit=3, now=100+i)

        msg4 = repeat_filter(state, "error A", reset_seconds=10, repeat_limit=3, now=104)
        self.assertIsNone(msg4)

    def test_repeat_filter_resets_after_timeout(self):
        """Should reset count after reset_seconds."""
        state = {}

        # First occurrence
        msg1 = repeat_filter(state, "error A", reset_seconds=10, repeat_limit=3, now=100)
        self.assertEqual(msg1, "error A")

        # After timeout, should reset
        msg2 = repeat_filter(state, "error A", reset_seconds=10, repeat_limit=3, now=120)
        self.assertEqual(msg2, "error A")


    def test_read_unix_timestamp_handles_zero_timestamp(self):
        """Should handle timestamp=0 (epoch start) correctly."""
        ts, raw = read_unix_timestamp({"timestamp": 0})
        self.assertEqual(ts, 0)
        self.assertEqual(raw, 0)

    def test_read_unix_timestamp_prefers_first_field(self):
        """Should use 'timestamp' field even when others present."""
        # When timestamp=0, should not fallback to ts field
        ts, raw = read_unix_timestamp({"timestamp": 0, "ts": 1234567890})
        self.assertEqual(ts, 0)

if __name__ == "__main__":
    unittest.main()
