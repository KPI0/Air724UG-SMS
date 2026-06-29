import unittest
import os
import tempfile
import hashlib
from datetime import datetime

from sms_core.sms_processing import (
    build_sms_ui_display_lines,
    sms_keyword_hit,
    build_unmatched_sms_log_entries,
    parse_sms_callback_metadata,
    repeat_count_message,
    sms_message_id,
    process_pending_sms,
)


class SmsProcessingTests(unittest.TestCase):
    def test_sms_keyword_hit_matches_keyword(self):
        """Should return True when message contains keyword."""
        self.assertTrue(sms_keyword_hit("Hello world", ["hello"]))
        self.assertTrue(sms_keyword_hit("Important notice", ["important"]))

    def test_sms_keyword_hit_case_insensitive(self):
        """Should perform case-insensitive matching."""
        self.assertTrue(sms_keyword_hit("HELLO WORLD", ["hello"]))
        self.assertTrue(sms_keyword_hit("hello world", ["HELLO"]))

    def test_sms_keyword_hit_returns_true_for_empty_keywords(self):
        """Should return True when keywords list is empty (match all)."""
        self.assertTrue(sms_keyword_hit("Any message", []))
        self.assertTrue(sms_keyword_hit("Any message", None))

    def test_sms_keyword_hit_returns_false_for_no_match(self):
        """Should return False when no keyword matches."""
        self.assertFalse(sms_keyword_hit("Hello world", ["urgent", "important"]))

    def test_sms_keyword_hit_matches_partial_word(self):
        """Should match substring (not just whole words)."""
        self.assertTrue(sms_keyword_hit("verification code", ["verif"]))

    def test_sms_keyword_hit_handles_multiple_keywords(self):
        """Should match if ANY keyword is present."""
        self.assertTrue(sms_keyword_hit("Hello world", ["goodbye", "world"]))

    def test_sms_keyword_hit_ignores_empty_keywords(self):
        """Should skip empty string keywords."""
        self.assertFalse(sms_keyword_hit("Hello", ["", "xyz"]))

    def test_build_unmatched_sms_log_entries_creates_two_entries(self):
        """Should create header and body entries."""
        entries = build_unmatched_sms_log_entries(
            "/tmp",
            "test",
            "Test message",
            now=datetime(2024, 1, 1, 12, 0, 0)
        )

        self.assertEqual(len(entries), 2)
        self.assertIn("未匹配拦截", entries[0][1])
        self.assertIn("Test message", entries[1][1])

    def test_build_unmatched_sms_log_entries_includes_timestamp(self):
        """Should include timestamp in log entries."""
        entries = build_unmatched_sms_log_entries(
            "/tmp",
            "test",
            "Message",
            now=datetime(2024, 1, 1, 12, 30, 45)
        )

        self.assertIn("2024-01-01 12:30:45", entries[0][1])

    def test_build_unmatched_sms_log_entries_uses_correct_filename(self):
        """Should generate filename with date."""
        entries = build_unmatched_sms_log_entries(
            "/tmp/logs",
            "system",
            "Message",
            now=datetime(2024, 1, 1, 12, 0, 0)
        )

        path = entries[0][0]
        self.assertIn("/tmp/logs", path)
        self.assertIn("sms_system_2024-01-01.txt", path)

    def test_build_unmatched_sms_log_entries_sanitizes_prefix(self):
        """Should replace : in prefix for valid filename."""
        entries = build_unmatched_sms_log_entries(
            "/tmp",
            "test:prefix",
            "Message"
        )

        path = entries[0][0]
        self.assertNotIn(":", os.path.basename(path))
        self.assertIn("test_prefix", path)

    def test_build_unmatched_sms_log_entries_returns_empty_for_empty_message(self):
        """Should return empty list for empty message."""
        self.assertEqual(build_unmatched_sms_log_entries("/tmp", "test", ""), [])
        self.assertEqual(build_unmatched_sms_log_entries("/tmp", "test", None), [])

    def test_repeat_count_message_allows_first_messages(self):
        """Should allow messages below limit."""
        state = {}

        msg1 = repeat_count_message(state, "key1", "Message A", limit=3)
        msg2 = repeat_count_message(state, "key1", "Message A", limit=3)

        self.assertEqual(msg1, "Message A")
        self.assertEqual(msg2, "Message A")

    def test_repeat_count_message_marks_limit_message(self):
        """Should mark the limit-th message."""
        state = {}

        for _ in range(2):
            repeat_count_message(state, "key1", "Message", limit=3)

        msg3 = repeat_count_message(state, "key1", "Message", limit=3)
        self.assertIn("后续同类消息已忽略", msg3)

    def test_repeat_count_message_suppresses_after_limit(self):
        """Should return None after limit."""
        state = {}

        for _ in range(3):
            repeat_count_message(state, "key1", "Message", limit=3)

        msg4 = repeat_count_message(state, "key1", "Message", limit=3)
        self.assertIsNone(msg4)

    def test_repeat_count_message_uses_custom_suppressed_message(self):
        """Should use custom suppressed message at limit."""
        state = {}

        for _ in range(2):
            repeat_count_message(state, "key1", "Message", limit=3)

        msg3 = repeat_count_message(
            state, "key1", "Message", limit=3,
            suppressed_message="Custom suppressed"
        )
        self.assertEqual(msg3, "Custom suppressed")

    def test_repeat_count_message_resets_for_different_key(self):
        """Should reset count when key changes."""
        state = {}

        repeat_count_message(state, "key1", "Message A", limit=3)
        repeat_count_message(state, "key1", "Message A", limit=3)

        # Different key, count should reset
        msg = repeat_count_message(state, "key2", "Message B", limit=3)
        self.assertEqual(msg, "Message B")

    def test_parse_sms_callback_metadata_extracts_sender_and_time(self):
        metadata = parse_sms_callback_metadata(
            "106598731 26/06/24,13:44:15+32 【中国电信】验证码461582"
        )

        self.assertEqual(metadata["sender"], "106598731")
        self.assertEqual(metadata["from"], "106598731")
        self.assertEqual(metadata["phone"], "106598731")
        self.assertEqual(metadata["sms_time"], "2026-06-24 13:44:15")
        self.assertEqual(metadata["local_number"], "")

    def test_build_sms_ui_display_lines_formats_three_display_lines(self):
        lines = build_sms_ui_display_lines(
            "106598731 26/06/24,13:44:15+32 【中国电信】验证码461582",
            "【中国电信】\n验证码461582",
        )

        self.assertEqual(lines, [
            "号码：106598731",
            "时间：2026-06-24 13:44:15",
            "内容：",
            "【中国电信】",
            "验证码461582",
        ])

    def test_build_sms_ui_display_lines_keeps_single_line_compact(self):
        lines = build_sms_ui_display_lines(
            "106598731 26/06/24,13:44:15+32 验证码461582",
            "验证码461582",
        )

        self.assertEqual(lines, [
            "号码：106598731",
            "时间：2026-06-24 13:44:15",
            "内容：验证码461582",
        ])

    def test_sms_message_id_keeps_plain_sms_hash_compatible(self):
        old_id = hashlib.sha1(
            "10086\n2026-06-28 12:00:00\nhello".encode("utf-8")
        ).hexdigest()
        new_plain_id = sms_message_id(
            "10086",
            "2026-06-28 12:00:00",
            "hello",
            None,
            None,
            None,
        )

        self.assertEqual(new_plain_id, old_id)

    def test_sms_message_id_uses_concat_reference_when_present(self):
        first = sms_message_id(
            "10086",
            "2026-06-28 12:00:00",
            "same body",
            0x2A,
            8,
            2,
        )
        second = sms_message_id(
            "10086",
            "2026-06-28 12:00:00",
            "same body",
            0x2B,
            8,
            2,
        )

        self.assertNotEqual(first, second)

    def test_process_pending_sms_message_id_uses_concat_metadata(self):
        calls = []

        class MockPending:
            full_msg = "same body"
            callback_head = "10086 26/06/28,12:00:00+32 same body"
            display_lines = []
            concat_reference = 0x2A
            concat_reference_bits = 8
            concat_total = 2
            message_trace_id = "abc123def456"

        process_pending_sms(
            MockPending(),
            [], False, "", "",
            {}, 5,
            lambda msg, **kwargs: calls.append((msg, kwargs)),
            lambda x, y: None,
            lambda x, y: None,
            lambda: None,
            lambda x: None,
            lambda x: None,
            lambda x, y: None
        )

        variables = calls[0][1]["variables"]
        self.assertEqual(
            variables["message_id"],
            sms_message_id("10086", "2026-06-28 12:00:00", "same body", 0x2A, 8, 2),
        )
        self.assertEqual(variables["message_trace_id"], "abc123def456")

    def test_process_pending_sms_passes_trace_id_to_cloud_metadata(self):
        calls = []

        class MockPending:
            full_msg = "same body"
            callback_head = "10086 26/06/28,12:00:00+32 same body"
            display_lines = []
            message_trace_id = "abc123def456"

        process_pending_sms(
            MockPending(),
            [], False, "", "",
            {}, 5,
            lambda msg, **kwargs: None,
            lambda head, msg, metadata=None: calls.append((head, msg, metadata)),
            lambda x, y: None,
            lambda: None,
            lambda x: None,
            lambda x: None,
            lambda x, y: None
        )

        self.assertEqual(calls[0][2], {"message_trace_id": "abc123def456"})

    def test_process_pending_sms_returns_empty_for_none(self):
        """Should return 'empty' when pending is None."""
        result = process_pending_sms(
            None, [], False, "", "", {}, 5,
            lambda x: None, lambda x, y: None, lambda x, y: None,
            lambda: None, lambda x: None, lambda x: None, lambda x, y: None
        )
        self.assertEqual(result, "empty")

    def test_process_pending_sms_shows_matched_message(self):
        """Should return 'shown' and trigger UI for matched messages."""
        calls = {"ui": [], "alert": 0, "popup": []}

        class MockPending:
            full_msg = "Verification code:\n123456"
            callback_head = "+8613812345678 26/06/24,13:44:15+32 Verification code:"
            display_lines = ["old display line"]

        def port_ui(text, style):
            calls["ui"].append((text, style))

        def play_alert():
            calls["alert"] += 1

        def show_popup(msg):
            calls["popup"].append(msg)

        result = process_pending_sms(
            MockPending(),
            ["verification"],  # Keyword match
            False, "", "",
            {}, 5,
            lambda x: None,
            lambda x, y: None,
            port_ui,
            play_alert,
            show_popup,
            lambda x: None,
            lambda x, y: None
        )

        self.assertEqual(result, "shown")
        self.assertEqual(calls["alert"], 1)
        self.assertEqual(calls["popup"], ["Verification code:\n123456"])
        self.assertEqual(calls["ui"], [
            ("📩 收到短信：", "normal"),
            ("号码：+8613812345678", "sms"),
            ("时间：2026-06-24 13:44:15", "sms"),
            ("内容：", "sms"),
            ("Verification code:", "sms"),
            ("123456", "sms"),
        ])

    def test_process_pending_sms_passes_push_template_variables(self):
        calls = []

        class MockPending:
            full_msg = "验证码461582"
            callback_head = "106598731 26/06/24,13:44:15+32 验证码461582"
            display_lines = []

        process_pending_sms(
            MockPending(),
            [], False, "", "",
            {}, 5,
            lambda msg, **kwargs: calls.append((msg, kwargs)),
            lambda x, y: None,
            lambda x, y: None,
            lambda: None,
            lambda x: None,
            lambda x: None,
            lambda x, y: None
        )

        self.assertEqual(calls[0][0], "验证码461582")
        self.assertEqual(calls[0][1]["variables"]["sender"], "106598731")
        self.assertEqual(calls[0][1]["variables"]["sms_time"], "2026-06-24 13:44:15")

    def test_process_pending_sms_keeps_multiline_body_for_push(self):
        calls = []

        class MockPending:
            full_msg = "第一行\n第二行\n第三行"
            callback_head = "106598731 26/06/24,13:44:15+32 第一行"
            display_lines = []

        process_pending_sms(
            MockPending(),
            [], False, "", "",
            {}, 5,
            lambda msg, **kwargs: calls.append((msg, kwargs)),
            lambda x, y: None,
            lambda x, y: None,
            lambda: None,
            lambda x: None,
            lambda x: None,
            lambda x, y: None
        )

        self.assertEqual(calls[0][0], "第一行\n第二行\n第三行")

    def test_process_pending_sms_ignores_unmatched_message(self):
        """Should return 'ignored' for messages not matching keywords."""
        calls = {"ui": [], "alert": 0}

        class MockPending:
            full_msg = "Regular message"
            callback_head = ""
            display_lines = []

        result = process_pending_sms(
            MockPending(),
            ["verification", "urgent"],  # No match
            False, "", "",
            {}, 5,
            lambda x: None,
            lambda x, y: None,
            lambda x, y: calls["ui"].append((x, y)),
            lambda: calls["alert"] + 1,
            lambda x: None,
            lambda x: None,
            lambda x, y: None
        )

        self.assertEqual(result, "ignored")
        self.assertEqual(calls["alert"], 0)

    def test_process_pending_sms_always_sends_to_push_and_cloud(self):
        """Should always send to third-party push and cloud, regardless of keywords."""
        calls = {"push": [], "cloud": []}

        class MockPending:
            full_msg = "Any message"
            callback_head = "+8613812345678"
            display_lines = []

        def enqueue_push(msg):
            calls["push"].append(msg)

        def send_cloud(head, msg):
            calls["cloud"].append((head, msg))

        process_pending_sms(
            MockPending(),
            ["nonmatch"],  # Won't match
            False, "", "",
            {}, 5,
            enqueue_push,
            send_cloud,
            lambda x, y: None,
            lambda: None,
            lambda x: None,
            lambda x: None,
            lambda x, y: None
        )

        self.assertEqual(len(calls["push"]), 1)
        self.assertEqual(len(calls["cloud"]), 1)


if __name__ == "__main__":
    unittest.main()
