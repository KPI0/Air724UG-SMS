import unittest

from sms_core.long_sms_assembler import LongSmsAssembler
from sms_core.serial_sms import PendingSms
from sms_core.sms_pdu import ConcatSmsInfo


def parse_head(text):
    parts = str(text or "").split(" ", 2)
    if len(parts) >= 3:
        return parts[0], parts[2]
    return "", str(text or "")


class LongSmsAssemblerTests(unittest.TestCase):
    def test_non_concat_message_returns_immediately(self):
        assembler = LongSmsAssembler(parse_head)
        pending = PendingSms("10086 26/06/28,12:00:00+32 hello", "hello", [])

        self.assertIs(assembler.add_message(pending, now=1.0), pending)

    def test_concat_parts_are_sorted_before_processing(self):
        logs = []
        assembler = LongSmsAssembler(parse_head)
        part2 = PendingSms(
            "10086 26/06/28,12:00:00+32 BB",
            "BB",
            ["header", "10086 26/06/28,12:00:00+32 BB"],
            ConcatSmsInfo(0x1234, 2, 2, 16),
        )
        part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 AA",
            "AA",
            ["header", "10086 26/06/28,12:00:00+32 AA"],
            ConcatSmsInfo(0x1234, 2, 1, 16),
        )

        self.assertIsNone(assembler.add_message(part2, now=1.0, log=logs.append))
        complete = assembler.add_message(part1, now=2.0, log=logs.append)

        self.assertEqual(complete.full_msg, "AABB")
        self.assertEqual(complete.callback_head, "10086 26/06/28,12:00:00+32 AABB")
        self.assertTrue(any("SMS CONCAT COMPLETE" in item for item in logs))

    def test_duplicate_part_is_ignored(self):
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms("10086 26/06/28,12:00:00+32 AA", "AA", [], ConcatSmsInfo(0x2A, 2, 1))
        duplicate_part1 = PendingSms("10086 26/06/28,12:00:00+32 XX", "XX", [], ConcatSmsInfo(0x2A, 2, 1))
        part2 = PendingSms("10086 26/06/28,12:00:00+32 BB", "BB", [], ConcatSmsInfo(0x2A, 2, 2))

        self.assertIsNone(assembler.add_message(part1, now=1.0))
        self.assertIsNone(assembler.add_message(duplicate_part1, now=2.0))
        complete = assembler.add_message(part2, now=3.0)

        self.assertEqual(complete.full_msg, "AABB")

    def test_near_timestamp_parts_share_one_session_and_ignore_duplicates(self):
        assembler = LongSmsAssembler(parse_head)
        parts = [
            PendingSms(
                "10086 26/06/28,12:00:11+32 A1",
                "A1",
                [],
                ConcatSmsInfo(0x2A, 5, 1),
                "A1",
                "10086",
                "26/06/28,12:00:11+32",
            ),
            PendingSms(
                "10086 26/06/28,12:00:12+32 B2",
                "B2",
                [],
                ConcatSmsInfo(0x2A, 5, 2),
                "B2",
                "10086",
                "26/06/28,12:00:12+32",
            ),
            PendingSms(
                "10086 26/06/28,12:00:15+32 E5",
                "E5",
                [],
                ConcatSmsInfo(0x2A, 5, 5),
                "E5",
                "10086",
                "26/06/28,12:00:15+32",
            ),
            PendingSms(
                "10086 26/06/28,12:00:13+32 C3",
                "C3",
                [],
                ConcatSmsInfo(0x2A, 5, 3),
                "C3",
                "10086",
                "26/06/28,12:00:13+32",
            ),
            PendingSms(
                "10086 26/06/28,12:00:15+32 E5",
                "E5",
                [],
                ConcatSmsInfo(0x2A, 5, 5),
                "E5",
                "10086",
                "26/06/28,12:00:15+32",
            ),
            PendingSms(
                "10086 26/06/28,12:00:14+32 D4",
                "D4",
                [],
                ConcatSmsInfo(0x2A, 5, 4),
                "D4",
                "10086",
                "26/06/28,12:00:14+32",
            ),
        ]

        for pending in parts[:-1]:
            self.assertIsNone(assembler.add_message(pending, now=1.0))
        complete = assembler.add_message(parts[-1], now=2.0)

        self.assertEqual(complete.full_msg, "A1B2C3D4E5")
        self.assertEqual(assembler._pending, {})

    def test_same_sender_reference_with_different_timestamp_does_not_mix_parts(self):
        assembler = LongSmsAssembler(parse_head)
        stale_part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 A1",
            "A1",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "A1",
        )
        new_part1 = PendingSms(
            "10086 26/06/28,12:01:00+32 B1",
            "B1",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "B1",
        )
        new_part2 = PendingSms(
            "10086 26/06/28,12:01:00+32 B2",
            "B2",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "B2",
        )

        self.assertIsNone(assembler.add_message(stale_part1, now=1.0))
        self.assertIsNone(assembler.add_message(new_part1, now=20.0))
        complete = assembler.add_message(new_part2, now=21.0)

        self.assertEqual(complete.full_msg, "B1B2")
        self.assertNotEqual(complete.full_msg, "A1B2")

    def test_late_part_after_session_window_does_not_complete_stale_message(self):
        assembler = LongSmsAssembler(parse_head)
        stale_part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 A1",
            "A1",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "A1",
        )
        new_part2 = PendingSms(
            "10086 26/06/28,12:00:20+32 B2",
            "B2",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "B2",
        )

        self.assertIsNone(assembler.add_message(stale_part1, now=1.0))
        self.assertIsNone(assembler.add_message(new_part2, now=70.0))
        self.assertEqual(len(assembler._pending), 2)

    def test_completed_duplicate_cache_is_scoped_by_timestamp(self):
        assembler = LongSmsAssembler(parse_head)
        first_part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 SAME",
            "SAME",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "SAME",
        )
        first_part2 = PendingSms(
            "10086 26/06/28,12:00:00+32 A2",
            "A2",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "A2",
        )
        second_part1 = PendingSms(
            "10086 26/06/28,12:01:00+32 SAME",
            "SAME",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "SAME",
        )
        second_part2 = PendingSms(
            "10086 26/06/28,12:01:00+32 B2",
            "B2",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "B2",
        )

        self.assertIsNone(assembler.add_message(first_part1, now=1.0))
        self.assertEqual(assembler.add_message(first_part2, now=2.0).full_msg, "SAMEA2")
        self.assertIsNone(assembler.add_message(second_part1, now=30.0))
        complete = assembler.add_message(second_part2, now=31.0)

        self.assertEqual(complete.full_msg, "SAMEB2")

    def test_duplicate_old_part_does_not_complete_new_message(self):
        assembler = LongSmsAssembler(parse_head)
        first_part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 A1",
            "A1",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "A1",
        )
        first_part2 = PendingSms(
            "10086 26/06/28,12:00:00+32 A2",
            "A2",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "A2",
        )
        duplicate_old_part2 = PendingSms(
            "10086 26/06/28,12:00:00+32 A2",
            "A2",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "A2",
        )
        new_part1 = PendingSms(
            "10086 26/06/28,12:00:20+32 B1",
            "B1",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "B1",
        )

        self.assertIsNone(assembler.add_message(first_part1, now=1.0))
        self.assertEqual(assembler.add_message(first_part2, now=2.0).full_msg, "A1A2")
        self.assertIsNone(assembler.add_message(duplicate_old_part2, now=3.0))
        self.assertIsNone(assembler.add_message(new_part1, now=20.0))
        self.assertEqual(len(assembler._pending), 1)

    def test_completed_duplicate_near_timestamp_does_not_create_pending(self):
        logs = []
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 A1",
            "A1",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "A1",
            "10086",
            "26/06/28,12:00:00+32",
        )
        part2 = PendingSms(
            "10086 26/06/28,12:00:00+32 A2",
            "A2",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "A2",
            "10086",
            "26/06/28,12:00:00+32",
        )
        duplicate_part2 = PendingSms(
            "10086 26/06/28,12:00:01+32 A2",
            "A2",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "A2",
            "10086",
            "26/06/28,12:00:01+32",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0, log=logs.append))
        self.assertEqual(assembler.add_message(part2, now=2.0, log=logs.append).full_msg, "A1A2")
        self.assertIsNone(assembler.add_message(duplicate_part2, now=3.0, log=logs.append))

        self.assertEqual(assembler._pending, {})
        self.assertEqual(len(assembler._completed), 1)
        self.assertTrue(any("duplicate complete" in item for item in logs))

    def test_default_cleanup_uses_fixed_incomplete_timeout(self):
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 AA",
            "AA",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "AA",
        )
        part2 = PendingSms(
            "10086 26/06/28,12:00:00+32 BB",
            "BB",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "BB",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0))
        assembler.cleanup(now=250.0)
        self.assertEqual(len(assembler._pending), 1)

        complete = assembler.add_message(part2, now=250.0)
        self.assertIsNone(complete)
        self.assertEqual(len(assembler._pending), 2)

        assembler.cleanup(now=302.0)
        self.assertEqual(len(assembler._pending), 1)

        assembler.cleanup(now=551.0)
        self.assertEqual(assembler._pending, {})

    def test_completed_duplicate_cache_expires_before_incomplete_cache(self):
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 AA",
            "AA",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "AA",
        )
        part2 = PendingSms(
            "10086 26/06/28,12:00:00+32 BB",
            "BB",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "BB",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0))
        self.assertEqual(assembler.add_message(part2, now=2.0).full_msg, "AABB")
        self.assertIsNone(assembler.add_message(part2, now=34.0))

        self.assertEqual(len(assembler._completed), 0)
        self.assertEqual(len(assembler._pending), 1)

    def test_concat_uses_pdu_body_instead_of_collected_text(self):
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 noisy AA",
            "noisy AA",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "AA",
        )
        part2 = PendingSms(
            "10086 26/06/28,12:00:00+32 noisy BB",
            "noisy BB",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "BB",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0))
        complete = assembler.add_message(part2, now=2.0)

        self.assertEqual(complete.full_msg, "AABB")
        self.assertNotIn("noisy", complete.full_msg)

    def test_concat_completion_replaces_corrupted_callback_body(self):
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 AA\ufffd",
            "AA\ufffd",
            ["header", "10086 26/06/28,12:00:00+32 AA\ufffd"],
            ConcatSmsInfo(0x2A, 2, 1),
            "AA",
        )
        part2 = PendingSms(
            "10086 26/06/28,12:00:00+32 BB",
            "BB",
            ["header", "10086 26/06/28,12:00:00+32 BB"],
            ConcatSmsInfo(0x2A, 2, 2),
            "BB",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0))
        complete = assembler.add_message(part2, now=2.0)

        self.assertEqual(complete.full_msg, "AABB")
        self.assertEqual(complete.callback_head, "10086 26/06/28,12:00:00+32 AABB")
        self.assertEqual(complete.display_lines, ["header", "10086 26/06/28,12:00:00+32 AABB"])
        self.assertNotIn("\ufffd", complete.callback_head)

    def test_concat_completion_preserves_message_id_metadata(self):
        logs = []
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 AA",
            "AA",
            [],
            ConcatSmsInfo(0x1234, 2, 1, 16),
            "AA",
        )
        part2 = PendingSms(
            "10086 26/06/28,12:00:00+32 BB",
            "BB",
            [],
            ConcatSmsInfo(0x1234, 2, 2, 16),
            "BB",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0, log=logs.append))
        complete = assembler.add_message(part2, now=2.0, log=logs.append)

        self.assertIsNone(complete.concat_info)
        self.assertEqual(complete.concat_reference, 0x1234)
        self.assertEqual(complete.concat_reference_bits, 16)
        self.assertEqual(complete.concat_total, 2)
        self.assertRegex(complete.message_trace_id, r"^[0-9a-f]{12}$")
        self.assertTrue(all(
            f"Trace={complete.message_trace_id}" in item
            for item in logs
            if "SMS CONCAT" in item
        ))

    def test_concat_parts_with_adjacent_timestamps_complete(self):
        logs = []
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 AA",
            "AA",
            [],
            ConcatSmsInfo(0x2A, 3, 1),
            "AA",
            "10086",
            "26/06/28,12:00:00+32",
        )
        part2 = PendingSms(
            "10086 26/06/28,12:00:00+32 BB",
            "BB",
            [],
            ConcatSmsInfo(0x2A, 3, 2),
            "BB",
            "10086",
            "26/06/28,12:00:01+32",
        )
        part3 = PendingSms(
            "10086 26/06/28,12:00:00+32 CC",
            "CC",
            [],
            ConcatSmsInfo(0x2A, 3, 3),
            "CC",
            "10086",
            "26/06/28,12:00:02+32",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0, log=logs.append))
        self.assertIsNone(assembler.add_message(part2, now=2.0, log=logs.append))
        complete = assembler.add_message(part3, now=3.0, log=logs.append)

        self.assertEqual(complete.full_msg, "AABBCC")
        self.assertEqual(assembler._pending, {})
        self.assertTrue(any("SMS CONCAT COMPLETE" in item for item in logs))

    def test_reference_reuse_after_session_window_does_not_mix(self):
        assembler = LongSmsAssembler(parse_head)
        a1 = PendingSms(
            "10086 26/06/28,12:00:00+32 A1",
            "A1",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "A1",
            "10086",
            "26/06/28,12:00:00+32",
        )
        b2 = PendingSms(
            "10086 26/06/28,12:00:01+32 B2",
            "B2",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "B2",
            "10086",
            "26/06/28,12:00:01+32",
        )
        a2 = PendingSms(
            "10086 26/06/28,12:00:00+32 A2",
            "A2",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "A2",
            "10086",
            "26/06/28,12:00:00+32",
        )
        b1 = PendingSms(
            "10086 26/06/28,12:00:01+32 B1",
            "B1",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "B1",
            "10086",
            "26/06/28,12:00:01+32",
        )

        self.assertIsNone(assembler.add_message(a1, now=1.0))
        self.assertIsNone(assembler.add_message(b2, now=70.0))
        self.assertEqual(len(assembler._pending), 2)

        self.assertIsNone(assembler.add_message(a2, now=71.0))
        complete_b = assembler.add_message(b1, now=72.0)

        self.assertEqual(complete_b.full_msg, "B1B2")
        self.assertEqual(len(assembler._pending), 1)

    def test_concat_parts_with_timestamp_gap_still_complete_within_session_window(self):
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 AA",
            "AA",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "AA",
            "10086",
            "26/06/28,12:00:00+32",
        )
        part2 = PendingSms(
            "10086 26/06/28,12:00:00+32 BB",
            "BB",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "BB",
            "10086",
            "26/06/28,12:00:02+32",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0))
        complete = assembler.add_message(part2, now=2.0)
        self.assertEqual(complete.full_msg, "AABB")

    def test_empty_sender_does_not_enter_concat_cache(self):
        assembler = LongSmsAssembler(parse_head)
        pending = PendingSms("not-a-callback", "AA", [], ConcatSmsInfo(0x2A, 2, 1), "AA")

        self.assertIs(assembler.add_message(pending, now=1.0), pending)
        self.assertEqual(assembler._pending, {})

    def test_pdu_sender_is_used_when_callback_head_has_no_sender(self):
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms(
            "not-a-callback",
            "AA",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "AA",
            "10086",
            "26/06/28,12:00:00+32",
        )
        part2 = PendingSms(
            "not-a-callback",
            "BB",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "BB",
            "10086",
            "26/06/28,12:00:00+32",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0))
        complete = assembler.add_message(part2, now=2.0)

        self.assertEqual(complete.full_msg, "AABB")
        self.assertEqual(complete.concat_sender, None)
        self.assertEqual(complete.concat_timestamp, None)

    def test_sender_key_normalizes_callback_and_pdu_sender(self):
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms(
            "+8613812345678 26/06/28,12:00:00+32 AA",
            "AA",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "AA",
        )
        part2 = PendingSms(
            "not-a-callback",
            "BB",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "BB",
            "13812345678",
            "26/06/28,12:00:00+32",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0))
        complete = assembler.add_message(part2, now=2.0)

        self.assertEqual(complete.full_msg, "AABB")

    def test_concat_without_timestamp_does_not_enter_cache(self):
        assembler = LongSmsAssembler(parse_head)
        pending = PendingSms(
            "10086 AA",
            "AA",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "AA",
        )

        self.assertIs(assembler.add_message(pending, now=1.0), pending)
        self.assertEqual(assembler._pending, {})

    def test_concat_with_pdu_sender_does_not_require_timestamp(self):
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms(
            "not-a-callback",
            "AA",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "AA",
            "10086",
            "",
        )
        part2 = PendingSms(
            "not-a-callback",
            "BB",
            [],
            ConcatSmsInfo(0x2A, 2, 2),
            "BB",
            "10086",
            "",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0))
        complete = assembler.add_message(part2, now=2.0)

        self.assertEqual(complete.full_msg, "AABB")

    def test_missing_single_part_emits_incomplete_sms_after_wait(self):
        logs = []
        assembler = LongSmsAssembler(parse_head)
        part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 AA",
            "AA",
            ["header", "10086 26/06/28,12:00:00+32 AA"],
            ConcatSmsInfo(0x2A, 3, 1),
            "AA",
        )
        part3 = PendingSms(
            "10086 26/06/28,12:00:02+32 CC",
            "CC",
            ["header", "10086 26/06/28,12:00:02+32 CC"],
            ConcatSmsInfo(0x2A, 3, 3),
            "CC",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0, log=logs.append))
        incomplete = assembler.add_message(part3, now=40.0, log=logs.append)

        self.assertIn("[INCOMPLETE SMS]", incomplete.full_msg)
        self.assertIn("[MISSING PART 2/3]", incomplete.full_msg)
        self.assertEqual(incomplete.concat_reference, 0x2A)
        self.assertEqual(assembler._pending, {})
        self.assertTrue(any("SMS CONCAT INCOMPLETE" in item for item in logs))

    def test_duplicate_part_does_not_extend_timeout(self):
        assembler = LongSmsAssembler(parse_head, timeout=30.0)
        part1 = PendingSms("10086 26/06/28,12:00:00+32 AA", "AA", [], ConcatSmsInfo(0x2A, 2, 1))
        duplicate_part1 = PendingSms("10086 26/06/28,12:00:00+32 XX", "XX", [], ConcatSmsInfo(0x2A, 2, 1))
        part2 = PendingSms("10086 26/06/28,12:00:00+32 BB", "BB", [], ConcatSmsInfo(0x2A, 2, 2))

        self.assertIsNone(assembler.add_message(part1, now=1.0))
        self.assertIsNone(assembler.add_message(duplicate_part1, now=20.0))
        self.assertIsNone(assembler.add_message(part2, now=32.0))

    def test_cleanup_drops_incomplete_messages(self):
        assembler = LongSmsAssembler(parse_head, timeout=30.0)
        part1 = PendingSms("10086 26/06/28,12:00:00+32 AA", "AA", [], ConcatSmsInfo(0x2A, 2, 1))
        part2 = PendingSms("10086 26/06/28,12:00:00+32 BB", "BB", [], ConcatSmsInfo(0x2A, 2, 2))

        self.assertIsNone(assembler.add_message(part1, now=1.0))
        self.assertIsNone(assembler.add_message(part2, now=32.0))

    def test_cleanup_logs_incomplete_message_timeout(self):
        logs = []
        assembler = LongSmsAssembler(parse_head, timeout=30.0)
        part1 = PendingSms(
            "10086 26/06/28,12:00:00+32 AA",
            "AA",
            [],
            ConcatSmsInfo(0x2A, 2, 1),
            "AA",
        )

        self.assertIsNone(assembler.add_message(part1, now=1.0))
        assembler.cleanup(now=32.0, log=logs.append)

        self.assertEqual(assembler._pending, {})
        self.assertEqual(len(logs), 1)
        self.assertIn("SMS CONCAT TIMEOUT", logs[0])
        self.assertIn("Sender=10086", logs[0])
        self.assertIn("Ref=0x2A", logs[0])
        self.assertIn("Parts=1/2", logs[0])

    def test_pending_cache_limit_evicts_oldest_incomplete_message(self):
        logs = []
        assembler = LongSmsAssembler(parse_head, max_pending_entries=2)
        parts = [
            PendingSms("10086 26/06/28,12:00:00+32 A1", "A1", [], ConcatSmsInfo(0x2A, 2, 1), "A1"),
            PendingSms("10086 26/06/28,12:00:10+32 B1", "B1", [], ConcatSmsInfo(0x2B, 2, 1), "B1"),
            PendingSms("10086 26/06/28,12:00:20+32 C1", "C1", [], ConcatSmsInfo(0x2C, 2, 1), "C1"),
        ]

        for index, pending in enumerate(parts, start=1):
            self.assertIsNone(assembler.add_message(pending, now=float(index), log=logs.append))

        self.assertEqual(len(assembler._pending), 2)
        self.assertFalse(any(key[2] == 0x2A for key in assembler._pending))
        self.assertTrue(any("SMS CONCAT EVICT" in item and "Ref=0x2A" in item for item in logs))


if __name__ == "__main__":
    unittest.main()
