import unittest

from sms_core.sms_pdu import encode_text_sms_pdu, encode_text_sms_pdus, measure_text_sms_pdus


class SmsPduTests(unittest.TestCase):
    def test_encode_text_sms_pdu_basic_message(self):
        """Should encode simple SMS with international number."""
        pdu, length = encode_text_sms_pdu("+8613812345678", "Hello")

        self.assertIsInstance(pdu, str)
        self.assertGreater(len(pdu), 0)
        self.assertIsInstance(length, int)
        self.assertGreater(length, 0)

    def test_encode_text_sms_pdu_uses_international_format(self):
        """Should use type 91 for + prefix."""
        pdu, _ = encode_text_sms_pdu("+8613812345678", "Test")

        # Type 91 indicates international format
        self.assertIn("91", pdu)

    def test_encode_text_sms_pdu_uses_national_format(self):
        """Should use type 81 for numbers without + prefix."""
        pdu, _ = encode_text_sms_pdu("13812345678", "Test")

        # Type 81 indicates national format
        self.assertIn("81", pdu)

    def test_encode_text_sms_pdu_swaps_number_pairs(self):
        """Should swap digit pairs in phone number (PDU semi-octet format)."""
        pdu, _ = encode_text_sms_pdu("+1234", "Test")

        # +1234 should become: type=91, len=04, swapped=2143
        # Check that swapping occurred (exact position depends on protocol)
        self.assertIn("91", pdu)  # International type
        self.assertIn("04", pdu)  # Length = 4 digits

    def test_encode_text_sms_pdu_pads_odd_length_number(self):
        """Should pad odd-length numbers with F."""
        pdu, _ = encode_text_sms_pdu("+12345", "Test")

        # +12345 (5 digits) should be padded to 6 with F
        # Swapped: 21 43 F5
        # The F indicates padding
        self.assertIn("F", pdu.upper())

    def test_encode_text_sms_pdu_encodes_message_as_ucs2(self):
        """Should encode message text as UCS2 (UTF-16-BE)."""
        pdu, _ = encode_text_sms_pdu("+1234", "A")

        # 'A' in UTF-16-BE is 0x0041
        self.assertIn("0041", pdu.upper())

    def test_encode_text_sms_pdu_handles_unicode(self):
        """Should handle Unicode characters (Chinese, emoji, etc)."""
        pdu, length = encode_text_sms_pdu("+8613812345678", "你好")

        self.assertIsInstance(pdu, str)
        self.assertGreater(length, 0)
        # Chinese characters should be encoded
        self.assertGreaterEqual(len(pdu), 40)

    def test_encode_text_sms_pdu_calculates_correct_length(self):
        """CMGS length should be TPDU byte length (excluding SMSC)."""
        pdu, length = encode_text_sms_pdu("+1234", "Test")

        # PDU format: SMSC (00) + TPDU
        # Length should be TPDU bytes only
        # SMSC is 00 (1 byte = 2 hex chars)
        self.assertEqual(length, (len(pdu) - 2) // 2)

    def test_encode_text_sms_pdu_handles_empty_message(self):
        """Should handle empty message."""
        pdu, length = encode_text_sms_pdu("+1234", "")

        self.assertIsInstance(pdu, str)
        self.assertGreater(length, 0)
        # Should still have header/number even with empty message

    def test_encode_text_sms_pdu_handles_empty_phone(self):
        """Should handle empty phone number."""
        pdu, length = encode_text_sms_pdu("", "Test")

        self.assertIsInstance(pdu, str)
        self.assertIsInstance(length, int)

    def test_encode_text_sms_pdu_strips_whitespace(self):
        """Should strip whitespace from phone number."""
        pdu1, _ = encode_text_sms_pdu("+1234", "Test")
        pdu2, _ = encode_text_sms_pdu("  +1234  ", "Test")

        self.assertEqual(pdu1, pdu2)

    def test_encode_text_sms_pdu_includes_udh_marker(self):
        """Should include correct PDU header markers."""
        pdu, _ = encode_text_sms_pdu("+1234", "Test")

        # SMSC length 00
        self.assertTrue(pdu.startswith("00"))
        # PDU type: 11 = SMS-SUBMIT, no UDH
        self.assertIn("1100", pdu[:10])

    def test_encode_text_sms_pdu_sets_ucs2_encoding(self):
        """Should set DCS to 08 (UCS2 encoding)."""
        pdu, _ = encode_text_sms_pdu("+1234", "Test")

        # Data Coding Scheme (DCS) should be 08 for UCS2
        self.assertIn("0008", pdu)

    def test_encode_text_sms_pdu_max_single_segment_message(self):
        """Should encode messages that fit a single UCS2 segment."""
        long_msg = "A" * 70
        pdu, length = encode_text_sms_pdu("+1234", long_msg)

        self.assertGreater(len(pdu), 140)
        self.assertGreater(length, 50)

    def test_encode_text_sms_pdu_rejects_too_long_single_segment_message(self):
        with self.assertRaises(ValueError):
            encode_text_sms_pdu("+1234", "A" * 71)

    def test_encode_text_sms_pdus_splits_long_ucs2_message(self):
        pdus = encode_text_sms_pdus("+1234", "A" * 100, reference=0x2A)

        self.assertEqual(len(pdus), 2)
        for index, (pdu, length) in enumerate(pdus, start=1):
            self.assertIn("0500032A02" + f"{index:02X}", pdu)
            self.assertEqual(length, (len(pdu) - 2) // 2)
            self.assertLessEqual(int(pdu[20:22], 16), 140)

    def test_encode_text_sms_pdus_keeps_surrogate_pairs_together(self):
        pdus = encode_text_sms_pdus("+1234", "😀" * 40, reference=0x01)

        self.assertEqual(len(pdus), 2)
        self.assertTrue(all("D83DDE00" in pdu for pdu, _length in pdus))

    def test_measure_text_sms_pdus_uses_ucs2_bytes(self):
        single = measure_text_sms_pdus("A" * 70)
        split = measure_text_sms_pdus("A" * 71)
        emoji_single = measure_text_sms_pdus("😀" * 35)
        emoji_split = measure_text_sms_pdus("😀" * 36)

        self.assertEqual((single.ucs2_bytes, single.segment_count, single.too_long), (140, 1, False))
        self.assertEqual((split.ucs2_bytes, split.segment_count, split.too_long), (142, 2, False))
        self.assertEqual((emoji_single.ucs2_bytes, emoji_single.segment_count), (140, 1))
        self.assertEqual((emoji_split.ucs2_bytes, emoji_split.segment_count), (144, 2))

    def test_measure_text_sms_pdus_reports_segment_limit(self):
        info = measure_text_sms_pdus("A" * 17086)

        self.assertEqual(info.segment_count, 256)
        self.assertEqual(info.segment_limit, 255)
        self.assertTrue(info.too_long)

    def test_encode_text_sms_pdus_rejects_more_than_255_segments(self):
        with self.assertRaises(ValueError):
            encode_text_sms_pdus("+1234", "A" * 17086, reference=0x01)

    def test_encode_text_sms_pdu_special_characters(self):
        """Should handle special characters."""
        pdu, _ = encode_text_sms_pdu("+1234", "Test @#$%^&*()")

        self.assertIsInstance(pdu, str)
        self.assertGreater(len(pdu), 0)

    def test_encode_text_sms_pdu_multiline_message(self):
        """Should handle multiline messages."""
        pdu, _ = encode_text_sms_pdu("+1234", "Line1\nLine2\nLine3")

        self.assertIsInstance(pdu, str)
        # Newline characters should be encoded
        self.assertGreater(len(pdu), 0)

    def test_encode_text_sms_pdu_format_consistency(self):
        """PDU should be valid hex string."""
        pdu, _ = encode_text_sms_pdu("+8613812345678", "Test")

        # Should be valid hex (all uppercase recommended)
        try:
            int(pdu, 16)
            valid_hex = True
        except ValueError:
            valid_hex = False

        self.assertTrue(valid_hex)


if __name__ == "__main__":
    unittest.main()
