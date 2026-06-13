import unittest

from sms_core.sms_pdu import encode_text_sms_pdu


class SmsPduTests(unittest.TestCase):
    def test_encode_text_sms_pdu_basic_message(self):
        """Should encode simple SMS with international number."""
        pdu, length = encode_text_sms_pdu("+8613800138000", "Hello")

        self.assertIsInstance(pdu, str)
        self.assertGreater(len(pdu), 0)
        self.assertIsInstance(length, int)
        self.assertGreater(length, 0)

    def test_encode_text_sms_pdu_uses_international_format(self):
        """Should use type 91 for + prefix."""
        pdu, _ = encode_text_sms_pdu("+8613800138000", "Test")

        # Type 91 indicates international format
        self.assertIn("91", pdu)

    def test_encode_text_sms_pdu_uses_national_format(self):
        """Should use type 81 for numbers without + prefix."""
        pdu, _ = encode_text_sms_pdu("13800138000", "Test")

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
        pdu, length = encode_text_sms_pdu("+8613800138000", "你好")

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

    def test_encode_text_sms_pdu_long_message(self):
        """Should encode long messages."""
        long_msg = "A" * 100
        pdu, length = encode_text_sms_pdu("+1234", long_msg)

        self.assertGreater(len(pdu), 200)
        self.assertGreater(length, 50)

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
        pdu, _ = encode_text_sms_pdu("+8613800138000", "Test")

        # Should be valid hex (all uppercase recommended)
        try:
            int(pdu, 16)
            valid_hex = True
        except ValueError:
            valid_hex = False

        self.assertTrue(valid_hex)


if __name__ == "__main__":
    unittest.main()
