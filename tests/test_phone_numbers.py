import unittest

from sms_core.phone_numbers import normalize_call_number


class PhoneNumbersTests(unittest.TestCase):
    def test_normalize_call_number_removes_plus86_prefix(self):
        """Should remove +86 prefix from Chinese numbers."""
        self.assertEqual(normalize_call_number("+8613812345678"), "13812345678")

    def test_normalize_call_number_removes_86_prefix_for_long_numbers(self):
        """Should remove 86 prefix from numbers longer than 11 digits."""
        self.assertEqual(normalize_call_number("8613812345678"), "13812345678")

    def test_normalize_call_number_preserves_86_for_short_numbers(self):
        """Should keep 86 prefix if total length <= 11 (likely not country code)."""
        # 86123456789 is 11 digits, might be local number starting with 86
        self.assertEqual(normalize_call_number("86123456789"), "86123456789")

    def test_normalize_call_number_preserves_local_numbers(self):
        """Should keep local numbers unchanged."""
        self.assertEqual(normalize_call_number("13812345678"), "13812345678")
        self.assertEqual(normalize_call_number("02012345678"), "02012345678")

    def test_normalize_call_number_handles_international_non_china(self):
        """Should keep non-Chinese international numbers unchanged."""
        self.assertEqual(normalize_call_number("+14155551234"), "+14155551234")
        self.assertEqual(normalize_call_number("+442071234567"), "+442071234567")

    def test_normalize_call_number_strips_whitespace(self):
        """Should strip leading/trailing whitespace."""
        self.assertEqual(normalize_call_number("  13812345678  "), "13812345678")
        self.assertEqual(normalize_call_number("\t+8613812345678\n"), "13812345678")

    def test_normalize_call_number_handles_empty_input(self):
        """Should return empty string for empty input."""
        self.assertEqual(normalize_call_number(""), "")
        self.assertEqual(normalize_call_number("   "), "")

    def test_normalize_call_number_handles_none_input(self):
        """Should handle None gracefully."""
        self.assertEqual(normalize_call_number(None), "")

    def test_normalize_call_number_distinguishes_by_length(self):
        """Should use length to distinguish country code from local prefix."""
        # 8613812345678 (13 digits) -> remove 86
        self.assertEqual(normalize_call_number("8613812345678"), "13812345678")

        # 8612345678 (10 digits) -> keep 86 (not long enough to be +86)
        self.assertEqual(normalize_call_number("8612345678"), "8612345678")

    def test_normalize_call_number_edge_case_11_digits_with_86(self):
        """Should keep 86 prefix when total is exactly 11 digits."""
        # Edge case: 86 + 9 digits = 11 total, kept as-is
        self.assertEqual(normalize_call_number("86123456789"), "86123456789")

    def test_normalize_call_number_edge_case_12_digits_with_86(self):
        """Should remove 86 prefix when total is 12+ digits."""
        # 86 + 10 digits = 12 total, 86 is removed
        self.assertEqual(normalize_call_number("861234567890"), "1234567890")

    def test_normalize_call_number_consistency(self):
        """Should produce consistent results for same input."""
        number = "+8613812345678"
        result1 = normalize_call_number(number)
        result2 = normalize_call_number(number)
        self.assertEqual(result1, result2)

    def test_normalize_call_number_idempotent(self):
        """Should be idempotent (normalizing twice = normalizing once)."""
        number = "+8613812345678"
        once = normalize_call_number(number)
        twice = normalize_call_number(once)
        self.assertEqual(once, twice)


if __name__ == "__main__":
    unittest.main()
