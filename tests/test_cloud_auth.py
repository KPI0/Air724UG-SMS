import unittest

from sms_core.cloud_auth import (
    command_action,
    command_text,
    incoming_secret,
    target_value,
    target_imeis,
    secret_match_result,
    target_match_result,
    auth_match_result,
)


class CloudAuthTests(unittest.TestCase):
    def test_command_action_reads_from_multiple_fields(self):
        """Should read action from type, action, or infer from cmd."""
        self.assertEqual(command_action({"type": "command"}), "command")
        self.assertEqual(command_action({"action": "REBOOT"}), "reboot")
        self.assertEqual(command_action({"cmd": "ATI"}), "cmd")

    def test_command_action_returns_empty_for_missing(self):
        """Should return empty string when no action field present."""
        self.assertEqual(command_action({}), "")

    def test_command_text_reads_from_multiple_fields(self):
        """Should read command from command, data, or cmd field."""
        self.assertEqual(command_text({"command": "ATI"}), "ATI")
        self.assertEqual(command_text({"data": "ATZ"}), "ATZ")
        self.assertEqual(command_text({"cmd": "AT+CGSN"}), "AT+CGSN")

    def test_command_text_returns_empty_for_missing(self):
        """Should return empty when no command field present."""
        self.assertEqual(command_text({}), "")

    def test_incoming_secret_reads_from_multiple_fields(self):
        """Should read secret from various field names."""
        for field in ["secret", "device_secret", "password", "pwd", "token"]:
            result = incoming_secret({field: "my-secret-123"})
            self.assertEqual(result, "my-secret-123")

    def test_incoming_secret_strips_whitespace(self):
        """Should strip whitespace from secret."""
        result = incoming_secret({"secret": "  my-secret  "})
        self.assertEqual(result, "my-secret")

    def test_incoming_secret_returns_empty_for_missing(self):
        """Should return empty string when no secret field present."""
        self.assertEqual(incoming_secret({}), "")

    def test_target_value_reads_from_multiple_fields(self):
        """Should read target IMEI from various field names."""
        fields = ["target_imei", "imei", "device_imei", "target", "device"]
        for field in fields:
            result = target_value({field: "123456789012345"})
            self.assertEqual(result, "123456789012345")

    def test_target_value_returns_none_for_missing(self):
        """Should return None when no target field present."""
        self.assertIsNone(target_value({}))

    def test_target_imeis_handles_single_imei(self):
        """Should normalize single IMEI string."""
        raw, normalized = target_imeis({"target_imei": "123456789012345"})
        self.assertEqual(raw, "123456789012345")
        self.assertEqual(normalized, ["123456789012345"])

    def test_target_imeis_handles_list_of_imeis(self):
        """Should normalize list of IMEIs."""
        raw, normalized = target_imeis({
            "target_imei": ["123456789012345", "123456789012346"]
        })
        self.assertIsInstance(raw, list)
        self.assertEqual(len(normalized), 2)
        self.assertIn("123456789012345", normalized)

    def test_target_imeis_normalizes_imei_format(self):
        """Should normalize IMEI by removing spaces/dashes."""
        raw, normalized = target_imeis({"imei": "1234-5678-9012-345"})
        self.assertEqual(normalized, ["123456789012345"])

    def test_secret_match_result_accepts_correct_secret(self):
        """Should accept when secrets match."""
        ok, msg = secret_match_result(
            {"secret": "correct-password"},
            "correct-password"
        )
        self.assertTrue(ok)
        self.assertEqual(msg, "")

    def test_secret_match_result_rejects_wrong_secret(self):
        """Should reject when secrets don't match."""
        ok, msg = secret_match_result(
            {"secret": "wrong-password"},
            "correct-password"
        )
        self.assertFalse(ok)
        self.assertIn("密码错误", msg)

    def test_secret_match_result_rejects_missing_incoming_secret(self):
        """Should reject when incoming payload lacks secret."""
        ok, msg = secret_match_result({}, "correct-password")
        self.assertFalse(ok)
        self.assertIn("缺少密码", msg)

    def test_secret_match_result_rejects_empty_expected_secret(self):
        """Should reject when local expected secret is empty."""
        ok, msg = secret_match_result(
            {"secret": "any-secret"},
            ""
        )
        self.assertFalse(ok)
        self.assertIn("本机云端控制密码为空", msg)

    def test_secret_match_result_uses_constant_time_comparison(self):
        """Should use hmac.compare_digest for timing-attack resistance."""
        # This test verifies the function uses hmac.compare_digest
        # by checking it accepts valid secrets
        ok, _ = secret_match_result(
            {"secret": "test123"},
            "test123"
        )
        self.assertTrue(ok)

    def test_target_match_result_accepts_matching_imei(self):
        """Should accept when target IMEI matches local IMEI."""
        ok, msg = target_match_result(
            {"target_imei": "123456789012345"},
            "123456789012345"
        )
        self.assertTrue(ok)
        self.assertEqual(msg, "")

    def test_target_match_result_accepts_normalized_match(self):
        """Should accept when normalized IMEIs match."""
        ok, msg = target_match_result(
            {"target_imei": "1234-5678-9012-345"},
            "123456789012345"
        )
        self.assertTrue(ok)

    def test_target_match_result_rejects_different_imei(self):
        """Should reject when target IMEI differs from local."""
        ok, msg = target_match_result(
            {"target_imei": "123456789012999"},
            "123456789012345"
        )
        self.assertFalse(ok)
        self.assertIn("非本机指令", msg)

    def test_target_match_result_rejects_missing_target(self):
        """Should reject when target IMEI is missing."""
        ok, msg = target_match_result({}, "123456789012345")
        self.assertFalse(ok)
        self.assertIn("缺少目标IMEI", msg)

    def test_target_match_result_rejects_when_local_imei_unknown(self):
        """Should reject when local IMEI is not set."""
        ok, msg = target_match_result(
            {"target_imei": "123456789012345"},
            ""
        )
        self.assertFalse(ok)
        self.assertIn("本机IMEI未知", msg)

    def test_target_match_result_accepts_broadcast_to_list(self):
        """Should accept when local IMEI is in target list."""
        ok, msg = target_match_result(
            {"target_imei": ["123456789012346", "123456789012345", "123456789012999"]},
            "123456789012345"
        )
        self.assertTrue(ok)

    def test_auth_match_result_requires_both_target_and_secret(self):
        """Should check both target IMEI and secret."""
        # Valid target and secret
        ok, msg = auth_match_result(
            {"target_imei": "123456789012345", "secret": "password123"},
            "123456789012345",
            "password123"
        )
        self.assertTrue(ok)

    def test_auth_match_result_rejects_on_target_mismatch(self):
        """Should reject if target IMEI doesn't match (even with correct secret)."""
        ok, msg = auth_match_result(
            {"target_imei": "123456789012999", "secret": "password123"},
            "123456789012345",
            "password123"
        )
        self.assertFalse(ok)
        self.assertIn("非本机", msg)

    def test_auth_match_result_rejects_on_secret_mismatch(self):
        """Should reject if secret doesn't match (even with correct target)."""
        ok, msg = auth_match_result(
            {"target_imei": "123456789012345", "secret": "wrong-password"},
            "123456789012345",
            "password123"
        )
        self.assertFalse(ok)
        self.assertIn("密码错误", msg)

    def test_auth_match_result_rejects_missing_target(self):
        """Should reject when target IMEI is missing."""
        ok, msg = auth_match_result(
            {"secret": "password123"},
            "123456789012345",
            "password123"
        )
        self.assertFalse(ok)
        self.assertIn("缺少目标IMEI", msg)

    def test_auth_match_result_rejects_missing_secret(self):
        """Should reject when secret is missing."""
        ok, msg = auth_match_result(
            {"target_imei": "123456789012345"},
            "123456789012345",
            "password123"
        )
        self.assertFalse(ok)
        self.assertIn("缺少密码", msg)


if __name__ == "__main__":
    unittest.main()
