import unittest

from sms_core.cloud_protocol import (
    CLOUD_WS_DEFAULT_PATH,
    cloud_ws_url_has_host,
    normalize_cloud_ws_url,
    normalize_imei,
    parse_sms_callback_head,
    auth_status_from_ack,
    version_tuple,
)


class CloudProtocolTests(unittest.TestCase):
    def test_normalize_cloud_ws_url_appends_default_path_to_root(self):
        """Should append the dedicated device path to root URLs."""
        self.assertEqual(
            normalize_cloud_ws_url("ws://localhost:8080"),
            "ws://localhost:8080/ws/device"
        )
        self.assertEqual(
            normalize_cloud_ws_url("ws://localhost:8080/"),
            "ws://localhost:8080/ws/device"
        )

    def test_normalize_cloud_ws_url_preserves_custom_path(self):
        """Should not modify URLs with custom paths."""
        url = "ws://localhost:8080/custom/path"
        self.assertEqual(normalize_cloud_ws_url(url), url)

    def test_normalize_cloud_ws_url_handles_wss_scheme(self):
        """Should work with wss:// (secure WebSocket)."""
        self.assertEqual(
            normalize_cloud_ws_url("wss://example.com"),
            "wss://example.com/ws/device"
        )

    def test_normalize_cloud_ws_url_preserves_non_ws_urls(self):
        """Should return non-WebSocket URLs unchanged."""
        http_url = "http://example.com"
        self.assertEqual(normalize_cloud_ws_url(http_url), http_url)

    def test_normalize_cloud_ws_url_handles_empty_input(self):
        """Should return empty string for empty input."""
        self.assertEqual(normalize_cloud_ws_url(""), "")
        self.assertEqual(normalize_cloud_ws_url("   "), "")

    def test_normalize_cloud_ws_url_handles_malformed_urls(self):
        """Should attempt to normalize even malformed ws:// URLs."""
        # The function tries to parse and may add default path
        result = normalize_cloud_ws_url("ws://:invalid")
        self.assertTrue(result.startswith("ws://"))

    def test_cloud_ws_url_has_host_rejects_missing_or_blank_host(self):
        for url in (
            "ws://",
            "ws:///ws/device",
            "wss://?query=value",
            "ws://:8765/ws/device",
            "ws://bad host/ws/device",
        ):
            with self.subTest(url=url):
                self.assertFalse(cloud_ws_url_has_host(url))

    def test_cloud_ws_url_has_host_rejects_invalid_port_and_fragment(self):
        for url in (
            "ws://example.com:abc/ws/device",
            "ws://example.com:70000/ws/device",
            "ws://example.com:/ws/device",
            "ws://example.com:+80/ws/device",
            "ws://example.com: 80/ws/device",
            "ws://example.com:0/ws/device",
            "ws://example.com/ws/device#fragment",
            "ws://example.com/ws/device#",
            "ws://example.com\\bad",
        ):
            with self.subTest(url=url):
                self.assertFalse(cloud_ws_url_has_host(url))

    def test_cloud_ws_url_has_host_accepts_hostname_and_ip_addresses(self):
        for url in (
            "ws://localhost:8765/ws/device",
            "wss://example.com/ws/device",
            "ws://127.0.0.1:8765/ws/device",
            "ws://[::1]:8765/ws/device",
        ):
            with self.subTest(url=url):
                self.assertTrue(cloud_ws_url_has_host(url))

    def test_normalize_imei_removes_non_digits(self):
        """Should remove all non-digit characters from IMEI."""
        self.assertEqual(normalize_imei("1234-5678-9012-345"), "123456789012345")
        self.assertEqual(normalize_imei("1234 5678 9012 345"), "123456789012345")
        self.assertEqual(normalize_imei("IMEI:123456789012345"), "123456789012345")

    def test_normalize_imei_handles_clean_imei(self):
        """Should pass through clean numeric IMEI."""
        self.assertEqual(normalize_imei("123456789012345"), "123456789012345")

    def test_normalize_imei_handles_empty_input(self):
        """Should return empty string for empty input."""
        self.assertEqual(normalize_imei(""), "")
        self.assertEqual(normalize_imei("   "), "")
        self.assertEqual(normalize_imei(None), "")

    def test_parse_sms_callback_head_extracts_phone_and_message(self):
        """Should extract phone number and message from callback format."""
        text = '+8613123123123 24/12/25,14:30:45+32 Hello World'
        phone, message = parse_sms_callback_head(text)

        self.assertEqual(phone, "+8613123123123")
        self.assertEqual(message, "Hello World")

    def test_parse_sms_callback_head_handles_multiline_message(self):
        """Should preserve multiline message content."""
        text = '+8613123123123 24/12/25,14:30:45+32 Line1\nLine2\nLine3'
        phone, message = parse_sms_callback_head(text)

        self.assertEqual(phone, "+8613123123123")
        self.assertIn("Line1", message)
        self.assertIn("Line3", message)

    def test_parse_sms_callback_head_handles_phone_without_plus(self):
        """Should handle phone numbers without + prefix."""
        text = '13123123123 24/12/25,14:30:45+32 Message'
        phone, message = parse_sms_callback_head(text)

        self.assertEqual(phone, "13123123123")
        self.assertEqual(message, "Message")

    def test_parse_sms_callback_head_handles_non_numeric_sender_token(self):
        text = '3A968BD3FABBFC8C52 26/07/22,17:33:02+32 Brand message'
        sender, message = parse_sms_callback_head(text)

        self.assertEqual(sender, "3A968BD3FABBFC8C52")
        self.assertEqual(message, "Brand message")

    def test_parse_sms_callback_head_uses_timestamp_boundary_for_spaced_sender(self):
        sender, message = parse_sms_callback_head(
            "Bank Alert 26/07/22,17:33:02+32 Brand message"
        )

        self.assertEqual(sender, "Bank Alert")
        self.assertEqual(message, "Brand message")

    def test_parse_sms_callback_head_returns_original_on_no_match(self):
        """Should return empty phone and original text when format doesn't match."""
        text = 'Not a callback format'
        phone, message = parse_sms_callback_head(text)

        self.assertEqual(phone, "")
        self.assertEqual(message, text)

    def test_parse_sms_callback_head_handles_empty_input(self):
        """Should handle empty input gracefully."""
        phone, message = parse_sms_callback_head("")
        self.assertEqual(phone, "")
        self.assertEqual(message, "")

    def test_auth_status_from_ack_recognizes_authorized(self):
        """Should recognize various authorized status indicators."""
        authorized_payloads = [
            {"auth_status": "authorized"},
            {"auth_status": "ok"},
            {"status": "authorized"},
            {"status": "auth_ok"},
            {"auth_label": "已授权"},
        ]

        for payload in authorized_payloads:
            status = auth_status_from_ack(payload)
            self.assertEqual(status, "authorized", f"Failed for {payload}")

    def test_auth_status_from_ack_recognizes_failed(self):
        """Should recognize various failed status indicators."""
        failed_payloads = [
            {"auth_status": "failed"},
            {"auth_status": "auth_failed"},
            {"auth_status": "unauthorized"},
            {"status": "failed"},
            {"status": "unauthorized"},
            {"message": "密码不一致"},
            {"message": "密码错误"},
        ]

        for payload in failed_payloads:
            status = auth_status_from_ack(payload)
            self.assertEqual(status, "failed", f"Failed for {payload}")

    def test_auth_status_from_ack_returns_waiting_for_unknown(self):
        """Should return 'waiting' for unrecognized status."""
        status = auth_status_from_ack({"status": "pending"})
        self.assertEqual(status, "waiting")

    def test_auth_status_from_ack_handles_empty_payload(self):
        """Should handle None or empty dict."""
        self.assertEqual(auth_status_from_ack(None), "waiting")
        self.assertEqual(auth_status_from_ack({}), "waiting")

    def test_auth_status_from_ack_case_insensitive(self):
        """Should handle case-insensitive status values."""
        self.assertEqual(auth_status_from_ack({"auth_status": "AUTHORIZED"}), "authorized")
        self.assertEqual(auth_status_from_ack({"status": "AUTH_OK"}), "authorized")
        self.assertEqual(auth_status_from_ack({"auth_status": "FAILED"}), "failed")

    def test_auth_status_from_ack_rejects_explicit_negative_ok_even_if_status_claims_success(self):
        """An explicit ok=false must never authorize contradictory fields."""
        self.assertEqual(
            auth_status_from_ack(
                {"ok": False, "status": "ok", "auth_status": "authorized"}
            ),
            "failed",
        )

    def test_auth_status_from_ack_requires_boolean_true_when_ok_is_present(self):
        """String/truthy ok values are not accepted as an authorization proof."""
        self.assertEqual(
            auth_status_from_ack(
                {"ok": "true", "status": "ok", "auth_status": "authorized"}
            ),
            "failed",
        )

    def test_version_tuple_parses_semantic_version(self):
        """Should parse semantic version into tuple."""
        self.assertEqual(version_tuple("1.2.3"), (1, 2, 3))
        self.assertEqual(version_tuple("v2.0.1"), (2, 0, 1))
        self.assertEqual(version_tuple("V10.5.8"), (10, 5, 8))

    def test_version_tuple_handles_missing_components(self):
        """Should pad missing version components with 0."""
        self.assertEqual(version_tuple("1"), (1, 0, 0))
        self.assertEqual(version_tuple("2.5"), (2, 5, 0))

    def test_version_tuple_handles_invalid_components(self):
        """Should treat invalid components as 0."""
        self.assertEqual(version_tuple("1.x.3"), (1, 0, 3))
        self.assertEqual(version_tuple("a.b.c"), (0, 0, 0))

    def test_version_tuple_strips_whitespace_and_prefix(self):
        """Should handle whitespace and version prefix."""
        self.assertEqual(version_tuple("  v1.2.3  "), (1, 2, 3))
        self.assertEqual(version_tuple("V 1.0.0"), (1, 0, 0))

    def test_version_tuple_handles_empty_input(self):
        """Should return (0, 0, 0) for empty input."""
        self.assertEqual(version_tuple(""), (0, 0, 0))
        self.assertEqual(version_tuple("   "), (0, 0, 0))

    def test_version_tuple_truncates_to_three_components(self):
        """Should only return first 3 version components."""
        self.assertEqual(version_tuple("1.2.3.4.5"), (1, 2, 3))


if __name__ == "__main__":
    unittest.main()
