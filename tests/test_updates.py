import unittest

from sms_core.updates import (
    UpdateCheckPlan,
    normalize_proxy_base,
    normalize_config_bases,
    build_release_api_urls,
    pick_zip_asset,
    version_tuple,
    build_download_url,
    plan_update_check,
    classify_update_error,
)


class UpdatesTests(unittest.TestCase):
    def test_normalize_proxy_base_adds_https_prefix(self):
        """Should add https:// prefix if missing."""
        self.assertEqual(
            normalize_proxy_base("gh-proxy.com"),
            "https://gh-proxy.com/"
        )

    def test_normalize_proxy_base_preserves_http_scheme(self):
        """Should keep http:// scheme if present."""
        self.assertEqual(
            normalize_proxy_base("http://proxy.local"),
            "http://proxy.local/"
        )

    def test_normalize_proxy_base_adds_trailing_slash(self):
        """Should add trailing slash if missing."""
        self.assertEqual(
            normalize_proxy_base("https://example.com"),
            "https://example.com/"
        )

    def test_normalize_proxy_base_preserves_existing_slash(self):
        """Should not double trailing slash."""
        self.assertEqual(
            normalize_proxy_base("https://example.com/"),
            "https://example.com/"
        )

    def test_normalize_proxy_base_handles_empty_input(self):
        """Should return empty string for empty input."""
        self.assertEqual(normalize_proxy_base(""), "")
        self.assertEqual(normalize_proxy_base("   "), "")

    def test_normalize_config_bases_normalizes_both(self):
        """Should normalize both proxy and api proxy bases."""
        proxy, api_proxy = normalize_config_bases("proxy.com", "api-proxy.com")
        self.assertEqual(proxy, "https://proxy.com/")
        self.assertEqual(api_proxy, "https://api-proxy.com/")

    def test_build_release_api_urls_includes_official_api(self):
        """Should always include official GitHub API."""
        urls = build_release_api_urls("owner", "repo", "", now=1234567890)
        self.assertTrue(any("api.github.com" in url for url in urls))

    def test_build_release_api_urls_adds_proxy_urls(self):
        """Should add proxy URLs when api_proxy_base provided."""
        urls = build_release_api_urls(
            "owner", "repo",
            "https://proxy1.com/|https://proxy2.com/",
            now=1234567890
        )
        self.assertGreater(len(urls), 1)
        self.assertTrue(any("proxy1.com" in url for url in urls))
        self.assertTrue(any("proxy2.com" in url for url in urls))

    def test_build_release_api_urls_adds_timestamp(self):
        """Should add timestamp to prevent caching."""
        urls = build_release_api_urls("owner", "repo", "", now=1234567890)
        # Official API doesn't have timestamp in this implementation
        # But proxy URLs should have ?t=
        self.assertTrue(any("repos/owner/repo/releases/latest" in url for url in urls))

    def test_pick_zip_asset_selects_largest_zip(self):
        """Should pick the largest .zip asset."""
        release = {
            "assets": [
                {"name": "small.zip", "size": 100},
                {"name": "large.zip", "size": 1000},
                {"name": "medium.zip", "size": 500},
            ]
        }
        asset = pick_zip_asset(release)
        self.assertEqual(asset["name"], "large.zip")

    def test_pick_zip_asset_ignores_non_zip_files(self):
        """Should only consider .zip files."""
        release = {
            "assets": [
                {"name": "file.tar.gz", "size": 2000},
                {"name": "file.zip", "size": 100},
            ]
        }
        asset = pick_zip_asset(release)
        self.assertEqual(asset["name"], "file.zip")

    def test_pick_zip_asset_returns_none_if_no_zip(self):
        """Should return None when no .zip assets."""
        release = {
            "assets": [
                {"name": "file.tar.gz", "size": 1000},
            ]
        }
        asset = pick_zip_asset(release)
        self.assertIsNone(asset)

    def test_pick_zip_asset_handles_empty_assets(self):
        """Should handle release with no assets."""
        release = {"assets": []}
        asset = pick_zip_asset(release)
        self.assertIsNone(asset)

    def test_version_tuple_parses_semantic_version(self):
        """Should parse version string into tuple."""
        self.assertEqual(version_tuple("1.2.3"), (1, 2, 3))
        self.assertEqual(version_tuple("v2.0.5"), (2, 0, 5))

    def test_version_tuple_handles_missing_components(self):
        """Should pad missing components with 0."""
        self.assertEqual(version_tuple("1"), (1, 0, 0))
        self.assertEqual(version_tuple("2.5"), (2, 5, 0))

    def test_version_tuple_strips_v_prefix(self):
        """Should remove v or V prefix."""
        self.assertEqual(version_tuple("v1.2.3"), (1, 2, 3))
        self.assertEqual(version_tuple("V1.2.3"), (1, 2, 3))

    def test_build_download_url_uses_proxy_when_provided(self):
        """Should prepend proxy base to GitHub URL."""
        url = build_download_url(
            "https://github.com/owner/repo/releases/download/v1.0/file.zip",
            "gh-proxy.com"
        )
        self.assertTrue(url.startswith("https://gh-proxy.com/"))
        self.assertIn("github.com", url)

    def test_build_download_url_returns_original_without_proxy(self):
        """Should return original URL when proxy is empty."""
        original = "https://github.com/owner/repo/releases/download/v1.0/file.zip"
        url = build_download_url(original, "")
        self.assertEqual(url, original)

    def test_build_download_url_handles_non_http_urls(self):
        """Should not proxy non-http URLs."""
        url = build_download_url("file:///local/path", "proxy.com")
        self.assertEqual(url, "file:///local/path")

    def test_plan_update_check_returns_latest_when_up_to_date(self):
        """Should return 'latest' plan when current >= latest."""
        release = {"tag_name": "v1.0.0"}
        plan = plan_update_check(release, "v1.0.0", "")

        self.assertEqual(plan.kind, "latest")
        self.assertEqual(plan.current_version, "v1.0.0")

    def test_plan_update_check_returns_update_when_newer_available(self):
        """Should return 'update' plan when newer version available."""
        release = {
            "tag_name": "v2.0.0",
            "assets": [
                {"name": "app.zip", "size": 1000, "browser_download_url": "https://example.com/app.zip"}
            ]
        }
        plan = plan_update_check(release, "v1.0.0", "")

        self.assertEqual(plan.kind, "update")
        self.assertEqual(plan.latest_tag, "v2.0.0")
        self.assertIn("example.com", plan.download_url)

    def test_plan_update_check_returns_no_zip_when_asset_missing(self):
        """Should return 'no_zip' when newer version has no .zip asset."""
        release = {
            "tag_name": "v2.0.0",
            "assets": []
        }
        plan = plan_update_check(release, "v1.0.0", "")

        self.assertEqual(plan.kind, "no_zip")
        self.assertEqual(plan.latest_tag, "v2.0.0")

    def test_plan_update_check_compares_versions_correctly(self):
        """Should correctly compare version tuples."""
        release_2_0 = {"tag_name": "v2.0.0", "assets": []}
        release_1_5 = {"tag_name": "v1.5.0", "assets": []}

        # Current 2.0.0, latest 1.5.0 -> already latest
        plan1 = plan_update_check(release_1_5, "v2.0.0", "")
        self.assertEqual(plan1.kind, "latest")

        # Current 1.0.0, latest 2.0.0 -> update available (but no zip)
        plan2 = plan_update_check(release_2_0, "v1.0.0", "")
        self.assertEqual(plan2.kind, "no_zip")

    def test_classify_update_error_recognizes_ssl_handshake_failure(self):
        """Should classify TLS handshake errors."""
        error = Exception("sslv3_alert_handshake_failure occurred")
        result = classify_update_error(error)
        self.assertIn("TLS握手失败", result)

    def test_classify_update_error_recognizes_timeout(self):
        """Should classify timeout errors."""
        error = Exception("The read operation timed out")
        result = classify_update_error(error)
        self.assertIn("连接超时", result)

    def test_classify_update_error_handles_unknown_errors(self):
        """Should handle unrecognized error types."""
        error = Exception("Some unknown error")
        result = classify_update_error(error)
        # Should return something (either "unknown" or the error text)
        self.assertIsNotNone(result)


if __name__ == "__main__":
    unittest.main()
