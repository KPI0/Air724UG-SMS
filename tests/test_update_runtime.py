import unittest
from io import BytesIO
from unittest.mock import Mock, patch

from sms_core import updates

from sms_core.updates import (
    build_download_url,
    check_latest_release,
    fetch_latest_release,
    plan_update_check,
    version_tuple,
)


class UpdateRuntimeTests(unittest.TestCase):
    def test_invalid_release_cannot_be_reported_as_latest(self):
        for release in (None, [], {}, {"message": "upstream unavailable"},
                        {"tag_name": ""}, {"tag_name": 399}, {"tag_name": "invalid"}):
            with self.subTest(release=release):
                with self.assertRaises(ValueError):
                    plan_update_check(release, "3.9.9", "")

    def test_invalid_proxy_release_falls_back_to_official_api(self):
        official = {"tag_name": "v4.0.0", "assets": []}
        get_json = Mock(side_effect=[{"message": "upstream unavailable"}, official])
        self.assertEqual(fetch_latest_release("owner", "repo", "proxy.example", get_json=get_json), official)
        self.assertEqual(get_json.call_count, 2)

    def test_invalid_asset_metadata_falls_back_to_official_api(self):
        official = {"tag_name": "v4.0.0", "assets": []}
        for assets in (None, {}, [None], [{"name": 1}], [{"name": "app.zip", "size": "bad"}]):
            with self.subTest(assets=assets):
                get_json = Mock(side_effect=[{"tag_name": "v4.0.0", "assets": assets}, official])
                self.assertEqual(fetch_latest_release("owner", "repo", "proxy.example", get_json=get_json), official)
                self.assertEqual(get_json.call_count, 2)

    def test_unparseable_numeric_tag_is_rejected(self):
        with self.assertRaises(ValueError):
            plan_update_check({"tag_name": "9" * 5000}, "3.9.9", "")

    def test_invalid_download_link_is_rejected(self):
        for url in ("", "file:///C:/fake.zip", "javascript:alert(1)", "https:///missing-host.zip"):
            with self.subTest(url=url):
                release = {"tag_name": "v4.0.0", "assets": [
                    {"name": "app.zip", "browser_download_url": url}]}
                with self.assertRaises(ValueError):
                    plan_update_check(release, "3.9.9", "proxy.example")

    def test_proxy_test_does_not_accept_error_json_as_success(self):
        result = updates.test_update_proxy_connectivity(
            "owner", "repo", "proxy.example", "",
            get_json=lambda *a, **kw: {"message": "upstream unavailable"},
            probe=Mock(side_effect=AssertionError("Must not probe an invalid release")),
        )
        self.assertFalse(result["ok_bases"])
        self.assertTrue(all(not ok for _name, ok, _info in result["checks"]))

    def test_oversized_release_response_is_rejected_with_bounded_read(self):
        stream = BytesIO(b'{"padding":"' + b"x" * (2 * 1024 * 1024) + b'"}')
        response = Mock(wraps=stream)
        response.__enter__ = Mock(return_value=response)
        response.__exit__ = Mock(return_value=False)
        opener = Mock()
        opener.open.return_value = response
        with patch.object(updates, "_build_no_proxy_opener", return_value=opener):
            with self.assertRaises(ValueError):
                updates.http_get_json("https://example.invalid/release", retries=1)
        self.assertLessEqual(stream.tell(), 1024 * 1024 + 1)

    def test_http_json_accepts_valid_utf8_and_rejects_corrupt_data(self):
        for payload in (b'{"tag_name":"v4.0.0"}', b'{"tag_name":"v4.0.0\xff"}', b'not json'):
            with self.subTest(payload=payload):
                opener = Mock()
                opener.open.return_value = BytesIO(payload)
                with patch.object(updates, "_build_no_proxy_opener", return_value=opener):
                    if payload == b'{"tag_name":"v4.0.0"}':
                        self.assertEqual(updates.http_get_json("https://example.invalid/release"), {"tag_name": "v4.0.0"})
                    else:
                        with self.assertRaises(ValueError):
                            updates.http_get_json("https://example.invalid/release")
                self.assertEqual(opener.open.call_count, 1)

    def test_version_tuple_normalizes_tags_and_bad_parts(self):
        self.assertEqual(version_tuple("v3.6.7"), (3, 6, 7))
        self.assertEqual(version_tuple("3.bad"), (3, 0, 0))
        self.assertEqual(version_tuple(""), (0, 0, 0))

    def test_build_download_url_uses_proxy_only_for_http_urls(self):
        self.assertEqual(
            build_download_url("https://github.com/repo/app.zip", "proxy.example"),
            "https://proxy.example/https://github.com/repo/app.zip",
        )
        self.assertEqual(build_download_url("file.zip", "proxy.example"), "file.zip")
        self.assertEqual(build_download_url("https://github.com/repo/app.zip", ""), "https://github.com/repo/app.zip")

    def test_plan_update_check_reports_latest(self):
        plan = plan_update_check({"tag_name": "v3.6.6"}, "3.6.6", "proxy.example")

        self.assertEqual(plan.kind, "latest")
        self.assertEqual(plan.current_version, "3.6.6")

    def test_plan_update_check_reports_missing_zip(self):
        plan = plan_update_check({"tag_name": "v3.6.7", "assets": []}, "3.6.6", "proxy.example")

        self.assertEqual(plan.kind, "no_zip")
        self.assertEqual(plan.latest_tag, "v3.6.7")

    def test_plan_update_check_builds_proxied_download_url(self):
        plan = plan_update_check(
            {
                "tag_name": "v3.6.7",
                "assets": [
                    {
                        "name": "sms.zip",
                        "size": 10,
                        "browser_download_url": "https://github.com/repo/sms.zip",
                    }
                ],
            },
            "3.6.6",
            "proxy.example",
        )

        self.assertEqual(plan.kind, "update")
        self.assertEqual(plan.download_url, "https://proxy.example/https://github.com/repo/sms.zip")

    def test_fetch_latest_release_tries_proxy_then_direct(self):
        calls = []

        def fake_get_json(url, timeout=0, retries=0):
            calls.append(url)
            if len(calls) == 1:
                raise RuntimeError("proxy failed")
            return {"tag_name": "v1"}

        result = fetch_latest_release("owner", "repo", "api.proxy", get_json=fake_get_json)

        self.assertEqual(result, {"tag_name": "v1"})
        self.assertEqual(len(calls), 2)
        self.assertIn("https://api.proxy/repos/owner/repo/releases/latest", calls[0])
        self.assertEqual(calls[1], "https://api.github.com/repos/owner/repo/releases/latest")

    def test_check_latest_release_combines_fetch_and_plan(self):
        plan = check_latest_release(
            "owner",
            "repo",
            "1.0.0",
            "",
            "",
            get_json=lambda *_args, **_kwargs: {
                "tag_name": "v1.0.1",
                "assets": [{"name": "app.zip", "browser_download_url": "https://example.com/app.zip"}],
            },
        )

        self.assertEqual(plan.kind, "update")
        self.assertEqual(plan.latest_tag, "v1.0.1")
        self.assertEqual(plan.download_url, "https://example.com/app.zip")


if __name__ == "__main__":
    unittest.main()
