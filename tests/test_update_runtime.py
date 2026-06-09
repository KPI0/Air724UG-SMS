import unittest

from sms_core.updates import (
    build_download_url,
    check_latest_release,
    fetch_latest_release,
    plan_update_check,
    version_tuple,
)


class UpdateRuntimeTests(unittest.TestCase):
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
