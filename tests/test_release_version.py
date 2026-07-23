import unittest

from sms_app.version import APP_VERSION
from tools.check_release_version import validate_release_tag


class ReleaseVersionTests(unittest.TestCase):
    def test_app_version_uses_stable_semver(self):
        self.assertRegex(APP_VERSION, r"^\d+\.\d+\.\d+$")

    def test_matching_release_tag_is_accepted(self):
        self.assertEqual(
            validate_release_tag(f"v{APP_VERSION}"),
            (True, f"v{APP_VERSION}"),
        )

    def test_mismatched_release_tag_is_rejected(self):
        ok, message = validate_release_tag("v9.9.9", APP_VERSION)
        self.assertFalse(ok)
        self.assertIn("does not match", message)

    def test_prerelease_and_malformed_tags_are_rejected(self):
        for tag in ("3.8.1", "v3.8", "v3.8.1-beta", "vfoo"):
            with self.subTest(tag=tag):
                self.assertFalse(validate_release_tag(tag, "3.8.1")[0])


if __name__ == "__main__":
    unittest.main()
