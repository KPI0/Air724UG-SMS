import json
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from tools.sms_replay import replay_lines


FIXTURE_PATH = ROOT / "tests" / "fixtures" / "sms_replay" / "cases.json"


def _load_fixture():
    return json.loads(FIXTURE_PATH.read_text(encoding="utf-8"))


def _calls_by_name(calls, name):
    return [args for kind, args, _kwargs in calls if kind == name]


class SmsSerialReplayFixtureTests(unittest.TestCase):
    def test_fixture_schema_is_append_only_friendly(self):
        data = _load_fixture()

        self.assertEqual(data.get("schema"), 1)
        names = [case.get("name") for case in data.get("cases", [])]
        self.assertEqual(len(names), len(set(names)))
        self.assertGreaterEqual(len(names), 1)
        for case in data["cases"]:
            with self.subTest(case=case.get("name")):
                self.assertIsInstance(case.get("lines"), list)
                self.assertGreater(len(case["lines"]), 0)
                self.assertIsInstance(case.get("expected_popups"), list)
                self.assertIsInstance(case.get("expected_cloud_bodies"), list)

    def test_replays_desensitized_air724ug_sms_logs(self):
        for case in _load_fixture()["cases"]:
            with self.subTest(case=case["name"]):
                calls = replay_lines(case["lines"])
                popups = [args[0] for args in _calls_by_name(calls, "sms_popup")]
                cloud_bodies = [args[1] for args in _calls_by_name(calls, "cloud_sms")]

                self.assertEqual(popups, case["expected_popups"])
                self.assertEqual(cloud_bodies, case["expected_cloud_bodies"])

                for forbidden in case.get("forbidden_popup_substrings", []):
                    self.assertFalse(
                        any(forbidden in popup for popup in popups),
                        f"{forbidden!r} leaked into replay popup output",
                    )


if __name__ == "__main__":
    unittest.main()
