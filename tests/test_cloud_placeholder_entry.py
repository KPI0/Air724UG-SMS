import unittest
from unittest.mock import patch

from sms_ui.cloud_placeholder_entry import CloudPlaceholderEntry


class FakeVar:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeWidget:
    def __init__(self):
        self.config_calls = []

    def config(self, **kwargs):
        self.config_calls.append(kwargs)


class CloudPlaceholderEntryTests(unittest.TestCase):
    def make_secret_entry(self):
        entry = CloudPlaceholderEntry.__new__(CloudPlaceholderEntry)
        entry.parent = "window"
        entry.variable = FakeVar("old-secret")
        entry.visible_var = FakeVar(False)
        entry.placeholder_active = True
        entry.entry = FakeWidget()
        entry.eye_button = FakeWidget()
        return entry

    def test_generate_secret_cancel_keeps_current_password(self):
        entry = self.make_secret_entry()

        with patch("sms_ui.cloud_placeholder_entry.confirm_secret_reset", return_value=False) as confirm, \
                patch("sms_ui.cloud_placeholder_entry.secrets.choice") as choice:
            entry.generate_secret()

        confirm.assert_called_once_with("window")
        choice.assert_not_called()
        self.assertEqual(entry.variable.get(), "old-secret")
        self.assertFalse(entry.visible_var.get())
        self.assertTrue(entry.placeholder_active)

    def test_generate_secret_after_confirm_replaces_password_and_shows_it(self):
        entry = self.make_secret_entry()

        with patch("sms_ui.cloud_placeholder_entry.confirm_secret_reset", return_value=True) as confirm, \
                patch("sms_ui.cloud_placeholder_entry.secrets.choice", return_value="A") as choice:
            entry.generate_secret()

        confirm.assert_called_once_with("window")
        self.assertEqual(choice.call_count, 16)
        self.assertEqual(entry.variable.get(), "A" * 16)
        self.assertTrue(entry.visible_var.get())
        self.assertFalse(entry.placeholder_active)
        self.assertIn({"style": "TEntry", "show": ""}, entry.entry.config_calls)
        self.assertIn({"text": "🙈"}, entry.eye_button.config_calls)


if __name__ == "__main__":
    unittest.main()
