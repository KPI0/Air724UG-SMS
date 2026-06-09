import unittest

from sms_ui.window_utils import sync_and_focus_existing_window


class FakeWindow:
    def __init__(self, exists=True):
        self.exists = exists
        self.calls = []
        self.synced = False

    def winfo_exists(self):
        return self.exists

    def _sync(self):
        self.synced = True

    def deiconify(self):
        self.calls.append("deiconify")

    def lift(self):
        self.calls.append("lift")

    def focus_force(self):
        self.calls.append("focus_force")


class WindowUtilsTests(unittest.TestCase):
    def test_sync_and_focus_existing_window_returns_false_for_missing_window(self):
        self.assertFalse(sync_and_focus_existing_window(None, "_sync"))
        self.assertFalse(sync_and_focus_existing_window(FakeWindow(exists=False), "_sync"))

    def test_sync_and_focus_existing_window_syncs_and_focuses(self):
        window = FakeWindow()

        self.assertTrue(sync_and_focus_existing_window(window, "_sync"))
        self.assertTrue(window.synced)
        self.assertEqual(window.calls, ["deiconify", "lift", "focus_force"])


if __name__ == "__main__":
    unittest.main()
