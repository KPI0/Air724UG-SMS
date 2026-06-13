import unittest

from sms_ui.window_utils import sync_and_focus_existing_window


class FakeWindow:
    def __init__(self, exists=True, fail_exists=False, fail_sync=False, fail_focus=False):
        self.exists = exists
        self.fail_exists = fail_exists
        self.fail_sync = fail_sync
        self.fail_focus = fail_focus
        self.calls = []
        self.synced = False

    def winfo_exists(self):
        if self.fail_exists:
            raise RuntimeError("exists failed")
        return self.exists

    def _sync(self):
        if self.fail_sync:
            raise RuntimeError("sync failed")
        self.synced = True

    def deiconify(self):
        if self.fail_focus:
            raise RuntimeError("focus failed")
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

    def test_sync_and_focus_existing_window_logs_exists_failure(self):
        logs = []

        self.assertFalse(sync_and_focus_existing_window(FakeWindow(fail_exists=True), "_sync", log_error=logs.append))
        self.assertEqual(len(logs), 1)
        self.assertIn("exists failed", logs[0])

    def test_sync_and_focus_existing_window_logs_sync_and_focus_failures(self):
        logs = []
        window = FakeWindow(fail_sync=True, fail_focus=True)

        self.assertTrue(sync_and_focus_existing_window(window, "_sync", log_error=logs.append))

        self.assertEqual(len(logs), 2)
        self.assertIn("sync failed", logs[0])
        self.assertIn("focus failed", logs[1])


if __name__ == "__main__":
    unittest.main()
