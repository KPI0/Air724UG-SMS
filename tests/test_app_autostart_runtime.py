import unittest

from sms_ui.app_autostart_runtime import set_autostart_runtime


class AppAutostartRuntimeTests(unittest.TestCase):
    def test_set_autostart_runtime_enables_startup_shortcut(self):
        calls = []

        ok = set_autostart_runtime(
            True,
            autostart_flag="--autostart",
            create_startup_shortcut=lambda flag: calls.append(("create", flag)),
            remove_startup_shortcut=lambda: calls.append(("remove",)),
            system_ui=lambda *args: calls.append(("ui", args)),
            show_error=lambda *args: calls.append(("error", args)),
        )

        self.assertTrue(ok)
        self.assertEqual(calls[0], ("create", "--autostart"))
        self.assertIn("开机自启", calls[1][1][0])

    def test_set_autostart_runtime_disables_startup_shortcut(self):
        calls = []

        ok = set_autostart_runtime(
            False,
            autostart_flag="--autostart",
            create_startup_shortcut=lambda flag: calls.append(("create", flag)),
            remove_startup_shortcut=lambda: calls.append(("remove",)),
            system_ui=lambda *args: calls.append(("ui", args)),
            show_error=lambda *args: calls.append(("error", args)),
        )

        self.assertTrue(ok)
        self.assertEqual(calls[0], ("remove",))
        self.assertIn("开机自启", calls[1][1][0])

    def test_set_autostart_runtime_reports_errors(self):
        calls = []

        ok = set_autostart_runtime(
            True,
            autostart_flag="--autostart",
            create_startup_shortcut=lambda flag: (_ for _ in ()).throw(RuntimeError("boom")),
            remove_startup_shortcut=lambda: calls.append(("remove",)),
            system_ui=lambda *args: calls.append(("ui", args)),
            show_error=lambda *args: calls.append(("error", args)),
        )

        self.assertFalse(ok)
        self.assertEqual(calls[0][0], "error")
        self.assertIn("boom", calls[0][1][1])


if __name__ == "__main__":
    unittest.main()
