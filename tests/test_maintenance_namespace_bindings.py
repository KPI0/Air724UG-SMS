import unittest
from unittest.mock import patch

import sms_ui.maintenance_namespace_bindings as bindings


class MaintenanceNamespaceBindingsTests(unittest.TestCase):
    def make_namespace(self):
        calls = []
        return {
            "calls": calls,
            "LOG_DIR": "logs",
            "clear_window": lambda: calls.append(("clear",)),
            "system_ui": lambda message: calls.append(("system", message)),
            "schedule_next_midnight_clear": lambda: calls.append(("schedule_midnight",)),
        }

    def test_install_registers_expected_names(self):
        namespace = self.make_namespace()

        result = bindings.install_maintenance_namespace_bindings(namespace)

        self.assertIs(result, namespace)
        for name in (
            "open_log_dir",
            "cleanup_old_logs",
            "open_log_cleanup_dialog",
            "open_update_proxy_dialog",
            "_auto_log_cleanup_tick",
            "schedule_auto_log_cleanup",
            "check_update_and_prompt",
            "clear_text_area_for_new_day",
            "schedule_next_midnight_clear",
        ):
            self.assertIn(name, namespace)
            self.assertTrue(callable(namespace[name]))

    def test_bindings_forward_namespace_and_arguments(self):
        namespace = self.make_namespace()
        bindings.install_maintenance_namespace_bindings(namespace)

        with patch.object(bindings, "open_log_dir_namespace_runtime", return_value="opened") as open_dir, \
                patch.object(bindings, "cleanup_old_logs_in_dir", return_value=3) as cleanup, \
                patch.object(bindings, "open_log_cleanup_dialog_namespace_runtime", return_value="cleanup_dialog") as cleanup_dialog, \
                patch.object(bindings, "open_update_proxy_dialog_namespace_runtime", return_value="proxy") as proxy, \
                patch.object(bindings, "run_auto_log_cleanup_tick_namespace_runtime", return_value="tick") as tick, \
                patch.object(bindings, "schedule_auto_log_cleanup_namespace_runtime", return_value="scheduled") as schedule, \
                patch.object(bindings, "check_update_and_prompt_namespace_runtime", return_value="checked") as check, \
                patch.object(bindings, "schedule_next_midnight_clear_namespace_runtime", return_value="midnight") as midnight:
            self.assertEqual(namespace["open_log_dir"](), "opened")
            self.assertEqual(namespace["cleanup_old_logs"](30), 3)
            self.assertEqual(namespace["open_log_cleanup_dialog"](), "cleanup_dialog")
            self.assertEqual(namespace["open_update_proxy_dialog"](), "proxy")
            self.assertEqual(namespace["_auto_log_cleanup_tick"](), "tick")
            self.assertEqual(namespace["schedule_auto_log_cleanup"](restart=False, first_delay_sec=5), "scheduled")
            self.assertEqual(namespace["check_update_and_prompt"](), "checked")
            self.assertEqual(namespace["schedule_next_midnight_clear"](), "midnight")

        open_dir.assert_called_once_with(namespace)
        cleanup.assert_called_once_with("logs", 30)
        cleanup_dialog.assert_called_once_with(namespace)
        proxy.assert_called_once_with(namespace)
        tick.assert_called_once_with(namespace)
        schedule.assert_called_once_with(namespace, restart=False, first_delay_sec=5)
        check.assert_called_once_with(namespace)
        midnight.assert_called_once_with(namespace)

    def test_clear_text_area_for_new_day_runs_ui_actions(self):
        namespace = self.make_namespace()
        bindings.install_maintenance_namespace_bindings(namespace)

        with patch.object(bindings, "schedule_next_midnight_clear_namespace_runtime", return_value="scheduled") as schedule:
            namespace["clear_text_area_for_new_day"]()

        self.assertEqual(namespace["calls"], [
            ("clear",),
            ("system", "📅 新的一天，窗口已清空"),
        ])
        schedule.assert_called_once_with(namespace)


if __name__ == "__main__":
    unittest.main()
