import unittest
from unittest.mock import patch

from sms_app import bootstrap
from sms_core.config_runtime import ConfigInitializationError


class BootstrapConfigFailureTests(unittest.TestCase):
    def test_main_reports_runtime_directory_failure_before_config_initialization(self):
        error = PermissionError("access denied")

        with patch.object(bootstrap, "_initialize_paths_and_constants", side_effect=error), \
                patch.object(bootstrap, "_initialize_config") as initialize_config, \
                patch.object(bootstrap, "_initialize_cloud_settings") as initialize_cloud, \
                patch.object(bootstrap.messagebox, "showerror") as show_error:
            result = bootstrap.main()

        self.assertFalse(result)
        initialize_config.assert_not_called()
        initialize_cloud.assert_not_called()
        show_error.assert_called_once()
        self.assertEqual(show_error.call_args.args[0], "运行目录初始化失败")
        self.assertIn("写入权限", show_error.call_args.args[1])

    def test_main_reports_config_initialization_failure_and_stops_startup(self):
        error = ConfigInitializationError("配置文件创建失败，无法写入：config.ini")

        with patch.object(bootstrap, "_initialize_paths_and_constants") as initialize_paths, \
                patch.object(bootstrap, "_initialize_config", side_effect=error), \
                patch.object(bootstrap, "_initialize_cloud_settings") as initialize_cloud, \
                patch.object(bootstrap.messagebox, "showerror") as show_error:
            result = bootstrap.main()

        self.assertFalse(result)
        initialize_paths.assert_called_once_with()
        initialize_cloud.assert_not_called()
        show_error.assert_called_once()
        self.assertEqual(show_error.call_args.args[0], "配置初始化失败")
        self.assertIn("写入权限", show_error.call_args.args[1])


if __name__ == "__main__":
    unittest.main()
