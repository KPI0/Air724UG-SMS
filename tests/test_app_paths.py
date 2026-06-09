import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from sms_core import app_paths


class AppPathsTests(unittest.TestCase):
    def test_get_app_dir_returns_project_root_for_nested_source_layout(self):
        current = Path(app_paths.__file__).resolve()
        self.assertEqual(current.parts[-3:], ("sms", "sms_core", "app_paths.py"))
        self.assertEqual(Path(app_paths.get_app_dir()).resolve(), current.parents[2])

    def test_resource_path_falls_back_to_project_root_in_source_mode(self):
        with tempfile.TemporaryDirectory() as tmp:
            app_dir = os.path.join(tmp, "sms")
            os.makedirs(app_dir)
            icon_path = os.path.join(tmp, "icon.ico")
            with open(icon_path, "wb") as file:
                file.write(b"icon")

            with patch.object(app_paths, "get_app_dir", return_value=app_dir), \
                    patch.object(app_paths.sys, "frozen", False, create=True):
                self.assertEqual(app_paths.resource_path("icon.ico"), icon_path)

    def test_resource_path_prefers_app_dir_when_file_exists(self):
        with tempfile.TemporaryDirectory() as tmp:
            app_dir = os.path.join(tmp, "sms")
            os.makedirs(app_dir)
            app_icon = os.path.join(app_dir, "icon.ico")
            with open(app_icon, "wb") as file:
                file.write(b"icon")

            with patch.object(app_paths, "get_app_dir", return_value=app_dir), \
                    patch.object(app_paths.sys, "frozen", False, create=True):
                self.assertEqual(app_paths.resource_path("icon.ico"), app_icon)


if __name__ == "__main__":
    unittest.main()
