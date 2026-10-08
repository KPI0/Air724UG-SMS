import configparser
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch

from sms_app import bootstrap
from sms_core.config_runtime import ConfigInitializationError, initialize_config_runtime


class StartupEncodingTests(unittest.TestCase):
    def test_invalid_encoding_preserves_file_and_reports_actionable_error(self):
        source = "[ui]\nvoice_text = 测试\n"
        for content in (source.encode("gbk"), source.encode("utf-16"), b"[ui]\nvoice_text=\xff\xfe"):
            with self.subTest(content_length=len(content)), tempfile.TemporaryDirectory() as temporary:
                path = Path(temporary) / "config.ini"
                path.write_bytes(content)
                saved = []
                config = configparser.ConfigParser(interpolation=None)
                with self.assertRaises(ConfigInitializationError) as raised:
                    initialize_config_runtime(
                        config=config, config_file=str(path), defaults_by_section={"ui": {}},
                        save_config=lambda: saved.append(True),
                    )
                self.assertIn("UTF-8", str(raised.exception))
                self.assertIn(str(path), str(raised.exception))
                self.assertEqual(path.read_bytes(), content)
                self.assertEqual(list(Path(temporary).iterdir()), [path])
                self.assertEqual(saved, [])
                self.assertEqual(config.sections(), [])

    def test_unreadable_existing_config_never_saves_defaults(self):
        for error in (PermissionError("denied"), None):
            config = configparser.ConfigParser(interpolation=None)
            with self.subTest(error=type(error).__name__), patch.object(
                config, "read", side_effect=error, return_value=[],
            ):
                saved = []
                with self.assertRaises(ConfigInitializationError):
                    initialize_config_runtime(
                        config=config, config_file="unreadable.ini", defaults_by_section={"ui": {}},
                        path_exists=lambda _path: True, save_config=lambda: saved.append(True),
                    )
                self.assertEqual(saved, [])

    def test_real_bootstrap_shows_error_and_stops_before_starting_services(self):
        with tempfile.TemporaryDirectory() as temporary:
            path = Path(temporary) / "config.ini"
            content = "[ui]\nvoice_text=测试\n".encode("utf-16")
            path.write_bytes(content)
            with patch.object(bootstrap, "_initialize_paths_and_constants"), patch.object(
                bootstrap, "CONFIG_FILE", str(path), create=True,
            ), patch.object(bootstrap.messagebox, "showerror") as show_error, patch.object(
                bootstrap, "_initialize_cloud_settings",
            ) as next_stage:
                self.assertFalse(bootstrap.main())
                show_error.assert_called_once()
                self.assertIn("UTF-8", show_error.call_args.args[1])
                next_stage.assert_not_called()
                self.assertEqual(path.read_bytes(), content)


if __name__ == "__main__":
    unittest.main()
