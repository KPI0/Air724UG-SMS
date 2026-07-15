import configparser
import os
import tempfile
import unittest

from sms_core.config_runtime import (
    initialize_config_runtime,
    read_startup_config_values,
    restore_config_section,
    safe_save_config_runtime,
    snapshot_config_section,
)


class DummyLock:
    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False


class FakeFile:
    def __init__(self, writes):
        self.writes = writes

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        return False

    def write(self, text):
        self.writes.append(text)


class ConfigRuntimeTests(unittest.TestCase):
    def test_config_section_snapshot_restores_values_and_missing_section(self):
        config = configparser.ConfigParser()
        config["serial"] = {"mode": "Manual", "port": "COM3", "extra": "keep"}
        snapshot = snapshot_config_section(config, "serial")

        config.set("serial", "mode", "Auto")
        config.remove_option("serial", "extra")
        config.set("serial", "baud", "115200")
        restore_config_section(config, "serial", snapshot)

        self.assertEqual(dict(config.items("serial", raw=True)), snapshot)

        missing_snapshot = snapshot_config_section(config, "cloud_control")
        config["cloud_control"] = {"enabled": "1"}
        restore_config_section(config, "cloud_control", missing_snapshot)
        self.assertFalse(config.has_section("cloud_control"))

    def test_read_startup_config_values_reads_ui_and_serial_settings(self):
        config = configparser.ConfigParser()
        config["ui"] = {
            "voice_text": " custom ",
            "popup_enabled": "0",
            "auto_log_cleanup": "0",
            "log_retention_days": "12",
            "allow_multi_instance": "1",
            "log_unmatched_sms": "1",
            "voice_enabled": "0",
            "sms_font_size": "24",
            "sms_font_color": " #00ff00 ",
            "keywords": '[" otp ", "bank"]',
            "call_filter_mode": "blacklist",
            "call_whitelist": '["10086"]',
            "call_blacklist": '["10010", " 95588 "]',
        }
        config["serial"] = {
            "port": " COM7 ",
            "baud": "9600",
            "mode": "manual",
        }

        values = read_startup_config_values(config, default_voice_text="default")

        self.assertEqual(values.voice_text, "custom")
        self.assertFalse(values.popup_enabled)
        self.assertFalse(values.auto_log_cleanup)
        self.assertEqual(values.log_retention_days, 12)
        self.assertTrue(values.allow_multi_instance)
        self.assertTrue(values.log_unmatched_sms)
        self.assertFalse(values.voice_enabled)
        self.assertEqual(values.sms_font_size, 24)
        self.assertEqual(values.sms_font_color, "#00ff00")
        self.assertEqual(values.keywords, ["otp", "bank"])
        self.assertEqual(values.call_filter_mode, "Blacklist")
        self.assertEqual(values.call_whitelist, ["10086"])
        self.assertEqual(values.call_blacklist, ["10010", "95588"])
        self.assertEqual(values.port, "COM7")
        self.assertEqual(values.baud, 9600)
        self.assertEqual(values.mode, "Manual")

    def test_read_startup_config_values_uses_fallbacks_and_normalizes_bad_mode(self):
        config = configparser.ConfigParser()
        config["ui"] = {
            "voice_text": " ",
            "popup_enabled": "bad",
            "auto_log_cleanup": "bad",
            "log_retention_days": "bad",
            "allow_multi_instance": "bad",
            "log_unmatched_sms": "bad",
            "voice_enabled": "bad",
            "sms_font_size": "bad",
            "sms_font_color": " ",
            "keywords": "[bad json",
            "call_filter_mode": "unknown",
            "call_whitelist": "bad json",
            "call_blacklist": "bad json",
        }
        config["serial"] = {
            "mode": "unknown",
        }

        values = read_startup_config_values(config, default_voice_text="default")

        self.assertEqual(values.voice_text, "default")
        self.assertTrue(values.popup_enabled)
        self.assertTrue(values.auto_log_cleanup)
        self.assertEqual(values.log_retention_days, 30)
        self.assertFalse(values.allow_multi_instance)
        self.assertFalse(values.log_unmatched_sms)
        self.assertTrue(values.voice_enabled)
        self.assertEqual(values.sms_font_size, 30)
        self.assertEqual(values.sms_font_color, "#ff0000")
        self.assertEqual(values.keywords, ["[bad json"])
        self.assertEqual(values.call_filter_mode, "Disabled")
        self.assertEqual(values.call_whitelist, [])
        self.assertEqual(values.call_blacklist, [])
        self.assertEqual(values.port, "")
        self.assertEqual(values.baud, 115200)
        self.assertEqual(values.mode, "Auto")

    def test_read_startup_config_values_logs_invalid_values_and_keeps_starting(self):
        config = configparser.ConfigParser()
        config["ui"] = {
            "popup_enabled": "bad",
            "call_whitelist": "{bad json",
        }
        config["serial"] = {
            "baud": "bad",
            "mode": "strange",
        }
        logs = []

        values = read_startup_config_values(
            config,
            default_voice_text="default",
            log_error=logs.append,
        )

        self.assertTrue(values.popup_enabled)
        self.assertEqual(values.call_whitelist, [])
        self.assertEqual(values.baud, 115200)
        self.assertEqual(values.mode, "Auto")
        self.assertTrue(any("ui.popup_enabled" in message for message in logs))
        self.assertTrue(any("ui.call_whitelist" in message for message in logs))
        self.assertTrue(any("serial.baud" in message for message in logs))
        self.assertTrue(any("serial.mode" in message for message in logs))

    def test_read_startup_config_values_rejects_non_positive_baud(self):
        config = configparser.ConfigParser()
        config["serial"] = {"baud": "0"}
        logs = []

        values = read_startup_config_values(
            config,
            default_voice_text="default",
            log_error=logs.append,
        )

        self.assertEqual(values.baud, 115200)
        self.assertTrue(any("must be positive" in message for message in logs))

    def test_initialize_config_runtime_creates_defaults_when_missing(self):
        config = configparser.ConfigParser()
        calls = []
        missing_config = tempfile.NamedTemporaryFile(delete=True)
        missing_config.close()

        result = initialize_config_runtime(
            config=config,
            config_file=missing_config.name,
            defaults_by_section={
                "serial": {"baud": "115200"},
                "ui": {"voice_enabled": "1"},
            },
            save_config=lambda: calls.append("save"),
            path_exists=lambda path: False,
        )

        self.assertTrue(result)
        self.assertEqual(calls, ["save"])
        self.assertEqual(config.get("serial", "baud"), "115200")
        self.assertEqual(config.get("ui", "voice_enabled"), "1")

    def test_initialize_config_runtime_reads_existing_config_without_saving(self):
        class FakeConfig:
            def __init__(self):
                self.sections = {}
                self.read_calls = []

            def __setitem__(self, key, value):
                self.sections[key] = value

            def read(self, path, encoding):
                self.read_calls.append((path, encoding))

        config = FakeConfig()
        calls = []

        result = initialize_config_runtime(
            config=config,
            config_file="config.ini",
            defaults_by_section={"ui": {"voice_enabled": "1"}},
            save_config=lambda: calls.append("save"),
            path_exists=lambda path: True,
        )

        self.assertFalse(result)
        self.assertEqual(calls, [])
        self.assertEqual(config.sections, {})
        self.assertEqual(config.read_calls, [("config.ini", "utf-8-sig")])

    def test_initialize_config_runtime_accepts_utf8_bom_config(self):
        with tempfile.TemporaryDirectory() as tmp:
            config_file = os.path.join(tmp, "config.ini")
            with open(config_file, "w", encoding="utf-8-sig") as file:
                file.write("[ui]\nvoice_enabled = 1\n")

            config = configparser.ConfigParser(interpolation=None)

            result = initialize_config_runtime(
                config=config,
                config_file=config_file,
                defaults_by_section={"ui": {"voice_enabled": "0"}},
                save_config=lambda: None,
            )

            self.assertFalse(result)
            self.assertEqual(config.get("ui", "voice_enabled"), "1")

    def test_initialize_config_runtime_recovers_corrupt_config_with_backup(self):
        with tempfile.TemporaryDirectory() as tmp:
            config_file = os.path.join(tmp, "config.ini")
            backup_file = os.path.join(tmp, "config.ini.broken.123.bak")
            with open(config_file, "w", encoding="utf-8") as file:
                file.write("this is not an ini file")

            config = configparser.ConfigParser(interpolation=None)
            logs = []

            def save_config():
                with open(config_file, "w", encoding="utf-8") as file:
                    config.write(file)

            result = initialize_config_runtime(
                config=config,
                config_file=config_file,
                defaults_by_section={"ui": {"voice_enabled": "1"}},
                save_config=save_config,
                backup_file=lambda src, dst: os.replace(src, backup_file),
                time_func=lambda: 123,
                log_error=logs.append,
            )

            self.assertTrue(result)
            self.assertTrue(os.path.exists(backup_file))
            self.assertEqual(config.get("ui", "voice_enabled"), "1")
            self.assertTrue(any("Config file invalid" in message for message in logs))

    def test_safe_save_config_runtime_writes_temp_and_replaces_target(self):
        config = configparser.ConfigParser()
        config["ui"] = {"voice_enabled": "1"}
        opened = []
        writes = []
        replaced = []

        result = safe_save_config_runtime(
            config=config,
            config_file="config.ini",
            config_lock=DummyLock(),
            getpid=lambda: 10,
            get_thread_id=lambda: 20,
            open_file=lambda path, mode, encoding: opened.append((path, mode, encoding)) or FakeFile(writes),
            replace_file=lambda src, dst: replaced.append((src, dst)),
        )

        self.assertTrue(result)
        self.assertEqual(opened, [("config.ini.10.20.tmp", "w", "utf-8")])
        self.assertEqual(replaced, [("config.ini.10.20.tmp", "config.ini")])
        self.assertTrue(any("[ui]" in item for item in writes))

    def test_safe_save_config_runtime_removes_temp_and_logs_on_error(self):
        logs = []
        removed = []

        result = safe_save_config_runtime(
            config=configparser.ConfigParser(),
            config_file="config.ini",
            config_lock=DummyLock(),
            log_error=logs.append,
            getpid=lambda: 10,
            get_thread_id=lambda: 20,
            open_file=lambda *_args, **_kwargs: (_ for _ in ()).throw(RuntimeError("disk")),
            path_exists=lambda path: True,
            remove_file=lambda path: removed.append(path),
        )

        self.assertFalse(result)
        self.assertEqual(removed, ["config.ini.10.20.tmp"])
        self.assertIn("disk", logs[0])

    def test_safe_save_config_runtime_swallows_cleanup_and_log_errors(self):
        result = safe_save_config_runtime(
            config=configparser.ConfigParser(),
            config_file="config.ini",
            config_lock=DummyLock(),
            log_error=lambda message: (_ for _ in ()).throw(RuntimeError("log")),
            getpid=lambda: 10,
            get_thread_id=lambda: 20,
            open_file=lambda *_args, **_kwargs: (_ for _ in ()).throw(RuntimeError("disk")),
            path_exists=lambda path: True,
            remove_file=lambda path: (_ for _ in ()).throw(RuntimeError("remove")),
        )

        self.assertFalse(result)


if __name__ == "__main__":
    unittest.main()
