import configparser
import os
import tempfile
import threading
import unittest

from sms_core.config_runtime import (
    CONFIG_SNAPSHOT_ATTR,
    ConfigInitializationError,
    ensure_config_defaults,
    initialize_config_runtime,
    load_config_snapshot,
    remember_config_snapshot,
    read_startup_config_values,
    reload_config_runtime,
    restore_config_section,
    safe_save_config_runtime,
    snapshot_config_section,
    snapshot_config_runtime,
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
    def test_ensure_config_defaults_fills_all_missing_values_without_overwriting(self):
        config = configparser.ConfigParser(interpolation=None)
        config["ui"] = {
            "voice_enabled": "0",
            "unknown_key": "keep",
        }
        defaults = {
            "ui": {
                "voice_enabled": "1",
                "popup_enabled": "1",
            },
            "cloud_control": {
                "enabled": "0",
            },
        }

        self.assertTrue(ensure_config_defaults(config, defaults))
        self.assertEqual(config.get("ui", "voice_enabled"), "0")
        self.assertEqual(config.get("ui", "popup_enabled"), "1")
        self.assertEqual(config.get("ui", "unknown_key"), "keep")
        self.assertEqual(config.get("cloud_control", "enabled"), "0")
        self.assertFalse(ensure_config_defaults(config, defaults))

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

    def test_load_config_snapshot_rejects_missing_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(FileNotFoundError):
                load_config_snapshot(os.path.join(tmp, "missing.ini"))

    def test_reload_config_runtime_removes_disk_deletions_and_preserves_local_changes(self):
        config = configparser.ConfigParser(interpolation=None)
        config["ui"] = {
            "unchanged": "old",
            "stale": "remove-me",
            "local": "old",
        }
        remember_config_snapshot(config)
        config.set("ui", "local", "pending")
        disk_snapshot = {
            "ui": {
                "unchanged": "disk",
                "local": "old",
                "external": "kept",
            }
        }

        values = reload_config_runtime(
            config=config,
            config_file="config.ini",
            config_lock=threading.RLock(),
            load_snapshot=lambda _path: disk_snapshot,
            read_values=lambda current: dict(current.items("ui", raw=True)),
        )

        self.assertEqual(values["unchanged"], "disk")
        self.assertEqual(values["local"], "pending")
        self.assertEqual(values["external"], "kept")
        self.assertNotIn("stale", values)
        self.assertEqual(getattr(config, CONFIG_SNAPSHOT_ATTR), disk_snapshot)

    def test_reload_config_runtime_rolls_back_config_and_baseline_when_commit_fails(self):
        config = configparser.ConfigParser(interpolation=None)
        config["third_push"] = {"enabled": "1", "old": "keep"}
        remember_config_snapshot(config)
        before = snapshot_config_runtime(config)
        before_baseline = getattr(config, CONFIG_SNAPSHOT_ATTR)

        with self.assertRaises(RuntimeError):
            reload_config_runtime(
                config=config,
                config_file="config.ini",
                config_lock=threading.RLock(),
                load_snapshot=lambda _path: {"third_push": {"enabled": "0"}},
                prepare_config=lambda staged: staged.set("third_push", "defaulted", "1"),
                commit_config=lambda: False,
            )

        self.assertEqual(snapshot_config_runtime(config), before)
        self.assertEqual(getattr(config, CONFIG_SNAPSHOT_ATTR), before_baseline)

    def test_reload_config_runtime_parse_failure_leaves_shared_config_untouched(self):
        config = configparser.ConfigParser(interpolation=None)
        config["cloud_control"] = {"enabled": "1", "url": "ws://old"}
        remember_config_snapshot(config)
        before = snapshot_config_runtime(config)
        before_baseline = getattr(config, CONFIG_SNAPSHOT_ATTR)

        with self.assertRaises(configparser.ParsingError):
            reload_config_runtime(
                config=config,
                config_file="config.ini",
                config_lock=threading.RLock(),
                load_snapshot=lambda _path: (_ for _ in ()).throw(
                    configparser.ParsingError("config.ini")
                ),
            )

        self.assertEqual(snapshot_config_runtime(config), before)
        self.assertEqual(getattr(config, CONFIG_SNAPSHOT_ATTR), before_baseline)

    def test_read_startup_config_values_reads_ui_and_serial_settings(self):
        config = configparser.ConfigParser()
        config["ui"] = {
            "voice_text": " custom ",
            "popup_enabled": "0",
            "call_popup_enabled": "0",
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
        self.assertFalse(values.call_popup_enabled)
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

    def test_old_config_without_call_popup_setting_keeps_existing_behavior(self):
        config = configparser.ConfigParser()
        config["ui"] = {"popup_enabled": "0"}

        values = read_startup_config_values(config, default_voice_text="default")

        self.assertFalse(values.popup_enabled)
        self.assertTrue(values.call_popup_enabled)

    def test_read_startup_config_values_uses_fallbacks_and_normalizes_bad_mode(self):
        config = configparser.ConfigParser()
        config["ui"] = {
            "voice_text": " ",
            "popup_enabled": "bad",
            "call_popup_enabled": "bad",
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
        self.assertTrue(values.call_popup_enabled)
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
            "call_popup_enabled": "bad",
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
        self.assertTrue(values.call_popup_enabled)
        self.assertEqual(values.call_whitelist, [])
        self.assertEqual(values.baud, 115200)
        self.assertEqual(values.mode, "Auto")
        self.assertTrue(any("ui.popup_enabled" in message for message in logs))
        self.assertTrue(any("ui.call_popup_enabled" in message for message in logs))
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

    def test_read_startup_config_values_rejects_unsafe_ui_numeric_ranges(self):
        config = configparser.ConfigParser()
        config["ui"] = {
            "log_retention_days": "-1",
            "sms_font_size": "100000",
        }
        logs = []

        values = read_startup_config_values(
            config,
            default_voice_text="default",
            log_error=logs.append,
        )

        self.assertEqual(values.log_retention_days, 30)
        self.assertEqual(values.sms_font_size, 30)
        self.assertTrue(any("log_retention_days" in message for message in logs))
        self.assertTrue(any("sms_font_size" in message for message in logs))

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

    def test_initialize_config_runtime_stops_when_missing_config_cannot_be_saved(self):
        config = configparser.ConfigParser()
        logs = []

        with self.assertRaises(ConfigInitializationError) as raised:
            initialize_config_runtime(
                config=config,
                config_file="config.ini",
                defaults_by_section={"ui": {"voice_enabled": "1"}},
                save_config=lambda: False,
                path_exists=lambda _path: False,
                log_error=logs.append,
            )

        self.assertIn("创建失败", str(raised.exception))
        self.assertEqual(config.get("ui", "voice_enabled"), "1")
        self.assertTrue(any("returned False" in message for message in logs))

    def test_initialize_config_runtime_reads_existing_config_without_saving(self):
        class TrackingConfig(configparser.ConfigParser):
            def __init__(self):
                super().__init__(interpolation=None)
                self.read_calls = []

            def read(self, path, encoding):
                self.read_calls.append((path, encoding))
                self["ui"] = {"voice_enabled": "0"}

        config = TrackingConfig()
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
        self.assertEqual(config.get("ui", "voice_enabled"), "0")
        self.assertEqual(config.read_calls, [("config.ini", "utf-8-sig")])

    def test_initialize_config_runtime_persists_all_missing_defaults_in_old_config(self):
        with tempfile.TemporaryDirectory() as tmp:
            config_file = os.path.join(tmp, "config.ini")
            with open(config_file, "w", encoding="utf-8") as file:
                file.write("[ui]\nvoice_enabled = 0\nunknown_key = keep\n")

            config = configparser.ConfigParser(interpolation=None)
            saves = []

            def save_config():
                saves.append("save")
                with open(config_file, "w", encoding="utf-8") as file:
                    config.write(file)
                return True

            result = initialize_config_runtime(
                config=config,
                config_file=config_file,
                defaults_by_section={
                    "ui": {
                        "voice_enabled": "1",
                        "popup_enabled": "1",
                    },
                    "cloud_control": {
                        "enabled": "0",
                    },
                },
                save_config=save_config,
            )

            self.assertFalse(result)
            self.assertEqual(saves, ["save"])
            self.assertEqual(config.get("ui", "voice_enabled"), "0")
            self.assertEqual(config.get("ui", "popup_enabled"), "1")
            self.assertEqual(config.get("ui", "unknown_key"), "keep")
            self.assertEqual(config.get("cloud_control", "enabled"), "0")

            disk = configparser.ConfigParser(interpolation=None)
            disk.read(config_file, encoding="utf-8")
            self.assertEqual(dict(disk.items("ui", raw=True)), {
                "voice_enabled": "0",
                "unknown_key": "keep",
                "popup_enabled": "1",
            })
            self.assertEqual(disk.get("cloud_control", "enabled"), "0")

    def test_initialize_config_runtime_restores_old_config_when_completion_save_fails(self):
        class TrackingConfig(configparser.ConfigParser):
            def read(self, path, encoding=None):
                self["ui"] = {
                    "voice_enabled": "0",
                    "unknown_key": "keep",
                }
                return [path]

        config = TrackingConfig(interpolation=None)

        with self.assertRaises(ConfigInitializationError) as raised:
            initialize_config_runtime(
                config=config,
                config_file="config.ini",
                defaults_by_section={
                    "ui": {
                        "voice_enabled": "1",
                        "popup_enabled": "1",
                    },
                    "cloud_control": {
                        "enabled": "0",
                    },
                },
                save_config=lambda: False,
                path_exists=lambda _path: True,
            )

        self.assertIn("补齐失败", str(raised.exception))
        self.assertEqual(dict(config.items("ui", raw=True)), {
            "voice_enabled": "0",
            "unknown_key": "keep",
        })
        self.assertFalse(config.has_section("cloud_control"))

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
            secret_marker = "DEVICE_SECRET_SUPER_PRIVATE"
            with open(config_file, "w", encoding="utf-8") as file:
                file.write(secret_marker + " = should-not-enter-logs")

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
            self.assertFalse(any(secret_marker in message for message in logs))

    def test_initialize_config_runtime_stops_when_repaired_config_cannot_be_saved(self):
        class BrokenConfig(configparser.ConfigParser):
            def read(self, filenames, encoding=None):
                raise configparser.MissingSectionHeaderError(str(filenames), 1, "broken")

        config = BrokenConfig(interpolation=None)
        backups = []

        with self.assertRaises(ConfigInitializationError) as raised:
            initialize_config_runtime(
                config=config,
                config_file="config.ini",
                defaults_by_section={"ui": {"voice_enabled": "1"}},
                save_config=lambda: False,
                path_exists=lambda _path: True,
                backup_file=lambda source, target: backups.append((source, target)),
                time_func=lambda: 123,
            )

        self.assertIn("修复失败", str(raised.exception))
        self.assertEqual(backups, [("config.ini", "config.ini.broken.123.bak")])
        self.assertEqual(config.get("ui", "voice_enabled"), "1")

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

    def test_safe_save_acquires_process_mutex_while_holding_config_lock(self):
        class TrackingLock:
            def __init__(self):
                self.held = False

            def __enter__(self):
                self.held = True
                return self

            def __exit__(self, exc_type, exc, tb):
                self.held = False
                return False

        lock = TrackingLock()
        process_lock_checks = []

        result = safe_save_config_runtime(
            config=configparser.ConfigParser(),
            config_file="config.ini",
            config_lock=lock,
            open_file=lambda *_args, **_kwargs: FakeFile([]),
            replace_file=lambda *_args: None,
            acquire_process_lock=lambda *_args, **_kwargs: process_lock_checks.append(lock.held) or ("mutex", 0),
            release_process_lock=lambda _lock: None,
        )

        self.assertTrue(result)
        self.assertEqual(process_lock_checks, [True])

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

    def test_safe_save_config_runtime_merges_stale_multi_instance_changes(self):
        with tempfile.TemporaryDirectory() as tmp:
            config_file = os.path.join(tmp, "config.ini")
            initial = configparser.ConfigParser(interpolation=None)
            initial["ui"] = {
                "voice_enabled": "1",
                "popup_enabled": "1",
                "unknown_key": "keep",
            }
            self.assertTrue(
                safe_save_config_runtime(
                    config=initial,
                    config_file=config_file,
                    config_lock=DummyLock(),
                )
            )

            instance_a = configparser.ConfigParser(interpolation=None)
            instance_b = configparser.ConfigParser(interpolation=None)
            instance_a.read(config_file, encoding="utf-8")
            instance_b.read(config_file, encoding="utf-8")
            remember_config_snapshot(instance_a)
            remember_config_snapshot(instance_b)

            instance_a.set("ui", "voice_enabled", "0")
            self.assertTrue(
                safe_save_config_runtime(
                    config=instance_a,
                    config_file=config_file,
                    config_lock=DummyLock(),
                )
            )

            instance_b.set("ui", "popup_enabled", "0")
            self.assertTrue(
                safe_save_config_runtime(
                    config=instance_b,
                    config_file=config_file,
                    config_lock=DummyLock(),
                )
            )

            final = configparser.ConfigParser(interpolation=None)
            final.read(config_file, encoding="utf-8")
            self.assertEqual(final.get("ui", "voice_enabled"), "0")
            self.assertEqual(final.get("ui", "popup_enabled"), "0")
            self.assertEqual(final.get("ui", "unknown_key"), "keep")
            self.assertEqual(instance_b.get("ui", "voice_enabled"), "0")

    def test_safe_save_default_completion_preserves_newer_disk_values(self):
        with tempfile.TemporaryDirectory() as tmp:
            config_file = os.path.join(tmp, "config.ini")
            with open(config_file, "w", encoding="utf-8") as file:
                file.write("[ui]\nvoice_enabled = 1\n")

            upgrading = configparser.ConfigParser(interpolation=None)
            upgrading.read(config_file, encoding="utf-8")
            remember_config_snapshot(upgrading)
            ensure_config_defaults(
                upgrading,
                {
                    "ui": {
                        "voice_enabled": "1",
                        "popup_enabled": "1",
                        "sms_font_size": "30",
                    },
                    "cloud_control": {
                        "enabled": "0",
                        "device_secret": "",
                    },
                },
            )

            with open(config_file, "w", encoding="utf-8") as file:
                file.write(
                    "[ui]\n"
                    "voice_enabled = 0\n"
                    "popup_enabled = 0\n"
                    "unknown_key = keep\n"
                    "[cloud_control]\n"
                    "enabled = 1\n"
                )

            self.assertTrue(
                safe_save_config_runtime(
                    config=upgrading,
                    config_file=config_file,
                    config_lock=DummyLock(),
                    defaults_by_section={
                        "ui": {
                            "voice_enabled": "1",
                            "popup_enabled": "1",
                            "sms_font_size": "30",
                        },
                        "cloud_control": {
                            "enabled": "0",
                            "device_secret": "",
                        },
                    },
                )
            )

            final = configparser.ConfigParser(interpolation=None)
            final.read(config_file, encoding="utf-8")
            self.assertEqual(final.get("ui", "voice_enabled"), "0")
            self.assertEqual(final.get("ui", "popup_enabled"), "0")
            self.assertEqual(final.get("ui", "sms_font_size"), "30")
            self.assertEqual(final.get("ui", "unknown_key"), "keep")
            self.assertEqual(final.get("cloud_control", "enabled"), "1")
            self.assertTrue(final.has_option("cloud_control", "device_secret"))


if __name__ == "__main__":
    unittest.main()
