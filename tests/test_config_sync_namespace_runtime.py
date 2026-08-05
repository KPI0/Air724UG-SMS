import configparser
import threading
import tkinter as tk
import unittest

from sms_ui.config_sync_namespace_runtime import (
    register_config_sync_refresher_namespace_runtime,
    reload_shared_ui_config_namespace_runtime,
    start_config_file_watch_namespace_runtime,
)
from sms_ui.config_sync_runtime import ConfigFileWatchState
from sms_core.config_runtime import CONFIG_SNAPSHOT_ATTR, remember_config_snapshot


class FakeVar:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeRoot:
    def winfo_rgb(self, color):
        if color == "not-a-color":
            raise tk.TclError("unknown color name")
        return (0, 0, 0)

    def after(self, *_args):
        return "after-id"

    def after_cancel(self, *_args):
        return None


class FakeStopEvent:
    def is_set(self):
        return False


class ConfigSyncNamespaceRuntimeTests(unittest.TestCase):
    def make_namespace(self):
        config = configparser.ConfigParser(interpolation=None)
        config["ui"] = {
            "popup_enabled": "1",
            "call_popup_enabled": "1",
            "local_only": "old",
        }
        remember_config_snapshot(config)
        calls = []
        return {
            "config": config,
            "CONFIG_FILE": "config.ini",
            "CONFIG_LOCK": threading.RLock(),
            "DEFAULT_VOICE_TEXT": "default voice",
            "POPUP_ENABLED": True,
            "CALL_POPUP_ENABLED": True,
            "VOICE_ENABLED": True,
            "VOICE_TEXT": "old voice",
            "SMS_FONT_SIZE": 30,
            "SMS_FONT_COLOR": "#ff0000",
            "KEYWORDS": ["old"],
            "LOG_UNMATCHED_SMS": False,
            "CALL_FILTER_MODE": "Disabled",
            "CALL_WHITELIST": ["10086"],
            "CALL_BLACKLIST": [],
            "AUTO_LOG_CLEANUP": True,
            "LOG_RETENTION_DAYS": 30,
            "ALLOW_MULTI_INSTANCE": True,
            "CLOUD_SENSITIVE_COMMAND_PERMISSIONS": {
                "sms": False,
                "call": False,
                "pin": False,
                "puk": False,
                "phone_number": False,
                "sn": False,
                "cell_location": False,
            },
            "PORT": "COM1",
            "BAUD": 115200,
            "MODE": "Manual",
            "popup_var": FakeVar(True),
            "call_popup_var": FakeVar(True),
            "multi_instance_var": FakeVar(True),
            "update_voice_menu_label": lambda: calls.append("voice_label"),
            "generate_alert_voice": lambda **kwargs: calls.append(("voice", kwargs)),
            "apply_sms_font_style": lambda: calls.append("font"),
            "schedule_auto_log_cleanup": lambda **kwargs: calls.append(("cleanup", kwargs)),
            "system_ui": lambda *args: calls.append(("ui", args)),
            "log_file_only": lambda message: calls.append(("log", message)),
            "close_call_popup": lambda: calls.append("close_call_popup"),
            "close_missed_call_popup": lambda: calls.append("close_missed_call_popup"),
            "calls": calls,
            "root": FakeRoot(),
            "tk_alive": lambda: True,
            "TK_SHUTDOWN": FakeStopEvent(),
            "is_exiting": False,
            "CONFIG_FILE_WATCH_STATE": ConfigFileWatchState(),
            "reload_shared_ui_config": lambda: calls.append("reload"),
        }

    def disk_snapshot(self):
        return {
            "ui": {
                "voice_text": "new voice",
                "popup_enabled": "0",
                "call_popup_enabled": "0",
                "auto_log_cleanup": "0",
                "log_retention_days": "7",
                "allow_multi_instance": "0",
                "log_unmatched_sms": "1",
                "voice_enabled": "0",
                "sms_font_size": "24",
                "sms_font_color": "#123456",
                "keywords": '["otp", "bank"]',
                "call_filter_mode": "Whitelist",
                "call_whitelist": '["10010"]',
                "call_blacklist": '["10000"]',
                "external_only": "kept",
            },
            "serial": {"mode": "Auto", "port": "COM9", "baud": "9600"},
        }

    def test_reload_updates_shared_ui_state_and_preserves_serial_runtime(self):
        namespace = self.make_namespace()
        keywords = namespace["KEYWORDS"]
        whitelist = namespace["CALL_WHITELIST"]
        blacklist = namespace["CALL_BLACKLIST"]

        changed = reload_shared_ui_config_namespace_runtime(
            namespace,
            load_snapshot=lambda _path: self.disk_snapshot(),
        )

        self.assertEqual(
            changed,
            (
                "短信弹窗",
                "电话弹窗",
                "语音播报",
                "语音内容",
                "短信字体",
                "关键词",
                "防骚扰",
                "日志清理",
                "程序多开",
            ),
        )
        self.assertFalse(namespace["POPUP_ENABLED"])
        self.assertFalse(namespace["popup_var"].get())
        self.assertFalse(namespace["CALL_POPUP_ENABLED"])
        self.assertFalse(namespace["call_popup_var"].get())
        self.assertIn("close_call_popup", namespace["calls"])
        self.assertIn("close_missed_call_popup", namespace["calls"])
        self.assertFalse(namespace["VOICE_ENABLED"])
        self.assertEqual(namespace["VOICE_TEXT"], "new voice")
        self.assertEqual((namespace["SMS_FONT_SIZE"], namespace["SMS_FONT_COLOR"]), (24, "#123456"))
        self.assertIs(namespace["KEYWORDS"], keywords)
        self.assertEqual(keywords, ["otp", "bank"])
        self.assertIs(namespace["CALL_WHITELIST"], whitelist)
        self.assertEqual(whitelist, ["10010"])
        self.assertIs(namespace["CALL_BLACKLIST"], blacklist)
        self.assertEqual(blacklist, ["10000"])
        self.assertEqual(namespace["CALL_FILTER_MODE"], "Whitelist")
        self.assertFalse(namespace["AUTO_LOG_CLEANUP"])
        self.assertEqual(namespace["LOG_RETENTION_DAYS"], 7)
        self.assertFalse(namespace["ALLOW_MULTI_INSTANCE"])
        self.assertFalse(namespace["multi_instance_var"].get())
        self.assertEqual((namespace["PORT"], namespace["BAUD"], namespace["MODE"]), ("COM1", 115200, "Manual"))
        self.assertEqual(namespace["config"].get("serial", "port"), "COM9")
        self.assertEqual(namespace["config"].get("ui", "external_only"), "kept")
        self.assertIn("voice_label", namespace["calls"])
        self.assertIn(("voice", {"force": True}), namespace["calls"])
        self.assertIn("font", namespace["calls"])
        self.assertIn(("cleanup", {"restart": True, "first_delay_sec": 60}), namespace["calls"])
        self.assertTrue(any(call[0] == "ui" and "短信弹窗" in call[1][0] for call in namespace["calls"] if isinstance(call, tuple)))
        self.assertTrue(any(call[0] == "ui" and "电话弹窗" in call[1][0] for call in namespace["calls"] if isinstance(call, tuple)))

    def test_reload_failure_preserves_existing_config_and_runtime(self):
        namespace = self.make_namespace()
        before = dict(namespace["config"].items("ui"))

        result = reload_shared_ui_config_namespace_runtime(
            namespace,
            load_snapshot=lambda _path: (_ for _ in ()).throw(ValueError("invalid config")),
        )

        self.assertFalse(result)
        self.assertEqual(dict(namespace["config"].items("ui")), before)
        self.assertTrue(namespace["POPUP_ENABLED"])
        failure_logs = [
            call[1]
            for call in namespace["calls"]
            if isinstance(call, tuple) and call[0] == "log"
        ]
        self.assertEqual(failure_logs, ["Reload shared UI config failed (ValueError)"])
        self.assertNotIn("invalid config", failure_logs[0])

    def test_reload_rejects_invalid_font_color_and_retains_valid_runtime_color(self):
        namespace = self.make_namespace()
        snapshot = self.disk_snapshot()
        snapshot["ui"]["sms_font_size"] = "30"
        snapshot["ui"]["sms_font_color"] = "not-a-color"

        changed = reload_shared_ui_config_namespace_runtime(
            namespace,
            load_snapshot=lambda _path: snapshot,
        )

        self.assertNotIn("短信字体", changed)
        self.assertEqual(namespace["SMS_FONT_COLOR"], "#ff0000")
        self.assertEqual(namespace["config"].get("ui", "sms_font_color"), "#ff0000")
        self.assertNotIn("font", namespace["calls"])
        repair_logs = [
            item[1]
            for item in namespace["calls"]
            if isinstance(item, tuple)
            and item[0] == "log"
            and "invalid synced SMS font color" in item[1]
        ]
        self.assertEqual(len(repair_logs), 1)
        self.assertNotIn("not-a-color", repair_logs[0])

    def test_reload_rejects_unsafe_font_size_and_log_retention_ranges(self):
        namespace = self.make_namespace()
        snapshot = self.disk_snapshot()
        snapshot["ui"]["sms_font_color"] = "#ff0000"
        snapshot["ui"]["auto_log_cleanup"] = "1"
        snapshot["ui"]["sms_font_size"] = "100000"
        snapshot["ui"]["log_retention_days"] = "-1"

        changed = reload_shared_ui_config_namespace_runtime(
            namespace,
            load_snapshot=lambda _path: snapshot,
        )

        self.assertEqual(namespace["SMS_FONT_SIZE"], 30)
        self.assertEqual(namespace["LOG_RETENTION_DAYS"], 30)
        self.assertNotIn("短信字体", changed)
        self.assertNotIn("日志清理", changed)
        logs = [
            item[1]
            for item in namespace["calls"]
            if isinstance(item, tuple) and item[0] == "log"
        ]
        self.assertTrue(any("sms_font_size" in message for message in logs))
        self.assertTrue(any("log_retention_days" in message for message in logs))

    def test_reload_syncs_cloud_sensitive_command_setting(self):
        namespace = self.make_namespace()
        snapshot = self.disk_snapshot()
        snapshot["cloud_control"] = {"allow_sensitive_commands": "1"}

        changed = reload_shared_ui_config_namespace_runtime(
            namespace,
            load_snapshot=lambda _path: snapshot,
        )

        self.assertIn("安全设置", changed)
        self.assertTrue(all(
            namespace["CLOUD_SENSITIVE_COMMAND_PERMISSIONS"].values()
        ))

    def test_reload_repeated_failure_is_suppressed_and_success_logs_recovery(self):
        namespace = self.make_namespace()
        namespace["time"] = type("FakeTime", (), {"monotonic": staticmethod(lambda: 0.0)})

        for _ in range(5):
            self.assertFalse(
                reload_shared_ui_config_namespace_runtime(
                    namespace,
                    load_snapshot=lambda _path: (_ for _ in ()).throw(ValueError("device_secret=secret")),
                )
            )

        failure_logs = [
            call[1]
            for call in namespace["calls"]
            if isinstance(call, tuple) and call[0] == "log"
        ]
        self.assertEqual(failure_logs, ["Reload shared UI config failed (ValueError)"])
        self.assertTrue(
            reload_shared_ui_config_namespace_runtime(
                namespace,
                load_snapshot=lambda _path: self.disk_snapshot(),
            )
        )
        all_logs = [
            call[1]
            for call in namespace["calls"]
            if isinstance(call, tuple) and call[0] == "log"
        ]
        self.assertEqual(all_logs[-1], "Reload shared UI config recovered after 5 failed attempts")
        self.assertTrue(all("device_secret" not in message for message in all_logs))

    def test_reload_preserves_pending_local_serial_change_and_uses_disk_as_new_baseline(self):
        namespace = self.make_namespace()
        namespace["config"]["serial"] = {"mode": "Manual", "port": "COM1", "baud": "115200"}
        remember_config_snapshot(namespace["config"])
        namespace["config"].set("serial", "port", "COM9")
        disk_snapshot = {
            "ui": {"popup_enabled": "0"},
            "serial": {"mode": "Manual", "port": "COM1", "baud": "115200"},
        }

        changed = reload_shared_ui_config_namespace_runtime(
            namespace,
            load_snapshot=lambda _path: disk_snapshot,
        )

        self.assertIn("短信弹窗", changed)
        self.assertEqual(namespace["config"].get("serial", "port"), "COM9")
        baseline = getattr(namespace["config"], CONFIG_SNAPSHOT_ATTR)
        self.assertEqual(baseline["serial"]["port"], "COM1")
        self.assertFalse(namespace["POPUP_ENABLED"])

    def test_reload_notifies_open_settings_refreshers_and_unregisters_cleanly(self):
        namespace = self.make_namespace()
        refreshes = []
        unregister_keywords = register_config_sync_refresher_namespace_runtime(
            namespace,
            "keywords",
            lambda: refreshes.append("keywords"),
        )
        register_config_sync_refresher_namespace_runtime(
            namespace,
            "call_filter",
            lambda: refreshes.append("call_filter"),
        )

        reload_shared_ui_config_namespace_runtime(
            namespace,
            load_snapshot=lambda _path: self.disk_snapshot(),
        )

        self.assertEqual(refreshes, ["keywords", "call_filter"])
        unregister_keywords()
        namespace["KEYWORDS"][:] = ["stale"]
        reload_shared_ui_config_namespace_runtime(
            namespace,
            load_snapshot=lambda _path: self.disk_snapshot(),
        )
        self.assertEqual(refreshes, ["keywords", "call_filter"])

    def test_start_watch_forwards_ui_lifecycle_and_callback(self):
        namespace = self.make_namespace()
        captured = {}

        def schedule_runtime(**kwargs):
            captured.update(kwargs)
            return "watch-id"

        result = start_config_file_watch_namespace_runtime(
            namespace,
            interval_ms=2500,
            schedule_runtime=schedule_runtime,
        )

        self.assertEqual(result, "watch-id")
        self.assertIs(captured["state"], namespace["CONFIG_FILE_WATCH_STATE"])
        self.assertEqual(captured["config_file"], "config.ini")
        self.assertEqual(captured["interval_ms"], 2500)
        self.assertFalse(captured["is_stopping"]())
        captured["on_change"]()
        self.assertIn("reload", namespace["calls"])


if __name__ == "__main__":
    unittest.main()
