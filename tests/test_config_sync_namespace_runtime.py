import configparser
import os
import queue
import tempfile
import threading
import tkinter as tk
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from sms_ui.config_sync_namespace_runtime import (
    register_config_sync_refresher_namespace_runtime,
    reload_shared_ui_config_namespace_runtime,
    start_config_file_watch_namespace_runtime,
)
from sms_ui.config_sync_runtime import ConfigFileWatchState
from sms_core.config_runtime import (
    CONFIG_SNAPSHOT_ATTR,
    remember_config_snapshot,
    safe_save_config_runtime,
    snapshot_config_runtime,
)
from sms_core.third_push import THIRD_PUSH_SETTINGS_KEYS, third_push_state
from sms_core.third_push_config import read_third_push_settings
from sms_ui.config_sync_namespace_bindings import install_config_sync_namespace_bindings
from sms_ui.third_push_namespace_bindings import install_third_push_namespace_bindings
from sms_ui.third_push_namespace_runtime import apply_third_push_settings_namespace_runtime
from tests.test_config_sync_runtime import FakeRootTimer
from tests.test_third_push_window_form import make_form


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
        push_settings = read_third_push_settings(config)
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
            "THIRD_PUSH_ENABLED": push_settings.enabled,
            "THIRD_PUSH_SMS_ENABLED": push_settings.sms_enabled,
            "THIRD_PUSH_CALL_ENABLED": push_settings.call_enabled,
            "THIRD_PUSH_TYPES": push_settings.channels,
            "THIRD_PUSH_SETTINGS": push_settings.settings,
            "third_push_win": None,
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

    def test_reload_notifies_security_window_after_permissions_change(self):
        namespace = self.make_namespace()
        observed = []
        unregister = register_config_sync_refresher_namespace_runtime(
            namespace, "security",
            lambda: observed.append(dict(namespace["CLOUD_SENSITIVE_COMMAND_PERMISSIONS"])),
        )
        snapshot = self.disk_snapshot()
        snapshot["cloud_control"] = {"allow_sensitive_sn": "1"}

        def load_snapshot(_path):
            return {section: dict(values) for section, values in snapshot.items()}

        reload_shared_ui_config_namespace_runtime(namespace, load_snapshot=load_snapshot)
        self.assertEqual(len(observed), 1)
        self.assertTrue(observed[0]["sn"])
        snapshot["cloud_control"]["allow_sensitive_sn"] = "0"
        reload_shared_ui_config_namespace_runtime(namespace, load_snapshot=load_snapshot)
        self.assertEqual(len(observed), 2)
        self.assertFalse(observed[1]["sn"])
        unregister()
        self.assertNotIn("security", namespace["_CONFIG_SYNC_REFRESHERS"])

    def test_reload_syncs_each_push_option_and_parameter_without_repeated_notices(self):
        updates = {
            "enabled": "1",
            "sms_enabled": "0",
            "call_enabled": "0",
            "notify_type": '["wecom", "email"]',
            **{key: "synthetic-update" for key in THIRD_PUSH_SETTINGS_KEYS},
        }
        for key, value in updates.items():
            with self.subTest(key=key):
                namespace = self.make_namespace()
                snapshot = snapshot_config_runtime(namespace["config"])
                snapshot["third_push"][key] = value

                changed = reload_shared_ui_config_namespace_runtime(
                    namespace, load_snapshot=lambda _path: snapshot,
                )

                self.assertIn("三方推送", changed)
                expected = read_third_push_settings(namespace["config"])
                self.assertEqual(self.push_state(namespace), third_push_state(
                    expected.enabled, expected.sms_enabled, expected.call_enabled,
                    expected.channels, expected.settings,
                ))
                self.assertTrue(any(
                    item[0] == "ui" and "三方推送" in item[1][0]
                    for item in namespace["calls"] if isinstance(item, tuple)
                ))
                namespace["calls"].clear()
                repeated = reload_shared_ui_config_namespace_runtime(
                    namespace, load_snapshot=lambda _path: snapshot,
                )
                self.assertNotIn("三方推送", repeated)
                self.assertFalse(any(
                    item[0] == "ui" for item in namespace["calls"]
                    if isinstance(item, tuple)
                ))

    @staticmethod
    def push_state(namespace):
        return third_push_state(
            namespace["THIRD_PUSH_ENABLED"], namespace["THIRD_PUSH_SMS_ENABLED"],
            namespace["THIRD_PUSH_CALL_ENABLED"], namespace["THIRD_PUSH_TYPES"],
            namespace["THIRD_PUSH_SETTINGS"],
        )

    def test_shared_file_watch_updates_subsequent_pushes_in_both_instances(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "config.ini")
            initial = self.make_namespace()["config"]
            initial.set("third_push", "enabled", "1")
            initial.set("third_push", "notify_type", '["dingtalk"]')
            initial.set("third_push", "dingtalk_webhook", "https://old.example.invalid")
            with open(path, "w", encoding="utf-8") as file_obj:
                initial.write(file_obj)
            namespaces = [self.make_namespace(), self.make_namespace()]
            timers = []
            for namespace in namespaces:
                namespace["config"].read(path, encoding="utf-8")
                namespace["CONFIG_FILE"] = path
                remember_config_snapshot(namespace["config"])
                apply_third_push_settings_namespace_runtime(
                    namespace, read_third_push_settings(namespace["config"]),
                )
                namespace["safe_save_config"] = lambda ns=namespace: safe_save_config_runtime(
                    config=ns["config"], config_file=path, config_lock=ns["CONFIG_LOCK"],
                )
                namespace["THIRD_PUSH_Q"] = queue.Queue()
                timer = FakeRootTimer()
                namespace["root"] = SimpleNamespace(
                    after=timer.after, after_cancel=timer.after_cancel,
                    winfo_rgb=FakeRoot().winfo_rgb,
                )
                timers.append(timer)
                install_third_push_namespace_bindings(namespace)
                install_config_sync_namespace_bindings(namespace)
                namespace["start_config_file_watch"]()
            first, second = namespaces

            self.assertTrue(second["enqueue_third_push"]("queued before sync"))
            self.assertIsNotNone(first["save_third_push_setting"](
                notify_type=["wecom"], sms_enabled=False,
                settings={"wecom_webhook": "https://new.example.invalid"},
            ))
            timers[1].run(second["CONFIG_FILE_WATCH_STATE"].after_id)
            self.assertEqual(second["THIRD_PUSH_TYPES"], ["wecom"])
            self.assertFalse(second["enqueue_third_push"]("disabled SMS"))
            self.assertTrue(second["enqueue_third_push"]("call", event_type="call"))
            old_payload = second["THIRD_PUSH_Q"].get_nowait()
            new_payload = second["THIRD_PUSH_Q"].get_nowait()
            self.assertEqual(old_payload["channels"], ["dingtalk"])
            self.assertEqual(old_payload["settings"]["dingtalk_webhook"], "https://old.example.invalid")
            self.assertEqual(new_payload["channels"], ["wecom"])
            self.assertEqual(new_payload["settings"]["wecom_webhook"], "https://new.example.invalid")

            self.assertIsNotNone(second["save_third_push_setting"](sms_enabled=True, call_enabled=False))
            timers[0].run(first["CONFIG_FILE_WATCH_STATE"].after_id)
            self.assertFalse(first["enqueue_third_push"]("disabled call", event_type="call"))
            self.assertTrue(first["enqueue_third_push"]("SMS"))
            self.assertEqual(first["THIRD_PUSH_Q"].get_nowait()["channels"], ["wecom"])
            self.assertIsNotNone(first["save_third_push_setting"](enabled=False))
            timers[1].run(second["CONFIG_FILE_WATCH_STATE"].after_id)
            self.assertFalse(second["enqueue_third_push"]("disabled push"))
            self.assertEqual((second["PORT"], second["BAUD"], second["MODE"]), ("COM1", 115200, "Manual"))

    def test_push_window_sync_refreshes_clean_form_and_preserves_unsaved_draft(self):
        for dirty in (False, True):
            with self.subTest(dirty=dirty):
                namespace = self.make_namespace()
                form = make_form(self.push_state(namespace))
                form.state_provider = lambda: self.push_state(namespace)
                if dirty:
                    form.entry_vars["wecom_webhook"].set("https://draft.example.invalid")
                namespace["third_push_win"] = SimpleNamespace(
                    winfo_exists=lambda: True,
                    _sync_form_from_globals=form.sync_from_state_if_clean,
                )
                snapshot = snapshot_config_runtime(namespace["config"])
                snapshot["third_push"].update(
                    enabled="1", notify_type='["wecom"]',
                    wecom_webhook="https://external.example.invalid",
                )

                changed = reload_shared_ui_config_namespace_runtime(
                    namespace, load_snapshot=lambda _path: snapshot,
                )

                self.assertIn("三方推送", changed)
                self.assertEqual(namespace["THIRD_PUSH_SETTINGS"]["wecom_webhook"], "https://external.example.invalid")
                self.assertEqual(form.entry_vars["wecom_webhook"].get(),
                                 "https://draft.example.invalid" if dirty else "https://external.example.invalid")
                self.assertEqual(form.is_dirty(), dirty)

    def test_saving_one_switch_before_watch_reload_keeps_external_push_configuration(self):
        with tempfile.TemporaryDirectory() as directory:
            path = os.path.join(directory, "config.ini")
            initial = self.make_namespace()["config"]
            initial.set("third_push", "enabled", "1")
            initial.set("third_push", "notify_type", '["dingtalk"]')
            initial.set("third_push", "dingtalk_webhook", "https://old.example.invalid")
            with open(path, "w", encoding="utf-8") as file_obj:
                initial.write(file_obj)
            namespaces = [self.make_namespace(), self.make_namespace()]
            for namespace in namespaces:
                namespace["config"].read(path, encoding="utf-8")
                remember_config_snapshot(namespace["config"])
                apply_third_push_settings_namespace_runtime(
                    namespace, read_third_push_settings(namespace["config"]),
                )
                namespace["safe_save_config"] = lambda ns=namespace: safe_save_config_runtime(
                    config=ns["config"], config_file=path, config_lock=ns["CONFIG_LOCK"],
                )
                namespace["THIRD_PUSH_Q"] = queue.Queue()
                install_third_push_namespace_bindings(namespace)
            first, second = namespaces
            first["save_third_push_setting"](
                notify_type=["wecom"], settings={"wecom_webhook": "https://external.example.invalid"},
            )
            self.assertEqual(second["THIRD_PUSH_TYPES"], ["dingtalk"])
            self.assertIsNotNone(second["save_third_push_setting"](call_enabled=False))
            self.assertEqual(second["THIRD_PUSH_TYPES"], ["wecom"])
            self.assertEqual(second["THIRD_PUSH_SETTINGS"]["wecom_webhook"], "https://external.example.invalid")
            self.assertFalse(second["enqueue_third_push"]("call", event_type="call"))
            self.assertTrue(second["enqueue_third_push"]("SMS"))
            self.assertEqual(second["THIRD_PUSH_Q"].get_nowait()["channels"], ["wecom"])
            second["save_third_push_setting"](sms_enabled=False)
            self.assertEqual(second["THIRD_PUSH_TYPES"], ["wecom"])
            self.assertFalse(second["enqueue_third_push"]("disabled SMS"))

            third = self.make_namespace()
            third["config"].read(path, encoding="utf-8")
            remember_config_snapshot(third["config"])
            apply_third_push_settings_namespace_runtime(third, read_third_push_settings(initial))
            third["safe_save_config"] = lambda: safe_save_config_runtime(
                config=third["config"], config_file=path, config_lock=third["CONFIG_LOCK"],
            )
            install_third_push_namespace_bindings(third)
            self.assertIsNotNone(third["save_third_push_setting"](enabled=False))
            self.assertEqual(third["THIRD_PUSH_TYPES"], ["wecom"])
            self.assertFalse(third["THIRD_PUSH_ENABLED"])

    def test_push_window_refresh_failure_keeps_synced_runtime_and_redacts_error(self):
        namespace = self.make_namespace()
        namespace["third_push_win"] = SimpleNamespace(
            winfo_exists=lambda: True,
            _sync_form_from_globals=lambda: (_ for _ in ()).throw(RuntimeError("synthetic-sensitive-value")),
        )
        snapshot = snapshot_config_runtime(namespace["config"])
        snapshot["third_push"]["enabled"] = "1"

        changed = reload_shared_ui_config_namespace_runtime(
            namespace, load_snapshot=lambda _path: snapshot,
        )

        self.assertIn("三方推送", changed)
        self.assertTrue(namespace["THIRD_PUSH_ENABLED"])
        logs = [item[1] for item in namespace["calls"] if isinstance(item, tuple) and item[0] == "log"]
        self.assertIn("Refresh synced third-push settings window failed (RuntimeError)", logs)
        self.assertTrue(all("synthetic-sensitive-value" not in message for message in logs))

    def test_push_parse_failure_rolls_back_config_and_preserves_runtime(self):
        namespace = self.make_namespace()
        before = snapshot_config_runtime(namespace["config"])
        before_push = self.push_state(namespace)
        with patch("sms_ui.config_sync_namespace_runtime.read_third_push_settings",
                   side_effect=ValueError("synthetic-sensitive-value")):
            result = reload_shared_ui_config_namespace_runtime(
                namespace, load_snapshot=lambda _path: self.disk_snapshot(),
            )
        self.assertFalse(result)
        self.assertEqual(snapshot_config_runtime(namespace["config"]), before)
        self.assertEqual(self.push_state(namespace), before_push)
        self.assertTrue(namespace["POPUP_ENABLED"])

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
