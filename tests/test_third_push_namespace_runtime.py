import configparser
import unittest
from unittest.mock import patch

from sms_core.third_push_config import ThirdPushSettings
import sms_ui.third_push_namespace_runtime as runtime


class ThirdPushNamespaceRuntimeTests(unittest.TestCase):
    def make_namespace(self):
        calls = []
        return {
            "calls": calls,
            "config": configparser.ConfigParser(),
            "CONFIG_FILE": "missing.ini",
            "safe_save_config": lambda: calls.append(("save_config",)),
            "THIRD_PUSH_ENABLED": True,
            "THIRD_PUSH_SMS_ENABLED": False,
            "THIRD_PUSH_CALL_ENABLED": True,
            "THIRD_PUSH_TYPES": ["dingtalk"],
            "THIRD_PUSH_SETTINGS": {"dingtalk_webhook": "url"},
            "LOCAL_NUMBER": "+8613812345678",
            "third_push_stop": "stop",
            "THIRD_PUSH_Q": "queue",
            "LOG_PREFIX": "COM5",
            "APP_VERSION": "3.6.6",
            "system_ui": lambda *args: calls.append(("system_ui", args)),
            "show_third_push_test_result": lambda *args: calls.append(("show_result", args)),
            "root": "root",
            "third_push_win": "old_win",
            "messagebox": "messagebox",
            "ui_post": "ui_post",
            "refresh_third_push_settings_from_config": lambda: calls.append(("refresh",)),
            "save_third_push_setting": lambda **kwargs: calls.append(("save_setting", kwargs)),
            "enqueue_third_push": lambda *args, **kwargs: calls.append(("enqueue_push", args, kwargs)),
            "sync_and_focus_existing_window": lambda win, attr, **kwargs: ("sync", win, attr, kwargs["log_error"]("sync log")),
            "center_window": "center",
            "log_file_only": lambda message: ("log", message),
        }

    def test_apply_updates_namespace_state(self):
        namespace = self.make_namespace()

        runtime.apply_third_push_settings_namespace_runtime(
            namespace,
            ThirdPushSettings(False, True, False, ["wecom"], {"wecom_webhook": "hook"}),
        )

        self.assertFalse(namespace["THIRD_PUSH_ENABLED"])
        self.assertTrue(namespace["THIRD_PUSH_SMS_ENABLED"])
        self.assertFalse(namespace["THIRD_PUSH_CALL_ENABLED"])
        self.assertEqual(namespace["THIRD_PUSH_TYPES"], ["wecom"])
        self.assertEqual(namespace["THIRD_PUSH_SETTINGS"], {"wecom_webhook": "hook"})

    def test_ensure_saves_when_defaults_changed(self):
        namespace = self.make_namespace()

        with patch.object(runtime, "ensure_third_push_config_values", return_value=True):
            changed = runtime.ensure_third_push_config_namespace_runtime(namespace, save=True)

        self.assertTrue(changed)
        self.assertIn(("save_config",), namespace["calls"])

    def test_refresh_reads_config_and_applies_settings(self):
        namespace = self.make_namespace()
        settings = ThirdPushSettings(False, True, True, ["bark"], {"bark_url": "url"})

        with patch.object(runtime, "ensure_third_push_config_values", return_value=False), \
                patch.object(runtime, "read_third_push_settings", return_value=settings):
            runtime.refresh_third_push_settings_namespace_runtime(namespace)

        self.assertFalse(namespace["THIRD_PUSH_ENABLED"])
        self.assertEqual(namespace["THIRD_PUSH_TYPES"], ["bark"])
        self.assertEqual(namespace["THIRD_PUSH_SETTINGS"], {"bark_url": "url"})

    def test_save_forwards_current_state_and_applies_result(self):
        namespace = self.make_namespace()

        def save_runtime(**kwargs):
            current = kwargs["current_settings"]()
            self.assertEqual(current, ThirdPushSettings(
                True,
                False,
                True,
                ["dingtalk"],
                {"dingtalk_webhook": "url"},
            ))
            next_settings = ThirdPushSettings(False, True, False, ["wecom"], {"wecom_webhook": "hook"})
            kwargs["apply_settings"](next_settings)
            self.assertIs(kwargs["config"], namespace["config"])
            kwargs["save_config"]()
            self.assertFalse(kwargs["enabled"])
            return next_settings

        with patch.object(runtime, "save_third_push_setting_runtime", side_effect=save_runtime):
            result = runtime.save_third_push_setting_namespace_runtime(namespace, enabled=False)

        self.assertEqual(result.channels, ["wecom"])
        self.assertFalse(namespace["THIRD_PUSH_ENABLED"])
        self.assertEqual(namespace["THIRD_PUSH_SETTINGS"], {"wecom_webhook": "hook"})
        self.assertIn(("save_config",), namespace["calls"])

    def test_worker_forwards_runtime_context(self):
        namespace = self.make_namespace()

        with patch.object(runtime, "third_push_worker_app_runtime", return_value="worked") as worker:
            result = runtime.third_push_worker_namespace_runtime(namespace)

        self.assertEqual(result, "worked")
        kwargs = worker.call_args.kwargs
        self.assertEqual(kwargs["stop_event"], "stop")
        self.assertEqual(kwargs["push_queue"], "queue")
        self.assertEqual(kwargs["get_log_prefix"](), "COM5")
        self.assertEqual(kwargs["app_version"], "3.6.6")
        kwargs["system_ui"]("msg", "normal")
        kwargs["show_result"](["ok"], [])
        self.assertIn(("system_ui", ("msg", "normal")), namespace["calls"])
        self.assertIn(("show_result", (["ok"], [])), namespace["calls"])

    def test_result_and_enqueue_forward_current_values(self):
        namespace = self.make_namespace()

        with patch.object(runtime, "show_third_push_test_result_app_runtime", return_value="shown") as show_runtime:
            self.assertEqual(
                runtime.show_third_push_test_result_namespace_runtime(namespace, ["ok"], ["fail"]),
                "shown",
            )

        show_kwargs = show_runtime.call_args.kwargs
        self.assertEqual(show_kwargs["root"], "root")
        self.assertEqual(show_kwargs["get_current_window"](), "old_win")
        self.assertEqual(show_kwargs["ok_channels"], ["ok"])

        with patch.object(runtime, "enqueue_third_push_app_runtime", return_value="queued") as enqueue_runtime:
            result = runtime.enqueue_third_push_namespace_runtime(
                namespace,
                "raw",
                show_success=True,
                show_result=True,
                channels=["bark"],
                settings={"bark_url": "url"},
                template="{msg}",
                event_type="call",
            )

        self.assertEqual(result, "queued")
        args, kwargs = enqueue_runtime.call_args
        self.assertEqual(args, ("raw",))
        self.assertEqual(kwargs["push_queue"], "queue")
        self.assertTrue(kwargs["enabled"]())
        self.assertFalse(kwargs["sms_enabled"]())
        self.assertEqual(kwargs["configured_channels"](), ["dingtalk"])
        self.assertEqual(kwargs["channels"], ["bark"])
        self.assertEqual(kwargs["event_type"], "call")

    def test_enqueue_uses_cached_local_number_when_variable_is_blank(self):
        namespace = self.make_namespace()

        with patch.object(runtime, "enqueue_third_push_app_runtime", return_value="queued") as enqueue_runtime:
            result = runtime.enqueue_third_push_namespace_runtime(
                namespace,
                "raw",
                variables={"sender": "106598731", "local_number": "", "self_number": ""},
            )

        self.assertEqual(result, "queued")
        variables = enqueue_runtime.call_args.kwargs["variables"]
        self.assertEqual(variables["sender"], "106598731")
        self.assertEqual(variables["local_number"], "+8613812345678")
        self.assertEqual(variables["self_number"], "+8613812345678")

    def test_open_window_forwards_callbacks_and_stores_window(self):
        namespace = self.make_namespace()

        def open_runtime(**kwargs):
            self.assertEqual(kwargs["parent"], "root")
            self.assertEqual(kwargs["current_window"](), "old_win")
            self.assertTrue(kwargs["enabled"]())
            self.assertEqual(kwargs["channels"](), ["dingtalk"])
            kwargs["refresh_settings"]()
            kwargs["save_setting"](enabled=True)
            kwargs["enqueue_push"]("msg", show_result=True)
            self.assertEqual(
                kwargs["sync_existing_window"]("win", "_sync"),
                ("sync", "win", "_sync", ("log", "sync log")),
            )
            kwargs["set_window"]("new_win")
            self.assertEqual(kwargs["center_window"], "center")
            return "opened"

        with patch.object(runtime, "open_third_push_values_app_runtime", side_effect=open_runtime):
            result = runtime.open_third_push_window_namespace_runtime(namespace)

        self.assertEqual(result, "opened")
        self.assertEqual(namespace["third_push_win"], "new_win")
        self.assertIn(("refresh",), namespace["calls"])
        self.assertIn(("save_setting", {"enabled": True}), namespace["calls"])
        self.assertIn(("enqueue_push", ("msg",), {"show_result": True}), namespace["calls"])


if __name__ == "__main__":
    unittest.main()
