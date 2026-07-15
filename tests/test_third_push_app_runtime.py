import unittest
import configparser

from sms_core.third_push_config import ThirdPushSettings
from sms_ui.third_push_app_runtime import (
    enqueue_third_push_app_runtime,
    format_third_push_message_runtime,
    open_third_push_app_runtime,
    open_third_push_values_app_runtime,
    save_third_push_setting_runtime,
    send_third_push_channel_runtime,
    show_third_push_test_result_app_runtime,
    third_push_settings_from_values,
    third_push_worker_app_runtime,
)


class ThirdPushAppRuntimeTests(unittest.TestCase):
    def test_third_push_settings_from_values_builds_settings(self):
        settings = third_push_settings_from_values(
            True,
            False,
            True,
            ["dingtalk"],
            {"dingtalk_webhook": "url"},
        )

        self.assertEqual(settings, ThirdPushSettings(
            True,
            False,
            True,
            ["dingtalk"],
            {"dingtalk_webhook": "url"},
        ))

    def test_save_third_push_setting_runtime_updates_writes_and_saves(self):
        config = configparser.ConfigParser()
        calls = []

        result = save_third_push_setting_runtime(
            current_settings=lambda: ThirdPushSettings(True, False, True, ["dingtalk"], {}),
            apply_settings=lambda settings: calls.append(("apply", settings)),
            config=config,
            save_config=lambda: calls.append(("save",)),
            enabled=False,
            sms_enabled=True,
            notify_type=["wecom", "bad", "bark"],
            settings={"wecom_webhook": "https://example.test/wecom"},
        )

        self.assertFalse(result.enabled)
        self.assertTrue(result.sms_enabled)
        self.assertEqual(result.channels, ["wecom", "bark"])
        self.assertEqual(calls[0], ("save",))
        self.assertEqual(calls[1], ("apply", result))
        self.assertEqual(config.get("third_push", "enabled"), "0")
        self.assertEqual(config.get("third_push", "sms_enabled"), "1")

    def test_save_third_push_setting_runtime_returns_none_on_save_error(self):
        calls = []
        config = configparser.ConfigParser()
        config["third_push"] = {
            "enabled": "1",
            "notify_type": '["dingtalk"]',
            "dingtalk_webhook": "https://old.example",
            "extra_key": "keep-me",
        }
        before = dict(config.items("third_push", raw=True))

        result = save_third_push_setting_runtime(
            current_settings=lambda: ThirdPushSettings(True, True, True, [], {}),
            apply_settings=lambda settings: calls.append(("apply", settings)),
            config=config,
            save_config=lambda: (_ for _ in ()).throw(RuntimeError("disk")),
            enabled=False,
        )

        self.assertIsNone(result)
        self.assertEqual(calls, [])
        self.assertEqual(dict(config.items("third_push", raw=True)), before)

    def test_save_third_push_setting_runtime_returns_none_on_false_save_result(self):
        calls = []
        config = configparser.ConfigParser()
        config["third_push"] = {"enabled": "1", "notify_type": '[]', "extra_key": "keep"}
        before = dict(config.items("third_push", raw=True))

        result = save_third_push_setting_runtime(
            current_settings=lambda: ThirdPushSettings(True, True, True, [], {}),
            apply_settings=lambda settings: calls.append(("apply", settings)),
            config=config,
            save_config=lambda: False,
            enabled=False,
        )

        self.assertIsNone(result)
        self.assertEqual(calls, [])
        self.assertEqual(dict(config.items("third_push", raw=True)), before)

    def test_open_third_push_app_runtime_builds_state_and_forwards_callbacks(self):
        calls = []

        def open_window_runtime(
            parent,
            current_window,
            state_provider,
            refresh_settings,
            save_setting,
            enqueue_push,
            system_ui,
            sync_existing_window,
            set_window,
            center_window,
        ):
            calls.append(("parent", parent))
            calls.append(("window", current_window))
            calls.append(("state", state_provider()))
            refresh_settings()
            save_setting(enabled=True)
            enqueue_push("msg")
            system_ui("saved", "normal")
            calls.append(("sync", sync_existing_window("win", "_sync")))
            set_window("new_win")
            calls.append(("center", center_window))
            return "new_win"

        result = open_third_push_app_runtime(
            parent="root",
            current_window="old_win",
            enabled=True,
            sms_enabled=False,
            call_enabled=True,
            channels=["dingtalk"],
            settings={"token": "t"},
            refresh_settings=lambda: calls.append(("refresh",)),
            save_setting=lambda **kwargs: calls.append(("save", kwargs)),
            enqueue_push=lambda *args: calls.append(("enqueue", args)),
            system_ui=lambda *args: calls.append(("system", args)),
            sync_existing_window=lambda win, attr: (win, attr),
            set_window=lambda win: calls.append(("set_window", win)),
            center_window="center",
            open_window_runtime=open_window_runtime,
        )

        self.assertEqual(result, "new_win")
        self.assertEqual(calls[0], ("parent", "root"))
        self.assertEqual(calls[1], ("window", "old_win"))
        self.assertEqual(calls[2], ("state", {
            "enabled": True,
            "sms_enabled": False,
            "call_enabled": True,
            "channels": ["dingtalk"],
            "settings": {"token": "t"},
        }))
        self.assertIn(("refresh",), calls)
        self.assertIn(("save", {"enabled": True}), calls)
        self.assertIn(("enqueue", ("msg",)), calls)
        self.assertIn(("system", ("saved", "normal")), calls)
        self.assertIn(("sync", ("win", "_sync")), calls)
        self.assertIn(("set_window", "new_win"), calls)
        self.assertIn(("center", "center"), calls)

    def test_format_and_send_third_push_runtime_add_app_context(self):
        formatted = format_third_push_message_runtime(
            "raw",
            "port={port}; {msg}",
            get_log_prefix=lambda: "COM5",
            format_message=lambda raw, template, port="": template.replace("{msg}", raw).replace("{port}", port),
        )
        self.assertEqual(formatted, "port=COM5; raw")

        result = send_third_push_channel_runtime(
            "dingtalk",
            "message",
            {"token": "t"},
            get_log_prefix=lambda: "COM7",
            app_version="3.6.6",
            send_channel=lambda channel, message, settings, user_agent, port: (
                channel,
                message,
                settings,
                user_agent,
                port,
            ),
        )

        self.assertEqual(result, (
            "dingtalk",
            "message",
            {"token": "t"},
            "Air724UG-SMS/3.6.6",
            "COM7",
        ))

    def test_worker_and_result_runtimes_forward_context(self):
        calls = []

        worker_result = third_push_worker_app_runtime(
            stop_event="stop",
            push_queue="queue",
            get_log_prefix=lambda: "COM5",
            app_version="3.6.6",
            system_ui=lambda *args: calls.append(("ui", args)),
            show_result=lambda *args: calls.append(("result", args)),
            worker_runtime=lambda **kwargs: calls.append(("worker", kwargs)) or "worked",
        )

        self.assertEqual(worker_result, "worked")
        worker_kwargs = calls[0][1]
        self.assertEqual(worker_kwargs["stop_event"], "stop")
        self.assertEqual(worker_kwargs["push_queue"], "queue")
        self.assertTrue(callable(worker_kwargs["send_channel_func"]))
        self.assertTrue(callable(worker_kwargs["format_message_func"]))

        show_result = show_third_push_test_result_app_runtime(
            root="root",
            get_current_window=lambda: "win",
            messagebox="messagebox",
            ui_post="post",
            ok_channels=["ok"],
            fail_infos=[],
            show_runtime=lambda **kwargs: calls.append(("show", kwargs)) or "shown",
        )

        self.assertEqual(show_result, "shown")
        self.assertEqual(calls[-1][1]["current_window"], "win")
        self.assertEqual(calls[-1][1]["ok_channels"], ["ok"])

    def test_enqueue_and_open_values_runtimes_read_current_values(self):
        calls = []

        enqueue_result = enqueue_third_push_app_runtime(
            "raw",
            push_queue="queue",
            enabled=lambda: True,
            sms_enabled=lambda: False,
            call_enabled=lambda: True,
            configured_channels=lambda: ["dingtalk"],
            current_settings=lambda: {"token": "t"},
            system_ui=lambda *args: calls.append(("ui", args)),
            channels=None,
            show_result=True,
            enqueue_runtime=lambda raw, **kwargs: calls.append(("enqueue", raw, kwargs)) or "queued",
        )

        self.assertEqual(enqueue_result, "queued")
        enqueue_kwargs = calls[0][2]
        self.assertTrue(enqueue_kwargs["enabled"])
        self.assertFalse(enqueue_kwargs["sms_enabled"])
        self.assertEqual(enqueue_kwargs["configured_channels"], ["dingtalk"])
        self.assertTrue(enqueue_kwargs["show_result"])

        open_result = open_third_push_values_app_runtime(
            parent="root",
            current_window=lambda: "old",
            enabled=lambda: True,
            sms_enabled=lambda: True,
            call_enabled=lambda: False,
            channels=lambda: ["wecom"],
            settings=lambda: {"wecom_webhook": "url"},
            refresh_settings=lambda: calls.append(("refresh",)),
            save_setting=lambda **kwargs: calls.append(("save", kwargs)),
            enqueue_push=lambda *args: calls.append(("push", args)),
            system_ui=lambda *args: calls.append(("system", args)),
            sync_existing_window=lambda win, attr: (win, attr),
            set_window=lambda win: calls.append(("window", win)),
            center_window="center",
            open_app_runtime=lambda **kwargs: calls.append(("open", kwargs)) or "opened",
        )

        self.assertEqual(open_result, "opened")
        open_kwargs = calls[-1][1]
        self.assertEqual(open_kwargs["current_window"], "old")
        self.assertTrue(open_kwargs["enabled"])
        self.assertEqual(open_kwargs["channels"], ["wecom"])
        self.assertEqual(open_kwargs["settings"], {"wecom_webhook": "url"})


if __name__ == "__main__":
    unittest.main()
