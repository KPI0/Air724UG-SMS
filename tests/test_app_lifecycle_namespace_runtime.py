import unittest

from sms_ui.app_lifecycle_namespace_runtime import (
    cleanup_and_exit_namespace_runtime,
    restart_software_namespace_runtime,
    set_autostart_namespace_runtime,
    toggle_multi_instance_namespace_runtime,
    toggle_popup_namespace_runtime,
    toggle_voice_broadcast_namespace_runtime,
)


class FakeRoot:
    def destroy(self):
        return "destroyed"


class FakeVar:
    def __init__(self, value):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeOs:
    @staticmethod
    def _exit(code):
        return ("exit", code)


class FakeThreadRegistry:
    def __init__(self, threads):
        self.threads = tuple(threads)

    def snapshot(self):
        return self.threads


class AppLifecycleNamespaceRuntimeTests(unittest.TestCase):
    def base_namespace(self):
        return {
            "AUTOSTART_FLAG": "--autostart",
            "RESTART_HELPER_FLAG": "--restart-helper",
            "root": FakeRoot(),
            "messagebox": "messagebox",
            "is_exiting": False,
            "serial_running": True,
            "TK_SHUTDOWN": "tk",
            "file_log_stop": "file",
            "third_push_stop": "third",
            "serial_stop_event": "serial",
            "serial_wakeup_event": "wakeup",
            "TTS_STOP": "tts",
            "THIRD_PUSH_Q": "third-queue",
            "third_push_thread": "third-thread",
            "serial_thread": "serial-thread",
            "TTS_THREAD": "tts-thread",
            "cloud_ws_thread": "cloud-thread",
            "tray_thread": "tray-thread",
            "FILE_LOG_Q": "queue",
            "VOICE_ENABLED": True,
            "ALLOW_MULTI_INSTANCE": False,
            "APP_DIR": "app-dir",
            "POPUP_ENABLED": True,
            "config": "config",
            "multi_instance_var": FakeVar(True),
            "popup_var": FakeVar(False),
            "os": FakeOs,
            "create_startup_shortcut": lambda flag: ("create", flag),
            "remove_startup_shortcut": lambda: "remove",
            "system_ui": lambda *args: ("ui", args),
            "ui_messagebox": lambda *args: ("message", args),
            "safe_save_config": lambda: "save",
            "update_voice_menu_label": lambda: "label",
            "safe_set_events": lambda *events: ("events", events),
            "stop_cloud_control": lambda **kwargs: ("cloud", kwargs),
            "safe_close_serial": lambda: "close",
            "stop_tray_icon": lambda **kwargs: ("tray", kwargs),
            "flush_log_queue": lambda queue: ("flush", queue),
            "log_file_only": lambda message: ("log", message),
            "app_mutex": "mutex",
            "release_mutex_handle": lambda mutex: ("release", mutex),
            "UPDATE_THREAD_REGISTRY": FakeThreadRegistry(
                ("update-1", "update-2")
            ),
            "SERIAL_COMMAND_THREAD_REGISTRY": FakeThreadRegistry(
                ("serial-command-1", "serial-command-2")
            ),
            "SMS_SEND_THREAD_REGISTRY": FakeThreadRegistry(
                ("sms-send-1", "sms-send-2")
            ),
        }

    def test_set_autostart_namespace_runtime_forwards_shortcut_dependencies(self):
        namespace = self.base_namespace()
        calls = []

        result = set_autostart_namespace_runtime(
            namespace,
            True,
            set_autostart_app_runtime=lambda enable, **kwargs: calls.append((enable, kwargs)) or "ok",
        )

        self.assertEqual(result, "ok")
        self.assertTrue(calls[0][0])
        forwarded = calls[0][1]
        self.assertEqual(forwarded["autostart_flag"], "--autostart")
        self.assertEqual(forwarded["show_error"]("title", "msg"), ("message", ("error", "title", "msg")))

    def test_cleanup_and_exit_namespace_runtime_forwards_and_mutates_state(self):
        namespace = self.base_namespace()
        calls = []

        result = cleanup_and_exit_namespace_runtime(
            namespace,
            cleanup_app_runtime=lambda **kwargs: calls.append(kwargs) or "exited",
        )

        self.assertEqual(result, "exited")
        forwarded = calls[0]
        self.assertEqual(forwarded["shutdown_events"], ("tk",))
        self.assertEqual(forwarded["worker_stop_events"], ("serial", "wakeup"))
        self.assertEqual(forwarded["tts_stop_event"], "tts")
        self.assertEqual(forwarded["deferred_worker_stop_events"], ("third",))
        self.assertEqual(forwarded["deferred_worker_threads"](), ("third-thread",))
        self.assertEqual(forwarded["deferred_worker_queues"], ("third-queue",))
        self.assertEqual(forwarded["file_log_stop_event"], "file")
        self.assertEqual(
            forwarded["worker_threads"]()[-4:],
            ("serial-command-1", "serial-command-2", "sms-send-1", "sms-send-2"),
        )
        self.assertIn("update-1", forwarded["worker_threads"]())
        self.assertIn("update-2", forwarded["worker_threads"]())
        self.assertIn("tray-thread", forwarded["worker_threads"]())
        forwarded["set_exiting"](True)
        forwarded["set_serial_running"](False)
        self.assertTrue(namespace["is_exiting"])
        self.assertFalse(namespace["serial_running"])
        self.assertEqual(forwarded["destroy_root"](), "destroyed")

    def test_toggle_namespace_runtimes_update_state(self):
        namespace = self.base_namespace()
        calls = []

        voice = toggle_voice_broadcast_namespace_runtime(
            namespace,
            toggle_runtime=lambda current, config, save, set_enabled, update_label, system_ui, **kwargs: (
                set_enabled(False),
                calls.append(("voice", current, config, save(), update_label(), system_ui("v"), kwargs["log_error"]("voice log"))),
                "voice-result",
            )[-1],
        )
        multi = toggle_multi_instance_namespace_runtime(
            namespace,
            toggle_runtime=lambda enabled, config, save, set_multi, system_ui, **kwargs: (
                set_multi(enabled),
                calls.append(("multi", enabled, config, save(), system_ui("m"), kwargs["log_error"]("multi log"))),
                "multi-result",
            )[-1],
        )
        popup = toggle_popup_namespace_runtime(
            namespace,
            toggle_runtime=lambda enabled, config, save, set_popup, system_ui, **kwargs: (
                set_popup(enabled),
                calls.append(("popup", enabled, config, save(), system_ui("p"), kwargs["log_error"]("popup log"))),
                "popup-result",
            )[-1],
        )

        self.assertEqual((voice, multi, popup), ("voice-result", "multi-result", "popup-result"))
        self.assertFalse(namespace["VOICE_ENABLED"])
        self.assertTrue(namespace["ALLOW_MULTI_INSTANCE"])
        self.assertFalse(namespace["POPUP_ENABLED"])
        self.assertEqual(calls[0][0], "voice")
        self.assertEqual(calls[1][0], "multi")
        self.assertEqual(calls[2][0], "popup")
        self.assertEqual(calls[0][-1], ("log", "voice log"))
        self.assertEqual(calls[1][-1], ("log", "multi log"))
        self.assertEqual(calls[2][-1], ("log", "popup log"))

    def test_toggle_namespace_runtimes_restore_menu_vars_when_save_fails(self):
        namespace = self.base_namespace()

        multi_result = toggle_multi_instance_namespace_runtime(
            namespace,
            toggle_runtime=lambda *args, **kwargs: None,
        )
        popup_result = toggle_popup_namespace_runtime(
            namespace,
            toggle_runtime=lambda *args, **kwargs: None,
        )

        self.assertIsNone(multi_result)
        self.assertIsNone(popup_result)
        self.assertFalse(namespace["multi_instance_var"].get())
        self.assertTrue(namespace["popup_var"].get())

    def test_disabling_multi_instance_acquires_mutex_before_saving(self):
        namespace = self.base_namespace()
        namespace["ALLOW_MULTI_INSTANCE"] = True
        namespace["multi_instance_var"] = FakeVar(False)
        namespace["app_mutex"] = None
        calls = []

        result = toggle_multi_instance_namespace_runtime(
            namespace,
            acquire_mutex=lambda name: calls.append(("acquire", name)) or ("new-mutex", 0),
            toggle_runtime=lambda enabled, config, save, set_multi, system_ui, **kwargs: (
                calls.append(("toggle", enabled)),
                set_multi(enabled),
                False,
            )[-1],
        )

        self.assertFalse(result)
        self.assertEqual(calls[0][0], "acquire")
        self.assertEqual(calls[1], ("toggle", False))
        self.assertEqual(namespace["app_mutex"], "new-mutex")
        self.assertFalse(namespace["ALLOW_MULTI_INSTANCE"])

    def test_disabling_multi_instance_releases_new_mutex_when_save_fails(self):
        namespace = self.base_namespace()
        namespace["ALLOW_MULTI_INSTANCE"] = True
        namespace["multi_instance_var"] = FakeVar(False)
        namespace["app_mutex"] = None
        released = []
        namespace["release_mutex_handle"] = released.append

        result = toggle_multi_instance_namespace_runtime(
            namespace,
            acquire_mutex=lambda _name: ("new-mutex", 0),
            toggle_runtime=lambda *args, **kwargs: None,
        )

        self.assertIsNone(result)
        self.assertEqual(released, ["new-mutex"])
        self.assertIsNone(namespace["app_mutex"])
        self.assertTrue(namespace["ALLOW_MULTI_INSTANCE"])
        self.assertTrue(namespace["multi_instance_var"].get())

    def test_disabling_multi_instance_releases_new_mutex_when_toggle_raises(self):
        namespace = self.base_namespace()
        namespace["ALLOW_MULTI_INSTANCE"] = True
        namespace["multi_instance_var"] = FakeVar(False)
        namespace["app_mutex"] = None
        released = []
        namespace["release_mutex_handle"] = released.append

        with self.assertRaisesRegex(RuntimeError, "save crashed"):
            toggle_multi_instance_namespace_runtime(
                namespace,
                acquire_mutex=lambda _name: ("new-mutex", 0),
                toggle_runtime=lambda *args, **kwargs: (_ for _ in ()).throw(RuntimeError("save crashed")),
            )

        self.assertEqual(released, ["new-mutex"])
        self.assertIsNone(namespace["app_mutex"])

    def test_disabling_multi_instance_keeps_setting_when_mutex_cannot_be_created(self):
        namespace = self.base_namespace()
        namespace["ALLOW_MULTI_INSTANCE"] = True
        namespace["multi_instance_var"] = FakeVar(False)
        namespace["app_mutex"] = None
        calls = []
        namespace["system_ui"] = lambda *args: calls.append(("ui", args))
        namespace["ui_messagebox"] = lambda *args: calls.append(("message", args))

        result = toggle_multi_instance_namespace_runtime(
            namespace,
            acquire_mutex=lambda _name: (None, 8),
            toggle_runtime=lambda *args, **kwargs: calls.append(("toggle",)) or False,
        )

        self.assertIsNone(result)
        self.assertTrue(namespace["ALLOW_MULTI_INSTANCE"])
        self.assertTrue(namespace["multi_instance_var"].get())
        self.assertFalse(any(call[0] == "toggle" for call in calls))
        self.assertTrue(any(call[0] == "message" for call in calls))

    def test_restart_software_namespace_runtime_forwards_shutdown_dependencies(self):
        namespace = self.base_namespace()
        calls = []

        result = restart_software_namespace_runtime(
            namespace,
            restart_app_runtime=lambda **kwargs: calls.append(kwargs) or "restarted",
        )

        self.assertEqual(result, "restarted")
        forwarded = calls[0]
        self.assertEqual(forwarded["autostart_flag"], "--autostart")
        self.assertEqual(forwarded["restart_helper_flag"], "--restart-helper")
        self.assertEqual(forwarded["stop_events"], ("tk", "serial", "wakeup", "tts"))
        self.assertEqual(forwarded["deferred_stop_events"], ("third",))
        self.assertEqual(forwarded["deferred_worker_threads"](), ("third-thread",))
        self.assertEqual(forwarded["deferred_worker_queues"], ("third-queue",))
        self.assertEqual(forwarded["file_log_stop_event"], "file")
        self.assertEqual(
            forwarded["worker_threads"]()[-4:],
            ("serial-command-1", "serial-command-2", "sms-send-1", "sms-send-2"),
        )
        self.assertIn("update-1", forwarded["worker_threads"]())
        self.assertIn("update-2", forwarded["worker_threads"]())
        self.assertIn("tray-thread", forwarded["worker_threads"]())
        self.assertEqual(forwarded["app_mutex"], "mutex")
        self.assertEqual(forwarded["file_log_queue"], "queue")
        forwarded["set_exiting"](True)
        forwarded["set_serial_running"](False)
        self.assertTrue(namespace["is_exiting"])
        self.assertFalse(namespace["serial_running"])
        self.assertEqual(forwarded["exit_process"](0), ("exit", 0))

    def test_shutdown_worker_snapshot_is_deferred_until_cleanup_requests_it(self):
        namespace = self.base_namespace()
        calls = []

        cleanup_and_exit_namespace_runtime(
            namespace,
            cleanup_app_runtime=lambda **kwargs: calls.append(kwargs) or "exited",
        )

        namespace["UPDATE_THREAD_REGISTRY"].threads += ("late-update",)
        namespace["SERIAL_COMMAND_THREAD_REGISTRY"].threads += ("late-command",)
        namespace["SMS_SEND_THREAD_REGISTRY"].threads += ("late-sms",)
        snapshot = calls[0]["worker_threads"]()
        self.assertIn("late-update", snapshot)
        self.assertIn("late-command", snapshot)
        self.assertIn("late-sms", snapshot)

    def test_shutdown_worker_snapshot_propagates_registry_failure(self):
        namespace = self.base_namespace()
        calls = []

        class FailingRegistry:
            def snapshot(self):
                raise RuntimeError("snapshot failed")

        namespace["SERIAL_COMMAND_THREAD_REGISTRY"] = FailingRegistry()
        cleanup_and_exit_namespace_runtime(
            namespace,
            cleanup_app_runtime=lambda **kwargs: calls.append(kwargs) or "wired",
        )

        with self.assertRaisesRegex(RuntimeError, "SERIAL_COMMAND_THREAD_REGISTRY"):
            calls[0]["worker_threads"]()


if __name__ == "__main__":
    unittest.main()
