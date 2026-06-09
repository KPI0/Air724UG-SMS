import unittest

from sms_ui.maintenance_namespace_runtime import (
    check_update_and_prompt_namespace_runtime,
    open_log_cleanup_dialog_namespace_runtime,
    open_log_dir_namespace_runtime,
    open_update_proxy_dialog_namespace_runtime,
    run_auto_log_cleanup_tick_namespace_runtime,
    schedule_auto_log_cleanup_namespace_runtime,
)


class FakePath:
    @staticmethod
    def abspath(path):
        return f"ABS/{path}"

    @staticmethod
    def exists(path):
        return True


class FakeOs:
    path = FakePath

    @staticmethod
    def startfile(path):
        return ("open", path)


class FakeRoot:
    def after(self, *args):
        return ("after", args)

    def after_cancel(self, *args):
        return ("cancel", args)


class FakeMessageBox:
    @staticmethod
    def showinfo(*args):
        return ("info", args)

    @staticmethod
    def showwarning(*args):
        return ("warning", args)

    @staticmethod
    def showerror(*args):
        return ("error", args)

    @staticmethod
    def askyesno(*args):
        return False


class FakeWebbrowser:
    @staticmethod
    def open(url):
        return ("open_url", url)


class MaintenanceNamespaceRuntimeTests(unittest.TestCase):
    def base_namespace(self):
        return {
            "LOG_DIR": "logs",
            "LOG_RETENTION_DAYS": 14,
            "AUTO_CLEANUP_INTERVAL_HOURS": 6,
            "AUTO_LOG_CLEANUP": True,
            "AUTO_LOG_CLEANUP_STATE": "cleanup_state",
            "GITHUB_OWNER": "owner",
            "GITHUB_REPO": "repo",
            "APP_VERSION": "1.0.0",
            "root": FakeRoot(),
            "config": "config",
            "os": FakeOs,
            "webbrowser": FakeWebbrowser,
            "messagebox": FakeMessageBox,
            "ui_messagebox": lambda *args: ("message", args),
            "safe_save_config": lambda: "saved",
            "system_ui": lambda *args: ("system", args),
            "schedule_auto_log_cleanup": lambda **kwargs: ("schedule", kwargs),
            "center_window": lambda win: ("center", win),
            "ui_post": lambda callback: callback(),
            "cleanup_old_logs": lambda days: ("cleanup", days),
            "tk_alive": lambda: True,
            "_auto_log_cleanup_tick": lambda: "tick",
        }

    def test_open_log_dir_namespace_runtime_forwards_os_callbacks(self):
        namespace = self.base_namespace()
        calls = []

        result = open_log_dir_namespace_runtime(
            namespace,
            open_dir_runtime=lambda log_dir, **kwargs: calls.append((log_dir, kwargs)) or "opened",
        )

        self.assertEqual(result, "opened")
        self.assertEqual(calls[0][0], "logs")
        self.assertEqual(calls[0][1]["path_abspath"]("logs"), "ABS/logs")
        self.assertTrue(calls[0][1]["path_exists"]("ABS/logs"))
        self.assertEqual(calls[0][1]["open_path"]("ABS/logs"), ("open", "ABS/logs"))

    def test_open_log_cleanup_dialog_namespace_runtime_forwards_state(self):
        namespace = self.base_namespace()
        calls = []

        result = open_log_cleanup_dialog_namespace_runtime(
            namespace,
            open_dialog_runtime=lambda parent, **kwargs: calls.append((parent, kwargs)) or "dialog",
        )

        self.assertEqual(result, "dialog")
        parent, forwarded = calls[0]
        self.assertIs(parent, namespace["root"])
        self.assertEqual(forwarded["get_retention_days"](), 14)
        self.assertEqual(forwarded["get_interval_hours"](), 6)
        forwarded["set_cleanup_state"](30, False)
        self.assertEqual(namespace["LOG_RETENTION_DAYS"], 30)
        self.assertFalse(namespace["AUTO_LOG_CLEANUP"])

    def test_open_update_proxy_dialog_namespace_runtime_forwards_values(self):
        namespace = self.base_namespace()
        calls = []

        result = open_update_proxy_dialog_namespace_runtime(
            namespace,
            open_dialog_runtime=lambda parent, **kwargs: calls.append((parent, kwargs)) or "proxy",
        )

        self.assertEqual(result, "proxy")
        parent, forwarded = calls[0]
        self.assertIs(parent, namespace["root"])
        self.assertEqual(forwarded["owner"], "owner")
        self.assertEqual(forwarded["repo"], "repo")
        self.assertEqual(forwarded["config"], "config")

    def test_auto_log_cleanup_namespace_runtimes_forward_callbacks(self):
        namespace = self.base_namespace()
        calls = []

        tick_result = run_auto_log_cleanup_tick_namespace_runtime(
            namespace,
            run_tick_runtime=lambda **kwargs: calls.append(("tick", kwargs)) or "ticked",
        )
        schedule_result = schedule_auto_log_cleanup_namespace_runtime(
            namespace,
            restart=False,
            first_delay_sec=9,
            schedule_runtime=lambda **kwargs: calls.append(("schedule", kwargs)) or "scheduled",
        )

        self.assertEqual(tick_result, "ticked")
        self.assertEqual(schedule_result, "scheduled")
        tick_kwargs = calls[0][1]
        self.assertEqual(tick_kwargs["state"], "cleanup_state")
        self.assertTrue(tick_kwargs["is_enabled"]())
        self.assertEqual(tick_kwargs["retention_days"](), 14)
        self.assertEqual(tick_kwargs["interval_hours"](), 6)
        self.assertEqual(tick_kwargs["root_after"]("tick"), ("after", ("tick",)))
        schedule_kwargs = calls[1][1]
        self.assertFalse(schedule_kwargs["restart"])
        self.assertEqual(schedule_kwargs["first_delay_sec"], 9)
        self.assertEqual(schedule_kwargs["root_after_cancel"]("old"), ("cancel", ("old",)))

    def test_check_update_and_prompt_namespace_runtime_forwards_messagebox_and_config(self):
        namespace = self.base_namespace()
        calls = []

        result = check_update_and_prompt_namespace_runtime(
            namespace,
            check_runtime=lambda **kwargs: calls.append(kwargs) or "checked",
            read_config_runtime=lambda config: ("proxy", f"api:{config}"),
        )

        self.assertEqual(result, "checked")
        forwarded = calls[0]
        self.assertEqual(forwarded["owner"], "owner")
        self.assertEqual(forwarded["repo"], "repo")
        self.assertEqual(forwarded["current_version"], "1.0.0")
        self.assertEqual(forwarded["get_update_config"](), ("proxy", "api:config"))
        self.assertEqual(forwarded["show_info"]("t", "m"), ("info", ("t", "m")))
        self.assertEqual(forwarded["open_url"]("https://example.test"), ("open_url", "https://example.test"))


if __name__ == "__main__":
    unittest.main()
