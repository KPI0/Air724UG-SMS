import configparser
import unittest

from sms_core.threading_runtime import WorkerThreadRegistry
from sms_ui.maintenance_runtime import (
    AutoLogCleanupState,
    apply_log_cleanup_runtime,
    ensure_update_config,
    log_cleanup_status,
    normalized_retention_days,
    open_log_dir_runtime,
    open_log_cleanup_dialog_app_runtime,
    open_log_cleanup_dialog_runtime,
    open_update_proxy_dialog_runtime,
    run_auto_log_cleanup_tick_app_runtime,
    run_auto_log_cleanup_tick_runtime,
    save_update_proxy_config,
    schedule_auto_log_cleanup_app_runtime,
    schedule_auto_log_cleanup_runtime,
    test_update_proxy_async,
)


class ImmediateThread:
    def __init__(self, target, daemon):
        self.target = target
        self.daemon = daemon
        self.started = False

    def start(self):
        self.started = True
        self.target()


class DeferredThread(ImmediateThread):
    def start(self):
        self.started = True


class MaintenanceRuntimeTests(unittest.TestCase):
    def test_log_cleanup_status_includes_values(self):
        message = log_cleanup_status(14, 6)

        self.assertIn("14", message)
        self.assertIn("6", message)

    def test_normalized_retention_days_uses_safe_default_for_negative_values(self):
        self.assertEqual(normalized_retention_days(-1), 30)
        self.assertEqual(normalized_retention_days(7), 7)

    def test_open_log_dir_runtime_opens_existing_path(self):
        calls = []

        result = open_log_dir_runtime(
            "logs",
            path_abspath=lambda path: f"ABS/{path}",
            path_exists=lambda path: True,
            open_path=lambda path: calls.append(("open", path)),
            show_message=lambda *args: calls.append(("message", args)),
        )

        self.assertEqual(result, "opened")
        self.assertEqual(calls, [("open", "ABS/logs")])

    def test_open_log_dir_runtime_reports_missing_path(self):
        calls = []

        result = open_log_dir_runtime(
            "logs",
            path_abspath=lambda path: f"ABS/{path}",
            path_exists=lambda path: False,
            open_path=lambda path: calls.append(("open", path)),
            show_message=lambda *args: calls.append(("message", args)),
        )

        self.assertEqual(result, "missing")
        self.assertEqual(calls[0][0], "message")
        self.assertEqual(calls[0][1][0], "warning")

    def test_open_log_dir_runtime_reports_open_errors(self):
        calls = []

        result = open_log_dir_runtime(
            "logs",
            path_abspath=lambda path: f"ABS/{path}",
            path_exists=lambda path: True,
            open_path=lambda path: (_ for _ in ()).throw(RuntimeError("boom")),
            show_message=lambda *args: calls.append(("message", args)),
        )

        self.assertEqual(result, "error")
        self.assertEqual(calls[0][1][0], "error")
        self.assertIn("boom", calls[0][1][2])

    def test_schedule_auto_log_cleanup_runtime_schedules_first_tick(self):
        state = AutoLogCleanupState()
        calls = []

        schedule_auto_log_cleanup_runtime(
            state=state,
            restart=True,
            first_delay_sec=12,
            is_enabled=lambda: True,
            tk_alive=lambda: True,
            root_after=lambda delay_ms, callback: calls.append((delay_ms, callback)) or "after-1",
            root_after_cancel=lambda after_id: calls.append(("cancel", after_id)),
            tick_callback="tick",
            is_main_thread=lambda: True,
            ui_post=lambda callback: callback(),
        )

        self.assertEqual(state.after_id, "after-1")
        self.assertEqual(calls, [(12000, "tick")])

    def test_schedule_auto_log_cleanup_runtime_cancels_existing_timer(self):
        state = AutoLogCleanupState(after_id="old")
        calls = []

        schedule_auto_log_cleanup_runtime(
            state=state,
            restart=True,
            first_delay_sec=1,
            is_enabled=lambda: True,
            tk_alive=lambda: True,
            root_after=lambda delay_ms, callback: calls.append(("after", delay_ms, callback)) or "new",
            root_after_cancel=lambda after_id: calls.append(("cancel", after_id)),
            tick_callback="tick",
            is_main_thread=lambda: True,
            ui_post=lambda callback: callback(),
        )

        self.assertEqual(state.after_id, "new")
        self.assertEqual(calls, [("cancel", "old"), ("after", 1000, "tick")])

    def test_schedule_auto_log_cleanup_runtime_posts_from_background_thread(self):
        state = AutoLogCleanupState()
        posted = []

        schedule_auto_log_cleanup_runtime(
            state=state,
            restart=True,
            first_delay_sec=1,
            is_enabled=lambda: True,
            tk_alive=lambda: True,
            root_after=lambda delay_ms, callback: "after",
            root_after_cancel=lambda after_id: None,
            tick_callback="tick",
            is_main_thread=lambda: False,
            ui_post=posted.append,
        )

        self.assertEqual(state.after_id, None)
        self.assertEqual(len(posted), 1)
        posted[0]()
        self.assertEqual(state.after_id, "after")

    def test_run_auto_log_cleanup_tick_runtime_cleans_and_reschedules(self):
        state = AutoLogCleanupState()
        ui_messages = []
        calls = []

        run_auto_log_cleanup_tick_runtime(
            state=state,
            is_enabled=lambda: True,
            retention_days=lambda: -3,
            interval_hours=lambda: 2,
            cleanup_old_logs=lambda days: calls.append(("cleanup", days)) or 4,
            system_ui=lambda message, tag="normal": ui_messages.append((message, tag)),
            tk_alive=lambda: True,
            root_after=lambda delay_ms, callback: calls.append(("after", delay_ms, callback)) or "after-2",
            tick_callback="tick",
            is_main_thread=lambda: True,
            ui_post=lambda callback: callback(),
        )

        self.assertEqual(calls, [("cleanup", 30), ("after", 7200000, "tick")])
        self.assertEqual(state.after_id, "after-2")
        self.assertEqual(ui_messages[0][1], "normal")
        self.assertIn("4", ui_messages[0][0])

    def test_run_auto_log_cleanup_tick_runtime_handles_disabled_and_errors(self):
        state = AutoLogCleanupState(after_id="old")

        run_auto_log_cleanup_tick_runtime(
            state=state,
            is_enabled=lambda: False,
            retention_days=lambda: 1,
            interval_hours=lambda: 1,
            cleanup_old_logs=lambda days: 0,
            system_ui=lambda *_: None,
            tk_alive=lambda: True,
            root_after=lambda delay_ms, callback: "after",
            tick_callback="tick",
            is_main_thread=lambda: True,
            ui_post=lambda callback: callback(),
        )

        self.assertIsNone(state.after_id)

        messages = []
        run_auto_log_cleanup_tick_runtime(
            state=state,
            is_enabled=lambda: True,
            retention_days=lambda: 1,
            interval_hours=lambda: 1,
            cleanup_old_logs=lambda days: (_ for _ in ()).throw(RuntimeError("boom")),
            system_ui=lambda message, tag="normal": messages.append((message, tag)),
            tk_alive=lambda: False,
            root_after=lambda delay_ms, callback: "after",
            tick_callback="tick",
            is_main_thread=lambda: True,
            ui_post=lambda callback: callback(),
        )

        self.assertIsNone(state.after_id)
        self.assertIn("boom", messages[0][0])

    def test_apply_log_cleanup_runtime_updates_config_and_schedules_cleanup(self):
        config = configparser.ConfigParser()
        saved = []
        state = []
        ui_messages = []
        scheduled = []

        apply_log_cleanup_runtime(
            30,
            config=config,
            save_config=lambda: saved.append("saved"),
            set_cleanup_state=lambda days, enabled: state.append((days, enabled)),
            system_ui=lambda message, tag="normal": ui_messages.append((message, tag)),
            schedule_cleanup=lambda **kwargs: scheduled.append(kwargs),
            interval_hours=12,
        )

        self.assertEqual(state, [(30, True)])
        self.assertEqual(config.get("ui", "auto_log_cleanup"), "1")
        self.assertEqual(config.get("ui", "log_retention_days"), "30")
        self.assertEqual(saved, ["saved"])
        self.assertEqual(ui_messages[0][1], "normal")
        self.assertEqual(scheduled, [{"restart": True, "first_delay_sec": 60}])

    def test_apply_log_cleanup_runtime_reports_save_failure(self):
        config = configparser.ConfigParser()
        config["ui"] = {
            "auto_log_cleanup": "0",
            "log_retention_days": "7",
            "extra": "keep",
        }
        state = []
        ui_messages = []
        scheduled = []

        result = apply_log_cleanup_runtime(
            30,
            config=config,
            save_config=lambda: False,
            set_cleanup_state=lambda days, enabled: state.append((days, enabled)),
            system_ui=lambda message, tag="normal": ui_messages.append((message, tag)),
            schedule_cleanup=lambda **kwargs: scheduled.append(kwargs),
            interval_hours=12,
        )

        self.assertFalse(result)
        self.assertEqual(
            dict(config.items("ui")),
            {
                "auto_log_cleanup": "0",
                "log_retention_days": "7",
                "extra": "keep",
            },
        )
        self.assertEqual(state, [])
        self.assertEqual(scheduled, [])
        self.assertIn("保存失败", ui_messages[0][0])

    def test_open_log_cleanup_dialog_runtime_wires_apply_callback(self):
        config = configparser.ConfigParser()
        applied = []
        opened = {}

        def open_dialog(parent, current_days, interval_hours, apply_cleanup, center_window):
            opened.update(
                parent=parent,
                current_days=current_days,
                interval_hours=interval_hours,
                center_window=center_window,
            )
            apply_cleanup(9)

        open_log_cleanup_dialog_runtime(
            "root",
            7,
            3,
            config=config,
            save_config=lambda: None,
            set_cleanup_state=lambda days, enabled: applied.append((days, enabled)),
            system_ui=lambda *_: None,
            schedule_cleanup=lambda **_: None,
            center_window="center",
            open_dialog=open_dialog,
        )

        self.assertEqual(opened, {
            "parent": "root",
            "current_days": 7,
            "interval_hours": 3,
            "center_window": "center",
        })
        self.assertEqual(applied, [(9, True)])
        self.assertEqual(config.get("ui", "log_retention_days"), "9")

    def test_open_log_cleanup_dialog_runtime_propagates_save_failure(self):
        results = []

        def open_dialog(_parent, _days, _interval, apply_cleanup, _center):
            results.append(apply_cleanup(9))

        open_log_cleanup_dialog_runtime(
            "root",
            7,
            3,
            config=configparser.ConfigParser(),
            save_config=lambda: False,
            set_cleanup_state=lambda *_: self.fail("state changed after failed save"),
            system_ui=lambda *_: None,
            schedule_cleanup=lambda **_: self.fail("cleanup scheduled after failed save"),
            center_window="center",
            open_dialog=open_dialog,
        )

        self.assertEqual(results, [False])

    def test_open_log_cleanup_dialog_app_runtime_reads_current_values(self):
        calls = []

        result = open_log_cleanup_dialog_app_runtime(
            "root",
            get_retention_days=lambda: 14,
            get_interval_hours=lambda: 6,
            config=configparser.ConfigParser(),
            save_config=lambda: calls.append("save"),
            set_cleanup_state=lambda *args: calls.append(("state", args)),
            system_ui=lambda *args: calls.append(("ui", args)),
            schedule_cleanup=lambda **kwargs: calls.append(("schedule", kwargs)),
            center_window="center",
            open_dialog=lambda *args: calls.append(("open", args)) or args[3](21),
        )

        self.assertIsNone(result)
        self.assertEqual(calls[0][0], "open")
        self.assertEqual(calls[0][1][1], 14)
        self.assertEqual(calls[0][1][2], 6)
        self.assertIn(("state", (21, True)), calls)

    def test_auto_log_cleanup_app_runtimes_forward_dependencies(self):
        calls = []
        state = AutoLogCleanupState()

        run_auto_log_cleanup_tick_app_runtime(
            state=state,
            is_enabled=lambda: True,
            retention_days=lambda: 7,
            interval_hours=lambda: 3,
            cleanup_old_logs=lambda days: calls.append(("cleanup", days)),
            system_ui=lambda *args: calls.append(("ui", args)),
            tk_alive=lambda: True,
            root_after=lambda *args: calls.append(("after", args)),
            tick_callback="tick",
            ui_post=lambda callback: callback(),
            run_tick_runtime=lambda **kwargs: calls.append(("tick_runtime", kwargs)) or "tick-result",
        )
        schedule_result = schedule_auto_log_cleanup_app_runtime(
            state=state,
            restart=False,
            first_delay_sec=5,
            is_enabled=lambda: True,
            tk_alive=lambda: True,
            root_after=lambda *args: calls.append(("after", args)),
            root_after_cancel=lambda *args: calls.append(("cancel", args)),
            tick_callback="tick",
            ui_post=lambda callback: callback(),
            schedule_runtime=lambda **kwargs: calls.append(("schedule_runtime", kwargs)) or "schedule-result",
        )

        self.assertEqual(calls[0][0], "tick_runtime")
        self.assertEqual(schedule_result, "schedule-result")
        self.assertEqual(calls[1][0], "schedule_runtime")

    def test_ensure_update_config_creates_defaults_only_when_missing(self):
        config = configparser.ConfigParser()

        self.assertTrue(ensure_update_config(config, {"api_proxy_base": "api", "proxy_base": "proxy"}))
        self.assertEqual(config.get("update", "api_proxy_base"), "api")

        self.assertFalse(ensure_update_config(config, {"api_proxy_base": "next", "proxy_base": "next"}))
        self.assertEqual(config.get("update", "api_proxy_base"), "api")

    def test_save_update_proxy_config_normalizes_values(self):
        config = configparser.ConfigParser()
        saved = []

        result = save_update_proxy_config(config, "api.example.com", "proxy.example.com", lambda: saved.append("saved"))

        self.assertEqual(config.get("update", "api_proxy_base"), "https://api.example.com/")
        self.assertEqual(config.get("update", "proxy_base"), "https://proxy.example.com/")
        self.assertEqual(saved, ["saved"])
        self.assertTrue(result)

    def test_save_update_proxy_config_rejects_save_failure(self):
        config = configparser.ConfigParser()

        with self.assertRaises(RuntimeError):
            save_update_proxy_config(config, "api.example.com", "proxy.example.com", lambda: False)
        self.assertFalse(config.has_section("update"))

    def test_test_update_proxy_async_posts_success_and_error_results(self):
        success = []
        errors = []

        thread = test_update_proxy_async(
            "owner",
            "repo",
            "api",
            "proxy",
            success.append,
            errors.append,
            ui_post=lambda callback: callback(),
            connectivity_func=lambda *args: {"args": args},
            formatter=lambda result: f"ok:{result['args'][0]}",
            thread_factory=ImmediateThread,
        )

        self.assertTrue(thread.started)
        self.assertEqual(success, ["ok:owner"])
        self.assertEqual(errors, [])

        test_update_proxy_async(
            "owner",
            "repo",
            "api",
            "proxy",
            success.append,
            errors.append,
            ui_post=lambda callback: callback(),
            connectivity_func=lambda *_: (_ for _ in ()).throw(RuntimeError("boom")),
            thread_factory=ImmediateThread,
        )

        self.assertEqual(errors, ["boom"])

    def test_update_proxy_worker_is_registered_until_it_finishes(self):
        registry = WorkerThreadRegistry()

        thread = test_update_proxy_async(
            "owner",
            "repo",
            "api",
            "proxy",
            lambda _value: None,
            lambda _value: None,
            ui_post=lambda callback: callback(),
            connectivity_func=lambda *args: {"args": args},
            thread_factory=DeferredThread,
            thread_registry=registry,
        )

        self.assertEqual(registry.snapshot(), (thread,))
        thread.target()
        self.assertEqual(registry.snapshot(), ())

    def test_update_proxy_skips_work_and_late_ui_callbacks_during_shutdown(self):
        state = {"stopping": True}
        posted = []
        calls = []

        thread = test_update_proxy_async(
            "owner",
            "repo",
            "api",
            "proxy",
            lambda value: calls.append(("success", value)),
            lambda value: calls.append(("error", value)),
            ui_post=posted.append,
            connectivity_func=lambda *args: self.fail("proxy test started during shutdown"),
            thread_factory=ImmediateThread,
            is_stopping=lambda: state["stopping"],
        )

        self.assertIsNone(thread)
        self.assertEqual(posted, [])

        state["stopping"] = False
        test_update_proxy_async(
            "owner",
            "repo",
            "api",
            "proxy",
            lambda value: calls.append(("success", value)),
            lambda value: calls.append(("error", value)),
            ui_post=posted.append,
            connectivity_func=lambda *args: "connected",
            formatter=lambda result: result,
            thread_factory=ImmediateThread,
            is_stopping=lambda: state["stopping"],
        )
        self.assertEqual(len(posted), 1)
        state["stopping"] = True
        posted[0]()
        self.assertEqual(calls, [])

    def test_open_update_proxy_dialog_runtime_wires_save_and_test_callbacks(self):
        config = configparser.ConfigParser()
        saved = []
        tested = []
        opened = {}

        def open_dialog(parent, api_proxy_base, proxy_base, save, test_connection, center_window):
            opened.update(
                parent=parent,
                api_proxy_base=api_proxy_base,
                proxy_base=proxy_base,
                center_window=center_window,
            )
            opened["save_result"] = save("api.example.com", "proxy.example.com", "win")
            test_connection("api", "proxy", lambda value: None, lambda value: None)

        open_update_proxy_dialog_runtime(
            "root",
            config=config,
            owner="owner",
            repo="repo",
            save_config=lambda: saved.append("saved"),
            ui_post=lambda callback: callback(),
            center_window="center",
            open_dialog=open_dialog,
            test_async=lambda *args: tested.append(args),
        )

        self.assertEqual(opened["parent"], "root")
        self.assertTrue(opened["api_proxy_base"])
        self.assertTrue(opened["proxy_base"])
        self.assertEqual(opened["center_window"], "center")
        self.assertTrue(opened["save_result"])
        self.assertEqual(config.get("update", "api_proxy_base"), "https://api.example.com/")
        self.assertEqual(config.get("update", "proxy_base"), "https://proxy.example.com/")
        self.assertEqual(saved, ["saved"])
        self.assertEqual(tested[0][:4], ("owner", "repo", "api", "proxy"))

    def test_open_update_proxy_dialog_runtime_propagates_save_failure(self):
        config = configparser.ConfigParser()
        opened = {}

        def open_dialog(_parent, _api_proxy_base, _proxy_base, save, _test_connection, _center_window):
            opened["save_result"] = save("api.example.com", "proxy.example.com", "win")

        with self.assertRaises(RuntimeError):
            open_update_proxy_dialog_runtime(
                "root",
                config=config,
                owner="owner",
                repo="repo",
                save_config=lambda: False,
                ui_post=lambda callback: callback(),
                center_window="center",
                open_dialog=open_dialog,
                test_async=lambda *args: None,
            )
        self.assertFalse(config.has_section("update"))


if __name__ == "__main__":
    unittest.main()
