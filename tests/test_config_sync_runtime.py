import unittest

from sms_ui.config_sync_runtime import (
    ConfigFileWatchState,
    ConfigReloadFailureLogState,
    clear_config_reload_failure_runtime,
    config_file_signature_runtime,
    report_config_reload_failure_runtime,
    schedule_config_file_watch_runtime,
    stop_config_file_watch_runtime,
)


class FakeRootTimer:
    def __init__(self):
        self.callbacks = {}
        self.cancelled = []
        self.next_id = 1

    def after(self, delay_ms, callback):
        after_id = f"after-{self.next_id}"
        self.next_id += 1
        self.callbacks[after_id] = (delay_ms, callback)
        return after_id

    def after_cancel(self, after_id):
        self.cancelled.append(after_id)
        self.callbacks.pop(after_id, None)

    def run(self, after_id):
        _delay, callback = self.callbacks.pop(after_id)
        callback()


class ConfigSyncRuntimeTests(unittest.TestCase):
    def test_reload_failure_logs_first_attempt_then_periodic_summary_and_recovery(self):
        state = ConfigReloadFailureLogState()
        logs = []
        times = iter((0.0, 1.0, 59.0, 60.0))

        for _ in range(4):
            report_config_reload_failure_runtime(
                state,
                "MissingSectionHeaderError",
                log_error=logs.append,
                monotonic=lambda: next(times),
                log_interval_seconds=60,
            )

        self.assertEqual(len(logs), 2)
        self.assertIn("MissingSectionHeaderError", logs[0])
        self.assertIn("suppressed 3 repeated attempts", logs[1])
        self.assertEqual(state.consecutive_failures, 4)

        self.assertEqual(
            clear_config_reload_failure_runtime(state, log_error=logs.append),
            4,
        )
        self.assertIn("recovered after 4 failed attempts", logs[2])
        self.assertEqual(state, ConfigReloadFailureLogState())

    def test_reload_failure_log_does_not_include_exception_body(self):
        logs = []
        state = ConfigReloadFailureLogState()

        report_config_reload_failure_runtime(
            state,
            "MissingSectionHeaderError",
            log_error=logs.append,
            monotonic=lambda: 0.0,
        )

        self.assertEqual(len(logs), 1)
        self.assertNotIn("device_secret", logs[0])
        self.assertEqual(logs[0], "Reload shared UI config failed (MissingSectionHeaderError)")

    def test_config_file_signature_uses_mtime_size_and_identity(self):
        stat_result = type(
            "StatResult",
            (),
            {"st_mtime_ns": 123, "st_size": 456, "st_ino": 789},
        )()

        self.assertEqual(
            config_file_signature_runtime("config.ini", stat_file=lambda _path: stat_result),
            (123, 456, 789),
        )

    def test_config_file_signature_returns_none_for_missing_file(self):
        def missing(_path):
            raise FileNotFoundError("missing")

        self.assertIsNone(config_file_signature_runtime("config.ini", stat_file=missing))

    def test_file_watch_only_reloads_after_signature_changes(self):
        root = FakeRootTimer()
        state = ConfigFileWatchState()
        signatures = iter(((1, 10, 1), (1, 10, 1), (2, 10, 2)))
        changes = []

        first_id = schedule_config_file_watch_runtime(
            state=state,
            config_file="config.ini",
            interval_ms=1000,
            root_after=root.after,
            root_after_cancel=root.after_cancel,
            tk_alive=lambda: True,
            is_stopping=lambda: False,
            on_change=lambda: changes.append("reload"),
            signature_func=lambda _path: next(signatures),
        )

        root.run(first_id)
        self.assertEqual(changes, [])
        second_id = state.after_id
        root.run(second_id)

        self.assertEqual(changes, ["reload"])
        self.assertNotEqual(state.after_id, second_id)

    def test_file_watch_stops_without_rescheduling_during_shutdown(self):
        root = FakeRootTimer()
        state = ConfigFileWatchState()
        stopping = {"value": False}

        after_id = schedule_config_file_watch_runtime(
            state=state,
            config_file="config.ini",
            interval_ms=1000,
            root_after=root.after,
            root_after_cancel=root.after_cancel,
            tk_alive=lambda: True,
            is_stopping=lambda: stopping["value"],
            on_change=lambda: self.fail("must not reload while stopping"),
            signature_func=lambda _path: (1, 1, 1),
        )

        stopping["value"] = True
        root.run(after_id)
        self.assertIsNone(state.after_id)
        self.assertEqual(root.callbacks, {})

    def test_file_watch_retries_same_signature_after_reload_failure(self):
        root = FakeRootTimer()
        state = ConfigFileWatchState()
        signatures = iter(((1, 10, 1), (2, 10, 2), (2, 10, 2), (2, 10, 2)))
        results = iter((False, ("短信弹窗",)))
        attempts = []

        first_id = schedule_config_file_watch_runtime(
            state=state,
            config_file="config.ini",
            interval_ms=1000,
            root_after=root.after,
            root_after_cancel=root.after_cancel,
            tk_alive=lambda: True,
            is_stopping=lambda: False,
            on_change=lambda: attempts.append("reload") or next(results),
            signature_func=lambda _path: next(signatures),
        )

        root.run(first_id)
        self.assertEqual(attempts, ["reload"])
        self.assertEqual(state.signature, (1, 10, 1))

        root.run(state.after_id)
        self.assertEqual(attempts, ["reload", "reload"])
        self.assertEqual(state.signature, (2, 10, 2))

        root.run(state.after_id)
        self.assertEqual(attempts, ["reload", "reload"])

    def test_stop_file_watch_invalidates_and_cancels_pending_callback(self):
        root = FakeRootTimer()
        state = ConfigFileWatchState(after_id="after-1", generation=2)
        root.callbacks["after-1"] = (1000, lambda: None)

        result = stop_config_file_watch_runtime(state, root_after_cancel=root.after_cancel)

        self.assertEqual(result, "after-1")
        self.assertEqual(state.generation, 3)
        self.assertIsNone(state.after_id)
        self.assertEqual(root.cancelled, ["after-1"])


if __name__ == "__main__":
    unittest.main()
