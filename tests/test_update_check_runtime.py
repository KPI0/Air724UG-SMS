import unittest
from types import SimpleNamespace

from sms_core.threading_runtime import SingleFlightTaskState, WorkerThreadRegistry
from sms_ui.update_check_runtime import (
    check_update_and_prompt_runtime,
    handle_update_check_plan,
    read_update_config_runtime,
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


class BrokenThread(ImmediateThread):
    def start(self):
        raise RuntimeError("start failed")


class UpdateCheckRuntimeTests(unittest.TestCase):
    def test_read_update_config_runtime_normalizes_config_values(self):
        class Config:
            def get(self, section, key, fallback=""):
                values = {
                    ("update", "proxy_base"): "https://proxy.example/path",
                    ("update", "api_proxy_base"): "https://api.example/path",
                }
                return values.get((section, key), fallback)

        proxy_base, api_proxy_base = read_update_config_runtime(Config())

        self.assertEqual(proxy_base, "https://proxy.example/path/")
        self.assertEqual(api_proxy_base, "https://api.example/path/")

    def test_handle_update_check_plan_reports_latest(self):
        calls = []

        result = handle_update_check_plan(
            SimpleNamespace(kind="latest", latest_tag="v1.0.0", download_url=""),
            "1.0.0",
            show_info=lambda title, message: calls.append(("info", title, message)),
            show_warning=lambda title, message: calls.append(("warning", title, message)),
            ask_open_download=lambda title, message: False,
            open_url=lambda url: calls.append(("open", url)),
        )

        self.assertEqual(result, "latest")
        self.assertEqual(calls[0][0], "info")
        self.assertIn("1.0.0", calls[0][2])

    def test_handle_update_check_plan_reports_missing_zip(self):
        calls = []

        result = handle_update_check_plan(
            SimpleNamespace(kind="no_zip", latest_tag="v1.0.1", download_url=""),
            "1.0.0",
            show_info=lambda title, message: calls.append(("info", title, message)),
            show_warning=lambda title, message: calls.append(("warning", title, message)),
            ask_open_download=lambda title, message: False,
            open_url=lambda url: calls.append(("open", url)),
        )

        self.assertEqual(result, "no_zip")
        self.assertEqual(calls[0][0], "warning")
        self.assertIn("v1.0.1", calls[0][2])

    def test_handle_update_check_plan_opens_download_when_confirmed(self):
        calls = []

        result = handle_update_check_plan(
            SimpleNamespace(kind="update", latest_tag="v1.0.1", download_url="https://example.com/app.zip"),
            "1.0.0",
            show_info=lambda title, message: calls.append(("info", title, message)),
            show_warning=lambda title, message: calls.append(("warning", title, message)),
            ask_open_download=lambda title, message: calls.append(("ask", title, message)) or True,
            open_url=lambda url: calls.append(("open", url)),
        )

        self.assertEqual(result, "update")
        self.assertEqual(calls[0][0], "ask")
        self.assertEqual(calls[1], ("open", "https://example.com/app.zip"))

    def test_handle_update_check_plan_surfaces_url_when_open_fails(self):
        # If the browser fails to open, the download URL must be shown so the
        # user can copy it manually instead of being left with nothing.
        calls = []

        result = handle_update_check_plan(
            SimpleNamespace(kind="update", latest_tag="v1.0.1", download_url="https://example.com/app.zip"),
            "1.0.0",
            show_info=lambda title, message: calls.append(("info", title, message)),
            show_warning=lambda title, message: calls.append(("warning", title, message)),
            ask_open_download=lambda title, message: True,
            open_url=lambda url: (_ for _ in ()).throw(RuntimeError("no browser")),
        )

        self.assertEqual(result, "update")
        warnings = [c for c in calls if c[0] == "warning"]
        self.assertEqual(len(warnings), 1)
        self.assertIn("https://example.com/app.zip", warnings[0][2])
        self.assertIn("no browser", warnings[0][2])

    def test_check_update_and_prompt_runtime_posts_plan_to_ui(self):
        calls = []

        thread = check_update_and_prompt_runtime(
            owner="owner",
            repo="repo",
            current_version="1.0.0",
            get_update_config=lambda: ("proxy", "api"),
            ui_post=lambda callback: calls.append(("posted", callback())) ,
            show_info=lambda title, message: calls.append(("info", title, message)),
            show_warning=lambda title, message: calls.append(("warning", title, message)),
            show_error=lambda title, message: calls.append(("error", title, message)),
            ask_open_download=lambda title, message: False,
            open_url=lambda url: calls.append(("open", url)),
            check_latest=lambda *args: SimpleNamespace(kind="latest", latest_tag="v1.0.0", download_url=""),
            thread_factory=ImmediateThread,
        )

        self.assertTrue(thread.started)
        self.assertEqual(calls[0][0], "info")
        self.assertEqual(calls[1], ("posted", "latest"))

    def test_check_update_and_prompt_runtime_posts_errors_to_ui(self):
        calls = []

        check_update_and_prompt_runtime(
            owner="owner",
            repo="repo",
            current_version="1.0.0",
            get_update_config=lambda: ("proxy", "api"),
            ui_post=lambda callback: callback(),
            show_info=lambda title, message: calls.append(("info", title, message)),
            show_warning=lambda title, message: calls.append(("warning", title, message)),
            show_error=lambda title, message: calls.append(("error", title, message)),
            ask_open_download=lambda title, message: False,
            open_url=lambda url: calls.append(("open", url)),
            check_latest=lambda *args: (_ for _ in ()).throw(RuntimeError("down")),
            thread_factory=ImmediateThread,
        )

        self.assertEqual(calls, [("error", "检测更新失败", "down")])

    def test_check_update_worker_is_registered_until_it_finishes(self):
        registry = WorkerThreadRegistry()

        thread = check_update_and_prompt_runtime(
            owner="owner",
            repo="repo",
            current_version="1.0.0",
            get_update_config=lambda: ("proxy", "api"),
            ui_post=lambda callback: callback(),
            show_info=lambda *_: None,
            show_warning=lambda *_: None,
            show_error=lambda *_: None,
            ask_open_download=lambda *_: False,
            open_url=lambda *_: None,
            check_latest=lambda *args: SimpleNamespace(
                kind="latest", latest_tag="v1.0.0", download_url=""
            ),
            thread_factory=DeferredThread,
            thread_registry=registry,
        )

        self.assertEqual(registry.snapshot(), (thread,))
        thread.target()
        self.assertEqual(registry.snapshot(), ())

    def test_check_update_allows_only_one_task_until_result_is_handled(self):
        registry = WorkerThreadRegistry()
        task_state = SingleFlightTaskState()
        posted = []
        calls = []

        def start_check():
            return check_update_and_prompt_runtime(
                owner="owner",
                repo="repo",
                current_version="1.0.0",
                get_update_config=lambda: ("proxy", "api"),
                ui_post=posted.append,
                show_info=lambda *args: calls.append(("info", args)),
                show_warning=lambda *args: calls.append(("warning", args)),
                show_error=lambda *args: calls.append(("error", args)),
                ask_open_download=lambda *_: False,
                open_url=lambda *_: None,
                check_latest=lambda *args: SimpleNamespace(
                    kind="latest", latest_tag="v1.0.0", download_url=""
                ),
                thread_factory=DeferredThread,
                thread_registry=registry,
                task_state=task_state,
            )

        first = start_check()
        self.assertIsNotNone(first)
        self.assertTrue(task_state.is_active())
        self.assertIsNone(start_check())
        self.assertEqual(registry.snapshot(), (first,))

        first.target()
        self.assertEqual(registry.snapshot(), ())
        self.assertTrue(task_state.is_active())
        self.assertIsNone(start_check())
        self.assertEqual(len(posted), 1)

        posted[0]()
        self.assertFalse(task_state.is_active())
        self.assertEqual(calls[0][0], "info")
        self.assertIsNotNone(start_check())

    def test_check_update_releases_single_flight_state_when_thread_start_fails(self):
        task_state = SingleFlightTaskState()

        with self.assertRaisesRegex(RuntimeError, "start failed"):
            check_update_and_prompt_runtime(
                owner="owner",
                repo="repo",
                current_version="1.0.0",
                get_update_config=lambda: ("proxy", "api"),
                ui_post=lambda callback: callback(),
                show_info=lambda *_: None,
                show_warning=lambda *_: None,
                show_error=lambda *_: None,
                ask_open_download=lambda *_: False,
                open_url=lambda *_: None,
                thread_factory=BrokenThread,
                task_state=task_state,
            )

        self.assertFalse(task_state.is_active())

    def test_check_update_releases_single_flight_state_when_ui_enqueue_fails(self):
        task_state = SingleFlightTaskState()
        calls = []

        thread = check_update_and_prompt_runtime(
            owner="owner",
            repo="repo",
            current_version="1.0.0",
            get_update_config=lambda: ("proxy", "api"),
            ui_post=lambda _callback: False,
            show_info=lambda *args: calls.append(("info", args)),
            show_warning=lambda *args: calls.append(("warning", args)),
            show_error=lambda *args: calls.append(("error", args)),
            ask_open_download=lambda *_: False,
            open_url=lambda *_: None,
            check_latest=lambda *args: SimpleNamespace(
                kind="latest", latest_tag="v1.0.0", download_url=""
            ),
            thread_factory=ImmediateThread,
            task_state=task_state,
        )

        self.assertIsNotNone(thread)
        self.assertFalse(task_state.is_active())
        self.assertEqual(calls, [])

    def test_check_update_skips_work_and_late_ui_callbacks_during_shutdown(self):
        state = {"stopping": True}
        task_state = SingleFlightTaskState()
        posted = []
        calls = []

        thread = check_update_and_prompt_runtime(
            owner="owner",
            repo="repo",
            current_version="1.0.0",
            get_update_config=lambda: ("proxy", "api"),
            ui_post=posted.append,
            show_info=lambda *args: calls.append(("info", args)),
            show_warning=lambda *args: calls.append(("warning", args)),
            show_error=lambda *args: calls.append(("error", args)),
            ask_open_download=lambda *_: False,
            open_url=lambda *_: None,
            check_latest=lambda *args: self.fail("network check started during shutdown"),
            thread_factory=ImmediateThread,
            is_stopping=lambda: state["stopping"],
            task_state=task_state,
        )

        self.assertIsNone(thread)
        self.assertEqual(posted, [])

        state["stopping"] = False
        check_update_and_prompt_runtime(
            owner="owner",
            repo="repo",
            current_version="1.0.0",
            get_update_config=lambda: ("proxy", "api"),
            ui_post=posted.append,
            show_info=lambda *args: calls.append(("info", args)),
            show_warning=lambda *args: calls.append(("warning", args)),
            show_error=lambda *args: calls.append(("error", args)),
            ask_open_download=lambda *_: False,
            open_url=lambda *_: None,
            check_latest=lambda *args: SimpleNamespace(
                kind="latest", latest_tag="v1.0.0", download_url=""
            ),
            thread_factory=ImmediateThread,
            is_stopping=lambda: state["stopping"],
            task_state=task_state,
        )
        self.assertEqual(len(posted), 1)
        self.assertTrue(task_state.is_active())
        state["stopping"] = True
        posted[0]()
        self.assertEqual(calls, [])
        self.assertFalse(task_state.is_active())


if __name__ == "__main__":
    unittest.main()
