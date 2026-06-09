import queue
import unittest

from sms_ui.ui_log_namespace_runtime import (
    flush_pending_ui_logs_namespace_runtime,
    log_early_namespace_runtime,
    log_file_only_namespace_runtime,
    log_namespace_runtime,
    main_text_available_namespace_runtime,
    safe_insert_main_text_namespace_runtime,
    system_ui_namespace_runtime,
    ui_only_namespace_runtime,
    ui_post_namespace_runtime,
)


class FakeFileQueue:
    def __init__(self):
        self.items = []

    def put_nowait(self, item):
        self.items.append(item)


class FakeTk:
    END = "END"


class FakeText:
    def __init__(self, exists=True):
        self.exists = exists
        self.calls = []

    def winfo_exists(self):
        return self.exists

    def yview(self):
        return (0.0, 1.0)

    def insert(self, *args):
        self.calls.append(("insert", args))

    def index(self, _index):
        return "1.0"

    def see(self, *args):
        self.calls.append(("see", args))


class UiLogNamespaceRuntimeTests(unittest.TestCase):
    def make_namespace(self):
        calls = []
        return {
            "calls": calls,
            "UI_TASK_QUEUE": queue.Queue(maxsize=2),
            "FILE_LOG_Q": FakeFileQueue(),
            "LOG_DIR": "logs",
            "LOG_PREFIX": "COM5",
            "PENDING_UI_LOGS": queue.Queue(),
            "tk": FakeTk,
            "text_area": FakeText(),
            "log_file_only": lambda msg: calls.append(("file_only", msg)),
            "main_text_available": lambda: True,
            "safe_insert_main_text": lambda *args: calls.append(("insert", args)),
            "run_on_ui_thread": lambda callback, ui_post: calls.append(("run", ui_post)) or callback(),
            "ui_post": "ui_post",
            "tk_alive": lambda: True,
            "schedule_delayed_ui": lambda callback: calls.append(("schedule", callback)) or callback(),
            "log_early": lambda *args: calls.append(("early", args)),
        }

    def test_ui_post_namespace_runtime_queues_task_and_logs_when_full(self):
        namespace = self.make_namespace()
        namespace["UI_TASK_QUEUE"].put_nowait(("a", (), {}))
        namespace["UI_TASK_QUEUE"].put_nowait(("b", (), {}))

        ui_post_namespace_runtime(namespace, lambda: None, 1, name="value")

        self.assertIn(("file_only", "⚠️ UI_TASK_QUEUE 已满：丢弃一次 UI 任务"), namespace["calls"])

    def test_log_file_only_namespace_runtime_writes_system_log(self):
        namespace = self.make_namespace()

        result = log_file_only_namespace_runtime(namespace, "msg")

        self.assertIsNotNone(result)
        self.assertEqual(namespace["FILE_LOG_Q"].items[0][0].startswith("logs"), True)
        self.assertIn("msg", namespace["FILE_LOG_Q"].items[0][1])

    def test_ui_only_and_system_ui_forward_namespace_callbacks(self):
        namespace = self.make_namespace()

        self.assertEqual(ui_only_namespace_runtime(namespace, "ui", "tag"), "inserted")
        self.assertEqual(system_ui_namespace_runtime(namespace, "system", "normal"), "scheduled")

        self.assertIn(("insert", ("ui", "tag")), namespace["calls"])
        self.assertIn(("file_only", "system"), namespace["calls"])
        self.assertIn(("insert", ("system", "normal")), namespace["calls"])

    def test_log_early_namespace_runtime_writes_file_and_pending_queue(self):
        namespace = self.make_namespace()

        log_early_namespace_runtime(namespace, "early", "warn")

        self.assertIn(("file_only", "early"), namespace["calls"])
        self.assertEqual(namespace["PENDING_UI_LOGS"].get_nowait(), ("early", "warn"))

    def test_main_text_available_namespace_runtime_checks_widget(self):
        namespace = self.make_namespace()

        self.assertTrue(main_text_available_namespace_runtime(namespace))
        namespace["text_area"] = FakeText(exists=False)
        self.assertFalse(main_text_available_namespace_runtime(namespace))
        namespace.pop("text_area")
        self.assertFalse(main_text_available_namespace_runtime(namespace))

    def test_safe_insert_and_flush_pending_namespace_runtime(self):
        namespace = self.make_namespace()
        namespace["PENDING_UI_LOGS"].put_nowait(("pending", "tag"))

        self.assertTrue(safe_insert_main_text_namespace_runtime(namespace, "msg", "normal"))
        self.assertEqual(flush_pending_ui_logs_namespace_runtime(namespace), 1)

        self.assertEqual(namespace["text_area"].calls[0][0], "insert")
        self.assertIn(("insert", ("pending", "tag")), namespace["calls"])

    def test_log_namespace_runtime_runs_ui_and_file_path(self):
        namespace = self.make_namespace()

        result = log_namespace_runtime(namespace, "msg", "tag")

        self.assertEqual(result, "logged")
        self.assertIn(("run", "ui_post"), namespace["calls"])
        self.assertIn(("insert", ("msg", "tag")), namespace["calls"])
        self.assertEqual(len(namespace["FILE_LOG_Q"].items), 1)


if __name__ == "__main__":
    unittest.main()
