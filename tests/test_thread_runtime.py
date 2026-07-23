import unittest
import queue

from sms_ui.thread_runtime import (
    run_on_ui_thread,
    post_ui_if_running_runtime,
    schedule_delayed_ui_runtime,
    tk_alive_runtime,
    ui_messagebox_runtime,
    ui_post_runtime,
    ui_pump_runtime,
)


class RootStub:
    def __init__(self, exists=True):
        self.exists = exists
        self.after_calls = []

    def winfo_exists(self):
        return self.exists

    def after(self, delay, callback):
        self.after_calls.append((delay, callback))


class MessageBoxStub:
    def __init__(self):
        self.calls = []

    def showinfo(self, title, message):
        self.calls.append(("info", title, message))
        return "info-result"

    def showwarning(self, title, message):
        self.calls.append(("warning", title, message))
        return "warning-result"

    def showerror(self, title, message):
        self.calls.append(("error", title, message))
        return "error-result"

    def askyesno(self, title, message):
        self.calls.append(("askyesno", title, message))
        return True


class ThreadRuntimeTests(unittest.TestCase):
    def test_run_on_ui_thread_executes_immediately_on_main_thread(self):
        calls = []
        main = object()

        result = run_on_ui_thread(
            lambda: calls.append("ran") or "result",
            lambda callback: calls.append(("post", callback)),
            current_thread=lambda: main,
            main_thread=lambda: main,
        )

        self.assertEqual(result, "result")
        self.assertEqual(calls, ["ran"])

    def test_run_on_ui_thread_posts_from_background_thread(self):
        calls = []
        main = object()
        worker = object()

        result = run_on_ui_thread(
            lambda: calls.append("ran") or "result",
            lambda callback: calls.append(("post", callback)),
            current_thread=lambda: worker,
            main_thread=lambda: main,
        )

        self.assertIsNone(result)
        self.assertEqual(calls[0][0], "post")
        calls[0][1]()
        self.assertEqual(calls[1], "ran")

    def test_tk_alive_runtime_checks_root_on_main_thread(self):
        shutdown = type("Event", (), {"is_set": lambda _self: False})()
        main = object()

        result = tk_alive_runtime(
            RootStub(exists=True),
            shutdown,
            current_thread=lambda: main,
            main_thread=lambda: main,
        )

        self.assertTrue(result)

    def test_tk_alive_runtime_skips_winfo_from_background_thread(self):
        shutdown = type("Event", (), {"is_set": lambda _self: False})()
        main = object()
        worker = object()

        result = tk_alive_runtime(
            RootStub(exists=False),
            shutdown,
            current_thread=lambda: worker,
            main_thread=lambda: main,
        )

        self.assertTrue(result)

    def test_tk_alive_runtime_returns_false_when_shutdown(self):
        shutdown = type("Event", (), {"is_set": lambda _self: True})()

        self.assertFalse(tk_alive_runtime(RootStub(), shutdown))

    def test_ui_post_runtime_queues_callback_with_arguments(self):
        task_queue = queue.Queue()

        result = ui_post_runtime(task_queue, lambda: None, (1,), {"name": "value"})

        self.assertTrue(result)
        _callback, args, kwargs = task_queue.get_nowait()
        self.assertEqual(args, (1,))
        self.assertEqual(kwargs, {"name": "value"})

    def test_ui_post_runtime_reports_full_queue(self):
        calls = []
        task_queue = queue.Queue(maxsize=1)
        task_queue.put_nowait(("existing", (), {}))

        result = ui_post_runtime(task_queue, lambda: None, on_full=lambda: calls.append("full"))

        self.assertFalse(result)
        self.assertEqual(calls, ["full"])

    def test_ui_post_runtime_logs_unexpected_queue_error(self):
        logs = []

        class BrokenQueue:
            def put_nowait(self, _item):
                raise RuntimeError("queue closed")

        result = ui_post_runtime(BrokenQueue(), lambda: None, log_error=logs.append)

        self.assertFalse(result)
        self.assertEqual(len(logs), 1)
        self.assertIn("queue closed", logs[0])

    def test_post_ui_if_running_reports_enqueue_failure_and_runs_skip_callback(self):
        calls = []

        result = post_ui_if_running_runtime(
            lambda _callback: False,
            lambda: calls.append("ran"),
            lambda: False,
            on_skipped=lambda: calls.append("skipped"),
        )

        self.assertFalse(result)
        self.assertEqual(calls, ["skipped"])

    def test_post_ui_if_running_checks_shutdown_before_queue_and_execution(self):
        state = {"stopping": False}
        posted = []
        calls = []
        skipped = []

        self.assertTrue(
            post_ui_if_running_runtime(
                posted.append,
                lambda: calls.append("ran"),
                lambda: state["stopping"],
                on_skipped=lambda: skipped.append("late"),
            )
        )
        state["stopping"] = True
        posted[0]()
        self.assertEqual(calls, [])
        self.assertEqual(skipped, ["late"])

        self.assertFalse(
            post_ui_if_running_runtime(
                posted.append,
                lambda: calls.append("late"),
                lambda: True,
                on_skipped=lambda: skipped.append("immediate"),
            )
        )
        self.assertEqual(len(posted), 1)
        self.assertEqual(skipped, ["late", "immediate"])

    def test_ui_pump_runtime_runs_tasks_and_reschedules(self):
        calls = []
        task_queue = queue.Queue()
        root = RootStub()
        task_queue.put_nowait((lambda value: calls.append(value), ("ran",), {}))

        processed = ui_pump_runtime(
            task_queue,
            root,
            tk_alive=lambda: True,
            schedule_self=lambda: calls.append("next"),
            max_batch=5,
        )

        self.assertEqual(processed, 1)
        self.assertEqual(calls, ["ran"])
        self.assertEqual(root.after_calls[0][0], 30)
        self.assertEqual(task_queue.unfinished_tasks, 0)

    def test_ui_pump_runtime_swallows_task_errors(self):
        calls = []
        task_queue = queue.Queue()
        task_queue.put_nowait((lambda: (_ for _ in ()).throw(RuntimeError("boom")), (), {}))
        task_queue.put_nowait((lambda: calls.append("next"), (), {}))

        processed = ui_pump_runtime(
            task_queue,
            RootStub(),
            tk_alive=lambda: False,
            schedule_self=lambda: None,
            max_batch=5,
        )

        self.assertEqual(processed, 2)
        self.assertEqual(calls, ["next"])
        self.assertEqual(task_queue.unfinished_tasks, 0)

    def test_ui_pump_runtime_logs_task_errors(self):
        logs = []
        task_queue = queue.Queue()
        task_queue.put_nowait((lambda: (_ for _ in ()).throw(RuntimeError("boom")), (), {}))

        processed = ui_pump_runtime(
            task_queue,
            RootStub(),
            tk_alive=lambda: False,
            schedule_self=lambda: None,
            log_error=logs.append,
        )

        self.assertEqual(processed, 1)
        self.assertEqual(len(logs), 1)
        self.assertIn("boom", logs[0])
        self.assertEqual(task_queue.unfinished_tasks, 0)

    def test_schedule_delayed_ui_runtime_schedules_remaining_startup_delay(self):
        calls = []

        result = schedule_delayed_ui_runtime(
            lambda: calls.append("callback"),
            app_start_mono=10.0,
            start_ui_delay=2.0,
            monotonic=lambda: 11.25,
            root_after=lambda delay, callback: calls.append(("after", delay, callback)) or "after-id",
            run_on_ui_thread=lambda callback, ui_post: calls.append(("run", ui_post)) or callback(),
            ui_post="post",
        )

        self.assertEqual(result, "after-id")
        self.assertEqual(calls[0], ("run", "post"))
        self.assertEqual(calls[1][0], "after")
        self.assertEqual(calls[1][1], 750)
        calls[1][2]()
        self.assertEqual(calls[2], "callback")

    def test_schedule_delayed_ui_runtime_runs_immediately_after_delay_elapsed(self):
        calls = []

        result = schedule_delayed_ui_runtime(
            lambda: calls.append("callback") or "done",
            app_start_mono=10.0,
            start_ui_delay=2.0,
            monotonic=lambda: 15.0,
            root_after=lambda delay, callback: calls.append(("after", delay)),
            run_on_ui_thread=lambda callback, ui_post: callback(),
            ui_post=None,
        )

        self.assertEqual(result, "done")
        self.assertEqual(calls, ["callback"])

    def test_schedule_delayed_ui_runtime_falls_back_when_after_fails(self):
        calls = []
        logs = []

        result = schedule_delayed_ui_runtime(
            lambda: calls.append("callback") or "done",
            app_start_mono=10.0,
            start_ui_delay=2.0,
            monotonic=lambda: 10.5,
            root_after=lambda delay, callback: (_ for _ in ()).throw(RuntimeError("dead tk")),
            run_on_ui_thread=lambda callback, ui_post: callback(),
            ui_post=None,
            log_error=logs.append,
        )

        self.assertEqual(result, "done")
        self.assertEqual(calls, ["callback"])
        self.assertTrue(any("dead tk" in message for message in logs))

    def test_schedule_delayed_ui_runtime_logs_callback_failure_once(self):
        calls = []
        logs = []

        def callback():
            calls.append("callback")
            raise RuntimeError("callback failed")

        result = schedule_delayed_ui_runtime(
            callback,
            app_start_mono=10.0,
            start_ui_delay=2.0,
            monotonic=lambda: 15.0,
            root_after=lambda delay, callback: None,
            run_on_ui_thread=lambda callback, ui_post: callback(),
            ui_post=None,
            log_error=logs.append,
        )

        self.assertIsNone(result)
        self.assertEqual(calls, ["callback"])
        self.assertTrue(any("callback failed" in message for message in logs))

    def test_schedule_delayed_ui_runtime_logs_scheduled_callback_failure(self):
        calls = []
        logs = []

        def callback():
            calls.append("callback")
            raise RuntimeError("scheduled callback failed")

        result = schedule_delayed_ui_runtime(
            callback,
            app_start_mono=10.0,
            start_ui_delay=2.0,
            monotonic=lambda: 11.0,
            root_after=lambda delay, scheduled: calls.append(("after", delay, scheduled)) or "after-id",
            run_on_ui_thread=lambda callback, ui_post: callback(),
            ui_post=None,
            log_error=logs.append,
        )

        self.assertEqual(result, "after-id")
        calls[0][2]()
        self.assertEqual(calls[:2], [("after", 1000, calls[0][2]), "callback"])
        self.assertTrue(any("scheduled callback failed" in message for message in logs))

    def test_schedule_delayed_ui_runtime_logs_time_source_failure(self):
        calls = []
        logs = []

        result = schedule_delayed_ui_runtime(
            lambda: calls.append("callback") or "done",
            app_start_mono=10.0,
            start_ui_delay=2.0,
            monotonic=lambda: (_ for _ in ()).throw(RuntimeError("clock failed")),
            root_after=lambda delay, callback: None,
            run_on_ui_thread=lambda callback, ui_post: callback(),
            ui_post=None,
            log_error=logs.append,
        )

        self.assertEqual(result, "done")
        self.assertEqual(calls, ["callback"])
        self.assertTrue(any("clock failed" in message for message in logs))

    def test_ui_messagebox_runtime_dispatches_known_kinds_on_ui_thread(self):
        messagebox = MessageBoxStub()
        calls = []

        self.assertEqual(
            ui_messagebox_runtime(
                "info",
                "Title",
                "Message",
                messagebox=messagebox,
                run_on_ui_thread=lambda callback, ui_post: calls.append(("run", ui_post)) or callback(),
                ui_post="post",
            ),
            "info-result",
        )
        self.assertEqual(
            ui_messagebox_runtime(
                "warning",
                "Title",
                "Message",
                messagebox=messagebox,
                run_on_ui_thread=lambda callback, ui_post: callback(),
                ui_post=None,
            ),
            "warning-result",
        )
        self.assertEqual(
            ui_messagebox_runtime(
                "error",
                "Title",
                "Message",
                messagebox=messagebox,
                run_on_ui_thread=lambda callback, ui_post: callback(),
                ui_post=None,
            ),
            "error-result",
        )
        self.assertTrue(
            ui_messagebox_runtime(
                "askyesno",
                "Title",
                "Message",
                messagebox=messagebox,
                run_on_ui_thread=lambda callback, ui_post: callback(),
                ui_post=None,
            )
        )

        self.assertEqual(calls, [("run", "post")])
        self.assertEqual([call[0] for call in messagebox.calls], ["info", "warning", "error", "askyesno"])

    def test_ui_messagebox_runtime_ignores_unknown_kind(self):
        messagebox = MessageBoxStub()

        result = ui_messagebox_runtime(
            "custom",
            "Title",
            "Message",
            messagebox=messagebox,
            run_on_ui_thread=lambda callback, ui_post: callback(),
            ui_post=None,
        )

        self.assertIsNone(result)
        self.assertEqual(messagebox.calls, [])


if __name__ == "__main__":
    unittest.main()
