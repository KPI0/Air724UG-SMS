import datetime
import os
import queue
import unittest

from sms_ui.ui_log_runtime import (
    clear_text_widget_runtime,
    flush_pending_ui_logs_runtime,
    insert_main_text_runtime,
    log_file_only_runtime,
    run_log_runtime,
    schedule_next_midnight_clear_runtime,
    show_sms_popup_runtime,
    system_ui_runtime,
    ui_only_runtime,
    write_port_log_runtime,
    write_system_log_runtime,
)


class FakeText:
    def __init__(self, exists=True, bottom=1.0, lines=1):
        self.exists = exists
        self.bottom = bottom
        self.lines = lines
        self.calls = []

    def winfo_exists(self):
        return self.exists

    def yview(self):
        return (0.0, self.bottom)

    def insert(self, end, text, tag):
        self.calls.append(("insert", end, text, tag))

    def index(self, index):
        return f"{self.lines}.0"

    def delete(self, start, end):
        self.calls.append(("delete", start, end))

    def see(self, end):
        self.calls.append(("see", end))


class PendingQueue:
    def __init__(self):
        self.items = []

    def put_nowait(self, item):
        self.items.append(item)


class UiLogRuntimeTests(unittest.TestCase):
    def test_insert_main_text_runtime_inserts_trims_and_scrolls_when_at_bottom(self):
        text = FakeText(bottom=0.99, lines=5)

        self.assertTrue(insert_main_text_runtime(text, "hello", "tag", max_lines=3, end_marker="END"))

        self.assertEqual(text.calls, [
            ("insert", "END", "hello\n", "tag"),
            ("delete", "1.0", "3.0"),
            ("see", "END"),
        ])

    def test_insert_main_text_runtime_does_not_scroll_when_not_at_bottom(self):
        text = FakeText(bottom=0.5, lines=1)

        self.assertTrue(insert_main_text_runtime(text, "hello", "tag", end_marker="END"))

        self.assertEqual(text.calls, [("insert", "END", "hello\n", "tag")])

    def test_insert_main_text_runtime_rejects_missing_widget(self):
        self.assertFalse(insert_main_text_runtime(None, "hello"))
        self.assertFalse(insert_main_text_runtime(FakeText(exists=False), "hello"))

    def test_write_port_log_runtime_formats_path_and_line(self):
        written = []
        now = datetime.datetime(2026, 6, 8, 12, 30, 5)

        path, line = write_port_log_runtime("msg", "logs", "COM5", written.append, now=now)

        self.assertEqual(path, os.path.join("logs", "sms_COM5_2026-06-08.txt"))
        self.assertEqual(line, "2026-06-08 12:30:05 msg\n")
        self.assertEqual(written, [(path, line)])

    def test_write_system_log_runtime_formats_path_and_line(self):
        written = []
        now = datetime.datetime(2026, 6, 8, 12, 30, 5)

        path, line = write_system_log_runtime("msg", "logs", written.append, now=now)

        self.assertEqual(path, os.path.join("logs", "sms_system_2026-06-08.txt"))
        self.assertEqual(line, "2026-06-08 12:30:05 msg\n")
        self.assertEqual(written, [(path, line)])

    def test_log_file_only_runtime_swallows_queue_errors(self):
        result = log_file_only_runtime(
            "msg",
            log_dir="logs",
            file_log=lambda item: (_ for _ in ()).throw(RuntimeError("full")),
            now=datetime.datetime(2026, 6, 8, 12, 30, 5),
        )

        self.assertIsNone(result)

    def test_flush_pending_ui_logs_runtime_drains_queue(self):
        pending = queue.Queue()
        pending.put_nowait(("one", "normal"))
        pending.put_nowait(("two", "warn"))
        calls = []

        flushed = flush_pending_ui_logs_runtime(pending, lambda *args: calls.append(args))

        self.assertEqual(flushed, 2)
        self.assertTrue(pending.empty())
        self.assertEqual(calls, [("one", "normal"), ("two", "warn")])

    def test_flush_pending_ui_logs_runtime_continues_after_insert_error(self):
        pending = queue.Queue()
        pending.put_nowait(("one", "normal"))
        pending.put_nowait(("two", "normal"))
        calls = []

        def insert(message, tag):
            calls.append((message, tag))
            if message == "one":
                raise RuntimeError("boom")

        flushed = flush_pending_ui_logs_runtime(pending, insert)

        self.assertEqual(flushed, 2)
        self.assertEqual(calls, [("one", "normal"), ("two", "normal")])

    def test_show_sms_popup_runtime_shows_when_enabled(self):
        calls = []

        result = show_sms_popup_runtime(
            "hello",
            popup_enabled=True,
            show_info=lambda *args: calls.append(("info", args)),
            show_window=lambda: calls.append(("show",)),
        )

        self.assertEqual(result, "shown")
        self.assertEqual(calls, [("info", ("短信提醒", "hello")), ("show",)])

    def test_show_sms_popup_runtime_skips_when_disabled(self):
        result = show_sms_popup_runtime(
            "hello",
            popup_enabled=False,
            show_info=lambda *_args: None,
            show_window=lambda: None,
        )

        self.assertEqual(result, "disabled")

    def test_clear_text_widget_runtime_deletes_existing_text(self):
        text = FakeText()

        self.assertTrue(clear_text_widget_runtime(text, end="END"))
        self.assertEqual(text.calls, [("delete", "1.0", "END")])

    def test_clear_text_widget_runtime_rejects_missing_widget(self):
        self.assertFalse(clear_text_widget_runtime(None))
        self.assertFalse(clear_text_widget_runtime(FakeText(exists=False)))

    def test_schedule_next_midnight_clear_runtime_schedules_delay(self):
        calls = []
        now = datetime.datetime(2026, 6, 8, 23, 59, 0)

        delay = schedule_next_midnight_clear_runtime(
            tk_alive=lambda: True,
            schedule_after=lambda ms, callback: calls.append((ms, callback)),
            clear_callback=lambda: None,
            now_func=lambda: now,
        )

        self.assertEqual(delay, 60000)
        self.assertEqual(calls[0][0], 60000)

    def test_schedule_next_midnight_clear_runtime_deduplicates_replaced_timer(self):
        scheduled = []
        cancelled = []
        cleared = []
        state = {}
        now = [datetime.datetime(2026, 6, 8, 23, 59, 0)]

        def schedule_after(delay_ms, callback):
            item = (len(scheduled) + 1, delay_ms, callback)
            scheduled.append(item)
            return item[0]

        schedule_next_midnight_clear_runtime(
            tk_alive=lambda: True,
            schedule_after=schedule_after,
            cancel_after=cancelled.append,
            clear_callback=lambda: cleared.append("clear"),
            now_func=lambda: now[0],
            state=state,
        )
        schedule_next_midnight_clear_runtime(
            tk_alive=lambda: True,
            schedule_after=schedule_after,
            cancel_after=cancelled.append,
            clear_callback=lambda: cleared.append("clear"),
            now_func=lambda: now[0],
            state=state,
        )

        scheduled[0][2]()
        self.assertEqual(cancelled, [1])
        self.assertEqual(cleared, [])

        now[0] = datetime.datetime(2026, 6, 9, 0, 0, 1)
        scheduled[1][2]()
        self.assertEqual(cleared, ["clear"])

        scheduled[1][2]()
        self.assertEqual(cleared, ["clear"])

        now[0] = datetime.datetime(2026, 6, 10, 0, 0, 1)
        scheduled[2][2]()
        self.assertEqual(cleared, ["clear", "clear"])

    def test_schedule_next_midnight_clear_runtime_skips_when_tk_dead(self):
        delay = schedule_next_midnight_clear_runtime(
            tk_alive=lambda: False,
            schedule_after=lambda *_args: None,
            clear_callback=lambda: None,
            now_func=lambda: datetime.datetime(2026, 6, 8, 23, 59, 0),
        )

        self.assertIsNone(delay)

    def test_run_log_runtime_uses_early_log_when_text_missing(self):
        calls = []

        result = run_log_runtime(
            "msg",
            "tag",
            has_text=lambda: False,
            insert_text=lambda *_: calls.append(("insert",)),
            log_early=lambda *args: calls.append(("early", args)),
            log_dir="logs",
            log_prefix="COM5",
            file_log=lambda item: calls.append(("file", item)),
        )

        self.assertEqual(result, "early")
        self.assertEqual(calls, [("early", ("msg", "tag"))])

    def test_run_log_runtime_inserts_and_writes_file(self):
        calls = []
        now = datetime.datetime(2026, 6, 8, 12, 30, 5)

        result = run_log_runtime(
            "msg",
            "tag",
            has_text=lambda: True,
            insert_text=lambda *args: calls.append(("insert", args)),
            log_early=lambda *args: calls.append(("early", args)),
            log_dir="logs",
            log_prefix="COM5",
            file_log=lambda item: calls.append(("file", item)),
            now=now,
        )

        self.assertEqual(result, "logged")
        self.assertEqual(calls[0], ("insert", ("msg", "tag")))
        self.assertEqual(calls[1][0], "file")

    def test_system_ui_runtime_caches_when_tk_is_not_alive(self):
        pending = PendingQueue()
        calls = []

        result = system_ui_runtime(
            "msg",
            "tag",
            tk_alive=lambda: False,
            log_file_only=lambda message: calls.append(("file", message)),
            pending_logs=pending,
            has_text=lambda: True,
            insert_text=lambda *_: calls.append(("insert",)),
            schedule_ui=lambda callback: calls.append(("schedule", callback)),
        )

        self.assertEqual(result, "pending")
        self.assertEqual(calls, [("file", "msg")])
        self.assertEqual(pending.items, [("msg", "tag")])

    def test_system_ui_runtime_schedules_insert_when_text_exists(self):
        pending = PendingQueue()
        calls = []

        result = system_ui_runtime(
            "msg",
            "tag",
            tk_alive=lambda: True,
            log_file_only=lambda message: calls.append(("file", message)),
            pending_logs=pending,
            has_text=lambda: True,
            insert_text=lambda *args: calls.append(("insert", args)),
            schedule_ui=lambda callback: callback(),
        )

        self.assertEqual(result, "scheduled")
        self.assertEqual(calls, [("file", "msg"), ("insert", ("msg", "tag"))])
        self.assertEqual(pending.items, [])

    def test_ui_only_runtime_inserts_when_text_exists(self):
        pending = PendingQueue()
        calls = []

        result = ui_only_runtime(
            "msg",
            "tag",
            pending_logs=pending,
            has_text=lambda: True,
            insert_text=lambda *args: calls.append(("insert", args)),
            run_on_ui_thread=lambda callback, ui_post: calls.append(("run", ui_post)) or callback(),
            ui_post="post",
        )

        self.assertEqual(result, "inserted")
        self.assertEqual(calls, [("run", "post"), ("insert", ("msg", "tag"))])
        self.assertEqual(pending.items, [])

    def test_ui_only_runtime_caches_when_text_missing_or_insert_fails(self):
        pending = PendingQueue()

        self.assertEqual(
            ui_only_runtime(
                "missing",
                "normal",
                pending_logs=pending,
                has_text=lambda: False,
                insert_text=lambda *_: None,
                run_on_ui_thread=lambda callback, ui_post: callback(),
                ui_post=None,
            ),
            "pending",
        )
        self.assertEqual(
            ui_only_runtime(
                "boom",
                "warn",
                pending_logs=pending,
                has_text=lambda: True,
                insert_text=lambda *_: (_ for _ in ()).throw(RuntimeError("dead text")),
                run_on_ui_thread=lambda callback, ui_post: callback(),
                ui_post=None,
            ),
            "pending",
        )

        self.assertEqual(pending.items, [("missing", "normal"), ("boom", "warn")])


if __name__ == "__main__":
    unittest.main()
