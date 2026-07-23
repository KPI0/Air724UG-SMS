import queue
import unittest

from sms_ui.serial_debug_panel import append_serial_debug_lines_once
from sms_ui.serial_debug_runtime import SerialDebugPauseController


class FakeVar:
    def __init__(self, value=False):
        self.value = value

    def get(self):
        return self.value

    def set(self, value):
        self.value = value


class FakeWidget:
    def __init__(self):
        self.calls = []

    def config(self, **kwargs):
        self.calls.append(("config", kwargs))

    def state(self, value):
        self.calls.append(("state", tuple(value)))

    def winfo_children(self):
        return []


class FakeText:
    def __init__(self, *, bottom=0.5, top_index="12.0"):
        self.bottom = bottom
        self.top_index = top_index
        self.calls = []
        self.line_count = 1

    def config(self, **kwargs):
        self.calls.append(("config", kwargs))

    def insert(self, index, text):
        self.calls.append(("insert", index, text))
        self.line_count += text.count("\n")

    def index(self, index):
        if index == "@0,0":
            return self.top_index
        return f"{self.line_count}.0"

    def yview(self, *args):
        if args:
            self.calls.append(("yview", args))
            self.top_index = args[0]
            return None
        return (0.0, self.bottom)

    def yview_moveto(self, value):
        self.calls.append(("yview_moveto", value))

    def delete(self, start, end):
        self.calls.append(("delete", start, end))

    def see(self, index):
        self.calls.append(("see", index))


class FakeFinder:
    term = ""

    def clear(self):
        pass

    def highlight_range(self, _start, _end):
        pass

    def find_all(self, _term):
        pass


class SerialDebugRuntimeTests(unittest.TestCase):
    def test_pause_banner_preserves_current_text_view(self):
        enabled_var = FakeVar(True)
        paused_var = FakeVar(False)
        text = FakeText(bottom=0.4, top_index="20.0")
        controller = SerialDebugPauseController(
            enabled_var,
            paused_var,
            FakeWidget(),
            FakeWidget(),
            FakeWidget(),
            FakeWidget(),
            FakeWidget(),
            FakeWidget(),
            FakeWidget(),
            text,
        )

        controller.set_paused(True)

        self.assertTrue(paused_var.get())
        self.assertIn(("yview", ("20.0",)), text.calls)
        self.assertFalse(any(call[0] == "see" for call in text.calls))

    def test_pause_banner_keeps_bottom_when_user_is_at_bottom(self):
        enabled_var = FakeVar(True)
        paused_var = FakeVar(False)
        text = FakeText(bottom=0.99, top_index="20.0")
        controller = SerialDebugPauseController(
            enabled_var,
            paused_var,
            FakeWidget(),
            FakeWidget(),
            FakeWidget(),
            FakeWidget(),
            FakeWidget(),
            FakeWidget(),
            FakeWidget(),
            text,
        )

        controller.set_paused(True)

        self.assertTrue(paused_var.get())
        self.assertIn(("see", "end"), text.calls)
        self.assertFalse(any(call[0] == "yview" for call in text.calls))

    def test_append_lines_does_not_scroll_when_user_is_not_at_bottom(self):
        serial_queue = queue.Queue()
        serial_queue.put_nowait("new line")
        text = FakeText(bottom=0.5)

        appended = append_serial_debug_lines_once(
            text,
            serial_queue,
            [],
            "",
            FakeFinder(),
            max_store_lines=100,
            max_visible_lines=100,
        )

        self.assertTrue(appended)
        self.assertFalse(any(call[0] == "see" for call in text.calls))
        self.assertEqual(serial_queue.unfinished_tasks, 0)

    def test_append_lines_scrolls_when_user_is_at_bottom(self):
        serial_queue = queue.Queue()
        serial_queue.put_nowait("new line")
        text = FakeText(bottom=0.99)

        appended = append_serial_debug_lines_once(
            text,
            serial_queue,
            [],
            "",
            FakeFinder(),
            max_store_lines=100,
            max_visible_lines=100,
        )

        self.assertTrue(appended)
        self.assertIn(("see", "end"), text.calls)
        self.assertEqual(serial_queue.unfinished_tasks, 0)


if __name__ == "__main__":
    unittest.main()
