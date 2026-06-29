import unittest

from sms_ui.main_window_layout import (
    build_main_window_layout_runtime,
    main_text_readonly_key_handler,
)


class FakeWidget:
    def __init__(self, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs
        self.grid_calls = []
        self.pack_calls = []
        self.bind_calls = []
        self.tag_calls = []
        self.mark_calls = []
        self.see_calls = []

    def grid(self, **kwargs):
        self.grid_calls.append(kwargs)

    def pack(self, **kwargs):
        self.pack_calls.append(kwargs)

    def bind(self, sequence, callback):
        self.bind_calls.append((sequence, callback))

    def tag_add(self, tag, start, end):
        self.tag_calls.append((tag, start, end))

    def mark_set(self, mark, index):
        self.mark_calls.append((mark, index))

    def see(self, index):
        self.see_calls.append(index)


class FakeVar:
    def __init__(self, value=""):
        self.value = value

    def get(self):
        return self.value


class FakeRoot:
    def __init__(self):
        self.rows = []
        self.columns = []

    def grid_rowconfigure(self, row, weight):
        self.rows.append((row, weight))

    def grid_columnconfigure(self, column, weight):
        self.columns.append((column, weight))


class FakeTkModule:
    BOTH = "both"
    LEFT = "left"
    Frame = FakeWidget
    Label = FakeWidget

    @staticmethod
    def StringVar(value=""):
        return FakeVar(value)


class FakeEvent:
    def __init__(self, *, keysym="", char="", state=0):
        self.keysym = keysym
        self.char = char
        self.state = state


class MainWindowLayoutTests(unittest.TestCase):
    def test_build_main_window_layout_runtime_returns_widget_references(self):
        root = FakeRoot()

        refs = build_main_window_layout_runtime(
            root,
            FakeTkModule,
            cloud_enabled=True,
            scrolled_text_class=FakeWidget,
        )

        self.assertEqual(root.rows, [(0, 1), (1, 0)])
        self.assertEqual(root.columns, [(0, 1)])
        self.assertEqual(refs["status_var"].get(), "🔍 启动中…")
        self.assertEqual(refs["temp_var"].get(), "🌡️ -- ℃")
        self.assertEqual(refs["signal_var"].get(), "📶 -- dBm")
        self.assertEqual(refs["cloud_var"].get(), "🌐 等待连接")
        self.assertIn("text_area", refs)
        self.assertIn("cloud_label", refs)
        self.assertEqual(
            [sequence for sequence, _callback in refs["text_area"].bind_calls],
            ["<Key>", "<<Paste>>", "<<Cut>>", "<<Clear>>", "<<PasteSelection>>", "<Button-2>"],
        )

    def test_build_main_window_layout_runtime_uses_closed_cloud_text(self):
        refs = build_main_window_layout_runtime(
            FakeRoot(),
            FakeTkModule,
            cloud_enabled=False,
            scrolled_text_class=FakeWidget,
        )

        self.assertEqual(refs["cloud_var"].get(), "🌐 已关闭")

    def test_main_text_readonly_key_handler_allows_copy_and_blocks_edits(self):
        text = FakeWidget()

        self.assertIsNone(main_text_readonly_key_handler(FakeEvent(keysym="c", state=0x0004), text))
        self.assertEqual(main_text_readonly_key_handler(FakeEvent(keysym="v", state=0x0004), text), "break")
        self.assertEqual(main_text_readonly_key_handler(FakeEvent(keysym="x", state=0x0004), text), "break")
        self.assertEqual(main_text_readonly_key_handler(FakeEvent(keysym="Delete"), text), "break")
        self.assertEqual(main_text_readonly_key_handler(FakeEvent(keysym="a", char="a"), text), "break")
        self.assertIsNone(main_text_readonly_key_handler(FakeEvent(keysym="Down"), text))

    def test_main_text_readonly_key_handler_supports_select_all_for_copy(self):
        text = FakeWidget()

        self.assertEqual(main_text_readonly_key_handler(FakeEvent(keysym="a", state=0x0004), text), "break")

        self.assertEqual(text.tag_calls, [("sel", "1.0", "end-1c")])
        self.assertEqual(text.mark_calls, [("insert", "1.0")])
        self.assertEqual(text.see_calls, ["insert"])


if __name__ == "__main__":
    unittest.main()
