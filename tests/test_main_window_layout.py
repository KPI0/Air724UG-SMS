import unittest

from sms_ui.main_window_layout import build_main_window_layout_runtime


class FakeWidget:
    def __init__(self, *args, **kwargs):
        self.args = args
        self.kwargs = kwargs
        self.grid_calls = []
        self.pack_calls = []

    def grid(self, **kwargs):
        self.grid_calls.append(kwargs)

    def pack(self, **kwargs):
        self.pack_calls.append(kwargs)


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

    def test_build_main_window_layout_runtime_uses_closed_cloud_text(self):
        refs = build_main_window_layout_runtime(
            FakeRoot(),
            FakeTkModule,
            cloud_enabled=False,
            scrolled_text_class=FakeWidget,
        )

        self.assertEqual(refs["cloud_var"].get(), "🌐 已关闭")


if __name__ == "__main__":
    unittest.main()
