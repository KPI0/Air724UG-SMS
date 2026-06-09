import unittest

from sms_ui.third_push_result_runtime import (
    show_third_push_test_result_runtime,
    third_push_result_dialog_plan,
    third_push_result_parent,
)


class FakeWindow:
    def __init__(self, exists=True, fail=False):
        self.exists = exists
        self.fail = fail

    def winfo_exists(self):
        if self.fail:
            raise RuntimeError("window failed")
        return self.exists


class FakeMessageBox:
    def __init__(self):
        self.calls = []

    def showwarning(self, title, message, parent=None):
        self.calls.append(("warning", title, message, parent))

    def showerror(self, title, message, parent=None):
        self.calls.append(("error", title, message, parent))

    def showinfo(self, title, message, parent=None):
        self.calls.append(("info", title, message, parent))


class ThirdPushResultRuntimeTests(unittest.TestCase):
    def test_third_push_result_parent_prefers_existing_window(self):
        root = object()
        win = FakeWindow(exists=True)

        self.assertIs(third_push_result_parent(root, win), win)
        self.assertIs(third_push_result_parent(root, FakeWindow(exists=False)), root)
        self.assertIs(third_push_result_parent(root, FakeWindow(fail=True)), root)

    def test_third_push_result_dialog_plan_formats_cases(self):
        self.assertEqual(
            third_push_result_dialog_plan(["dingtalk"], ["bark: fail"])[0:2],
            ("warning", "测试部分成功"),
        )
        self.assertEqual(
            third_push_result_dialog_plan([], ["bark: fail"])[0:2],
            ("error", "测试推送失败"),
        )
        self.assertEqual(
            third_push_result_dialog_plan(["dingtalk"], [])[0:2],
            ("info", "测试推送成功"),
        )
        self.assertEqual(
            third_push_result_dialog_plan([], [])[0:2],
            ("warning", "测试推送失败"),
        )

    def test_show_third_push_test_result_runtime_posts_dialog(self):
        posted = []
        messagebox = FakeMessageBox()
        root = object()
        win = FakeWindow()

        show_third_push_test_result_runtime(
            root=root,
            current_window=win,
            messagebox=messagebox,
            ui_post=lambda fn: posted.append(fn),
            ok_channels=["dingtalk"],
            fail_infos=[],
        )

        self.assertEqual(len(posted), 1)
        posted[0]()
        self.assertEqual(messagebox.calls[0][0], "info")
        self.assertEqual(messagebox.calls[0][3], win)


if __name__ == "__main__":
    unittest.main()
