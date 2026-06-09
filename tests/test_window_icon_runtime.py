import unittest

from sms_ui.window_icon_runtime import install_window_icon_runtime


class FakeRoot:
    def __init__(self, fail_icon=False):
        self.fail_icon = fail_icon
        self.icon_calls = []

    def iconbitmap(self, path):
        self.icon_calls.append(path)
        if self.fail_icon:
            raise RuntimeError("root icon failed")


class FakeWindow:
    def __init__(self, fail_after=False):
        self.fail_after = fail_after
        self.icon_calls = []
        self.after_calls = []

    def iconbitmap(self, path):
        self.icon_calls.append(path)

    def after(self, delay, callback):
        self.after_calls.append(delay)
        if self.fail_after:
            raise RuntimeError("after failed")
        callback()


class FakeTkModule:
    def __init__(self, window):
        self.window = window
        self.Toplevel = self.create_toplevel

    def create_toplevel(self, *args, **kwargs):
        return self.window


class FakeMessageBox:
    def __init__(self):
        self.calls = []

    def showinfo(self, title, message, **options):
        self.calls.append(("info", title, message, options))
        return "info"

    def showwarning(self, title, message, **options):
        self.calls.append(("warning", title, message, options))
        return "warning"

    def showerror(self, title, message, **options):
        self.calls.append(("error", title, message, options))
        return "error"

    def askyesno(self, title, message, **options):
        self.calls.append(("ask", title, message, options))
        return True


class WindowIconRuntimeTests(unittest.TestCase):
    def test_install_window_icon_runtime_patches_toplevel_and_messagebox_parent(self):
        root = FakeRoot()
        window = FakeWindow()
        tk_module = FakeTkModule(window)
        messagebox = FakeMessageBox()
        logs = []

        result = install_window_icon_runtime(
            root,
            tk_module,
            messagebox,
            icon_path="icon.ico",
            path_exists=lambda path: True,
            log_error=logs.append,
        )

        self.assertTrue(result)
        self.assertEqual(root.icon_calls, ["icon.ico"])
        created = tk_module.Toplevel("parent")
        self.assertIs(created, window)
        self.assertEqual(window.icon_calls, ["icon.ico"])
        self.assertEqual(messagebox.showinfo("title", "message"), "info")
        self.assertIs(messagebox.calls[-1][3]["parent"], root)
        self.assertEqual(messagebox.showwarning("title", "message", parent="custom"), "warning")
        self.assertEqual(messagebox.calls[-1][3]["parent"], "custom")
        self.assertEqual(logs, [])

    def test_install_window_icon_runtime_logs_root_icon_failure_and_handles_after_failure(self):
        root = FakeRoot(fail_icon=True)
        window = FakeWindow(fail_after=True)
        tk_module = FakeTkModule(window)
        messagebox = FakeMessageBox()
        logs = []

        result = install_window_icon_runtime(
            root,
            tk_module,
            messagebox,
            icon_path="icon.ico",
            path_exists=lambda path: True,
            log_error=logs.append,
        )

        self.assertTrue(result)
        tk_module.Toplevel()
        self.assertEqual(window.icon_calls, ["icon.ico"])
        self.assertTrue(any("root icon failed" in message for message in logs))


if __name__ == "__main__":
    unittest.main()
