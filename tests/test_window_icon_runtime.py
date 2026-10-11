import unittest
import gc
from pathlib import Path
from types import SimpleNamespace
import tkinter as tk
import weakref

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
    def __init__(self):
        self.icon_calls = []
        self.bindings = {}

    def iconbitmap(self, path):
        self.icon_calls.append(path)

    def bind(self, event, callback, add=None):
        self.bindings[event] = callback
        return 'icon-map'

    def unbind(self, event, binding):
        self.bindings.pop(event, None)

    def map(self, widget=None):
        callback = self.bindings.get('<Map>')
        if callback:
            callback(SimpleNamespace(widget=self if widget is None else widget))


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
        self.assertEqual(window.icon_calls, [])
        window.map(widget=object())
        self.assertEqual(window.icon_calls, [])
        window.map()
        self.assertEqual(window.icon_calls, ["icon.ico"])
        window.map()
        self.assertEqual(window.icon_calls, ["icon.ico"])
        self.assertEqual(window.bindings, {})
        self.assertEqual(messagebox.showinfo("title", "message"), "info")
        self.assertIs(messagebox.calls[-1][3]["parent"], root)
        self.assertEqual(messagebox.showwarning("title", "message", parent="custom"), "warning")
        self.assertEqual(messagebox.calls[-1][3]["parent"], "custom")
        self.assertEqual(logs, [])

    def test_install_window_icon_runtime_logs_root_failure_and_sets_child_icon(self):
        root = FakeRoot(fail_icon=True)
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
        tk_module.Toplevel()
        window.map()
        self.assertEqual(window.icon_calls, ["icon.ico"])
        self.assertTrue(any("root icon failed" in message for message in logs))


class WindowIconMappingTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest('Tk display is unavailable')
        self.root.withdraw()
        self.root.attributes('-alpha', 0)
        self.addCleanup(self.root.destroy)
        self.root.deiconify()
        self.root.update()
        self.module = SimpleNamespace(Toplevel=tk.Toplevel)
        self.logs = []
        self.errors = []
        self.root.report_callback_exception = lambda *args, errors=self.errors: errors.append(args)
        self.assertTrue(install_window_icon_runtime(
            self.root, self.module, FakeMessageBox(),
            icon_path=str(Path(__file__).resolve().parents[1] / 'icon.ico'),
            path_exists=lambda path: Path(path).is_file(), log_error=self.logs.append))

    def test_constructor_does_not_map_unconfigured_window(self):
        window = self.module.Toplevel(self.root)
        self.assertFalse(window.winfo_ismapped(), 'Icon setup displayed an unconfigured 200x200 window')
        window.withdraw()
        window.title('Prepared icon regression dialog')
        window.geometry('360x180+400+300')
        icon_calls = []
        original_iconbitmap = window.iconbitmap

        def record_icon(*args, **kwargs):
            icon_calls.append((window.winfo_ismapped(), window.winfo_width(), window.winfo_height()))
            return original_iconbitmap(*args, **kwargs)

        window.iconbitmap = record_icon
        self.addCleanup(delattr, window, 'iconbitmap')
        window.update_idletasks()
        self.assertEqual(icon_calls, [])
        window.deiconify()
        self.root.update()
        self.assertEqual(icon_calls, [(1, 360, 180)])
        self.assertFalse(window.bind('<Map>'))
        window.withdraw()
        window.deiconify()
        self.root.update()
        self.assertEqual(len(icon_calls), 1)
        self.assertEqual(self.errors, [])

    def test_destroy_before_mapping_releases_callback_without_timer(self):
        timers = self.root.tk.call('after', 'info')
        window = self.module.Toplevel(self.root)
        reference = weakref.ref(window)
        window.destroy()
        del window
        gc.collect()
        self.root.update()
        self.assertIsNone(reference())
        self.assertEqual(self.root.tk.call('after', 'info'), timers)
        self.assertEqual(self.errors, [])


if __name__ == "__main__":
    unittest.main()
