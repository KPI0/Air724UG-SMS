"""Font dialog resources must be released when a long-running client closes it."""
import gc
import tkinter as tk
import unittest
import weakref

from sms_ui.sms_font_dialog import open_sms_font_dialog


class FontDialogLifecycleTests(unittest.TestCase):
    def test_large_font_preview_fits_at_multiple_dpi_without_changing_saved_size(self):
        try:
            root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk display unavailable: {exc}")
        root.attributes("-alpha", 0)
        root.update()
        saved = []
        try:
            for scale in (4 / 3, 5 / 3, 8 / 3, 4):
                with self.subTest(scale=scale):
                    root.tk.call("tk", "scaling", scale)
                    open_sms_font_dialog(root, 72, "#ff0000", lambda size, color: saved.append((size, color)), lambda *_: None)
                    win = root.grab_current()
                    win.attributes("-alpha", 0)
                    root.update()

                    def descendants(parent):
                        for child in parent.winfo_children():
                            yield child
                            yield from descendants(child)

                    widgets = list(descendants(win))
                    canvas = next(widget for widget in widgets if isinstance(widget, tk.Canvas))
                    self.assertTrue(canvas.winfo_ismapped())
                    x1, y1, x2, y2 = canvas.bbox("all")
                    self.assertGreaterEqual(x1, 0)
                    self.assertGreaterEqual(y1, 0)
                    self.assertLessEqual(x2, canvas.winfo_width())
                    self.assertLessEqual(y2, canvas.winfo_height())
                    color_entry = next(widget for widget in widgets if isinstance(widget, tk.Entry) and not isinstance(widget, tk.Spinbox))
                    color_variable = tk.StringVar(root, name=color_entry.cget("textvariable"))
                    color_variable.set("not-a-color")
                    root.update()
                    x1, y1, x2, y2 = canvas.bbox("all")
                    self.assertGreaterEqual(x1, 0)
                    self.assertGreaterEqual(y1, 0)
                    self.assertLessEqual(x2, canvas.winfo_width())
                    self.assertLessEqual(y2, canvas.winfo_height())
                    color_variable.set("#ff0000")
                    next(widget for widget in widgets if isinstance(widget, tk.Button) and widget.cget("text") == "保存").invoke()
                    self.assertEqual(saved[-1], (72, "#ff0000"))
        finally:
            root.destroy()

    def test_repeated_close_releases_windows_and_tcl_variables(self):
        try:
            root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk display unavailable: {exc}")
        root.withdraw()
        references = []
        baseline = set(root.tk.call("info", "vars", "PY_VAR*"))
        try:
            for _ in range(3):
                open_sms_font_dialog(root, 30, "#ff0000", lambda *_: True, lambda *_: None)
                win = root.grab_current()
                references.append(weakref.ref(win))
                win.destroy()
                del win
                root.update()
            gc.collect()
            self.assertEqual(sum(ref() is not None for ref in references), 0)
            self.assertEqual(set(root.tk.call("info", "vars", "PY_VAR*")), baseline)
        finally:
            root.destroy()


if __name__ == "__main__":
    unittest.main()
