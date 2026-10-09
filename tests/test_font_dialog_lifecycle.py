"""Font dialog resources must be released when a long-running client closes it."""
import gc
import tkinter as tk
import unittest
import weakref

from sms_ui.sms_font_dialog import open_sms_font_dialog


class FontDialogLifecycleTests(unittest.TestCase):
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
