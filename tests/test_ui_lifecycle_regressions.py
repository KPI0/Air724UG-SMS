import gc
from pathlib import Path
import queue
from types import SimpleNamespace
import tkinter as tk
from tkinter import ttk
import unittest
from unittest.mock import patch
import weakref

from sms_ui import security_settings_dialog, sms_font_dialog
from sms_ui.serial_debug_panel import create_serial_debug_body, reset_serial_debug_window_state
from sms_ui.serial_debug_runtime import start_serial_debug_append_loop
from sms_ui.window_icon_runtime import install_window_icon_runtime


class UiLifecycleRegressionTests(unittest.TestCase):
    def setUp(self):
        security_settings_dialog._active_security_settings_refs = None
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("Tk display is unavailable")
        self.root.withdraw()
        self.addCleanup(self.root.destroy)
        self.errors = []
        self.root.report_callback_exception = lambda *args, errors=self.errors: errors.append(args)
        native = tk.Toplevel

        class InvisibleWindow(native):
            def __init__(window, *args, **kwargs):
                super().__init__(*args, **kwargs)
                window.attributes("-alpha", 0)

            def focus_force(window):
                pass

        self.patch = patch.object(tk, "Toplevel", InvisibleWindow)
        self.patch.start()
        self.addCleanup(self.patch.stop)
        self.addCleanup(setattr, security_settings_dialog, "_active_security_settings_refs", None)

    def tearDown(self):
        self.assertEqual(self.errors, [])

    def timers(self):
        return set(self.root.tk.call("after", "info"))

    def open_font(self):
        sms_font_dialog.open_sms_font_dialog(self.root, 30, "#ff0000", lambda *_: True, lambda *_: None)
        return next(w for w in self.root.winfo_children() if w.winfo_class() == "Toplevel")

    def test_icon_install_does_not_leave_callback_after_immediate_close(self):
        module = SimpleNamespace(Toplevel=tk.Toplevel)
        dialogs = SimpleNamespace(**{name: lambda *_args, **_kwargs: None
                                     for name in ("showinfo", "showwarning", "showerror", "askyesno")})
        logs = []
        self.assertTrue(install_window_icon_runtime(
            self.root, module, dialogs,
            icon_path=str(Path(__file__).resolve().parents[1] / "icon.ico"),
            path_exists=lambda p: Path(p).is_file(), log_error=logs.append))
        before = self.timers()
        win = module.Toplevel(self.root)
        win.destroy()
        self.assertEqual(self.timers(), before)
        self.assertEqual(logs, [])

    def test_font_immediate_close_cancels_preview(self):
        before = self.timers()
        self.open_font().destroy()
        self.assertEqual(self.timers(), before)

    def test_font_color_cancel_then_close_has_no_lift_callback(self):
        win = self.open_font()
        self.root.update()
        before = self.timers()
        with patch.object(sms_font_dialog.colorchooser, "askcolor", return_value=(None, None)):
            picker = next(widget for frame in win.winfo_children() for widget in frame.winfo_children()
                          if widget.winfo_class() == "Button" and widget.cget("text") == "选颜色")
            picker.invoke()
        win.destroy()
        self.assertEqual(self.timers(), before)

    def test_serial_refresh_cancels_timer_on_destroy_in_both_pause_states(self):
        for paused in (False, True):
            with self.subTest(paused=paused):
                win = tk.Toplevel(self.root)
                text = tk.Text(win)
                before = self.timers()
                start_serial_debug_append_loop(
                    win, lambda: text, queue.Queue(), [], tk.BooleanVar(win, value=paused),
                    ttk.Label(win), tk.StringVar(win, value=""), None, 100, 100, lambda: 0)
                self.assertNotEqual(self.timers(), before)
                win.destroy()
                self.assertEqual(self.timers(), before)

    def open_debug_body(self):
        win = tk.Toplevel(self.root)
        win.geometry("900x320")
        text, panel, frame = create_serial_debug_body(win, lambda _command: None)
        panel.grid(row=0, column=1, sticky="ns")
        self.root.update()
        return win, text, panel, frame

    def test_serial_hover_and_close_preserve_global_wheel_binding(self):
        marker = self.root.bind_all("<MouseWheel>", lambda _event: None)
        expected = self.root.bind_all("<MouseWheel>")
        win, text, panel, _frame = self.open_debug_body()
        canvas = next(w for w in panel.winfo_children() if w.winfo_class() == "Canvas")
        for _ in range(3):
            canvas.event_generate("<Enter>")
            canvas.event_generate("<Leave>")
        self.assertEqual(self.root.bind_all("<MouseWheel>"), expected)
        reset_serial_debug_window_state(win, text, queue.Queue(), [], SimpleNamespace(close=lambda: None))
        win.destroy()
        self.assertEqual(self.root.bind_all("<MouseWheel>"), expected)
        self.root.unbind_all("<MouseWheel>")
        self.root.deletecommand(marker)

    def test_serial_hover_does_not_retain_destroyed_canvas(self):
        win, text, panel, frame = self.open_debug_body()
        canvas = next(w for w in panel.winfo_children() if w.winfo_class() == "Canvas")
        ref = weakref.ref(canvas)
        canvas.event_generate("<Enter>")
        canvas.event_generate("<Leave>")
        win.destroy()
        del win, text, panel, frame, canvas
        gc.collect()
        self.assertIsNone(ref())

    def test_serial_wheel_scrolls_buttons_without_affecting_other_content(self):
        win, text, panel, frame = self.open_debug_body()
        canvas = next(w for w in panel.winfo_children() if w.winfo_class() == "Canvas")
        canvas.yview_moveto(0)
        frame.winfo_children()[0].event_generate("<MouseWheel>", delta=-120)
        self.assertGreater(canvas.yview()[0], 0)
        before = canvas.yview()
        text.event_generate("<MouseWheel>", delta=-120)
        self.assertEqual(canvas.yview(), before)
        win.destroy()

    def test_font_color_return_after_window_destroy_does_not_touch_widgets(self):
        for result in ((None, None), ((1, 2, 3), "#010203")):
            with self.subTest(result=result):
                win = self.open_font()
                picker = next(widget for frame in win.winfo_children() for widget in frame.winfo_children()
                              if widget.winfo_class() == "Button" and widget.cget("text") == "选颜色")

                def close_during_picker(**_kwargs):
                    win.destroy()
                    return result

                with patch.object(sms_font_dialog.colorchooser, "askcolor", side_effect=close_during_picker):
                    picker.invoke()
                self.assertFalse(win.winfo_exists())
                self.assertEqual(self.errors, [])

    def test_security_success_refreshes_from_committed_runtime_state(self):
        current = {"pin": True, "sn": True}
        def save(_permissions):
            current.update(pin=False, sn=False)
            return True
        refs = security_settings_dialog.open_security_settings_dialog(
            self.root, current, save, lambda *_: None, get_permissions=lambda: current)
        refs["permission_vars"]["sn"].set(False)
        self.assertTrue(refs["toggle"]("sn"))
        self.assertFalse(refs["permission_vars"]["pin"].get())
        self.assertFalse(refs["permission_vars"]["sn"].get())
        refs["close"]()

    def test_security_confirmation_after_window_destroy_does_not_change_permissions(self):
        for action in ("one", "all"):
            with self.subTest(action=action):
                changed = []
                refs = security_settings_dialog.open_security_settings_dialog(
                    self.root, {}, changed.append, lambda *_: None)

                def close_during_confirmation(*_args, **_kwargs):
                    refs["window"].destroy()
                    return True

                with patch.object(security_settings_dialog.messagebox, "askyesno", side_effect=close_during_confirmation):
                    if action == "one":
                        refs["permission_vars"]["sn"].set(True)
                        result = refs["toggle"]("sn")
                    else:
                        result = refs["set_all"](True)
                self.assertFalse(result)
                self.assertEqual(changed, [])

    def test_security_window_destroyed_during_save_does_not_refresh_closed_widgets(self):
        holder = {}
        def save(_permissions):
            holder["window"].destroy()
            return True
        refs = security_settings_dialog.open_security_settings_dialog(self.root, {}, save, lambda *_: None)
        holder.update(refs)
        with patch.object(security_settings_dialog.messagebox, "askyesno", return_value=True):
            refs["permission_vars"]["sn"].set(True)
            self.assertTrue(refs["toggle"]("sn"))
        self.assertEqual(self.errors, [])

    def test_security_parent_destroy_releases_active_window(self):
        parent = tk.Toplevel(self.root)
        refs = security_settings_dialog.open_security_settings_dialog(
            parent, {}, lambda _permissions: True, lambda *_: None)
        ref = weakref.ref(refs["window"])
        parent.destroy()
        del parent, refs
        gc.collect()
        self.assertIsNone(security_settings_dialog._active_security_settings_refs)
        self.assertIsNone(ref())

    def test_security_refresh_keeps_revoked_permissions_off_when_toggling_another(self):
        current = {"pin": True}
        callbacks = []

        def register(callback):
            callbacks.append(callback)
            return lambda: callbacks.remove(callback)

        def save(permissions):
            current.clear()
            current.update(permissions)
            return True

        refs = security_settings_dialog.open_security_settings_dialog(
            self.root, current, save, lambda *_: None,
            register_external_refresh=register, get_permissions=lambda: current)
        current["pin"] = False
        callbacks[0]()
        self.assertFalse(refs["permission_vars"]["pin"].get())
        refs["permission_vars"]["sn"].set(True)
        with patch.object(security_settings_dialog.messagebox, "askyesno", return_value=True):
            self.assertTrue(refs["toggle"]("sn"))
        self.assertEqual({key for key, value in current.items() if value}, {"sn"})
        refs["window"].destroy()
        self.assertEqual(callbacks, [])

    def test_security_refresh_during_confirmation_keeps_latest_value_on_cancel(self):
        current = {"sn": False}
        callbacks = []
        refs = security_settings_dialog.open_security_settings_dialog(
            self.root, current, lambda _permissions: True, lambda *_: None,
            register_external_refresh=lambda callback: callbacks.append(callback),
            get_permissions=lambda: current)

        def confirm(*_args, **_kwargs):
            current["sn"] = True
            callbacks[0]()
            return False

        refs["permission_vars"]["sn"].set(True)
        with patch.object(security_settings_dialog.messagebox, "askyesno", side_effect=confirm):
            self.assertFalse(refs["toggle"]("sn"))
        self.assertTrue(refs["permission_vars"]["sn"].get())
        refs["window"].destroy()


if __name__ == "__main__":
    unittest.main()
