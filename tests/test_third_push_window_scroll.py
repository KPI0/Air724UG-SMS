import tkinter as tk
import threading
import unittest
import weakref
from tkinter import ttk
from unittest.mock import patch

from sms_core.config_schema import THIRD_PUSH_DEFAULTS
from sms_core.third_push import THIRD_PUSH_CHANNELS
from sms_ui import third_push_window as ui


def descendants(parent):
    for child in parent.winfo_children():
        yield child
        yield from descendants(child)


class ThirdPushWindowScrollTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("Tk display is unavailable")
        self.root.withdraw()
        self.addCleanup(self.root.destroy)
        previous_scale = self.root.tk.call("tk", "scaling")
        self.addCleanup(self.root.tk.call, "tk", "scaling", previous_scale)
        self.root.tk.call("tk", "scaling", 4 / 3)
        self.errors = []
        self.root.report_callback_exception = lambda *args, errors=self.errors: errors.append(args)
        for name in ("TkDefaultFont", "TkTextFont"):
            self.root.tk.call("font", "configure", name, "-size", 12)
        state = dict(enabled=False, sms_enabled=True, call_enabled=True,
                     channels=[], settings=dict(THIRD_PUSH_DEFAULTS))
        native_toplevel = tk.Toplevel

        class InvisibleWindow(native_toplevel):
            def deiconify(window):
                window.attributes("-alpha", 0)
                super().deiconify()

            def focus_force(window):
                pass

        with patch.object(ui.tk, "Toplevel", InvisibleWindow):
            self.win = ui.open_third_push_window_dialog(
                self.root, lambda: state, lambda *_: True, lambda *_: False,
                lambda window: window.destroy(), lambda *_: None)
        self.form = self.win._sync_form_from_globals.__self__
        self.root.update()

    def tearDown(self):
        self.assertEqual(self.errors, [])

    def select(self, channel):
        self.form.select_channel(channel)
        self.root.update()

    def field(self, key):
        variable = str(self.form.entry_vars[key])
        return next(widget for widget in descendants(self.form.param_box)
                    if isinstance(widget, ttk.Entry)
                    and str(widget.cget("textvariable")) == variable)

    def assert_visible(self, widget):
        canvas = self.form.param_canvas
        self.assertGreaterEqual(widget.winfo_rooty(), canvas.winfo_rooty())
        self.assertLessEqual(widget.winfo_rooty() + widget.winfo_height(),
                             canvas.winfo_rooty() + canvas.winfo_height())
        self.assertGreaterEqual(widget.winfo_rootx(), canvas.winfo_rootx())
        self.assertLessEqual(widget.winfo_rootx() + widget.winfo_width(),
                             canvas.winfo_rootx() + canvas.winfo_width())
        self.assertGreater(widget.winfo_width(), 40)

    def test_long_forms_reach_last_field_and_keep_footer_visible(self):
        for geometry in ("680x520", "780x580", "1000x760"):
            self.win.geometry(geometry)
            for channel, key in (("next-smtp-proxy", "next_smtp_proxy_subject"),
                                 ("email", "email_subject")):
                with self.subTest(geometry=geometry, channel=channel):
                    self.select(channel)
                    scrollbar = self.form.param_scrollbar
                    self.assertTrue(scrollbar.winfo_ismapped())
                    scrollbar.tk.call(scrollbar.cget("command"), "moveto", 1)
                    self.root.update()
                    self.assert_visible(self.field(key))
                    for button in descendants(self.win):
                        if isinstance(button, ttk.Button):
                            self.assertLessEqual(button.winfo_rooty() + button.winfo_height(),
                                                 self.win.winfo_rooty() + self.win.winfo_height())

    def test_keyboard_focus_reveals_hidden_field(self):
        self.select("next-smtp-proxy")
        last = self.field("next_smtp_proxy_subject")
        canvas = self.form.param_canvas
        self.assertGreater(last.winfo_rooty() + last.winfo_height(),
                           canvas.winfo_rooty() + canvas.winfo_height())
        last.event_generate("<FocusIn>")
        self.root.update()
        self.assert_visible(last)
        first = self.field("next_smtp_proxy_api")
        first.event_generate("<FocusIn>")
        self.root.update()
        self.assert_visible(first)

    def test_tab_navigation_skips_labels_and_reveals_parameters(self):
        self.select("next-smtp-proxy")
        first = self.field("next_smtp_proxy_api")
        last = self.field("next_smtp_proxy_subject")
        tk.Toplevel.focus_force(self.win)
        first.focus_set()
        self.root.update()
        self.assertIs(self.root.focus_get(), first)
        for _ in range(7):
            self.root.focus_get().event_generate("<Tab>")
            self.root.update()
        self.assertIs(self.root.focus_get(), last)
        self.assert_visible(last)
        for _ in range(7):
            self.root.focus_get().event_generate("<Shift-Tab>")
            self.root.update()
        self.assertIs(self.root.focus_get(), first)
        self.assert_visible(first)

    def test_long_status_keeps_footer_buttons_visible_at_large_font(self):
        status_label = next(widget for widget in descendants(self.win)
                            if isinstance(widget, ttk.Label) and widget.cget("textvariable"))
        buttons = [widget for widget in descendants(self.win) if isinstance(widget, ttk.Button)]
        for scale, size in ((4 / 3, 16), (8 / 3, 9)):
            self.root.tk.call("tk", "scaling", scale)
            for name in ("TkDefaultFont", "TkTextFont"):
                self.root.tk.call("font", "configure", name, "-size", size)
            for geometry in ("680x520", "780x580"):
                with self.subTest(scale=scale, size=size, geometry=geometry):
                    self.win.geometry(geometry)
                    self.root.setvar(str(status_label.cget("textvariable")), "✅ 配置已保存，测试已加入队列")
                    self.root.update()
                    for button in buttons:
                        self.assertTrue(button.winfo_ismapped())
                        self.assertGreaterEqual(button.winfo_width(), button.winfo_reqwidth())
                        self.assertLessEqual(button.winfo_rootx() + button.winfo_width(),
                                             self.win.winfo_rootx() + self.win.winfo_width())
                        self.assertLessEqual(button.winfo_rooty() + button.winfo_height(),
                                             self.win.winfo_rooty() + self.win.winfo_height())
                    self.assertGreaterEqual(status_label.winfo_height(), status_label.winfo_reqheight())

    def test_wheel_scrolls_form_without_changing_combobox_or_draft(self):
        self.select("email")
        combo = self.field("email_encryption")
        before = self.form.collect(validate=False)
        position = self.form.param_canvas.yview()
        combo.event_generate("<MouseWheel>", delta=-120)
        self.root.update()
        self.assertGreater(self.form.param_canvas.yview()[0], position[0])
        self.assertEqual(self.form.collect(validate=False), before)
        self.assertFalse(self.form.is_dirty())

    def test_channel_switch_resets_scroll_and_retains_custom_body(self):
        self.select("custom_post")
        text = self.form.custom_body_text
        body = "\n".join(f"draft line {index}" for index in range(40))
        text.delete("1.0", "end")
        text.insert("1.0", body)
        text.yview_moveto(0)
        self.root.update()
        position = self.form.param_canvas.yview()
        text.event_generate("<MouseWheel>", delta=-120)
        self.root.update()
        self.assertGreater(text.yview()[0], 0)
        self.assertEqual(self.form.param_canvas.yview(), position)
        self.select("email")
        self.form.param_canvas.yview_moveto(1)
        self.select("wecom")
        self.assertEqual(self.form.param_canvas.yview(), (0.0, 1.0))
        self.field("wecom_webhook").event_generate("<MouseWheel>", delta=-120)
        self.root.update()
        self.assertEqual(self.form.param_canvas.yview(), (0.0, 1.0))
        self.select("custom_post")
        self.assertEqual(self.form.custom_body_text.get("1.0", "end-1c"), body)
        self.assertEqual(self.form.param_canvas.yview()[0], 0)

    def test_all_channels_selectable_and_scroll_bindings_stay_local(self):
        self.win.geometry("680x520")
        global_binding = self.win.bind_all("<MouseWheel>")
        for channel, _label in THIRD_PUSH_CHANNELS:
            self.select(channel)
            index = self.form.channel_index[channel]
            self.assertIsNotNone(self.form.channel_list.bbox(index))
            self.assertEqual(self.form.current_channel, channel)
        self.win.destroy()
        self.root.update()
        self.assertEqual(self.root.bind_all("<MouseWheel>"), global_binding)

    def test_closed_forms_release_variables_and_callbacks_on_ui_thread(self):
        owner = threading.get_ident()
        finalized = []
        original = tk.Variable.__del__

        def finalize(variable):
            finalized.append(threading.get_ident())
            original(variable)

        with patch.object(tk.Variable, "__del__", new=finalize):
            for _ in range(3):
                reference = weakref.ref(self.form)
                variable_count = len(self.root.tk.call("info", "globals", "PY_VAR*"))
                self.form = None
                self.win.destroy()
                self.assertIsNone(reference())
                self.assertFalse(self.root.tk.call("info", "globals", "PY_VAR*"))
                self.assertEqual(len(finalized), variable_count)
                self.assertEqual(set(finalized), {owner})
                finalized.clear()
                state = dict(enabled=False, sms_enabled=True, call_enabled=True,
                             channels=[], settings=dict(THIRD_PUSH_DEFAULTS))
                self.win = ui.open_third_push_window_dialog(
                    self.root, lambda: state, lambda *_: True, lambda *_: False,
                    lambda window: window.destroy(), lambda *_: None)
                self.form = self.win._sync_form_from_globals.__self__
                self.form.entry_vars["wecom_webhook"].set("synthetic draft")
                self.assertTrue(self.form.is_dirty())


if __name__ == "__main__":
    unittest.main()
