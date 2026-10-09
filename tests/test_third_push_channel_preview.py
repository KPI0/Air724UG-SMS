import tkinter as tk
from tkinter import ttk
import unittest

from sms_core.config_schema import THIRD_PUSH_DEFAULTS
from sms_core.third_push import THIRD_PUSH_CHANNELS
from sms_ui.third_push_window_form import ThirdPushFormController


class ThirdPushChannelPreviewTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(str(exc))
        self.root.withdraw()
        self.addCleanup(self.root.destroy)
        self.win = tk.Toplevel(self.root)
        self.win.attributes('-alpha', 0)
        self.win.geometry('780x580')
        frame = ttk.Frame(self.win)
        frame.pack(fill='both', expand=True)
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(2, weight=1)
        self.state = dict(enabled=True, sms_enabled=True, call_enabled=True,
                          channels=['wecom'], settings=dict(THIRD_PUSH_DEFAULTS))
        self.option_calls, self.errors = [], []
        self.root.report_callback_exception = lambda *args, errors=self.errors: errors.append(args)
        self.form = ThirdPushFormController(self.win, frame, self.state, lambda: self.state,
                                           on_option_changed=lambda *args: self.option_calls.append(args))
        self.form.select_channel('wecom')
        self.root.update()

    def tearDown(self):
        self.assertEqual(self.errors, [])

    def checkbox(self, channel):
        label = dict(THIRD_PUSH_CHANNELS)[channel]
        return next(widget for widget in self.form.channel_box.winfo_children() if widget.cget('text') == label)

    def point(self, widget, element):
        y = widget.winfo_height() // 2
        xs = [x for x in range(widget.winfo_width()) if element in widget.identify(x, y)]
        self.assertTrue(xs, (widget.cget('text'), element))
        return xs[len(xs) // 2], y

    def click(self, channel, element):
        widget = self.checkbox(channel)
        x, y = self.point(widget, element)
        widget.event_generate('<ButtonPress-1>', x=x, y=y)
        widget.event_generate('<ButtonRelease-1>', x=x, y=y)
        self.root.update()

    def test_channel_names_only_preview_and_leave_saved_values_clean(self):
        before = self.form.collect(validate=False)
        for channel, _label in THIRD_PUSH_CHANNELS:
            with self.subTest(channel=channel):
                self.click(channel, 'label')
                self.click(channel, 'label')
                self.assertEqual(self.form.current_channel, channel)
                self.assertEqual(self.form.channel_list.curselection(), (self.form.channel_index[channel],))
                self.assertEqual(self.form.collect(validate=False), before)
                self.assertFalse(self.form.is_dirty())
        self.assertEqual(self.option_calls, [])

    def test_only_indicator_toggles_once_and_restores_clean_state_when_reverted(self):
        for channel, _label in THIRD_PUSH_CHANNELS:
            with self.subTest(channel=channel):
                initial = self.form.channel_vars[channel].get()
                self.click(channel, 'indicator')
                self.assertEqual(self.form.channel_vars[channel].get(), not initial)
                self.assertEqual(self.form.current_channel, channel)
                self.assertTrue(self.form.is_dirty())
                self.click(channel, 'label')
                self.assertEqual(self.form.channel_vars[channel].get(), not initial)
                self.click(channel, 'indicator')
                self.assertEqual(self.form.channel_vars[channel].get(), initial)
                self.assertFalse(self.form.is_dirty())

    def test_keyboard_space_still_toggles_native_checkbox(self):
        widget = self.checkbox('inotify')
        widget.focus_force()
        self.root.update()
        for expected in (True, False):
            widget.event_generate('<KeyPress-space>')
            widget.event_generate('<KeyRelease-space>')
            self.root.update()
            self.assertEqual(self.form.channel_vars['inotify'].get(), expected)
        self.assertFalse(self.form.is_dirty())

    def test_preview_preserves_existing_multiline_draft_and_channel_selection(self):
        self.form.select_channel('custom_post')
        draft = '{\n  "message": "draft content"\n}'
        self.form.custom_body_text.delete('1.0', 'end')
        self.form.custom_body_text.insert('1.0', draft)
        before = self.form.collect(validate=False)
        for channel, _label in THIRD_PUSH_CHANNELS:
            self.click(channel, 'label')
        self.click('custom_post', 'label')
        self.assertEqual(self.form.custom_body_text.get('1.0', 'end-1c'), draft)
        self.assertEqual(self.form.collect(validate=False), before)
        self.assertTrue(self.form.is_dirty())

    def test_disabled_checkbox_does_not_preview_or_toggle(self):
        widget = self.checkbox('inotify')
        widget.state(['disabled'])
        for element in ('label', 'indicator'):
            self.click('inotify', element)
            self.assertEqual(self.form.current_channel, 'wecom')
            self.assertFalse(self.form.channel_vars['inotify'].get())
        self.assertFalse(self.form.is_dirty())
