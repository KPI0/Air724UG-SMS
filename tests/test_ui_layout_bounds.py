import queue
import threading
import tkinter as tk
from tkinter import ttk
import unittest

from sms_ui.app_ui_namespace_runtime import center_window_namespace_runtime
from sms_ui.cloud_control_window import open_cloud_control_window_dialog
from sms_ui.main_window_layout import build_main_window_layout_runtime
from sms_ui.serial_debug_window import open_serial_debug_window_dialog
from sms_ui.serial_debug_panel import create_serial_debug_body


def descendants(widget):
    for child in widget.winfo_children():
        yield child
        yield from descendants(child)


class UiLayoutBoundsTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError as exc:
            self.skipTest(f"Tk display unavailable: {exc}")
        self.root.attributes("-alpha", 0)
        self.root.tk.call("tk", "scaling", 8 / 3)
        for name in ("TkDefaultFont", "TkTextFont"):
            self.root.tk.call("font", "configure", name, "-size", 12)
        self.root.geometry("520x280+100+100")
        self.errors = []
        self.root.report_callback_exception = lambda *args: self.errors.append(args)
        self.root.update()

    def tearDown(self):
        self.root.destroy()
        self.assertEqual(self.errors, [])

    def assert_visible(self, win, widget):
        self.assertTrue(widget.winfo_ismapped(), str(widget))
        x = widget.winfo_rootx() - win.winfo_rootx()
        y = widget.winfo_rooty() - win.winfo_rooty()
        self.assertGreaterEqual(x, 0)
        self.assertGreaterEqual(y, 0)
        self.assertLessEqual(x + widget.winfo_width(), win.winfo_width())
        self.assertLessEqual(y + widget.winfo_height(), win.winfo_height(),
                             f"{widget}: {widget.cget('text') if 'text' in widget.keys() else ''}")
        if widget.winfo_class() not in ("TEntry", "Entry"):
            self.assertGreaterEqual(widget.winfo_width(), widget.winfo_reqwidth())
            self.assertGreaterEqual(widget.winfo_height(), widget.winfo_reqheight())

    def test_main_statuses_remain_readable_when_narrow_and_updated(self):
        refs = build_main_window_layout_runtime(self.root, tk, cloud_enabled=True)
        self.root.geometry("520x200")
        self.root.update()
        refs["status_var"].set("串口已断开，正在等待设备重新连接")
        refs["cloud_var"].set("云端连接失败，正在重试")
        self.root.update()
        for name in ("status_label", "temp_label", "signal_label", "cloud_label"):
            self.assert_visible(self.root, refs[name])
        self.assertGreater(refs["text_area"].winfo_height(), 20)
        self.root.geometry("1800x480")
        refs["status_var"].set("已连接")
        refs["cloud_var"].set("已连接")
        self.root.update()
        self.assertEqual(len({refs[name].winfo_y() for name in
                             ("status_label", "temp_label", "signal_label", "cloud_label")}), 1)

    def open_debug(self):
        win, text = open_serial_debug_window_dialog(
            self.root, None, None, True, lambda: 100,
            queue.Queue(), threading.Lock(), lambda: None,
            lambda *_: None, lambda *_: None, lambda *_: None,
            lambda _: "已连接", lambda: "COM999", lambda *_: None,
            lambda *_: None, lambda *_: None, lambda: None,
            lambda *_: None,
        )
        win.attributes("-alpha", 0)
        return win, text

    def debug_controls(self, win):
        labels = ("启用原始输出旁路（不做任何过滤）", "清空", "⏸ 暂停",
                  "清除筛选", "加回车换行(\\r\\n)", "发送", "快捷命令 ▶")
        controls = {label: next(w for w in descendants(win) if "text" in w.keys()
                               and str(w.cget("text")) == label) for label in labels}
        entries = [w for w in descendants(win) if isinstance(w, ttk.Entry)]
        controls["filter"], controls["send"] = entries
        return controls

    def assert_same_row(self, *widgets):
        centers = [w.winfo_rooty() + w.winfo_height() / 2 for w in widgets]
        self.assertLessEqual(max(centers) - min(centers), 2)

    def test_serial_default_layout_keeps_original_rows_and_button_sizes(self):
        for scale in (4 / 3, 5 / 3):
            with self.subTest(scale=scale):
                self.root.tk.call("tk", "scaling", scale)
                for name in ("TkDefaultFont", "TkTextFont"):
                    self.root.tk.call("font", "configure", name, "-size", 9)
                win, text = self.open_debug()
                win.geometry("900x520")
                self.root.update()
                controls = self.debug_controls(win)
                self.assert_same_row(*(controls[name] for name in (
                    "启用原始输出旁路（不做任何过滤）", "清空", "⏸ 暂停", "filter", "清除筛选")))
                self.assert_same_row(*(controls[name] for name in (
                    "send", "加回车换行(\\r\\n)", "发送", "快捷命令 ▶")))
                for label in ("清空", "⏸ 暂停", "清除筛选"):
                    self.assertEqual(controls[label].winfo_width(), controls["发送"].winfo_width())
                self.assertEqual(controls["filter"].winfo_width(), controls["filter"].winfo_reqwidth())
                self.assertLess(controls["发送"].winfo_rootx(), controls["快捷命令 ▶"].winfo_rootx())
                self.assertGreater(text.winfo_height(), 360)
                win.destroy()

    def test_serial_layout_wraps_only_when_needed_and_restores_when_widened(self):
        win, _text = self.open_debug()
        controls = self.debug_controls(win)
        controls["filter"].insert(0, "test filter")
        controls["send"].insert(0, "AT")
        for _ in range(3):
            win.geometry("800x300")
            self.root.update()
            self.assertGreater(controls["清空"].winfo_rooty(),
                               controls["启用原始输出旁路（不做任何过滤）"].winfo_rooty())
            self.assertGreater(controls["发送"].winfo_rooty(), controls["send"].winfo_rooty())
            for widget in controls.values():
                self.assert_visible(win, widget)
            win.geometry("1800x520")
            self.root.update()
            self.assert_same_row(*(controls[name] for name in (
                "启用原始输出旁路（不做任何过滤）", "清空", "⏸ 暂停", "filter", "清除筛选")))
            self.assert_same_row(*(controls[name] for name in (
                "send", "加回车换行(\\r\\n)", "发送", "快捷命令 ▶")))
        self.assertEqual(controls["filter"].get(), "test filter")
        self.assertEqual(controls["send"].get(), "AT")

    def test_serial_tab_navigation_follows_send_row_in_both_layouts(self):
        win, _text = self.open_debug()
        controls = self.debug_controls(win)
        order = [controls[name] for name in ("send", "加回车换行(\\r\\n)", "发送", "快捷命令 ▶")]
        for size, width in ((9, 900), (12, 800), (12, 1800), (9, 900)):
            self.root.tk.call("tk", "scaling", 5 / 3 if size == 9 else 8 / 3)
            for name in ("TkDefaultFont", "TkTextFont"):
                self.root.tk.call("font", "configure", name, "-size", size)
            win.geometry(f"{width}x520")
            self.root.update()
            for previous, expected in zip(order, order[1:]):
                self.assertEqual(str(self.root.tk.call("tk_focusNext", str(previous))), str(expected))
                self.assertEqual(str(self.root.tk.call("tk_focusPrev", str(expected))), str(previous))
                previous.focus_force()
                self.root.update()
                previous.event_generate("<Tab>")
                self.root.update()
                self.assertIs(self.root.focus_get(), expected)

    def test_serial_actions_remain_visible_at_minimum_size(self):
        win, text = self.open_debug()
        win.geometry("800x300")
        self.root.update()
        for widget in descendants(win):
            if widget.winfo_class() in ("TButton", "TCheckbutton", "TEntry", "TLabel"):
                ancestor = widget.master
                while ancestor is not win and ancestor.winfo_class() != "Canvas":
                    ancestor = ancestor.master
                if ancestor is not win:
                    continue
                if widget.winfo_ismapped():
                    self.assert_visible(win, widget)
                else:
                    self.fail(f"Hidden action: {widget}")
        self.assertGreater(text.winfo_height(), 20)
        quick = next(widget for widget in descendants(win)
                     if isinstance(widget, ttk.Button) and str(widget.cget("text")) == "快捷命令 ▶")
        quick.invoke()
        self.root.update()
        canvas = next(widget for widget in descendants(win) if isinstance(widget, tk.Canvas))
        content = canvas.nametowidget(canvas.itemcget(canvas.find_all()[0], "window"))
        last = content.winfo_children()[-1]
        last.focus_force()
        self.root.update()
        self.assertGreaterEqual(canvas.winfo_height(), last.winfo_height())
        self.assertLessEqual(last.winfo_rooty() + last.winfo_height(),
                             canvas.winfo_rooty() + canvas.winfo_height())

    def test_cloud_actions_remain_visible_when_resized(self):
        state = dict(enabled=False, auto_upload=False, url="", secret="", reconnect_interval=30)
        status = tk.StringVar(self.root, value="未连接")
        win = open_cloud_control_window_dialog(
            self.root, lambda: state, status, lambda *_: None,
            lambda *_: None, lambda *_: None, lambda *_: None,
            lambda w: w.destroy(), lambda *_: None,
        )
        win.attributes("-alpha", 0)
        self.root.update()
        win.geometry(f"540x{win.winfo_height()}")
        self.root.update()
        for widget in descendants(win):
            if isinstance(widget, (ttk.Button, ttk.Checkbutton)):
                self.assert_visible(win, widget)

    def test_child_window_fits_screen_when_parent_is_near_edge(self):
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"520x280+{sw - 560}+{sh - 400}")
        self.root.update()
        win = tk.Toplevel(self.root)
        win.withdraw()
        win.attributes("-alpha", 0)
        win.geometry("800x600")
        center_window_namespace_runtime({}, win, self.root)
        self.assertFalse(win.winfo_ismapped())
        win.deiconify()
        self.root.update()
        self.assertLessEqual(win.winfo_rootx() + win.winfo_width(), sw)
        self.assertLessEqual(win.winfo_rooty() + win.winfo_height(), sh)

    def test_pending_minimum_size_is_used_before_window_maps(self):
        sw, sh = self.root.winfo_screenwidth(), self.root.winfo_screenheight()
        self.root.geometry(f"520x280+{sw - 560}+{sh - 400}")
        self.root.update()
        win = tk.Toplevel(self.root)
        win.withdraw()
        win.attributes("-alpha", 0)
        win.geometry("780x580")
        win.update_idletasks()
        win.minsize(780, 692)
        center_window_namespace_runtime({}, win, self.root)
        self.assertFalse(win.winfo_ismapped())
        win.deiconify()
        self.root.update()
        self.assertLessEqual(win.winfo_rooty() + win.winfo_height(), sh)

    def test_quick_panel_labels_are_scrollable_and_keyboard_focus_is_revealed(self):
        self.root.geometry("900x520")
        text, panel, content = create_serial_debug_body(self.root, lambda *_: None)
        panel.grid(row=0, column=1, sticky="ns")
        self.root.update()
        canvas = content.master
        buttons = content.winfo_children()
        for button in buttons:
            self.assertGreaterEqual(button.winfo_width(), button.winfo_reqwidth())
        self.assertLess(canvas.xview()[1], 1)
        canvas.xview_moveto(1)
        self.assertEqual(canvas.xview()[1], 1)
        buttons[-1].focus_force()
        self.root.update()
        self.assertLessEqual(buttons[-1].winfo_rooty() + buttons[-1].winfo_height(),
                             canvas.winfo_rooty() + canvas.winfo_height())
        self.assertGreater(text.winfo_width(), 100)


if __name__ == "__main__":
    unittest.main()
