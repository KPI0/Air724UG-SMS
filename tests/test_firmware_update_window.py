import gc
import threading
import tkinter as tk
import unittest
from unittest.mock import patch

from sms_ui import firmware_update_window as ui


class Updater:
    def __init__(self):
        self.state = dict(phase="idle", busy=False, can_start=False, can_cancel=False,
            config_changed=False, device="COM1 · IMEI 尾号 000001", current_version="1.0.13",
            target_version="1.0.14", filename="synthetic.dfota.bin", progress=0,
            message="连接设备后读取固件，或选择本地升级包")
        self.started = 0
        self.prepared = []

    def snapshot(self):
        return dict(self.state)

    def prepare(self, path=None):
        self.prepared.append(path)

    def start(self):
        self.started += 1

    def cancel(self):
        self.state.update(message="正在取消，等待当前 USB 操作结束…")


def descendants(parent):
    for child in parent.winfo_children():
        yield child
        yield from descendants(child)


class FirmwareUpdateWindowTests(unittest.TestCase):
    def setUp(self):
        try:
            self.root = tk.Tk()
        except tk.TclError:
            self.skipTest("Tk display is unavailable")
        self.root.withdraw()
        self.errors = []
        self.root.report_callback_exception = lambda *args: self.errors.append(args)
        native_toplevel = tk.Toplevel

        class InvisibleWindow(native_toplevel):
            def deiconify(window):
                window.attributes("-alpha", 0)
                super().deiconify()

            def winfo_screenheight(window):
                return 600

            def lift(window):
                pass

            def focus_set(window):
                pass

            def focus_force(window):
                pass

        self.patcher = patch.object(ui.tk, "Toplevel", InvisibleWindow)
        self.patcher.start()
        self.updater = Updater()
        self.namespace = dict(root=self.root, local_firmware_updater=self.updater, is_exiting=False)

    def tearDown(self):
        if hasattr(self, "patcher"):
            self.root.destroy()
            self.patcher.stop()

    def open(self, scale=2.0):
        self.root.tk.call("tk", "scaling", scale)
        win = ui.open_firmware_update_window(self.namespace)
        win.geometry("460x520")
        self.root.update()
        return win

    def buttons(self, win):
        return {w.cget("text"): w for w in descendants(win) if w.winfo_class() == "TButton"}

    def poll(self):
        self.root.after(220, self.root.quit)
        self.root.mainloop()
        self.root.update_idletasks()
        self.assertFalse(self.errors)

    def test_short_screen_large_font_preserves_full_status_and_footer(self):
        for scale in (1.0, 1.5, 2.0, 3.0):
            with self.subTest(scale=scale):
                self.updater.state.update(phase="installing", busy=True,
                    message="正在安装并等待重启确认，请保持供电；此阶段不能取消")
                win = self.open(scale)
                for label in descendants(win):
                    if label.winfo_class() == "TLabel":
                        self.assertGreaterEqual(label.winfo_height(), label.winfo_reqheight(),
                                                str(label.cget("text")))
                for text in ("关闭", "开始更新", "取消更新"):
                    button = self.buttons(win)[text]
                    x, y = button.winfo_rootx() - win.winfo_rootx(), button.winfo_rooty() - win.winfo_rooty()
                    self.assertGreaterEqual(x, 0)
                    self.assertGreaterEqual(y, 0)
                    self.assertLessEqual(x + button.winfo_width(), win.winfo_width())
                    self.assertLessEqual(y + button.winfo_height(), win.winfo_height())
                self.buttons(win)["关闭"].invoke()

    def test_actions_state_changes_and_close_reopen(self):
        win = self.open()
        buttons = self.buttons(win)
        self.assertEqual(str(buttons["开始更新"]["state"]), "disabled")
        self.updater.state.update(phase="ready", can_start=True)
        self.poll()
        with patch.object(ui.messagebox, "askyesno", return_value=False):
            buttons["开始更新"].invoke()
        self.assertEqual(self.updater.started, 0)
        with patch.object(ui.messagebox, "askyesno", return_value=True):
            buttons["开始更新"].invoke()
        self.assertEqual(self.updater.started, 1)
        with patch.object(ui.filedialog, "askopenfilename", return_value="synthetic.dfota.bin"):
            buttons["选择固件…"].invoke()
        self.assertEqual(self.updater.prepared, ["synthetic.dfota.bin"])
        for phase in ("reading", "transferring", "installing", "failed", "unconfirmed", "success"):
            busy = phase in {"reading", "transferring", "installing"}
            self.updater.state.update(phase=phase, busy=busy, can_start=False,
                                     can_cancel=phase in {"reading", "transferring"})
            self.poll()
            self.assertEqual(str(buttons["选择固件…"]["state"]), "disabled" if busy else "normal")
            self.assertEqual(str(buttons["取消更新"]["state"]),
                             "normal" if self.updater.state["can_cancel"] else "disabled")
        buttons["关闭"].invoke()
        self.poll()
        self.assertIsNone(self.namespace["firmware_update_window"])
        reopened = self.open()
        self.assertIs(self.namespace["local_firmware_updater"], self.updater)
        self.assertIs(ui.open_firmware_update_window(self.namespace), reopened)

    def test_overflow_scrolls_to_full_message_and_resets_after_resize(self):
        self.updater.state.update(phase="failed", message="设备空间不足或写入失败，更新已停止。" * 10)
        win = self.open(3.0)
        canvas = next(w for w in descendants(win) if w.winfo_class() == "Canvas")
        scrollbar = next(w for w in descendants(win) if w.winfo_class() == "TScrollbar")
        status = next(w for w in descendants(win)
                      if w.winfo_class() == "TLabel" and w.cget("text") == self.updater.state["message"])
        self.assertTrue(scrollbar.winfo_ismapped())
        before = canvas.yview()
        win.event_generate("<MouseWheel>", delta=-120)
        self.root.update()
        self.assertGreater(canvas.yview()[0], before[0])
        status.event_generate("<FocusIn>")
        self.root.update()
        self.assertLessEqual(status.winfo_rooty() + status.winfo_height(),
                             canvas.winfo_rooty() + canvas.winfo_height())
        canvas.yview_moveto(1)
        self.root.update()
        body = canvas.nametowidget(canvas.itemcget(canvas.find_all()[0], "window"))
        self.assertLessEqual(body.winfo_rooty() + body.winfo_height(),
                             canvas.winfo_rooty() + canvas.winfo_height())
        win.geometry("1000x1000")
        self.root.update()
        self.assertFalse(scrollbar.winfo_ismapped())
        self.assertEqual(canvas.yview(), (0.0, 1.0))
        self.assertFalse(self.errors)

    def test_rejected_package_keeps_actual_updater_device_labels(self):
        from tests.test_local_firmware_update import LocalFirmwareUpdateTests, FIXTURES
        fixture = LocalFirmwareUpdateTests()
        fixture.setUp()
        fixture.capability["version"] = "1.0.8"
        self.updater = fixture.updater
        self.namespace.update(fixture.ns, local_firmware_updater=self.updater)
        win = self.open()
        buttons = self.buttons(win)
        buttons["读取设备"].invoke()
        self.poll()
        device = self.updater.snapshot()["device"]
        with patch.object(ui.filedialog, "askopenfilename", return_value=str(FIXTURES / "synthetic.dfota.bin")):
            buttons["选择固件…"].invoke()
        self.poll()
        labels = {widget.cget("text") for widget in descendants(win) if widget.winfo_class() == "TLabel"}
        self.assertIn(device, labels)
        self.assertIn("1.0.8", labels)
        self.assertIn("目标版本必须高于当前版本；不支持同版本更新或降级", labels)
        self.assertEqual(str(buttons["开始更新"]["state"]), "disabled")
        buttons["关闭"].invoke()
        reopened = self.open()
        self.assertEqual(self.updater.snapshot()["device"], device)
        self.assertEqual(self.updater.snapshot()["current_version"], "1.0.8")
        self.assertEqual(str(self.buttons(reopened)["开始更新"]["state"]), "disabled")
        self.assertFalse(self.errors)

    def test_closed_window_releases_tk_variables_on_ui_thread(self):
        gc.collect()
        gc_enabled = gc.isenabled()
        gc.disable()
        finalized_on = []
        original = tk.Variable.__del__
        ui_thread = threading.get_ident()

        def finalized(variable):
            finalized_on.append(threading.get_ident())
            original(variable)

        def close_window():
            win = self.open()
            self.buttons(win)["关闭"].invoke()

        worker = threading.Thread(target=gc.collect, daemon=True)
        try:
            with patch.object(tk.Variable, "__del__", finalized):
                close_window()
                self.root.after(0, worker.start)

                def finished():
                    if worker.is_alive():
                        self.root.after(10, finished)
                    else:
                        self.root.quit()

                self.root.after(10, finished)
                self.root.mainloop()
                worker.join(3)
            self.assertGreaterEqual(len(finalized_on), 5)
            self.assertEqual(set(finalized_on), {ui_thread})
        finally:
            if gc_enabled:
                gc.enable()

    def test_dialog_result_after_window_destroyed_does_not_start_work(self):
        self.updater.state.update(phase="ready", can_start=True)
        for action in ("开始更新", "选择固件…"):
            with self.subTest(action=action):
                self.updater.started = 0
                self.updater.prepared.clear()
                win = self.open()

                def destroyed_dialog(**_kwargs):
                    win.destroy()
                    return "synthetic.dfota.bin"

                def destroyed_confirmation(*_args, **_kwargs):
                    win.destroy()
                    return True

                with patch.object(ui.filedialog, "askopenfilename", destroyed_dialog), \
                        patch.object(ui.messagebox, "askyesno", destroyed_confirmation):
                    self.buttons(win)[action].invoke()
                self.assertEqual(self.updater.started, 0)
                self.assertEqual(self.updater.prepared, [])
                self.assertFalse(self.errors)


if __name__ == "__main__":
    unittest.main()
