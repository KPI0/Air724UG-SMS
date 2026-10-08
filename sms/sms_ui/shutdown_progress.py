"""The UI-thread-owned progress view for graceful application shutdown."""
import tkinter as tk
from tkinter import ttk


class ShutdownProgress:
    def __init__(self, root, *, title="正在退出", message="正在停止接收并保存已收到的短信…"):
        root.deiconify()
        self.window = tk.Toplevel(root)
        self.window.withdraw()
        self.window.title(title)
        self.window.transient(root)
        self.window.resizable(False, False)
        self.window.protocol("WM_DELETE_WINDOW", lambda: None)
        body = ttk.Frame(self.window, padding=20)
        body.pack(fill="both", expand=True)
        self.status = ttk.Label(body, text=message, wraplength=340)
        self.status.pack(fill="x", pady=(0, 12))
        self.indicator = ttk.Progressbar(body, mode="indeterminate", length=340)
        self.indicator.pack(fill="x")
        self.window.tk.call("tk::PlaceWindow", self.window, "widget", root)
        # Shorter status messages must not shrink the dialog away from its center.
        self.window.minsize(self.window.winfo_reqwidth(), self.window.winfo_reqheight())
        self.indicator.start(30)
        self.indicator.bind("<Destroy>", self._stop_destroyed_indicator, add="+")
        self.window.grab_set()

    def _stop_destroyed_indicator(self, _event):
        try:
            # The widget command may already be gone. Ttk's stop procedure
            # cancels its timer before attempting to reset the widget value.
            self.indicator.tk.call("ttk::progressbar::stop", str(self.indicator))
        except tk.TclError:
            pass

    def update(self, message):
        self.status.configure(text=message)

    def close(self):
        self.indicator.stop()
        self.window.grab_release()
        self.window.destroy()
