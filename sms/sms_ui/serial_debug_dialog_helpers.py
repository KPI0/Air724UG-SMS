import tkinter as tk
from tkinter import messagebox, ttk


def ensure_debug_enabled(enabled_var, parent):
    if enabled_var.get():
        return True
    messagebox.showwarning("提示", "请先勾选顶部的“启用原始输出旁路”", parent=parent)
    return False


def create_debug_dialog(parent):
    win = tk.Toplevel(parent)
    win.withdraw()
    return win


def show_centered_debug_dialog(win, center_window, parent, focus_widget):
    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    focus_widget.focus_set()


def finish_debug_dialog(win, center_window, parent, focus_widget, command):
    buttons = ttk.Frame(win.winfo_children()[0])
    buttons.pack(anchor="e")
    ttk.Button(buttons, text="发送指令", command=command).pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="取消", command=win.destroy).pack(side="left")

    show_centered_debug_dialog(win, center_window, parent, focus_widget)
    win.bind("<Return>", lambda _e: command())
    win.bind("<Escape>", lambda _e: win.destroy())
