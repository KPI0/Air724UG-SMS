import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.serial_debug import (
    build_pin_change_command,
    build_pin_lock_command,
    build_pin_unlock_command,
    build_puk_unlock_command,
)
from sms_ui.serial_debug_dialog_helpers import (
    create_debug_dialog,
    ensure_debug_enabled,
    finish_debug_dialog,
)


def open_input_pin_dialog(parent, enabled_var, quick_send, center_window):
    win = create_debug_dialog(parent)
    win.title("输入PIN码解锁")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frm = ttk.Frame(win, padding=15)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="请输入 SIM 卡 PIN 码：").pack(anchor="w", pady=(0, 10))
    pin_var = tk.StringVar()
    ent = ttk.Entry(frm, textvariable=pin_var, width=28)
    ent.pack(fill="x", pady=(0, 5))

    tk.Label(
        frm,
        text="⚠️ 警告：连续3次错误将锁定，需PUK码解锁！",
        fg="#d9534f",
        font=("微软雅黑", 9),
    ).pack(anchor="w", pady=(0, 15))

    def submit():
        if not ensure_debug_enabled(enabled_var, win):
            return
        pin = pin_var.get().strip()
        if not pin:
            messagebox.showerror("错误", "PIN码不能为空", parent=win)
            return
        win.destroy()
        quick_send(build_pin_unlock_command(pin))

    finish_debug_dialog(win, center_window, parent, ent, submit)


def open_input_puk_dialog(parent, enabled_var, quick_send, center_window):
    win = create_debug_dialog(parent)
    win.title("输入PUK码解锁")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frm = ttk.Frame(win, padding=15)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="请输入 PUK 码 (通常为8位)：").pack(anchor="w")
    puk_var = tk.StringVar()
    ent_puk = ttk.Entry(frm, textvariable=puk_var, width=28)
    ent_puk.pack(fill="x", pady=(2, 2))

    tk.Label(
        frm,
        text="⚠️ 致命警告：连续10次错误将永久烧毁SIM卡！",
        fg="#d9534f",
        font=("微软雅黑", 9, "bold"),
    ).pack(anchor="w", pady=(0, 12))

    ttk.Label(frm, text="请设置 新 PIN 码 (通常为4-8位数字)：").pack(anchor="w")
    new_pin_var = tk.StringVar()
    ent_new_pin = ttk.Entry(frm, textvariable=new_pin_var, width=28)
    ent_new_pin.pack(fill="x", pady=(2, 15))

    def submit():
        if not ensure_debug_enabled(enabled_var, win):
            return
        puk = puk_var.get().strip()
        new_pin = new_pin_var.get().strip()
        if not puk or not new_pin:
            messagebox.showerror("错误", "PUK码和新PIN码都不能为空！", parent=win)
            return
        win.destroy()
        quick_send(build_puk_unlock_command(puk, new_pin))

    finish_debug_dialog(win, center_window, parent, ent_puk, submit)


def open_pin_lock_dialog(parent, enabled_var, quick_send, center_window, enable: bool):
    title = "开启PIN码锁" if enable else "关闭PIN码锁"
    tip = "💡 提示：开启后模组每次开机均需输入PIN码。" if enable else "💡 提示：关闭后模组开机将自动联网，不再拦截。"

    win = create_debug_dialog(parent)
    win.title(title)
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frm = ttk.Frame(win, padding=15)
    frm.pack(fill="both", expand=True)

    prompt = "请输入当前的 PIN 码以开启锁定：" if enable else "请输入当前的 PIN 码以关闭锁定："
    ttk.Label(frm, text=prompt).pack(anchor="w", pady=(0, 10))
    pin_var = tk.StringVar()
    ent = ttk.Entry(frm, textvariable=pin_var, width=28)
    ent.pack(fill="x", pady=(0, 5))

    tk.Label(frm, text=tip, fg="gray", font=("微软雅黑", 9)).pack(anchor="w", pady=(0, 15))

    def submit():
        if not ensure_debug_enabled(enabled_var, win):
            return
        pin = pin_var.get().strip()
        if not pin:
            messagebox.showerror("错误", "PIN码不能为空", parent=win)
            return
        win.destroy()
        quick_send(build_pin_lock_command(pin, enable=enable))

    finish_debug_dialog(win, center_window, parent, ent, submit)


def open_modify_pin_dialog(parent, enabled_var, quick_send, center_window):
    win = create_debug_dialog(parent)
    win.title("修改PIN码")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frm = ttk.Frame(win, padding=15)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="请输入 旧 PIN 码：").pack(anchor="w")
    old_pin_var = tk.StringVar()
    ent_old = ttk.Entry(frm, textvariable=old_pin_var, width=28)
    ent_old.pack(fill="x", pady=(2, 10))

    ttk.Label(frm, text="请输入 新 PIN 码 (通常为4-8位数字)：").pack(anchor="w")
    new_pin_var = tk.StringVar()
    ent_new = ttk.Entry(frm, textvariable=new_pin_var, width=28)
    ent_new.pack(fill="x", pady=(2, 15))

    def submit():
        if not ensure_debug_enabled(enabled_var, win):
            return
        old_pin = old_pin_var.get().strip()
        new_pin = new_pin_var.get().strip()
        if not old_pin or not new_pin:
            messagebox.showerror("错误", "旧PIN码和新PIN码都不能为空", parent=win)
            return
        win.destroy()
        quick_send(build_pin_change_command(old_pin, new_pin))

    finish_debug_dialog(win, center_window, parent, ent_old, submit)
