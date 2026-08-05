import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.serial_debug import (
    build_information_center_command,
    build_manual_operator_command,
    build_own_number_commands,
    build_sn_command,
)
from sms_ui.serial_debug_dialog_helpers import (
    create_debug_dialog,
    ensure_debug_enabled,
    finish_debug_dialog,
)


def open_modify_number_dialog(parent, enabled_var, send_commands, center_window):
    win = create_debug_dialog(parent)
    win.title("修改本机号码")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frm = ttk.Frame(win, padding=15)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="请输入新的手机号码：").pack(anchor="w", pady=(0, 10))
    num_var = tk.StringVar()
    ent = ttk.Entry(frm, textvariable=num_var, width=28)
    ent.pack(fill="x", pady=(0, 5))

    tk.Label(
        frm,
        text="💡 提示：需加 '+' 国际前缀 (如 +8618888888...)。",
        fg="gray",
        font=("微软雅黑", 9),
    ).pack(anchor="w", pady=(0, 15))

    def submit():
        if not ensure_debug_enabled(enabled_var, win):
            return
        phone = num_var.get().strip()
        if not phone:
            messagebox.showerror("错误", "手机号码不能为空", parent=win)
            return
        win.destroy()
        send_commands(build_own_number_commands(phone))

    finish_debug_dialog(win, center_window, parent, ent, submit)


def open_modify_information_center_dialog(
    parent,
    enabled_var,
    quick_send,
    center_window,
):
    win = create_debug_dialog(parent)
    win.title("修改信息中心号码")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frm = ttk.Frame(win, padding=15)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="请输入新的信息中心号码：").pack(
        anchor="w",
        pady=(0, 10),
    )
    number_var = tk.StringVar()
    ent = ttk.Entry(frm, textvariable=number_var, width=28)
    ent.pack(fill="x", pady=(0, 5))

    tk.Label(
        frm,
        text="💡 提示：请填写运营商提供的完整号码，通常以 +86 开头。",
        fg="gray",
        font=("微软雅黑", 9),
    ).pack(anchor="w", pady=(0, 15))

    def submit():
        if not ensure_debug_enabled(enabled_var, win):
            return
        try:
            command = build_information_center_command(number_var.get())
        except ValueError as exc:
            messagebox.showerror("错误", str(exc), parent=win)
            return
        win.destroy()
        quick_send(command)

    finish_debug_dialog(win, center_window, parent, ent, submit)


def open_manual_operator_dialog(parent, enabled_var, quick_send, center_window):
    win = create_debug_dialog(parent)
    win.title("手动切换运营商")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frm = ttk.Frame(win, padding=15)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="请输入运营商 PLMN：").pack(anchor="w", pady=(0, 10))
    plmn_var = tk.StringVar()
    ent = ttk.Entry(frm, textvariable=plmn_var, width=28)
    ent.pack(fill="x", pady=(0, 5))

    tk.Label(
        frm,
        text="💡 请填写附近运营商查询结果中的 5-6 位数字，切换时可能短暂掉线。",
        fg="gray",
        font=("微软雅黑", 9),
    ).pack(anchor="w", pady=(0, 15))

    def submit():
        if not ensure_debug_enabled(enabled_var, win):
            return
        try:
            command = build_manual_operator_command(plmn_var.get())
        except ValueError as exc:
            messagebox.showerror("错误", str(exc), parent=win)
            return
        win.destroy()
        quick_send(command)

    finish_debug_dialog(win, center_window, parent, ent, submit)


def open_modify_sn_dialog(parent, enabled_var, quick_send, center_window):
    win = create_debug_dialog(parent)
    win.title("修改设备SN码")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frm = ttk.Frame(win, padding=15)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="请输入新的 SN 码 (最长64位)：").pack(anchor="w", pady=(0, 10))
    sn_var = tk.StringVar()
    ent = ttk.Entry(frm, textvariable=sn_var, width=28)
    ent.pack(fill="x", pady=(0, 5))

    tk.Label(
        frm,
        text="⚠️ 警告：修改SN码可能导致串口异常，请尝试多次插拔。",
        fg="#d9534f",
        justify="left",
        font=("微软雅黑", 9),
    ).pack(anchor="w", pady=(0, 15))

    def submit():
        if not ensure_debug_enabled(enabled_var, win):
            return
        new_sn = sn_var.get().strip()
        if not new_sn:
            messagebox.showerror("错误", "SN码不能为空", parent=win)
            return
        if len(new_sn) > 64:
            messagebox.showerror("错误", "SN码最长不能超过64位", parent=win)
            return
        win.destroy()
        quick_send(build_sn_command(new_sn))

    finish_debug_dialog(win, center_window, parent, ent, submit)
