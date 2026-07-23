import tkinter as tk
from tkinter import ttk


def open_call_popup(
    parent,
    caller_num,
    center_window,
    on_answer,
    on_hangup,
    on_ignore,
    on_close,
):
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("来电提醒")
    win.minsize(300, 0)
    win.resizable(False, False)
    win.attributes("-topmost", True)

    frm = ttk.Frame(win, padding=15)
    frm.pack(fill="both", expand=True)

    title_label = tk.Label(
        frm,
        text="📞 收到新来电",
        font=("微软雅黑", 11, "bold"),
        fg="#0052cc",
    )
    title_label.pack(pady=(0, 8))

    tk.Label(
        frm,
        text=f"{caller_num}",
        font=("微软雅黑", 16, "bold"),
        fg="#d9534f",
    ).pack(pady=(0, 20))

    btn_frm = ttk.Frame(frm)
    btn_frm.pack(anchor="center")

    def mark_connected():
        try:
            if not win.winfo_exists():
                return
            title_label.config(text="📞 正在通话中...", fg="#2ecc71")
            btn_answer.pack_forget()
            btn_ignore.pack_forget()
            btn_hangup.config(state="normal")
            win.protocol("WM_DELETE_WINDOW", hangup)
        except Exception:
            pass

    def restore_answer():
        try:
            if btn_answer.winfo_exists():
                btn_answer.config(state="normal")
        except Exception:
            pass

    def restore_hangup():
        try:
            if btn_hangup.winfo_exists():
                btn_hangup.config(state="normal")
        except Exception:
            pass

    def answer():
        try:
            btn_answer.config(state="disabled")
        except Exception:
            pass
        on_answer(mark_connected, restore_answer)

    def hangup():
        try:
            btn_hangup.config(state="disabled")
        except Exception:
            pass
        on_hangup(restore_hangup)

    def ignore():
        on_ignore()

    btn_answer = ttk.Button(btn_frm, text="✅ 接听", command=answer)
    btn_answer.pack(side="left", padx=6)

    btn_hangup = ttk.Button(btn_frm, text="❌ 挂断", command=hangup)
    btn_hangup.pack(side="left", padx=6)

    btn_ignore = ttk.Button(btn_frm, text="忽略", command=ignore)
    btn_ignore.pack(side="left", padx=6)

    win.protocol("WM_DELETE_WINDOW", on_close)
    center_window(win, parent)
    win.deiconify()
    win.lift()
    return win
