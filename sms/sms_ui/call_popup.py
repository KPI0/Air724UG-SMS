import time
import tkinter as tk
from tkinter import ttk


def _format_call_duration(elapsed_seconds):
    total_seconds = max(0, int(elapsed_seconds))
    hours, remainder = divmod(total_seconds, 3600)
    minutes, seconds = divmod(remainder, 60)
    if hours:
        return f"{hours}:{minutes:02d}:{seconds:02d}"
    return f"{minutes:02d}:{seconds:02d}"


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
    on_close_callback = on_close
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

    duration_label = tk.Label(
        frm,
        text="00:00",
        font=("微软雅黑", 12),
        fg="#666666",
    )

    duration_after_id = None
    call_started_at = None
    duration_timer_running = False

    def stop_duration_timer():
        nonlocal duration_after_id, duration_timer_running
        duration_timer_running = False
        after_id = duration_after_id
        duration_after_id = None
        if after_id is None:
            return
        try:
            win.after_cancel(after_id)
        except Exception:
            pass

    def update_duration():
        nonlocal duration_after_id, duration_timer_running
        duration_after_id = None
        if not duration_timer_running:
            return
        try:
            if not win.winfo_exists():
                duration_timer_running = False
                return
            duration_label.config(text=_format_call_duration(time.monotonic() - call_started_at))
            duration_after_id = win.after(1000, update_duration)
        except Exception:
            duration_timer_running = False

    def start_duration_timer():
        nonlocal call_started_at, duration_timer_running
        if duration_timer_running:
            return
        try:
            call_started_at = time.monotonic()
            duration_timer_running = True
            duration_label.config(text="00:00")
            duration_label.pack(before=btn_frm, pady=(0, 14))
            update_duration()
        except Exception:
            stop_duration_timer()

    btn_frm = ttk.Frame(frm)
    btn_frm.pack(anchor="center")

    def mark_connected():
        try:
            if not win.winfo_exists():
                return
            title_label.config(text="📞 正在通话中...", fg="#2ecc71")
            start_duration_timer()
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

    def hide():
        on_ignore()

    def close_popup():
        stop_duration_timer()
        on_close_callback()

    btn_answer = ttk.Button(btn_frm, text="✅ 接听", command=answer)
    btn_answer.pack(side="left", padx=6)

    btn_hangup = ttk.Button(btn_frm, text="❌ 挂断", command=hangup)
    btn_hangup.pack(side="left", padx=6)

    btn_ignore = ttk.Button(btn_frm, text="隐藏", command=hide)
    btn_ignore.pack(side="left", padx=6)

    win._call_popup_cleanup = stop_duration_timer
    win.protocol("WM_DELETE_WINDOW", close_popup)
    center_window(win, parent)
    win.deiconify()
    win.lift()
    return win
