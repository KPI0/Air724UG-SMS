import tkinter as tk
from tkinter import ttk


def additional_missed_call_notice(total_count):
    try:
        additional_count = max(1, int(total_count)) - 1
    except (TypeError, ValueError):
        additional_count = 0
    if additional_count <= 0:
        return ""
    return f"另有 {additional_count} 个未接来电已显示在主窗口"


def format_call_time(value):
    if value is None:
        return ""
    try:
        return value.strftime("%Y-%m-%d %H:%M:%S")
    except (AttributeError, TypeError, ValueError):
        return str(value)


def open_missed_call_popup(
    parent,
    caller_num,
    started_at,
    center_window,
    on_close,
):
    """Open a non-modal, reusable missed-call reminder."""
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("未接来电提醒")
    win.minsize(380, 215)
    win.resizable(False, False)
    try:
        win.attributes("-topmost", True)
    except Exception:
        pass

    body = tk.Frame(win, bg="white", padx=24, pady=22)

    title_label = tk.Label(
        body,
        text="📵 未接来电",
        bg="white",
        fg="#c9302c",
        font=("Microsoft YaHei UI", 12, "bold"),
    )
    title_label.pack(pady=(0, 8))

    caller_label = tk.Label(
        body,
        bg="white",
        fg="#d9534f",
        font=("Microsoft YaHei UI", 17, "bold"),
    )
    caller_label.pack(pady=(0, 6))

    time_label = tk.Label(
        body,
        bg="white",
        fg="#555555",
        font=("Microsoft YaHei UI", 9),
    )
    time_label.pack(pady=(0, 8))

    additional_label = tk.Label(
        body,
        bg="white",
        fg="#555555",
        font=("Microsoft YaHei UI", 9),
    )

    footer = tk.Frame(win, bg="#f0f0f0", height=58)
    footer.pack(fill="x", side="bottom")
    footer.pack_propagate(False)
    close_button = ttk.Button(footer, text="确定", command=on_close, width=12)
    close_button.pack(pady=13)
    body.pack(fill="both", expand=True)

    def update_missed_call(next_caller_num, next_started_at, total_count=1):
        try:
            caller_label.configure(text=str(next_caller_num or "未知号码"))
            time_text = format_call_time(next_started_at)
            time_label.configure(text=f"来电时间：{time_text}" if time_text else "")
            notice = additional_missed_call_notice(total_count)
            additional_label.configure(text=notice)
            if notice:
                additional_label.pack(pady=(2, 0))
            else:
                additional_label.pack_forget()
            win.update_idletasks()
            center_window(win, parent)
            win.deiconify()
            win.lift()
            win.focus_force()
            close_button.focus_set()
        except Exception:
            pass
        return win

    win.missed_call_popup_update = update_missed_call
    win.protocol("WM_DELETE_WINDOW", on_close)
    win.bind("<Escape>", lambda _event: on_close())
    update_missed_call(caller_num, started_at, 1)
    return win
