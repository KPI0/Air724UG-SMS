import queue
import tkinter as tk
from tkinter import ttk

from sms_core.serial_debug import COMMON_SERIAL_COMMANDS, quick_command_label
from sms_core.threading_runtime import task_done_safely
from sms_ui.serial_debug_finder import SerialDebugFinder
from sms_ui.serial_debug_dialogs import (
    open_dial_dialog,
    open_input_pin_dialog,
    open_input_puk_dialog,
    open_modify_number_dialog,
    open_modify_pin_dialog,
    open_modify_sn_dialog,
    open_pin_lock_dialog,
    open_send_sms_dialog,
)


def create_serial_debug_body(parent, quick_send):
    body = ttk.Frame(parent)
    body.pack(fill="both", expand=True, padx=8, pady=6)

    body.grid_rowconfigure(0, weight=1)
    body.grid_columnconfigure(0, weight=1)
    body.grid_columnconfigure(1, weight=0)

    text_frame = ttk.Frame(body)
    text_frame.grid(row=0, column=0, sticky="nsew")

    yscroll = ttk.Scrollbar(text_frame, orient="vertical")
    yscroll.pack(side="right", fill="y")

    xscroll = ttk.Scrollbar(text_frame, orient="horizontal")
    xscroll.pack(side="bottom", fill="x")

    serial_text = tk.Text(
        text_frame,
        wrap="none",
        yscrollcommand=yscroll.set,
        xscrollcommand=xscroll.set,
    )
    serial_text.pack(side="left", fill="both", expand=True)
    yscroll.config(command=serial_text.yview)
    xscroll.config(command=serial_text.xview)

    quick_panel = ttk.LabelFrame(body, text="常用指令")
    quick_canvas = tk.Canvas(quick_panel, highlightthickness=0, width=330)
    quick_scrollbar = ttk.Scrollbar(quick_panel, orient="vertical", command=quick_canvas.yview)
    quick_scroll_frame = ttk.Frame(quick_canvas)

    quick_scroll_window = quick_canvas.create_window((0, 0), window=quick_scroll_frame, anchor="nw")
    quick_canvas.configure(yscrollcommand=quick_scrollbar.set)

    quick_scrollbar.pack(side="right", fill="y")
    quick_canvas.pack(side="left", fill="both", expand=True)

    quick_scroll_frame.bind(
        "<Configure>",
        lambda _e: quick_canvas.configure(scrollregion=quick_canvas.bbox("all")),
    )
    quick_canvas.bind(
        "<Configure>",
        lambda e: quick_canvas.itemconfig(quick_scroll_window, width=e.width),
    )

    def bind_mousewheel(_event):
        quick_canvas.bind_all(
            "<MouseWheel>",
            lambda e: quick_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"),
        )

    def unbind_mousewheel(_event):
        quick_canvas.unbind_all("<MouseWheel>")

    quick_canvas.bind("<Enter>", bind_mousewheel)
    quick_canvas.bind("<Leave>", unbind_mousewheel)

    for cmd, desc in COMMON_SERIAL_COMMANDS:
        ttk.Button(
            quick_scroll_frame,
            text=quick_command_label(cmd, desc),
            command=lambda c=cmd: quick_send(c),
        ).pack(fill="x", padx=6, pady=3)

    return serial_text, quick_panel, quick_scroll_frame


def create_serial_debug_quick_actions(
    parent,
    enabled_var,
    quick_scroll_frame,
    quick_panel,
    quick_button,
    quick_send,
    send_command_sequence,
    send_text_sms_pdu,
    dial_phone,
    hangup_dialed_phone,
    close_active_dial_call,
    center_window,
):
    ttk.Button(
        quick_scroll_frame,
        text="输入PIN码解锁 🔑",
        command=lambda: open_input_pin_dialog(parent, enabled_var, quick_send, center_window),
    ).pack(fill="x", padx=6, pady=(6, 6))
    ttk.Button(
        quick_scroll_frame,
        text="输入PUK码解锁 🔐",
        command=lambda: open_input_puk_dialog(parent, enabled_var, quick_send, center_window),
    ).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(
        quick_scroll_frame,
        text="开启PIN码锁 🔒",
        command=lambda: open_pin_lock_dialog(parent, enabled_var, quick_send, center_window, enable=True),
    ).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(
        quick_scroll_frame,
        text="关闭PIN码锁 🔓",
        command=lambda: open_pin_lock_dialog(parent, enabled_var, quick_send, center_window, enable=False),
    ).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(
        quick_scroll_frame,
        text="修改PIN码 ✏️",
        command=lambda: open_modify_pin_dialog(parent, enabled_var, quick_send, center_window),
    ).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(
        quick_scroll_frame,
        text="修改本机号码 ☎",
        command=lambda: open_modify_number_dialog(
            parent,
            enabled_var,
            send_command_sequence,
            center_window,
        ),
    ).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(
        quick_scroll_frame,
        text="修改SN码 🏷️",
        command=lambda: open_modify_sn_dialog(parent, enabled_var, quick_send, center_window),
    ).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(
        quick_scroll_frame,
        text="发送短信 ✉️",
        command=lambda: open_send_sms_dialog(
            parent,
            enabled_var,
            send_text_sms_pdu,
            center_window,
        ),
    ).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(
        quick_scroll_frame,
        text="拨打电话 📞",
        command=lambda: open_dial_dialog(
            parent,
            enabled_var,
            dial_phone,
            hangup_dialed_phone,
            close_active_dial_call,
            center_window,
        ),
    ).pack(fill="x", padx=6, pady=(0, 6))

    panel_visible = False

    def toggle_quick_panel():
        nonlocal panel_visible
        if panel_visible:
            quick_panel.grid_remove()
            quick_button.config(text="快捷命令 ▶")
            panel_visible = False
        else:
            quick_panel.grid(row=0, column=1, sticky="ns", padx=(8, 0))
            quick_button.config(text="快捷命令 ◀")
            panel_visible = True

    quick_button.config(command=toggle_quick_panel)
    return toggle_quick_panel


def redraw_serial_debug_filter(text_widget, all_lines, filter_text: str):
    text_widget.config(state="normal")
    text_widget.delete("1.0", "end")

    for line in all_lines:
        if filter_text and filter_text not in line:
            continue
        if not line.endswith("\n"):
            line += "\n"
        text_widget.insert("end", line)

    text_widget.see("end")
    text_widget.config(state="disabled")


def append_serial_debug_lines_once(
    text_widget,
    serial_queue,
    all_lines,
    filter_text: str,
    finder: SerialDebugFinder,
    max_store_lines: int,
    max_visible_lines: int,
    batch_size: int = 200,
):
    lines = []
    for _ in range(batch_size):
        try:
            line = serial_queue.get_nowait()
        except queue.Empty:
            break
        task_done_safely(serial_queue)
        lines.append(line)

    if not lines:
        return False

    was_at_bottom = _text_view_is_at_bottom(text_widget)
    all_lines.extend(lines)
    if len(all_lines) > max_store_lines:
        all_lines[:] = all_lines[-max_store_lines:]

    text_widget.config(state="normal")
    insert_start = text_widget.index("end-1c")
    for line in lines:
        if filter_text and filter_text not in line:
            continue
        if not line.endswith("\n"):
            line += "\n"
        text_widget.insert("end", line)
    insert_end = text_widget.index("end-1c")

    if finder.term:
        finder.highlight_range(insert_start, insert_end)

    try:
        cur_lines = int(text_widget.index("end-1c").split(".")[0])
        if cur_lines > max_visible_lines:
            del_lines = cur_lines - max_visible_lines
            text_widget.delete("1.0", f"{del_lines + 1}.0")
            finder.clear()
            if finder.term:
                finder.find_all(finder.term)
    except Exception:
        pass

    if was_at_bottom:
        text_widget.see("end")
    text_widget.config(state="disabled")
    return True


def _text_view_is_at_bottom(text_widget, threshold=0.98):
    try:
        return text_widget.yview()[1] >= threshold
    except Exception:
        return True


def reset_serial_debug_window_state(
    window,
    text_widget,
    serial_queue,
    all_lines,
    finder: SerialDebugFinder,
):
    try:
        while True:
            serial_queue.get_nowait()
            task_done_safely(serial_queue)
    except queue.Empty:
        pass

    try:
        all_lines.clear()
    except Exception:
        pass

    try:
        window.unbind_all("<MouseWheel>")
    except Exception:
        pass

    try:
        if text_widget is not None and text_widget.winfo_exists():
            text_widget.config(state="normal")
            text_widget.delete("1.0", "end")
            text_widget.config(state="disabled")
    except Exception:
        pass

    try:
        finder.close()
    except Exception:
        pass
