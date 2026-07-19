import tkinter as tk
from tkinter import ttk

from sms_core.serial_debug import (
    HANGUP_COMMAND,
    SERIAL_DEBUG_MAX_STORE_LINES,
    SERIAL_DEBUG_MAX_VISIBLE_LINES,
    build_dial_command,
    normalize_dial_number,
)
from sms_core.serial_sender import (
    send_command_async,
    send_command_sequence_async,
    send_text_sms_pdu_async,
)
from sms_ui.serial_debug_panel import (
    SerialDebugFinder,
    create_serial_debug_body,
    create_serial_debug_quick_actions,
    redraw_serial_debug_filter,
)
from sms_ui.serial_debug_runtime import (
    SerialDebugPauseController,
    close_serial_debug_runtime,
    start_serial_debug_append_loop,
)


def open_serial_debug_window_dialog(
    parent,
    current_window,
    current_text,
    debug_enabled,
    debug_drop_count,
    serial_queue,
    serial_lock,
    get_serial_obj,
    push_serial_debug,
    port_ui,
    set_status,
    format_connected_status,
    get_port,
    set_current_dial_num,
    set_debug_enabled,
    set_drop_count,
    clear_window_refs,
    center_window,
    window_title="串口调试",
    log_error=None,
):
    if current_window is not None and current_window.winfo_exists():
        current_window.deiconify()
        current_window.lift()
        current_window.focus_force()
        return current_window, current_text

    win = tk.Toplevel(parent)
    win.withdraw()
    win.title(window_title)
    win.geometry("900x520")
    win.minsize(800, 300)
    win.lift()
    win.focus_force()

    top = ttk.Frame(win)
    top.pack(fill="x", padx=8, pady=6)

    enabled_var = tk.BooleanVar(value=debug_enabled)
    pause_controller = None

    def toggle_enabled():
        set_debug_enabled(bool(enabled_var.get()))
        if pause_controller is not None:
            pause_controller.refresh()

    chk = ttk.Checkbutton(
        top,
        text="启用原始输出旁路（不做任何过滤）",
        variable=enabled_var,
        command=toggle_enabled,
    )
    chk.pack(side="left")

    all_debug_lines = []

    def clear_text():
        all_debug_lines.clear()
        serial_text.config(state="normal")
        serial_text.delete("1.0", "end")
        serial_text.config(state="disabled")

    ttk.Button(top, text="清空", width=8, command=clear_text).pack(side="left", padx=8)

    paused_var = tk.BooleanVar(value=False)
    btn_pause = ttk.Button(top, text="⏸ 暂停", width=8)
    btn_pause.pack(side="left")

    right_frame = ttk.Frame(top)
    right_frame.pack(side="right", padx=(8, 8))

    filter_var = tk.StringVar(value="")

    ttk.Label(right_frame, text="筛选：").grid(row=0, column=0, padx=(0, 4))
    filter_entry = ttk.Entry(right_frame, textvariable=filter_var, width=16)
    filter_entry.grid(row=0, column=1, padx=(0, 6))

    def redraw_by_filter():
        redraw_serial_debug_filter(serial_text, all_debug_lines, filter_var.get().strip())

    def clear_filter():
        filter_var.set("")
        redraw_by_filter()

    filter_var.trace_add("write", lambda *_: redraw_by_filter())
    ttk.Button(right_frame, text="清除筛选", width=8, command=clear_filter).grid(row=0, column=2)

    serial_status_bar = ttk.Frame(win)
    serial_status_bar.pack(side="bottom", fill="x", padx=8, pady=(0, 6))

    state_label = ttk.Label(serial_status_bar, text="")
    state_label.pack(side="left")

    drop_label = ttk.Label(top, text="")
    drop_label.pack(side="right")

    send_frame = ttk.Frame(win)
    send_frame.pack(side="bottom", fill="x", padx=8, pady=(0, 6))

    send_var = tk.StringVar()
    ttk.Label(send_frame, text="发送指令：").pack(side="left")

    send_entry = ttk.Entry(send_frame, textvariable=send_var)
    send_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

    crlf_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(send_frame, text="加回车换行(\\r\\n)", variable=crlf_var).pack(side="left", padx=(0, 8))

    def send_cmd(_event=None):
        if not enabled_var.get():
            return "break"
        cmd = send_var.get()
        if not cmd:
            return "break"

        send_command_async(
            serial_lock,
            get_serial_obj,
            cmd,
            append_crlf=crlf_var.get(),
            push_debug=push_serial_debug,
            log_error=log_error,
        )
        return "break"

    btn_send = ttk.Button(send_frame, text="发送", width=8, command=send_cmd)
    btn_send.pack(side="left")

    def quick_send(cmd):
        send_var.set(cmd)
        send_cmd()

    def send_serial_command_sequence(commands, delay_sec=0.3):
        send_command_sequence_async(
            serial_lock,
            get_serial_obj,
            commands,
            push_debug=push_serial_debug,
            delay_sec=delay_sec,
            log_error=log_error,
        )

    def send_text_sms_pdu(phone, msg):
        send_text_sms_pdu_async(
            serial_lock,
            get_serial_obj,
            phone,
            msg,
            push_debug=push_serial_debug,
            port_ui=port_ui,
            log_error=log_error,
        )

    btn_quick = ttk.Button(send_frame, text="快捷命令 ▶")
    btn_quick.pack(side="left", padx=(8, 0))

    send_entry.bind("<Return>", send_cmd)

    serial_text, quick_panel, quick_scroll_frame = create_serial_debug_body(
        win,
        quick_send,
    )

    def dial_phone(phone):
        phone = normalize_dial_number(phone)
        set_current_dial_num(phone)
        port_ui(f"📞 主动呼叫：拨打号码 {phone}", "normal")
        set_status(f"📞 呼叫中：{phone}", "blue")
        quick_send(build_dial_command(phone))

    def hangup_dialed_phone():
        port_ui("📞 已发送挂机指令 (ATH)", "normal")
        set_status(format_connected_status(get_port()), "green")
        quick_send(HANGUP_COMMAND)

    def close_active_dial_call():
        set_status(format_connected_status(get_port()), "green")
        quick_send(HANGUP_COMMAND)

    create_serial_debug_quick_actions(
        win,
        enabled_var,
        quick_scroll_frame,
        quick_panel,
        btn_quick,
        quick_send,
        send_serial_command_sequence,
        send_text_sms_pdu,
        dial_phone,
        hangup_dialed_phone,
        close_active_dial_call,
        center_window,
    )
    serial_text.config(state="disabled")

    pause_controller = SerialDebugPauseController(
        enabled_var,
        paused_var,
        state_label,
        btn_pause,
        chk,
        btn_send,
        send_entry,
        btn_quick,
        quick_scroll_frame,
        serial_text,
    )
    btn_pause.config(command=pause_controller.toggle)
    pause_controller.refresh()

    finder = SerialDebugFinder(win, serial_text)
    win.bind("<Control-f>", lambda _e: (finder.open(), "break"))
    win.bind("<Control-F>", lambda _e: (finder.open(), "break"))

    def on_close():
        close_serial_debug_runtime(
            win,
            serial_text,
            serial_queue,
            all_debug_lines,
            finder,
            enabled_var,
            pause_controller,
            drop_label,
            set_debug_enabled,
            set_drop_count,
            clear_window_refs,
        )

    win.protocol("WM_DELETE_WINDOW", on_close)
    win.bind("<Escape>", lambda _e: on_close())

    win.update_idletasks()
    try:
        center_window(win, parent)
    except Exception:
        pass

    win.deiconify()
    win.lift()
    win.focus_force()

    start_serial_debug_append_loop(
        win,
        lambda: serial_text,
        serial_queue,
        all_debug_lines,
        paused_var,
        drop_label,
        filter_var,
        finder,
        SERIAL_DEBUG_MAX_STORE_LINES,
        SERIAL_DEBUG_MAX_VISIBLE_LINES,
        lambda: debug_drop_count(),
    )

    return win, serial_text
