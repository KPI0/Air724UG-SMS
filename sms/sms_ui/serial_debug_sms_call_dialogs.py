import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.sms_pdu import measure_text_sms_pdus
from sms_ui.serial_debug_dialog_helpers import ensure_debug_enabled


def format_sms_pdu_counter(message):
    message_text = str(message or "")
    if not message_text:
        return "0 字 | UCS2 0 字节 | 0 段", False
    info = measure_text_sms_pdus(message_text)
    return (
        f"{info.char_count} 字 | UCS2 {info.ucs2_bytes} 字节 | {info.segment_count}/{info.segment_limit} 段",
        info.too_long,
    )


def open_send_sms_dialog(parent, enabled_var, send_sms, center_window):
    win = tk.Toplevel(parent)
    win.title("发送短信")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frm = ttk.Frame(win, padding=15)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="接收方手机号：").pack(anchor="w")
    phone_var = tk.StringVar()
    ent_phone = ttk.Entry(frm, textvariable=phone_var, width=32)
    ent_phone.pack(fill="x", pady=(2, 10))

    ttk.Label(frm, text="短信内容：").pack(anchor="w")
    txt_msg = tk.Text(frm, height=4, width=30, font=("微软雅黑", 9))
    txt_msg.pack(fill="x", pady=(2, 2))

    count_var = tk.StringVar(value="0 字 | UCS2 0 字节 | 0 段")
    count_label = tk.Label(
        frm,
        textvariable=count_var,
        fg="gray",
        font=("微软雅黑", 8),
        justify="right",
        wraplength=320,
    )
    count_label.pack(anchor="e", pady=(0, 5))

    def update_char_count(_event=None):
        content = txt_msg.get("1.0", "end-1c")
        counter_text, too_long = format_sms_pdu_counter(content)
        count_var.set(counter_text)
        count_label.config(fg="#d9534f" if too_long else "gray")

    txt_msg.bind("<KeyRelease>", update_char_count)

    tk.Label(
        frm,
        text="长短信将按 UCS2 PDU 自动分段发送。\n💡 提示：支持 '+' 国际前缀 (如 +8618888888...)。",
        fg="gray",
        justify="left",
        font=("微软雅黑", 9),
        wraplength=320,
    ).pack(anchor="w", pady=(5, 10))

    def submit():
        if not ensure_debug_enabled(enabled_var, win):
            return

        phone = phone_var.get().strip()
        msg = txt_msg.get("1.0", "end-1c").strip()
        if not phone or not msg:
            messagebox.showerror("错误", "手机号和短信内容不能为空！", parent=win)
            return
        info = measure_text_sms_pdus(msg)
        if info.too_long:
            messagebox.showerror(
                "短信过长",
                f"当前内容需要 {info.segment_count} 个分段，超过 {info.segment_limit} 段限制。\n请删减后再发送。",
                parent=win,
            )
            return

        win.destroy()
        send_sms(phone, msg)

    buttons = ttk.Frame(frm)
    buttons.pack(anchor="e")
    ttk.Button(buttons, text="发送指令", command=submit).pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="取消", command=win.destroy).pack(side="left")

    win.update_idletasks()
    center_window(win, parent)
    ent_phone.focus_set()
    win.bind("<Control-Return>", lambda _e: submit())
    win.bind("<Escape>", lambda _e: win.destroy())


def open_dial_dialog(parent, enabled_var, dial, hangup, close_active_call, center_window):
    win = tk.Toplevel(parent)
    win.title("拨打电话")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    is_dialing = False

    frm = ttk.Frame(win, padding=15)
    frm.pack(fill="both", expand=True)

    ttk.Label(frm, text="请输入要拨打的手机/电话号码：").pack(anchor="w", pady=(0, 10))
    phone_var = tk.StringVar()
    ent_phone = ttk.Entry(frm, textvariable=phone_var, width=28)
    ent_phone.pack(fill="x", pady=(0, 5))

    tk.Label(
        frm,
        text="💡 提示：支持 '+' 国际前缀 (如 +8618888888...)。\n注意: 需确认 SIM 卡已开通语音/长途权限。",
        fg="gray",
        justify="left",
        font=("微软雅黑", 9),
    ).pack(anchor="w", pady=(0, 15))

    def submit_dial():
        nonlocal is_dialing
        if not ensure_debug_enabled(enabled_var, win):
            return
        phone = phone_var.get().strip()
        if not phone:
            messagebox.showerror("错误", "号码不能为空", parent=win)
            return
        is_dialing = True
        dial(phone)

    def submit_hangup():
        nonlocal is_dialing
        if not ensure_debug_enabled(enabled_var, win):
            return
        if not is_dialing:
            return
        is_dialing = False
        hangup()

    def close_dialog():
        nonlocal is_dialing
        if is_dialing and enabled_var.get():
            is_dialing = False
            close_active_call()
        win.destroy()

    buttons = ttk.Frame(frm)
    buttons.pack(anchor="e")
    ttk.Button(buttons, text="📞 拨号", command=submit_dial).pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="挂断", command=submit_hangup).pack(side="left", padx=(0, 8))
    ttk.Button(buttons, text="取消", command=close_dialog).pack(side="left")

    win.update_idletasks()
    center_window(win, parent)
    ent_phone.focus_set()
    win.bind("<Return>", lambda _e: submit_dial())
    win.bind("<Escape>", lambda _e: close_dialog())
    win.protocol("WM_DELETE_WINDOW", close_dialog)
