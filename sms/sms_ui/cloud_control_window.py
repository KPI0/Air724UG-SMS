import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.cloud_protocol import normalize_cloud_ws_url
from sms_core.cloud_runtime import (
    cloud_auto_upload_action,
    cloud_control_save_kwargs,
)
from sms_ui.cloud_placeholder_entry import CloudPlaceholderEntry, configure_placeholder_style


def show_url_reference(parent):
    messagebox.showinfo(
        "WebSocket 地址参考",
        "参考格式：\n"
        "wss://example.com/websocket\n"
        "ws://192.168.1.100:8000/websocket\n\n"
        "如果只填写 ws://主机:端口，程序会自动补 /websocket。\n"
        "地址必须以 ws:// 或 wss:// 开头。",
        parent=parent,
    )


def open_cloud_control_window_dialog(
    parent,
    state_provider,
    status_var,
    on_auto_upload_changed,
    on_save,
    on_connect,
    on_disconnect,
    on_close,
    center_window,
):
    state = state_provider()
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("云端控制")
    win.minsize(480, 260)
    win.resizable(False, False)
    win.transient(parent)

    frame = ttk.Frame(win, padding=12)
    frame.pack(fill="both", expand=True)
    frame.grid_columnconfigure(1, weight=1)

    enabled_var = tk.BooleanVar(win, value=state["enabled"])
    auto_upload_var = tk.BooleanVar(win, value=state["auto_upload"])
    url_var = tk.StringVar(win, value=state["url"])
    secret_var = tk.StringVar(win, value=state["secret"])
    reconnect_var = tk.StringVar(win, value=str(state["reconnect_interval"]))

    win._enabled_var = enabled_var
    win._auto_upload_var = auto_upload_var
    win._url_var = url_var
    win._secret_var = secret_var
    win._reconnect_var = reconnect_var

    def sync_form_from_state():
        latest = state_provider()
        enabled_var.set(latest["enabled"])
        auto_upload_var.set(latest["auto_upload"])
        url_var.set(latest["url"])
        secret_var.set(latest["secret"])
        reconnect_var.set(str(latest["reconnect_interval"]))

    win._sync_form_from_state = sync_form_from_state

    def auto_upload_toggled():
        on_auto_upload_changed(bool(auto_upload_var.get()), win)

    top_opts_frame = ttk.Frame(frame)
    top_opts_frame.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))

    ttk.Checkbutton(
        top_opts_frame,
        text="启用云端控制",
        variable=enabled_var,
    ).pack(side="left", padx=(0, 20))
    ttk.Checkbutton(
        top_opts_frame,
        text="主动公开设备",
        variable=auto_upload_var,
        command=auto_upload_toggled,
    ).pack(side="left")

    configure_placeholder_style(win)
    url_field = CloudPlaceholderEntry(
        win,
        frame,
        1,
        "WebSocket 地址：",
        url_var,
        "wss://example.com/websocket",
        visible=True,
        help_command=lambda: show_url_reference(win),
    )

    ttk.Label(frame, text="重连间隔(秒)：").grid(row=2, column=0, sticky="w", pady=(0, 8))
    tk.Spinbox(frame, textvariable=reconnect_var, from_=1, to=3600, width=8).grid(
        row=2, column=1, sticky="w", pady=(0, 8)
    )

    secret_field = CloudPlaceholderEntry(
        win,
        frame,
        3,
        "控制密码：",
        secret_var,
        "自定义",
        visible=False,
        random_secret=True,
    )

    ttk.Label(frame, text="当前状态：").grid(row=4, column=0, sticky="w", pady=(0, 10))
    ttk.Label(frame, textvariable=status_var).grid(row=4, column=1, sticky="w", pady=(0, 10))

    btn_frame = ttk.Frame(frame)
    btn_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 12))
    for col in range(4):
        btn_frame.grid_columnconfigure(col, weight=1, uniform="cloud_actions")

    def read_form(force_enabled=False):
        url = normalize_cloud_ws_url(url_field.get())
        if url:
            url_field.set_value(url)
        enabled = bool(enabled_var.get()) or bool(force_enabled)
        try:
            interval = int(reconnect_var.get().strip())
            if interval < 1:
                raise ValueError
        except Exception:
            messagebox.showerror("错误", "重连间隔必须是大于 0 的整数。", parent=win)
            return None

        secret = secret_field.get()
        if enabled and not secret:
            messagebox.showerror("错误", "启用云端控制时，控制密码不能为空。", parent=win)
            return None

        if enabled and url.lower().startswith("ws://"):
            if not messagebox.askyesno(
                "安全警告",
                "您正在使用未加密的 ws:// 协议！\n\n"
                "设备控制密码将在网络中明文传输，容易被同网络环境中的第三方窃听并劫持设备。\n\n"
                "强烈建议配置 SSL 并使用 wss://。\n是否仍要继续保存？",
                parent=win,
            ):
                return None

        return enabled, url, interval, secret, bool(auto_upload_var.get())

    def save_only():
        values = read_form()
        if values is not None:
            on_save(values, win)

    def connect():
        values = read_form(force_enabled=True)
        if values is not None:
            enabled_var.set(True)
            on_connect(values, win)

    def disconnect():
        enabled_var.set(False)
        values = (
            False,
            url_field.get(),
            reconnect_var.get(),
            secret_field.get(),
            bool(auto_upload_var.get()),
        )
        on_disconnect(values, win)

    def close():
        on_close(win)

    action_buttons = (
        ("保存", save_only),
        ("连接", connect),
        ("断开", disconnect),
        ("关闭", close),
    )
    for col, (text, command) in enumerate(action_buttons):
        ttk.Button(btn_frame, text=text, width=10, command=command).grid(
            row=0, column=col, padx=4, sticky="ew"
        )

    win.protocol("WM_DELETE_WINDOW", close)
    win.bind("<Escape>", lambda _e: close())
    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    return win


def open_cloud_control_window_runtime(
    parent,
    current_window,
    state_provider,
    status_var,
    refresh_settings,
    save_setting,
    connection_state_provider,
    register_current,
    schedule_unregister,
    restart_control,
    stop_control,
    cloud_log,
    sync_existing_window,
    set_window,
    center_window,
):
    refresh_settings()

    if sync_existing_window(current_window, "_sync_form_from_state"):
        return current_window

    def on_auto_upload_changed(auto_upload, _win):
        was_public = bool(state_provider().get("auto_upload"))
        save_setting(auto_upload=auto_upload)
        try:
            connected, has_loop, has_conn = connection_state_provider()
            action = cloud_auto_upload_action(was_public, auto_upload, connected, has_loop, has_conn)
            if action == "register":
                register_current()
            elif action == "unregister":
                schedule_unregister("auto_upload_disabled")
        except Exception:
            pass

    def save_values(values, win):
        kwargs = cloud_control_save_kwargs(values)
        save_setting(**kwargs)
        messagebox.showinfo("配置已保存", "云端控制配置已成功保存！", parent=win)
        if kwargs["enabled"]:
            restart_control(show_errors=True)
        else:
            stop_control()
        cloud_log("配置已保存")

    def connect_values(values, _win):
        save_setting(**cloud_control_save_kwargs(values, enabled_override=True))
        restart_control(show_errors=True)

    def disconnect_values(values, _win):
        save_setting(**cloud_control_save_kwargs(values, enabled_override=False))
        stop_control()
        cloud_log("已手动断开")

    def close_window(win):
        try:
            win.destroy()
        except Exception:
            pass
        set_window(None)

    win = open_cloud_control_window_dialog(
        parent,
        state_provider,
        status_var,
        on_auto_upload_changed,
        save_values,
        connect_values,
        disconnect_values,
        close_window,
        center_window,
    )
    set_window(win)
    return win
