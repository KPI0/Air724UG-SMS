import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.cloud_protocol import normalize_cloud_ws_url
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


class CloudControlFormController:
    def __init__(self, win, frame, state, state_provider, status_var, on_auto_upload_changed):
        self.win = win
        self.state_provider = state_provider
        self.on_auto_upload_changed = on_auto_upload_changed
        self.enabled_var = tk.BooleanVar(win, value=state["enabled"])
        self.auto_upload_var = tk.BooleanVar(win, value=state["auto_upload"])
        self.url_var = tk.StringVar(win, value=state["url"])
        self.secret_var = tk.StringVar(win, value=state["secret"])
        self.reconnect_var = tk.StringVar(win, value=str(state["reconnect_interval"]))

        self._expose_legacy_window_vars()
        self._build_options(frame)
        self._build_fields(frame)
        self._build_status(frame, status_var)

    def _expose_legacy_window_vars(self):
        self.win._enabled_var = self.enabled_var
        self.win._auto_upload_var = self.auto_upload_var
        self.win._url_var = self.url_var
        self.win._secret_var = self.secret_var
        self.win._reconnect_var = self.reconnect_var

    def _build_options(self, frame):
        top_opts_frame = ttk.Frame(frame)
        top_opts_frame.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 8))
        ttk.Checkbutton(
            top_opts_frame,
            text="启用云端控制",
            variable=self.enabled_var,
        ).pack(side="left", padx=(0, 20))
        ttk.Checkbutton(
            top_opts_frame,
            text="主动公开设备",
            variable=self.auto_upload_var,
            command=self._auto_upload_toggled,
        ).pack(side="left")

    def _build_fields(self, frame):
        configure_placeholder_style(self.win)
        self.url_field = CloudPlaceholderEntry(
            self.win,
            frame,
            1,
            "WebSocket 地址：",
            self.url_var,
            "wss://example.com/websocket",
            visible=True,
            help_command=lambda: show_url_reference(self.win),
        )

        ttk.Label(frame, text="重连间隔(秒)：").grid(row=2, column=0, sticky="w", pady=(0, 8))
        tk.Spinbox(frame, textvariable=self.reconnect_var, from_=1, to=3600, width=8).grid(
            row=2,
            column=1,
            sticky="w",
            pady=(0, 8),
        )

        self.secret_field = CloudPlaceholderEntry(
            self.win,
            frame,
            3,
            "控制密码：",
            self.secret_var,
            "自定义",
            visible=False,
            random_secret=True,
        )

    def _build_status(self, frame, status_var):
        ttk.Label(frame, text="当前状态：").grid(row=4, column=0, sticky="w", pady=(0, 10))
        ttk.Label(frame, textvariable=status_var).grid(row=4, column=1, sticky="w", pady=(0, 10))

    def _auto_upload_toggled(self):
        self.on_auto_upload_changed(bool(self.auto_upload_var.get()), self.win)

    def sync_from_state(self):
        latest = self.state_provider()
        self.enabled_var.set(latest["enabled"])
        self.auto_upload_var.set(latest["auto_upload"])
        self.url_var.set(latest["url"])
        self.secret_var.set(latest["secret"])
        self.reconnect_var.set(str(latest["reconnect_interval"]))

    def read(self, force_enabled=False):
        url = normalize_cloud_ws_url(self.url_field.get())
        if url:
            self.url_field.set_value(url)
        enabled = bool(self.enabled_var.get()) or bool(force_enabled)
        interval = self._read_interval()
        if interval is None:
            return None

        secret = self.secret_field.get()
        if enabled and not secret:
            messagebox.showerror("错误", "启用云端控制时，控制密码不能为空。", parent=self.win)
            return None

        if enabled and url.lower().startswith("ws://") and not confirm_insecure_ws(self.win):
            return None

        return enabled, url, interval, secret, bool(self.auto_upload_var.get())

    def _read_interval(self):
        try:
            interval = int(self.reconnect_var.get().strip())
            if interval < 1:
                raise ValueError
            return interval
        except Exception:
            messagebox.showerror("错误", "重连间隔必须是大于 0 的整数。", parent=self.win)
            return None

    def disconnect_values(self):
        self.enabled_var.set(False)
        return (
            False,
            self.url_field.get(),
            self.reconnect_var.get(),
            self.secret_field.get(),
            bool(self.auto_upload_var.get()),
        )


def confirm_insecure_ws(parent):
    return messagebox.askyesno(
        "安全警告",
        "您正在使用未加密的 ws:// 协议！\n\n"
        "设备控制密码将在网络中明文传输，容易被同网络环境中的第三方窃听并劫持设备。\n\n"
        "强烈建议配置 SSL 并使用 wss://。\n是否仍要继续保存？",
        parent=parent,
    )


def build_cloud_action_buttons(frame, save_command, connect_command, disconnect_command, close_command):
    btn_frame = ttk.Frame(frame)
    btn_frame.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 12))
    for col in range(4):
        btn_frame.grid_columnconfigure(col, weight=1, uniform="cloud_actions")
    action_buttons = (
        ("保存", save_command),
        ("连接", connect_command),
        ("断开", disconnect_command),
        ("关闭", close_command),
    )
    for col, (text, command) in enumerate(action_buttons):
        ttk.Button(btn_frame, text=text, width=10, command=command).grid(
            row=0,
            column=col,
            padx=4,
            sticky="ew",
        )
    return btn_frame
