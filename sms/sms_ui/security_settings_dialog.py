import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.cloud_command_security import (
    CLOUD_COMMAND_PERMISSION_SPECS,
    cloud_sensitive_commands_status,
    normalize_cloud_command_permissions,
)


_active_security_settings_refs = None


def _focus_active_security_settings_dialog():
    global _active_security_settings_refs
    refs = _active_security_settings_refs
    if not isinstance(refs, dict):
        return None
    win = refs.get("window")
    try:
        if win is None or not win.winfo_exists():
            raise RuntimeError("window no longer exists")
        win.deiconify()
        win.lift()
        win.focus_force()
        return refs
    except Exception:
        _active_security_settings_refs = None
        return None


def _fit_security_window(win, frame, *, min_width=620, min_height=285):
    """Size the fixed dialog from its requested content, with DPI-safe minimums."""
    try:
        requested_width = int(frame.winfo_reqwidth())
    except Exception:
        requested_width = min_width
    try:
        requested_height = int(frame.winfo_reqheight())
    except Exception:
        requested_height = min_height

    width = max(min_width, requested_width)
    height = max(min_height, requested_height)
    win.geometry(f"{width}x{height}")


def open_security_settings_dialog(
    parent,
    current_permissions,
    on_change,
    center_window,
):
    global _active_security_settings_refs
    active_refs = _focus_active_security_settings_dialog()
    if active_refs is not None:
        return active_refs

    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("安全设置")
    win.geometry("620x285")
    win.resizable(False, False)

    frame = ttk.Frame(win, padding=12)
    frame.pack(fill=tk.BOTH, expand=True)

    permissions = normalize_cloud_command_permissions(current_permissions)
    permission_vars = {
        spec.category: tk.BooleanVar(value=permissions[spec.category])
        for spec in CLOUD_COMMAND_PERMISSION_SPECS
    }
    status_var = tk.StringVar(value=cloud_sensitive_commands_status(permissions))

    def close():
        global _active_security_settings_refs
        if (
            isinstance(_active_security_settings_refs, dict)
            and _active_security_settings_refs.get("window") is win
        ):
            _active_security_settings_refs = None
        win.destroy()

    def apply_permissions(next_permissions):
        nonlocal permissions
        previous_permissions = dict(permissions)
        try:
            if on_change(next_permissions) is False:
                raise RuntimeError("配置保存失败")
        except Exception as exc:
            for category, value in previous_permissions.items():
                permission_vars[category].set(value)
            messagebox.showerror(
                "设置失败",
                f"安全设置未能保存，已恢复原状态：{exc}",
                parent=win,
            )
            return False

        permissions = dict(next_permissions)
        for category, value in permissions.items():
            permission_vars[category].set(value)
        status_var.set(cloud_sensitive_commands_status(permissions))
        return True

    def set_all(enabled):
        enabled = bool(enabled)
        next_permissions = {
            spec.category: enabled
            for spec in CLOUD_COMMAND_PERMISSION_SPECS
        }
        if next_permissions == permissions:
            return True
        if enabled:
            confirmed = messagebox.askyesno(
                "确认开启全部敏感权限",
                "即将允许云端发送短信、拨打电话、PIN 码操作、PUK 码操作、"
                "修改本机号码、修改 SN 码、查询基站定位数据、使用 USSD 服务、"
                "设置呼叫转移或呼叫限制、修改信息中心号码、删除设备数据以及重置或关闭设备。"
                "\n\n确认全部开启吗？",
                parent=win,
            )
            if not confirmed:
                return False
        return apply_permissions(next_permissions)

    all_buttons = ttk.Frame(frame)
    all_buttons.pack(fill=tk.X, pady=(0, 10))
    ttk.Button(
        all_buttons,
        text="开启全部",
        width=12,
        command=lambda: set_all(True),
    ).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Button(
        all_buttons,
        text="关闭全部",
        width=12,
        command=lambda: set_all(False),
    ).pack(side=tk.LEFT, padx=(0, 16))
    ttk.Label(all_buttons, textvariable=status_var).pack(side=tk.RIGHT)

    options_frame = ttk.LabelFrame(frame, text="敏感指令", padding=8)
    options_frame.pack(fill=tk.BOTH, expand=True)
    for column in range(3):
        options_frame.grid_columnconfigure(column, weight=1)

    def toggle(category):
        nonlocal permissions
        next_enabled = bool(permission_vars[category].get())
        previous_enabled = bool(permissions[category])
        if next_enabled == previous_enabled:
            return True

        spec = next(
            item for item in CLOUD_COMMAND_PERMISSION_SPECS
            if item.category == category
        )
        if next_enabled:
            confirmed = messagebox.askyesno(
                "确认开启敏感权限",
                f"即将允许云端执行：{spec.label}\n\n{spec.description}\n\n确认继续开启吗？",
                parent=win,
            )
            if not confirmed:
                permission_vars[category].set(previous_enabled)
                return False

        next_permissions = dict(permissions)
        next_permissions[category] = next_enabled
        return apply_permissions(next_permissions)

    for index, spec in enumerate(CLOUD_COMMAND_PERMISSION_SPECS):
        ttk.Checkbutton(
            options_frame,
            text=spec.label,
            variable=permission_vars[spec.category],
            command=lambda category=spec.category: toggle(category),
        ).grid(
            row=index // 3,
            column=index % 3,
            sticky="w",
            padx=(4, 24),
            pady=(8, 8),
        )

    footer = ttk.Frame(frame)
    footer.pack(fill=tk.X, pady=(10, 0))
    footer.grid_columnconfigure(0, weight=1)
    ttk.Button(footer, text="关闭", width=10, command=close).grid(
        row=0,
        column=0,
    )

    refs = {
        "window": win,
        "permission_vars": permission_vars,
        "status_var": status_var,
        "toggle": toggle,
        "set_all": set_all,
        "close": close,
    }
    _active_security_settings_refs = refs

    win.update_idletasks()
    _fit_security_window(win, frame)
    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    win.protocol("WM_DELETE_WINDOW", close)
    win.bind("<Escape>", lambda _event: close())

    return refs
