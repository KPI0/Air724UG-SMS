import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.cloud_runtime import (
    CloudControlSettings,
    cloud_control_connection_action,
    cloud_control_save_kwargs,
)
from sms_ui.cloud_control_window_form import (
    CloudControlFormController,
    build_cloud_action_buttons,
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
    on_enabled_changed=None,
    open_security_settings=None,
):
    state = state_provider()
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("云端控制")
    win.minsize(540, 260)
    win.resizable(True, False)

    frame = ttk.Frame(win, padding=12)
    frame.pack(fill="both", expand=True)
    frame.grid_columnconfigure(1, weight=1)

    form = CloudControlFormController(
        win,
        frame,
        state,
        state_provider,
        status_var,
        on_auto_upload_changed,
        on_enabled_changed,
        open_security_settings,
    )
    win._sync_form_from_state = form.refresh_synced_fields
    win._refresh_synced_fields = form.refresh_synced_fields

    def save_only():
        values = form.read()
        if values is not None:
            if on_save(values, win) is not False:
                form.refresh_after_save()

    def connect():
        values = form.read(force_enabled=True)
        if values is not None:
            form.enabled_var.set(True)
            if on_connect(values, win) is False:
                form.restore_toggles()
            else:
                form.refresh_after_save()

    def disconnect():
        if on_disconnect(form.disconnect_values(), win) is False:
            form.restore_toggles()
        else:
            form.refresh_after_save(("enabled",))

    def close():
        on_close(win)

    build_cloud_action_buttons(frame, save_only, connect, disconnect, close)

    win.protocol("WM_DELETE_WINDOW", close)
    win.bind("<Escape>", lambda _e: close())
    win.update_idletasks()
    win.minsize(max(540, win.winfo_reqwidth()), max(260, win.winfo_reqheight()))
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
    open_security_settings=None,
):
    refresh_settings()

    if sync_existing_window(current_window, "_sync_form_from_state"):
        return current_window

    def settings_from_state():
        state = state_provider()
        return CloudControlSettings(
            enabled=bool(state["enabled"]), auto_upload=bool(state["auto_upload"]),
            url=state["url"], device_secret=state["secret"],
            reconnect_interval=state["reconnect_interval"],
        )

    def save_and_update_connection(kwargs, *, reconnect=False, explicit_fields=()):
        previous = settings_from_state()
        kwargs = {
            key: value for key, value in kwargs.items()
            if key in explicit_fields or value != getattr(previous, key)
        }
        saved = save_setting(**kwargs)
        if saved in (None, False):
            return False
        if not isinstance(saved, CloudControlSettings):
            # Retain compatibility with callbacks that return only success.
            saved = CloudControlSettings(**{**vars(previous), **kwargs})
        action = cloud_control_connection_action(previous, saved)
        if reconnect:
            action = "restart" if saved.enabled else "stop"
        try:
            if action == "restart":
                restart_control(show_errors=True)
            elif action == "stop":
                stop_control()
            elif action == "register":
                if all(connection_state_provider()):
                    register_current()
            elif action == "unregister":
                schedule_unregister("auto_upload_disabled")
        except Exception as exc:
            cloud_log(f"云端配置已保存，连接状态更新失败（{type(exc).__name__}），请重新连接", show_main=True)
        return True

    def on_auto_upload_changed(auto_upload, _win):
        return save_and_update_connection({"auto_upload": auto_upload}, explicit_fields=("auto_upload",))

    def on_enabled_changed(enabled, values, _win):
        if enabled:
            kwargs = cloud_control_save_kwargs(values, enabled_override=True)
        else:
            kwargs = {"enabled": False}

        if not save_and_update_connection(kwargs, reconnect=True, explicit_fields=("enabled",)):
            return False

        if enabled:
            cloud_log("云端控制已启用")
        else:
            cloud_log("云端控制已关闭")
        return True

    def save_values(values, win):
        kwargs = cloud_control_save_kwargs(values)
        if not save_and_update_connection(kwargs, reconnect=True):
            messagebox.showerror("保存失败", "云端控制配置保存失败，请检查配置文件是否可写。", parent=win)
            return False
        messagebox.showinfo("配置已保存", "云端控制配置已成功保存！", parent=win)
        cloud_log("配置已保存")
        return True

    def connect_values(values, _win):
        return save_and_update_connection(
            cloud_control_save_kwargs(values, enabled_override=True), reconnect=True,
            explicit_fields=("enabled",),
        )

    def disconnect_values(values, _win):
        if not save_and_update_connection({"enabled": False}, reconnect=True, explicit_fields=("enabled",)):
            return False
        cloud_log("已手动断开")
        return True

    def close_window(win):
        try:
            win.destroy()
        except Exception:
            pass
        set_window(None)

    dialog_kwargs = {}
    if on_enabled_changed is not None:
        dialog_kwargs["on_enabled_changed"] = on_enabled_changed
    if open_security_settings is not None:
        dialog_kwargs["open_security_settings"] = open_security_settings
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
        **dialog_kwargs,
    )
    set_window(win)
    return win
