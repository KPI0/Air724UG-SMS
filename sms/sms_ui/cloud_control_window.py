import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.cloud_runtime import (
    cloud_auto_upload_action,
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

    form = CloudControlFormController(
        win,
        frame,
        state,
        state_provider,
        status_var,
        on_auto_upload_changed,
    )
    win._sync_form_from_state = form.sync_from_state

    def save_only():
        values = form.read()
        if values is not None:
            on_save(values, win)

    def connect():
        values = form.read(force_enabled=True)
        if values is not None:
            form.enabled_var.set(True)
            on_connect(values, win)

    def disconnect():
        on_disconnect(form.disconnect_values(), win)

    def close():
        on_close(win)

    build_cloud_action_buttons(frame, save_only, connect, disconnect, close)

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
