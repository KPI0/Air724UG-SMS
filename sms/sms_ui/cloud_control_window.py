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
    on_enabled_changed=None,
):
    state = state_provider()
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("云端控制")
    win.minsize(540, 260)
    win.resizable(True, False)
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
        on_enabled_changed,
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
            if on_connect(values, win) is False:
                form.sync_from_state()

    def disconnect():
        if on_disconnect(form.disconnect_values(), win) is False:
            form.sync_from_state()

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
        if save_setting(auto_upload=auto_upload) in (None, False):
            # The checkbox is toggled by Tk before the transactional save
            # completes.  Keep the form and any dependent cloud state aligned
            # with the last persisted settings when the save is rejected.
            try:
                _win._sync_form_from_state()
            except Exception:
                pass
            return False
        try:
            connected, has_loop, has_conn = connection_state_provider()
            action = cloud_auto_upload_action(was_public, auto_upload, connected, has_loop, has_conn)
            if action == "register":
                register_current()
            elif action == "unregister":
                schedule_unregister("auto_upload_disabled")
        except Exception:
            pass
        return True

    def on_enabled_changed(enabled, values, _win):
        if enabled:
            kwargs = cloud_control_save_kwargs(values, enabled_override=True)
        else:
            kwargs = {"enabled": False}

        if save_setting(**kwargs) in (None, False):
            try:
                _win._sync_form_from_state()
            except Exception:
                pass
            return False

        if enabled:
            restart_control(show_errors=True)
            cloud_log("云端控制已启用")
        else:
            stop_control()
            cloud_log("云端控制已关闭")
        return True

    def save_values(values, win):
        kwargs = cloud_control_save_kwargs(values)
        if save_setting(**kwargs) in (None, False):
            messagebox.showerror("保存失败", "云端控制配置保存失败，请检查配置文件是否可写。", parent=win)
            try:
                win._sync_form_from_state()
            except Exception:
                pass
            return False
        messagebox.showinfo("配置已保存", "云端控制配置已成功保存！", parent=win)
        if kwargs["enabled"]:
            restart_control(show_errors=True)
        else:
            stop_control()
        cloud_log("配置已保存")
        return True

    def connect_values(values, _win):
        if save_setting(**cloud_control_save_kwargs(values, enabled_override=True)) in (None, False):
            return False
        restart_control(show_errors=True)
        return True

    def disconnect_values(values, _win):
        if save_setting(**cloud_control_save_kwargs(values, enabled_override=False)) in (None, False):
            return False
        stop_control()
        cloud_log("已手动断开")
        return True

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
        on_enabled_changed=on_enabled_changed,
    )
    set_window(win)
    return win
