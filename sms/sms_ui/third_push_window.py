import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.third_push import (
    THIRD_PUSH_TEST_MESSAGE,
    third_push_save_kwargs,
    third_push_saved_status,
)
from sms_ui.third_push_window_form import ThirdPushFormController, build_action_buttons


def open_third_push_window_dialog(
    parent,
    state_provider,
    on_save,
    on_test,
    on_close,
    center_window,
):
    state = state_provider()
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("三方推送")
    win.geometry("780x580")
    win.minsize(680, 520)
    win.resizable(True, True)

    frame = ttk.Frame(win, padding=12)
    frame.pack(fill="both", expand=True)
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(2, weight=1)

    form = ThirdPushFormController(win, frame, state, state_provider)
    win._sync_form_from_globals = form.sync_from_state

    def save_only():
        values = form.collect(validate=True)
        if values is not None:
            on_save(values, win)

    def test_push():
        values = form.collect(validate=True)
        if values is None:
            return
        if not values[3]:
            messagebox.showwarning("提示", "请先选择至少一个通知通道。", parent=win)
            return
        on_test(values, win)

    def close():
        form.store_custom_body_text()
        on_close(win)

    build_action_buttons(frame, save_only, test_push, close)

    win.protocol("WM_DELETE_WINDOW", close)
    win.bind("<Escape>", lambda _e: close())
    form.select_channel(form.current_channel)
    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    return win


def open_third_push_window_runtime(
    parent,
    current_window,
    state_provider,
    refresh_settings,
    save_setting,
    enqueue_push,
    system_ui,
    sync_existing_window,
    set_window,
    center_window,
):
    refresh_settings()

    if sync_existing_window(current_window, "_sync_form_from_globals"):
        return current_window

    def save_values(values, win):
        kwargs = third_push_save_kwargs(values)
        save_setting(**kwargs)
        messagebox.showinfo("配置已保存", "三方推送配置已成功保存！", parent=win)
        system_ui(third_push_saved_status(kwargs["enabled"], kwargs["notify_type"]), "normal")

    def test_values(values, win):
        kwargs = third_push_save_kwargs(values)
        save_setting(**kwargs)
        queued = enqueue_push(
            THIRD_PUSH_TEST_MESSAGE,
            show_success=True,
            show_result=True,
            channels=kwargs["notify_type"],
            settings=kwargs["settings"],
        )
        if queued:
            system_ui("📡 三方推送配置已保存，测试已加入队列", "normal")
        else:
            messagebox.showerror("测试推送失败", "三方推送队列已满，测试未发送。", parent=win)

    def close_window(win):
        try:
            win.destroy()
        except Exception:
            pass
        set_window(None)

    win = open_third_push_window_dialog(
        parent,
        state_provider,
        save_values,
        test_values,
        close_window,
        center_window,
    )
    set_window(win)
    return win
