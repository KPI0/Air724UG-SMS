import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.third_push import (
    THIRD_PUSH_TEST_MESSAGE,
    third_push_save_kwargs,
    third_push_saved_status,
)
from sms_ui.third_push_window_form import ThirdPushFormController, build_action_buttons


def confirm_close_with_unsaved_changes(form, win, confirm=messagebox.askyesno):
    form.store_custom_body_text()
    if not form.is_dirty():
        return True
    return bool(confirm(
        "未保存修改",
        "当前三方推送配置尚未保存，是否放弃这些修改并关闭窗口？",
        parent=win,
    ))


def open_third_push_window_dialog(
    parent,
    state_provider,
    on_save,
    on_test,
    on_close,
    center_window,
    on_option_changed=None,
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

    status_var = tk.StringVar(win, value="")
    action_refs = {}

    def dirty_changed(dirty):
        win.title("三方推送 *" if dirty else "三方推送")
        save_button = action_refs.get("save_button")
        if save_button is not None:
            save_button.config(text="保存修改" if dirty else "保存")
        if dirty:
            status_var.set("● 有未保存修改")
        elif status_var.get() == "● 有未保存修改":
            status_var.set("")

    def option_changed(option, value, values, option_win):
        if on_option_changed is None:
            return False
        result = on_option_changed(option, value, values, option_win)
        if result is not False:
            labels = {
                "enabled": "三方推送",
                "sms_enabled": "短信事件推送",
                "call_enabled": "电话事件推送",
            }
            status_var.set(f"✅ {labels.get(option, option)}已{'开启' if value else '关闭'}")
        return result

    form = ThirdPushFormController(
        win,
        frame,
        state,
        state_provider,
        on_option_changed=option_changed,
        on_dirty_changed=dirty_changed,
    )
    win._sync_form_from_globals = form.sync_from_state_if_clean
    win._force_sync_form_from_globals = form.sync_from_state

    def save_only():
        values = form.collect(validate=True)
        if values is not None:
            if on_save(values, win) is not False:
                form.mark_all_saved()
                status_var.set("✅ 配置已保存")

    def test_push():
        values = form.collect(validate=True)
        if values is None:
            return
        if not values[3]:
            messagebox.showwarning("提示", "请先选择至少一个通知通道。", parent=win)
            return
        result = on_test(values, win)
        if result is not False:
            form.mark_all_saved()
            status_var.set("✅ 配置已保存，测试已加入队列" if result == "queued" else "✅ 配置已保存")

    def close():
        if not confirm_close_with_unsaved_changes(form, win):
            return
        on_close(win)

    action_refs.update(build_action_buttons(frame, save_only, test_push, close, status_var=status_var))

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
        if save_setting(**kwargs) is None:
            messagebox.showerror("保存失败", "三方推送配置保存失败，请检查配置文件是否可写。", parent=win)
            return False
        system_ui(third_push_saved_status(kwargs["enabled"], kwargs["notify_type"]), "normal")
        return True

    def option_values(option, value, values, win):
        if option == "enabled" and value:
            kwargs = third_push_save_kwargs(values)
        else:
            kwargs = {option: bool(value)}
        if save_setting(**kwargs) is None:
            messagebox.showerror("保存失败", "三方推送配置保存失败，请检查配置文件是否可写。", parent=win)
            try:
                sync_form = getattr(win, "_force_sync_form_from_globals", None)
                if sync_form is None:
                    sync_form = win._sync_form_from_globals
                sync_form()
            except Exception:
                pass
            return False
        labels = {
            "enabled": "三方推送",
            "sms_enabled": "短信事件推送",
            "call_enabled": "电话事件推送",
        }
        system_ui(f"📡 {labels.get(option, option)}已{'开启' if value else '关闭'}", "normal")
        return True

    def test_values(values, win):
        kwargs = third_push_save_kwargs(values)
        if save_setting(**kwargs) is None:
            messagebox.showerror("保存失败", "三方推送配置保存失败，请检查配置文件是否可写。", parent=win)
            return False
        queued = enqueue_push(
            THIRD_PUSH_TEST_MESSAGE,
            show_success=True,
            show_result=True,
            channels=kwargs["notify_type"],
            settings=kwargs["settings"],
        )
        if queued:
            system_ui("📡 三方推送配置已保存，测试已加入队列", "normal")
            return "queued"
        else:
            messagebox.showerror("测试推送失败", "三方推送队列已满，测试未发送。", parent=win)
            return True

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
        on_option_changed=option_values,
    )
    set_window(win)
    return win
