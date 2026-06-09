import json
import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.third_push import (
    THIRD_PUSH_CHANNELS,
    THIRD_PUSH_TEST_MESSAGE,
    THIRD_PUSH_SETTINGS_KEYS,
    push_label,
    third_push_save_kwargs,
    third_push_saved_status,
    validate_push_settings,
)
from sms_ui.third_push_fields import THIRD_PUSH_CHANNEL_PARAM_DEFS


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
    win.geometry("780x640")
    win.resizable(False, False)

    frame = ttk.Frame(win, padding=12)
    frame.pack(fill="both", expand=True)
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_rowconfigure(2, weight=1)

    enabled_var = tk.IntVar(win, value=1 if state["enabled"] else 0)
    sms_push_var = tk.IntVar(win, value=1 if state["sms_enabled"] else 0)
    call_push_var = tk.IntVar(win, value=1 if state["call_enabled"] else 0)
    channel_vars = {}
    entry_vars = {
        key: tk.StringVar(win, value=state["settings"].get(key, ""))
        for key in THIRD_PUSH_SETTINGS_KEYS
    }
    current_channel = {
        "value": state["channels"][0] if state["channels"] else THIRD_PUSH_CHANNELS[0][0]
    }
    custom_body_text = {"widget": None}

    push_opts = ttk.Frame(frame)
    push_opts.grid(row=0, column=0, sticky="w", pady=(0, 8))
    ttk.Checkbutton(push_opts, text="启用三方推送", variable=enabled_var).pack(side="left", padx=(0, 18))
    ttk.Checkbutton(push_opts, text="短信事件推送", variable=sms_push_var).pack(side="left", padx=(0, 18))
    ttk.Checkbutton(push_opts, text="电话事件推送", variable=call_push_var).pack(side="left")

    channel_box = ttk.LabelFrame(frame, text="通知通道", padding=8)
    channel_box.grid(row=1, column=0, sticky="ew", pady=(0, 8))
    for col in range(3):
        channel_box.grid_columnconfigure(col, weight=1)

    body_frame = ttk.Frame(frame)
    body_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
    body_frame.grid_columnconfigure(1, weight=1)
    body_frame.grid_rowconfigure(0, weight=1)

    list_box = ttk.LabelFrame(body_frame, text="参数页", padding=8)
    list_box.grid(row=0, column=0, sticky="nsw", padx=(0, 10))
    list_box.grid_rowconfigure(0, weight=1)

    channel_list = tk.Listbox(list_box, width=16, height=14, exportselection=False)
    channel_list.grid(row=0, column=0, sticky="ns")

    param_box = ttk.LabelFrame(body_frame, text="参数", padding=10)
    param_box.grid(row=0, column=1, sticky="nsew")
    param_box.grid_columnconfigure(1, weight=1)

    channel_index = {}

    def store_custom_body_text():
        text = custom_body_text["widget"]
        if text is None:
            return
        try:
            if text.winfo_exists():
                entry_vars["custom_post_body"].set(text.get("1.0", "end-1c"))
        except Exception:
            pass

    def set_custom_body_text(value):
        text = custom_body_text["widget"]
        if text is None:
            return
        try:
            if text.winfo_exists():
                text.delete("1.0", "end")
                text.insert("1.0", value)
        except Exception:
            pass

    def render_channel(channel):
        store_custom_body_text()
        for child in param_box.winfo_children():
            child.destroy()
        custom_body_text["widget"] = None

        label = push_label(channel)
        param_box.configure(text=f"{label} 参数")
        spec = THIRD_PUSH_CHANNEL_PARAM_DEFS.get(channel, {})
        row = 0

        tip = spec.get("tip")
        if tip:
            ttk.Label(param_box, text=tip, foreground="#666666", wraplength=460).grid(
                row=row, column=0, columnspan=2, sticky="w", pady=(0, 8)
            )
            row += 1

        for field_label, key, kind, show in spec.get("fields", ()):
            ttk.Label(param_box, text=field_label).grid(row=row, column=0, sticky="w", pady=5)
            if kind == "text":
                text = tk.Text(param_box, height=5, width=56, wrap="word")
                text.grid(row=row, column=1, sticky="ew", pady=5, padx=(8, 0))
                text.insert("1.0", entry_vars[key].get())
                text.bind("<KeyRelease>", lambda _e, k=key, w=text: entry_vars[k].set(w.get("1.0", "end-1c")))
                text.bind("<FocusOut>", lambda _e, k=key, w=text: entry_vars[k].set(w.get("1.0", "end-1c")))
                custom_body_text["widget"] = text
            else:
                ttk.Entry(param_box, textvariable=entry_vars[key], width=56, show=show).grid(
                    row=row, column=1, sticky="ew", pady=5, padx=(8, 0)
                )
            row += 1

    def select_channel(channel, update_list=True):
        if channel not in channel_index:
            return
        current_channel["value"] = channel
        if update_list:
            try:
                idx = channel_index[channel]
                channel_list.selection_clear(0, "end")
                channel_list.selection_set(idx)
                channel_list.see(idx)
            except Exception:
                pass
        render_channel(channel)

    for idx, (channel, label) in enumerate(THIRD_PUSH_CHANNELS):
        channel_index[channel] = idx
        channel_list.insert("end", label)
        var = tk.BooleanVar(win, value=channel in state["channels"])
        channel_vars[channel] = var
        ttk.Checkbutton(
            channel_box,
            text=label,
            variable=var,
            command=lambda ch=channel: select_channel(ch),
        ).grid(row=idx // 3, column=idx % 3, sticky="w", padx=(0, 8), pady=(4, 4))

    def on_channel_select(_event=None):
        try:
            sel = channel_list.curselection()
            if sel:
                select_channel(THIRD_PUSH_CHANNELS[sel[0]][0], update_list=False)
        except Exception:
            pass

    channel_list.bind("<<ListboxSelect>>", on_channel_select)

    def sync_form_from_state():
        latest = state_provider()
        enabled_var.set(1 if latest["enabled"] else 0)
        sms_push_var.set(1 if latest["sms_enabled"] else 0)
        call_push_var.set(1 if latest["call_enabled"] else 0)
        for channel, var in channel_vars.items():
            var.set(channel in latest["channels"])
        for key, var in entry_vars.items():
            var.set(latest["settings"].get(key, ""))
        set_custom_body_text(entry_vars["custom_post_body"].get())

    win._sync_form_from_globals = sync_form_from_state

    def focus_channel_params(channel):
        select_channel(channel)

    def collect_form(validate=True):
        store_custom_body_text()
        selected = [channel for channel, var in channel_vars.items() if var.get()]
        if validate and bool(enabled_var.get()) and not selected:
            messagebox.showerror("错误", "启用三方推送时，请至少选择一个通知通道。", parent=win)
            return None

        settings = {key: var.get().strip() for key, var in entry_vars.items()}
        if validate and "custom_post" in selected:
            content_type = settings.get("custom_post_content_type", "")
            if "json" in content_type.lower() and settings.get("custom_post_body"):
                try:
                    json.loads(settings["custom_post_body"])
                except Exception as exc:
                    focus_channel_params("custom_post")
                    messagebox.showerror("错误", f"自定义 POST Body 不是有效 JSON：\n{exc}", parent=win)
                    return None

        if validate:
            missing = validate_push_settings(selected, settings)
            if missing:
                messagebox.showerror(
                    "缺少参数",
                    "请先填写所选通道的必填参数：\n\n" + "\n".join(missing),
                    parent=win,
                )
                first_channel = _first_channel_from_missing(selected, missing)
                if first_channel:
                    focus_channel_params(first_channel)
                return None

        return (
            bool(enabled_var.get()),
            bool(sms_push_var.get()),
            bool(call_push_var.get()),
            selected,
            settings,
        )

    def save_only():
        values = collect_form(validate=True)
        if values is not None:
            on_save(values, win)

    def test_push():
        values = collect_form(validate=True)
        if values is None:
            return
        if not values[3]:
            messagebox.showwarning("提示", "请先选择至少一个通知通道。", parent=win)
            return
        on_test(values, win)

    def close():
        store_custom_body_text()
        on_close(win)

    btn_frame = ttk.Frame(frame)
    btn_frame.grid(row=3, column=0, sticky="e")
    ttk.Button(btn_frame, text="保存", width=10, command=save_only).pack(side="left", padx=(0, 8))
    ttk.Button(btn_frame, text="测试推送", width=10, command=test_push).pack(side="left", padx=(0, 8))
    ttk.Button(btn_frame, text="关闭", width=10, command=close).pack(side="left")

    win.protocol("WM_DELETE_WINDOW", close)
    win.bind("<Escape>", lambda _e: close())
    select_channel(current_channel["value"])
    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    return win


def _first_channel_from_missing(selected, missing):
    for item in missing:
        label = item.split(":", 1)[0]
        for channel in selected:
            if push_label(channel) == label:
                return channel
    return None


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
