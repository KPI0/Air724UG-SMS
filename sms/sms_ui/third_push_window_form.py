import json
import tkinter as tk
from tkinter import messagebox, ttk

from sms_core.third_push import (
    THIRD_PUSH_CHANNELS,
    THIRD_PUSH_SETTINGS_KEYS,
    push_label,
    validate_push_settings,
)
from sms_ui.third_push_fields import THIRD_PUSH_CHANNEL_PARAM_DEFS


class ThirdPushFormController:
    def __init__(self, win, frame, state, state_provider):
        self.win = win
        self.state_provider = state_provider
        self.enabled_var = tk.IntVar(win, value=1 if state["enabled"] else 0)
        self.sms_push_var = tk.IntVar(win, value=1 if state["sms_enabled"] else 0)
        self.call_push_var = tk.IntVar(win, value=1 if state["call_enabled"] else 0)
        self.channel_vars = {}
        self.entry_vars = {
            key: tk.StringVar(win, value=state["settings"].get(key, ""))
            for key in THIRD_PUSH_SETTINGS_KEYS
        }
        self.current_channel = state["channels"][0] if state["channels"] else THIRD_PUSH_CHANNELS[0][0]
        self.custom_body_text = None
        self.channel_index = {}

        self._build_push_options(frame)
        self._build_channel_selector(frame)
        self._build_parameter_area(frame)
        self._populate_channels()

    def _build_push_options(self, frame):
        push_opts = ttk.Frame(frame)
        push_opts.grid(row=0, column=0, sticky="w", pady=(0, 8))
        ttk.Checkbutton(push_opts, text="启用三方推送", variable=self.enabled_var).pack(
            side="left",
            padx=(0, 18),
        )
        ttk.Checkbutton(push_opts, text="短信事件推送", variable=self.sms_push_var).pack(
            side="left",
            padx=(0, 18),
        )
        ttk.Checkbutton(push_opts, text="电话事件推送", variable=self.call_push_var).pack(side="left")

    def _build_channel_selector(self, frame):
        channel_box = ttk.LabelFrame(frame, text="通知通道", padding=8)
        channel_box.grid(row=1, column=0, sticky="ew", pady=(0, 8))
        for col in range(3):
            channel_box.grid_columnconfigure(col, weight=1)
        self.channel_box = channel_box

    def _build_parameter_area(self, frame):
        body_frame = ttk.Frame(frame)
        body_frame.grid(row=2, column=0, sticky="nsew", pady=(0, 10))
        body_frame.grid_columnconfigure(1, weight=1)
        body_frame.grid_rowconfigure(0, weight=1)

        list_box = ttk.LabelFrame(body_frame, text="参数页", padding=8)
        list_box.grid(row=0, column=0, sticky="nsw", padx=(0, 10))
        list_box.grid_rowconfigure(0, weight=1)

        self.channel_list = tk.Listbox(list_box, width=16, height=14, exportselection=False)
        self.channel_list.grid(row=0, column=0, sticky="ns")
        self.channel_list.bind("<<ListboxSelect>>", self._on_channel_select)

        self.param_box = ttk.LabelFrame(body_frame, text="参数", padding=10)
        self.param_box.grid(row=0, column=1, sticky="nsew")
        self.param_box.grid_columnconfigure(1, weight=1)

    def _populate_channels(self):
        for idx, (channel, label) in enumerate(THIRD_PUSH_CHANNELS):
            self.channel_index[channel] = idx
            self.channel_list.insert("end", label)
            var = tk.BooleanVar(self.win, value=channel in self.state_provider()["channels"])
            self.channel_vars[channel] = var
            ttk.Checkbutton(
                self.channel_box,
                text=label,
                variable=var,
                command=lambda ch=channel: self.select_channel(ch),
            ).grid(row=idx // 3, column=idx % 3, sticky="w", padx=(0, 8), pady=(4, 4))

    def store_custom_body_text(self):
        text = self.custom_body_text
        if text is None:
            return
        try:
            if text.winfo_exists():
                self.entry_vars["custom_post_body"].set(text.get("1.0", "end-1c"))
        except Exception:
            pass

    def _set_custom_body_text(self, value):
        text = self.custom_body_text
        if text is None:
            return
        try:
            if text.winfo_exists():
                text.delete("1.0", "end")
                text.insert("1.0", value)
        except Exception:
            pass

    def render_channel(self, channel):
        self.store_custom_body_text()
        for child in self.param_box.winfo_children():
            child.destroy()
        self.custom_body_text = None

        label = push_label(channel)
        self.param_box.configure(text=f"{label} 参数")
        spec = THIRD_PUSH_CHANNEL_PARAM_DEFS.get(channel, {})
        render_channel_fields(
            self.param_box,
            self.entry_vars,
            spec,
            lambda widget: setattr(self, "custom_body_text", widget),
        )

    def select_channel(self, channel, update_list=True):
        if channel not in self.channel_index:
            return
        self.current_channel = channel
        if update_list:
            try:
                idx = self.channel_index[channel]
                self.channel_list.selection_clear(0, "end")
                self.channel_list.selection_set(idx)
                self.channel_list.see(idx)
            except Exception:
                pass
        self.render_channel(channel)

    def _on_channel_select(self, _event=None):
        try:
            sel = self.channel_list.curselection()
            if sel:
                self.select_channel(THIRD_PUSH_CHANNELS[sel[0]][0], update_list=False)
        except Exception:
            pass

    def sync_from_state(self):
        latest = self.state_provider()
        self.enabled_var.set(1 if latest["enabled"] else 0)
        self.sms_push_var.set(1 if latest["sms_enabled"] else 0)
        self.call_push_var.set(1 if latest["call_enabled"] else 0)
        for channel, var in self.channel_vars.items():
            var.set(channel in latest["channels"])
        for key, var in self.entry_vars.items():
            var.set(latest["settings"].get(key, ""))
        self._set_custom_body_text(self.entry_vars["custom_post_body"].get())

    def collect(self, validate=True):
        self.store_custom_body_text()
        selected = [channel for channel, var in self.channel_vars.items() if var.get()]
        if validate and bool(self.enabled_var.get()) and not selected:
            messagebox.showerror("错误", "启用三方推送时，请至少选择一个通知通道。", parent=self.win)
            return None

        settings = {key: var.get().strip() for key, var in self.entry_vars.items()}
        if validate and "custom_post" in selected:
            json_error = validate_custom_post_body(settings)
            if json_error:
                self.select_channel("custom_post")
                messagebox.showerror("错误", f"自定义 POST Body 不是有效 JSON：\n{json_error}", parent=self.win)
                return None

        if validate:
            missing = validate_push_settings(selected, settings)
            if missing:
                messagebox.showerror(
                    "缺少参数",
                    "请先填写所选通道的必填参数：\n\n" + "\n".join(missing),
                    parent=self.win,
                )
                first_channel = first_channel_from_missing(selected, missing)
                if first_channel:
                    self.select_channel(first_channel)
                return None

        return (
            bool(self.enabled_var.get()),
            bool(self.sms_push_var.get()),
            bool(self.call_push_var.get()),
            selected,
            settings,
        )


def first_channel_from_missing(selected, missing):
    for item in missing:
        label = item.split(":", 1)[0]
        for channel in selected:
            if push_label(channel) == label:
                return channel
    return None


def validate_custom_post_body(settings):
    content_type = settings.get("custom_post_content_type", "")
    if "json" not in content_type.lower() or not settings.get("custom_post_body"):
        return None
    try:
        json.loads(settings["custom_post_body"])
    except Exception as exc:
        return str(exc)
    return None


def render_channel_fields(param_box, entry_vars, spec, set_custom_text_widget):
    row = 0
    tip = spec.get("tip")
    if tip:
        ttk.Label(param_box, text=tip, foreground="#666666", wraplength=460).grid(
            row=row,
            column=0,
            columnspan=2,
            sticky="w",
            pady=(0, 8),
        )
        row += 1

    for field_label, key, kind, show in spec.get("fields", ()):
        ttk.Label(param_box, text=field_label).grid(row=row, column=0, sticky="w", pady=5)
        if kind == "text":
            text = tk.Text(param_box, height=5, width=56, wrap="word")
            text.grid(row=row, column=1, sticky="ew", pady=5, padx=(8, 0))
            text.insert("1.0", entry_vars[key].get())
            text.bind(
                "<KeyRelease>",
                lambda _e, k=key, w=text: entry_vars[k].set(w.get("1.0", "end-1c")),
            )
            text.bind(
                "<FocusOut>",
                lambda _e, k=key, w=text: entry_vars[k].set(w.get("1.0", "end-1c")),
            )
            set_custom_text_widget(text)
        else:
            ttk.Entry(param_box, textvariable=entry_vars[key], width=56, show=show).grid(
                row=row,
                column=1,
                sticky="ew",
                pady=5,
                padx=(8, 0),
            )
        row += 1


def build_action_buttons(frame, save_command, test_command, close_command):
    btn_frame = ttk.Frame(frame)
    btn_frame.grid(row=3, column=0, sticky="e")
    ttk.Button(btn_frame, text="保存", width=10, command=save_command).pack(side="left", padx=(0, 8))
    ttk.Button(btn_frame, text="测试推送", width=10, command=test_command).pack(side="left", padx=(0, 8))
    ttk.Button(btn_frame, text="关闭", width=10, command=close_command).pack(side="left")
    return btn_frame
