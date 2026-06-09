from sms_ui.serial_debug_panel import append_serial_debug_lines_once, reset_serial_debug_window_state


class SerialDebugPauseController:
    def __init__(
        self,
        enabled_var,
        paused_var,
        state_label,
        pause_button,
        bypass_checkbox,
        send_button,
        send_entry,
        quick_button,
        quick_scroll_frame,
        text_widget,
    ):
        self.enabled_var = enabled_var
        self.paused_var = paused_var
        self.state_label = state_label
        self.pause_button = pause_button
        self.bypass_checkbox = bypass_checkbox
        self.send_button = send_button
        self.send_entry = send_entry
        self.quick_button = quick_button
        self.quick_scroll_frame = quick_scroll_frame
        self.text_widget = text_widget
        self.pause_banner_shown = False

    def refresh(self):
        running = bool(self.enabled_var.get())
        if not running:
            self.state_label.config(text="○ 未运行")
        else:
            self.state_label.config(text="⏸ 已暂停显示" if self.paused_var.get() else "● 运行中")

        try:
            self.pause_button.state(["!disabled"] if running else ["disabled"])
        except Exception:
            pass

        try:
            state_val = ["!disabled"] if running else ["disabled"]
            self.send_button.state(state_val)
            self.send_entry.state(state_val)
            self.quick_button.state(state_val)

            for child in self.quick_scroll_frame.winfo_children():
                try:
                    child.state(state_val)
                except Exception:
                    pass
        except Exception:
            pass

    def set_paused(self, is_paused):
        is_paused = bool(is_paused)
        if self.paused_var.get() == is_paused:
            return

        self.paused_var.set(is_paused)
        if is_paused:
            self.pause_button.config(text="▶ 继续")
            try:
                self.bypass_checkbox.state(["!disabled"])
            except Exception:
                pass

            if not self.pause_banner_shown:
                self.pause_banner_shown = True
                self._append_banner("\n—— 已暂停显示（串口仍在采集，旁路开关已锁定）——\n")
        else:
            self.pause_button.config(text="⏸ 暂停")
            try:
                self.bypass_checkbox.state(["disabled"])
            except Exception:
                pass
            self._append_banner("\n—— 已继续显示（串口仍在采集，旁路开关已锁定）——\n")
            self.pause_banner_shown = False

        self.refresh()

    def toggle(self):
        self.set_paused(not self.paused_var.get())

    def reset(self):
        try:
            self.paused_var.set(False)
        except Exception:
            pass
        try:
            self.pause_button.config(text="⏸ 暂停")
        except Exception:
            pass
        self.pause_banner_shown = False

    def _append_banner(self, text):
        try:
            self.text_widget.config(state="normal")
            self.text_widget.insert("end", text)
            self.text_widget.see("end")
            self.text_widget.config(state="disabled")
        except Exception:
            pass


def start_serial_debug_append_loop(
    window,
    get_text_widget,
    serial_queue,
    all_lines,
    paused_var,
    drop_label,
    filter_var,
    finder,
    max_store_lines,
    max_visible_lines,
    get_drop_count,
    interval_ms=100,
):
    def append_lines():
        text_widget = get_text_widget()
        if text_widget is None or not text_widget.winfo_exists():
            return

        drop_count = get_drop_count()
        if paused_var.get():
            _update_drop_label(drop_label, drop_count)
            try:
                window.after(interval_ms, append_lines)
            except Exception:
                return
            return

        append_serial_debug_lines_once(
            text_widget,
            serial_queue,
            all_lines,
            filter_var.get().strip(),
            finder,
            max_store_lines,
            max_visible_lines,
        )
        _update_drop_label(drop_label, drop_count)

        try:
            window.after(interval_ms, append_lines)
        except Exception:
            return

    append_lines()
    return append_lines


def close_serial_debug_runtime(
    window,
    text_widget,
    serial_queue,
    all_lines,
    finder,
    enabled_var,
    pause_controller,
    drop_label,
    set_enabled,
    set_drop_count,
    clear_window_refs,
):
    set_enabled(False)

    try:
        enabled_var.set(False)
    except Exception:
        pass

    try:
        pause_controller.reset()
    except Exception:
        pass

    reset_serial_debug_window_state(
        window,
        text_widget,
        serial_queue,
        all_lines,
        finder,
    )

    set_drop_count(0)
    try:
        drop_label.config(text="")
    except Exception:
        pass

    try:
        if window is not None and window.winfo_exists():
            window.destroy()
    finally:
        clear_window_refs()


def _update_drop_label(drop_label, drop_count):
    if drop_count > 0:
        drop_label.config(text=f"队列满丢弃：{drop_count} 行")
    else:
        drop_label.config(text="")
