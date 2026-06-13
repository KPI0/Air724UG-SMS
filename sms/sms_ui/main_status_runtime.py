def _safe_log(log_error, message):
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def update_temperature_status_runtime(temp_str, *, tk_alive, temp_var, run_on_ui_thread, ui_post, log_error=None):
    if not tk_alive():
        return False

    def update():
        try:
            temp_var.set(f"🌡️ {temp_str} ℃")
        except Exception as exc:
            _safe_log(log_error, f"Update temperature status failed: {exc!r}")

    run_on_ui_thread(update, ui_post)
    return True


def format_signal_text(rsrp_val):
    try:
        value = int(rsrp_val)
        if value == 255:
            return "📶 未知"
        return f"📶 {value - 140} dBm"
    except Exception:
        return "📶 -- dBm"


def update_signal_status_runtime(rsrp_val, *, tk_alive, signal_var, run_on_ui_thread, ui_post, log_error=None):
    if not tk_alive():
        return False

    def update():
        try:
            signal_var.set(format_signal_text(rsrp_val))
        except Exception as exc:
            _safe_log(log_error, f"Update signal status failed: {exc!r}")
            try:
                signal_var.set("📶 -- dBm")
            except Exception as fallback_exc:
                _safe_log(log_error, f"Update signal fallback status failed: {fallback_exc!r}")

    run_on_ui_thread(update, ui_post)
    return True


def update_label_status_runtime(text, color, *, tk_alive, text_var, label, run_on_ui_thread, ui_post, log_error=None):
    if not tk_alive():
        return False

    def update():
        try:
            text_var.set(text)
            label.config(fg=color)
        except Exception as exc:
            _safe_log(log_error, f"Update label status failed: {exc!r}")

    run_on_ui_thread(update, ui_post)
    return True


def apply_sms_font_style_runtime(text_area, font_size, font_color, *, log_error=None):
    try:
        text_area.tag_config("sms", foreground=font_color, font=("微软雅黑", font_size))
        return True
    except Exception as exc:
        _safe_log(log_error, f"Apply SMS font style failed: {exc!r}")
        return False
