from sms_ui.device_reset_runtime import send_reset_command_runtime
from sms_ui.main_status_runtime import (
    apply_sms_font_style_runtime,
    update_label_status_runtime,
    update_signal_status_runtime,
    update_temperature_status_runtime,
)
from sms_ui.thread_runtime import ui_messagebox_runtime
from sms_ui.tray_runtime import create_tray_icon_runtime, stop_tray_icon_runtime
from sms_ui.ui_log_runtime import clear_text_widget_runtime


def stop_tray_icon_namespace_runtime(namespace, *, wait_after=0.45):
    return stop_tray_icon_runtime(
        tray_icon=namespace["tray_icon"],
        clear_tray_icon=lambda: namespace.__setitem__("tray_icon", None),
        wait_after=wait_after,
    )


def show_window_namespace_runtime(namespace):
    def do_show():
        try:
            namespace["root"].deiconify()
            namespace["root"].lift()
            namespace["root"].focus_force()
        except Exception:
            pass

    return namespace["run_on_ui_thread"](do_show, namespace["ui_post"])


def hide_window_namespace_runtime(namespace):
    def do_hide():
        try:
            namespace["root"].withdraw()
        except Exception:
            pass

    return namespace["run_on_ui_thread"](do_hide, namespace["ui_post"])


def create_tray_namespace_runtime(namespace):
    return create_tray_icon_runtime(
        icon_path=namespace["resource_path"]("icon.ico"),
        title=namespace["APP_WINDOW_TITLE"],
        show_window=namespace["show_window"],
        hide_window=namespace["hide_window"],
        cleanup_and_exit=namespace["cleanup_and_exit"],
        set_tray_icon=lambda icon: namespace.__setitem__("tray_icon", icon),
    )


def center_on_screen_namespace_runtime(namespace, win, width=None, height=None):
    win.update_idletasks()
    if width is None or height is None:
        width = win.winfo_reqwidth()
        height = win.winfo_reqheight()

    screen_width = win.winfo_screenwidth()
    screen_height = win.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    win.geometry(f"{width}x{height}+{x}+{y}")


def center_window_namespace_runtime(namespace, win, parent):
    win.update_idletasks()
    width = win.winfo_width()
    height = win.winfo_height()
    if width <= 1 or height <= 1:
        width = win.winfo_reqwidth()
        height = win.winfo_reqheight()

    x = parent.winfo_rootx() + (parent.winfo_width() - width) // 2
    y = parent.winfo_rooty() + (parent.winfo_height() - height) // 2
    win.geometry(f"+{x}+{y}")


def ui_messagebox_namespace_runtime(namespace, kind, title, message):
    return ui_messagebox_runtime(
        kind,
        title,
        message,
        messagebox=namespace["messagebox"],
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
    )


def set_temperature_namespace_runtime(namespace, temp_str):
    return update_temperature_status_runtime(
        temp_str,
        tk_alive=namespace["tk_alive"],
        temp_var=namespace["temp_var"],
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
    )


def set_signal_namespace_runtime(namespace, rsrp_val):
    return update_signal_status_runtime(
        rsrp_val,
        tk_alive=namespace["tk_alive"],
        signal_var=namespace["signal_var"],
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
    )


def set_cloud_status_namespace_runtime(namespace, text, color="#666666"):
    return update_label_status_runtime(
        text,
        color,
        tk_alive=namespace["tk_alive"],
        text_var=namespace["cloud_var"],
        label=namespace["cloud_label"],
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
    )


def set_cloud_auth_status_from_ack_namespace_runtime(namespace, data):
    status = namespace["_cloud_auth_status_from_ack"](data)
    if status == "authorized":
        return namespace["set_cloud_status"]("🌐 已授权", "#008000")
    if status == "failed":
        return namespace["set_cloud_status"]("🌐 授权失败", "#cc0000")
    return namespace["set_cloud_status"]("🌐 等待授权", "#b26a00")


def set_status_namespace_runtime(namespace, text, color="black"):
    return update_label_status_runtime(
        text,
        color,
        tk_alive=namespace["tk_alive"],
        text_var=namespace["status_var"],
        label=namespace["status_label"],
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
    )


def apply_sms_font_style_namespace_runtime(namespace):
    return apply_sms_font_style_runtime(
        namespace["text_area"],
        namespace["SMS_FONT_SIZE"],
        namespace["SMS_FONT_COLOR"],
    )


def clear_window_namespace_runtime(namespace):
    def do_clear():
        return clear_text_widget_runtime(namespace.get("text_area"), end=namespace["tk"].END)

    return namespace["run_on_ui_thread"](do_clear, namespace["ui_post"])


def send_reset_cmd_namespace_runtime(namespace):
    return send_reset_command_runtime(
        confirm_reset=lambda: namespace["messagebox"].askyesno(
            "重启硬件",
            "确定要重启底层通信模组吗？\n\n(设备重启期间将会短暂断开连接，随后自动重连)",
            parent=namespace["root"],
        ),
        send_command_async=namespace["send_command_with_result_async"],
        serial_lock=namespace["serial_lock"],
        get_serial=lambda: namespace["serial_obj"],
        ui_post=namespace["ui_post"],
        system_ui=namespace["system_ui"],
        show_warning=lambda title, message: namespace["messagebox"].showwarning(
            title,
            message,
            parent=namespace["root"],
        ),
        log_error=namespace.get("log_file_only"),
    )
