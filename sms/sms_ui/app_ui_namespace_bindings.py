from sms_ui.app_ui_namespace_runtime import (
    apply_sms_font_style_namespace_runtime,
    center_on_screen_namespace_runtime,
    center_window_namespace_runtime,
    clear_window_namespace_runtime,
    create_tray_namespace_runtime,
    hide_window_namespace_runtime,
    send_reset_cmd_namespace_runtime,
    set_cloud_auth_status_from_ack_namespace_runtime,
    set_cloud_status_namespace_runtime,
    set_signal_namespace_runtime,
    set_status_namespace_runtime,
    set_temperature_namespace_runtime,
    show_window_namespace_runtime,
    stop_tray_icon_namespace_runtime,
    ui_messagebox_namespace_runtime,
)


def install_app_ui_namespace_bindings(namespace):
    def center_on_screen(win, w=None, h=None):
        return center_on_screen_namespace_runtime(namespace, win, w, h)

    def stop_tray_icon(wait_after=0.45):
        return stop_tray_icon_namespace_runtime(namespace, wait_after=wait_after)

    def show_window():
        return show_window_namespace_runtime(namespace)

    def hide_window():
        return hide_window_namespace_runtime(namespace)

    def create_tray():
        return create_tray_namespace_runtime(namespace)

    def center_window(win, parent):
        return center_window_namespace_runtime(namespace, win, parent)

    def show_about():
        return namespace["_ui_open_about_dialog"](
            namespace["root"],
            namespace["APP_VERSION"],
            "https://github.com/KPI0/Air724UG-SMS",
            namespace["center_window"],
        )

    def ui_messagebox(kind, title, message):
        return ui_messagebox_namespace_runtime(namespace, kind, title, message)

    def set_temperature(temp_str):
        return set_temperature_namespace_runtime(namespace, temp_str)

    def set_signal(rsrp_val):
        return set_signal_namespace_runtime(namespace, rsrp_val)

    def set_cloud_status(text, color="#666666"):
        return set_cloud_status_namespace_runtime(namespace, text, color)

    def set_cloud_auth_status_from_ack(data):
        return set_cloud_auth_status_from_ack_namespace_runtime(namespace, data)

    def set_status(text, color="black"):
        return set_status_namespace_runtime(namespace, text, color)

    def apply_sms_font_style():
        return apply_sms_font_style_namespace_runtime(namespace)

    def clear_window():
        return clear_window_namespace_runtime(namespace)

    def send_reset_cmd():
        return send_reset_cmd_namespace_runtime(namespace)

    namespace.update({
        "center_on_screen": center_on_screen,
        "stop_tray_icon": stop_tray_icon,
        "show_window": show_window,
        "hide_window": hide_window,
        "create_tray": create_tray,
        "center_window": center_window,
        "show_about": show_about,
        "ui_messagebox": ui_messagebox,
        "set_temperature": set_temperature,
        "set_signal": set_signal,
        "set_cloud_status": set_cloud_status,
        "set_cloud_auth_status_from_ack": set_cloud_auth_status_from_ack,
        "set_status": set_status,
        "apply_sms_font_style": apply_sms_font_style,
        "clear_window": clear_window,
        "send_reset_cmd": send_reset_cmd,
    })
    return namespace
