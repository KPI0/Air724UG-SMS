from sms_core.namespace_binding import make_namespace_runtime_binder
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
    bind = make_namespace_runtime_binder(namespace, globals())

    def center_on_screen(win, w=None, h=None):
        return center_on_screen_namespace_runtime(namespace, win, w, h)

    def show_about():
        return namespace["_ui_open_about_dialog"](
            namespace["root"],
            namespace["APP_VERSION"],
            "https://github.com/KPI0/Air724UG-SMS",
            namespace["center_window"],
        )

    namespace.update({
        "center_on_screen": center_on_screen,
        "stop_tray_icon": bind(
            "stop_tray_icon_namespace_runtime",
            positional_keywords=("wait_after",),
        ),
        "show_window": bind("show_window_namespace_runtime"),
        "hide_window": bind("hide_window_namespace_runtime"),
        "create_tray": bind("create_tray_namespace_runtime"),
        "center_window": bind("center_window_namespace_runtime"),
        "show_about": show_about,
        "ui_messagebox": bind("ui_messagebox_namespace_runtime"),
        "set_temperature": bind("set_temperature_namespace_runtime"),
        "set_signal": bind("set_signal_namespace_runtime"),
        "set_cloud_status": bind("set_cloud_status_namespace_runtime"),
        "set_cloud_auth_status_from_ack": bind("set_cloud_auth_status_from_ack_namespace_runtime"),
        "set_status": bind("set_status_namespace_runtime"),
        "apply_sms_font_style": bind("apply_sms_font_style_namespace_runtime"),
        "clear_window": bind("clear_window_namespace_runtime"),
        "send_reset_cmd": bind("send_reset_cmd_namespace_runtime"),
    })
    return namespace
