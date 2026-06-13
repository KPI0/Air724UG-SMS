from sms_core.namespace_binding import make_namespace_runtime_binder
from sms_ui.app_infrastructure_namespace_runtime import (
    check_single_instance_namespace_runtime,
    lock_port_mutex_namespace_runtime,
    safe_close_serial_namespace_runtime,
    safe_save_config_namespace_runtime,
    show_sms_popup_namespace_runtime,
    unlock_port_mutex_namespace_runtime,
)
from sms_ui.repeat_notice_runtime import emit_repeat_notice


def install_app_infrastructure_namespace_bindings(namespace):
    bind = make_namespace_runtime_binder(namespace, globals())

    def auto_connect_ui(msg):
        return emit_repeat_notice(namespace["_auto_connect_notice"], msg, namespace["system_ui"])

    namespace.update({
        "safe_save_config": bind("safe_save_config_namespace_runtime"),
        "safe_close_serial": bind("safe_close_serial_namespace_runtime"),
        "auto_connect_ui": auto_connect_ui,
        "lock_port_mutex": bind("lock_port_mutex_namespace_runtime"),
        "unlock_port_mutex": bind("unlock_port_mutex_namespace_runtime"),
        "check_single_instance": bind("check_single_instance_namespace_runtime"),
        "show_sms_popup": bind("show_sms_popup_namespace_runtime"),
    })
    return namespace
