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
    def safe_save_config():
        return safe_save_config_namespace_runtime(namespace)

    def safe_close_serial():
        return safe_close_serial_namespace_runtime(namespace)

    def auto_connect_ui(msg):
        return emit_repeat_notice(namespace["_auto_connect_notice"], msg, namespace["system_ui"])

    def lock_port_mutex(port_name):
        return lock_port_mutex_namespace_runtime(namespace, port_name)

    def unlock_port_mutex():
        return unlock_port_mutex_namespace_runtime(namespace)

    def check_single_instance():
        return check_single_instance_namespace_runtime(namespace)

    def show_sms_popup(msg):
        return show_sms_popup_namespace_runtime(namespace, msg)

    namespace.update({
        "safe_save_config": safe_save_config,
        "safe_close_serial": safe_close_serial,
        "auto_connect_ui": auto_connect_ui,
        "lock_port_mutex": lock_port_mutex,
        "unlock_port_mutex": unlock_port_mutex,
        "check_single_instance": check_single_instance,
        "show_sms_popup": show_sms_popup,
    })
    return namespace
