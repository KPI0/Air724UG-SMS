from sms_ui.app_lifecycle_namespace_runtime import (
    cleanup_and_exit_namespace_runtime,
    restart_software_namespace_runtime,
    set_autostart_namespace_runtime,
    toggle_multi_instance_namespace_runtime,
    toggle_popup_namespace_runtime,
    toggle_voice_broadcast_namespace_runtime,
)


def install_app_lifecycle_namespace_bindings(namespace):
    def set_autostart(enable):
        return set_autostart_namespace_runtime(namespace, enable)

    def cleanup_and_exit():
        def do_cleanup():
            return cleanup_and_exit_namespace_runtime(namespace)

        return namespace["run_on_ui_thread"](do_cleanup, namespace["ui_post"])

    def toggle_voice_broadcast():
        return toggle_voice_broadcast_namespace_runtime(namespace)

    def toggle_multi_instance():
        return toggle_multi_instance_namespace_runtime(namespace)

    def toggle_autostart():
        return namespace["set_autostart"](namespace["autostart_var"].get())

    def toggle_popup():
        return toggle_popup_namespace_runtime(namespace)

    def restart_software():
        return restart_software_namespace_runtime(namespace)

    namespace.update({
        "set_autostart": set_autostart,
        "cleanup_and_exit": cleanup_and_exit,
        "toggle_voice_broadcast": toggle_voice_broadcast,
        "toggle_multi_instance": toggle_multi_instance,
        "toggle_autostart": toggle_autostart,
        "toggle_popup": toggle_popup,
        "restart_software": restart_software,
    })
    return namespace
