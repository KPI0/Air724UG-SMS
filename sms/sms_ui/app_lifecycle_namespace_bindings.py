from sms_core.namespace_binding import make_namespace_runtime_binder
from sms_ui.app_lifecycle_namespace_runtime import (
    cleanup_and_exit_namespace_runtime,
    restart_software_namespace_runtime,
    set_autostart_namespace_runtime,
    toggle_call_popup_namespace_runtime,
    toggle_multi_instance_namespace_runtime,
    toggle_popup_namespace_runtime,
    toggle_voice_broadcast_namespace_runtime,
)


def install_app_lifecycle_namespace_bindings(namespace):
    bind = make_namespace_runtime_binder(namespace, globals())

    def cleanup_and_exit():
        # pystray invokes menu callbacks on its worker thread. Keep the whole
        # confirmation and shutdown flow on the Tk thread for every caller.
        def do_cleanup():
            return cleanup_and_exit_namespace_runtime(namespace)

        return namespace["run_on_ui_thread"](do_cleanup, namespace["ui_post"])

    def toggle_autostart():
        return namespace["set_autostart"](namespace["autostart_var"].get())

    namespace.update({
        "set_autostart": bind("set_autostart_namespace_runtime"),
        "cleanup_and_exit": cleanup_and_exit,
        "toggle_voice_broadcast": bind("toggle_voice_broadcast_namespace_runtime"),
        "toggle_multi_instance": bind("toggle_multi_instance_namespace_runtime"),
        "toggle_autostart": toggle_autostart,
        "toggle_popup": bind("toggle_popup_namespace_runtime"),
        "toggle_call_popup": bind("toggle_call_popup_namespace_runtime"),
        "restart_software": bind("restart_software_namespace_runtime"),
    })
    return namespace
