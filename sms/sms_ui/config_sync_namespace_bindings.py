from sms_core.namespace_binding import make_namespace_runtime_binder
from sms_ui.config_sync_namespace_runtime import (
    register_config_sync_refresher_namespace_runtime,
    reload_shared_ui_config_namespace_runtime,
    start_config_file_watch_namespace_runtime,
)


def install_config_sync_namespace_bindings(namespace):
    bind = make_namespace_runtime_binder(namespace, globals())
    namespace.update({
        "register_config_sync_refresher": bind(
            "register_config_sync_refresher_namespace_runtime",
            positional_keywords=("group", "callback"),
        ),
        "reload_shared_ui_config": bind("reload_shared_ui_config_namespace_runtime"),
        "start_config_file_watch": bind(
            "start_config_file_watch_namespace_runtime",
            positional_keywords=("interval_ms",),
        ),
    })
    return namespace
