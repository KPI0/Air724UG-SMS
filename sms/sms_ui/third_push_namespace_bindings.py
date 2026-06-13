from sms_core.namespace_binding import make_namespace_runtime_binder
from sms_ui.third_push_namespace_runtime import (
    apply_third_push_settings_namespace_runtime,
    enqueue_third_push_namespace_runtime,
    ensure_third_push_config_namespace_runtime,
    open_third_push_window_namespace_runtime,
    refresh_third_push_settings_namespace_runtime,
    save_third_push_setting_namespace_runtime,
    show_third_push_test_result_namespace_runtime,
    third_push_worker_namespace_runtime,
)


def install_third_push_namespace_bindings(namespace):
    bind = make_namespace_runtime_binder(namespace, globals())

    namespace.update({
        "ensure_third_push_config": bind(
            "ensure_third_push_config_namespace_runtime",
            positional_keywords=("save",),
        ),
        "apply_third_push_settings": bind("apply_third_push_settings_namespace_runtime"),
        "refresh_third_push_settings_from_config": bind("refresh_third_push_settings_namespace_runtime"),
        "save_third_push_setting": bind(
            "save_third_push_setting_namespace_runtime",
            positional_keywords=("enabled", "sms_enabled", "call_enabled", "notify_type", "settings"),
        ),
        "_third_push_worker": bind("third_push_worker_namespace_runtime"),
        "show_third_push_test_result": bind("show_third_push_test_result_namespace_runtime"),
        "enqueue_third_push": bind(
            "enqueue_third_push_namespace_runtime",
            positional_keywords=(
                "show_success",
                "show_result",
                "channels",
                "settings",
                "template",
                "event_type",
            ),
            positional_prefix_count=1,
        ),
        "open_third_push_window": bind("open_third_push_window_namespace_runtime"),
    })
    return namespace
