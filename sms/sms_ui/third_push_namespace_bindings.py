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
    def ensure_third_push_config(save=False):
        return ensure_third_push_config_namespace_runtime(namespace, save=save)

    def apply_third_push_settings(settings):
        return apply_third_push_settings_namespace_runtime(namespace, settings)

    def refresh_third_push_settings_from_config():
        return refresh_third_push_settings_namespace_runtime(namespace)

    def save_third_push_setting(
        enabled=None,
        sms_enabled=None,
        call_enabled=None,
        notify_type=None,
        settings=None,
    ):
        return save_third_push_setting_namespace_runtime(
            namespace,
            enabled=enabled,
            sms_enabled=sms_enabled,
            call_enabled=call_enabled,
            notify_type=notify_type,
            settings=settings,
        )

    def third_push_worker():
        return third_push_worker_namespace_runtime(namespace)

    def show_third_push_test_result(ok_channels, fail_infos):
        return show_third_push_test_result_namespace_runtime(namespace, ok_channels, fail_infos)

    def enqueue_third_push(
        raw_msg,
        show_success=False,
        show_result=False,
        channels=None,
        settings=None,
        template=None,
        event_type="sms",
    ):
        return enqueue_third_push_namespace_runtime(
            namespace,
            raw_msg,
            channels=channels,
            settings=settings,
            template=template,
            event_type=event_type,
            show_success=show_success,
            show_result=show_result,
        )

    def open_third_push_window():
        return open_third_push_window_namespace_runtime(namespace)

    namespace.update({
        "ensure_third_push_config": ensure_third_push_config,
        "apply_third_push_settings": apply_third_push_settings,
        "refresh_third_push_settings_from_config": refresh_third_push_settings_from_config,
        "save_third_push_setting": save_third_push_setting,
        "_third_push_worker": third_push_worker,
        "show_third_push_test_result": show_third_push_test_result,
        "enqueue_third_push": enqueue_third_push,
        "open_third_push_window": open_third_push_window,
    })
    return namespace
