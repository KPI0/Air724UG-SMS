import datetime
import time

from sms_core.config_runtime import safe_save_config_runtime
from sms_core.serial_io_runtime import safe_close_serial_runtime
from sms_core.windows_runtime import close_windows_handle, create_named_mutex, port_mutex_name
from sms_ui.app_instance_runtime import check_single_instance_app_runtime
from sms_ui.thread_runtime import schedule_delayed_ui_runtime
from sms_ui.ui_log_runtime import schedule_next_midnight_clear_runtime, show_sms_popup_runtime


def _safe_log(namespace, message):
    log_error = namespace.get("log_file_only")
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def safe_save_config_namespace_runtime(namespace):
    return safe_save_config_runtime(
        config=namespace["config"],
        config_file=namespace["CONFIG_FILE"],
        config_lock=namespace["CONFIG_LOCK"],
        log_error=lambda message: namespace["log_file_only"](message),
    )


def safe_close_serial_namespace_runtime(namespace):
    return safe_close_serial_runtime(
        namespace["serial_lock"],
        lambda: namespace["serial_obj"],
        lambda value: namespace.__setitem__("serial_obj", value),
        namespace["unlock_port_mutex"],
    )


def schedule_delayed_ui_namespace_runtime(namespace, callback):
    return schedule_delayed_ui_runtime(
        callback,
        app_start_mono=namespace["APP_START_MONO"],
        start_ui_delay=namespace["START_UI_DELAY"],
        monotonic=namespace.get("time", time).monotonic,
        root_after=namespace["root"].after,
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
        log_error=namespace.get("log_file_only"),
    )


def lock_port_mutex_namespace_runtime(namespace, port_name):
    namespace["unlock_port_mutex"]()
    if not port_name:
        return None
    mutex = create_named_mutex(port_mutex_name(port_name))
    namespace.__setitem__("current_port_mutex", mutex)
    return mutex


def unlock_port_mutex_namespace_runtime(namespace):
    mutex = namespace["current_port_mutex"]
    if mutex:
        try:
            close_windows_handle(mutex)
        except Exception as exc:
            _safe_log(namespace, f"Close port mutex failed: {exc!r}")
        namespace.__setitem__("current_port_mutex", None)


def check_single_instance_namespace_runtime(namespace):
    mutex = check_single_instance_app_runtime(
        allow_multi_instance=namespace["ALLOW_MULTI_INSTANCE"],
        window_title=namespace["APP_WINDOW_TITLE"],
        log_error=namespace.get("log_file_only"),
    )
    namespace.__setitem__("app_mutex", mutex)
    return mutex


def show_sms_popup_namespace_runtime(namespace, message):
    def do_show():
        return show_sms_popup_runtime(
            message,
            popup_enabled=namespace["POPUP_ENABLED"],
            show_info=namespace["messagebox"].showinfo,
            show_window=namespace["show_window"],
        )

    return namespace["run_on_ui_thread"](do_show, namespace["ui_post"])


def schedule_next_midnight_clear_namespace_runtime(namespace):
    datetime_source = namespace.get("datetime", datetime)
    now_func = getattr(datetime_source, "now", None)
    if now_func is None:
        now_func = datetime_source.datetime.now

    def do_schedule():
        return schedule_next_midnight_clear_runtime(
            tk_alive=namespace["tk_alive"],
            schedule_after=namespace["root"].after,
            clear_callback=namespace["clear_text_area_for_new_day"],
            now_func=now_func,
        )

    return namespace["run_on_ui_thread"](do_schedule, namespace["ui_post"])
