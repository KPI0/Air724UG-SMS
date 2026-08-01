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


def safe_save_config_namespace_runtime(namespace, *, defaults_by_section=None):
    return safe_save_config_runtime(
        config=namespace["config"],
        config_file=namespace["CONFIG_FILE"],
        config_lock=namespace["CONFIG_LOCK"],
        log_error=lambda message: namespace["log_file_only"](message),
        defaults_by_section=defaults_by_section,
    )


def safe_close_serial_namespace_runtime(namespace):
    namespace.__setitem__(
        "serial_connection_generation",
        int(namespace.get("serial_connection_generation", 0)) + 1,
    )
    coordinator = namespace.get("SMS_SEND_COORDINATOR")
    if coordinator is not None:
        try:
            coordinator.cancel_active("串口连接已关闭，短信发送已取消")
        except Exception as exc:
            _safe_log(namespace, f"Cancel active SMS send before serial close failed: {exc!r}")
    command_coordinator = namespace.get("SERIAL_COMMAND_RESPONSE_COORDINATOR")
    if command_coordinator is not None:
        try:
            command_coordinator.cancel_active("串口连接已关闭，AT 指令事务已取消")
        except Exception as exc:
            _safe_log(namespace, f"Cancel active AT transaction before serial close failed: {exc!r}")
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
        app_dir=namespace["APP_DIR"],
        log_error=namespace.get("log_file_only"),
    )
    namespace.__setitem__("app_mutex", mutex)
    return mutex


def show_sms_popup_namespace_runtime(namespace, message):
    def do_show():
        return show_sms_popup_runtime(
            message,
            popup_enabled=namespace["POPUP_ENABLED"],
            parent=namespace["root"],
            current_popup=namespace.get("sms_popup_win"),
            set_popup=lambda win: namespace.__setitem__("sms_popup_win", win),
            center_on_screen=namespace["center_on_screen"],
            show_window=namespace["show_window"],
        )

    return namespace["run_on_ui_thread"](do_show, namespace["ui_post"])


def schedule_next_midnight_clear_namespace_runtime(namespace):
    datetime_source = namespace.get("datetime", datetime)
    now_func = getattr(datetime_source, "now", None)
    if now_func is None:
        now_func = datetime_source.datetime.now
    state = namespace.setdefault("_MIDNIGHT_CLEAR_STATE", {})
    root = namespace["root"]

    def do_schedule():
        return schedule_next_midnight_clear_runtime(
            tk_alive=namespace["tk_alive"],
            schedule_after=root.after,
            cancel_after=getattr(root, "after_cancel", None),
            clear_callback=namespace["clear_text_area_for_new_day"],
            now_func=now_func,
            state=state,
        )

    return namespace["run_on_ui_thread"](do_schedule, namespace["ui_post"])
