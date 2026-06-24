from sms_core.serial_runtime import (
    SerialRuntimeCallbacks,
    SerialRuntimeConfig,
    run_serial_runtime_thread,
)
from sms_core.status_text import format_connecting_status


def build_serial_runtime_config(settings):
    return SerialRuntimeConfig(
        keywords=settings["keywords"](),
        log_unmatched_sms=settings["log_unmatched_sms"](),
        log_dir=settings["log_dir"](),
        log_prefix=settings["log_prefix"](),
        error_repeat_limit=settings["error_repeat_limit"](),
        call_filter_mode=settings["call_filter_mode"](),
        call_whitelist=settings["call_whitelist"](),
        call_blacklist=settings["call_blacklist"](),
    )


def build_serial_runtime_callbacks(callbacks):
    return SerialRuntimeCallbacks(
        enqueue_third_push=callbacks["enqueue_third_push"],
        send_cloud_sms_event=callbacks["send_cloud_sms_event"],
        port_ui=callbacks["port_ui"],
        play_alert=callbacks["play_alert"],
        show_sms_popup=callbacks["show_sms_popup"],
        file_log=callbacks["file_log"],
        system_ui=callbacks["system_ui"],
        push_serial_debug=callbacks["push_serial_debug"],
        send_cloud_serial_log=callbacks["send_cloud_serial_log"],
        capture_cloud_device_imei=callbacks["capture_cloud_device_imei"],
        set_temperature=callbacks["set_temperature"],
        set_signal=callbacks["set_signal"],
        set_status=callbacks["set_status"],
        close_call_popup=callbacks["close_call_popup"],
        send_call_hangup=callbacks["send_call_hangup"],
        show_call_popup=callbacks["show_call_popup"],
        set_local_number=callbacks.get("set_local_number", lambda *_args: None),
    )


def run_serial_app_runtime(
    *,
    parse_callback_head,
    settings,
    callbacks,
    state,
    io,
    reconnect,
    apply_disconnect_effects,
    run_runtime_thread=run_serial_runtime_thread,
):
    runtime_callbacks = build_serial_runtime_callbacks(callbacks)

    def get_runtime_config():
        return build_serial_runtime_config(settings)

    def on_connected_port(target_port):
        state["set_log_prefix"](target_port.replace(":", "_"))
        callbacks.get("set_local_number", lambda *_args: None)("")
        callbacks["schedule_connected_log"](state["port"](), state["baud"]())

    def handle_disconnect(error, target_port):
        state["set_log_prefix"]("system")
        callbacks.get("set_local_number", lambda *_args: None)("")
        callbacks["close_call_popup"]()
        apply_disconnect_effects(
            error,
            target_port,
            state["port"](),
            callbacks["serial_error_ui"],
            callbacks["set_status"],
            callbacks["set_temperature"],
            callbacks["set_signal"],
            io["safe_close_serial"],
        )
        return reconnect["try_manual_rebind_after_error"](error)

    def wait_before_retry():
        reconnect["wakeup_wait"](timeout=reconnect["interval"]())
        reconnect["wakeup_clear"]()

    return run_runtime_thread(
        parse_callback_head=parse_callback_head,
        get_runtime_config=get_runtime_config,
        callbacks=runtime_callbacks,
        get_call_state=state["get_call_state"],
        set_call_state=state["set_call_state"],
        popup_active=state["popup_active"],
        ignore_repeat_state=state["ignore_repeat_state"],
        should_continue=lambda: state["serial_running"]() and not reconnect["stop_requested"](),
        get_target_port=state["port"],
        resolve_target_port=io["resolve_target_port"],
        set_connecting_status=lambda target_port: callbacks["set_status"](
            format_connecting_status(target_port),
            "orange",
        ),
        open_and_initialize_serial=io["open_and_initialize_serial"],
        on_connected_port=on_connected_port,
        read_serial_line=io["read_serial_line"],
        handle_disconnect=handle_disconnect,
        wait_before_retry=wait_before_retry,
        safe_close_serial=io["safe_close_serial"],
    )


def build_serial_app_wiring(
    *,
    values,
    callbacks,
    state_access,
    io_callbacks,
    reconnect_callbacks,
):
    settings = {
        "keywords": values["keywords"],
        "log_unmatched_sms": values["log_unmatched_sms"],
        "log_dir": values["log_dir"],
        "log_prefix": values["log_prefix"],
        "error_repeat_limit": values["error_repeat_limit"],
        "call_filter_mode": values["call_filter_mode"],
        "call_whitelist": values["call_whitelist"],
        "call_blacklist": values["call_blacklist"],
    }
    runtime_callbacks = {
        "enqueue_third_push": callbacks["enqueue_third_push"],
        "send_cloud_sms_event": callbacks["send_cloud_sms_event"],
        "port_ui": callbacks["port_ui"],
        "play_alert": callbacks["play_alert"],
        "show_sms_popup": callbacks["show_sms_popup"],
        "file_log": callbacks["file_log"],
        "system_ui": callbacks["system_ui"],
        "push_serial_debug": callbacks["push_serial_debug"],
        "send_cloud_serial_log": callbacks["send_cloud_serial_log"],
        "capture_cloud_device_imei": callbacks["capture_cloud_device_imei"],
        "set_temperature": callbacks["set_temperature"],
        "set_signal": callbacks["set_signal"],
        "set_local_number": callbacks.get("set_local_number", lambda *_args: None),
        "set_status": callbacks["set_status"],
        "close_call_popup": callbacks["close_call_popup"],
        "send_call_hangup": callbacks["send_call_hangup"],
        "show_call_popup": callbacks["show_call_popup"],
        "schedule_connected_log": callbacks["schedule_connected_log"],
        "serial_error_ui": callbacks["serial_error_ui"],
    }
    state = {
        "get_call_state": state_access["get_call_state"],
        "set_call_state": state_access["set_call_state"],
        "popup_active": state_access["popup_active"],
        "ignore_repeat_state": state_access["ignore_repeat_state"],
        "serial_running": state_access["serial_running"],
        "port": state_access["port"],
        "baud": state_access["baud"],
        "set_log_prefix": state_access["set_log_prefix"],
    }
    io = {
        "resolve_target_port": io_callbacks["resolve_target_port"],
        "open_and_initialize_serial": io_callbacks["open_and_initialize_serial"],
        "read_serial_line": io_callbacks["read_serial_line"],
        "safe_close_serial": io_callbacks["safe_close_serial"],
    }
    reconnect = {
        "stop_requested": reconnect_callbacks["stop_requested"],
        "interval": reconnect_callbacks["interval"],
        "wakeup_wait": reconnect_callbacks["wakeup_wait"],
        "wakeup_clear": reconnect_callbacks["wakeup_clear"],
        "try_manual_rebind_after_error": reconnect_callbacks["try_manual_rebind_after_error"],
    }
    return settings, runtime_callbacks, state, io, reconnect


def run_serial_app_from_wiring(
    *,
    parse_callback_head,
    values,
    callbacks,
    state_access,
    io_callbacks,
    reconnect_callbacks,
    apply_disconnect_effects,
    build_wiring=build_serial_app_wiring,
    run_app=run_serial_app_runtime,
):
    settings, runtime_callbacks, state, io, reconnect = build_wiring(
        values=values,
        callbacks=callbacks,
        state_access=state_access,
        io_callbacks=io_callbacks,
        reconnect_callbacks=reconnect_callbacks,
    )
    return run_app(
        parse_callback_head=parse_callback_head,
        settings=settings,
        callbacks=runtime_callbacks,
        state=state,
        io=io,
        reconnect=reconnect,
        apply_disconnect_effects=apply_disconnect_effects,
    )


def run_serial_reader_namespace_runtime(
    namespace,
    *,
    parse_callback_head,
    apply_disconnect_effects,
    run_reader=run_serial_app_from_wiring,
):
    return run_reader(
        parse_callback_head=parse_callback_head,
        values={
            "keywords": lambda: namespace["KEYWORDS"],
            "log_unmatched_sms": lambda: namespace["LOG_UNMATCHED_SMS"],
            "log_dir": lambda: namespace["LOG_DIR"],
            "log_prefix": lambda: namespace["LOG_PREFIX"],
            "error_repeat_limit": lambda: namespace["ERROR_REPEAT_LIMIT"],
            "call_filter_mode": lambda: namespace["CALL_FILTER_MODE"],
            "call_whitelist": lambda: namespace["CALL_WHITELIST"],
            "call_blacklist": lambda: namespace["CALL_BLACKLIST"],
        },
        callbacks={
            "enqueue_third_push": namespace["enqueue_third_push"],
            "send_cloud_sms_event": namespace["_cloud_send_sms_event"],
            "port_ui": namespace["port_ui"],
            "play_alert": namespace["play_alert"],
            "show_sms_popup": namespace["show_sms_popup"],
            "file_log": namespace["FILE_LOG_Q"].put_nowait,
            "system_ui": namespace["system_ui"],
            "push_serial_debug": namespace["_push_serial_debug"],
            "send_cloud_serial_log": namespace["_cloud_send_serial_log"],
            "capture_cloud_device_imei": namespace["_maybe_capture_cloud_device_imei"],
            "set_temperature": namespace["set_temperature"],
            "set_signal": namespace["set_signal"],
            "set_local_number": lambda value: namespace.__setitem__("LOCAL_NUMBER", value),
            "set_status": namespace["set_status"],
            "close_call_popup": namespace["close_call_popup"],
            "send_call_hangup": namespace["send_call_hangup_command"],
            "show_call_popup": namespace["show_call_popup"],
            "schedule_connected_log": namespace["schedule_delayed_connected_log"],
            "serial_error_ui": namespace["serial_error_ui"],
        },
        state_access={
            "get_call_state": namespace["get_serial_call_state"],
            "set_call_state": namespace["set_serial_call_state"],
            "popup_active": lambda: namespace["current_call_popup"] is not None,
            "ignore_repeat_state": namespace["_sms_ignore_repeat_state"],
            "serial_running": lambda: namespace["serial_running"],
            "port": lambda: namespace["PORT"],
            "baud": lambda: namespace["BAUD"],
            "set_log_prefix": namespace["set_serial_log_prefix"],
        },
        io_callbacks={
            "resolve_target_port": namespace["resolve_serial_target_port"],
            "open_and_initialize_serial": namespace["open_and_initialize_serial"],
            "read_serial_line": namespace["read_serial_line_safely"],
            "safe_close_serial": namespace["safe_close_serial"],
        },
        reconnect_callbacks={
            "stop_requested": namespace["serial_stop_event"].is_set,
            "interval": lambda: namespace["RECONNECT_INTERVAL"],
            "wakeup_wait": namespace["serial_wakeup_event"].wait,
            "wakeup_clear": namespace["serial_wakeup_event"].clear,
            "try_manual_rebind_after_error": namespace["try_manual_rebind_after_error"],
        },
        apply_disconnect_effects=apply_disconnect_effects,
    )
