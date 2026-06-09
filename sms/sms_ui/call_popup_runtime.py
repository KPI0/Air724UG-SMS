from sms_core.call_effects import apply_call_answer_result, apply_call_hangup_result
from sms_core.serial_sender import send_command_with_result_async
from sms_ui.call_popup import open_call_popup


def popup_exists(window):
    try:
        return window is not None and window.winfo_exists()
    except Exception:
        return False


def close_call_popup_runtime(current_popup, set_popup):
    try:
        if popup_exists(current_popup):
            current_popup.destroy()
    finally:
        set_popup(None)


def close_call_popup_app_runtime(
    *,
    get_popup,
    set_popup,
    run_on_ui_thread,
    ui_post,
    close_runtime=close_call_popup_runtime,
):
    def close_on_ui():
        return close_runtime(get_popup(), set_popup)

    return run_on_ui_thread(close_on_ui, ui_post)


def show_call_popup_runtime(
    parent,
    caller_num,
    current_popup,
    set_popup,
    center_window,
    serial_lock,
    serial_getter,
    port_ui,
    set_status,
    ui_post,
    close_popup,
    set_ring_timeout,
    *,
    open_popup=open_call_popup,
    send_command_async=send_command_with_result_async,
    log_error=None,
):
    if popup_exists(current_popup):
        return current_popup

    def answer(mark_connected, restore_answer):
        def on_result(result):
            apply_call_answer_result(
                result,
                caller_num,
                restore_answer,
                mark_connected,
                port_ui,
                set_status,
                ui_post,
                set_ring_timeout,
            )

        kwargs = {"on_result": on_result}
        if log_error is not None:
            kwargs["log_error"] = log_error
        send_command_async(serial_lock, serial_getter, "ATA", **kwargs)

    def hangup(restore_hangup):
        def on_result(result):
            apply_call_hangup_result(
                result,
                restore_hangup,
                port_ui,
                ui_post,
                close_popup,
            )

        kwargs = {"on_result": on_result}
        if log_error is not None:
            kwargs["log_error"] = log_error
        send_command_async(serial_lock, serial_getter, "ATH", **kwargs)

    win = open_popup(
        parent,
        caller_num,
        center_window,
        answer,
        hangup,
        close_popup,
        close_popup,
    )
    set_popup(win)
    return win


def show_call_popup_app_runtime(
    *,
    parent,
    caller_num,
    get_popup,
    set_popup,
    center_window,
    serial_lock,
    get_serial,
    port_ui,
    set_status,
    ui_post,
    close_popup,
    set_ring_timeout,
    run_on_ui_thread,
    log_error=None,
    show_runtime=show_call_popup_runtime,
):
    def show_on_ui():
        args = (
            parent,
            caller_num,
            get_popup(),
            set_popup,
            center_window,
            serial_lock,
            get_serial,
            port_ui,
            set_status,
            ui_post,
            close_popup,
            set_ring_timeout,
        )
        if log_error is None:
            return show_runtime(*args)
        return show_runtime(*args, log_error=log_error)

    return run_on_ui_thread(show_on_ui, ui_post)
