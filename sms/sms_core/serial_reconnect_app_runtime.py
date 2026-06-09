from sms_core.serial_reconnect import manual_rebind_runtime


def try_rebind_manual_port_runtime(
    reason,
    *,
    mode,
    current_port,
    baud,
    find_luat_best_port,
    list_ports,
    choose_candidate,
    config,
    save_config,
    set_port,
    system_ui,
    set_status,
    wake_serial,
    reset_rebind_hint,
    hint_formatter,
    runtime=manual_rebind_runtime,
):
    return runtime(
        mode=mode,
        current_port=current_port,
        baud=baud,
        reason=reason,
        find_luat_best_port=find_luat_best_port,
        list_ports=list_ports,
        choose_candidate=choose_candidate,
        config=config,
        save_config=save_config,
        set_port=set_port,
        system_ui=system_ui,
        set_status=set_status,
        wake_serial=wake_serial,
        reset_rebind_hint=reset_rebind_hint,
        hint_formatter=hint_formatter,
    )
