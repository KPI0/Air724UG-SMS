from sms_core.serial_parsers import (
    parse_cesq_rsrp,
    parse_local_number,
    parse_serial_debug_insights,
    parse_temperature,
)


def apply_serial_line_effects(
    line,
    push_serial_debug,
    send_cloud_serial_log,
    capture_cloud_device_imei,
    set_temperature,
    set_signal,
    set_local_number=None,
):
    push_serial_debug(line)
    send_cloud_serial_log(line)
    capture_cloud_device_imei(line)

    temp_val = parse_temperature(line)
    if temp_val is not None:
        set_temperature(temp_val)

    rsrp_str = parse_cesq_rsrp(line)
    if rsrp_str is not None:
        set_signal(rsrp_str)

    local_number = parse_local_number(line)
    if local_number is not None and set_local_number is not None:
        set_local_number(local_number)


def push_serial_debug_insights(line, push_serial_debug):
    for insight in parse_serial_debug_insights(line):
        push_serial_debug(insight)
