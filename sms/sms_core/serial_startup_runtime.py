import time

from sms_core.serial_ports import unlocked_ports
from sms_core.serial_reconnect import is_serial_open_denied, serial_open_denied_repeat_key
from sms_core.serial_sender import DEFAULT_SERIAL_TRANSACTION_LOCK, write_serial_command_result


def resolve_serial_target_port_runtime(
    *,
    mode,
    current_port,
    reconnect_interval,
    find_luat_best_port,
    list_ports,
    is_port_locked,
    auto_connect_ui,
    set_status,
    wakeup_wait,
    wakeup_clear,
    sleep=time.sleep,
    unlocked_ports_func=unlocked_ports,
):
    if mode == "Auto":
        dev, desc = find_luat_best_port()
        if dev:
            auto_connect_ui(f"🔌 检测到 LUAT Modem 标识，自动连接：{dev}（{desc}）")
            return dev

        available_ports = unlocked_ports_func(list_ports(), is_port_locked)
        if len(available_ports) == 1:
            single = available_ports[0]
            auto_connect_ui(f"🔌 未检测到 LUAT Modem 标识，但仅发现单一串口，自动连接：{single.device}")
            return single.device

        set_status("🔍 扫描 LUAT Modem 中…", "orange")
        wakeup_wait(timeout=reconnect_interval)
        wakeup_clear()
        return None

    if not current_port:
        set_status("🔒 手动模式：未指定串口", "red")
        sleep(reconnect_interval)
        return None

    return current_port


def open_and_initialize_serial_runtime(
    *,
    target_port,
    baud,
    mode,
    serial_lock,
    open_serial,
    set_serial_obj,
    set_port,
    lock_port_mutex,
    set_cloud_imei_query_deadline,
    serial_error_ui,
    set_status,
    write_command=write_serial_command_result,
    transaction_lock=DEFAULT_SERIAL_TRANSACTION_LOCK,
    is_open_denied=is_serial_open_denied,
    open_denied_repeat_key=serial_open_denied_repeat_key,
    monotonic=time.monotonic,
):
    with transaction_lock:
        try:
            with serial_lock:
                serial_obj = open_serial(target_port, baud)
                set_serial_obj(serial_obj)
        except Exception as exc:
            if is_open_denied(str(exc)):
                serial_error_ui(
                    f"⚠️ 端口占用：{target_port} 已被其他程序或本软件其他实例占用。",
                    repeat_key=open_denied_repeat_key(target_port),
                )
                set_status("🔴 端口占用，等待释放…", "red")
            raise

        lock_port_mutex(target_port)
        write_command(serial_obj, "AT+CLIP=1")
        try:
            set_cloud_imei_query_deadline(monotonic() + 6.0)
            write_command(serial_obj, "AT+CGSN")
        except Exception:
            pass
        try:
            write_command(serial_obj, "AT+CNUM")
        except Exception:
            pass

        if mode == "Auto":
            set_port(target_port)

        return serial_obj
