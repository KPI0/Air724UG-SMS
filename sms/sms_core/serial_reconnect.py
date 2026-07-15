from sms_core.status_text import format_connecting_status
from sms_core.config_runtime import restore_config_section, snapshot_config_section


SERIAL_PORT_GONE_MARKERS = (
    "could not open port",
    "file not found",
    "no such file",
    "the system cannot find the file specified",
    "cleartcommerror",
    "clearcommerror",
    "semaphore timeout",
    "timeout period has expired",
    "device does not recognize",
    "access is denied",
    "winerror 2",
    "winerror 5",
    "winerror 31",
    "winerror 1167",
    "device not functioning",
    "设备",
    "不存在",
    "找不到",
    "系统找不到指定的文件",
)


def manual_rebind_runtime(
    *,
    mode,
    current_port,
    baud,
    reason,
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
):
    if mode != "Manual":
        return False

    try:
        dev, desc = find_luat_best_port()
    except Exception:
        dev, desc = (None, None)

    try:
        candidate = choose_candidate(dev, desc, list_ports(), current_port=current_port)
    except Exception:
        candidate = choose_candidate(dev, desc, [], current_port=current_port)
    if not candidate.found:
        return False

    old_port = current_port
    new_port = candidate.device
    config_snapshot = snapshot_config_section(config, "serial")

    try:
        if not config.has_section("serial"):
            config["serial"] = {}
        config.set("serial", "mode", "Manual")
        config.set("serial", "port", new_port)
        config.set("serial", "baud", str(baud))
        if save_config() is False:
            raise RuntimeError("配置保存失败")
    except Exception as exc:
        restore_config_section(config, "serial", config_snapshot)
        system_ui(f"⚠️ 串口重绑定已取消，配置保存失败：{exc}", "normal")
        return False

    set_port(new_port)
    system_ui(hint_formatter(old_port, new_port, candidate.description, reason), "normal")
    set_status(format_connecting_status(new_port), "orange")

    try:
        wake_serial()
    except Exception as exc:
        # wake_serial is the action that actually reconnects to the new port.
        # If it fails silently the device looks offline with no explanation,
        # which is the worst case for diagnosis. Report and keep going so the
        # status/hint already shown to the user stays consistent.
        system_ui(f"⚠️ 串口重绑定到 {new_port} 后唤醒失败：{exc}", "normal")
    reset_rebind_hint()
    return True


def is_serial_open_denied(error_text: str) -> bool:
    text = str(error_text or "")
    lower = text.lower()
    return (
        "access is denied" in lower
        or "permission" in lower
        or "拒绝访问" in text
        or "winerror 5" in lower
    )


def serial_open_denied_repeat_key(port) -> str:
    return f"serial-open-denied:{str(port or '').strip().upper()}"


def serial_disconnect_status(port) -> str:
    return f"🔴 断开/失败：{port}（自动重连中…）"


def apply_serial_disconnect_effects(
    error,
    target_port,
    current_port,
    serial_error_ui,
    set_status,
    set_temperature,
    set_signal,
    close_serial,
):
    err_msg = str(error)
    repeat_key = ""
    if is_serial_open_denied(err_msg):
        repeat_key = serial_open_denied_repeat_key(target_port or current_port)

    serial_error_ui(f"⚠️ 串口异常：{err_msg}", repeat_key=repeat_key)
    set_status(serial_disconnect_status(current_port), "red")
    set_temperature("--")
    set_signal("--")
    close_serial()


def is_serial_port_gone_error(error) -> bool:
    text = (repr(error) + " " + str(error)).lower()
    return any(marker in text for marker in SERIAL_PORT_GONE_MARKERS)
