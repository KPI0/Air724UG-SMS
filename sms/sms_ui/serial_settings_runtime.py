from sms_core.config_runtime import restore_config_section, snapshot_config_section
from sms_ui.serial_settings_dialog import open_serial_setting_dialog, parse_positive_baud


def serial_settings_status(mode, port, baud):
    return f"⚙️ 串口设置已更新：mode={mode} port={port or '(Auto)'} baud={baud}"


def serial_settings_save_failed_status():
    return "❌ 串口设置保存失败，已保留原设置"


def serial_settings_invalid_baud_status():
    return "❌ 串口设置无效：波特率必须是大于 0 的整数"


def apply_serial_setting_runtime(
    mode,
    port,
    baud,
    *,
    config,
    save_config,
    set_serial_state,
    set_status,
    safe_close_serial,
    wake_serial,
    system_ui,
):
    try:
        baud = parse_positive_baud(baud)
    except ValueError:
        system_ui(serial_settings_invalid_baud_status())
        return False

    config_snapshot = snapshot_config_section(config, "serial")
    if not config.has_section("serial"):
        config["serial"] = {}
    config.set("serial", "mode", mode)
    config.set("serial", "port", port)
    config.set("serial", "baud", str(baud))
    try:
        if save_config() is False:
            raise RuntimeError("配置保存失败")
    except Exception:
        restore_config_section(config, "serial", config_snapshot)
        system_ui(serial_settings_save_failed_status())
        return False

    set_serial_state(mode, port, baud)
    set_status("🟡 应用中，重连…", "orange")
    safe_close_serial()
    try:
        wake_serial()
    except Exception:
        pass

    system_ui(serial_settings_status(mode, port, baud))
    return True


def open_serial_setting_runtime(
    parent,
    *,
    current_mode,
    current_port,
    current_baud,
    scan_ports,
    center_window,
    config,
    save_config,
    set_serial_state,
    set_status,
    safe_close_serial,
    wake_serial,
    system_ui,
    open_dialog=open_serial_setting_dialog,
):
    def apply(mode, port, baud):
        return apply_serial_setting_runtime(
            mode,
            port,
            baud,
            config=config,
            save_config=save_config,
            set_serial_state=set_serial_state,
            set_status=set_status,
            safe_close_serial=safe_close_serial,
            wake_serial=wake_serial,
            system_ui=system_ui,
        )

    open_dialog(
        parent,
        current_mode,
        current_port,
        current_baud,
        scan_ports,
        apply,
        center_window,
    )
