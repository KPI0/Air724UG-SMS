RESET_COMMAND = "AT+RESET"


def send_reset_command_runtime(
    *,
    confirm_reset,
    send_command_async,
    serial_lock,
    get_serial,
    ui_post,
    system_ui,
    show_warning,
    log_error=None,
):
    if not confirm_reset():
        return "cancelled"

    def on_result(result):
        if not result.ok:
            error = str(getattr(result, "error", ""))
            ui_post(lambda error=error: show_warning("提示", f"串口当前未连接或发送失败：{error}"))
            system_ui(f"❌ 发送重启指令失败：{error}", "normal")
            return

        system_ui(f"🔄 已发送重启指令：{RESET_COMMAND}", "normal")

    send_command_async(
        serial_lock,
        get_serial,
        RESET_COMMAND,
        on_result=on_result,
        log_error=log_error,
    )
    return "submitted"
