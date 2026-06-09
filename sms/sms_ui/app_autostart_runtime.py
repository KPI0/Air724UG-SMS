def set_autostart_runtime(
    enable,
    *,
    autostart_flag,
    create_startup_shortcut,
    remove_startup_shortcut,
    system_ui,
    show_error,
):
    try:
        if enable:
            create_startup_shortcut(autostart_flag)
            message = "✅️ 开机自启：已打开"
        else:
            remove_startup_shortcut()
            message = "❌ 开机自启：已关闭"
        system_ui(message, "normal")
        return True
    except Exception as exc:
        show_error("错误", f"设置开机自启失败：\n{exc}")
        return False
