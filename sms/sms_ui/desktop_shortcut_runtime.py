DEFAULT_DESKTOP_SHORTCUT_NAME = "短信监听系统"

from sms_core.config_runtime import restore_config_section, snapshot_config_section


def desktop_shortcut_default_name(config, fallback=DEFAULT_DESKTOP_SHORTCUT_NAME):
    return config.get("ui", "desktop_shortcut_name", fallback=fallback)


def _save_config_or_raise(save_config):
    result = save_config()
    if result is False:
        raise RuntimeError("配置保存失败")
    return True


def save_desktop_shortcut_name_runtime(config, name, save_config):
    config_snapshot = snapshot_config_section(config, "ui")
    try:
        if not config.has_section("ui"):
            config["ui"] = {}
        config.set("ui", "desktop_shortcut_name", name)
        return _save_config_or_raise(save_config)
    except Exception:
        restore_config_section(config, "ui", config_snapshot)
        raise


def open_desktop_shortcut_dialog_runtime(
    parent,
    *,
    config,
    save_config,
    create_shortcut,
    system_ui,
    center_window,
    open_dialog,
):
    default_name = desktop_shortcut_default_name(config)

    def save_name(name):
        save_desktop_shortcut_name_runtime(config, name, save_config)

    def apply_now(name):
        save_name(name)
        create_shortcut(name)
        system_ui(f"✅ 桌面快捷方式已创建：{name}.lnk", "normal")

    def save_only(name):
        save_name(name)
        system_ui(f"💾 已保存桌面快捷方式：{name}", "normal")

    open_dialog(parent, default_name, apply_now, save_only, center_window)
