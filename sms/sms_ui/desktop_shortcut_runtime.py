DEFAULT_DESKTOP_SHORTCUT_NAME = "短信监听系统"


def desktop_shortcut_default_name(config, fallback=DEFAULT_DESKTOP_SHORTCUT_NAME):
    return config.get("ui", "desktop_shortcut_name", fallback=fallback)


def save_desktop_shortcut_name_runtime(config, name, save_config):
    if not config.has_section("ui"):
        config["ui"] = {}
    config.set("ui", "desktop_shortcut_name", name)
    save_config()


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
        create_shortcut(name)
        save_name(name)
        system_ui(f"✅ 桌面快捷方式已创建：{name}.lnk", "normal")

    def save_only(name):
        save_name(name)
        system_ui(f"💾 已保存桌面快捷方式：{name}", "normal")

    open_dialog(parent, default_name, apply_now, save_only, center_window)
