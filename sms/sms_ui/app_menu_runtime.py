def build_main_menu_runtime(
    root,
    tk_module,
    *,
    is_autostart_enabled,
    allow_multi_instance,
    popup_var,
    call_popup_var,
    commands,
):
    menu_bar = tk_module.Menu(root)

    file_menu = tk_module.Menu(menu_bar, tearoff=0)
    file_menu.add_command(label="清空窗口", command=commands["clear_window"])
    file_menu.add_command(label="打开日志", command=commands["open_log_dir"])
    file_menu.add_separator()
    file_menu.add_command(label="重启软件", command=commands["restart_software"])
    file_menu.add_command(label="重启硬件", command=commands["send_reset_cmd"])
    file_menu.add_command(label="退出软件", command=commands["cleanup_and_exit"])
    menu_bar.add_cascade(label="文件", menu=file_menu)

    menu_bar.add_command(label="串口设置", command=commands["open_serial_setting"])
    menu_bar.add_command(label="关键词设置", command=commands["open_keywords_setting"])
    menu_bar.add_command(label="防骚扰设置", command=commands["open_call_filter_setting"])
    menu_bar.add_command(label="🔊 语音播报", command=commands["toggle_voice_broadcast"])
    voice_menu_index = menu_bar.index("end")

    settings_menu = tk_module.Menu(menu_bar, tearoff=0)
    autostart_var = tk_module.BooleanVar(value=is_autostart_enabled())
    multi_instance_var = tk_module.BooleanVar(value=allow_multi_instance)

    settings_menu.add_checkbutton(
        label="开机自启",
        variable=autostart_var,
        command=commands["toggle_autostart"],
    )
    settings_menu.add_checkbutton(
        label="程序多开",
        variable=multi_instance_var,
        command=commands["toggle_multi_instance"],
    )
    settings_menu.add_checkbutton(
        label="短信弹窗",
        variable=popup_var,
        command=commands["toggle_popup"],
    )
    settings_menu.add_checkbutton(
        label="电话弹窗",
        variable=call_popup_var,
        command=commands["toggle_call_popup"],
    )
    settings_menu.add_separator()
    settings_menu.add_command(label="日志清理", command=commands["open_log_cleanup_dialog"])
    settings_menu.add_command(label="代理设置", command=commands["open_update_proxy_dialog"])
    settings_menu.add_command(label="快捷方式", command=commands["open_desktop_shortcut_dialog"])
    settings_menu.add_command(label="语音播报", command=commands["open_voice_text_dialog"])
    settings_menu.add_command(label="短信字体", command=commands["open_sms_font_dialog"])
    settings_menu.add_command(label="云端控制", command=commands["open_cloud_control_window"])
    settings_menu.add_command(label="三方推送", command=commands["open_third_push_window"])
    settings_menu.add_command(label="串口调试", command=commands["open_serial_debug_window"])
    menu_bar.add_cascade(label="设置", menu=settings_menu)

    help_menu = tk_module.Menu(menu_bar, tearoff=0)
    help_menu.add_command(label="关于", command=commands["show_about"])
    help_menu.add_command(label="检测更新", command=commands["check_update_and_prompt"])
    menu_bar.add_cascade(label="帮助", menu=help_menu)

    root.config(menu=menu_bar)
    return {
        "menu_bar": menu_bar,
        "voice_menu_index": voice_menu_index,
        "autostart_var": autostart_var,
        "multi_instance_var": multi_instance_var,
    }
