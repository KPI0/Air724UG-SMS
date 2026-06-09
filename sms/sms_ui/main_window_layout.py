from tkinter.scrolledtext import ScrolledText


def build_main_window_layout_runtime(root, tk_module, *, cloud_enabled, scrolled_text_class=ScrolledText):
    root.grid_rowconfigure(0, weight=1)
    root.grid_rowconfigure(1, weight=0)
    root.grid_columnconfigure(0, weight=1)

    main_frame = tk_module.Frame(root)
    main_frame.grid(row=0, column=0, sticky="nsew")

    text_area = scrolled_text_class(main_frame, font=("微软雅黑", 10))
    text_area.pack(fill=tk_module.BOTH, expand=True)

    status_frame = tk_module.Frame(root)
    status_frame.grid(row=1, column=0, sticky="ew")

    status_var = tk_module.StringVar(value="🔍 启动中…")
    status_label = tk_module.Label(status_frame, textvariable=status_var, anchor="w")
    status_label.pack(side=tk_module.LEFT, padx=6)

    temp_var = tk_module.StringVar(value="🌡️ -- ℃")
    temp_label = tk_module.Label(status_frame, textvariable=temp_var, anchor="w", fg="#008000")
    temp_label.pack(side=tk_module.LEFT, padx=(20, 6))

    signal_var = tk_module.StringVar(value="📶 -- dBm")
    signal_label = tk_module.Label(status_frame, textvariable=signal_var, anchor="w", fg="#008000")
    signal_label.pack(side=tk_module.LEFT, padx=(20, 6))

    cloud_var = tk_module.StringVar(value="🌐 等待连接" if cloud_enabled else "🌐 已关闭")
    cloud_label = tk_module.Label(status_frame, textvariable=cloud_var, anchor="w", fg="#666666")
    cloud_label.pack(side=tk_module.LEFT, padx=(20, 6))

    return {
        "main_frame": main_frame,
        "text_area": text_area,
        "status_frame": status_frame,
        "status_var": status_var,
        "status_label": status_label,
        "temp_var": temp_var,
        "temp_label": temp_label,
        "signal_var": signal_var,
        "signal_label": signal_label,
        "cloud_var": cloud_var,
        "cloud_label": cloud_label,
    }
