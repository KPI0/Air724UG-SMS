from tkinter.scrolledtext import ScrolledText


_CONTROL_MASK = 0x0004
_MUTATING_TEXT_KEYSYMS = {
    "BackSpace",
    "Delete",
    "Return",
    "KP_Enter",
    "Tab",
    "Clear",
}


def _block_text_edit_event(_event=None):
    return "break"


def _has_control_modifier(event):
    try:
        return bool(int(getattr(event, "state", 0) or 0) & _CONTROL_MASK)
    except Exception:
        return False


def _select_all_text(text_widget):
    try:
        text_widget.tag_add("sel", "1.0", "end-1c")
        text_widget.mark_set("insert", "1.0")
        text_widget.see("insert")
    except Exception:
        pass
    return "break"


def main_text_readonly_key_handler(event, text_widget=None):
    keysym = str(getattr(event, "keysym", "") or "")
    char = str(getattr(event, "char", "") or "")

    if _has_control_modifier(event):
        lowered = keysym.lower()
        if lowered == "c" or keysym == "Insert":
            return None
        if lowered == "a" and text_widget is not None:
            return _select_all_text(text_widget)
        return "break"

    if keysym in _MUTATING_TEXT_KEYSYMS:
        return "break"
    if char:
        return "break"
    return None


def protect_main_text_widget_runtime(text_widget):
    text_widget.bind("<Key>", lambda event: main_text_readonly_key_handler(event, text_widget))
    text_widget.bind("<<Paste>>", _block_text_edit_event)
    text_widget.bind("<<Cut>>", _block_text_edit_event)
    text_widget.bind("<<Clear>>", _block_text_edit_event)
    text_widget.bind("<<PasteSelection>>", _block_text_edit_event)
    text_widget.bind("<Button-2>", _block_text_edit_event)
    return text_widget


def _flow_status_labels(root, frame, labels):
    def layout(_event=None):
        available = max(1, frame.winfo_width() - 12)
        x, y, row_height = 6, 0, 0
        for label in labels:
            label.configure(wraplength=max(1, available - 4), justify="left")
            width = min(available, label.winfo_reqwidth())
            height = label.winfo_reqheight()
            if x > 6 and x + width > available + 6:
                x, y, row_height = 6, y + row_height, 0
            label.place(x=x, y=y, width=width, height=height)
            x += width + 20
            row_height = max(row_height, height)
        frame.configure(height=y + row_height)
        root.minsize(root.minsize()[0], max(200, y + row_height + 80))

    # Label size changes also cover updated text and named fonts. Widget-owned
    # bindings are released on destruction; no variable traces or timers remain.
    frame.bind("<Configure>", layout, add="+")
    for label in labels:
        label.bind("<Configure>", layout, add="+")


def build_main_window_layout_runtime(root, tk_module, *, cloud_enabled, scrolled_text_class=ScrolledText):
    root.grid_rowconfigure(0, weight=1)
    root.grid_rowconfigure(1, weight=0)
    root.grid_columnconfigure(0, weight=1)

    main_frame = tk_module.Frame(root)
    main_frame.grid(row=0, column=0, sticky="nsew")

    text_area = scrolled_text_class(main_frame, font=("微软雅黑", 10))
    protect_main_text_widget_runtime(text_area)
    text_area.pack(fill=tk_module.BOTH, expand=True)

    status_frame = tk_module.Frame(root)
    status_frame.grid(row=1, column=0, sticky="ew")

    status_var = tk_module.StringVar(value="🔍 启动中…")
    status_label = tk_module.Label(status_frame, textvariable=status_var, anchor="w")

    temp_var = tk_module.StringVar(value="🌡️ -- ℃")
    temp_label = tk_module.Label(status_frame, textvariable=temp_var, anchor="w", fg="#008000")

    signal_var = tk_module.StringVar(value="📶 -- dBm")
    signal_label = tk_module.Label(status_frame, textvariable=signal_var, anchor="w", fg="#008000")

    cloud_var = tk_module.StringVar(value="🌐 等待连接" if cloud_enabled else "🌐 已关闭")
    cloud_label = tk_module.Label(status_frame, textvariable=cloud_var, anchor="w", fg="#666666")
    _flow_status_labels(root, status_frame, (status_label, temp_label, signal_label, cloud_label))

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
