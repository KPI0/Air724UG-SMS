import tkinter as tk
from tkinter import colorchooser, messagebox


def open_sms_font_dialog(parent, current_size, current_color, on_save, center_window):
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("短信字体设置")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    bottom_line = tk.Frame(win, height=1, bg="#d4d4d4")
    bottom_line.pack(side="bottom", fill="x")
    bottom_line.pack_propagate(False)

    frame = tk.Frame(win, padx=14, pady=12)
    frame.pack(side="top", fill=tk.BOTH, expand=True)

    tk.Label(frame, text="字号：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w")

    size_var = tk.StringVar(value=str(current_size))
    size_spin = tk.Spinbox(frame, from_=8, to=72, width=8, textvariable=size_var)
    size_spin.grid(row=0, column=1, sticky="w", padx=(8, 0))

    tk.Label(frame, text="颜色：", font=("微软雅黑", 10)).grid(row=1, column=0, sticky="w", pady=(10, 0))

    color_var = tk.StringVar(value=current_color)
    color_entry = tk.Entry(frame, textvariable=color_var, width=14)
    color_entry.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(10, 0))

    preview_box = tk.LabelFrame(frame, text="预览", padx=8, pady=8)
    preview_box.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(12, 0))
    preview_box.grid_columnconfigure(0, weight=1)

    preview_canvas = tk.Canvas(preview_box, width=560, height=110, highlightthickness=1)
    preview_canvas.grid(row=0, column=0, sticky="ew")

    def refresh_preview():
        preview_canvas.update()
        try:
            size = int(size_var.get().strip())
        except Exception:
            size = current_size

        color = color_var.get().strip() or current_color
        width = preview_canvas.winfo_width()
        height = preview_canvas.winfo_height()
        if width <= 1:
            width = 560
        if height <= 1:
            height = 110

        preview_size = min(size, max(8, int(height * 0.7)))
        preview_canvas.delete("all")
        try:
            preview_canvas.create_text(
                width // 2,
                height // 2,
                text="短信内容",
                anchor="c",
                font=("微软雅黑", preview_size),
                fill=color,
            )
        except Exception:
            preview_canvas.create_text(
                width // 2,
                height // 2,
                text="短信内容",
                anchor="c",
                font=("微软雅黑", 30),
                fill="#ff0000",
            )

    def pick_color():
        color = color_var.get().strip() or current_color
        win.lift()
        win.after(0, lambda: win.lift())

        try:
            win.grab_release()
        except Exception:
            pass

        chosen = colorchooser.askcolor(parent=win, initialcolor=color, title="选择短信颜色")

        try:
            win.grab_set()
        except Exception:
            pass

        win.lift()
        win.after(0, lambda: win.lift())

        if chosen and chosen[1]:
            color_var.set(chosen[1])
            refresh_preview()

    tk.Button(frame, text="选颜色", width=10, command=pick_color).grid(row=1, column=2, padx=(8, 0), pady=(10, 0))

    def save():
        try:
            size = int(size_var.get().strip())
            if size < 8 or size > 72:
                raise ValueError
        except Exception:
            messagebox.showerror("错误", "字号必须是 8~72 的整数", parent=win)
            return

        color = color_var.get().strip() or "#ff0000"
        if on_save(size, color) is False:
            return
        win.destroy()

    buttons = tk.Frame(frame)
    buttons.grid(row=3, column=0, columnspan=3, sticky="e", pady=(14, 0))
    frame.grid_columnconfigure(1, weight=1)

    tk.Button(buttons, text="保存", width=10, command=save).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(buttons, text="取消", width=10, command=win.destroy).pack(side=tk.LEFT)

    size_var.trace_add("write", lambda *_: refresh_preview())
    color_var.trace_add("write", lambda *_: refresh_preview())

    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    win.after(0, refresh_preview)
    size_spin.focus_set()
    win.bind("<Return>", lambda _e: save())
    win.bind("<Escape>", lambda _e: win.destroy())
