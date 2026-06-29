import tkinter as tk
import webbrowser
from tkinter import messagebox


def open_about_dialog(parent, app_version, project_url, center_window):
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("关于")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    bottom_line = tk.Frame(win, height=1, bg="#d4d4d4")
    bottom_line.pack(side="bottom", fill="x")
    bottom_line.pack_propagate(False)

    frame = tk.Frame(win, padx=20, pady=15)
    frame.pack(side="top", fill=tk.BOTH, expand=True)

    tk.Label(frame, text="短信监听系统", font=("微软雅黑", 12, "bold")).pack(pady=(0, 8))
    tk.Label(
        frame,
        text=f"版本：v{app_version}",
        justify="left",
        font=("微软雅黑", 10),
    ).pack(anchor="w")

    link_frame = tk.Frame(frame)
    link_frame.pack(anchor="w")

    tk.Label(
        link_frame,
        text="软件地址：",
        font=("微软雅黑", 10),
    ).pack(side="left")

    link = tk.Label(
        link_frame,
        text=project_url,
        fg="blue",
        cursor="hand2",
        font=("微软雅黑", 10, "underline"),
    )
    link.pack(side="left")
    link.bind("<Button-1>", lambda _e: webbrowser.open(project_url))

    tk.Button(frame, text="确定", width=10, command=win.destroy).pack(pady=(12, 0))

    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    win.bind("<Escape>", lambda _e: win.destroy())


def open_voice_text_dialog(parent, current_text, on_preview, on_save, center_window):
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("语音播报自定义")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    tk.Label(win, text="播报内容：").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 4))

    text = tk.Text(win, width=42, height=4, font=("微软雅黑", 10))
    text.grid(row=1, column=0, columnspan=3, padx=10, pady=(0, 10))
    text.insert("1.0", current_text)

    def read_text():
        value = text.get("1.0", "end").strip()
        if not value:
            messagebox.showerror("错误", "播报内容不能为空", parent=win)
            return None
        return value

    def preview():
        value = read_text()
        if value is not None:
            on_preview(value)

    def save():
        value = read_text()
        if value is None:
            return
        on_save(value)
        win.destroy()

    tk.Button(win, text="试听", width=10, command=preview).grid(
        row=2, column=0, padx=10, pady=(0, 10), sticky="w"
    )
    tk.Button(win, text="保存", width=10, command=save).grid(row=2, column=1, pady=(0, 10))
    tk.Button(win, text="取消", width=10, command=win.destroy).grid(
        row=2, column=2, padx=10, pady=(0, 10), sticky="e"
    )

    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    text.focus_set()
    win.bind("<Escape>", lambda _e: win.destroy())


def open_log_cleanup_dialog(
    parent,
    current_days,
    interval_hours,
    on_apply,
    center_window,
):
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("日志自动清理")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frame = tk.Frame(win, padx=14, pady=12)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="保留最近 N 天日志：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w")

    days_var = tk.StringVar(value=str(current_days))
    days_entry = tk.Entry(frame, textvariable=days_var, width=10)
    days_entry.grid(row=0, column=1, sticky="w", padx=(8, 0))
    tk.Label(frame, text="天", font=("微软雅黑", 10)).grid(row=0, column=2, sticky="w", padx=(6, 0))

    tk.Label(
        frame,
        text="说明：会删除 sms_logs 目录下超过 N 天的 sms_*.txt 日志（含 sms_system / sms_COMx）。",
        fg="gray",
        font=("微软雅黑", 9),
        wraplength=360,
        justify="left",
    ).grid(row=1, column=0, columnspan=3, sticky="w", pady=(10, 6))

    def apply_cleanup():
        try:
            days = int(days_var.get().strip())
            if days < 0:
                raise ValueError
        except Exception:
            messagebox.showerror("错误", "天数必须是非负整数（例如 30）", parent=win)
            return

        if not messagebox.askyesno(
            "确认",
            f"确定设置为自动清理，并保留最近 {days} 天日志吗？",
            parent=win,
        ):
            return

        on_apply(days)
        messagebox.showinfo(
            "完成",
            "已启用自动日志清理（程序运行期间会定期清理）。",
            parent=win,
        )
        win.destroy()

    btns = tk.Frame(frame)
    btns.grid(row=2, column=0, columnspan=3, sticky="e", pady=(10, 0))
    tk.Button(btns, text="确认", width=10, command=apply_cleanup).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btns, text="取消", width=10, command=win.destroy).pack(side=tk.LEFT)

    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    days_entry.focus_set()
    win.bind("<Return>", lambda _e: apply_cleanup())
    win.bind("<Escape>", lambda _e: win.destroy())


def open_desktop_shortcut_dialog(parent, default_name, on_apply, on_save, center_window):
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("创建桌面快捷方式")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    bottom_line = tk.Frame(win, height=1, bg="#d4d4d4")
    bottom_line.pack(side="bottom", fill="x")
    bottom_line.pack_propagate(False)

    frame = tk.Frame(win, padx=14, pady=12)
    frame.pack(side="top", fill=tk.BOTH, expand=True)

    tk.Label(frame, text="快捷方式名称：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w")

    name_var = tk.StringVar(value=default_name)
    entry = tk.Entry(frame, textvariable=name_var, width=28)
    entry.grid(row=1, column=0, pady=(6, 12), sticky="w")

    def read_name():
        name = name_var.get().strip()
        if not name:
            messagebox.showerror("错误", "名称不能为空", parent=win)
            return None
        return name

    def apply_now():
        name = read_name()
        if name is None:
            return
        try:
            on_apply(name)
            messagebox.showinfo("完成", "桌面快捷方式已创建", parent=win)
        except Exception as exc:
            messagebox.showerror("失败", str(exc), parent=win)

    def save_only():
        name = read_name()
        if name is None:
            return
        try:
            result = on_save(name)
            if result is False:
                raise RuntimeError("配置保存失败")
            messagebox.showinfo("已保存", "名称已保存，下次可直接应用", parent=win)
        except Exception as exc:
            messagebox.showerror("失败", str(exc), parent=win)

    btns = tk.Frame(frame)
    btns.grid(row=2, column=0, sticky="e")
    tk.Button(btns, text="应用", width=10, command=apply_now).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btns, text="保存", width=10, command=save_only).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btns, text="取消", width=10, command=win.destroy).pack(side=tk.LEFT, padx=(0, 8))

    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    entry.focus_set()
    win.bind("<Return>", lambda _e: apply_now())
    win.bind("<Escape>", lambda _e: win.destroy())
