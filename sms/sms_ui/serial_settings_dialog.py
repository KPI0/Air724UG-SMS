import tkinter as tk
from tkinter import messagebox, ttk


def parse_positive_baud(value):
    try:
        baud = int(str(value).strip())
    except (TypeError, ValueError) as exc:
        raise ValueError("baud must be a positive integer") from exc
    if baud <= 0:
        raise ValueError("baud must be a positive integer")
    return baud


def open_serial_setting_dialog(
    parent,
    current_mode,
    current_port,
    current_baud,
    scan_ports,
    on_apply,
    center_window,
):
    win = tk.Toplevel(parent)
    win.withdraw()
    win.title("串口设置")
    win.resizable(False, False)
    win.transient(parent)
    win.grab_set()

    frame = tk.Frame(win, padx=12, pady=10)
    frame.pack(fill=tk.BOTH, expand=True)

    def refresh_ports():
        ports = scan_ports()
        port_box["values"] = ports
        if ports and port_var.get() not in ports:
            port_var.set(ports[0])

    def apply():
        mode = mode_var.get()

        try:
            baud = parse_positive_baud(baud_entry.get())
        except ValueError:
            messagebox.showerror("错误", "波特率必须是大于 0 的整数", parent=win)
            return

        if mode == "Manual":
            if not port_var.get():
                messagebox.showerror("错误", "手动模式必须选择串口", parent=win)
                return
            port = port_var.get()
        else:
            port = ""

        if on_apply(mode, port, baud) is False:
            return
        win.destroy()

    tk.Label(frame, text="连接模式：").grid(row=0, column=0, sticky="w", pady=(0, 6))
    mode_var = tk.StringVar(value=current_mode)
    mode_box = ttk.Combobox(frame, values=["Auto", "Manual"], textvariable=mode_var, state="readonly", width=18)
    mode_box.grid(row=0, column=1, sticky="w", pady=(0, 6))

    tk.Label(frame, text="串口号（手动模式）：").grid(row=1, column=0, sticky="w", pady=(0, 6))
    ports = scan_ports()
    port_var = tk.StringVar(value=current_port if current_port in ports else (ports[0] if ports else ""))
    port_box = ttk.Combobox(frame, values=ports, textvariable=port_var, state="readonly", width=18)
    port_box.grid(row=1, column=1, sticky="w", pady=(0, 6))

    tk.Label(frame, text="波特率：").grid(row=2, column=0, sticky="w", pady=(0, 6))
    baud_entry = tk.Entry(frame, width=21)
    baud_entry.insert(0, str(current_baud))
    baud_entry.grid(row=2, column=1, sticky="w", pady=(0, 6))

    btn_row = tk.Frame(frame)
    btn_row.grid(row=3, column=0, columnspan=2, pady=(10, 0))
    tk.Button(btn_row, text="刷新", width=10, command=refresh_ports).pack(side=tk.LEFT, padx=8)
    tk.Button(btn_row, text="应用", width=10, command=apply).pack(side=tk.LEFT, padx=8)
    tk.Button(btn_row, text="取消", width=10, command=win.destroy).pack(side=tk.LEFT, padx=8)

    tip_frame = tk.Frame(frame)
    tip_frame.grid(row=4, column=0, columnspan=2, sticky="w", pady=(12, 0))
    tk.Label(
        tip_frame,
        text="💡 提示：\nAuto 自动优先识别 LUAT Modem\nManual 手动锁定所选 COM",
        fg="gray",
        justify="left",
        font=("微软雅黑", 9),
        anchor="w",
    ).pack(anchor="w")

    win.update_idletasks()
    center_window(win, parent)
    win.deiconify()
    win.lift()
    win.focus_force()
    mode_box.focus_set()
    win.bind("<Return>", lambda _e: apply())
    win.bind("<Escape>", lambda _e: win.destroy())
