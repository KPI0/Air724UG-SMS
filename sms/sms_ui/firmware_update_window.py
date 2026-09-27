"""A single local firmware window; workers never access Tk widgets."""
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from sms_core.local_firmware_update import LocalFirmwareUpdater
from sms_ui.window_utils import sync_and_focus_existing_window


def open_firmware_update_window(namespace):
    current = namespace.get("firmware_update_window")
    if sync_and_focus_existing_window(current):
        return current
    updater = namespace.get("local_firmware_updater")
    if updater is None:
        updater = LocalFirmwareUpdater(namespace)
        namespace["local_firmware_updater"] = updater
    win = tk.Toplevel(namespace["root"])
    namespace["firmware_update_window"] = win
    win.withdraw()
    win.title("固件更新")
    win.resizable(True, True)
    win.minsize(460, 390)
    footer = ttk.Frame(win, padding=(18, 10, 18, 18))
    footer.pack(side="bottom", fill="x")
    viewport = ttk.Frame(win)
    viewport.pack(fill="both", expand=True)
    canvas = tk.Canvas(viewport, highlightthickness=0, width=760, height=390,
                       background=ttk.Style(win).lookup("TFrame", "background"))
    scrollbar = ttk.Scrollbar(viewport, orient="vertical", command=canvas.yview)
    canvas.pack(side="left", fill="both", expand=True)
    canvas.configure(yscrollcommand=scrollbar.set)
    body = ttk.Frame(canvas, padding=18)
    content = canvas.create_window((0, 0), window=body, anchor="nw")
    body.columnconfigure(1, weight=1)
    ttk.Label(body, text="设备固件更新", font=("Microsoft YaHei UI", 12, "bold")).grid(
        row=0, column=0, columnspan=3, sticky="w", pady=(0, 16))
    notes = []
    values = {key: tk.StringVar(master=win, value="—") for key in ("device", "current_version", "target_version", "filename")}
    for row, (key, label) in enumerate((("device", "当前设备"), ("current_version", "当前版本"),
                                      ("target_version", "目标版本"), ("filename", "升级文件")), start=2):
        ttk.Label(body, text=label).grid(row=row, column=0, sticky="nw", padx=(0, 12), pady=5)
        value = (ttk.Entry(body, textvariable=values[key], state="readonly") if key == "filename"
                 else ttk.Label(body, textvariable=values[key], wraplength=380))
        value.grid(row=row, column=1, columnspan=2, sticky="ew", pady=5)
        if key != "filename":
            notes.append((value, 130))
    actions = ttk.Frame(body)
    actions.grid(row=6, column=0, columnspan=3, sticky="w", pady=(8, 12))

    def select_file():
        if updater.snapshot()["busy"]:
            return
        path = filedialog.askopenfilename(parent=win, title="选择本地固件",
            filetypes=[("Air724UG 固件", "*.dfota.bin *.air724ota"), ("所有文件", "*.*")])
        if path and not closed:
            updater.prepare(path)
            render()

    read_button = ttk.Button(actions, text="读取设备", command=lambda: updater.prepare())
    read_button.pack(side="left", padx=(0, 8))
    file_button = ttk.Button(actions, text="选择固件…", command=select_file)
    file_button.pack(side="left")
    progress = ttk.Progressbar(body, maximum=100, mode="determinate")
    progress.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(0, 8))
    status = tk.StringVar(master=win)
    status_label = ttk.Label(body, textvariable=status, wraplength=490)
    status_label.grid(row=8, column=0, columnspan=3, sticky="new", pady=(0, 12))
    notes.append((status_label, 36))
    hint = ttk.Label(body, text="支持 .dfota.bin / .air724ota，目标版本须高于当前版本。更新期间请保持供电，暂停拨号和发信。", wraplength=490)
    hint.grid(row=9, column=0, columnspan=3, sticky="ew", pady=(0, 14))
    notes.append((hint, 36))

    def start():
        state = updater.snapshot()
        if not state["can_start"]:
            return
        detail = f"设备：{state['device']}\n版本：{state['current_version']} → {state['target_version']}"
        if state["config_changed"]:
            detail += "\n\n包内配置与设备不同，将随固件一并更新。"
        confirmed = messagebox.askyesno("开始固件更新", detail + "\n\n设备将安装并重启，请保持供电和 USB 连接。是否开始？", parent=win)
        if confirmed and not closed:
            updater.start()
            render()

    timer = None
    last_state = None
    closed = False

    def close():
        namespace["firmware_update_window"] = None
        win.destroy()

    def destroyed(event):
        nonlocal timer, status, closed
        if event.widget is win:
            closed = True
            if timer is not None:
                win.after_cancel(timer)
                timer = None
            if namespace.get("firmware_update_window") is win:
                namespace["firmware_update_window"] = None
            # Release Tk variables here, before closure cycles can be collected
            # by a worker thread after this window has gone away.
            values.clear()
            status = None

    ttk.Button(footer, text="关闭", width=8, padding=(12, 6), command=close).pack(side="right", padx=(12, 0))
    start_button = ttk.Button(footer, text="开始更新", width=8, padding=(12, 6), command=start)
    start_button.pack(side="right", padx=(12, 0))
    cancel_button = ttk.Button(footer, text="取消更新", width=8, padding=(12, 6), command=updater.cancel)
    cancel_button.pack(side="right")

    def render():
        nonlocal last_state
        if closed:
            return
        state = updater.snapshot()
        if state == last_state:
            return
        last_state = state
        for key, variable in values.items():
            variable.set(state[key] or "—")
        progress.configure(value=state["progress"])
        status.set(state["message"] + ("\n关闭窗口后任务继续，可从设置重新打开查看。" if state["busy"] else ""))
        for button in (read_button, file_button):
            button.configure(state="disabled" if state["busy"] else "normal")
        start_button.configure(state="normal" if state["can_start"] else "disabled")
        cancel_button.configure(state="normal" if state["can_cancel"] else "disabled")

    def poll():
        nonlocal timer
        if closed or namespace.get("is_exiting"):
            timer = None
            return
        render()
        timer = win.after(200, poll)

    def resize(event):
        canvas.itemconfigure(content, width=event.width)
        for label, gutter in notes:
            label.configure(wraplength=max(100, event.width - gutter))
        update_scroll()

    def update_scroll(event=None):
        canvas.configure(scrollregion=canvas.bbox("all"))
        overflow = body.winfo_reqheight() > canvas.winfo_height()
        if overflow and not scrollbar.winfo_manager():
            scrollbar.pack(side="right", fill="y", before=canvas)
        elif not overflow and scrollbar.winfo_manager():
            scrollbar.pack_forget()
        min_height = min(max(390, body.winfo_reqheight() + footer.winfo_reqheight()),
                         win.winfo_screenheight() - 80)
        win.minsize(max(460, footer.winfo_reqwidth()), min_height)

    def scroll(event):
        if body.winfo_reqheight() > canvas.winfo_height():
            delta = (-1 if event.num == 4 else 1) if event.num in (4, 5) else -int(event.delta / 120)
            canvas.yview_scroll(delta, "units")
            return "break"

    def reveal_focus(event):
        widget = event.widget
        if str(widget).startswith(str(body) + "."):
            top = widget.winfo_rooty() - body.winfo_rooty()
            bottom = top + widget.winfo_height()
            visible_top = canvas.canvasy(0)
            if top < visible_top:
                canvas.yview_moveto(top / max(1, body.winfo_height()))
            elif bottom > visible_top + canvas.winfo_height():
                canvas.yview_moveto((bottom - canvas.winfo_height()) / max(1, body.winfo_height()))

    canvas.bind("<Configure>", resize)
    body.bind("<Configure>", update_scroll)
    win.bind("<MouseWheel>", scroll)
    win.bind("<Button-4>", scroll)
    win.bind("<Button-5>", scroll)
    win.bind("<FocusIn>", reveal_focus)
    win.bind("<Destroy>", destroyed)
    win.bind("<Escape>", lambda _event: close())
    win.protocol("WM_DELETE_WINDOW", close)
    win.geometry("760x450")
    win.update_idletasks()
    center = namespace.get("center_window")
    if center:
        center(win, namespace["root"])
    poll()
    win.deiconify()
    win.lift()
    win.focus_set()
    return win
