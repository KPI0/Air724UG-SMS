import serial
import threading
import tkinter as tk
import os
import sys
import winsound
import pyttsx3
import configparser
import time
import webbrowser
import winreg
import pystray
from PIL import Image
from tkinter.scrolledtext import ScrolledText
from tkinter import messagebox, ttk
from datetime import datetime, timedelta
from serial.tools import list_ports


# ====== 版本说明 V3.1.4 ======
# - 严格优先自动识别 LUAT Modem 口（description + hwid 兜底）
# - 识别不到时回退到配置串口（手动指定）
# - 串口掉线/换设备/COM 变化：自动重连 + 自动重新扫描
# - 串口设置/关于弹窗居中（模态）
# - 左下角显示当前连接状态（颜色）
# - 增加托盘功能

# ================= 配置 =================
CONFIG_FILE = "config.ini"
KEYWORDS = ["【四川安播中心】"]
LOG_DIR = "sms_logs"
TTS_FILE = "sichuan_alert.wav"
RECONNECT_INTERVAL = 2  # 秒


# ================= 语音播报开关 =================
VOICE_ENABLED = True
os.makedirs(LOG_DIR, exist_ok=True)

# ================= 读取配置 =================
config = configparser.ConfigParser()
if not os.path.exists(CONFIG_FILE):
    config["serial"] = {
        "port": "",
        "baud": "115200",
        "mode": "Auto",  # Auto / Manual
    }
    config["ui"] = {"voice_enabled": "1"}
    # 新增：关键词配置（可选）
    config["keywords"] = {"items": "|".join(KEYWORDS)}
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)

config.read(CONFIG_FILE, encoding="utf-8")
PORT = config.get("serial", "port", fallback="").strip()
BAUD = config.getint("serial", "baud", fallback=115200)
MODE = config.get("serial", "mode", fallback="Auto").strip().lower()
if MODE not in ("Auto", "Manual"):
    MODE = "Auto"

# ================= 语音播报开关（配置记忆） =================
# 默认开启；若 config.ini 存在上次状态，则以配置为准
if not config.has_section("ui"):
    config["ui"] = {"voice_enabled": "1"}
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)
    except Exception:
        pass

try:
    VOICE_ENABLED = config.getboolean("ui", "voice_enabled", fallback=True)
except Exception:
    VOICE_ENABLED = True

# ================= 关键词（配置记忆，可选） =================
# 读取 config.ini 中的 keywords.items（用 | 分隔）；不存在则使用默认 KEYWORDS
try:
    items = config.get("keywords", "items", fallback="").strip()
    if items:
        loaded = [x.strip() for x in items.split("|") if x.strip()]
        if loaded:
            KEYWORDS = loaded
except Exception:
    pass


# ================= 串口控制 =================
serial_obj = None
serial_running = True

# ================= 日志 =================
def get_log_file():
    today = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(LOG_DIR, f"sms_{today}.txt")

# ================= TTS =================
def generate_alert_voice():
    if not os.path.exists(TTS_FILE):
        engine = pyttsx3.init()
        engine.setProperty("rate", 150)
        engine.save_to_file("注意！四川安播中心预警短信，请及时查看。", TTS_FILE)
        engine.runAndWait()

generate_alert_voice()

# ================= GUI =================
root = tk.Tk()
root.withdraw()
root.minsize(500, 200)

def resource_path(relative):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, relative)
    return os.path.join(os.path.abspath("."), relative)

root.iconbitmap(resource_path("icon.ico"))

root.title("四川安播中心预警短信接收显示 V3.1.4")
root.geometry("760x520")

root.update_idletasks()
root.deiconify()


# ================= 托盘 / 退出 / 隐藏 =================
tray_icon = None
is_exiting = False

def show_window():
    root.after(0, lambda: (root.deiconify(), root.lift(), root.focus_force()))

def hide_window():
    root.after(0, root.withdraw)

def cleanup_and_exit():
    """真正退出：停止串口线程、关闭串口、停止托盘、销毁窗口"""
    global serial_running, serial_obj, is_exiting, tray_icon
    if is_exiting:
        return
    is_exiting = True

    try:
        serial_running = False
    except Exception:
        pass

    try:
        if serial_obj:
            serial_obj.close()
    except Exception:
        pass

    try:
        if tray_icon:
            tray_icon.stop()
    except Exception:
        pass

    try:
        root.after(0, root.destroy)
    except Exception:
        pass

def on_close():
    """点右上角×：隐藏到托盘，不退出"""
    hide_window()

root.protocol("WM_DELETE_WINDOW", on_close)

def create_tray():
    global tray_icon
    try:
        img = Image.open(resource_path("icon.ico"))
    except Exception:
        img = None

    menu = pystray.Menu(
        pystray.MenuItem("显示", lambda: show_window(), default=True),  # 双击托盘
        pystray.MenuItem("隐藏", lambda: hide_window()),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("退出", lambda: cleanup_and_exit()),
    )

    tray_icon = pystray.Icon("sms_tray", img, "短信接收系统", menu)
    tray_icon.run_detached()


threading.Thread(target=create_tray, daemon=True).start()

def center_window(win, parent):
    """将子窗口居中到父窗口（主窗口）上。"""
    win.update_idletasks()
    w = win.winfo_width()
    h = win.winfo_height()
    px = parent.winfo_rootx()
    py = parent.winfo_rooty()
    pw = parent.winfo_width()
    ph = parent.winfo_height()
    x = px + (pw - w) // 2
    y = py + (ph - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")

def show_about():
    """在主窗口正中显示“关于”弹窗（模态）。"""
    win = tk.Toplevel(root)
    win.title("关于")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    frame = tk.Frame(win, padx=20, pady=15)
    frame.pack(fill=tk.BOTH, expand=True)

    # 版本信息
    tk.Label(frame, text="四川安播中心预警短信接收显示", font=("微软雅黑", 12, "bold")).pack(pady=(0, 8))
    tk.Label(
        frame,
        text="版本：V3.1.4",
        justify="left",
        font=("微软雅黑", 10),
    ).pack(anchor="w")

    # 容器，用来横向放两个 Label
    link_frame = tk.Frame(frame)
    link_frame.pack(anchor="w")

    # 普通文字
    tk.Label(
        link_frame,
        text="软件地址：",
        font=("微软雅黑", 10),
    ).pack(side="left")

    # 超链接
    link = tk.Label(
        link_frame,
        text="https://github.com/KPI0/Air724UG-SMS",
        fg="blue",
        cursor="hand2",
        font=("微软雅黑", 10, "underline"),
    )
    link.pack(side="left")

    # 点击事件
    link.bind(
        "<Button-1>",
        lambda e: webbrowser.open("https://github.com/KPI0/Air724UG-SMS")
    )

    tk.Button(frame, text="确定", width=10, command=win.destroy).pack(pady=(12, 0))
    win.bind("<Escape>", lambda _e: win.destroy())

    win.update_idletasks()
    center_window(win, root)

# ===== 用 grid 布局：内容区永远不会盖住状态栏 =====
root.grid_rowconfigure(0, weight=1)   # 内容区可伸缩
root.grid_rowconfigure(1, weight=0)   # 状态栏固定
root.grid_columnconfigure(0, weight=1)

# 中间内容区域
main_frame = tk.Frame(root)
main_frame.grid(row=0, column=0, sticky="nsew")

text_area = ScrolledText(main_frame, font=("微软雅黑", 10))
text_area.pack(fill=tk.BOTH, expand=True)  # 这里用 pack 没问题，因为只在 main_frame 内部

# 底部状态栏
status_frame = tk.Frame(root)
status_frame.grid(row=1, column=0, sticky="ew")

status_var = tk.StringVar(value="🔍 启动中…")
status_label = tk.Label(status_frame, textvariable=status_var, anchor="w")
status_label.pack(side=tk.LEFT, padx=6)


def set_status(text, color="black"):
    root.after(0, lambda: (status_var.set(text), status_label.config(fg=color)))

text_area.tag_config("normal", foreground="black", font=("微软雅黑", 10))
text_area.tag_config("sms", foreground="red", font=("微软雅黑", 30))

def log(msg, tag="normal"):
    text_area.insert(tk.END, msg + "\n", tag)
    text_area.see(tk.END)
    with open(get_log_file(), "a", encoding="utf-8") as f:
        f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")

# ================= 声音 =================
def play_alert():
    global VOICE_ENABLED
    if not VOICE_ENABLED:
        return
    winsound.MessageBeep()
    winsound.PlaySound(TTS_FILE, winsound.SND_FILENAME | winsound.SND_ASYNC)

def show_sms_popup(msg: str):
    """弹窗确认后，自动显示主程序窗口"""
    global VOICE_ENABLED
    if not VOICE_ENABLED:
        return

    def _popup_and_show():
        messagebox.showinfo("预警短信", msg)  # 用户点“确定”前会阻塞
        show_window()  # 👈 关键：确认后自动打开主窗口

    try:
        root.after(0, _popup_and_show)
    except Exception:
        pass

# ================= 清空窗口 =================
def clear_window():
    text_area.delete("1.0", tk.END)

# ================= 打开日志目录 =================
def open_log_dir():
    log_path = os.path.abspath(LOG_DIR)
    if os.path.exists(log_path):
        os.startfile(log_path)   # Windows 下直接打开文件夹
    else:
        messagebox.showwarning("提示", "日志目录不存在")

# ================= 每日清空 =================
def clear_text_area_for_new_day():
    clear_window()
    log("📅 新的一天，窗口已清空")
    schedule_next_midnight_clear()

def schedule_next_midnight_clear():
    now = datetime.now()
    next_midnight = now.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)
    root.after(int((next_midnight - now).total_seconds() * 1000), clear_text_area_for_new_day)

# ================= 串口扫描 =================
def scan_com_ports_all():
    """设置窗口用：显示系统所有 COM 口"""
    return [p.device for p in list_ports.comports()]

def find_luat_best_port():
    """
    自动识别 LUAT 可用 Modem 口（最终稳定策略）：
    1) 必须是 LUAT（desc 或 hwid 中包含 LUAT）
    2) 明确排除：DIAG/NPI/MOS/DEBUG/DOWNLOAD/CP/AP 等诊断口，以及 AT 口
    3) 优先选择 description 包含 MODEM 的口
    返回： (device, desc) 或 (None, None)
    """
    exclude_tokens = [
        "DIAG", "NPI", "MOS", "DEBUG", "DOWNLOAD",
        "CP ", "CP_", "AP ", "AP_",  # 有些驱动会写 CP/AP
    ]

    candidates = []
    for p in list_ports.comports():
        dev = p.device
        desc = (p.description or "")
        hwid = (p.hwid or "")

        desc_u = desc.upper()
        hwid_u = hwid.upper()

        # 必须是 LUAT（description 或 hwid 任一包含）
        if "LUAT" not in desc_u and "LUAT" not in hwid_u:
            continue

        # 排除明显非业务口
        if any(tok in desc_u for tok in exclude_tokens):
            continue

        # 排除 AT（业务建议只用 Modem）
        # 注意：desc 可能是 "LUAT USB Device 1 AT"
        if " AT" in desc_u or desc_u.endswith("AT"):
            continue

        score = 0
        if "MODEM" in desc_u:
            score += 100
        # 轻微偏好 Device 0（很多 LUAT 的 Modem 是 0）
        if "USB DEVICE 0" in desc_u:
            score += 10

        candidates.append((score, dev, desc))

    if not candidates:
        return None, None

    candidates.sort(reverse=True, key=lambda x: x[0])
    _, dev, desc = candidates[0]
    return dev, desc

# ================= 串口线程（自动识别 + 自动重连） =================
def read_serial():
    """
    串口读取线程（严格模式）：
    - 仅当串口行中包含 [I]-[handler_sms.smsCallback] 才认为“短信有效”
    - 命中后会收集同一条短信的多行输出，合并后再进行【关键词过滤】与弹窗/播报（避免弹窗不完整）
    - 关键词过滤规则：full_msg 只要包含 KEYWORDS 任意一项即放行；否则忽略不显示/不弹窗/不播报
    - 其它所有串口日志全部忽略
    """
    global serial_obj, serial_running, PORT

    callback_prefix = "[I]-[handler_sms.smsCallback]"

    follow_lines_left = 0
    pending_parts = []
    pending_display_lines = []
    pending_deadline = 0.0
    pending_active = False

    def extract_sms_body(full_msg: str) -> str:
        if not full_msg:
            return ""
        idx = full_msg.find("【")
        if idx != -1:
            return full_msg[idx:]
        return full_msg

    def keyword_hit(full_msg: str) -> bool:
        body = extract_sms_body(full_msg)
        if not KEYWORDS:
            return True
        return any(k and (k in body) for k in KEYWORDS)

    def flush_pending():
        nonlocal pending_parts, pending_display_lines, pending_deadline, pending_active, follow_lines_left
        if not pending_active:
            return

        full_msg = "".join([p for p in pending_parts if p]).strip()

        if full_msg and keyword_hit(full_msg):
            if pending_display_lines:
                first = True
                for ln in pending_display_lines:
                    if first:
                        log(ln, tag="normal")
                        first = False
                    else:
                        log(ln, tag="sms")
            else:
                log("📩 收到短信：", tag="normal")
                log(full_msg, tag="sms")

            play_alert()
            show_sms_popup(full_msg)
        else:
            log("🚫 短信未命中关键词，已忽略", tag="normal")

        pending_parts = []
        pending_display_lines = []
        pending_deadline = 0.0
        pending_active = False
        follow_lines_left = 0

    while serial_running:
        try:
            if MODE == "Auto":
                dev, desc = find_luat_best_port()
                if not dev:
                    set_status("🔍 扫描 LUAT Modem 中…", "orange")
                    time.sleep(RECONNECT_INTERVAL)
                    continue
                PORT = dev
                set_status(f"🟡 连接中：{PORT}（{desc}） @ {BAUD}", "orange")
            else:
                if not PORT:
                    set_status("🔒 手动模式：未指定串口", "red")
                    time.sleep(RECONNECT_INTERVAL)
                    continue
                set_status(f"🟡 连接中：{PORT} @ {BAUD}", "orange")

            serial_obj = serial.Serial(PORT, BAUD, timeout=1)
            log(f"🔌 串口已连接：{PORT} @ {BAUD}")
            if MODE == "Auto":
                set_status(f"🟢 已连接 Modem：{PORT} @ {BAUD}", "green")
            else:
                set_status(f"🟢 已连接：{PORT} @ {BAUD}", "green")

            while serial_running:
                try:
                    raw = serial_obj.readline()
                except (PermissionError, OSError, serial.SerialException) as e:
                    raise e

                line = raw.decode("utf-8", "ignore").strip()
                if not line:
                    if pending_active and time.monotonic() > pending_deadline:
                        flush_pending()
                    continue

                if callback_prefix in line:
                    msg = line.split(callback_prefix, 1)[1].strip()
                    if msg:
                        pending_parts = [msg]
                        pending_display_lines = ["📩 收到短信：", msg]
                        pending_active = True
                        pending_deadline = time.monotonic() + 0.6
                        follow_lines_left = 8
                    else:
                        pending_parts = []
                        pending_display_lines = []
                        pending_active = False
                        follow_lines_left = 0
                    continue

                if follow_lines_left > 0 and pending_active:
                    has_cjk = any(0x4e00 <= ord(ch) <= 0x9fff for ch in line) or ("【" in line) or ("】" in line)
                    if has_cjk:
                        pending_parts.append(line)
                        pending_display_lines.append(line)
                        pending_deadline = time.monotonic() + 0.6
                        follow_lines_left -= 1

                        if follow_lines_left <= 0:
                            flush_pending()
                    else:
                        flush_pending()
                    continue

                continue

        except Exception as e:
            log(f"⚠️ 串口异常：{e}")
            set_status(f"🔴 断开/失败：{PORT}（自动重连中…）", "red")

            try:
                if serial_obj:
                    serial_obj.close()
            except Exception:
                pass

            if MODE == "Auto":
                PORT = ""

            time.sleep(RECONNECT_INTERVAL)

    try:
        if serial_obj:
            serial_obj.close()
    except Exception:
        pass

# ================= 串口设置窗口 =================
def open_serial_setting():
    def refresh_ports():
        ports = scan_com_ports_all()
        port_box["values"] = ports
        if ports and (port_var.get() not in ports):
            port_var.set(ports[0])

    def apply():
        global PORT, BAUD, MODE, serial_running

        MODE = mode_var.get()

        try:
            BAUD = int(baud_entry.get())
        except ValueError:
            messagebox.showerror("错误", "波特率必须是数字")
            return

        if MODE == "Manual":
            if not port_var.get():
                messagebox.showerror("错误", "手动模式必须选择串口")
                return
            PORT = port_var.get()
        else:
            PORT = ""

        config.set("serial", "mode", MODE)
        config.set("serial", "port", PORT)
        config.set("serial", "baud", str(BAUD))
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)

        set_status("🟡 应用中，重连…", "orange")
        try:
            if serial_obj:
                serial_obj.close()
        except:
            pass

        log(f"⚙️ 串口设置已更新：mode={MODE} port={PORT or '(Auto)'} baud={BAUD}")
        win.destroy()

    win = tk.Toplevel(root)
    win.title("串口设置")
    win.geometry("340x240")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    frame = tk.Frame(win, padx=12, pady=10)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="连接模式：").grid(row=0, column=0, sticky="w", pady=(0, 6))
    mode_var = tk.StringVar(value=MODE)
    mode_box = ttk.Combobox(frame, values=["Auto", "Manual"], textvariable=mode_var, state="readonly", width=18)
    mode_box.grid(row=0, column=1, sticky="w", pady=(0, 6))

    tk.Label(frame, text="串口号（手动模式）：").grid(row=1, column=0, sticky="w", pady=(0, 6))
    ports = scan_com_ports_all()
    port_var = tk.StringVar(value=PORT if PORT in ports else (ports[0] if ports else ""))
    port_box = ttk.Combobox(frame, values=ports, textvariable=port_var, state="readonly", width=18)
    port_box.grid(row=1, column=1, sticky="w", pady=(0, 6))

    tk.Label(frame, text="波特率：").grid(row=2, column=0, sticky="w", pady=(0, 6))
    baud_entry = tk.Entry(frame, width=21)
    baud_entry.insert(0, str(BAUD))
    baud_entry.grid(row=2, column=1, sticky="w", pady=(0, 6))

    btn_row = tk.Frame(frame)
    btn_row.grid(row=3, column=0, columnspan=2, pady=(10, 0), sticky="w")
    tk.Button(btn_row, text="刷新端口", width=10, command=refresh_ports).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btn_row, text="应用", width=10, command=apply).pack(side=tk.LEFT)

    tip_frame = tk.Frame(frame)
    tip_frame.grid(row=4, column=0, columnspan=2, sticky="w", pady=(12, 0))

    tk.Label(
        tip_frame,
        text="提示：",
        fg="gray",
        font=("微软雅黑", 9, "bold"),
        anchor="w",
    ).pack(anchor="w")

    tk.Label(
        tip_frame,
        text="Auto 自动优先识别 LUAT Modem",
        fg="gray",
        font=("微软雅黑", 9),
        anchor="w",
    ).pack(anchor="w")

    tk.Label(
        tip_frame,
        text="Manual 手动锁定所选 COM",
        fg="gray",
        font=("微软雅黑", 9),
        anchor="w",
    ).pack(anchor="w")

    win.update_idletasks()
    center_window(win, root)

# ================= 新增：关键词设置窗口（增加/删除/修改 + 居中模态） =================
def open_keywords_setting():
    def refresh_list(select_index=None):
        listbox.delete(0, tk.END)
        for k in KEYWORDS:
            listbox.insert(tk.END, k)
        if select_index is not None and 0 <= select_index < len(KEYWORDS):
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(select_index)
            listbox.see(select_index)

    def save_keywords_to_config():
        try:
            if not config.has_section("keywords"):
                config["keywords"] = {}
            config.set("keywords", "items", "|".join(KEYWORDS))
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                config.write(f)
        except Exception:
            pass

    def get_entry_value():
        return entry_var.get().strip()

    def on_select(_evt=None):
        sel = listbox.curselection()
        if not sel:
            return
        idx = sel[0]
        try:
            entry_var.set(KEYWORDS[idx])
        except Exception:
            pass

    def add_kw():
        global KEYWORDS
        v = get_entry_value()
        if not v:
            messagebox.showerror("错误", "关键词不能为空")
            return
        if v in KEYWORDS:
            messagebox.showwarning("提示", "该关键词已存在")
            return
        KEYWORDS.append(v)
        save_keywords_to_config()
        refresh_list(select_index=len(KEYWORDS) - 1)

    def del_kw():
        global KEYWORDS
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请选择要删除的关键词")
            return
        idx = sel[0]
        if idx < 0 or idx >= len(KEYWORDS):
            return
        KEYWORDS.pop(idx)
        save_keywords_to_config()
        entry_var.set("")
        refresh_list(select_index=min(idx, len(KEYWORDS) - 1))

    def edit_kw():
        global KEYWORDS
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请选择要修改的关键词")
            return
        idx = sel[0]
        v = get_entry_value()
        if not v:
            messagebox.showerror("错误", "关键词不能为空")
            return
        if v in KEYWORDS and KEYWORDS[idx] != v:
            messagebox.showwarning("提示", "该关键词已存在")
            return
        KEYWORDS[idx] = v
        save_keywords_to_config()
        refresh_list(select_index=idx)

    win = tk.Toplevel(root)
    win.title("关键词设置")
    win.geometry("420x290")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    frame = tk.Frame(win, padx=12, pady=10)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="关键词列表：").grid(row=0, column=0, sticky="w")

    listbox = tk.Listbox(frame, height=8, width=38)
    listbox.grid(row=1, column=0, rowspan=4, sticky="nsew", pady=(6, 0))
    listbox.bind("<<ListboxSelect>>", on_select)

    right = tk.Frame(frame)
    right.grid(row=1, column=1, sticky="n", padx=(12, 0), pady=(6, 0))

    tk.Label(right, text="关键词：").pack(anchor="w")
    entry_var = tk.StringVar()
    entry = tk.Entry(right, textvariable=entry_var, width=22)
    entry.pack(anchor="w", pady=(4, 10))

    tk.Button(right, text="增加", width=10, command=add_kw).pack(anchor="w", pady=(0, 6))
    tk.Button(right, text="删除", width=10, command=del_kw).pack(anchor="w", pady=(0, 6))
    tk.Button(right, text="修改", width=10, command=edit_kw).pack(anchor="w")


    # ===== 关键词规则提示 =====
    tip = tk.Label(
        frame,
        text="提示：关键词为空时，全部短信都会显示",
        fg="gray",
        font=("微软雅黑", 9),
        anchor="w"
    )
    tip.grid(row=5, column=0, columnspan=2, sticky="w", pady=(10, 6))

    bottom = tk.Frame(frame)
    bottom.grid(row=6, column=0, columnspan=2, sticky="e", pady=(0, 10))
    tk.Button(bottom, text="关闭", width=10, command=win.destroy).pack()

    frame.grid_columnconfigure(0, weight=1)

    refresh_list()
    win.update_idletasks()
    center_window(win, root)

# ================= 语音播报开关（菜单按钮） =================
def update_voice_menu_label():
    """刷新菜单栏语音播报按钮文案"""
    try:
        label = "🔊 语音播报" if VOICE_ENABLED else "🔇 语音播报"
        menu_bar.entryconfig(voice_menu_index, label=label)
    except Exception:
        pass

def save_voice_setting():
    """保存语音播报开关到 config.ini（用于下次启动记忆）"""
    try:
        if not config.has_section("ui"):
            config["ui"] = {}
        config.set("ui", "voice_enabled", "1" if VOICE_ENABLED else "0")
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)
    except Exception:
        pass

def toggle_voice_broadcast():
    """切换语音播报开关"""
    global VOICE_ENABLED
    VOICE_ENABLED = not VOICE_ENABLED
    update_voice_menu_label()
    save_voice_setting()
    if VOICE_ENABLED:
        log("🔊 语音播报：已开启")
    else:
        log("🔇 语音播报：已关闭")

# ================= 菜单（一级串口设置） =================
menu_bar = tk.Menu(root)

file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="清空窗口", command=clear_window)
file_menu.add_command(label="打开日志", command=open_log_dir)
file_menu.add_separator()
file_menu.add_command(label="退出", command=cleanup_and_exit)
menu_bar.add_cascade(label="文件", menu=file_menu)

# 串口设置
menu_bar.add_command(label="串口设置", command=open_serial_setting)

# 关键词设置
menu_bar.add_command(label="关键词设置", command=open_keywords_setting)

# 语音播报
voice_menu_index = menu_bar.index("end") + 1
menu_bar.add_command(label="🔊 语音播报", command=toggle_voice_broadcast)

# 帮助
help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="关于", command=show_about)
menu_bar.add_cascade(label="帮助", menu=help_menu)


root.config(menu=menu_bar)
update_voice_menu_label()

# ================= 启动 =================
schedule_next_midnight_clear()

if MODE == "Auto":
    set_status("🔍 自动模式：扫描 LUAT Modem 中…", "orange")
else:
    set_status(f"🔒 手动模式：{PORT or '未指定'} @ {BAUD}", "orange")

threading.Thread(target=read_serial, daemon=True).start()
root.mainloop()

