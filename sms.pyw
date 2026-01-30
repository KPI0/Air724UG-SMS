# ---- 标准库 ----
import configparser
import json
import os
import re
import socket
import ssl
import subprocess
import sys
import tempfile
import threading
import time
import urllib.error
import urllib.request
import winsound
import webbrowser
import queue
from datetime import datetime, timedelta

# ---- 第三方库 ----
import serial
import pystray
import pyttsx3
from PIL import Image
from serial.tools import list_ports

# ---- tkinter ----
import tkinter as tk
from tkinter import messagebox, ttk, colorchooser
from tkinter.scrolledtext import ScrolledText

# ================= 配置 =================
CONFIG_FILE = "config.ini"  # 软件配置文件
LOG_DIR = "sms_logs" # 短信日志文件夹
TTS_DIR = "tts" # 语音播报文件夹
TTS_FILE = os.path.join(TTS_DIR, "alert.wav")
RECONNECT_INTERVAL = 2  # 秒
APP_VERSION = "3.2.8"  # 软件版本号
GITHUB_OWNER = "KPI0"
GITHUB_REPO = "Air724UG-SMS"

# 启动参数：开机自启时是否默认最小化到托盘
AUTOSTART_FLAG = "--autostart"
START_MINIMIZED = AUTOSTART_FLAG in sys.argv

# ================= 开机自启 =================
def get_startup_dir():
    # 用环境变量 APPDATA 获取当前用户启动文件夹
    appdata = os.environ.get("APPDATA", "")
    return os.path.join(appdata, r"Microsoft\Windows\Start Menu\Programs\Startup")

def get_startup_lnk():
    return os.path.join(get_startup_dir(), "sms.lnk")

def is_autostart_enabled():
    return os.path.exists(get_startup_lnk())

def _get_launch_target_and_args():
    """
    返回 (target_path, arguments, working_dir)
    - 打包 exe：target=exe, args=""
    - 脚本运行：target=pythonw.exe, args=脚本路径（不带引号）
    """
    if getattr(sys, "frozen", False):
        exe_path = sys.executable
        return exe_path, "", os.path.dirname(exe_path)

    # 脚本模式：用 pythonw.exe 最稳（不弹黑窗）
    pyw = os.path.join(os.path.dirname(sys.executable), "pythonw.exe")
    if not os.path.exists(pyw):
        pyw = sys.executable

    script_path = os.path.abspath(sys.argv[0])

    return pyw, script_path, os.path.dirname(script_path)

def create_startup_shortcut():
    startup_dir = get_startup_dir()
    os.makedirs(startup_dir, exist_ok=True)

    lnk_path = get_startup_lnk()
    target, args, workdir = _get_launch_target_and_args()

    # vbs 用双引号包裹字符串；内部双引号要变成 ""
    def vbs_quote(s: str) -> str:
        return '"' + s.replace('"', '""') + '"'

    # 生成临时 vbs（wscript 执行默认无窗口，不闪）
    vbs = f'''
        Set WshShell = CreateObject("WScript.Shell")
        Set Shortcut = WshShell.CreateShortcut({vbs_quote(lnk_path)})
        Shortcut.TargetPath = {vbs_quote(target)}
        Shortcut.WorkingDirectory = {vbs_quote(workdir)}
        Shortcut.WindowStyle = 1
        '''

    if args:
        # 脚本模式：pythonw.exe "脚本路径" --autostart
        arg_line = f'"{args}" {AUTOSTART_FLAG}'
        vbs += f'Shortcut.Arguments = {vbs_quote(arg_line)}\n'
    else:
        # exe 模式：sms.exe --autostart
        vbs += f'Shortcut.Arguments = {vbs_quote(AUTOSTART_FLAG)}\n'

    vbs += 'Shortcut.Save\n'

    vbs_path = os.path.join(tempfile.gettempdir(), "sms_autostart_create.vbs")
    with open(vbs_path, "w", encoding="mbcs") as f:
        f.write(vbs)

    # 用 wscript.exe 执行（无控制台窗口）
    r = subprocess.run(
        ["wscript.exe", "//Nologo", vbs_path],
        capture_output=True,
        text=True
    )

    # 校验是否真的创建成功
    if not os.path.exists(lnk_path):
        err = ((r.stderr or "") + "\n" + (r.stdout or "")).strip()
        raise RuntimeError(
            "创建快捷方式失败：\n"
            f"returncode={r.returncode}\n"
            f"{err or '（stderr/stdout 为空，lnk 未生成）'}"
        )

def remove_startup_shortcut():
    lnk = get_startup_lnk()
    if os.path.exists(lnk):
        os.remove(lnk)

# ================= 语音播报开关 =================
VOICE_ENABLED = True
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(TTS_DIR, exist_ok=True)

# ================= 读取配置 =================
config = configparser.ConfigParser()
if not os.path.exists(CONFIG_FILE):
    config["serial"] = {
        "port": "",
        "baud": "115200",
        "mode": "Auto",  # Auto / Manual
    }

    config["ui"] = {
    "voice_enabled": "1",         # 0=关闭语音播报，1=打开语音播报（默认）
    "voice_text": "注意！四川安播中心预警短信，请及时查看。",   # 默认语音播报内容
    "allow_multi_instance": "0",  # 0=禁止程序多开（默认），1=允许程序多开
    "auto_log_cleanup": "1",      # 0=关闭日志清理，1=打开日志清理（默认）
    "log_retention_days": "30",   # 日志保留时间，单位：天
    "desktop_shortcut_name": "短信监听系统",  # 默认桌面快捷方式名称
    "keywords": "【四川安播中心】",  # 默认关键词
    "sms_font_size": "30",        # 默认字体大小
    "sms_font_color": "#ff0000",  # 默认字体颜色

    }

    # 更新代理配置
    config["update"] = {
        "api_proxy_base": "https://github-api.daybyday.top/",
        "proxy_base": "https://gh-proxy.com/",
    }

    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)

config.read(CONFIG_FILE, encoding="utf-8")

# ===== 语音播报内容（从配置读取）=====
DEFAULT_VOICE_TEXT = "注意！四川安播中心预警短信，请及时查看。"
try:
    VOICE_TEXT = config.get("ui", "voice_text", fallback=DEFAULT_VOICE_TEXT).strip()
    if not VOICE_TEXT:
        VOICE_TEXT = DEFAULT_VOICE_TEXT
except Exception:
    VOICE_TEXT = DEFAULT_VOICE_TEXT

# ===== 自动日志清理（从配置读取）=====
try:
    AUTO_LOG_CLEANUP = config.getboolean("ui", "auto_log_cleanup", fallback=True)
except Exception:
    AUTO_LOG_CLEANUP = True

try:
    LOG_RETENTION_DAYS = config.getint("ui", "log_retention_days", fallback=30)
except Exception:
    LOG_RETENTION_DAYS = 30

try:
    ALLOW_MULTI_INSTANCE = config.getboolean(
        "ui", "allow_multi_instance", fallback=False
    )
except Exception:
    ALLOW_MULTI_INSTANCE = False

PORT = config.get("serial", "port", fallback="").strip()
BAUD = config.getint("serial", "baud", fallback=115200)
MODE = config.get("serial", "mode", fallback="Auto").strip().lower()
if MODE not in ("auto", "manual"):
    MODE = "auto"
MODE = "Auto" if MODE == "auto" else "Manual"

# ===== 短信字体（从配置读取）=====
try:
    SMS_FONT_SIZE = config.getint("ui", "sms_font_size", fallback=30)
except Exception:
    SMS_FONT_SIZE = 30

try:
    SMS_FONT_COLOR = config.get("ui", "sms_font_color", fallback="#ff0000").strip() or "#ff0000"
except Exception:
    SMS_FONT_COLOR = "#ff0000"

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

# ================= 关键词（配置记忆） =================
# 读取 config.ini 中的 ui.keywords（用 | 分隔）
# 注意：允许 items 为空（表示不过滤：显示全部短信）
KEYWORDS = []
try:
    raw = config.get("ui", "keywords", fallback="").strip()
    KEYWORDS = [x.strip() for x in raw.split("|") if x.strip()]
except Exception:
    pass

# ================= 串口控制 =================
serial_obj = None
serial_running = True

# ================= 全局变量 =================
PENDING_UI_LOGS = []  # 用于 text_area 未创建前缓存要显示到窗口的提示
LOG_PREFIX = "system"
AUTO_CLEANUP_INTERVAL_HOURS = 24 # 自动清理频率：24小时一次
AUTO_CLEANUP_AFTER_ID = None     # 记录 after() 的任务ID，用于避免重复定时器
SERIAL_DEBUG_ENABLED = False
serial_debug_queue = queue.Queue(maxsize=5000)  # 防止无限涨
serial_debug_win = None
serial_debug_text = None
serial_debug_drop_count = 0

# ================= 日志 =================
def get_log_file():
    today = datetime.now().strftime("%Y-%m-%d")
    return os.path.join(
        LOG_DIR,
        f"sms_{LOG_PREFIX}_{today}.txt"
    )

def log_file_only(msg: str):
    """系统级日志：固定写入 sms_system_YYYY-MM-DD.txt（不依赖 text_area / LOG_PREFIX）"""
    try:
        today = datetime.now().strftime("%Y-%m-%d")
        system_log = os.path.join(LOG_DIR, f"sms_system_{today}.txt")
        with open(system_log, "a", encoding="utf-8") as f:
            f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")
    except Exception:
        pass

def ui_only(msg: str, tag="normal"):
    """只显示到窗口，不写任何日志文件（不写 COM 日志）"""
    try:
        text_area.insert(tk.END, msg + "\n", tag)
        text_area.see(tk.END)
    except Exception:
        pass

def log_early(msg: str, tag: str = "normal"):
    """早期日志：先写文件，再缓存，等 text_area 创建后补到窗口"""
    log_file_only(msg)
    try:
        PENDING_UI_LOGS.append((msg, tag))
    except Exception:
        pass

def system_ui(message: str, tag="normal"):
    """
    - UI 未就绪/窗口已销毁：走 log_early（system + 缓存，等 UI 创建后补）
    - UI 就绪：写 system，然后
        - 主线程：直接 ui_only
        - 非主线程：root.after 调回主线程 ui_only
    """
    # --- 1) 判断 root 是否可用 ---
    root_ok = False
    try:
        root_ok = ("root" in globals()) and (root is not None) and root.winfo_exists()
    except Exception:
        root_ok = False

    # --- 2) 判断 text_area 是否可用 ---
    text_ok = False
    try:
        text_ok = ("text_area" in globals()) and (text_area is not None) and text_area.winfo_exists()
    except Exception:
        text_ok = False

    # UI 不可用：直接走早期日志（system + 缓存）
    if not (root_ok and text_ok):
        log_early(message, tag)
        return

    # --- 3) UI 可用：写 system（只写一次） ---
    log_file_only(message)

    # --- 4) UI 更新：区分主线程/非主线程 ---
    def _do_ui():
        try:
            ui_only(message, tag)
        except Exception:
            pass

    try:
        if threading.current_thread() is threading.main_thread():
            _do_ui()
        else:
            root.after(0, _do_ui)
    except Exception:
        # after 不可用/竞态：退回 early
        log_early(message, tag)

def port_ui(message: str, tag="normal"):
    """
    线程安全写“COM 分日志 + 窗口”
    - UI 不可用：退回 log_early
    - UI 可用：走 log()
    """
    # --- 1) 判断 root 是否可用 ---
    root_ok = False
    try:
        root_ok = ("root" in globals()) and (root is not None) and root.winfo_exists()
    except Exception:
        root_ok = False

    # --- 2) 判断 text_area 是否可用 ---
    text_ok = False
    try:
        text_ok = ("text_area" in globals()) and (text_area is not None) and text_area.winfo_exists()
    except Exception:
        text_ok = False

    if not (root_ok and text_ok):
        log_early(message, tag)
        return

    def _do():
        try:
            log(message, tag=tag)   # 关键：走 log() -> get_log_file() -> sms_{LOG_PREFIX}_*.txt
        except Exception:
            # 极端情况：log 失败也别丢，至少写 system
            log_file_only(message)

    try:
        if threading.current_thread() is threading.main_thread():
            _do()
        else:
            root.after(0, _do)
    except Exception:
        log_early(message, tag)

def set_autostart(enable: bool):
    try:
        if enable:
            create_startup_shortcut()
            msg = "🚀 开机自启：已打开"
        else:
            remove_startup_shortcut()
            msg = "⛔ 开机自启：已关闭"

        system_ui(msg, "normal")

    except Exception as e:
        messagebox.showerror("错误", f"设置开机自启失败：\n{e}")

# ================= TTS语音播报 =================
def generate_alert_voice(force: bool = False):
    """生成/更新语音播报 wav（VOICE_TEXT）"""
    if (not force) and os.path.exists(TTS_FILE):
        return

    try:
        os.makedirs(os.path.dirname(TTS_FILE), exist_ok=True)
        engine = pyttsx3.init()
        engine.setProperty("rate", 150)
        engine.save_to_file(VOICE_TEXT, TTS_FILE)
        engine.runAndWait()
    except Exception as e:
        log_file_only(f"TTS 生成失败，使用系统声音兜底：{e}")

# ================= 获取桌面路径 =================
def get_desktop_dir():
    # 优先用 Windows 注册表拿 “真实桌面路径”，兼容 OneDrive 重定向
    try:
        import winreg
        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders"
        ) as k:
            desktop = winreg.QueryValueEx(k, "Desktop")[0]
        desktop = os.path.expandvars(desktop)
        if desktop and os.path.isdir(desktop):
            return desktop
    except Exception:
        pass

    # 兜底：传统路径
    return os.path.join(os.path.expanduser("~"), "Desktop")

# ================= 创建桌面快捷方式 =================  
def create_desktop_shortcut(shortcut_name: str):
    shortcut_name = re.sub(r'[\\/:*?"<>|]', "_", shortcut_name.strip())
    desktop = get_desktop_dir()
    os.makedirs(desktop, exist_ok=True)

    if not shortcut_name.lower().endswith(".lnk"):
        shortcut_name += ".lnk"

    lnk_path = os.path.join(desktop, shortcut_name)

    target, args, workdir = _get_launch_target_and_args()

    def vbs_quote(s: str) -> str:
        return '"' + s.replace('"', '""') + '"'

    vbs = f'''
Set WshShell = CreateObject("WScript.Shell")
Set Shortcut = WshShell.CreateShortcut({vbs_quote(lnk_path)})
Shortcut.TargetPath = {vbs_quote(target)}
Shortcut.WorkingDirectory = {vbs_quote(workdir)}
Shortcut.WindowStyle = 1
'''

    if args:
        vbs += f'Shortcut.Arguments = {vbs_quote(args)}\n'

    vbs += 'Shortcut.Save\n'

    vbs_path = os.path.join(tempfile.gettempdir(), "sms_desktop_shortcut.vbs")
    with open(vbs_path, "w", encoding="mbcs") as f:
        f.write(vbs)

    # 只执行一次
    r = subprocess.run(
        ["cscript.exe", "//Nologo", vbs_path],
        capture_output=True,
        text=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
    )

    # 校验必须在函数内部
    if not os.path.exists(lnk_path):
        detail = ((r.stderr or "") + "\n" + (r.stdout or "")).strip()
        raise RuntimeError(
            "桌面快捷方式创建失败：\n" +
            (detail or "（cscript 未返回错误信息，但 .lnk 未生成）")
        )

def save_voice_text_setting():
    try:
        if "ui" not in config:
            config["ui"] = {}
        config.set("ui", "voice_text", VOICE_TEXT)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)
    except Exception:
        pass

def save_sms_font_setting():
    try:
        if not config.has_section("ui"):
            config["ui"] = {}
        config.set("ui", "sms_font_size", str(SMS_FONT_SIZE))
        config.set("ui", "sms_font_color", SMS_FONT_COLOR)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)
    except Exception:
        pass

# ================= 短信字体设置 =================  
def open_sms_font_dialog():
    win = tk.Toplevel(root)
    win.withdraw()
    win.title("短信字体设置")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    frame = tk.Frame(win, padx=14, pady=12)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="字号：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w")

    size_var = tk.StringVar(value=str(SMS_FONT_SIZE))
    size_spin = tk.Spinbox(frame, from_=8, to=72, width=8, textvariable=size_var)
    size_spin.grid(row=0, column=1, sticky="w", padx=(8, 0))

    tk.Label(frame, text="颜色：", font=("微软雅黑", 10)).grid(row=1, column=0, sticky="w", pady=(10, 0))

    color_var = tk.StringVar(value=SMS_FONT_COLOR)
    color_entry = tk.Entry(frame, textvariable=color_var, width=14)
    color_entry.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(10, 0))

    # ===== 预览区：固定尺寸，不随字体撑大窗口 =====
    preview_box = tk.LabelFrame(frame, text="预览", padx=8, pady=8)
    preview_box.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(12, 0))
    preview_box.grid_columnconfigure(0, weight=1)

    preview_canvas = tk.Canvas(preview_box, width=560, height=110, highlightthickness=1)
    preview_canvas.grid(row=0, column=0, sticky="ew")

    PREVIEW_TEXT = "短信内容"

    def refresh_preview():
        preview_canvas.update_idletasks()

        try:
            s = int(size_var.get().strip())
        except Exception:
            s = SMS_FONT_SIZE

        c = (color_var.get().strip() or SMS_FONT_COLOR)

        # 预览用字号：避免裁剪（高度的 70% 比较合适）
        max_size = max(8, int(preview_canvas.winfo_height() * 0.7))
        s_preview = min(s, max_size)

        preview_canvas.delete("all")
        try:
            preview_canvas.create_text(
                preview_canvas.winfo_width() // 2,
                preview_canvas.winfo_height() // 2,
                text=PREVIEW_TEXT,
                anchor="c",
                font=("微软雅黑", s_preview),
                fill=c
            )
        except Exception:
            preview_canvas.create_text(
                preview_canvas.winfo_width() // 2,
                preview_canvas.winfo_height() // 2,
                text=PREVIEW_TEXT,
                anchor="c",
                font=("微软雅黑", 30),
                fill="#ff0000"
            )

    def pick_color():
        c = color_var.get().strip() or SMS_FONT_COLOR
        
        win.lift()
        win.after(0, lambda: win.lift())

        # 关键：临时释放 grab，避免系统颜色对话框闪烁/抢焦点异常
        try:
            win.grab_release()
        except Exception:
            pass

        # 关键：指定 parent，避免额外的“左上角小框/幽灵窗口”
        chosen = colorchooser.askcolor(parent=win, initialcolor=c, title="选择短信颜色")

        # 选完后把模态抓取恢复
        try:
            win.grab_set()
        except Exception:
            pass

        win.lift()
        win.after(0, lambda: win.lift())

        if chosen and chosen[1]:
            color_var.set(chosen[1])  # #RRGGBB
            refresh_preview()

    tk.Button(frame, text="选颜色", width=10, command=pick_color).grid(row=1, column=2, padx=(8, 0), pady=(10, 0))

    def do_save():
        global SMS_FONT_SIZE, SMS_FONT_COLOR

        try:
            s = int(size_var.get().strip())
            if s < 8 or s > 72:
                raise ValueError
        except Exception:
            messagebox.showerror("错误", "字号必须是 8~72 的整数")
            return

        c = color_var.get().strip() or "#ff0000"

        SMS_FONT_SIZE = s
        SMS_FONT_COLOR = c

        save_sms_font_setting()
        apply_sms_font_style()

        system_ui(f"🎨 已更新短信字体：字号 {SMS_FONT_SIZE}，颜色 {SMS_FONT_COLOR}", "normal")
        win.destroy()

    btns = tk.Frame(frame)
    btns.grid(row=3, column=0, columnspan=3, sticky="e", pady=(14, 0))
    frame.grid_columnconfigure(1, weight=1)
    
    tk.Button(btns, text="保存", width=10, command=do_save).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btns, text="取消", width=10, command=win.destroy).pack(side=tk.LEFT)

    # 交互：改值即更新预览
    size_var.trace_add("write", lambda *_: refresh_preview())
    color_var.trace_add("write", lambda *_: refresh_preview())
    
    win.update_idletasks()
    center_window(win, root)
    win.deiconify()
    win.lift()
    win.focus_force()
    win.after(0, refresh_preview)
    size_spin.focus_set()
    win.bind("<Return>", lambda _e: do_save())
    win.bind("<Escape>", lambda _e: win.destroy())

# ================= 串口调试 =================
def open_serial_debug_window():
    global serial_debug_win, serial_debug_text

    if serial_debug_win is not None and serial_debug_win.winfo_exists():
        serial_debug_win.deiconify()
        serial_debug_win.lift()
        serial_debug_win.focus_force()
        return

    serial_debug_win = tk.Toplevel(root)
    serial_debug_win.withdraw()
    serial_debug_win.title("串口调试")
    serial_debug_win.geometry("900x520")
    serial_debug_win.minsize(630, 300)
    serial_debug_win.lift()
    serial_debug_win.focus_force()
    top = ttk.Frame(serial_debug_win)
    top.pack(fill="x", padx=8, pady=6)

    enabled_var = tk.BooleanVar(value=SERIAL_DEBUG_ENABLED)

    def _toggle():
        global SERIAL_DEBUG_ENABLED
        SERIAL_DEBUG_ENABLED = bool(enabled_var.get())
        _update_state_label()

    chk = ttk.Checkbutton(
        top,
        text="启用原始输出旁路（不做任何过滤）",
        variable=enabled_var,
        command=_toggle
    )
    chk.pack(side="left")
    all_debug_lines = []   # list[str]
    MAX_STORE_LINES = 20000  # 防止内存无限增长

    def _clear():
        all_debug_lines.clear()  # 清缓存
        serial_debug_text.config(state="normal")
        serial_debug_text.delete("1.0", "end")
        serial_debug_text.config(state="disabled")

    ttk.Button(top, text="清空", width=8, command=_clear).pack(side="left", padx=8)

    # 状态 + 暂停/继续
    paused_var = tk.BooleanVar(value=False)
    pause_banner_shown = False  # 防止重复插入“已暂停显示”提示

    btn_pause = ttk.Button(top, text="⏸ 暂停", width=8)
    btn_pause.pack(side="left")

    # ===== 右侧筛选区（整体靠右）=====
    right_frame = ttk.Frame(top)
    right_frame.pack(side="right", padx=(8, 8))

    filter_var = tk.StringVar(value="")

    ttk.Label(right_frame, text="筛选：").grid(row=0, column=0, padx=(0, 4))
    filter_entry = ttk.Entry(right_frame, textvariable=filter_var, width=16)
    filter_entry.grid(row=0, column=1, padx=(0, 6))

    def _clear_filter():
        filter_var.set("")
        _redraw_by_filter()

    def _redraw_by_filter():
        kw = filter_var.get().strip()
        serial_debug_text.config(state="normal")
        serial_debug_text.delete("1.0", "end")

        for ln in all_debug_lines:
            if kw and (kw not in ln):
                continue
            if not ln.endswith("\n"):
                ln += "\n"
            serial_debug_text.insert("end", ln)

        serial_debug_text.see("end")
        serial_debug_text.config(state="disabled")

    filter_var.trace_add("write", lambda *_: _redraw_by_filter())

    ttk.Button(right_frame,text="清除筛选",width=8,command=_clear_filter).grid(row=0, column=2)

    # 根据旁路/暂停状态刷新状态标签
    def _update_state_label():
        running = bool(enabled_var.get())

        if not running:
            state_label.config(text="○ 未运行")
        else:
            state_label.config(text="⏸ 已暂停显示" if paused_var.get() else "● 运行中")

        try:
            btn_pause.state(["!disabled"] if running else ["disabled"])
        except Exception:
            pass

    def _set_pause_state(is_paused: bool):
        nonlocal pause_banner_shown
        is_paused = bool(is_paused)

        # 状态没变化，直接返回（防刷）
        if paused_var.get() == is_paused:
            return

        paused_var.set(is_paused)

        if is_paused:
            btn_pause.config(text="▶ 继续")

            # 暂停时锁定旁路开关
            try:
                chk.state(["!disabled"])
            except Exception:
                pass

            # 只插入一次“已暂停显示”
            if not pause_banner_shown:
                pause_banner_shown = True
                try:
                    serial_debug_text.config(state="normal")
                    serial_debug_text.insert(
                        "end",
                        "\n—— 已暂停显示（串口仍在采集，旁路开关已锁定）——\n"
                    )
                    serial_debug_text.see("end")
                    serial_debug_text.config(state="disabled")
                except Exception:
                    pass

        else:
            btn_pause.config(text="⏸ 暂停")

            # 继续时仍保持旁路开关锁定
            try:
                chk.state(["disabled"])
            except Exception:
                pass

            # 插入“已继续显示”
            try:
                serial_debug_text.config(state="normal")
                serial_debug_text.insert(
                    "end",
                    "\n—— 已继续显示（串口仍在采集，旁路开关已锁定）——\n"
                )
                serial_debug_text.see("end")
                serial_debug_text.config(state="disabled")
            except Exception:
                pass

            pause_banner_shown = False

        # 统一在这里刷新状态标签
        _update_state_label()

    def _toggle_pause():
        _set_pause_state(not paused_var.get())

    # 串口调试区底部状态栏（左下角）
    serial_status_bar = ttk.Frame(serial_debug_win)
    serial_status_bar.pack(side="bottom", fill="x", padx=8, pady=(0, 6))

    state_label = ttk.Label(serial_status_bar, text="")
    state_label.pack(side="left")

    btn_pause.config(command=_toggle_pause)

    drop_label = ttk.Label(top, text="")
    drop_label.pack(side="right")

    body = ttk.Frame(serial_debug_win)
    body.pack(fill="both", expand=True, padx=8, pady=6)

    yscroll = ttk.Scrollbar(body, orient="vertical")
    yscroll.pack(side="right", fill="y")

    serial_debug_text = tk.Text(body, wrap="none", yscrollcommand=yscroll.set)
    serial_debug_text.pack(side="left", fill="both", expand=True)
    yscroll.config(command=serial_debug_text.yview)

    serial_debug_text.config(state="disabled")
    _update_state_label()
    serial_debug_text.tag_config("find_hit", background="yellow")

    find_win = None
    find_var = tk.StringVar(value="")
    last_find_index = "1.0"

    def _clear_find_highlight():
        try:
            serial_debug_text.tag_remove("find_hit", "1.0", "end")
        except Exception:
            pass

    def _find_all(term: str):
        _clear_find_highlight()
        if not term:
            return
        start = "1.0"
        while True:
            pos = serial_debug_text.search(term, start, stopindex="end", nocase=True)
            if not pos:
                break
            endpos = f"{pos}+{len(term)}c"
            serial_debug_text.tag_add("find_hit", pos, endpos)
            start = endpos

    def _find_next(_event=None):
        nonlocal last_find_index
        term = find_var.get().strip()
        if not term:
            return "break"

        pos = serial_debug_text.search(term, last_find_index, stopindex="end", nocase=True)
        if not pos:
            pos = serial_debug_text.search(term, "1.0", stopindex="end", nocase=True)
            if not pos:
                return "break"

        endpos = f"{pos}+{len(term)}c"
        serial_debug_text.see(pos)
        serial_debug_text.mark_set("insert", endpos)
        serial_debug_text.tag_remove("sel", "1.0", "end")
        serial_debug_text.tag_add("sel", pos, endpos)
        last_find_index = endpos
        return "break"

    def _find_prev(_event=None):
        nonlocal last_find_index
        term = find_var.get().strip()
        if not term:
            return "break"

        cur = serial_debug_text.index("insert")
        start = "1.0"
        last = None
        while True:
            pos = serial_debug_text.search(term, start, stopindex=cur, nocase=True)
            if not pos:
                break
            last = pos
            start = f"{pos}+1c"

        if last is None:
            cur = "end"
            start = "1.0"
            while True:
                pos = serial_debug_text.search(term, start, stopindex=cur, nocase=True)
                if not pos:
                    break
                last = pos
                start = f"{pos}+1c"
            if last is None:
                return "break"

        endpos = f"{last}+{len(term)}c"
        serial_debug_text.see(last)
        serial_debug_text.mark_set("insert", endpos)
        serial_debug_text.tag_remove("sel", "1.0", "end")
        serial_debug_text.tag_add("sel", last, endpos)
        last_find_index = endpos
        return "break"

    def _open_find():
        nonlocal find_win, last_find_index
        if find_win is not None and find_win.winfo_exists():
            find_win.deiconify()
            find_win.lift()
            return

        find_win = tk.Toplevel(serial_debug_win)
        find_win.title("查找 (Ctrl+F)")
        find_win.resizable(False, False)

        frm = ttk.Frame(find_win, padding=10)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="查找：").grid(row=0, column=0, sticky="w")
        ent = ttk.Entry(frm, textvariable=find_var, width=28)
        ent.grid(row=0, column=1, padx=(6, 6))
        ttk.Button(frm, text="上一个", command=_find_prev).grid(row=0, column=2, padx=(0, 6))
        ttk.Button(frm, text="下一个", command=_find_next).grid(row=0, column=3)

        def _on_change(*_):
            nonlocal last_find_index
            last_find_index = "1.0"
            _find_all(find_var.get().strip())

        find_var.trace_add("write", _on_change)

        ent.focus_set()
        ent.bind("<Return>", _find_next)
        ent.bind("<Shift-Return>", _find_prev)
        find_win.protocol("WM_DELETE_WINDOW", lambda: (find_win.destroy(), _clear_find_highlight()))

    # Ctrl+F 打开查找窗口
    serial_debug_win.bind("<Control-f>", lambda _e: (_open_find(), "break"))
    serial_debug_win.bind("<Control-F>", lambda _e: (_open_find(), "break"))
    # 控制最大行数，避免跑久了内存爆
    MAX_LINES = 5000

    def _append_lines():
        global serial_debug_drop_count

        if serial_debug_text is None or not serial_debug_text.winfo_exists():
            return

        # 暂停时：不取队列、不插入、不滚动；但仍刷新丢弃计数
        if paused_var.get():
            if serial_debug_drop_count > 0:
                drop_label.config(text=f"队列满丢弃：{serial_debug_drop_count} 行")
            else:
                drop_label.config(text="")
            serial_debug_win.after(100, _append_lines)
            return

        lines = []
        # 一次最多拿 200 行，避免 UI 卡顿
        for _ in range(200):
            try:
                lines.append(serial_debug_queue.get_nowait())
            except queue.Empty:
                break

        if lines:
            kw = filter_var.get().strip()

            for ln in lines:
                all_debug_lines.append(ln)

            if len(all_debug_lines) > MAX_STORE_LINES:
                # 保留最后 MAX_STORE_LINES 行
                all_debug_lines[:] = all_debug_lines[-MAX_STORE_LINES:]

            serial_debug_text.config(state="normal")
            for ln in lines:
                if kw and (kw not in ln):
                    continue
                # 确保有换行
                if not ln.endswith("\n"):
                    ln += "\n"
                serial_debug_text.insert("end", ln)
            # 行数裁剪
            try:
                cur_lines = int(serial_debug_text.index("end-1c").split(".")[0])
                if cur_lines > MAX_LINES:
                    # 删除前面多余的行
                    del_lines = cur_lines - MAX_LINES
                    serial_debug_text.delete("1.0", f"{del_lines + 1}.0")
            except Exception:
                pass

            serial_debug_text.see("end")
            serial_debug_text.config(state="disabled")

        if serial_debug_drop_count > 0:
            drop_label.config(text=f"队列满丢弃：{serial_debug_drop_count} 行")
        else:
            drop_label.config(text="")

        # 100ms 刷新一次
        serial_debug_win.after(100, _append_lines)

    def _on_close():
        global SERIAL_DEBUG_ENABLED, serial_debug_drop_count, serial_debug_win, serial_debug_text
        nonlocal pause_banner_shown

        # 1) 关闭旁路输出（全局开关）
        SERIAL_DEBUG_ENABLED = False

        # 2) 复位 UI 状态：旁路勾选 + 暂停状态
        try:
            enabled_var.set(False)
        except Exception:
            pass

        try:
            paused_var.set(False)
        except Exception:
            pass

        try:
            btn_pause.config(text="⏸ 暂停")
        except Exception:
            pass

        pause_banner_shown = False

        # 3) 清空队列
        try:
            while True:
                serial_debug_queue.get_nowait()
        except queue.Empty:
            pass

        # 4) 清空缓存 & 文本框
        try:
            all_debug_lines.clear()
        except Exception:
            pass

        try:
            if serial_debug_text is not None and serial_debug_text.winfo_exists():
                serial_debug_text.config(state="normal")
                serial_debug_text.delete("1.0", "end")
                serial_debug_text.config(state="disabled")
        except Exception:
            pass

        # 5) 清零丢弃计数 & 顶部提示
        serial_debug_drop_count = 0
        try:
            drop_label.config(text="")
        except Exception:
            pass

        # 6) 最后关闭窗口并清引用
        try:
            if serial_debug_win is not None and serial_debug_win.winfo_exists():
                serial_debug_win.destroy()
        finally:
            serial_debug_win = None
            serial_debug_text = None

    serial_debug_win.protocol("WM_DELETE_WINDOW", _on_close)
    serial_debug_win.bind("<Escape>", lambda _e: _on_close())

    # 相对主窗口居中
    serial_debug_win.update_idletasks()
    try:
        center_window(serial_debug_win, root)
    except Exception:
        pass

    serial_debug_win.deiconify()         # 居中后再显示
    serial_debug_win.lift()
    serial_debug_win.focus_force()

    _append_lines()

def open_voice_text_dialog():
    win = tk.Toplevel(root)
    win.withdraw()
    win.title("语音播报自定义")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    tk.Label(win, text="播报内容：").grid(row=0, column=0, sticky="w", padx=10, pady=(10, 4))

    text = tk.Text(win, width=42, height=4, font=("微软雅黑", 10))
    text.grid(row=1, column=0, columnspan=3, padx=10, pady=(0, 10))
    text.insert("1.0", VOICE_TEXT)

    def do_preview():
        # 预览用当前输入内容，不一定保存
        tmp = text.get("1.0", "end").strip()
        if not tmp:
            messagebox.showerror("错误", "播报内容不能为空")
            return
        # 临时替换生成试听
        global VOICE_TEXT
        old = VOICE_TEXT
        VOICE_TEXT = tmp
        try:
            generate_alert_voice(force=True)
            play_alert()
        finally:
            VOICE_TEXT = old  # 不保存时恢复

    def do_save():
        tmp = text.get("1.0", "end").strip()
        if not tmp:
            messagebox.showerror("错误", "播报内容不能为空")
            return

        global VOICE_TEXT
        VOICE_TEXT = tmp
        save_voice_text_setting()
        generate_alert_voice(force=True)

        msg = "🔊 已更新语音播报内容：" + tmp
        system_ui(msg, "normal")

        win.destroy()

    tk.Button(win, text="试听", width=10, command=do_preview).grid(row=2, column=0, padx=10, pady=(0, 10), sticky="w")
    tk.Button(win, text="保存", width=10, command=do_save).grid(row=2, column=1, pady=(0, 10))
    tk.Button(win, text="取消", width=10, command=win.destroy).grid(row=2, column=2, padx=10, pady=(0, 10), sticky="e")

    win.update_idletasks()
    center_window(win, root)
    win.deiconify()
    win.lift()
    win.focus_force()
    text.focus_set()
    win.bind("<Escape>", lambda _e: win.destroy())

# ================= 保存快捷方式名称 =================
def save_desktop_shortcut_name(name: str):
    if not config.has_section("ui"):
        config["ui"] = {}
    config.set("ui", "desktop_shortcut_name", name)
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        config.write(f)

# ================= 单实例：二次启动时唤醒已有实例 =================
SINGLE_INSTANCE_HOST = "127.0.0.1"

# 端口文件：记录“主实例当前使用的端口”，让二次启动能找到它
PORT_FILE = os.path.join(tempfile.gettempdir(), "sms_single_instance_port.txt")

# 端口尝试范围（足够小，不会乱；足够大，基本不冲突）
PORT_RANGE = range(45678, 45699)

def _read_saved_port():
    try:
        with open(PORT_FILE, "r", encoding="utf-8") as f:
            p = int(f.read().strip())
            return p
    except Exception:
        return None

def _save_port(port: int):
    try:
        with open(PORT_FILE, "w", encoding="utf-8") as f:
            f.write(str(port))
    except Exception:
        pass

def _pick_free_port():
    """从范围里挑一个能 bind 的端口"""
    for p in PORT_RANGE:
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        try:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            s.bind((SINGLE_INSTANCE_HOST, p))
            s.close()
            return p
        except OSError:
            try:
                s.close()
            except Exception:
                pass
            continue
    return None

def _try_notify_existing_instance() -> bool:
    """如果已有实例在监听，则发送 SHOW 并返回 True；否则返回 False"""
    port = _read_saved_port()
    if not port:
        return False

    try:
        with socket.create_connection((SINGLE_INSTANCE_HOST, port), timeout=0.3) as s:
            s.sendall(b"SHOW")
        return True

    except OSError:
        # 连接失败：大概率是旧的 port 文件残留，清理一下
        try:
            os.remove(PORT_FILE)
        except Exception:
            pass
        return False

def _start_single_instance_server(port: int, show_callback):
    """
    本实例成为“主实例”：在后台监听端口。
    收到 SHOW 就调用 show_callback()（用 root.after 调回主线程）。
    """
    _save_port(port)

    def _server():
        srv = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        srv.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)

        try:
            srv.bind((SINGLE_INSTANCE_HOST, port))
        except OSError:
            # 极小概率：端口在“检测后-真正 bind 前”被抢占
            try:
                srv.close()
            except Exception:
                pass
            return   # 👈 直接放弃单实例监听，但程序本身继续运行

        srv.listen(5)

        while True:
            try:
                conn, _addr = srv.accept()
                with conn:
                    data = conn.recv(1024) or b""
                    if b"SHOW" in data:
                        try:
                            show_callback()
                        except Exception:
                            pass
            except Exception:
                time.sleep(0.2)

    threading.Thread(target=_server, daemon=True).start()

def center_on_screen(win, w=None, h=None):
    """将窗口居中到屏幕（主窗口建议传入 w/h，避免 withdraw 状态取到 minsize）。"""
    win.update_idletasks()

    # withdraw 状态下 winfo_width/height 可能等于 minsize；优先用传入值，其次用 reqwidth/reqheight
    if w is None or h is None:
        w = win.winfo_reqwidth()
        h = win.winfo_reqheight()

    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    x = (sw - w) // 2
    y = (sh - h) // 2
    win.geometry(f"{w}x{h}+{x}+{y}")
    
# 如果检测到已有实例：通知它显示窗口，然后直接退出当前进程（避免多开）
if not ALLOW_MULTI_INSTANCE:
    if _try_notify_existing_instance():
        sys.exit(0)

# ================= GUI =================
root = tk.Tk()
root.withdraw()
root.minsize(500, 200)

threading.Thread(target=generate_alert_voice, daemon=True).start()

def resource_path(relative):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, relative)
    # 脚本模式：用文件本身所在目录
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative)

try:
    root.iconbitmap(resource_path("icon.ico"))
except Exception as e:
    print("icon.ico 加载失败：", e)

# 更改弹窗左上角图标：让所有弹窗继承 icon.ico
try:
    _ICON_ICO_PATH = resource_path("icon.ico")

    def _apply_window_icon(_win):
        try:
            if _ICON_ICO_PATH and os.path.exists(_ICON_ICO_PATH):
                _win.iconbitmap(_ICON_ICO_PATH)
        except Exception:
            # 仅图标失败，不影响弹窗功能
            pass

    # 1) 所有 tk.Toplevel 弹窗：创建后自动设置图标
    _orig_Toplevel = tk.Toplevel

    def _patched_Toplevel(*args, **kwargs):
        _win = _orig_Toplevel(*args, **kwargs)
        try:
            _win.after(0, lambda w=_win: _apply_window_icon(w))
        except Exception:
            _apply_window_icon(_win)
        return _win

    tk.Toplevel = _patched_Toplevel

    # 2) messagebox 弹窗：补 parent=root 继承图标
    _mb_showinfo = messagebox.showinfo
    _mb_showwarning = messagebox.showwarning
    _mb_showerror = messagebox.showerror
    _mb_askyesno = messagebox.askyesno

    def _mb_wrap(fn):
        def _inner(title, message, **options):
            if "parent" not in options:
                options["parent"] = root
            return fn(title, message, **options)
        return _inner

    messagebox.showinfo = _mb_wrap(_mb_showinfo)
    messagebox.showwarning = _mb_wrap(_mb_showwarning)
    messagebox.showerror = _mb_wrap(_mb_showerror)
    messagebox.askyesno = _mb_wrap(_mb_askyesno)

except Exception as _e:
    # 任何异常都不能影响主程序和弹窗正常使用
    print("弹窗图标补丁加载失败：", _e)

root.title("短信监听系统")
root.geometry("760x520")

root.update_idletasks()
if not START_MINIMIZED:
    center_on_screen(root, 760, 520)
    root.deiconify()
else:
    # 自启：保持隐藏，托盘可“显示”
    root.withdraw()

# ================= 托盘 / 退出 / 隐藏 =================
tray_icon = None
is_exiting = False

def show_window():
    root.after(0, lambda: (root.deiconify(), root.lift(), root.focus_force()))

if not ALLOW_MULTI_INSTANCE:
    port = _pick_free_port()
    if port is None:
        # 极端情况：范围内全占用，就不做单实例（至少不崩）
        msg = "⚠️ 单实例端口被占用，已降级为允许多开"
        log_early(msg, tag="normal")

    else:
        _start_single_instance_server(port, lambda: root.after(0, show_window))

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

    try:
        if os.path.exists(PORT_FILE):
            os.remove(PORT_FILE)
    except Exception:
        pass

def on_close():
    """点右上角×：隐藏到托盘，不退出"""
    hide_window()

root.protocol("WM_DELETE_WINDOW", on_close)
root.bind("<Escape>", lambda _e: on_close())

def create_tray():
    global tray_icon
    def _load_tray_image():
        # 1) 优先使用 icon.ico
        try:
            return Image.open(resource_path("icon.ico"))
        except Exception:
            pass

        # 2) 兜底：生成一个简单的 16x16 图标
        try:
            img = Image.new("RGB", (16, 16), color=(200, 30, 30))  # 深红色
            return img
        except Exception:
            return None

    img = _load_tray_image()
    if img is None:
        # 极端兜底：理论上几乎不会到这一步
        return

    menu = pystray.Menu(
        pystray.MenuItem("显示", lambda: show_window(), default=True),  # 双击托盘
        pystray.MenuItem("隐藏", lambda: hide_window()),
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("退出", lambda: cleanup_and_exit()),
    )

    tray_icon = pystray.Icon("sms_tray", img, "短信监听系统", menu)
    tray_icon.run()

threading.Thread(target=create_tray, daemon=True).start()

def center_window(win, parent):
    win.update_idletasks()

    w = win.winfo_width()
    h = win.winfo_height()
    if w <= 1 or h <= 1:
        w = win.winfo_reqwidth()
        h = win.winfo_reqheight()

    px, py = parent.winfo_rootx(), parent.winfo_rooty()
    pw, ph = parent.winfo_width(), parent.winfo_height()

    x = px + (pw - w) // 2
    y = py + (ph - h) // 2
    win.geometry(f"+{x}+{y}")

def show_about():
    """在主窗口正中显示“关于”弹窗（模态）。"""
    win = tk.Toplevel(root)
    win.withdraw()
    win.title("关于")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    frame = tk.Frame(win, padx=20, pady=15)
    frame.pack(fill=tk.BOTH, expand=True)

    # 版本信息
    tk.Label(frame, text="短信监听系统", font=("微软雅黑", 12, "bold")).pack(pady=(0, 8))
    tk.Label(
        frame,
        text=f"版本：v{APP_VERSION}",
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

    win.update_idletasks()
    center_window(win, root)
    win.deiconify()
    win.lift()
    win.focus_force()
    win.bind("<Escape>", lambda _e: win.destroy())

# ===== 用 grid 布局：内容区永远不会盖住状态栏 =====
root.grid_rowconfigure(0, weight=1)   # 内容区可伸缩
root.grid_rowconfigure(1, weight=0)   # 状态栏固定
root.grid_columnconfigure(0, weight=1)

# 中间内容区域
main_frame = tk.Frame(root)
main_frame.grid(row=0, column=0, sticky="nsew")

text_area = ScrolledText(main_frame, font=("微软雅黑", 10))
text_area.pack(fill=tk.BOTH, expand=True)  # 这里用 pack 没问题，因为只在 main_frame 内部

# 把早期提示补到窗口
for m, t in PENDING_UI_LOGS:
    try:
        text_area.insert(tk.END, m + "\n", t)
    except Exception:
        pass
PENDING_UI_LOGS.clear()

# 底部状态栏
status_frame = tk.Frame(root)
status_frame.grid(row=1, column=0, sticky="ew")

status_var = tk.StringVar(value="🔍 启动中…")
status_label = tk.Label(status_frame, textvariable=status_var, anchor="w")
status_label.pack(side=tk.LEFT, padx=6)

def set_status(text, color="black"):
    root.after(0, lambda: (status_var.set(text), status_label.config(fg=color)))

text_area.tag_config("normal", foreground="black", font=("微软雅黑", 10))

def apply_sms_font_style():
    try:
        text_area.tag_config("sms", foreground=SMS_FONT_COLOR, font=("微软雅黑", SMS_FONT_SIZE))
    except Exception:
        pass

apply_sms_font_style()

def log(msg, tag="normal"):
    """线程安全：在子线程调用时自动切回主线程执行"""
    def _do():
        # 1) UI
        try:
            text_area.insert(tk.END, msg + "\n", tag)
            text_area.see(tk.END)
        except Exception:
            pass

        # 2) 文件（COM 分日志，走 get_log_file -> sms_{LOG_PREFIX}_*.txt）
        try:
            with open(get_log_file(), "a", encoding="utf-8") as f:
                f.write(f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n")
        except Exception:
            pass

    # --- root / text_area 是否可用 ---
    root_ok = False
    try:
        root_ok = ("root" in globals()) and (root is not None) and root.winfo_exists()
    except Exception:
        root_ok = False

    text_ok = False
    try:
        text_ok = ("text_area" in globals()) and (text_area is not None) and text_area.winfo_exists()
    except Exception:
        text_ok = False

    # UI 不可用：至少别丢（写 system + 缓存）
    if not (root_ok and text_ok):
        try:
            log_early(msg, tag)
        except Exception:
            pass
        return

    # --- 主线程直接做；子线程丢回主线程 ---
    try:
        if threading.current_thread() is threading.main_thread():
            _do()
        else:
            root.after(0, _do)
    except Exception:
        # after 不可用/竞态兜底
        try:
            log_early(msg, tag)
        except Exception:
            pass

# ================= 声音 =================
def play_alert():
    if not VOICE_ENABLED:
        return

    try:
        if os.path.exists(TTS_FILE):
            winsound.PlaySound(TTS_FILE,winsound.SND_FILENAME | winsound.SND_ASYNC)
        else:
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except Exception:
        winsound.MessageBeep(winsound.MB_ICONASTERISK)

def show_sms_popup(msg: str):
    """弹窗确认后，自动显示主程序窗口"""
    global VOICE_ENABLED
    if not VOICE_ENABLED:
        return

    def _popup_and_show():
        messagebox.showinfo("短信提醒", msg)  # 用户点“确定”前会阻塞
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

# ================= 日志清理 =================
def _parse_date_from_log_filename(filename: str):
    """
    从文件名中解析日期：支持 sms_system_YYYY-MM-DD.txt / sms_COM5_YYYY-MM-DD.txt / sms_xxx_YYYY-MM-DD.txt
    解析失败返回 None
    """
    m = re.search(r"_(\d{4}-\d{2}-\d{2})\.txt$", filename)
    if not m:
        return None
    try:
        return datetime.strptime(m.group(1), "%Y-%m-%d").date()
    except Exception:
        return None

def cleanup_old_logs(days: int) -> int:
    """
    删除 LOG_DIR 中超过 days 天的 .txt 日志，返回删除数量
    规则：根据文件名末尾的 YYYY-MM-DD 判断；解析失败则用文件修改时间判断
    """
    if days < 0:
        days = 0

    cutoff = (datetime.now() - timedelta(days=days)).date()
    deleted = 0

    if not os.path.isdir(LOG_DIR):
        return 0

    for name in os.listdir(LOG_DIR):
        path = os.path.join(LOG_DIR, name)
        if not os.path.isfile(path):
            continue
        if not name.lower().endswith(".txt"):
            continue
        if not name.lower().startswith("sms_"):
            continue

        file_date = _parse_date_from_log_filename(name)
        try:
            if file_date is None:
                # fallback：用文件修改时间
                mtime = datetime.fromtimestamp(os.path.getmtime(path)).date()
                file_date = mtime

            # 早于 cutoff 才删（例如保留 7 天：删 7 天之前的）
            if file_date < cutoff:
                os.remove(path)
                deleted += 1
        except Exception:
            # 单个文件删失败不影响整体
            pass

    return deleted

def open_log_cleanup_dialog():
    """弹窗：设置保留天数并清理日志"""
    win = tk.Toplevel(root)
    win.withdraw()
    win.title("日志自动清理")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    frame = tk.Frame(win, padx=14, pady=12)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="保留最近 N 天日志：", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w")

    days_var = tk.StringVar(value=str(LOG_RETENTION_DAYS))
    days_entry = tk.Entry(frame, textvariable=days_var, width=10)
    days_entry.grid(row=0, column=1, sticky="w", padx=(8, 0))
    tk.Label(frame, text="天", font=("微软雅黑", 10)).grid(row=0, column=2, sticky="w", padx=(6, 0))

    tip = tk.Label(
        frame,
        text="说明：会删除 sms_logs 目录下超过 N 天的 sms_*.txt 日志（含 sms_system / sms_COMx）。",
        fg="gray",
        font=("微软雅黑", 9),
        wraplength=360,
        justify="left",
    )
    tip.grid(row=1, column=0, columnspan=3, sticky="w", pady=(10, 6))

    def do_cleanup():
        global LOG_RETENTION_DAYS, AUTO_LOG_CLEANUP

        try:
            days = int(days_var.get().strip())
            if days < 0:
                raise ValueError
        except Exception:
            messagebox.showerror("错误", "天数必须是非负整数（例如 30）")
            return

        # 确认开启自动清理
        if not messagebox.askyesno("确认", f"确定设置为自动清理，并保留最近 {days} 天日志吗？"):
            return

        LOG_RETENTION_DAYS = days
        AUTO_LOG_CLEANUP = True

        # 保存到 config.ini，重启后仍然生效
        try:
            if not config.has_section("ui"):
                config["ui"] = {}
            config.set("ui", "auto_log_cleanup", "1")
            config.set("ui", "log_retention_days", str(LOG_RETENTION_DAYS))
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                config.write(f)
        except Exception:
            pass

        # 记录到 system，并在窗口显示
        msg = f"✅ 已启用自动日志清理：保留 {LOG_RETENTION_DAYS} 天（每 {AUTO_CLEANUP_INTERVAL_HOURS} 小时执行一次）"
        system_ui(msg, "normal")

        # 启动/重启自动定时器（以后每24小时自动清理）
        schedule_auto_log_cleanup(restart=True, first_delay_sec=60)

        messagebox.showinfo("完成", "已启用自动日志清理（程序运行期间会定期清理）。")
        win.destroy()

    btns = tk.Frame(frame)
    btns.grid(row=2, column=0, columnspan=3, sticky="e", pady=(10, 0))

    tk.Button(btns, text="确认", width=10, command=do_cleanup).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btns, text="取消", width=10, command=win.destroy).pack(side=tk.LEFT)

    win.update_idletasks()
    center_window(win, root)
    win.deiconify()
    win.lift()
    win.focus_force()
    days_entry.focus_set()
    win.bind("<Return>", lambda _e: do_cleanup())
    win.bind("<Escape>", lambda _e: win.destroy())

def open_update_proxy_dialog():
    """弹窗：编辑 GitHub Proxy 下载前缀与 API 前缀"""
    if not config.has_section("update"):
        config["update"] = {
            "api_proxy_base": "https://github-api.daybyday.top/",
            "proxy_base": "https://gh-proxy.com/",
        }

    win = tk.Toplevel(root)
    win.withdraw()
    win.title("检查更新代理设置")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    frame = tk.Frame(win, padx=14, pady=12)
    frame.pack(fill=tk.BOTH, expand=True)

    proxy_var = tk.StringVar(value=config.get("update", "proxy_base", fallback=""))
    api_var = tk.StringVar(value=config.get("update", "api_proxy_base", fallback=""))

    tk.Label(frame, text="API 代理前缀 api_proxy_base：").grid(row=0, column=0, sticky="w")
    api_entry = tk.Entry(frame, textvariable=api_var, width=44)
    api_entry.grid(row=1, column=0, pady=(4, 10), sticky="w")

    tk.Label(frame, text="下载代理前缀 proxy_base：").grid(row=2, column=0, sticky="w")
    proxy_entry = tk.Entry(frame, textvariable=proxy_var, width=44)
    proxy_entry.grid(row=3, column=0, pady=(4, 10), sticky="w")

    def _normalize(s: str) -> str:
        s = (s or "").strip()
        if not s:
            return ""
        # 自动补协议
        if not (s.startswith("http://") or s.startswith("https://")):
            s = "https://" + s
        # 补 /
        if not s.endswith("/"):
            s += "/"
        return s

    def save():
        config.set("update", "proxy_base", _normalize(proxy_var.get()))
        config.set("update", "api_proxy_base", _normalize(api_var.get()))
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)
        messagebox.showinfo("完成", "代理设置已保存")

    def test_connection():
        # 先禁用按钮，避免重复点（需要 btn_test 变量，下面按钮处我也给你改法）
        try:
            btn_test.config(state="disabled", text="测试中…")
        except Exception:
            pass

        api_raw = api_var.get().strip()

        def classify_err(e: Exception) -> str:
            s = str(e)
            if "SSLV3_ALERT_HANDSHAKE_FAILURE" in s or "sslv3 alert handshake failure" in s:
                return "TLS握手失败（代理节点/线路不兼容或被干扰）"
            if "timed out" in s.lower():
                return "连接超时（线路慢/被阻断）"
            if "name or service not known" in s.lower() or "getaddrinfo failed" in s.lower():
                return "DNS 解析失败"
            return s

        def worker():
            owner, repo = GITHUB_OWNER, GITHUB_REPO
            api_path = f"/repos/{owner}/{repo}/releases/latest"

            checks = []
            ok_api = False
            release_json = None

            # 1) 先测 api_proxy_base（支持 |）
            bases = [x.strip() for x in api_raw.split("|") if x.strip()] if api_raw else []
            ok_bases = []

            for base in bases:
                base_n = _normalize(base)
                url = base_n.rstrip("/") + api_path
                try:
                    release_json = _http_get_json(url, timeout=6, retries=2)
                    checks.append((base_n, True, "OK"))
                    ok_bases.append(base_n)
                    ok_api = True
                    break
                except Exception as e:
                    checks.append((base_n, False, classify_err(e)))

            # 2) API 代理都失败 -> 再测直连兜底
            if not ok_api:
                direct_url = "https://api.github.com" + api_path
                try:
                    release_json = _http_get_json(direct_url, timeout=6, retries=2)
                    checks.append(("直连 api.github.com", True, "OK"))
                    ok_api = True
                except Exception as e:
                    checks.append(("直连 api.github.com", False, classify_err(e)))

            # 3) 只有 API 成功，才测试下载代理 proxy_base
            if ok_api and release_json:
                asset = _pick_exe_asset(release_json)
                if asset:
                    raw_url = asset.get("browser_download_url") or ""
                    pb = _normalize(proxy_var.get())
                    if pb and raw_url.startswith("http"):
                        test_url = pb + raw_url
                        try:
                            _http_probe(test_url, timeout=6, retries=2)
                            checks.append((f"下载代理 {pb}", True, "OK"))
                        except Exception as e:
                            checks.append((f"下载代理 {pb}", False, classify_err(e)))
                    else:
                        checks.append(("下载代理 proxy_base", False, "未填写或获取不到下载链接"))
                else:
                    checks.append(("下载代理 proxy_base", False, "Release 无 .zip 附件"))

            def done():
                try:
                    btn_test.config(state="normal", text="测试连接")
                except Exception:
                    pass

                lines = []
                for name, ok, info in checks:
                    lines.append(("✅ " if ok else "❌ ") + f"{name}：{info}")

                # 是否下载代理 OK（你 checks 里成功时 name 是 "下载代理 {pb}"）
                download_ok = any((ok is True) and isinstance(name, str) and name.startswith("下载代理 ")
                                  for (name, ok, _info) in checks)

                if ok_bases or download_ok:
                    lines.append("")

                if ok_bases:
                    lines.append("提示：API 代理可用，检测更新将优先使用它。")

                if download_ok:
                    lines.append("提示：下载代理可用，下载链接将优先使用它。")

                messagebox.showinfo("测试结果", "\n".join(lines))

            root.after(0, done)

        threading.Thread(target=worker, daemon=True).start()

    def reset_default():
        api_var.set("https://github-api.daybyday.top/")
        proxy_var.set("https://gh-proxy.com/")

    btns = tk.Frame(frame)
    btns.grid(row=4, column=0, sticky="e", pady=(6, 0))
    btn_test = tk.Button(btns, text="测试连接", width=10, command=test_connection)
    btn_test.pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btns, text="恢复默认", width=10, command=reset_default).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btns, text="保存", width=10, command=save).pack(side=tk.LEFT, padx=(0, 8))
    tk.Button(btns, text="取消", width=10, command=win.destroy).pack(side=tk.LEFT)

    win.update_idletasks()
    center_window(win, root)
    win.deiconify()
    win.lift()
    win.focus_force()
    api_entry.focus_set()
    win.bind("<Return>", lambda _e: save())
    win.bind("<Escape>", lambda _e: win.destroy())

def _auto_log_cleanup_tick():
    """一次自动清理 + 重新安排下一次"""
    global AUTO_CLEANUP_AFTER_ID

    if not AUTO_LOG_CLEANUP:
        AUTO_CLEANUP_AFTER_ID = None
        return

    days = LOG_RETENTION_DAYS if LOG_RETENTION_DAYS >= 0 else 0

    try:
        n = cleanup_old_logs(days)
        msg = f"🧹 自动日志清理：已删除 {n} 个旧日志文件（保留 {days} 天）"
        # 显示到窗口 + 写入 system 日志
        system_ui(msg, "normal")

    except Exception as e:
        system_ui(f"⚠️ 自动日志清理失败：{e}")

    # 下一次
    AUTO_CLEANUP_AFTER_ID = root.after(AUTO_CLEANUP_INTERVAL_HOURS * 3600 * 1000, _auto_log_cleanup_tick)

def schedule_auto_log_cleanup(restart: bool = True, first_delay_sec: int = 60):
    """
    开启/重启自动清理定时器
    - restart=True：会先取消旧定时器，避免重复跑
    - first_delay_sec：首次执行延迟（避免刚启动就占资源）
    """
    global AUTO_CLEANUP_AFTER_ID

    if restart and AUTO_CLEANUP_AFTER_ID is not None:
        try:
            root.after_cancel(AUTO_CLEANUP_AFTER_ID)
        except Exception:
            pass
        AUTO_CLEANUP_AFTER_ID = None

    if not AUTO_LOG_CLEANUP:
        return
    AUTO_CLEANUP_AFTER_ID = root.after(first_delay_sec * 1000, _auto_log_cleanup_tick)

# ================= 检测更新 =================
def _ver_tuple(v: str):
    # 允许 "v1.1.1" / "1.1.1"
    v = (v or "").strip().lstrip("vV")
    parts = []
    for x in v.split("."):
        try:
            parts.append(int(x))
        except Exception:
            parts.append(0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])

def _http_get_json(url: str, timeout=8, retries=3):
    last_err = None

    ctx = ssl.create_default_context()

    # 禁用系统代理（避免 v2rayng/system proxy 影响 urllib TLS）
    opener = urllib.request.build_opener(
        urllib.request.ProxyHandler({}),          # 关键：空代理=不走系统代理
        urllib.request.HTTPSHandler(context=ctx)  # 保持 TLS 上下文
    )

    for i in range(retries):
        try:
            req = urllib.request.Request(
                url,
                headers={"User-Agent": "sms-updater", "Accept": "application/vnd.github+json"},
                method="GET",
            )
            with opener.open(req, timeout=timeout) as resp:
                data = resp.read().decode("utf-8", "ignore")
                return json.loads(data)

        except Exception as e:
            last_err = e
            try:
                time.sleep(0.6 * (2 ** i))
            except Exception:
                pass

    raise last_err

def _http_probe(url: str, timeout=8, retries=2):
    """
    探测 URL 是否可访问：
    - 优先 HEAD（更快，不下载正文）
    - 部分代理不支持 HEAD，则 fallback GET 读取少量字节
    - 禁用系统代理（与 _http_get_json 一致）
    """
    last_err = None
    ctx = ssl.create_default_context()
    opener = urllib.request.build_opener(
        urllib.request.ProxyHandler({}),
        urllib.request.HTTPSHandler(context=ctx)
    )

    for i in range(retries):
        try:
            # 1) HEAD
            req = urllib.request.Request(url, headers={"User-Agent": "sms-updater"}, method="HEAD")
            with opener.open(req, timeout=timeout) as resp:
                # 2xx/3xx 基本都算可达
                return True, f"HTTP {getattr(resp, 'status', 200)}"
        except Exception as e_head:
            try:
                # 2) fallback GET (读取少量字节即可)
                req = urllib.request.Request(url, headers={"User-Agent": "sms-updater"}, method="GET")
                with opener.open(req, timeout=timeout) as resp:
                    resp.read(64)
                    return True, f"HTTP {getattr(resp, 'status', 200)}"
            except Exception as e_get:
                last_err = e_get
                try:
                    time.sleep(0.4 * (2 ** i))
                except Exception:
                    pass

    raise last_err

def _get_update_config():
    proxy_base = config.get("update", "proxy_base", fallback="https://gh-proxy.com/").strip()
    api_proxy_base = config.get("update", "api_proxy_base", fallback="").strip()
    # 规范：确保以 / 结尾
    if proxy_base and not proxy_base.endswith("/"):
        proxy_base += "/"
    if api_proxy_base and not api_proxy_base.endswith("/"):
        api_proxy_base += "/"
    return proxy_base, api_proxy_base

def _get_latest_release():
    owner, repo = GITHUB_OWNER, GITHUB_REPO
    api_path = f"/repos/{owner}/{repo}/releases/latest"
    direct = "https://api.github.com" + api_path
    _proxy_base, api_proxy_base = _get_update_config()

    urls = []

    # 1) 代理优先（支持 | 分隔多个候选）
    if api_proxy_base:
        for base in (x.strip() for x in api_proxy_base.split("|") if x.strip()):
            if not (base.startswith("http://") or base.startswith("https://")):
                base = "https://" + base
            base = base.rstrip("/")
            urls.append(base + api_path + f"?t={int(time.time())}")

    # 2) 最后再直连兜底
    urls.append(direct)

    last_err = None
    for u in urls:
        try:
            return _http_get_json(u, timeout=8, retries=3)
        except Exception as e:
            last_err = e
            continue

    raise RuntimeError(f"获取最新版本信息失败：{last_err}")

def _pick_exe_asset(release_json: dict):
    assets = release_json.get("assets") or []
    zip_assets = [a for a in assets if (a.get("name","").lower().endswith(".zip"))]
    if not zip_assets:
        return None
    zip_assets.sort(key=lambda a: -int(a.get("size", 0) or 0))
    return zip_assets[0]

def check_update_and_prompt():
    def worker():
        try:
            rel = _get_latest_release()
            tag = rel.get("tag_name") or ""
            latest = _ver_tuple(tag)
            current = _ver_tuple(APP_VERSION)

            if latest <= current:
                root.after(0, lambda: messagebox.showinfo("检测更新", f"当前已是最新版本：V{APP_VERSION}"))
                return

            asset = _pick_exe_asset(rel)
            if not asset:
                root.after(0, lambda: messagebox.showwarning(
                    "检测更新",
                    f"发现新版本：{tag}\n但 Release 里没有 .zip 附件。"
                ))
                return

            raw_url = asset.get("browser_download_url") or ""
            proxy_base, _api_proxy_base = _get_update_config()

            # 下载链接：优先走代理（大陆可用），同时给用户一个“原始链接”
            proxy_url = (proxy_base + raw_url) if (proxy_base and raw_url.startswith("http")) else raw_url

            def ask():
                ok = messagebox.askyesno(
                    "发现新版本",
                    f"当前：V{APP_VERSION}\n最新：{tag}\n\n是否打开下载链接？（如已配置下载代理，将优先使用）"
                )
                if ok:
                    try:
                        webbrowser.open(proxy_url)
                    except Exception:
                        pass

            root.after(0, ask)

        except Exception as e:
            root.after(0, lambda: messagebox.showerror("检测更新失败", str(e)))

    threading.Thread(target=worker, daemon=True).start()

# ================= 每日清空 =================
def clear_text_area_for_new_day():
    clear_window()
    system_ui("📅 新的一天，窗口已清空")
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

def _push_serial_debug(raw_line: str):
    global serial_debug_drop_count
    if not SERIAL_DEBUG_ENABLED:
        return

    # 空行/纯空白直接忽略，避免调试窗口大量空白行
    if raw_line is None:
        return
    if not str(raw_line).strip():
        return

    try:
        serial_debug_queue.put_nowait(raw_line)
    except queue.Full:
        serial_debug_drop_count += 1

# ================= 串口线程（自动识别 + 自动重连） =================
def read_serial():
    """
    串口读取线程（严格模式）：
    - 仅当串口行中包含 [I]-[handler_sms.smsCallback] 才认为“短信有效”
    - 命中后会收集同一条短信的多行输出，合并后再进行【关键词过滤】与弹窗/播报（避免弹窗不完整）
    - 关键词过滤规则：full_msg 只要包含 KEYWORDS 任意一项即放行；否则忽略不显示/不弹窗/不播报
    - 其它所有串口日志全部忽略
    """
    global serial_obj, serial_running, PORT, LOG_PREFIX

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
                        port_ui(ln, "normal")   
                        first = False
                    else:
                        port_ui(ln, "sms")      
            else:
                port_ui("📩 收到短信：", "normal")  
                port_ui(full_msg, "sms")          

            play_alert()               
            show_sms_popup(full_msg)   
        else:
            system_ui("🚫 短信未命中关键词，已忽略", "normal")  

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

            LOG_PREFIX = PORT.replace(":", "_")

            system_ui(f"🔌 串口已连接：{PORT} @ {BAUD}")
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
                _push_serial_debug(line)
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
            LOG_PREFIX = "system"
            system_ui(f"⚠️ 串口异常：{e}")
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

        system_ui(f"⚙️ 串口设置已更新：mode={MODE} port={PORT or '(Auto)'} baud={BAUD}")
        win.destroy()

    win = tk.Toplevel(root)
    win.withdraw()
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
    btn_row.grid(row=3, column=0, columnspan=2, pady=(10, 0))
    tk.Button(btn_row, text="刷新", width=10, command=refresh_ports).pack(side=tk.LEFT, padx=8)
    tk.Button(btn_row, text="应用", width=10, command=apply).pack(side=tk.LEFT, padx=8)
    tk.Button(btn_row,text="取消",width=10,command=win.destroy).pack(side=tk.LEFT, padx=8)

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
    win.deiconify()
    win.lift()
    win.focus_force()
    mode_box.focus_set()
    win.bind("<Return>", lambda _e: apply())
    win.bind("<Escape>", lambda _e: win.destroy())

# ================= 弹窗：快捷方式设置窗口 =================
def open_desktop_shortcut_dialog():
    default_name = config.get(
        "ui", "desktop_shortcut_name", fallback="短信监听系统"
    )

    win = tk.Toplevel(root)
    win.withdraw()
    win.title("创建桌面快捷方式")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    frame = tk.Frame(win, padx=14, pady=12)
    frame.pack(fill=tk.BOTH, expand=True)

    tk.Label(frame, text="快捷方式名称：", font=("微软雅黑", 10)).grid(
        row=0, column=0, sticky="w"
    )

    name_var = tk.StringVar(value=default_name)
    entry = tk.Entry(frame, textvariable=name_var, width=28)
    entry.grid(row=1, column=0, pady=(6, 12), sticky="w")

    def apply_now():
        name = name_var.get().strip()
        if not name:
            messagebox.showerror("错误", "名称不能为空")
            return
        try:
            create_desktop_shortcut(name)
            save_desktop_shortcut_name(name)
            msg = f"✅ 桌面快捷方式已创建：{name}.lnk"
            system_ui(msg, "normal")
            messagebox.showinfo("完成", "桌面快捷方式已创建")
        except Exception as e:
            messagebox.showerror("失败", str(e))

    def save_only():
        name = name_var.get().strip()
        if not name:
            messagebox.showerror("错误", "名称不能为空")
            return
        save_desktop_shortcut_name(name)
        # 窗口显示 + system 日志（不写 COM 日志）
        msg = f"💾 已保存桌面快捷方式：{name}"
        system_ui(msg, "normal")
        messagebox.showinfo("已保存", "名称已保存，下次可直接应用")

    btns = tk.Frame(frame)
    btns.grid(row=2, column=0, sticky="e")

    tk.Button(btns, text="应用", width=10, command=apply_now).pack(
        side=tk.LEFT, padx=(0, 8)
    )
    tk.Button(btns, text="保存", width=10, command=save_only).pack(
        side=tk.LEFT, padx=(0, 8)
    )
    tk.Button(btns, text="取消", width=10, command=win.destroy).pack(
        side=tk.LEFT, padx=(0, 8)
    )

    win.update_idletasks()
    center_window(win, root)
    win.deiconify()
    win.lift()
    win.focus_force()
    entry.focus_set()
    win.bind("<Return>", lambda _e: apply_now())
    win.bind("<Escape>", lambda _e: win.destroy())

# ================= 关键词设置窗口（增加/删除/修改 + 居中模态） =================
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
            if not config.has_section("ui"):
                config["ui"] = {}
            config.set("ui", "keywords", "|".join(KEYWORDS))
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
        system_ui(f"🧷 关键词增加：{v}")

    def del_kw():
        global KEYWORDS
        sel = listbox.curselection()
        if not sel:
            messagebox.showwarning("提示", "请选择要删除的关键词")
            return
        idx = sel[0]
        if idx < 0 or idx >= len(KEYWORDS):
            return
        old = KEYWORDS[idx]
        KEYWORDS.pop(idx)
        save_keywords_to_config()
        entry_var.set("")
        refresh_list(select_index=min(idx, len(KEYWORDS) - 1))
        system_ui(f"🧷 关键词删除：{old}")

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
        old = KEYWORDS[idx] 
        KEYWORDS[idx] = v
        save_keywords_to_config()
        refresh_list(select_index=idx)
        system_ui(f"🧷 关键词修改：{old} -> {v}")

    win = tk.Toplevel(root)
    win.withdraw()
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
    win.deiconify()
    win.lift()
    win.focus_force()
    entry.focus_set()

    win.bind("<Return>", lambda _e: edit_kw())
    listbox.bind("<Delete>", lambda _e: del_kw())
    win.bind("<Escape>", lambda _e: win.destroy())

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
    msg = "🔊 语音播报：已开启" if VOICE_ENABLED else "🔇 语音播报：已关闭"

    system_ui(msg, "normal")

def toggle_multi_instance():
    global ALLOW_MULTI_INSTANCE
    ALLOW_MULTI_INSTANCE = multi_instance_var.get()
    try:
        if not config.has_section("ui"):
            config["ui"] = {}
        config.set(
            "ui",
            "allow_multi_instance",
            "1" if ALLOW_MULTI_INSTANCE else "0"
        )
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)
    except Exception:
        pass

    msg = "🧩 程序多开：已开启" if ALLOW_MULTI_INSTANCE else "🔒 程序多开：已关闭"

    system_ui(msg, "normal")

def toggle_autostart():
    set_autostart(autostart_var.get())

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
menu_bar.add_command(label="🔊 语音播报", command=toggle_voice_broadcast)
voice_menu_index = menu_bar.index("end")

# ================= 设置 菜单 =================
settings_menu = tk.Menu(menu_bar, tearoff=0)

autostart_var = tk.BooleanVar(value=is_autostart_enabled())

multi_instance_var = tk.BooleanVar(value=ALLOW_MULTI_INSTANCE)

settings_menu.add_checkbutton(
    label="开机自启",
    variable=autostart_var,
    command=toggle_autostart
)

settings_menu.add_checkbutton(
    label="程序多开",
    variable=multi_instance_var,
    command=toggle_multi_instance
)

settings_menu.add_separator()
settings_menu.add_command(
    label="日志清理", 
    command=open_log_cleanup_dialog
)

settings_menu.add_command(
    label="代理设置",
    command=open_update_proxy_dialog
)

settings_menu.add_command(
    label="快捷方式",
    command=open_desktop_shortcut_dialog
)

settings_menu.add_command(
    label="语音播报",
    command=open_voice_text_dialog
)

settings_menu.add_command(
    label="短信字体",
    command=open_sms_font_dialog
)

settings_menu.add_command(
    label="串口调试", 
    command=open_serial_debug_window
)

menu_bar.add_cascade(label="设置", menu=settings_menu)

# 帮助
help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="关于", command=show_about)
help_menu.add_command(label="检测更新", command=check_update_and_prompt)
menu_bar.add_cascade(label="帮助", menu=help_menu)

root.config(menu=menu_bar)
update_voice_menu_label()

# ================= 启动 =================
schedule_next_midnight_clear()

if MODE == "Auto":
    set_status("🔍 自动模式：扫描 LUAT Modem 中…", "orange")
else:
    set_status(f"✍️ 手动模式：{PORT or '未指定'} @ {BAUD}", "orange")

threading.Thread(target=read_serial, daemon=True).start()
# 启动后自动清理定时器（默认60秒后首次运行）
schedule_auto_log_cleanup(restart=True, first_delay_sec=60)

root.mainloop()