# ================================================================
#  短信监听系统
#
#  功能简介：
#  ------------------------------------------------
#  本软件用于监听 LUAT 模组（如 Air724UG）串口输出，
#  实时识别指定短信回调日志，并进行如下处理：
#
#  1. 自动识别 LUAT Modem 串口（支持 Auto / Manual 模式）
#  2. 严格匹配短信回调标识 [handler_sms.smsCallback]
#  3. 支持关键词过滤（命中后才显示 / 弹窗 / 播报）
#  4. 支持语音播报（可自定义播报内容）
#  5. 支持短信弹窗提醒，电话弹窗提醒
#  6. 支持串口调试窗口（原始数据旁路）
#  7. 支持日志分端口记录与自动清理
#  8. 支持单实例运行（可选允许多开）
#  9. 支持开机自启与桌面快捷方式创建
# 10. 支持在线检测更新（支持代理）
#
#  作者：ChatGPT、Gemini、KPI0
#  GitHub：https://github.com/KPI0/Air724UG-SMS
# ================================================================

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
APP_VERSION = "3.5.4"  # 软件版本号
GITHUB_OWNER = "KPI0"
GITHUB_REPO = "Air724UG-SMS"

# 启动参数：开机自启时是否默认最小化到托盘
AUTOSTART_FLAG = "--autostart"
START_MINIMIZED = AUTOSTART_FLAG in sys.argv

# 程序启动时间（用于延迟启动时的 UI 日志显示）
APP_START_MONO = time.monotonic()
# 启动后延迟多少秒才在窗口显示日志（但仍会立即写文件）
START_UI_DELAY = 2.0

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
config = configparser.ConfigParser(interpolation=None)
if not os.path.exists(CONFIG_FILE):
    config["serial"] = {
        "port": "",
        "baud": "115200",
        "mode": "Auto",  # Auto / Manual
    }

    config["ui"] = {
        "voice_enabled": "1",         # 0=关闭语音播报，1=打开语音播报（默认）
        "popup_enabled": "1",         # 0=关闭短信弹窗，1=打开短信弹窗（默认）
        "voice_text": "注意！四川安播中心预警短信，请及时查看。",   # 默认语音播报内容
        "allow_multi_instance": "0",  # 0=禁止程序多开（默认），1=允许程序多开
        "auto_log_cleanup": "1",      # 0=关闭日志清理，1=打开日志清理（默认）
        "log_retention_days": "30",   # 日志保留时间，单位：天
        "desktop_shortcut_name": "短信监听系统",  # 默认桌面快捷方式名称
        "keywords": '["【四川安播中心】"]',  # 默认关键词
        "sms_font_size": "30",        # 默认字体大小
        "sms_font_color": "#ff0000",  # 默认字体颜色
        "call_filter_mode": "Disabled", # Disabled=关闭过滤，Whitelist=白名单，Blacklist=黑名单
        "call_whitelist": "[]",         # 默认来电白名单
        "call_blacklist": "[]",         # 默认来电黑名单

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

# ================= 短信弹窗开关（配置记忆） =================
try:
    POPUP_ENABLED = config.getboolean("ui", "popup_enabled", fallback=True)
except Exception:
    POPUP_ENABLED = True

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
KEYWORDS = []
try:
    raw = config.get("ui", "keywords", fallback="").strip()
    if raw:
        # 如果是 JSON 数组格式（新版存储），则用 json 解析
        if raw.startswith("[") and raw.endswith("]"):
            try:
                KEYWORDS = json.loads(raw)
            except Exception:
                # 解析失败兜底
                KEYWORDS = [x.strip() for x in raw.split("|") if x.strip()]
        else:
            # 兼容老版本的 "|" 分隔格式
            KEYWORDS = [x.strip() for x in raw.split("|") if x.strip()]
except Exception:
    pass

# ================= 防骚扰黑白名单（配置记忆） =================
try:
    CALL_FILTER_MODE = config.get("ui", "call_filter_mode", fallback="Disabled").strip()
except Exception:
    CALL_FILTER_MODE = "Disabled"

CALL_WHITELIST = []
try:
    raw = config.get("ui", "call_whitelist", fallback="").strip()
    if raw:
        CALL_WHITELIST = json.loads(raw)
except Exception:
    pass

CALL_BLACKLIST = []
try:
    raw = config.get("ui", "call_blacklist", fallback="").strip()
    if raw:
        CALL_BLACKLIST = json.loads(raw)
except Exception:
    pass

# ================= 串口控制 =================
serial_obj = None
serial_running = True
ring_timeout_target = 0.0 
current_dial_num = ""      

# ================= 串口对象并发保护（避免多线程 close/read 竞态） =================
serial_lock = threading.Lock()

def safe_close_serial():
    """线程安全关闭串口并置空 serial_obj"""
    global serial_obj
    with serial_lock:
        try:
            if serial_obj is not None:
                serial_obj.close()
        except Exception:
            pass
        serial_obj = None
        unlock_port_mutex()

# ================= 自动连接提示重复抑制 =================
_last_auto_connect_msg = None
_last_auto_connect_count = 0

def auto_connect_ui(msg: str):
    """自动连接提示：重复抑制，避免刷屏"""
    global _last_auto_connect_msg, _last_auto_connect_count
    try:
        if _last_auto_connect_msg == msg:
            _last_auto_connect_count += 1
        else:
            _last_auto_connect_msg = msg
            _last_auto_connect_count = 1

        if _last_auto_connect_count < ERROR_REPEAT_LIMIT:
            system_ui(msg, "normal")
        elif _last_auto_connect_count == ERROR_REPEAT_LIMIT:
            system_ui(msg + "（后续同类提示已忽略）", "normal")
        else:
            pass
    except Exception:
        try:
            system_ui(msg, "normal")
        except Exception:
            pass

# ================= 串口错误重复抑制 =================
# 用于抑制连续重复显示相同的串口异常（避免日志刷屏）
ERROR_REPEAT_LIMIT = 4  # 1~3 次显示详细错误；第 4 次显示“后续忽略”
_last_serial_error_msg = None
_last_serial_error_count = 0

# ================= 短信忽略重复抑制 =================
# 用于抑制连续重复显示相同的“短信未命中关键词，已忽略”提示（避免日志刷屏）
_last_sms_ignore_msg = None
_last_sms_ignore_count = 0

# ================= Manual 重绑提示重复抑制 =================
_last_rebind_hint_msg = None
_last_rebind_hint_count = 0

# ================= 全局变量 =================
PENDING_UI_LOGS = queue.Queue(maxsize=20000)  # 用于 text_area 未创建前缓存要显示到窗口的提示
LOG_PREFIX = "system"
AUTO_CLEANUP_INTERVAL_HOURS = 24 # 自动清理频率：24小时一次
AUTO_CLEANUP_AFTER_ID = None     # 记录 after() 的任务ID，用于避免重复定时器
SERIAL_DEBUG_ENABLED = False
serial_debug_queue = queue.Queue(maxsize=5000)  # 防止无限涨
serial_debug_win = None
serial_debug_text = None
serial_debug_drop_count = 0
serial_stop_event = threading.Event()
serial_wakeup_event = threading.Event()
TTS_LOCK = threading.Lock()
TTS_REQ_Q = queue.Queue(maxsize=50)
TTS_STOP = threading.Event()

# ================= Tk 线程安全：UI 任务队列（所有 Tk 操作只能在主线程） =================
UI_TASK_QUEUE = queue.Queue(maxsize=10000)
FILE_LOG_Q = queue.Queue(maxsize=50000)
file_log_stop = threading.Event()

def file_log_worker():
    while not file_log_stop.is_set():
        try:
            # 阻塞获取第一条，如果0.5秒没日志就休息
            path, line = FILE_LOG_Q.get(timeout=0.5)
        except queue.Empty:
            continue
            
        # 核心优化：瞬间吸干当前队列里所有积压的日志，按文件路径打包
        batch = {path: [line]}
        while True:
            try:
                p, l = FILE_LOG_Q.get_nowait()
                if p not in batch:
                    batch[p] = []
                batch[p].append(l)
            except queue.Empty:
                break # 吸干了，跳出循环
                
        # 批量写入：无论刚才吸了 10 行还是 100 行，每个文件只打开/关闭一次！
        for p, lines in batch.items():
            try:
                with open(p, "a", encoding="utf-8") as f:
                    f.writelines(lines)
            except Exception:
                pass

threading.Thread(target=file_log_worker, daemon=True).start()

# ================= root 是否可用（避免退出过程中 after 抛异常） =================
TK_SHUTDOWN = threading.Event()

def tk_alive() -> bool:
    # 后台线程：绝对不调用任何 Tk 方法
    if TK_SHUTDOWN.is_set():
        return False
    if ("root" not in globals()) or (root is None):
        return False

    # 只有主线程才允许 winfo_exists()
    if threading.current_thread() is threading.main_thread():
        try:
            return bool(root.winfo_exists())
        except Exception:
            return False

    return True

def ui_post(fn, *args, **kwargs):
    try:
        UI_TASK_QUEUE.put_nowait((fn, args, kwargs))
    except queue.Full:
        log_file_only("⚠️ UI_TASK_QUEUE 已满：丢弃一次 UI 任务")
    except Exception:
        pass

def ui_pump(max_batch=200):
    """
    主线程定时执行队列里的 UI 操作。
    """
    n = 0
    while n < max_batch:
        try:
            fn, args, kwargs = UI_TASK_QUEUE.get_nowait()
        except queue.Empty:
            break
        try:
            fn(*args, **kwargs)
        except Exception:
            pass
        n += 1

    # 继续下一轮
    if tk_alive():
        try:
            root.after(30, ui_pump)
        except Exception:
            pass

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
        line = f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n"
        FILE_LOG_Q.put_nowait((system_log, line))
    except Exception:
        pass

# ================= ui_only：永远线程安全（只 UI 不写文件） =================
def ui_only(msg: str, tag="normal"):
    """只显示到窗口，不写任何日志文件（不写 COM 日志）"""
    def _do():
        try:
            if ("text_area" in globals()) and (text_area is not None) and text_area.winfo_exists():
                safe_insert_main_text(msg, tag)
            else:
                # UI 未就绪则缓存
                try:
                    PENDING_UI_LOGS.put_nowait((msg, tag))
                except Exception:
                    pass
        except Exception:
            pass

    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

# ================= log_early：写文件 + 线程安全缓存 =================
def log_early(msg: str, tag: str = "normal"):
    """早期日志：先写文件，再缓存，等 text_area 创建后补到窗口"""
    log_file_only(msg)
    try:
        PENDING_UI_LOGS.put_nowait((msg, tag))
    except queue.Full:
        pass
    except Exception:
        pass

# ================= system_ui：写 system 文件 + 投递 UI（线程安全） =================
def system_ui(message: str, tag="normal"):
    # 0) Tk 尚未就绪/退出过程中：只写文件 + 缓存（不碰 root.after）
    if not tk_alive():
        log_file_only(message)
        try:
            PENDING_UI_LOGS.put_nowait((message, tag))
        except Exception:
            pass
        return
    # 1) 永远先写 system 日志（文件写入不依赖 Tk）
    log_file_only(message)

    def _do_ui():
        try:
            if ("text_area" in globals()) and (text_area is not None) and text_area.winfo_exists():
                safe_insert_main_text(message, tag)
            else:
                # UI 不可用：缓存给后续补显示
                PENDING_UI_LOGS.put_nowait((message, tag))
        except Exception:
            try:
                PENDING_UI_LOGS.put_nowait((message, tag))
            except Exception:
                pass

    # 2) respect START_UI_DELAY：延迟也要在主线程安排
    def _schedule_in_main():
        try:
            elapsed = time.monotonic() - APP_START_MONO
            delay_ms = int(max(0.0, START_UI_DELAY - elapsed) * 1000)
        except Exception:
            delay_ms = 0

        try:
            if delay_ms > 0:
                root.after(delay_ms, _do_ui)
            else:
                _do_ui()
        except Exception:
            _do_ui()

    if threading.current_thread() is threading.main_thread():
        _schedule_in_main()
    else:
        ui_post(_schedule_in_main)

# ================= port_ui：统一走 log()（线程安全） =================
def port_ui(message: str, tag="normal"):
    # log() 已经线程安全了，这里无需再做各种 root/text_area 判断
    log(message, tag=tag)

def set_autostart(enable: bool):
    try:
        if enable:
            create_startup_shortcut()
            msg = "✅️ 开机自启：已打开"
        else:
            remove_startup_shortcut()
            msg = "❌ 开机自启：已关闭"

        system_ui(msg, "normal")

    except Exception as e:
        ui_messagebox("error", "错误", f"设置开机自启失败：\n{e}")

# ================= TTS语音播报 =================
def _tts_worker():
    while not TTS_STOP.is_set():
        try:
            task = TTS_REQ_Q.get(timeout=0.5)
        except queue.Empty:
            continue

        # 兼容旧参数和新参数解包
        if len(task) == 3:
            text, force, play_after = task
        else:
            text, force = task
            play_after = False

        if not text:
            text = DEFAULT_VOICE_TEXT

        # 如果不强制生成，且文件已存在，直接跳过
        if (not force) and os.path.exists(TTS_FILE):
            try:
                TTS_REQ_Q.task_done()
            except Exception:
                pass
            if play_after:
                play_alert(force=True)
            continue

        try:
            with TTS_LOCK:
                os.makedirs(os.path.dirname(TTS_FILE), exist_ok=True)
                tmp_path = TTS_FILE + ".tmp.wav"
                
                # 核心修复 1：每次重新 init，避开 pyttsx3 长驻缓存的 Bug
                engine = pyttsx3.init()
                engine.setProperty("rate", 150)
                engine.save_to_file(text, tmp_path)
                engine.runAndWait()
                engine.stop()
                del engine  # 核心修复 2：强制释放 COM 内存，防止内存泄漏
                
                # 原子替换，只有成功生成了才覆盖原文件
                if os.path.exists(tmp_path):
                    os.replace(tmp_path, TTS_FILE)
                    
        except Exception as e:
            # 清理可能损坏的 tmp 文件
            try:
                if os.path.exists(TTS_FILE + ".tmp.wav"):
                    os.remove(TTS_FILE + ".tmp.wav")
            except Exception:
                pass
            log_file_only(f"TTS 生成失败，使用系统声音兜底：{e}")
        finally:
            try:
                TTS_REQ_Q.task_done()
            except Exception:
                pass
        
        # 核心修复 3：文件安全替换完成后，再进行回调播放（彻底解决并发文件锁问题）
        if play_after:
            play_alert(force=True)

threading.Thread(target=_tts_worker, daemon=True).start()

def generate_alert_voice(force: bool = False, text: str = None, play_after: bool = False):
    """
    对外接口：任何线程都可调用。
    - play_after: 专门用于试听，生成完毕后立刻回调播放，防止文件锁冲突
    """
    try:
        if text is None:
            text_snapshot = (VOICE_TEXT or DEFAULT_VOICE_TEXT).strip() or DEFAULT_VOICE_TEXT
        else:
            text_snapshot = (text or "").strip() or DEFAULT_VOICE_TEXT
    except Exception:
        text_snapshot = DEFAULT_VOICE_TEXT

    try:
        # 核心防御：清空积压队列（防抖）。如果用户狂点“试听”，直接丢弃旧任务，只执行最后一次
        while not TTS_REQ_Q.empty():
            try:
                TTS_REQ_Q.get_nowait()
                TTS_REQ_Q.task_done()
            except Exception:
                break

        # 把 play_after 也打包塞进队列
        TTS_REQ_Q.put_nowait((text_snapshot, bool(force), bool(play_after)))
    except queue.Full:
        log_file_only("⚠️ TTS 请求队列已满，已丢弃一次生成请求")

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
        arg_line = f'"{args}"'   # 脚本路径加引号，防空格
        vbs += f'Shortcut.Arguments = {vbs_quote(arg_line)}\n'

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

    bottom_line = tk.Frame(win, height=1, bg="#d4d4d4")
    bottom_line.pack(side="bottom", fill="x")
    bottom_line.pack_propagate(False)

    frame = tk.Frame(win, padx=14, pady=12)
    frame.pack(side="top", fill=tk.BOTH, expand=True)

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
        # 1. 强制系统处理所有事件并完成重绘，获取真实物理尺寸
        preview_canvas.update()

        try:
            s = int(size_var.get().strip())
        except Exception:
            s = SMS_FONT_SIZE

        c = (color_var.get().strip() or SMS_FONT_COLOR)

        # 2. 终极防呆兜底：如果极端情况下依然返回 1，则使用创建时的默认物理像素 (560x110)
        cw = preview_canvas.winfo_width()
        ch = preview_canvas.winfo_height()
        if cw <= 1: cw = 560
        if ch <= 1: ch = 110

        # 预览用字号：避免裁剪（高度的 70% 比较合适）
        max_size = max(8, int(ch * 0.7))
        s_preview = min(s, max_size)

        preview_canvas.delete("all")
        try:
            preview_canvas.create_text(
                cw // 2,
                ch // 2,
                text=PREVIEW_TEXT,
                anchor="c",
                font=("微软雅黑", s_preview),
                fill=c
            )
        except Exception:
            preview_canvas.create_text(
                cw // 2,
                ch // 2,
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
    serial_debug_win.minsize(800, 300)
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

        # ============ 控制发送按钮和输入框变灰 ============
        try:
            # 简化写法：根据 running 状态决定是恢复(!disabled)还是禁用(disabled)
            state_val = ["!disabled"] if running else ["disabled"]
            
            btn_send.state(state_val)
            send_entry.state(state_val)
            btn_quick.state(state_val)
            
            # 遍历右侧面板的所有快捷按钮，统统跟随变灰/恢复
            for child in quick_scroll_frame.winfo_children():
                if isinstance(child, ttk.Button):
                    child.state(state_val)
        except Exception:
            pass
        # ================================================

    def _set_pause_state(is_paused: bool):
        nonlocal pause_banner_shown
        is_paused = bool(is_paused)

        # 状态没变化，直接返回（防刷）
        if paused_var.get() == is_paused:
            return

        paused_var.set(is_paused)

        if is_paused:
            btn_pause.config(text="▶ 继续")

            # 暂停时仍保持旁路开关锁定
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

            # 继续时锁定旁路开关
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

    # ================= 发送命令区 =================
    send_frame = ttk.Frame(serial_debug_win)
    send_frame.pack(side="bottom", fill="x", padx=8, pady=(0, 6))

    send_var = tk.StringVar()
    ttk.Label(send_frame, text="发送指令：").pack(side="left")
    
    # 输入框
    send_entry = ttk.Entry(send_frame, textvariable=send_var)
    send_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

    # AT 指令通常需要回车换行
    crlf_var = tk.BooleanVar(value=True)
    ttk.Checkbutton(send_frame, text="加回车换行(\\r\\n)", variable=crlf_var).pack(side="left", padx=(0, 8))

    def _send_cmd(_event=None):
        if not enabled_var.get():
            return "break"
        cmd = send_var.get()
        if not cmd:
            return "break"
        
        # 处理换行符并转码
        cmd_bytes = (cmd + "\r\n").encode("utf-8", "ignore") if crlf_var.get() else cmd.encode("utf-8", "ignore")
        display_suffix = "\\r\\n" if crlf_var.get() else ""
        
        # 将发送动作包裹在一个独立函数里，丢进后台线程执行
        def _send_task():
            global serial_obj, serial_lock
            with serial_lock:
                if serial_obj is not None and serial_obj.is_open:
                    try:
                        serial_obj.write(cmd_bytes)
                        serial_obj.flush()
                        _push_serial_debug(f">>> 发送: {cmd}{display_suffix}")
                    except Exception as e:
                        _push_serial_debug(f">>> 发送失败: {e}")
                else:
                    _push_serial_debug(">>> 发送失败: 串口未连接")
                    
        # 开辟守护线程发送指令，彻底解放 GUI 主线程
        threading.Thread(target=_send_task, daemon=True).start()
        return "break"

    btn_send = ttk.Button(send_frame, text="发送", width=8, command=_send_cmd)
    btn_send.pack(side="left")

    # ============ 快捷命令发送函数与展开按钮 ============
    def _quick_send(cmd):
        send_var.set(cmd)
        _send_cmd()

    btn_quick = ttk.Button(send_frame, text="快捷命令 ▶")
    btn_quick.pack(side="left", padx=(8, 0))
    # ==========================================================

    # 绑定回车键快捷发送
    send_entry.bind("<Return>", _send_cmd)
    # ====================================================

    body = ttk.Frame(serial_debug_win)
    body.pack(fill="both", expand=True, padx=8, pady=6)

    # 告诉 body 使用网格布局：左边（第0列）自动伸缩，右边（第1列）自适应大小
    body.grid_rowconfigure(0, weight=1)
    body.grid_columnconfigure(0, weight=1)
    body.grid_columnconfigure(1, weight=0)

    # ==== 1. 左侧文本框容器 ====
    text_frame = ttk.Frame(body)
    text_frame.grid(row=0, column=0, sticky="nsew")

    yscroll = ttk.Scrollbar(text_frame, orient="vertical")
    yscroll.pack(side="right", fill="y")

    serial_debug_text = tk.Text(text_frame, wrap="none", yscrollcommand=yscroll.set)
    serial_debug_text.pack(side="left", fill="both", expand=True)
    yscroll.config(command=serial_debug_text.yview)

    # ==== 2. 右侧常用指令面板 ====
    quick_panel = ttk.LabelFrame(body, text="常用指令")

    # 给定一个适当的初始宽度防止太窄
    quick_canvas = tk.Canvas(quick_panel, highlightthickness=0, width=330) 
    quick_scrollbar = ttk.Scrollbar(quick_panel, orient="vertical", command=quick_canvas.yview)
    quick_scroll_frame = ttk.Frame(quick_canvas)

    # 将 Frame 放入 Canvas
    quick_scroll_window = quick_canvas.create_window((0, 0), window=quick_scroll_frame, anchor="nw")
    quick_canvas.configure(yscrollcommand=quick_scrollbar.set)

    quick_scrollbar.pack(side="right", fill="y")
    quick_canvas.pack(side="left", fill="both", expand=True)

    # 动态更新滚动区域（当按钮增多时，自动拉长可滚动范围）
    quick_scroll_frame.bind("<Configure>", lambda e: quick_canvas.configure(scrollregion=quick_canvas.bbox("all")))
    # 动态撑满宽度（让按钮横向填满面板）
    quick_canvas.bind("<Configure>", lambda e: quick_canvas.itemconfig(quick_scroll_window, width=e.width))

    # 绑定鼠标滚轮 (只有当鼠标悬浮在指令区时才生效)
    def _bind_mousewheel(event):
        quick_canvas.bind_all("<MouseWheel>", lambda e: quick_canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
    def _unbind_mousewheel(event):
        quick_canvas.unbind_all("<MouseWheel>")
        
    quick_canvas.bind("<Enter>", _bind_mousewheel)
    quick_canvas.bind("<Leave>", _unbind_mousewheel)
    # ==========================================================
    
    common_cmds = [
        ("AT", "测试通信"),
        ("ATI", "查模块信息"),
        ("AT+CGMR", "查固件版本"),
        ("AT+CSQ", "查信号(RSSI/通用)"),
        ("AT+CESQ", "查精确信号(4G RSRP)"),
        ("AT+CGSN", "查模组IMEI"),
        ("AT+WISN?", "查模组SN"),
        ("AT+MIFIMAC=R", "查WiFi热点MAC地址"),
        ("AT+CGPADDR", "查PDP上下文IP地址"),
        ("AT+CGDCONT?", "查APN配置"),
        ("AT+RFTEMPERATURE?", "查模组温度"),
        ("AT+CNUM", "查本机号码"),
        ("AT+CSCA?", "查短信中心号码"),
        ("AT+COPS?", "查运营商"),
        ("AT+CPIN?", "查PIN码锁状态"),
        ("AT+ICCID", "查SIM卡ICCID"),
        ("AT+CIMI", "查SIM卡IMSI"),
        ("AT+CGATT?", "查网络附着"),
        ("AT+CFUN=1,1", "重启基带"),
        ("AT+RESET", "重启模组"),
        ("AT+CFUN?", "查看当前飞行模式状态"),
        ("AT+CFUN=0", "打开飞行模式"),
        ("AT+CFUN=1", "关闭飞行模式"),
        ("AT+EEMGINFO?", "查基站定位数据"),

    ]
    
    # 循环生成竖排按钮
    for cmd, desc in common_cmds:
        btn_text = f"{cmd}  ({desc})"
        ttk.Button(quick_scroll_frame, text=btn_text, command=lambda c=cmd: _quick_send(c)).pack(fill="x", padx=6, pady=3)
    
    # ================= 输入PIN码解锁弹窗 =================
    def _open_input_pin_dialog():
        win = tk.Toplevel(serial_debug_win)
        win.title("输入PIN码解锁")
        win.resizable(False, False)
        win.transient(serial_debug_win)
        win.grab_set()

        frm = ttk.Frame(win, padding=15)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="请输入 SIM 卡 PIN 码：").pack(anchor="w", pady=(0, 10))
        pin_var = tk.StringVar()
        ent = ttk.Entry(frm, textvariable=pin_var, width=28)
        ent.pack(fill="x", pady=(0, 5))  # 把底部的 pady 改小，让提示紧贴输入框

        # ================= 红色警告提示文本 =================
        tk.Label(
            frm, 
            text="⚠️ 警告：连续3次错误将锁定，需PUK码解锁！", 
            fg="#d9534f",  # 使用醒目的偏红色作为警告
            font=("微软雅黑", 9)
        ).pack(anchor="w", pady=(0, 15))
        def _do_unlock():
            if not enabled_var.get():
                messagebox.showwarning("提示", "请先勾选顶部的“启用原始输出旁路”", parent=win)
                return

            pin = pin_var.get().strip()
            if not pin:
                messagebox.showerror("错误", "PIN码不能为空", parent=win)
                return
            
            win.destroy()
            # 自动拼接成 AT+CPIN="xxxx" 的格式并直接发送
            _quick_send(f'AT+CPIN="{pin}"')

        btn_frm = ttk.Frame(frm)
        btn_frm.pack(anchor="e")
        ttk.Button(btn_frm, text="发送指令", command=_do_unlock).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frm, text="取消", command=win.destroy).pack(side="left")

        win.update_idletasks()
        center_window(win, serial_debug_win)
        ent.focus_set()
        win.bind("<Return>", lambda e: _do_unlock())
        win.bind("<Escape>", lambda e: win.destroy())

    # ================= 输入PUK码解锁弹窗 =================
    def _open_input_puk_dialog():
        win = tk.Toplevel(serial_debug_win)
        win.title("输入PUK码解锁")
        win.resizable(False, False)
        win.transient(serial_debug_win)
        win.grab_set()

        frm = ttk.Frame(win, padding=15)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="请输入 PUK 码 (通常为8位)：").pack(anchor="w")
        puk_var = tk.StringVar()
        ent_puk = ttk.Entry(frm, textvariable=puk_var, width=28)
        ent_puk.pack(fill="x", pady=(2, 2))  # 把底部的边距调小，让警告语紧贴它

        # ================= 红色警告提示文本移到这里 =================
        tk.Label(
            frm, 
            text="⚠️ 致命警告：连续10次错误将永久烧毁SIM卡！", 
            fg="#d9534f",  # 醒目红色
            font=("微软雅黑", 9, "bold")
        ).pack(anchor="w", pady=(0, 12))
        # =========================================================

        ttk.Label(frm, text="请设置 新 PIN 码 (通常为4-8位数字)：").pack(anchor="w")
        new_pin_var = tk.StringVar()
        ent_new_pin = ttk.Entry(frm, textvariable=new_pin_var, width=28)
        ent_new_pin.pack(fill="x", pady=(2, 15))

        def _do_unlock_puk():
            if not enabled_var.get():
                messagebox.showwarning("提示", "请先勾选顶部的“启用原始输出旁路”", parent=win)
                return

            puk = puk_var.get().strip()
            new_pin = new_pin_var.get().strip()
            
            if not puk or not new_pin:
                messagebox.showerror("错误", "PUK码和新PIN码都不能为空！", parent=win)
                return
            
            win.destroy()
            # 自动拼接成 AT+CPIN="puk","new_pin" 的格式并直接发送
            _quick_send(f'AT+CPIN="{puk}","{new_pin}"')

        btn_frm = ttk.Frame(frm)
        btn_frm.pack(anchor="e")
        ttk.Button(btn_frm, text="发送指令", command=_do_unlock_puk).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frm, text="取消", command=win.destroy).pack(side="left")

        win.update_idletasks()
        center_window(win, serial_debug_win)
        ent_puk.focus_set()
        win.bind("<Return>", lambda e: _do_unlock_puk())
        win.bind("<Escape>", lambda e: win.destroy())

    # ================= 开启PIN码锁专属功能 =================
    def _open_enable_pin_dialog():
        win = tk.Toplevel(serial_debug_win)
        win.title("开启PIN码锁")
        win.resizable(False, False)
        win.transient(serial_debug_win)
        win.grab_set()

        frm = ttk.Frame(win, padding=15)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="请输入当前的 PIN 码以开启锁定：").pack(anchor="w", pady=(0, 10))
        pin_var = tk.StringVar()
        ent = ttk.Entry(frm, textvariable=pin_var, width=28)
        ent.pack(fill="x", pady=(0, 5))

        # ================= 灰色提示文本 =================
        tk.Label(
            frm, 
            text="💡 提示：开启后模组每次开机均需输入PIN码。", 
            fg="gray", 
            font=("微软雅黑", 9)
        ).pack(anchor="w", pady=(0, 15))
        # ===============================================

        def _do_enable():
            if not enabled_var.get():
                messagebox.showwarning("提示", "请先勾选顶部的“启用原始输出旁路”", parent=win)
                return

            pin = pin_var.get().strip()
            if not pin:
                messagebox.showerror("错误", "PIN码不能为空", parent=win)
                return
            
            win.destroy()
            # 自动拼接成 AT+CLCK="SC",1,"xxxx" 的格式并直接发送
            _quick_send(f'AT+CLCK="SC",1,"{pin}"')

        btn_frm = ttk.Frame(frm)
        btn_frm.pack(anchor="e")
        ttk.Button(btn_frm, text="发送指令", command=_do_enable).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frm, text="取消", command=win.destroy).pack(side="left")

        win.update_idletasks()
        center_window(win, serial_debug_win)
        ent.focus_set()
        win.bind("<Return>", lambda e: _do_enable())
        win.bind("<Escape>", lambda e: win.destroy())

    # ================= 关闭PIN码锁专属功能 =================
    def _open_disable_pin_dialog():
        win = tk.Toplevel(serial_debug_win)
        win.title("关闭PIN码锁")
        win.resizable(False, False)
        win.transient(serial_debug_win)
        win.grab_set()

        frm = ttk.Frame(win, padding=15)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="请输入当前的 PIN 码以关闭锁定：").pack(anchor="w", pady=(0, 10))
        pin_var = tk.StringVar()
        ent = ttk.Entry(frm, textvariable=pin_var, width=28)
        ent.pack(fill="x", pady=(0, 5))

        # ================= 灰色提示文本 =================
        tk.Label(
            frm, 
            text="💡 提示：关闭后模组开机将自动联网，不再拦截。", 
            fg="gray", 
            font=("微软雅黑", 9)
        ).pack(anchor="w", pady=(0, 15))
        # ===============================================

        def _do_disable():
            if not enabled_var.get():
                messagebox.showwarning("提示", "请先勾选顶部的“启用原始输出旁路”", parent=win)
                return

            pin = pin_var.get().strip()
            if not pin:
                messagebox.showerror("错误", "PIN码不能为空", parent=win)
                return
            
            win.destroy()
            # 自动拼接成 AT+CLCK="SC",0,"xxxx" 的格式并直接发送
            _quick_send(f'AT+CLCK="SC",0,"{pin}"')

        btn_frm = ttk.Frame(frm)
        btn_frm.pack(anchor="e")
        ttk.Button(btn_frm, text="发送指令", command=_do_disable).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frm, text="取消", command=win.destroy).pack(side="left")

        win.update_idletasks()
        center_window(win, serial_debug_win)
        ent.focus_set()
        win.bind("<Return>", lambda e: _do_disable())
        win.bind("<Escape>", lambda e: win.destroy())

    # ================= 修改PIN码专属功能 =================
    def _open_modify_pin_dialog():
        win = tk.Toplevel(serial_debug_win)
        win.title("修改PIN码")
        win.resizable(False, False)
        win.transient(serial_debug_win)
        win.grab_set()

        frm = ttk.Frame(win, padding=15)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="请输入 旧 PIN 码：").pack(anchor="w")
        old_pin_var = tk.StringVar()
        ent_old = ttk.Entry(frm, textvariable=old_pin_var, width=28)
        ent_old.pack(fill="x", pady=(2, 10))

        ttk.Label(frm, text="请输入 新 PIN 码 (通常为4-8位数字)：").pack(anchor="w")
        new_pin_var = tk.StringVar()
        ent_new = ttk.Entry(frm, textvariable=new_pin_var, width=28)
        ent_new.pack(fill="x", pady=(2, 15))

        def _do_modify_pin():
            if not enabled_var.get():
                messagebox.showwarning("提示", "请先勾选顶部的“启用原始输出旁路”", parent=win)
                return

            old_pin = old_pin_var.get().strip()
            new_pin = new_pin_var.get().strip()

            if not old_pin or not new_pin:
                messagebox.showerror("错误", "旧PIN码和新PIN码都不能为空", parent=win)
                return
            
            win.destroy()
            # 自动拼接成 AT+CPWD="SC","old","new" 的格式并直接发送
            _quick_send(f'AT+CPWD="SC","{old_pin}","{new_pin}"')

        btn_frm = ttk.Frame(frm)
        btn_frm.pack(anchor="e")
        ttk.Button(btn_frm, text="发送指令", command=_do_modify_pin).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frm, text="取消", command=win.destroy).pack(side="left")

        win.update_idletasks()
        center_window(win, serial_debug_win)
        ent_old.focus_set()
        win.bind("<Return>", lambda e: _do_modify_pin())
        win.bind("<Escape>", lambda e: win.destroy())

    # ================= 修改本机号码专属功能 =================
    def _open_modify_number_dialog():
        win = tk.Toplevel(serial_debug_win)
        win.title("修改本机号码")
        win.resizable(False, False)
        win.transient(serial_debug_win)
        win.grab_set()

        frm = ttk.Frame(win, padding=15)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="请输入新的手机号码：").pack(anchor="w", pady=(0, 10))
        num_var = tk.StringVar()
        ent = ttk.Entry(frm, textvariable=num_var, width=28)
        ent.pack(fill="x", pady=(0, 5))  # 把底部的 pady 改小，让提示紧贴输入框

        # ================= 灰色提示文本 =================
        tk.Label(
            frm, 
            text="💡 提示：需加 '+' 国际前缀 (如 +8618888888...)。", 
            fg="gray", 
            font=("微软雅黑", 9)
        ).pack(anchor="w", pady=(0, 15))
        # ===============================================

        def _do_modify():
            if not enabled_var.get():
                messagebox.showwarning("提示", "请先勾选顶部的“启用原始输出旁路”", parent=win)
                return

            phone = num_var.get().strip()
            if not phone:
                messagebox.showerror("错误", "手机号码不能为空", parent=win)
                return
            
            # ================= 防呆/自动容错设计 =================
            # 如果用户忘了输 + 号，系统自动帮他补上国内的 +86 前缀
            if not phone.startswith("+"):
                phone = "+86" + phone
            # ====================================================
            
            win.destroy()
            
            # 准备要发送的两条指令 (保持使用 145 国际格式)
            cmd1 = 'AT+CPBS="ON"'
            cmd2 = f'AT+CPBW=1,"{phone}",145'

            # 放到后台线程发送，防止 time.sleep 卡死软件界面
            def _send_task():
                global serial_obj, serial_lock
                with serial_lock:
                    if serial_obj is not None and serial_obj.is_open:
                        try:
                            # 发送第一条：设置电话本为本机
                            serial_obj.write((cmd1 + "\r\n").encode("utf-8"))
                            serial_obj.flush()
                            _push_serial_debug(f">>> 发送: {cmd1}\\r\\n")
                            
                            # 延时 0.3 秒，给模组一点处理时间
                            time.sleep(0.3)
                            
                            # 发送第二条：写入号码
                            serial_obj.write((cmd2 + "\r\n").encode("utf-8"))
                            serial_obj.flush()
                            _push_serial_debug(f">>> 发送: {cmd2}\\r\\n")
                        except Exception as e:
                            _push_serial_debug(f">>> 发送失败: {e}")
                    else:
                        _push_serial_debug(">>> 发送失败: 串口未连接")
                        
            threading.Thread(target=_send_task, daemon=True).start()

        btn_frm = ttk.Frame(frm)
        btn_frm.pack(anchor="e")
        ttk.Button(btn_frm, text="发送指令", command=_do_modify).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frm, text="取消", command=win.destroy).pack(side="left")

        win.update_idletasks()
        center_window(win, serial_debug_win)
        ent.focus_set()
        win.bind("<Return>", lambda e: _do_modify())
        win.bind("<Escape>", lambda e: win.destroy())

    # ================= 修改设备SN码专属功能 =================
    def _open_modify_sn_dialog():
        win = tk.Toplevel(serial_debug_win)
        win.title("修改设备SN码")
        win.resizable(False, False)
        win.transient(serial_debug_win)
        win.grab_set()

        frm = ttk.Frame(win, padding=15)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="请输入新的 SN 码 (最长64位)：").pack(anchor="w", pady=(0, 10))
        sn_var = tk.StringVar()
        ent = ttk.Entry(frm, textvariable=sn_var, width=28)
        ent.pack(fill="x", pady=(0, 5))

        # ================= 警告与提示文本 =================
        tk.Label(
            frm, 
            text="⚠️ 警告：修改SN码可能导致串口异常，请尝试多次插拔。", 
            fg="#d9534f", 
            justify="left",
            font=("微软雅黑", 9)
        ).pack(anchor="w", pady=(0, 15))
        # ===============================================

        def _do_modify():
            if not enabled_var.get():
                messagebox.showwarning("提示", "请先勾选顶部的“启用原始输出旁路”", parent=win)
                return

            new_sn = sn_var.get().strip()
            if not new_sn:
                messagebox.showerror("错误", "SN码不能为空", parent=win)
                return
                
            # 根据官方手册增加长度防呆拦截
            if len(new_sn) > 64:
                messagebox.showerror("错误", "SN码最长不能超过64位", parent=win)
                return
            
            win.destroy()
            
            # 使用官方手册提供的专有指令：AT+WISN
            _quick_send(f'AT+WISN={new_sn}')

        btn_frm = ttk.Frame(frm)
        btn_frm.pack(anchor="e")
        ttk.Button(btn_frm, text="发送指令", command=_do_modify).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frm, text="取消", command=win.destroy).pack(side="left")

        win.update_idletasks()
        center_window(win, serial_debug_win)
        ent.focus_set()
        win.bind("<Return>", lambda e: _do_modify())
        win.bind("<Escape>", lambda e: win.destroy())
    
    # ================= 发送短信专属功能 (PDU 模式完美兼容版) =================
    def _open_send_sms_dialog():
        win = tk.Toplevel(serial_debug_win)
        win.title("发送短信")
        win.resizable(False, False)
        win.transient(serial_debug_win)
        win.grab_set()

        frm = ttk.Frame(win, padding=15)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="接收方手机号：").pack(anchor="w")
        phone_var = tk.StringVar()
        ent_phone = ttk.Entry(frm, textvariable=phone_var, width=32)
        ent_phone.pack(fill="x", pady=(2, 10))

        ttk.Label(frm, text="短信内容：").pack(anchor="w")
        txt_msg = tk.Text(frm, height=4, width=30, font=("微软雅黑", 9))
        txt_msg.pack(fill="x", pady=(2, 2))  # 缩小底边距，让字数统计贴紧输入框

        # ================= 实时字数统计指示器 =================
        count_var = tk.StringVar(value="已输入: 0 / 70 字")
        count_label = tk.Label(frm, textvariable=count_var, fg="gray", font=("微软雅黑", 8))
        count_label.pack(anchor="e", pady=(0, 5))  # 靠右(e)对齐

        def _update_char_count(event=None):
            # end-1c 是为了不把 Text 控件自带的末尾换行符算进去
            content = txt_msg.get("1.0", "end-1c")
            length = len(content)
            count_var.set(f"已输入: {length} / 70 字")
            
            # 如果超长，字数统计变红警告
            if length > 70:
                count_label.config(fg="#d9534f")
            else:
                count_label.config(fg="gray")

        # 绑定键盘松开事件，每次打字都会触发字数刷新
        txt_msg.bind("<KeyRelease>", _update_char_count)

        # ================= 提示文本更新 =================
        tk.Label(
            frm, 
            text="注意: 当前仅支持单条发送，最大限制 70 字符。\n💡 提示：支持 '+' 国际前缀 (如 +8618888888...)。", 
            fg="gray", 
            justify="left",
            font=("微软雅黑", 9)
        ).pack(anchor="w", pady=(5, 10))
        # ===============================================

        def _do_send_sms():
            if not enabled_var.get():
                messagebox.showwarning("提示", "请先勾选顶部的“启用原始输出旁路”", parent=win)
                return

            phone = phone_var.get().strip()
            # 获取时去除首尾空白
            msg = txt_msg.get("1.0", "end-1c").strip()
            
            if not phone or not msg:
                messagebox.showerror("错误", "手机号和短信内容不能为空！", parent=win)
                return
                
            # ================= 发送前拦截超长限制 =================
            if len(msg) > 70:
                messagebox.showerror(
                    "字数超限", 
                    f"当前输入了 {len(msg)} 个字符，已超过 70 字符最大限制！\n请删减后再发送。", 
                    parent=win
                )
                return
            # =====================================================
            
            win.destroy()
            
            # 放到后台线程发送
            def _send_task():
                global serial_obj, serial_lock
                with serial_lock:
                    if serial_obj is not None and serial_obj.is_open:
                        try:
                            # =============================================
                            # 核心算法：将明文和手机号转化为标准 PDU 十六进制码
                            # =============================================
                            def _encode_pdu(p_str, m_str):
                                p_str = p_str.strip()
                                # 91 代表带+的国际格式，81 代表国内通用格式
                                p_type = "91" if p_str.startswith("+") else "81"
                                p_num = p_str.lstrip("+")
                                p_len = f"{len(p_num):02X}" # 长度
                                
                                # 电话号码奇偶位反转 (如 138 变 31F8)
                                if len(p_num) % 2 != 0:
                                    p_num += "F"
                                p_swap = "".join([p_num[i+1] + p_num[i] for i in range(0, len(p_num), 2)])
                                
                                # 短信内容使用 UCS2 (utf-16-be) 编码
                                m_bytes = m_str.encode("utf-16-be")
                                udl = f"{len(m_bytes):02X}" # 内容长度
                                ud = m_bytes.hex().upper()  # 内容HEX
                                
                                # 1. 短信中心号码 (SMSC) 信息。00 表示使用 SIM 卡内设定的默认中心号码
                                smsc = "00"
                                
                                # 2. 传送协议数据单元 (TPDU)
                                # 11(Type) 00(MR) [长度] [类型] [反转号码] 00(PID) 08(UCS2编码) C0(有效期) [正文长度] [正文HEX]
                                tpdu = f"1100{p_len}{p_type}{p_swap}0008C0{udl}{ud}"
                                
                                # 完整的 PDU 是 SMSC + TPDU
                                p_data = smsc + tpdu
                                
                                # 3. AT+CMGS 所需的长度是严格的 TPDU 字节数（完全排除 SMSC）
                                c_len = len(tpdu) // 2
                                
                                return p_data, c_len
                            # ==========================================

                            pdu_str, cmgs_len = _encode_pdu(phone, msg)

                            # ======== 将发送行为写进主日志和界面 ========
                            port_ui(f"📤 发送短信至 {phone}：", "normal")
                            port_ui(msg, "sms")
                            # =================================================

                            # 1. 切换到 PDU 模式 (0)
                            cmd1 = "AT+CMGF=0"
                            serial_obj.write((cmd1 + "\r\n").encode("utf-8"))
                            serial_obj.flush()
                            _push_serial_debug(f">>> 发送: {cmd1}\\r\\n")
                            time.sleep(0.3) # 稍微给模组一点反应时间
                            
                            # 2. 发送目标数据长度
                            cmd2 = f"AT+CMGS={cmgs_len}"
                            serial_obj.write((cmd2 + "\r\n").encode("utf-8"))
                            serial_obj.flush()
                            _push_serial_debug(f">>> 发送: {cmd2}\\r\\n")
                            
                            # 【极其关键】：发送完 CMGS 后，必须强行等 1 秒，让模组吐出 "> "
                            time.sleep(1.0)
                            
                            # 3. 发送 PDU 字符串，并追加 \x1a (Ctrl+Z) 结束符
                            payload = pdu_str.encode("utf-8") + b"\x1a"
                            serial_obj.write(payload)
                            serial_obj.flush()
                            _push_serial_debug(f">>> 发送 PDU 正文及 Ctrl+Z，等待模组响应...")
                        except Exception as e:
                            _push_serial_debug(f">>> 发送失败: {e}")
                            # 失败时也在主界面提示
                            port_ui(f"❌ 发送短信失败：{e}", "normal")
                    else:
                        _push_serial_debug(">>> 发送失败: 串口未连接")
                        # 失败时也在主界面提示
                        port_ui("❌ 发送短信失败：串口未连接", "normal")
            threading.Thread(target=_send_task, daemon=True).start()

        btn_frm = ttk.Frame(frm)
        btn_frm.pack(anchor="e")
        ttk.Button(btn_frm, text="发送指令", command=_do_send_sms).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frm, text="取消", command=win.destroy).pack(side="left")

        win.update_idletasks()
        center_window(win, serial_debug_win)
        ent_phone.focus_set()
        
        win.bind("<Control-Return>", lambda e: _do_send_sms())
        win.bind("<Escape>", lambda e: win.destroy())

    # ================= 拨打电话专属功能 =================
    def _open_dial_dialog():
        win = tk.Toplevel(serial_debug_win)
        win.title("拨打电话")
        win.resizable(False, False)
        win.transient(serial_debug_win)
        win.grab_set()

        # ======== 通话状态标记 ========
        is_dialing = False
        # =============================

        frm = ttk.Frame(win, padding=15)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="请输入要拨打的手机/电话号码：").pack(anchor="w", pady=(0, 10))
        phone_var = tk.StringVar()
        ent_phone = ttk.Entry(frm, textvariable=phone_var, width=28)
        ent_phone.pack(fill="x", pady=(0, 5))

        # ================= 灰色提示文本 =================
        tk.Label(
            frm, 
            text="💡 提示：支持 '+' 国际前缀 (如 +8618888888...)。\n注意: 需确认 SIM 卡已开通语音/长途权限。",
            fg="gray", 
            justify="left",
            font=("微软雅黑", 9)
        ).pack(anchor="w", pady=(0, 15))
        # ===============================================

        def _do_dial():
            nonlocal is_dialing  
            global current_dial_num  # 引入全局变量
            
            if not enabled_var.get():
                messagebox.showwarning("提示", "请先勾选顶部的“启用原始输出旁路”", parent=win)
                return

            phone = phone_var.get().strip()
            if not phone:
                messagebox.showerror("错误", "号码不能为空", parent=win)
                return
            
            # ================= 智能拨号防呆容错 =================
            if phone.startswith("+86"):
                phone = phone[3:]
            elif phone.startswith("86") and len(phone) == 13:
                phone = phone[2:]
            # =================================================

            is_dialing = True
            current_dial_num = phone  # 把号码存到全局变量里，等接通时用
            
            port_ui(f"📞 主动呼叫：拨打号码 {phone}", "normal")
            set_status(f"📞 呼叫中：{phone}", "blue")
            
            # 发送拨号指令
            _quick_send(f"ATD{phone};")

        def _do_hangup():
            nonlocal is_dialing
            if not enabled_var.get():
                messagebox.showwarning("提示", "请先勾选顶部的“启用原始输出旁路”", parent=win)
                return
            
            # ======== 如果没有在拨号，直接拦截，不发多余指令 ========
            if not is_dialing:
                # 没拨号的情况下点挂断，直接 return 忽略
                # (如果你希望此时点击挂断能顺便把窗口关了，可以换成 win.destroy())
                return
            # =====================================================
            
            # ======== 标记为结束通话 ========
            is_dialing = False
            # ===============================
            
            port_ui("📞 已发送挂机指令 (ATH)", "normal")
            global PORT, BAUD
            set_status(f"🟢 已连接：{PORT} @ {BAUD}", "green")  # 点击后立刻强制恢复绿色状态
            
            # 发送挂机指令
            _quick_send("ATH")

        # ================= 优化：智能兜底挂机逻辑 =================
        def _on_dial_close():
            nonlocal is_dialing
            # 只有在“真正在通话（点过拨号且没点挂断）”时，关闭窗口才发 ATH 兜底
            if is_dialing and enabled_var.get():
                global PORT, BAUD
                set_status(f"🟢 已连接：{PORT} @ {BAUD}", "green")
                _quick_send("ATH")
            
            win.destroy()
        # =========================================================

        btn_frm = ttk.Frame(frm)
        btn_frm.pack(anchor="e")
        # 左侧放拨号，中间放挂断，右侧放取消
        ttk.Button(btn_frm, text="📞 拨号", command=_do_dial).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frm, text="挂断", command=_do_hangup).pack(side="left", padx=(0, 8))
        ttk.Button(btn_frm, text="取消", command=_on_dial_close).pack(side="left")

        win.update_idletasks()
        center_window(win, serial_debug_win)
        ent_phone.focus_set()
        win.bind("<Return>", lambda e: _do_dial())        
        win.bind("<Escape>", lambda e: _on_dial_close())
        win.protocol("WM_DELETE_WINDOW", _on_dial_close)
    # ============================================================

    ttk.Button(quick_scroll_frame, text="输入PIN码解锁 🔑", command=_open_input_pin_dialog).pack(fill="x", padx=6, pady=(6, 6))
    ttk.Button(quick_scroll_frame, text="输入PUK码解锁 🔐", command=_open_input_puk_dialog).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(quick_scroll_frame, text="开启PIN码锁 🔒", command=_open_enable_pin_dialog).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(quick_scroll_frame, text="关闭PIN码锁 🔓", command=_open_disable_pin_dialog).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(quick_scroll_frame, text="修改PIN码 ✏️", command=_open_modify_pin_dialog).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(quick_scroll_frame, text="修改本机号码 ☎", command=_open_modify_number_dialog).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(quick_scroll_frame, text="修改SN码 🏷️", command=_open_modify_sn_dialog).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(quick_scroll_frame, text="发送短信 ✉️", command=_open_send_sms_dialog).pack(fill="x", padx=6, pady=(0, 6))
    ttk.Button(quick_scroll_frame, text="拨打电话 📞", command=_open_dial_dialog).pack(fill="x", padx=6, pady=(0, 6))

    # ==== 3. 展开/收起控制逻辑 ====
    panel_visible = False
    def _toggle_quick_panel():
        nonlocal panel_visible
        if panel_visible:
            # grid_remove() 会完美隐藏面板，并把位置“让”出来给文本框
            quick_panel.grid_remove() 
            btn_quick.config(text="快捷命令 ▶")
            panel_visible = False
        else:
            # 重新显示在第1列
            quick_panel.grid(row=0, column=1, sticky="ns", padx=(8, 0))
            btn_quick.config(text="快捷命令 ◀")
            panel_visible = True

    # 绑定到底部的快捷按钮上
    btn_quick.config(command=_toggle_quick_panel)
    serial_debug_text.config(state="disabled")
    _update_state_label()
    serial_debug_text.tag_config("find_hit", background="yellow")
    serial_debug_text.tag_config("find_cur", background="#ff9f1a")  # 橙色，当前命中
    serial_debug_text.tag_raise("find_cur")  # 保证盖在 find_hit 上面

    find_win = None
    find_var = tk.StringVar(value="")
    last_find_index = "1.0"
    find_trace_id = None

    def _clear_find_highlight():
        try:
            serial_debug_text.tag_remove("find_hit", "1.0", "end")
            serial_debug_text.tag_remove("find_cur", "1.0", "end")
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

    def _highlight_range(term: str, start_idx: str, end_idx: str):
        if not term:
            return
        start = start_idx
        while True:
            pos = serial_debug_text.search(term, start, stopindex=end_idx, nocase=True)
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
        serial_debug_text.tag_remove("find_cur", "1.0", "end")
        serial_debug_text.tag_add("find_cur", pos, endpos)
        serial_debug_text.tag_raise("find_cur", "find_hit")
        last_find_index = endpos
        return "break"

    def _find_prev(_event=None):
        nonlocal last_find_index
        term = find_var.get().strip()
        if not term:
            return "break"

        # 用当前高亮(find_cur)的起点作为边界，避免重复命中自己
        cur_start = None
        try:
            ranges = serial_debug_text.tag_ranges("find_cur")
            cur_start = ranges[0] if ranges else serial_debug_text.index("insert")
        except Exception:
            cur_start = serial_debug_text.index("insert")

        start = serial_debug_text.index(f"{cur_start}-1c")

        pos = serial_debug_text.search(
            term, start, stopindex="1.0", nocase=True, backwards=True
        )
        if not pos:
            pos = serial_debug_text.search(
                term, "end-1c", stopindex="1.0", nocase=True, backwards=True
            )
            if not pos:
                return "break"

        endpos = f"{pos}+{len(term)}c"
        serial_debug_text.see(pos)
        serial_debug_text.mark_set("insert", endpos)

        serial_debug_text.tag_remove("find_cur", "1.0", "end")
        serial_debug_text.tag_add("find_cur", pos, endpos)
        serial_debug_text.tag_raise("find_cur", "find_hit")

        last_find_index = endpos
        return "break"

    def _open_find():
        nonlocal find_win, last_find_index, find_trace_id

        # 已打开就置顶
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

        # 只绑定一次，避免越绑越多
        if find_trace_id is None:
            find_trace_id = find_var.trace_add("write", _on_change)

        def _close_find(_event=None):
            nonlocal find_win, last_find_index, find_trace_id

            # 关闭前先清高亮
            _clear_find_highlight()
            last_find_index = "1.0"

            # 先解绑 trace，避免下面 set("") 触发回调
            if find_trace_id is not None:
                try:
                    find_var.trace_remove("write", find_trace_id)
                except Exception:
                    pass
                find_trace_id = None

            # 再清空输入
            try:
                find_var.set("")
            except Exception:
                pass

            # 关窗
            try:
                if find_win is not None and find_win.winfo_exists():
                    find_win.destroy()
            except Exception:
                pass
            find_win = None

            # 焦点回文本框
            try:
                serial_debug_text.focus_set()
            except Exception:
                pass
            return "break"

        ent.focus_set()
        ent.bind("<Return>", _find_next)
        ent.bind("<Shift-Return>", _find_prev)
        ent.bind("<Escape>", _close_find)
        find_win.bind("<Escape>", _close_find)
        find_win.protocol("WM_DELETE_WINDOW", _close_find)

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
            try:
                serial_debug_win.after(100, _append_lines)
            except Exception:
                return            
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
            insert_start = serial_debug_text.index("end-1c")
            for ln in lines:
                if kw and (kw not in ln):
                    continue
                # 确保有换行
                if not ln.endswith("\n"):
                    ln += "\n"
                serial_debug_text.insert("end", ln)
            insert_end = serial_debug_text.index("end-1c")
            term = find_var.get().strip()
            if term:
                _highlight_range(term, insert_start, insert_end)

            # 行数裁剪
            try:
                cur_lines = int(serial_debug_text.index("end-1c").split(".")[0])
                if cur_lines > MAX_LINES:
                    # 删除前面多余的行
                    del_lines = cur_lines - MAX_LINES
                    serial_debug_text.delete("1.0", f"{del_lines + 1}.0")
                    _clear_find_highlight()
                    term = find_var.get().strip()
                    if term:
                        _find_all(term)
            except Exception:
                pass

            serial_debug_text.see("end")
            serial_debug_text.config(state="disabled")

        if serial_debug_drop_count > 0:
            drop_label.config(text=f"队列满丢弃：{serial_debug_drop_count} 行")
        else:
            drop_label.config(text="")

        # 100ms 刷新一次
        try:
            serial_debug_win.after(100, _append_lines)
        except Exception:
            return        

    def _on_close():
        global SERIAL_DEBUG_ENABLED, serial_debug_drop_count, serial_debug_win, serial_debug_text
        nonlocal pause_banner_shown, find_win, find_trace_id

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

        # ===== 强制解绑滚轮，防止幽灵报错 =====
        try:
            serial_debug_win.unbind_all("<MouseWheel>")
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

        # 关闭调试窗口时，确保查找 trace 解绑（防残留）
        try:
            if find_trace_id is not None:
                find_var.trace_remove("write", find_trace_id)
                find_trace_id = None
        except Exception:
            pass

        # 如果查找窗口还开着，也关掉
        try:
            if find_win is not None and find_win.winfo_exists():
                find_win.destroy()
        except Exception:
            pass
        find_win = None

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

    serial_debug_win.deiconify() # 居中后再显示
    serial_debug_win.lift()
    serial_debug_win.focus_force()

    _append_lines()

# ================= 语音播报开关（菜单按钮） =================
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
        tmp = text.get("1.0", "end").strip()
        if not tmp:
            messagebox.showerror("错误", "播报内容不能为空", parent=win)
            return

        # 走队列生成，并在生成完成后自动回调播放（传入 play_after=True）
        generate_alert_voice(force=True, text=tmp, play_after=True)

    def do_save():
        tmp = text.get("1.0", "end").strip()
        if not tmp:
            messagebox.showerror("错误", "播报内容不能为空", parent=win)
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

# ================= 多开并发：物理端口保护锁 =================
current_port_mutex = None

def lock_port_mutex(port_name):
    global current_port_mutex
    unlock_port_mutex()
    if not port_name: return
    # 删除了 Global\ 前缀，避免需要管理员权限才能上锁
    mutex_name = f"Air724UG_PORT_{port_name}"
    current_port_mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)

def unlock_port_mutex():
    global current_port_mutex
    if current_port_mutex:
        try:
            ctypes.windll.kernel32.CloseHandle(current_port_mutex)
        except Exception:
            pass
        current_port_mutex = None

def is_port_locked_by_other(port_name):
    if not port_name: return False
    mutex_name = f"Air724UG_PORT_{port_name}"
    # 终极修复：使用 CreateMutexW 并检查错误码 183，这比 OpenMutex 更加稳定，无视权限组差异
    handle = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
    err = ctypes.windll.kernel32.GetLastError()
    if handle:
        ctypes.windll.kernel32.CloseHandle(handle)
    return err == 183  # 183 = ERROR_ALREADY_EXISTS (说明绝对被其他多开实例占用了)

# ================= 单实例：Windows Mutex 锁 =================
import ctypes

app_mutex = None

def check_single_instance():
    global app_mutex
    # 创建一个系统全局唯一的互斥量名称
    mutex_name = "Global\\Air724UG_SMS_Monitor_Mutex_V3"
    
    # 调用 Windows 内核 API 创建互斥量
    app_mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
    last_error = ctypes.windll.kernel32.GetLastError()
    
    # 183 (ERROR_ALREADY_EXISTS) 表示互斥量已被另一个实例创建
    if last_error == 183:
        # 发现已有实例运行，弹出系统原生提示框并立刻退出
        ctypes.windll.user32.MessageBoxW(0, "程序已经在运行中，请在右下角托盘查看。", "提示", 0x30)
        sys.exit(0)

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

# ================= 单实例：二次启动拦截（应放在 Tk 创建之前） =================
if not ALLOW_MULTI_INSTANCE:
    check_single_instance()

# ================= 开启 Windows 高DPI 极致清晰支持 (全世代兼容) =================
try:
    import ctypes
    # 优先尝试 Windows 8.1 / 10 / 11 的现代 API
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        # 如果失败（说明是 Windows 7 或更老），回退使用老版本 API
        ctypes.windll.user32.SetProcessDPIAware()
except Exception:
    pass

# ================= GUI =================
root = tk.Tk()
root.withdraw()
root.minsize(500, 200)

popup_var = tk.BooleanVar(value=POPUP_ENABLED)
generate_alert_voice(force=False)   # 或 force=True，程序启动时发一个“生成任务”（可选）

def resource_path(relative):
    if getattr(sys, 'frozen', False):
        return os.path.join(sys._MEIPASS, relative)
    # 脚本模式：用文件本身所在目录
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), relative)

try:
    root.iconbitmap(resource_path("icon.ico"))
except Exception as e:
    log_file_only(f"icon.ico 加载失败：{e}")

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
    log_file_only(f"弹窗图标补丁加载失败：{_e}")

root.title("短信监听系统")
root.geometry("800x520")

root.update_idletasks()
if not START_MINIMIZED:
    center_on_screen(root, 800, 520)
    root.deiconify()
else:
    # 自启：保持隐藏，托盘可“显示”
    root.withdraw()

# ================= 托盘 / 退出 / 隐藏 =================
tray_icon = None
is_exiting = False

# ================= 托盘回调：强制回主线程 =================
def show_window():
    def _do():
        try:
            root.deiconify()
            root.lift()
            root.focus_force()
        except Exception:
            pass
    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

def hide_window():
    def _do():
        try:
            root.withdraw()
        except Exception:
            pass
    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

def cleanup_and_exit():
    """真正退出：停止串口线程、关闭串口、停止托盘、销毁窗口（主线程执行更稳）"""
    def _do():
        global serial_running, serial_obj, is_exiting, tray_icon
        # 拦截提前：如果已经在退出了，就不再弹第二个确认框！
        if is_exiting:
            return
            
        # ==== 防误点确认弹窗 ====
        if not messagebox.askyesno("退出软件", "确定要完全退出软件吗？\n\n退出后将停止监听短信和来电。", parent=root):
            return
            
        is_exiting = True
        
        try:
            TK_SHUTDOWN.set()
        except Exception:
            pass

        try:
            serial_running = False
        except Exception:
            pass

        try:
            file_log_stop.set()
        except Exception:
            pass

        try:
            serial_stop_event.set()
            serial_wakeup_event.set()
        except Exception:
            pass

        safe_close_serial()

        try:
            if tray_icon:
                tray_icon.stop()
        except Exception:
            pass

        try:
            while not FILE_LOG_Q.empty():
                p, l = FILE_LOG_Q.get_nowait()
                with open(p, "a", encoding="utf-8") as f:
                    f.write(l)
        except Exception:
            pass

        try:
            root.destroy()
        except Exception:
            pass

        try:
            TTS_STOP.set()
        except Exception:
            pass

    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

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

    bottom_line = tk.Frame(win, height=1, bg="#d4d4d4")
    bottom_line.pack(side="bottom", fill="x")
    bottom_line.pack_propagate(False)

    frame = tk.Frame(win, padx=20, pady=15)
    frame.pack(side="top", fill=tk.BOTH, expand=True)

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

# ================= 线程安全 messagebox 调用（后台线程投递） =================
def ui_messagebox(kind: str, title: str, message: str):
    """
    kind: 'info' | 'warning' | 'error' | 'askyesno'
    """
    def _do():
        if kind == "info":
            return messagebox.showinfo(title, message)
        if kind == "warning":
            return messagebox.showwarning(title, message)
        if kind == "error":
            return messagebox.showerror(title, message)
        if kind == "askyesno":
            return messagebox.askyesno(title, message)
        return None

    if threading.current_thread() is threading.main_thread():
        return _do()
    else:
        ui_post(_do)
        return None

# ===== 用 grid 布局：内容区永远不会盖住状态栏 =====
root.grid_rowconfigure(0, weight=1)   # 内容区可伸缩
root.grid_rowconfigure(1, weight=0)   # 状态栏固定
root.grid_columnconfigure(0, weight=1)

# 中间内容区域
main_frame = tk.Frame(root)
main_frame.grid(row=0, column=0, sticky="nsew")

text_area = ScrolledText(main_frame, font=("微软雅黑", 10))
text_area.pack(fill=tk.BOTH, expand=True)
root.after(30, ui_pump)

# ================= 把早期提示补到窗口：从队列取出 =================
try:
    while True:
        m, t = PENDING_UI_LOGS.get_nowait()
        try:
            safe_insert_main_text(m, t)
        except Exception:
            pass
except queue.Empty:
    pass
except Exception:
    pass

# 底部状态栏
status_frame = tk.Frame(root)
status_frame.grid(row=1, column=0, sticky="ew")

status_var = tk.StringVar(value="🔍 启动中…")
status_label = tk.Label(status_frame, textvariable=status_var, anchor="w")
status_label.pack(side=tk.LEFT, padx=6)

# ================= 温度 UI 与更新函数 =================
temp_var = tk.StringVar(value="🌡️ -- ℃")
temp_label = tk.Label(status_frame, textvariable=temp_var, anchor="w", fg="#0052cc") 
temp_label.pack(side=tk.LEFT, padx=(20, 6))

def set_temperature(temp_str):
    if not tk_alive():
        return
    def _do():
        try:
            temp_var.set(f"🌡️ {temp_str} ℃")
        except Exception:
            pass
    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

# ================= 信号强度 UI 与更新函数 =================
signal_var = tk.StringVar(value="📶 -- dBm")
# 紧跟在温度标签左侧排列，用绿色显示
signal_label = tk.Label(status_frame, textvariable=signal_var, anchor="w", fg="#008000") 
signal_label.pack(side=tk.LEFT, padx=(20, 6))

def set_signal(rsrp_val):
    if not tk_alive():
        return
    def _do():
        try:
            val = int(rsrp_val) 
            if val == 255:  # CESQ 协议中 255 代表未知/无信号
                signal_var.set("📶 未知")
            else:
                # 4G RSRP 转 dBm 公式: RSRP参数 - 140
                dbm = val - 140
                signal_var.set(f"📶 {dbm} dBm")
        except Exception:
            signal_var.set("📶 -- dBm")
            
    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)
# ==========================================================

def set_status(text, color="black"):
    if not tk_alive():
        return

    def _do():
        try:
            status_var.set(text)
            status_label.config(fg=color)
        except Exception:
            pass

    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

text_area.tag_config("normal", foreground="black", font=("微软雅黑", 10))

def apply_sms_font_style():
    try:
        text_area.tag_config("sms", foreground=SMS_FONT_COLOR, font=("微软雅黑", SMS_FONT_SIZE))
    except Exception:
        pass

apply_sms_font_style()

# ================= 主窗口智能插入防爆内存 =================
def safe_insert_main_text(msg, tag="normal"):
    if ("text_area" not in globals()) or (text_area is None) or not text_area.winfo_exists():
        return False
    
    # 1. 智能滚动：判断当前滚动条是不是在最底部
    try:
        is_at_bottom = text_area.yview()[1] >= 0.98
    except Exception:
        is_at_bottom = True

    text_area.insert(tk.END, msg + "\n", tag)

    # 2. 内存保护：主屏幕永远只保留最近的 3000 行
    try:
        total_lines = int(text_area.index("end-1c").split(".")[0])
        if total_lines > 3000:
            text_area.delete("1.0", f"{total_lines - 3000 + 1}.0")
    except Exception:
        pass

    # 3. 只有当用户原本就在底部时，才自动往下滚；否则绝不强行拉底画面！
    if is_at_bottom:
        text_area.see(tk.END)
    
    return True

# ================= 统一线程安全 log：后台线程只投递，主线程执行 Tk =================
def log(msg, tag="normal"):
    def _ui_and_file():
        # --- UI ---
        try:
            if ("text_area" in globals()) and (text_area is not None) and text_area.winfo_exists():
                safe_insert_main_text(msg, tag)
            else:
                log_early(msg, tag)
                return
        except Exception:
            try:
                log_early(msg, tag)
            except Exception:
                pass
            return

        # --- 文件（COM 分日志）---
        try:
            prefix_snapshot = LOG_PREFIX
            today = datetime.now().strftime("%Y-%m-%d")
            path = os.path.join(LOG_DIR, f"sms_{prefix_snapshot}_{today}.txt")
            line = f"{datetime.now():%Y-%m-%d %H:%M:%S} {msg}\n"
            FILE_LOG_Q.put_nowait((path, line))
        except queue.Full:
            # 队列满：丢弃（避免阻塞 UI/串口线程）
            pass
        except Exception:
            pass

    # 主线程直接执行；非主线程投递到 UI 队列
    if threading.current_thread() is threading.main_thread():
        _ui_and_file()
    else:
        ui_post(_ui_and_file)

# ================= 声音 =================
_last_play_time = 0.0  # 记录上次播报时间（防抖用）

def play_alert(force: bool = False):
    global _last_play_time
    if (not force) and (not VOICE_ENABLED):
        return

    # 3秒冷却防抖。避免瞬间涌入十几条短信导致播报疯狂重叠卡顿
    now = time.monotonic()
    if (not force) and (now - _last_play_time < 3.0):
        return
    _last_play_time = now

    try:
        if os.path.exists(TTS_FILE):
            winsound.PlaySound(TTS_FILE, winsound.SND_FILENAME | winsound.SND_ASYNC)
        else:
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
    except Exception:
        winsound.MessageBeep(winsound.MB_ICONASTERISK)

def show_sms_popup(msg: str):
    if not POPUP_ENABLED:
        return

    def _do():
        try:
            messagebox.showinfo("短信提醒", msg)
            show_window()
        except Exception:
            pass

    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

# ================= 清空窗口 clear_window：永远线程安全 =================
def clear_window():
    def _do():
        try:
            if ("text_area" in globals()) and (text_area is not None) and text_area.winfo_exists():
                text_area.delete("1.0", tk.END)
        except Exception:
            pass

    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

# ================= 重启设备 =================
def send_reset_cmd():
    """发送重启指令 AT+RESET"""
    # ==== 防误点确认弹窗 ====
    # 加上 parent=root，保证即便窗口最小化也能正常弹出在最顶层
    if not messagebox.askyesno("重启硬件", "确定要重启底层通信模组吗？\n\n(设备重启期间将会短暂断开连接，随后自动重连)", parent=root):
        return
    # ==============================
    global serial_obj, serial_lock
    # 检查串口是否连接
    with serial_lock:
        if serial_obj is not None and serial_obj.is_open:
            try:
                # 发送指令，AT指令通常需要回车换行
                serial_obj.write(b"AT+RESET\r\n")
                serial_obj.flush()
                
                # 在主窗口日志中记录
                system_ui("🔄 已发送重启指令：AT+RESET", "normal")
            except Exception as e:
                system_ui(f"❌ 发送重启指令失败：{e}", "normal")
        else:
            messagebox.showwarning("提示", "串口当前未连接，无法发送指令")

# ================= 打开日志目录 =================
def open_log_dir():
    log_path = os.path.abspath(LOG_DIR)
    if os.path.exists(log_path):
        os.startfile(log_path)   # Windows 下直接打开文件夹
    else:
        ui_messagebox("warning", "提示", "日志目录不存在")

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
        # 先禁用按钮，避免重复点
        try:
            ui_post(lambda: btn_test.config(state="disabled", text="测试中…"))
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

                # 是否下载代理 OK
                download_ok = any((ok is True) and isinstance(name, str) and name.startswith("下载代理 ")
                                  for (name, ok, _info) in checks)

                if ok_bases or download_ok:
                    lines.append("")

                if ok_bases:
                    lines.append("提示：API 代理可用，检测更新将优先使用它。")

                if download_ok:
                    lines.append("提示：下载代理可用，下载链接将优先使用它。")

                messagebox.showinfo("测试结果", "\n".join(lines))

            ui_post(done)

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
    """一次自动清理 + 重新安排下一次（线程安全：确保 after 在主线程安排）"""
    def _do():
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
        try:
            if tk_alive():
                AUTO_CLEANUP_AFTER_ID = root.after(
                    AUTO_CLEANUP_INTERVAL_HOURS * 3600 * 1000,
                    _auto_log_cleanup_tick
                )
            else:
                AUTO_CLEANUP_AFTER_ID = None
        except Exception:
            AUTO_CLEANUP_AFTER_ID = None

    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

def schedule_auto_log_cleanup(restart: bool = True, first_delay_sec: int = 60):
    """
    开启/重启自动清理定时器
    - restart=True：会先取消旧定时器，避免重复跑
    - first_delay_sec：首次执行延迟（避免刚启动就占资源）
    """
    def _do():
        global AUTO_CLEANUP_AFTER_ID

        if restart and AUTO_CLEANUP_AFTER_ID is not None:
            if tk_alive():
                try:
                    root.after_cancel(AUTO_CLEANUP_AFTER_ID)
                except Exception:
                    pass
            AUTO_CLEANUP_AFTER_ID = None

        if not AUTO_LOG_CLEANUP:
            return

        # 关键：root 不可用时不要 after
        if not tk_alive():
            AUTO_CLEANUP_AFTER_ID = None
            return

        try:
            AUTO_CLEANUP_AFTER_ID = root.after(first_delay_sec * 1000, _auto_log_cleanup_tick)
        except Exception:
            AUTO_CLEANUP_AFTER_ID = None

    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

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
                ui_post(lambda: messagebox.showinfo("检测更新", f"当前已是最新版本：v{APP_VERSION}"))
                return

            asset = _pick_exe_asset(rel)
            if not asset:
                ui_post(lambda: messagebox.showwarning(
                    "检测更新",
                    f"发现新版本：{tag}\n但 Release 里没有 .zip 附件。"
                ))
                return

            raw_url = asset.get("browser_download_url") or ""
            proxy_base, _api_proxy_base = _get_update_config()
            proxy_url = (proxy_base + raw_url) if (proxy_base and raw_url.startswith("http")) else raw_url

            def ask():
                ok = messagebox.askyesno(
                    "发现新版本",
                    f"当前：v{APP_VERSION}\n最新：{tag}\n\n是否打开下载链接？（如已配置下载代理，将优先使用）"
                )
                if ok:
                    try:
                        webbrowser.open(proxy_url)
                    except Exception:
                        pass

            ui_post(ask)

        except Exception as e:
            ui_post(lambda: messagebox.showerror("检测更新失败", str(e)))

    threading.Thread(target=worker, daemon=True).start()

# ================= 每日清空 =================
def clear_text_area_for_new_day():
    clear_window()
    system_ui("📅 新的一天，窗口已清空")
    schedule_next_midnight_clear()

def schedule_next_midnight_clear():
    def _do():
        now = datetime.now()
        next_midnight = now.replace(hour=0, minute=0, second=0, microsecond=0) + timedelta(days=1)
        try:
            if tk_alive():
                ms = int((next_midnight - now).total_seconds() * 1000)
                root.after(ms, clear_text_area_for_new_day)
        except Exception:
            pass

    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

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
    global PORT # 引入当前的记忆端口
    exclude_tokens = [
        "DIAG", "NPI", "MOS", "DEBUG", "DOWNLOAD",
        "CP ", "CP_", "AP ", "AP_",  # 有些驱动会写 CP/AP
    ]

    candidates = []
    for p in list_ports.comports():
        dev = p.device
        if is_port_locked_by_other(dev):
            continue
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
            
        # 终极修复：必须严格包含 MODEM 字样。Device 7 等无用诊断口全部屏蔽！
        if "MODEM" not in desc_u:
            if dev != PORT:  # 除非它是我们记忆里真正连过的老相好
                continue
                
        score = 0
        if "MODEM" in desc_u:
            score += 100
        # 轻微偏好 Device 0（很多 LUAT 的 Modem 是 0）
        if "USB DEVICE 0" in desc_u:
            score += 10
        if dev == PORT:
            score += 1000
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

def try_rebind_manual_port(reason: str = "") -> bool:
    """
    Manual 模式下：端口失效时尝试自动重绑到新的端口（仍保持 Manual）。
    优先：LUAT Modem -> 若无则单一串口。
    成功返回 True，失败返回 False。
    """
    global PORT, MODE

    if MODE != "Manual":
        return False

    # 1) 优先找 LUAT Modem
    dev, desc = (None, None)
    try:
        dev, desc = find_luat_best_port()
    except Exception:
        dev, desc = (None, None)

    # 2) 否则：仅一个串口就用它
    if not dev:
        try:
            ports_all = list(list_ports.comports())
            if len(ports_all) == 1:
                dev = ports_all[0].device
                desc = ports_all[0].description or ""
        except Exception:
            dev = None

    # 找不到候选
    if not dev:
        return False

    # 候选就是当前端口也没意义
    if dev == PORT:
        return False

    old_port = PORT
    PORT = dev

    # 写回配置：保持 Manual，但 port 更新
    try:
        if not config.has_section("serial"):
            config["serial"] = {}
        config.set("serial", "mode", "Manual")
        config.set("serial", "port", PORT)
        config.set("serial", "baud", str(BAUD))
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)
    except Exception:
        pass

    # UI 提示（线程安全：用 system_ui / set_status，它们自己已处理线程投递）
    hint = f"🔁 手动模式端口失效，已自动重绑：{old_port} -> {PORT}"
    if desc:
        hint += f"（{desc}）"
    if reason:
        hint += f"；原因：{reason}"
    system_ui(hint, "normal")
    set_status(f"🟡 连接中：{PORT} @ {BAUD}", "orange")

    # 让串口线程立刻醒来重连
    try:
        serial_wakeup_event.set()
    except Exception:
        pass
    # 重绑成功：清零提示抑制，让下次还能正常提示
    global _last_rebind_hint_msg, _last_rebind_hint_count
    _last_rebind_hint_msg = None
    _last_rebind_hint_count = 0

    return True

def rebind_hint_ui(msg: str):
    """Manual 重绑提示：重复抑制，避免刷屏"""
    global _last_rebind_hint_msg, _last_rebind_hint_count

    try:
        if _last_rebind_hint_msg == msg:
            _last_rebind_hint_count += 1
        else:
            _last_rebind_hint_msg = msg
            _last_rebind_hint_count = 1

        if _last_rebind_hint_count < ERROR_REPEAT_LIMIT:
            system_ui(msg, "normal")
        elif _last_rebind_hint_count == ERROR_REPEAT_LIMIT:
            system_ui(msg + "（后续同类提示已忽略）", "normal")
        else:
            # 超过阈值：彻底静默
            pass
    except Exception:
        # 兜底：抑制逻辑崩了也别刷屏
        try:
            system_ui(msg, "normal")
        except Exception:
            pass

def serial_error_ui(msg: str):
    """串口异常提示：重复抑制，避免刷屏"""
    global _last_serial_error_msg, _last_serial_error_count

    try:
        if _last_serial_error_msg == msg:
            _last_serial_error_count += 1
        else:
            _last_serial_error_msg = msg
            _last_serial_error_count = 1

        if _last_serial_error_count < ERROR_REPEAT_LIMIT:
            system_ui(msg, "normal")
        elif _last_serial_error_count == ERROR_REPEAT_LIMIT:
            system_ui(msg + "（后续同类错误已忽略）", "normal")
        else:
            pass
    except Exception:
        try:
            system_ui(msg, "normal")
        except Exception:
            pass

# ================= 来电提醒互动弹窗 =================
current_call_popup = None

def close_call_popup():
    """安全关闭来电弹窗"""
    def _do():
        global current_call_popup
        if current_call_popup is not None and current_call_popup.winfo_exists():
            current_call_popup.destroy()
        current_call_popup = None
    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

def show_call_popup(caller_num):
    """显示来电弹窗（置顶显示）"""
    def _do():
        global current_call_popup
        # 如果已经有弹窗存在，就不重复弹了
        if current_call_popup is not None and current_call_popup.winfo_exists():
            return

        win = tk.Toplevel(root)
        win.title("来电提醒")
        win.minsize(300, 0)
        win.resizable(False, False)
        # 不使用 grab_set() 模态，允许用户同时操作主界面
        win.attributes("-topmost", True)  # 保持置顶显示

        frm = ttk.Frame(win, padding=15)
        frm.pack(fill="both", expand=True)

        # 提取为变量，方便接通后修改文字和颜色
        lbl_title = tk.Label(frm, text="📞 收到新来电", font=("微软雅黑", 11, "bold"), fg="#0052cc")
        lbl_title.pack(pady=(0, 8))
        tk.Label(frm, text=f"{caller_num}", font=("微软雅黑", 16, "bold"), fg="#d9534f").pack(pady=(0, 20))

        btn_frm = ttk.Frame(frm)
        btn_frm.pack(anchor="center")

        def _answer():
            global serial_obj, serial_lock, ring_timeout_target
            with serial_lock:
                if serial_obj is not None and serial_obj.is_open:
                    try:
                        serial_obj.write(b"ATA\r\n")
                        serial_obj.flush()
                    except Exception:
                        pass
            port_ui("📞 已发送接听指令 (ATA)", "normal")
            
            # 将状态切换为“通话中”，并设为 -1.0 激活免疫状态
            set_status(f"📞 通话中：{caller_num}", "blue")
            ring_timeout_target = -1.0
            
            # 🛑 核心变形魔法：不关闭弹窗，变成“通话中控制面板” 🛑
            lbl_title.config(text="📞 正在通话中...", fg="#2ecc71") # 标题变成原谅绿
            btn_answer.pack_forget() # 隐藏接听按钮
            btn_ignore.pack_forget() # 隐藏忽略按钮
            # 此时界面上就只剩下一个“挂断”按钮了！
            
            # 终极防呆：接通后，如果用户点右上角的 X 关闭窗口，也等同于挂机
            win.protocol("WM_DELETE_WINDOW", _hangup)

        def _hangup():
            global serial_obj, serial_lock
            with serial_lock:
                if serial_obj is not None and serial_obj.is_open:
                    try:
                        serial_obj.write(b"ATH\r\n")
                        serial_obj.flush()
                    except Exception:
                        pass
            port_ui("📞 已发送挂机指令 (ATH)", "normal")
            close_call_popup()

        def _ignore():
            # 仅仅关掉弹窗，不发任何指令，让模组继续响铃
            close_call_popup()

        btn_answer = ttk.Button(btn_frm, text="✅ 接听", command=_answer)
        btn_answer.pack(side="left", padx=6)
        
        btn_hangup = ttk.Button(btn_frm, text="❌ 挂断", command=_hangup)
        btn_hangup.pack(side="left", padx=6)
        
        btn_ignore = ttk.Button(btn_frm, text="忽略", command=_ignore)
        btn_ignore.pack(side="left", padx=6)
        
        win.protocol("WM_DELETE_WINDOW", close_call_popup)
        center_window(win, root)
        win.lift()
        current_call_popup = win

    if threading.current_thread() is threading.main_thread():
        _do()
    else:
        ui_post(_do)

# ================= 串口线程（自动识别 + 自动重连） =================
def read_serial():
    """
    串口读取线程（严格模式）：
    - 仅当串口行中包含 [I]-[handler_sms.smsCallback] 才认为“短信有效”
    - 命中后会收集同一条短信的多行输出，合并后再进行【关键词过滤】与弹窗/播报（避免弹窗不完整）
    - 关键词过滤规则：full_msg 只要包含 KEYWORDS 任意一项即放行；否则忽略不显示/不弹窗/不播报
    - 其它所有串口日志全部忽略
    """
    global serial_obj, serial_running, PORT, LOG_PREFIX, _last_serial_error_msg, _last_serial_error_count, ring_timeout_target, current_dial_num

    callback_prefix = "[I]-[handler_sms.smsCallback]"

    follow_lines_left = 0
    pending_parts = []
    pending_display_lines = []
    pending_deadline = 0.0
    pending_active = False

    # ======== 来电与挂机防抖记录变量 ========
    last_clip_time = 0.0
    last_clip_num = ""
    last_hangup_time = 0.0
    ring_timeout_target = 0.0
    current_dial_num = ""

    def keyword_hit(full_msg: str) -> bool:
        if not KEYWORDS:
            return True
        # 将短信内容统一转为小写，实现大小写不敏感匹配
        msg_lower = full_msg.lower()
        return any(k and (k.lower() in msg_lower) for k in KEYWORDS)
    
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
            try:
                global _last_sms_ignore_msg, _last_sms_ignore_count
                msg_text = "🚫 短信未命中关键词，已忽略"
                # 重复抑制逻辑（与串口错误抑制规则一致）
                if _last_sms_ignore_msg == msg_text:
                    _last_sms_ignore_count += 1
                else:
                    _last_sms_ignore_msg = msg_text
                    _last_sms_ignore_count = 1

                if _last_sms_ignore_count < ERROR_REPEAT_LIMIT:
                    system_ui(msg_text, "normal")
                elif _last_sms_ignore_count == ERROR_REPEAT_LIMIT:
                    system_ui("🚫 短信未命中关键词，已忽略（后续同类消息已忽略）", "normal")
                else:
                    # 超过阈值：继续抑制，不再重复显示
                    pass
            except Exception:
                try:
                    system_ui("🚫 短信未命中关键词，已忽略", "normal")
                except Exception:
                    pass

        pending_parts = []
        pending_display_lines = []
        pending_deadline = 0.0
        pending_active = False
        follow_lines_left = 0
        
    while serial_running and (not serial_stop_event.is_set()):
        try:
            target_port = PORT
            desc_str = ""

            if MODE == "Auto":
                dev, desc = find_luat_best_port()
                if dev:
                    auto_connect_ui(f"🔌 检测到 LUAT Modem 标识，自动连接：{dev}（{desc}）")
                    target_port = dev
                    desc_str = f"（{desc}）"
                if not dev:
                    ports_all = list(list_ports.comports())
                    # 过滤掉被其他多开实例占用的单一串口
                    available_ports = [p for p in ports_all if not is_port_locked_by_other(p.device)]
                    if len(available_ports) == 1:
                        single = available_ports[0]
                        target_port = single.device
                        desc = single.description or ""
                        auto_connect_ui(f"🔌 未检测到 LUAT Modem 标识，但仅发现单一串口，自动连接：{target_port}")
                        desc_str = f"（{desc}）"
                    else:
                        set_status("🔍 扫描 LUAT Modem 中…", "orange")
                        serial_wakeup_event.wait(timeout=RECONNECT_INTERVAL)
                        serial_wakeup_event.clear()
                        continue
            else:
                if not target_port:
                    set_status("🔒 手动模式：未指定串口", "red")
                    time.sleep(RECONNECT_INTERVAL)
                    continue

            set_status(f"🟡 连接中：{target_port}{desc_str} @ {BAUD}", "orange")

            if is_port_locked_by_other(target_port):
                serial_error_ui(f"⚠️ 端口冲突：{target_port} 已被本软件的其他实例占用，已绕过。")
                set_status(f"🔴 端口冲突绕过中", "red")
                time.sleep(RECONNECT_INTERVAL)
                continue

            # 创建串口也要加锁，避免与 safe_close_serial() 并发
            with serial_lock:
                # 使用目标端口尝试连接
                serial_obj = serial.Serial(target_port, BAUD, timeout=0.3, write_timeout=0.5)
                lock_port_mutex(target_port)  # 成功打开后再上锁
                
                # 自动发指令：尝试与模组进行真实通信
                serial_obj.write(b"AT+CLIP=1\r\n")
                serial_obj.flush() # 强行等待写入完成，测试端口是否真的是活的

                if MODE == "Auto":
                    PORT = target_port

            LOG_PREFIX = target_port.replace(":", "_")

            # 延迟 2 秒再输出已连接日志，与检测日志隔开（不阻塞主线程）
            def _delayed_connected_log(port, baud, delay=2):
                try:
                    time.sleep(delay)
                    # 一旦真正连接成功，立刻清零所有的防抖拦截计数器，让下次断开时还能正常提示
                    global _last_auto_connect_msg, _last_auto_connect_count, _last_serial_error_msg, _last_serial_error_count
                    _last_auto_connect_msg = None
                    _last_auto_connect_count = 0
                    _last_serial_error_msg = None
                    _last_serial_error_count = 0

                    system_ui(f"🔌 串口已连接：{port} @ {baud}")

                    def _update_status():
                        try:
                            # 终极细节：如果这2秒内刚好进来了电话，千万不要去覆盖“响铃”或“呼叫”状态！
                            current_state = status_var.get()
                            if "响铃" not in current_state and "通话" not in current_state and "呼叫" not in current_state:
                                status_var.set(f"🟢 已连接：{port} @ {baud}")
                                status_label.config(fg="green")
                        except Exception:
                            pass

                    # 这里不要在后台线程 root.after
                    def _schedule():
                        try:
                            elapsed = time.monotonic() - APP_START_MONO
                            delay_ms = int(max(0.0, START_UI_DELAY - elapsed) * 1000)
                        except Exception:
                            delay_ms = 0
                        try:
                            if delay_ms > 0:
                                root.after(delay_ms, _update_status)
                            else:
                                root.after(0, _update_status)
                        except Exception:
                            _update_status()

                    ui_post(_schedule)   # 关键：回主线程安排 after
                except Exception:
                    pass

            threading.Thread(target=_delayed_connected_log, args=(PORT, BAUD), daemon=True).start()

            while serial_running and (not serial_stop_event.is_set()):
                try:
                    # 仅在获取串口对象时加锁，绝对不要把耗时的 readline 放在锁里阻塞！
                    with serial_lock:
                        if serial_obj is None or not serial_obj.is_open:
                            raise serial.SerialException("serial_obj is None (closed)")
                        current_serial = serial_obj
                    
                    # 脱离锁的作用域后再进行 I/O 阻塞读取，极大提升 UI 并发响应速度
                    try:
                        raw = current_serial.readline()
                    except Exception as inner_e:
                        # 专门捕获由于并发 safe_close_serial 导致底层句柄被抽走引发的非标准异常
                        # 例如 ValueError: Attempting to use a port that is not open 等
                        raise serial.SerialException(f"并发读取被中断: {inner_e}")
                        
                except Exception as e:
                    # 统一拦截所有 I/O 和 并发抽锁 造成的异常，并抛给外层重连机制
                    raise e
                line = raw.decode("utf-8", "ignore").strip()
                
                # ======== 纯线程安全看门狗（检查是否未接听挂断） ========
                if ring_timeout_target > 0 and time.monotonic() > ring_timeout_target:
                    ring_timeout_target = 0.0  # 超时触发后复位
                    port_ui("📞 呼叫已取消或未接听", "normal")
                    set_status(f"🟢 已连接：{PORT} @ {BAUD}", "green")
                    last_clip_num = "" 
                    close_call_popup()  # 超时对方挂断，自动关掉弹窗
                # =====================================================

                if not line:
                    if pending_active and time.monotonic() > pending_deadline:
                        flush_pending()
                    continue
                _push_serial_debug(line)

                # ================= 解析温度数据 =================
                if "+RFTEMPERATURE:" in line:
                    try:
                        # 解析示例: "[I]-[ril.proatc] +RFTEMPERATURE: 28.84"
                        temp_val = line.split("+RFTEMPERATURE:")[1].strip()
                        set_temperature(temp_val)
                    except Exception:
                        pass

                # ================= 解析信号强度(CESQ 的 RSRP) =================
                if "+CESQ:" in line:
                    try:
                        # 解析示例: "[I]-[ril.proatc] +CESQ: 99,99,255,255,26,49"
                        # 取冒号后面的部分，再用逗号分割
                        parts = line.split("+CESQ:")[1].split(",")
                        # 确保分割出来的数据够长，取第6个参数(索引为5)
                        if len(parts) >= 6:
                            rsrp_str = parts[5].strip()
                            set_signal(rsrp_str)
                    except Exception:
                        pass

                # ================= 解析来电提醒 (RING & CLIP) =================
                if "+CLIP:" in line:
                    try:
                        m = re.search(r'\+CLIP:\s*"?(\+?\d+)"?', line)
                        if m:
                            caller_num = m.group(1)
                        else:
                            caller_num = "未知号码"

                        # ================= 智能号码归一化 (防 +86 坑) =================
                        def _norm_phone(n_str):
                            n_str = n_str.strip()
                            if n_str.startswith("+86"): return n_str[3:]
                            if n_str.startswith("86") and len(n_str) > 11: return n_str[2:]
                            return n_str
                            
                        norm_caller = _norm_phone(caller_num)
                        # ==============================================================

                        # 核心防御：黑白名单智能拦截逻辑
                        blocked = False
                        block_reason = ""
                        
                        if CALL_FILTER_MODE == "Whitelist":
                            # 默认拦截，除非在白名单找到匹配
                            blocked = True
                            block_reason = "不在白名单"
                            for w_num in CALL_WHITELIST:
                                if _norm_phone(w_num) == norm_caller:
                                    blocked = False
                                    block_reason = ""
                                    break
                                    
                        elif CALL_FILTER_MODE == "Blacklist":
                            # 默认放行，一旦在黑名单找到匹配则立刻拦截
                            for b_num in CALL_BLACKLIST:
                                if _norm_phone(b_num) == norm_caller:
                                    blocked = True
                                    block_reason = "命中黑名单"
                                    break

                        now = time.monotonic()

                        if blocked:
                            if (caller_num != last_clip_num) or (now - last_clip_time > 4.0):
                                port_ui(f"🚫 防骚扰拦截：拒接 {caller_num} ({block_reason})", "warning")
                                last_clip_num = caller_num
                                last_clip_time = now
                            
                            # 最快速度夺取串口控制权，强行静默挂机
                            with serial_lock:
                                if serial_obj is not None and serial_obj.is_open:
                                    try:
                                        serial_obj.write(b"ATH\r\n")
                                        serial_obj.flush()
                                    except Exception:
                                        pass
                            continue  # ⛔ 强行中断！阻止后续的弹窗、响铃和看门狗刷新！
                        
                        # (被拦截后就不会执行到这里，正常来电才会继续触发弹窗)                        
                        if (caller_num != last_clip_num) or (now - last_clip_time > 4.0):
                            port_ui(f"📞 收到来电：来自 {caller_num}", "normal")
                            set_status(f"🔔 响铃中：{caller_num}", "blue")
                            last_clip_num = caller_num
                            last_clip_time = now
                            show_call_popup(caller_num)

                        # 每次收到响铃，把超时目标时间推迟到 12 秒后
                        ring_timeout_target = time.monotonic() + 12.0
                    except Exception:
                        pass

                elif "RING" == line or line.endswith("RING"):
                    # 只要收到 RING，说明电话还在响，刷新看门狗时间
                    # 核心防呆：如果已经接听（-1.0 免疫态），则直接无视缓存里滞后的 RING！
                    if ring_timeout_target != -1.0:
                        ring_timeout_target = time.monotonic() + 12.0

                # ================= 解析真实挂断 (NO CARRIER / BUSY / NO ANSWER) =================
                if "NO CARRIER" in line or "BUSY" in line or "NO ANSWER" in line:
                    ring_timeout_target = 0.0  
                    current_dial_num = ""      # 电话挂断时，清空主动呼出记录
                    
                    now = time.monotonic()
                    if (now - last_hangup_time > 3.0):
                        port_ui("📞 语音通话已结束", "normal")
                        set_status(f"🟢 已连接：{PORT} @ {BAUD}", "green")
                        last_hangup_time = now
                        last_clip_num = ""
                        close_call_popup()

                # ================= 解析对方接听状态 (+CIEV: "CALL",1) =================
                if line == "CONNECT" or ("+CIEV:" in line and '"CALL",1' in line.replace(" ", "").upper()):
                    if current_dial_num:
                        port_ui(f"📞 对方已接听：{current_dial_num}", "normal")
                        set_status(f"📞 通话中：{current_dial_num}", "blue")
                        ring_timeout_target = 0.0  # 确保强制关闭看门狗倒计时

                # ================= 解析运营商(COPS) 并直接在窗口提示 =================
                if "+COPS:" in line and '"' in line:
                    try:
                        # 解析示例: [I]-[ril.proatc] +COPS: 0,2,"46011",7
                        plmn = line.split('+COPS:')[1].split('"')[1]
                        
                        # 运营商匹配字典
                        carrier_map = {
                            "46000": "中国移动", "46002": "中国移动", "46007": "中国移动","46008": "中国移动",
                            "46001": "中国联通", "46006": "中国联通", "46009": "中国联通",
                            "46003": "中国电信", "46005": "中国电信", "46011": "中国电信",
                            "46013": "中国电信(物联卡)",
                            "46004": "中国移动(物联卡)",
                            "46015": "中国广电"
                        }
                        
                        c_name = carrier_map.get(plmn, "未知运营商")
                        # 追加一条醒目的提示打印到当前的串口调试窗口中
                        _push_serial_debug(f">>> 识别到网络运营商：{c_name} ({plmn})")
                    except Exception:
                        pass

                # ================= 解析SIM卡状态(CPIN) 并直接在窗口提示 =================
                if "+CPIN:" in line:
                    try:
                        # 解析示例: [I]-[ril.proatc] +CPIN: READY
                        cpin_status = line.split("+CPIN:")[1].strip()
                        
                        # 常见 PIN 状态字典
                        cpin_map = {
                            "READY": "PIN码锁未开启 (SIM卡正常)",
                            "SIM PIN": "等待输入PIN码 (卡已被锁)",
                            "SIM PUK": "等待输入PUK码 (需PUK解锁)",
                            "NOT INSERTED": "未检测到SIM卡",
                            "NOT READY": "SIM卡未准备好"
                        }
                        
                        status_cn = cpin_map.get(cpin_status, f"未知状态 ({cpin_status})")
                        # 追加一条醒目的提示打印到当前的串口调试窗口中
                        _push_serial_debug(f">>> 识别到SIM卡状态：{status_cn}")
                    except Exception:
                        pass

                # ================= 解析网络附着状态(CGATT) 并直接在窗口提示 =================
                if "+CGATT:" in line:
                    try:
                        # 解析示例: [I]-[ril.proatc] +CGATT: 1
                        cgatt_status = line.split("+CGATT:")[1].strip()
                        
                        if cgatt_status == "1":
                            status_cn = "已附着数据网络 (通道已打通，可以上网)"
                        elif cgatt_status == "0":
                            status_cn = "未附着数据网络 (无法连接互联网)"
                        else:
                            status_cn = f"未知状态 ({cgatt_status})"
                            
                        # 追加一条醒目的提示打印到当前的串口调试窗口中
                        _push_serial_debug(f">>> 识别到网络附着状态：{status_cn}")
                    except Exception:
                        pass

                # ================= 解析射频/飞行模式状态(CFUN) 并直接在窗口提示 =================
                if "+CFUN:" in line:
                    try:
                        # 解析示例: [I]-[ril.proatc] +CFUN: 1
                        cfun_status = line.split("+CFUN:")[1].strip()
                        
                        # 常见 CFUN 状态字典
                        cfun_map = {
                            "0": "飞行模式 (射频关闭)",
                            "1": "正常模式 (射频开启)",
                        }
                        
                        status_cn = cfun_map.get(cfun_status, f"未知状态 ({cfun_status})")
                        # 追加一条醒目的提示打印到当前的串口调试窗口中
                        _push_serial_debug(f">>> 识别到射频/飞行模式状态：{status_cn}")
                    except Exception:
                        pass

                # ================= 解析基站定位数据(LTE) =================
                if "+EEMLTESVC:" in line:
                    try:
                        parts = line.split("+EEMLTESVC:")[1].split(",")
                        if len(parts) >= 10:
                            raw_mcc = int(parts[0].strip())
                            raw_mnc = int(parts[1].strip())
                            tac = parts[3].strip()
                            ci = parts[9].strip()
                            
                            # LUAT 模组底层特色修正：把被强行转成十进制的转回十六进制
                            mcc = f"{raw_mcc:x}" 
                            mnc = f"{raw_mnc:02x}"
                            
                            # 分行追加到当前的串口调试窗口中
                            _push_serial_debug(">>> 解析到基站定位数据：")
                            _push_serial_debug(f"    MCC (国家代码) : {mcc}")
                            _push_serial_debug(f"    MNC (网络代码) : {mnc}")
                            _push_serial_debug(f"    LAC/TAC (区域) : {tac}")
                            _push_serial_debug(f"    CI  (小区ID)   : {ci}")
                    except Exception:
                        pass

                # ================= 短信接收核心逻辑 =================
                if callback_prefix in line:
                    msg = line.split(callback_prefix, 1)[1].strip()
                    if msg:
                        pending_parts = [msg]
                        pending_display_lines = ["📩 收到短信：", msg]
                        pending_active = True
                        pending_deadline = time.monotonic() + 1.0
                        follow_lines_left = 40
                    else:
                        pending_parts = []
                        pending_display_lines = []
                        pending_active = False
                        follow_lines_left = 0
                    continue

                # ================= 多行短信断行收集逻辑 =================
                if pending_active:
                    # 如果新的一行以系统日志标志开头，说明上一条短信已经彻底打印完了
                    # (收紧 [ 的判断，防止短信内容中以 [温馨提示] 开头的行被误杀)
                    if line.startswith("[I]-") or line.startswith("[W]-") or line.startswith("[E]-") or \
                       line.startswith(">>>") or line.startswith("AT+") or line.startswith("+"):
                        flush_pending()
                        # 注意：这里千万不要写 continue，必须让这行新日志继续往下走去解析
                    else:
                        # 如果没有系统前缀，说明它肯定是短信的断行正文（不管是纯英文、数字还是链接）
                        if follow_lines_left > 0:
                            pending_parts.append(line)
                            pending_display_lines.append(line)
                            pending_deadline = time.monotonic() + 0.4
                            follow_lines_left -= 1
                            if follow_lines_left <= 0:
                                flush_pending()
                        else:
                            flush_pending()
                            
                        # 短信碎片收集完毕，直接 continue 读取下一行串口数据
                        continue
                # ===============================================================
                
        except Exception as e:
            LOG_PREFIX = "system"

            # 物理断开时，强行重置所有通话状态，并瞬间销毁弹窗防卡死
            ring_timeout_target = 0.0
            current_dial_num = ""
            close_call_popup() 
            # =======================================================

            # 1) 打印原始异常
            err_msg = str(e)
            serial_error_ui(f"⚠️ 串口异常：{err_msg}")

            set_status(f"🔴 断开/失败：{PORT}（自动重连中…）", "red")
            set_temperature("--")  # 断开时重置温度显示
            set_signal("--")       # 断开时重置信号显示
            safe_close_serial()

            # 2) Manual：先尝试重绑（在 wait 之前做！）
            if MODE == "Manual":
                s = (repr(e) + " " + str(e)).lower()

                # Windows/pyserial 常见拔插错误关键词
                is_port_gone = any(x in s for x in [
                    "could not open port",
                    "file not found",
                    "no such file",
                    "the system cannot find the file specified",
                    "cleartcommerror",                 # ClearCommError failed
                    "clearcommerror",
                    "semaphore timeout",               # Semaphore timeout period has expired
                    "timeout period has expired",
                    "device does not recognize",       # The device does not recognize the command
                    "access is denied",                # 被系统/占用/权限
                    "winerror 2",                      # 系统找不到文件
                    "winerror 5",                      # 拒绝访问
                    "winerror 31",                     # 设备未正常工作
                    "winerror 1167",                   # device not connected
                    "device not functioning",
                    "设备", "不存在", "找不到", "系统找不到指定的文件"
                ])

                if is_port_gone:
                    rebind_hint_ui("🧠 Manual：检测到疑似拔插/端口变化，尝试自动重绑…")

                    ok = try_rebind_manual_port("端口号变化或设备插拔")
                    if ok:
                        # 重绑成功：立刻回到 while 外层，重新 open 新 PORT
                        continue
                else:
                    # system_ui(f"DEBUG: is_port_gone=False, e={repr(e)}")
                    pass

            # 3) Auto：维持原逻辑（清 PORT 让它重新扫描）
            if MODE == "Auto":
                pass

            # 4) 等待下一轮（只有在没有 continue 的情况下才会走到这里）
            serial_wakeup_event.wait(timeout=RECONNECT_INTERVAL)
            serial_wakeup_event.clear()            

    safe_close_serial()

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

        # 线程安全关闭串口，触发 read_serial() 进入异常->重连
        safe_close_serial()

        try:
            serial_wakeup_event.set()
        except Exception:
            pass

        system_ui(f"⚙️ 串口设置已更新：mode={MODE} port={PORT or '(Auto)'} baud={BAUD}")
        win.destroy()

    win = tk.Toplevel(root)
    win.withdraw()
    win.title("串口设置")
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
        text="💡 提示：\nAuto 自动优先识别 LUAT Modem\nManual 手动锁定所选 COM",
        fg="gray",
        justify="left",
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

    bottom_line = tk.Frame(win, height=1, bg="#d4d4d4")
    bottom_line.pack(side="bottom", fill="x")
    bottom_line.pack_propagate(False)

    frame = tk.Frame(win, padx=14, pady=12)
    frame.pack(side="top", fill=tk.BOTH, expand=True)

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
            # 使用 json 序列化保存，完美支持包含 "|"、引号等任何特殊字符
            config.set("ui", "keywords", json.dumps(KEYWORDS, ensure_ascii=False))
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
        system_ui(f"💬 关键词 增加：{v}")

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
        system_ui(f"💬 关键词 删除：{old}")

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
        system_ui(f"💬 关键词 修改：{old} -> {v}")

    win = tk.Toplevel(root)
    win.withdraw()
    win.title("短信关键词设置")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    bottom_line = tk.Frame(win, height=1, bg="#d4d4d4")
    bottom_line.pack(side="bottom", fill="x")
    bottom_line.pack_propagate(False)
    frame = tk.Frame(win, padx=12, pady=10)
    frame.pack(side="top", fill=tk.BOTH, expand=True)

    tk.Label(frame, text="关键词列表：").grid(row=0, column=0, sticky="w")

    listbox = tk.Listbox(frame, height=8, width=20)
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
        text="💡 提示：关键词为空时，不做过滤全部短信都会显示。",
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

# ================= 防骚扰黑白名单设置窗口 =================
def open_call_filter_setting():
    win = tk.Toplevel(root)
    win.withdraw()
    win.title("来电防骚扰设置")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    frm = tk.Frame(win, padx=15, pady=15)
    frm.pack(fill=tk.BOTH, expand=True)

    # --- 模式选择区 ---
    mode_frm = tk.LabelFrame(frm, text="过滤模式 (即时生效)", padx=10, pady=8)
    mode_frm.pack(fill="x", pady=(0, 15))

    mode_var = tk.StringVar(value=CALL_FILTER_MODE)

    def on_mode_change():
        global CALL_FILTER_MODE
        CALL_FILTER_MODE = mode_var.get()
        if "ui" not in config: config["ui"] = {}
        config.set("ui", "call_filter_mode", CALL_FILTER_MODE)
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                config.write(f)
        except Exception:
            pass
        
        mode_cn = {"Disabled": "关闭过滤", "Whitelist": "白名单模式", "Blacklist": "黑名单模式"}[CALL_FILTER_MODE]
        system_ui(f"📞 防骚扰模式已切换为：{mode_cn}")

    tk.Radiobutton(mode_frm, text="关闭过滤 (允许所有)", variable=mode_var, value="Disabled", command=on_mode_change).pack(side="left", padx=5)
    tk.Radiobutton(mode_frm, text="白名单 (仅限名单内)", variable=mode_var, value="Whitelist", command=on_mode_change).pack(side="left", padx=5)
    tk.Radiobutton(mode_frm, text="黑名单 (拦截名单内)", variable=mode_var, value="Blacklist", command=on_mode_change).pack(side="left", padx=5)

    # --- 名单管理区 (使用 Notebook 选项卡) ---
    notebook = ttk.Notebook(frm)
    notebook.pack(fill="both", expand=True)

    tab_white = tk.Frame(notebook, padx=10, pady=10)
    tab_black = tk.Frame(notebook, padx=10, pady=10)
    notebook.add(tab_white, text="白名单管理")
    notebook.add(tab_black, text="黑名单管理")

    # 辅助函数：构建左右布局的列表管理
    def build_list_tab(parent_tab, target_list, config_key, list_name):
        listbox = tk.Listbox(parent_tab, height=10)
        listbox.pack(side="left", fill="both", expand=True, padx=(0, 10))

        right_frm = tk.Frame(parent_tab)
        right_frm.pack(side="right", fill="y")

        tk.Label(right_frm, text="手机/电话号码：").pack(anchor="w")

        # ================= 提示语 =================
        tk.Label(
            right_frm, 
            text="💡 提示：需包含国际前缀\n请与模块日志上报完全一致\n(如: +8618888888...)。",
            fg="gray", 
            justify="left",
            font=("微软雅黑", 8)
        ).pack(anchor="w", pady=(0, 5))
        # ==========================================

        entry_var = tk.StringVar()
        entry = ttk.Entry(right_frm, textvariable=entry_var, width=18)
        entry.pack(anchor="w", pady=(0, 10))

        def refresh():
            listbox.delete(0, tk.END)
            for num in target_list:
                listbox.insert(tk.END, num)

        def save():
            if "ui" not in config: config["ui"] = {}
            config.set("ui", config_key, json.dumps(target_list, ensure_ascii=False))
            try:
                with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                    config.write(f)
            except Exception:
                pass

        def on_select(_e=None):
            sel = listbox.curselection()
            if sel: entry_var.set(target_list[sel[0]])

        listbox.bind("<<ListboxSelect>>", on_select)

        def add_num():
            val = entry_var.get().strip()
            if not val: return
            if val in target_list:
                messagebox.showwarning("提示", "该号码已在名单中", parent=win)
                return
            target_list.append(val)
            save()
            refresh()
            entry_var.set("")
            system_ui(f"📞 {list_name} 增加：{val}")

        def del_num():
            sel = listbox.curselection()
            if not sel: return
            idx = sel[0]
            val = target_list.pop(idx)
            save()
            refresh()
            entry_var.set("")
            system_ui(f"📞 {list_name} 删除：{val}")

        def edit_num():
            sel = listbox.curselection()
            if not sel:
                messagebox.showwarning("提示", "请先在左侧列表中选择要修改的号码", parent=win)
                return
            idx = sel[0]
            old_val = target_list[idx]
            new_val = entry_var.get().strip()

            if not new_val:
                messagebox.showerror("错误", "号码不能为空", parent=win)
                return
            
            # 如果没改动，直接返回
            if new_val == old_val:
                return

            # 防止修改后的号码和列表里其他的冲突
            if new_val in target_list:
                messagebox.showwarning("提示", "该号码已在名单中", parent=win)
                return

            target_list[idx] = new_val
            save()
            refresh()
            # 修改完后让它继续保持选中状态
            listbox.selection_set(idx)
            system_ui(f"📞 {list_name} 修改：{old_val} -> {new_val}")

        ttk.Button(right_frm, text="增加", command=add_num).pack(fill="x", pady=4)
        ttk.Button(right_frm, text="删除", command=del_num).pack(fill="x", pady=4)
        ttk.Button(right_frm, text="修改", command=edit_num).pack(fill="x", pady=4)
        refresh()

    build_list_tab(tab_white, CALL_WHITELIST, "call_whitelist", "白名单")
    build_list_tab(tab_black, CALL_BLACKLIST, "call_blacklist", "黑名单")

    # --- 底部 ---
    tk.Button(frm, text="关闭窗口", width=12, command=win.destroy).pack(anchor="e", pady=(12, 0))

    center_window(win, root)
    win.deiconify()
    win.lift()
    win.focus_force()
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

    if VOICE_ENABLED:
        msg = "🔊 语音播报：已开启"
    else:
        msg = "🔇 语音播报：已关闭"

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

    if ALLOW_MULTI_INSTANCE:
        msg = "✅️ 程序多开：已开启"
    else:
        msg = "❌ 程序多开：已关闭"

    system_ui(msg, "normal")

def toggle_autostart():
    set_autostart(autostart_var.get())

def toggle_popup():
    global POPUP_ENABLED
    POPUP_ENABLED = bool(popup_var.get())

    try:
        if not config.has_section("ui"):
            config["ui"] = {}
        config.set("ui", "popup_enabled", "1" if POPUP_ENABLED else "0")
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            config.write(f)
    except Exception:
        pass

    if POPUP_ENABLED:
        msg = "✅️ 短信弹窗：已开启"
    else:
        msg = "❌ 短信弹窗：已关闭"
    system_ui(msg, "normal")

# ================= 重启软件 =================
def restart_software():
    global is_exiting, serial_running, tray_icon, app_mutex
    # 拦截提前：防止用户疯狂连点导致弹出多个重启确认框
    if is_exiting:
        return
        
    if not messagebox.askyesno("重启软件", "确定要重启软件吗？", parent=root):
        return
    
    is_exiting = True
    system_ui("🔄 正在重启软件...", "normal")
    
    # 1. 标记退出状态并释放底层串口资源
    serial_running = False
    safe_close_serial()

    try:
        if app_mutex:
            import ctypes
            ctypes.windll.kernel32.ReleaseMutex(app_mutex)
            ctypes.windll.kernel32.CloseHandle(app_mutex)
    except Exception:
        pass

    # 2. 销毁右下角托盘图标 (防止留下幽灵图标)
    try:
        if tray_icon:
            tray_icon.stop()
    except Exception:
        pass

    # 3. 拉起一个新的自己 (过滤掉自启参数，防止重启后再次隐藏窗口)
    try:
        new_args = [arg for arg in sys.argv if arg != AUTOSTART_FLAG]
        if getattr(sys, 'frozen', False):
            # 如果是打包好的 exe
            subprocess.Popen(new_args)
        else:
            # 如果是脚本模式
            subprocess.Popen([sys.executable] + new_args)
    except Exception:
        pass

    try:
        while not FILE_LOG_Q.empty():
            p, l = FILE_LOG_Q.get_nowait()
            with open(p, "a", encoding="utf-8") as f:
                f.write(l)
    except Exception:
        pass

    # 4. 瞬间结束当前老进程 (让 OS 立即回收所有内存和 socket 端口)
    os._exit(0)

# ================= 菜单（一级串口设置） =================
menu_bar = tk.Menu(root)

file_menu = tk.Menu(menu_bar, tearoff=0)
file_menu.add_command(label="清空窗口", command=clear_window)
file_menu.add_command(label="打开日志", command=open_log_dir)
file_menu.add_separator()
file_menu.add_command(label="重启软件", command=restart_software)
file_menu.add_command(label="重启硬件", command=send_reset_cmd)
file_menu.add_command(label="退出软件", command=cleanup_and_exit)
menu_bar.add_cascade(label="文件", menu=file_menu)

# 串口设置
menu_bar.add_command(label="串口设置", command=open_serial_setting)

# 短信关键词设置
menu_bar.add_command(label="关键词设置", command=open_keywords_setting)

# 防骚扰设置
menu_bar.add_command(label="防骚扰设置", command=open_call_filter_setting)

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

settings_menu.add_checkbutton(
    label="短信弹窗",
    variable=popup_var,
    command=toggle_popup
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

# 帮助 菜单
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
