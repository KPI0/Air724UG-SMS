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
#  作者：ChatGPT、Gemini、Codex、KPI0
#  GitHub：https://github.com/KPI0/Air724UG-SMS
# ================================================================

# ---- 标准库 ----
import asyncio
import configparser
import os
import re
import sys
import threading
import time
import winsound
import webbrowser
import queue

# ---- 第三方库 ----
import serial
from serial.tools import list_ports

# ---- tkinter ----
import tkinter as tk
from tkinter import messagebox

from sms_core.app_launch import (
    maybe_run_restart_helper_mode,
)
from sms_core.app_paths import get_app_dir, resource_path
from sms_core.app_shutdown import flush_log_queue, safe_set_events
from sms_core.file_log_runtime import start_file_log_worker
from sms_core.cloud_protocol import (
    auth_status_from_ack as _cloud_auth_status_from_ack,
    normalize_imei as _normalize_imei,
    parse_sms_callback_head as _parse_cloud_sms_callback_head,
)
from sms_core.cloud_runtime import (
    CloudControlSettings,
    read_cloud_control_settings,
)
from sms_core.cloud_namespace_bindings import install_cloud_namespace_bindings
from sms_core.cloud_auth import auth_match_result as _cloud_auth_match_result
from sms_core.cloud_payloads import (
    build_register_payload as _cloud_build_register_payload,
    build_serial_log_payload as _cloud_build_serial_log_payload,
    build_sms_event_payload as _cloud_build_sms_event_payload,
    build_status_payload as _cloud_build_status_payload,
    build_unregister_payload as _cloud_build_unregister_payload,
    identity_payload as _cloud_identity_payload_core,
)
from sms_core.cloud_serial_log_runtime import (
    CloudSerialLogDrainState,
)
from sms_core.cloud_security import (
    check_replay_window as _cloud_check_replay_window_core,
    safe_preview as _cloud_safe_preview,
)
from sms_core.config_schema import (
    DEFAULT_CLOUD_CONTROL_CONFIG,
    DEFAULT_SERIAL_CONFIG,
    DEFAULT_UI_CONFIG,
    DEFAULT_UPDATE_CONFIG,
    DEFAULT_VOICE_TEXT,
    THIRD_PUSH_DEFAULTS,
)
from sms_core.config_runtime import (
    initialize_config_runtime,
    read_startup_config_values,
)
from sms_core.serial_namespace_bindings import install_serial_namespace_bindings
from sms_core.serial_sender import send_command_with_result_async, write_serial_command_result
from sms_core.status_text import format_connected_status
from sms_core.windows_shortcuts import (
    create_desktop_shortcut,
    create_startup_shortcut,
    is_autostart_enabled,
    remove_startup_shortcut,
)
from sms_core.windows_runtime import (
    is_port_locked_by_other,
    release_mutex_handle,
    request_dpi_awareness,
)
from sms_ui.serial_debug_namespace_bindings import install_serial_debug_namespace_bindings
from sms_ui.settings_dialogs import (
    open_sms_font_dialog as _ui_open_sms_font_dialog,
)
from sms_ui.settings_namespace_bindings import install_settings_namespace_bindings
from sms_ui.third_push_namespace_bindings import install_third_push_namespace_bindings
from sms_ui.maintenance_runtime import (
    AutoLogCleanupState,
)
from sms_ui.maintenance_namespace_bindings import install_maintenance_namespace_bindings
from sms_ui.app_infrastructure_namespace_bindings import install_app_infrastructure_namespace_bindings
from sms_ui.app_lifecycle_namespace_bindings import install_app_lifecycle_namespace_bindings
from sms_ui.app_ui_namespace_bindings import install_app_ui_namespace_bindings
from sms_ui.app_menu_runtime import build_main_menu_runtime
from sms_ui.main_window_layout import build_main_window_layout_runtime
from sms_ui.window_icon_runtime import install_window_icon_runtime
from sms_ui.audio_namespace_bindings import install_audio_namespace_bindings
from sms_ui.repeat_notice_runtime import (
    ConsecutiveRepeatNotice,
    TimedRepeatNotice,
)
from sms_ui.ui_log_namespace_runtime import (
    flush_pending_ui_logs_namespace_runtime,
)
from sms_ui.ui_log_namespace_bindings import install_ui_log_namespace_bindings
from sms_ui.utility_dialogs import (
    open_about_dialog as _ui_open_about_dialog,
    open_desktop_shortcut_dialog as _ui_open_desktop_shortcut_dialog,
    open_voice_text_dialog as _ui_open_voice_text_dialog,
)
from sms_ui.window_utils import sync_and_focus_existing_window

# ---- 预编译正则 ----
IMEI_REGEX = re.compile(r"\b(\d{14,17})\b")

# ================= 配置 =================
APP_DIR = get_app_dir()
CONFIG_FILE = os.path.join(APP_DIR, "config.ini")  # 软件配置文件
LOG_DIR = os.path.join(APP_DIR, "sms_logs") # 短信日志文件夹
TTS_DIR = os.path.join(APP_DIR, "tts") # 语音播报文件夹
TTS_FILE = os.path.join(TTS_DIR, "alert.wav")
APP_WINDOW_TITLE = "短信监听系统"
RECONNECT_INTERVAL = 2  # 秒
APP_VERSION = "3.6.6"  # 软件版本号
GITHUB_OWNER = "KPI0"
GITHUB_REPO = "Air724UG-SMS"

# 启动参数：开机自启时是否默认最小化到托盘
AUTOSTART_FLAG = "--autostart"
RESTART_HELPER_FLAG = "--restart-helper"
START_MINIMIZED = AUTOSTART_FLAG in sys.argv

# 程序启动时间（用于延迟启动时的 UI 日志显示）
APP_START_MONO = time.monotonic()
# 启动后延迟多少秒才在窗口显示日志（但仍会立即写文件）
START_UI_DELAY = 2.0

# ================= 语音播报开关 =================
VOICE_ENABLED = True
os.makedirs(LOG_DIR, exist_ok=True)
os.makedirs(TTS_DIR, exist_ok=True)

# ================= 读取配置 =================
config = configparser.ConfigParser(interpolation=None)
CONFIG_LOCK = threading.RLock()
install_app_infrastructure_namespace_bindings(globals())

initialize_config_runtime(
    config=config,
    config_file=CONFIG_FILE,
    defaults_by_section={
        "serial": DEFAULT_SERIAL_CONFIG,
        "ui": DEFAULT_UI_CONFIG,
        "update": DEFAULT_UPDATE_CONFIG,
        "cloud_control": DEFAULT_CLOUD_CONTROL_CONFIG,
        "third_push": THIRD_PUSH_DEFAULTS,
    },
    save_config=safe_save_config,
)

startup_config = read_startup_config_values(config, default_voice_text=DEFAULT_VOICE_TEXT)
VOICE_TEXT = startup_config.voice_text
POPUP_ENABLED = startup_config.popup_enabled
AUTO_LOG_CLEANUP = startup_config.auto_log_cleanup
LOG_RETENTION_DAYS = startup_config.log_retention_days
ALLOW_MULTI_INSTANCE = startup_config.allow_multi_instance
LOG_UNMATCHED_SMS = startup_config.log_unmatched_sms
VOICE_ENABLED = startup_config.voice_enabled
SMS_FONT_SIZE = startup_config.sms_font_size
SMS_FONT_COLOR = startup_config.sms_font_color
KEYWORDS = startup_config.keywords
CALL_FILTER_MODE = startup_config.call_filter_mode
CALL_WHITELIST = startup_config.call_whitelist
CALL_BLACKLIST = startup_config.call_blacklist
PORT = startup_config.port
BAUD = startup_config.baud
MODE = startup_config.mode

# ===== 云端控制（WebSocket）依赖 =====
try:
    import websockets
except Exception:
    websockets = None

# ===== 云端控制（WebSocket）配置 =====
# IMEI 只来自当前串口设备的 AT+CGSN 响应，不写入配置，避免多开实例共用 config.ini 时串号。
CLOUD_DEVICE_IMEI = ""

def apply_cloud_control_settings(settings: CloudControlSettings):
    global CLOUD_CONTROL_ENABLED, CLOUD_WS_URL, CLOUD_DEVICE_SECRET
    global CLOUD_WS_RECONNECT_INTERVAL, CLOUD_AUTO_UPLOAD
    CLOUD_CONTROL_ENABLED = settings.enabled
    CLOUD_WS_URL = settings.url
    CLOUD_DEVICE_SECRET = settings.device_secret
    CLOUD_WS_RECONNECT_INTERVAL = settings.reconnect_interval
    CLOUD_AUTO_UPLOAD = settings.auto_upload

def refresh_cloud_control_settings_from_config():
    """重新从 config.ini 读取云端控制配置，避免窗口复用时显示旧状态。"""
    try:
        config.read(CONFIG_FILE, encoding="utf-8")
    except Exception:
        pass
    apply_cloud_control_settings(read_cloud_control_settings(config))

apply_cloud_control_settings(read_cloud_control_settings(config))

# ===== 三方推送配置 =====
install_third_push_namespace_bindings(globals())
ensure_third_push_config(save=True)
refresh_third_push_settings_from_config()

# ================= 串口控制 =================
serial_obj = None
serial_running = True
ring_timeout_target = 0.0 
current_dial_num = ""      

serial_lock = threading.Lock()

ERROR_REPEAT_LIMIT = 4  # 1~3 次显示详细错误；第 4 次显示“后续忽略”
SERIAL_ERROR_REPEAT_RESET_SECONDS = 60.0
_auto_connect_notice = ConsecutiveRepeatNotice(
    limit=ERROR_REPEAT_LIMIT,
    suffix="（后续同类提示已忽略）",
)
_serial_error_notice = TimedRepeatNotice(
    limit=ERROR_REPEAT_LIMIT,
    reset_seconds=SERIAL_ERROR_REPEAT_RESET_SECONDS,
    suffix="（后续同类错误已忽略）",
)
_rebind_hint_notice = ConsecutiveRepeatNotice(
    limit=ERROR_REPEAT_LIMIT,
    suffix="（后续同类提示已忽略）",
)

# ================= 短信忽略重复抑制 =================
# 用于抑制连续重复显示相同的“短信未命中关键词，已忽略”提示（避免日志刷屏）
_sms_ignore_repeat_state = {}

CLOUD_LOG_REPEAT_LIMIT = 4  # 1~3 次输出详细日志；第 4 次输出“后续忽略”
CLOUD_MAIN_REPEAT_RESET_SECONDS = 60.0
_cloud_main_notice = TimedRepeatNotice(
    limit=CLOUD_LOG_REPEAT_LIMIT,
    reset_seconds=CLOUD_MAIN_REPEAT_RESET_SECONDS,
    suffix="（后续同类消息已忽略）",
)
_cloud_file_notice = TimedRepeatNotice(
    limit=CLOUD_LOG_REPEAT_LIMIT,
    reset_seconds=CLOUD_MAIN_REPEAT_RESET_SECONDS,
    suffix="（后续同类消息已忽略）",
)

# ================= 全局变量 =================
PENDING_UI_LOGS = queue.Queue(maxsize=20000)  # 用于 text_area 未创建前缓存要显示到窗口的提示
LOG_PREFIX = "system"
AUTO_CLEANUP_INTERVAL_HOURS = 24 # 自动清理频率：24小时一次
AUTO_LOG_CLEANUP_STATE = AutoLogCleanupState()
SERIAL_DEBUG_ENABLED = False
serial_debug_queue = queue.Queue(maxsize=5000)  # 防止无限涨
serial_debug_win = None
serial_debug_text = None
serial_debug_drop_count = 0
cloud_control_win = None
third_push_win = None
cloud_ws_loop = None
cloud_ws_conn = None
cloud_ws_thread = None
cloud_ws_lock = threading.Lock()
cloud_stop_event = threading.Event()
cloud_restart_seq = 0
CLOUD_SERIAL_LOG_Q = queue.Queue(maxsize=1000)
CLOUD_SERIAL_LOG_DRAIN_BATCH = 100
CLOUD_REPLAY_WINDOW_SECONDS = 60
CLOUD_REPLAY_CACHE_MAX = 512
cloud_replay_seen = {}
CLOUD_SERIAL_LOG_DRAIN_STATE = CloudSerialLogDrainState()
cloud_connected = False
cloud_device_authorized = False
cloud_imei_verified = False
cloud_imei_query_deadline = 0.0
serial_stop_event = threading.Event()
serial_wakeup_event = threading.Event()
TTS_LOCK = threading.Lock()
TTS_REQ_Q = queue.Queue(maxsize=50)
TTS_STOP = threading.Event()
TTS_THREAD = None
THIRD_PUSH_Q = queue.Queue(maxsize=200)
third_push_stop = threading.Event()

# ================= Tk 线程安全：UI 任务队列（所有 Tk 操作只能在主线程） =================
UI_TASK_QUEUE = queue.Queue(maxsize=10000)
FILE_LOG_Q = queue.Queue(maxsize=50000)
file_log_stop = threading.Event()

start_file_log_worker(log_queue=FILE_LOG_Q, stop_event=file_log_stop)

# ================= root 是否可用（避免退出过程中 after 抛异常） =================
TK_SHUTDOWN = threading.Event()
install_ui_log_namespace_bindings(globals())

install_audio_namespace_bindings(globals())
install_settings_namespace_bindings(globals())
install_app_lifecycle_namespace_bindings(globals())
install_app_ui_namespace_bindings(globals())
install_serial_debug_namespace_bindings(globals())
install_maintenance_namespace_bindings(globals())
ensure_tts_worker()

# ================= 多开并发：物理端口保护锁 =================
current_port_mutex = None
app_mutex = None

# ================= 单实例：二次启动拦截（应放在 Tk 创建之前） =================
maybe_run_restart_helper_mode(RESTART_HELPER_FLAG)

if not ALLOW_MULTI_INSTANCE:
    check_single_instance()

# ================= 开启 Windows 高DPI 极致清晰支持 (全世代兼容) =================
request_dpi_awareness()

# ================= GUI =================
root = tk.Tk()
root.withdraw()
root.minsize(500, 200)

popup_var = tk.BooleanVar(value=POPUP_ENABLED)
generate_alert_voice(force=False)   # 或 force=True，程序启动时发一个“生成任务”（可选）

install_window_icon_runtime(
    root,
    tk,
    messagebox,
    icon_path=resource_path("icon.ico"),
    path_exists=os.path.exists,
    log_error=log_file_only,
)

root.title(APP_WINDOW_TITLE)
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

def on_close():
    """点右上角×：隐藏到托盘，不退出"""
    hide_window()

root.protocol("WM_DELETE_WINDOW", on_close)
root.bind("<Escape>", lambda _e: on_close())

threading.Thread(target=create_tray, daemon=True).start()

_layout_refs = build_main_window_layout_runtime(root, tk, cloud_enabled=CLOUD_CONTROL_ENABLED)
main_frame = _layout_refs["main_frame"]
text_area = _layout_refs["text_area"]
status_frame = _layout_refs["status_frame"]
status_var = _layout_refs["status_var"]
status_label = _layout_refs["status_label"]
temp_var = _layout_refs["temp_var"]
temp_label = _layout_refs["temp_label"]
signal_var = _layout_refs["signal_var"]
signal_label = _layout_refs["signal_label"]
cloud_var = _layout_refs["cloud_var"]
cloud_label = _layout_refs["cloud_label"]
root.after(30, ui_pump)

text_area.tag_config("normal", foreground="black", font=("微软雅黑", 10))

apply_sms_font_style()

# ================= 把早期提示补到窗口：从队列取出 =================
flush_pending_ui_logs_namespace_runtime(globals())

_last_play_time = 0.0  # 记录上次播报时间（防抖用）

# ================= 三方推送 =================
threading.Thread(target=_third_push_worker, daemon=True).start()

# ================= 云端控制（WebSocket） =================
install_cloud_namespace_bindings(globals())

current_call_popup = None
install_serial_namespace_bindings(globals())

_menu_state = build_main_menu_runtime(
    root,
    tk,
    is_autostart_enabled=is_autostart_enabled,
    allow_multi_instance=ALLOW_MULTI_INSTANCE,
    popup_var=popup_var,
    commands={
        "clear_window": clear_window,
        "open_log_dir": open_log_dir,
        "restart_software": restart_software,
        "send_reset_cmd": send_reset_cmd,
        "cleanup_and_exit": cleanup_and_exit,
        "open_serial_setting": open_serial_setting,
        "open_keywords_setting": open_keywords_setting,
        "open_call_filter_setting": open_call_filter_setting,
        "toggle_voice_broadcast": toggle_voice_broadcast,
        "toggle_autostart": toggle_autostart,
        "toggle_multi_instance": toggle_multi_instance,
        "toggle_popup": toggle_popup,
        "open_log_cleanup_dialog": open_log_cleanup_dialog,
        "open_update_proxy_dialog": open_update_proxy_dialog,
        "open_desktop_shortcut_dialog": open_desktop_shortcut_dialog,
        "open_voice_text_dialog": open_voice_text_dialog,
        "open_sms_font_dialog": open_sms_font_dialog,
        "open_cloud_control_window": open_cloud_control_window,
        "open_third_push_window": open_third_push_window,
        "open_serial_debug_window": open_serial_debug_window,
        "show_about": show_about,
        "check_update_and_prompt": check_update_and_prompt,
    },
)
menu_bar = _menu_state["menu_bar"]
voice_menu_index = _menu_state["voice_menu_index"]
autostart_var = _menu_state["autostart_var"]
multi_instance_var = _menu_state["multi_instance_var"]
update_voice_menu_label()

# ================= 启动 =================
schedule_next_midnight_clear()

if MODE == "Auto":
    set_status("🔍 自动模式：扫描 LUAT Modem 中…", "orange")
else:
    set_status(f"✍️ 手动模式：{PORT or '未指定'}", "orange")

if CLOUD_CONTROL_ENABLED:
    start_cloud_control()

threading.Thread(target=read_serial, daemon=True).start()
# 启动后自动清理定时器（默认60秒后首次运行）
schedule_auto_log_cleanup(restart=True, first_delay_sec=60)

root.mainloop()


