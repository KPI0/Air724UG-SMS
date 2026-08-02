# ---- 标准库 ----
import asyncio
import configparser
import os
import re
import sys
import threading
import time
import webbrowser
import queue

# ---- 跨平台兼容性 ----
try:
    import winsound
except ImportError:
    # Linux/Mac: 创建winsound stub
    class winsound:
        @staticmethod
        def Beep(frequency, duration):
            pass
        @staticmethod
        def MessageBeep(*args, **kwargs):
            pass
        @staticmethod
        def PlaySound(*args, **kwargs):
            pass
        MB_OK = 0
        MB_ICONASTERISK = 0
        SND_FILENAME = 0
        SND_ASYNC = 0

# ---- 第三方库 ----
import serial
from serial.tools import list_ports

# ---- tkinter (跨平台兼容) ----
try:
    import tkinter as tk
    from tkinter import messagebox
except ImportError:
    # Headless环境: 创建tkinter stub
    class tk:
        class Tk:
            def __init__(self):
                raise RuntimeError("Tkinter不可用: headless环境")
        class StringVar:
            def __init__(self, *args, **kwargs):
                self.value = ""
            def get(self):
                return self.value
            def set(self, v):
                self.value = v
    class messagebox:
        @staticmethod
        def showinfo(*args, **kwargs):
            pass
        @staticmethod
        def showerror(*args, **kwargs):
            pass

from sms_core.app_launch import (
    maybe_run_restart_helper_mode,
)
from sms_core.autostart_instances import (
    AUTOSTART_CHILD_FLAG,
    get_autostart_state_path,
    launch_autostart_companions,
    register_autostart_instance,
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
from sms_core.cloud_command_security import read_cloud_command_permissions
from sms_app.cloud_namespace_bindings import install_cloud_namespace_bindings
from sms_app.version import APP_VERSION as CLIENT_VERSION
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
from sms_core.call_session import IncomingCallSessionTracker
from sms_core.config_schema import (
    DEFAULT_CLOUD_CONTROL_CONFIG,
    DEFAULT_SERIAL_CONFIG,
    DEFAULT_UI_CONFIG,
    DEFAULT_UPDATE_CONFIG,
    DEFAULT_VOICE_TEXT,
    THIRD_PUSH_DEFAULTS,
)
from sms_core.config_runtime import (
    ConfigInitializationError,
    initialize_config_runtime,
    read_startup_config_values,
)
from sms_app.serial_namespace_bindings import install_serial_namespace_bindings
from sms_core.serial_sender import (
    DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR,
    DEFAULT_SERIAL_COMMAND_THREAD_REGISTRY,
    DEFAULT_SMS_PDU_SEND_COORDINATOR,
    DEFAULT_SMS_SEND_THREAD_REGISTRY,
    send_command_with_result_async,
    write_serial_command_result,
)
from sms_core.status_text import format_connected_status
from sms_core.threading_runtime import (
    SingleFlightTaskState,
    WorkerThreadRegistry,
    start_daemon_thread,
)
from sms_core.tts_runtime import instance_tts_file_path
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
from sms_ui.app_instance_runtime import (
    claim_instance_number_app_runtime,
    format_instance_window_title,
    is_instance_number_active_app_runtime,
)
from sms_ui.settings_dialogs import (
    open_sms_font_dialog as _ui_open_sms_font_dialog,
)
from sms_ui.security_settings_dialog import (
    open_security_settings_dialog as _ui_open_security_settings_dialog,
)
from sms_ui.settings_namespace_bindings import install_settings_namespace_bindings
from sms_ui.config_sync_namespace_bindings import install_config_sync_namespace_bindings
from sms_ui.config_sync_runtime import ConfigFileWatchState
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
from sms_ui.ui_log_runtime import system_log_prefix_runtime
from sms_ui.ui_log_namespace_bindings import install_ui_log_namespace_bindings
from sms_ui.utility_dialogs import (
    open_about_dialog as _ui_open_about_dialog,
    open_desktop_shortcut_dialog as _ui_open_desktop_shortcut_dialog,
    open_voice_text_dialog as _ui_open_voice_text_dialog,
)
from sms_ui.window_utils import sync_and_focus_existing_window



def _initialize_paths_and_constants():
    global APP_DIR, CONFIG_FILE, LOG_DIR, TTS_DIR, TTS_FILE, APP_WINDOW_TITLE
    global APP_DISPLAY_TITLE, SERIAL_DEBUG_WINDOW_TITLE, APP_INSTANCE_NUMBER
    global RECONNECT_INTERVAL, APP_VERSION, GITHUB_OWNER, GITHUB_REPO
    global AUTOSTART_FLAG, RESTART_HELPER_FLAG, START_MINIMIZED
    global AUTOSTART_STATE_FILE, AUTOSTART_IS_LAUNCH, AUTOSTART_IS_LEADER
    global APP_START_MONO, START_UI_DELAY, VOICE_ENABLED, IMEI_REGEX

    IMEI_REGEX = re.compile(r"\b(\d{14,17})\b")
    APP_DIR = get_app_dir()
    CONFIG_FILE = os.path.join(APP_DIR, "config.ini")
    LOG_DIR = os.path.join(APP_DIR, "sms_logs")
    TTS_DIR = os.path.join(APP_DIR, "tts")
    TTS_FILE = os.path.join(TTS_DIR, "alert.wav")
    APP_WINDOW_TITLE = "短信监听系统"
    APP_DISPLAY_TITLE = APP_WINDOW_TITLE
    SERIAL_DEBUG_WINDOW_TITLE = "串口调试"
    APP_INSTANCE_NUMBER = 1
    RECONNECT_INTERVAL = 2
    APP_VERSION = CLIENT_VERSION
    GITHUB_OWNER = "KPI0"
    GITHUB_REPO = "Air724UG-SMS"
    AUTOSTART_FLAG = "--autostart"
    RESTART_HELPER_FLAG = "--restart-helper"
    AUTOSTART_IS_LAUNCH = AUTOSTART_FLAG in sys.argv
    AUTOSTART_IS_LEADER = AUTOSTART_IS_LAUNCH and AUTOSTART_CHILD_FLAG not in sys.argv
    AUTOSTART_STATE_FILE = get_autostart_state_path(APP_DIR)
    START_MINIMIZED = AUTOSTART_IS_LAUNCH
    APP_START_MONO = time.monotonic()
    START_UI_DELAY = 2.0
    VOICE_ENABLED = True
    os.makedirs(LOG_DIR, exist_ok=True)
    os.makedirs(TTS_DIR, exist_ok=True)


def _initialize_config():
    global config, CONFIG_LOCK, startup_config, VOICE_TEXT, POPUP_ENABLED
    global CALL_POPUP_ENABLED
    global AUTO_LOG_CLEANUP, LOG_RETENTION_DAYS, ALLOW_MULTI_INSTANCE
    global LOG_UNMATCHED_SMS, VOICE_ENABLED, SMS_FONT_SIZE, SMS_FONT_COLOR
    global KEYWORDS, CALL_FILTER_MODE, CALL_WHITELIST, CALL_BLACKLIST
    global PORT, BAUD, MODE

    config = configparser.ConfigParser(interpolation=None)
    CONFIG_LOCK = threading.RLock()
    install_app_infrastructure_namespace_bindings(globals())
    defaults_by_section = {
        "serial": DEFAULT_SERIAL_CONFIG,
        "ui": DEFAULT_UI_CONFIG,
        "update": DEFAULT_UPDATE_CONFIG,
        "cloud_control": DEFAULT_CLOUD_CONTROL_CONFIG,
        "third_push": THIRD_PUSH_DEFAULTS,
    }
    initialize_config_runtime(
        config=config,
        config_file=CONFIG_FILE,
        defaults_by_section=defaults_by_section,
        save_config=lambda: safe_save_config(defaults_by_section=defaults_by_section),
    )

    startup_config = read_startup_config_values(config, default_voice_text=DEFAULT_VOICE_TEXT)
    VOICE_TEXT = startup_config.voice_text
    POPUP_ENABLED = startup_config.popup_enabled
    CALL_POPUP_ENABLED = startup_config.call_popup_enabled
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


def apply_cloud_control_settings(settings: CloudControlSettings):
    global CLOUD_CONTROL_ENABLED, CLOUD_WS_URL, CLOUD_DEVICE_SECRET
    global CLOUD_WS_RECONNECT_INTERVAL, CLOUD_AUTO_UPLOAD

    CLOUD_CONTROL_ENABLED = settings.enabled
    CLOUD_WS_URL = settings.url
    CLOUD_DEVICE_SECRET = settings.device_secret
    CLOUD_WS_RECONNECT_INTERVAL = settings.reconnect_interval
    CLOUD_AUTO_UPLOAD = settings.auto_upload


def _initialize_cloud_settings():
    global CLOUD_DEVICE_IMEI, CLOUD_SENSITIVE_COMMAND_PERMISSIONS, websockets

    try:
        import websockets
    except Exception:
        websockets = None
    CLOUD_DEVICE_IMEI = ""
    CLOUD_SENSITIVE_COMMAND_PERMISSIONS = read_cloud_command_permissions(config)
    apply_cloud_control_settings(read_cloud_control_settings(config))


def _initialize_third_push_settings():
    install_third_push_namespace_bindings(globals())
    ensure_third_push_config(save=True)
    refresh_third_push_settings_from_config()


def _initialize_serial_state():
    global serial_obj, serial_running, ring_timeout_target, current_dial_num, serial_lock
    global serial_stop_event, serial_wakeup_event, serial_connection_generation

    serial_obj = None
    serial_running = True
    ring_timeout_target = 0.0
    current_dial_num = ""
    serial_lock = threading.Lock()
    serial_stop_event = threading.Event()
    serial_wakeup_event = threading.Event()
    serial_connection_generation = 0


def _initialize_notice_state():
    global ERROR_REPEAT_LIMIT, SERIAL_ERROR_REPEAT_RESET_SECONDS
    global _auto_connect_notice, _serial_error_notice, _rebind_hint_notice
    global _sms_ignore_repeat_state, CLOUD_LOG_REPEAT_LIMIT, CLOUD_MAIN_REPEAT_RESET_SECONDS
    global _cloud_main_notice, _cloud_file_notice

    ERROR_REPEAT_LIMIT = 4
    SERIAL_ERROR_REPEAT_RESET_SECONDS = 60.0
    _auto_connect_notice = ConsecutiveRepeatNotice(limit=ERROR_REPEAT_LIMIT, suffix="（后续同类提示已忽略）")
    _serial_error_notice = TimedRepeatNotice(
        limit=ERROR_REPEAT_LIMIT,
        reset_seconds=SERIAL_ERROR_REPEAT_RESET_SECONDS,
        suffix="（后续同类错误已忽略）",
    )
    _rebind_hint_notice = ConsecutiveRepeatNotice(limit=ERROR_REPEAT_LIMIT, suffix="（后续同类提示已忽略）")
    _sms_ignore_repeat_state = {}
    CLOUD_LOG_REPEAT_LIMIT = 4
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


def _initialize_ui_state():
    global PENDING_UI_LOGS, LOG_PREFIX, LOCAL_NUMBER
    global AUTO_CLEANUP_INTERVAL_HOURS, AUTO_LOG_CLEANUP_STATE, SERIAL_DEBUG_ENABLED
    global CONFIG_FILE_WATCH_STATE
    global serial_debug_queue, serial_debug_win, serial_debug_text, serial_debug_drop_count
    global cloud_control_win, third_push_win, sms_popup_win

    PENDING_UI_LOGS = queue.Queue(maxsize=20000)
    LOG_PREFIX = "system"
    LOCAL_NUMBER = ""
    AUTO_CLEANUP_INTERVAL_HOURS = 24
    AUTO_LOG_CLEANUP_STATE = AutoLogCleanupState()
    CONFIG_FILE_WATCH_STATE = ConfigFileWatchState()
    SERIAL_DEBUG_ENABLED = False
    serial_debug_queue = queue.Queue(maxsize=5000)
    serial_debug_win = None
    serial_debug_text = None
    serial_debug_drop_count = 0
    cloud_control_win = None
    third_push_win = None
    sms_popup_win = None


def _initialize_cloud_runtime_state():
    global cloud_ws_loop, cloud_ws_conn, cloud_ws_thread
    global cloud_ws_lock, cloud_stop_event, cloud_restart_seq, CLOUD_SERIAL_LOG_Q
    global CLOUD_SERIAL_LOG_DRAIN_BATCH, CLOUD_REPLAY_WINDOW_SECONDS, CLOUD_REPLAY_CACHE_MAX
    global cloud_replay_seen, CLOUD_SERIAL_LOG_DRAIN_STATE, cloud_connected
    global cloud_device_authorized, cloud_imei_verified, cloud_imei_query_deadline

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


def _initialize_worker_state():
    global TTS_LOCK, TTS_REQ_Q, TTS_STOP, TTS_THREAD, THIRD_PUSH_Q, third_push_stop
    global third_push_thread, serial_thread
    global UI_TASK_QUEUE, FILE_LOG_Q, file_log_stop, file_log_thread
    global TK_SHUTDOWN, current_port_mutex, app_mutex
    global instance_number_mutex, SMS_SEND_COORDINATOR, SMS_SEND_THREAD_REGISTRY
    global autostart_spawn_thread, AUTOSTART_INSTANCE_REGISTERED
    global AUTOSTART_DESIRED_INSTANCE_COUNT
    global SERIAL_COMMAND_RESPONSE_COORDINATOR
    global SERIAL_COMMAND_THREAD_REGISTRY, UPDATE_THREAD_REGISTRY, UPDATE_CHECK_TASK_STATE

    TTS_LOCK = threading.Lock()
    TTS_REQ_Q = queue.Queue(maxsize=50)
    TTS_STOP = threading.Event()
    TTS_THREAD = None
    THIRD_PUSH_Q = queue.Queue(maxsize=200)
    third_push_stop = threading.Event()
    third_push_thread = None
    serial_thread = None
    UI_TASK_QUEUE = queue.Queue(maxsize=10000)
    FILE_LOG_Q = queue.Queue(maxsize=50000)
    file_log_stop = threading.Event()
    file_log_thread = start_file_log_worker(log_queue=FILE_LOG_Q, stop_event=file_log_stop)
    TK_SHUTDOWN = threading.Event()
    current_port_mutex = None
    app_mutex = None
    instance_number_mutex = None
    autostart_spawn_thread = None
    AUTOSTART_INSTANCE_REGISTERED = False
    AUTOSTART_DESIRED_INSTANCE_COUNT = 1
    SMS_SEND_COORDINATOR = DEFAULT_SMS_PDU_SEND_COORDINATOR
    SERIAL_COMMAND_RESPONSE_COORDINATOR = DEFAULT_AT_COMMAND_RESPONSE_COORDINATOR
    SMS_SEND_THREAD_REGISTRY = DEFAULT_SMS_SEND_THREAD_REGISTRY
    SERIAL_COMMAND_THREAD_REGISTRY = DEFAULT_SERIAL_COMMAND_THREAD_REGISTRY
    UPDATE_THREAD_REGISTRY = WorkerThreadRegistry()
    UPDATE_CHECK_TASK_STATE = SingleFlightTaskState()


def _initialize_runtime_state():
    _initialize_serial_state()
    _initialize_notice_state()
    _initialize_ui_state()
    _initialize_cloud_runtime_state()
    _initialize_worker_state()


def _install_runtime_bindings():
    install_ui_log_namespace_bindings(globals())
    install_audio_namespace_bindings(globals())
    install_settings_namespace_bindings(globals())
    install_config_sync_namespace_bindings(globals())
    install_app_lifecycle_namespace_bindings(globals())
    install_app_ui_namespace_bindings(globals())
    install_serial_debug_namespace_bindings(globals())
    install_maintenance_namespace_bindings(globals())
    ensure_tts_worker()


def _run_startup_guards():
    global APP_INSTANCE_NUMBER, instance_number_mutex
    global APP_DISPLAY_TITLE, SERIAL_DEBUG_WINDOW_TITLE
    global LOG_PREFIX, TTS_FILE
    global AUTOSTART_DESIRED_INSTANCE_COUNT, AUTOSTART_INSTANCE_REGISTERED

    maybe_run_restart_helper_mode(RESTART_HELPER_FLAG)
    check_single_instance()
    APP_INSTANCE_NUMBER, instance_number_mutex = claim_instance_number_app_runtime(
        app_dir=APP_DIR,
        log_error=log_file_only,
    )
    registration = register_autostart_instance(
        app_dir=APP_DIR,
        state_path=AUTOSTART_STATE_FILE,
        instance_number=APP_INSTANCE_NUMBER,
        preserve_desired=AUTOSTART_IS_LAUNCH,
        allow_multi_instance=ALLOW_MULTI_INSTANCE,
        is_instance_active=lambda number: is_instance_number_active_app_runtime(
            app_dir=APP_DIR,
            instance_number=number,
        ),
        log_error=log_file_only,
    )
    AUTOSTART_DESIRED_INSTANCE_COUNT = registration.desired_count
    AUTOSTART_INSTANCE_REGISTERED = registration.registered
    APP_DISPLAY_TITLE = format_instance_window_title(APP_WINDOW_TITLE, APP_INSTANCE_NUMBER)
    SERIAL_DEBUG_WINDOW_TITLE = format_instance_window_title("串口调试", APP_INSTANCE_NUMBER)
    if LOG_PREFIX == "system":
        LOG_PREFIX = system_log_prefix_runtime(APP_INSTANCE_NUMBER)
    TTS_FILE = instance_tts_file_path(TTS_DIR, APP_INSTANCE_NUMBER)
    request_dpi_awareness()


def _create_root_window():
    global root, popup_var, call_popup_var, tray_icon, tray_thread, is_exiting, on_close

    root = tk.Tk()
    root.withdraw()
    root.minsize(520, 200)
    popup_var = tk.BooleanVar(value=POPUP_ENABLED)
    call_popup_var = tk.BooleanVar(value=CALL_POPUP_ENABLED)
    generate_alert_voice(force=False)
    install_window_icon_runtime(
        root,
        tk,
        messagebox,
        icon_path=resource_path("icon.ico"),
        path_exists=os.path.exists,
        log_error=log_file_only,
    )
    root.title(APP_DISPLAY_TITLE)
    root.geometry("800x520")
    root.update_idletasks()
    if not START_MINIMIZED:
        center_on_screen(root, 800, 520)
        root.deiconify()
    else:
        root.withdraw()

    tray_icon = None
    is_exiting = False

    def on_close():
        """点右上角×：隐藏到托盘，不退出"""
        hide_window()

    root.protocol("WM_DELETE_WINDOW", on_close)
    root.bind("<Escape>", lambda _e: on_close())
    tray_thread = start_daemon_thread("tray", create_tray, log_error=log_file_only)


def _build_main_layout():
    global _layout_refs, main_frame, text_area, status_frame, status_var, status_label
    global temp_var, temp_label, signal_var, signal_label, cloud_var, cloud_label

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
    flush_pending_ui_logs_namespace_runtime(globals())


def _install_cloud_and_serial_bindings():
    global _last_play_time, current_call_popup, current_missed_call_popup
    global INCOMING_CALL_SESSION, third_push_thread

    _last_play_time = 0.0
    third_push_thread = start_daemon_thread(
        "third_push_worker",
        _third_push_worker,
        log_error=log_file_only,
    )
    install_cloud_namespace_bindings(globals())
    current_call_popup = None
    current_missed_call_popup = None
    INCOMING_CALL_SESSION = IncomingCallSessionTracker()
    install_serial_namespace_bindings(globals())


def _build_main_menu():
    global _menu_state, menu_bar, voice_menu_index, autostart_var, multi_instance_var

    _menu_state = build_main_menu_runtime(
        root,
        tk,
        is_autostart_enabled=is_autostart_enabled,
        allow_multi_instance=ALLOW_MULTI_INSTANCE,
        popup_var=popup_var,
        call_popup_var=call_popup_var,
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
            "toggle_call_popup": toggle_call_popup,
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


def _start_services():
    global serial_thread, autostart_spawn_thread

    schedule_next_midnight_clear()
    start_config_file_watch()
    if MODE == "Auto":
        set_status("🔍 自动模式：扫描 LUAT Modem 中…", "orange")
    else:
        set_status(f"✍️ 手动模式：{PORT or '未指定'}", "orange")
    if CLOUD_CONTROL_ENABLED:
        start_cloud_control()
    serial_thread = start_daemon_thread("serial_reader", read_serial, log_error=log_file_only)
    if (
        AUTOSTART_IS_LEADER
        and APP_INSTANCE_NUMBER == 1
        and ALLOW_MULTI_INSTANCE
        and AUTOSTART_DESIRED_INSTANCE_COUNT > 1
    ):
        autostart_spawn_thread = start_daemon_thread(
            "autostart_instance_launcher",
            lambda: launch_autostart_companions(
                desired_count=AUTOSTART_DESIRED_INSTANCE_COUNT,
                allow_multi_instance=ALLOW_MULTI_INSTANCE,
                is_leader=True,
                autostart_flag=AUTOSTART_FLAG,
                child_flag=AUTOSTART_CHILD_FLAG,
                wait_before_launch=TK_SHUTDOWN.wait,
                log_error=log_file_only,
            ),
            log_error=log_file_only,
        )
    schedule_auto_log_cleanup(restart=True, first_delay_sec=60)


def main():
    """Start the SMS listener application."""
    try:
        _initialize_paths_and_constants()
    except OSError as exc:
        try:
            messagebox.showerror(
                "运行目录初始化失败",
                "无法创建程序运行所需的日志或语音目录。\n\n"
                f"{exc}\n\n"
                "请将软件放到有写入权限的目录，或调整当前目录权限后重新启动。",
            )
        except Exception:
            pass
        return False
    try:
        _initialize_config()
    except ConfigInitializationError as exc:
        try:
            messagebox.showerror(
                "配置初始化失败",
                f"{exc}\n\n请检查程序目录是否有写入权限、配置文件是否被占用，然后重新启动软件。",
            )
        except Exception:
            pass
        return False
    _initialize_cloud_settings()
    _initialize_third_push_settings()
    _initialize_runtime_state()
    _install_runtime_bindings()
    _run_startup_guards()
    _create_root_window()
    _build_main_layout()
    _install_cloud_and_serial_bindings()
    _build_main_menu()
    _start_services()
    root.mainloop()
    return True
