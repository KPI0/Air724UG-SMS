import os
import re

from sms_core.app_launch import get_launch_target_and_args
from sms_core.app_paths import get_startup_dir, get_startup_lnk


INVALID_SHORTCUT_CHARS = re.compile(r'[\\/:*?"<>|]')


def sanitize_shortcut_name(shortcut_name: str, default: str = "sms"):
    name = INVALID_SHORTCUT_CHARS.sub("_", (shortcut_name or "").strip())
    if not name:
        name = default
    if not name.lower().endswith(".lnk"):
        name += ".lnk"
    return name


def create_windows_shortcut(
    lnk_path: str,
    target: str,
    arguments: str = "",
    working_dir: str = "",
    window_style: int = 1,
):
    """Create a Windows .lnk through COM without generating temporary VBS."""
    try:
        import pythoncom
        from win32com.client import Dispatch
    except Exception as e:
        raise RuntimeError(f"缺少创建快捷方式所需组件 pywin32：{e}") from e

    os.makedirs(os.path.dirname(lnk_path), exist_ok=True)

    com_initialized = False
    shell = None
    shortcut = None
    try:
        pythoncom.CoInitialize()
        com_initialized = True

        shell = Dispatch("WScript.Shell")
        shortcut = shell.CreateShortcut(lnk_path)
        shortcut.TargetPath = target
        shortcut.WorkingDirectory = working_dir or os.path.dirname(target)
        shortcut.WindowStyle = int(window_style)
        if arguments:
            shortcut.Arguments = arguments
        shortcut.Save()
    except Exception as e:
        raise RuntimeError(f"创建快捷方式失败：{e}") from e
    finally:
        shortcut = None
        shell = None
        if com_initialized:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass

    if not os.path.exists(lnk_path):
        raise RuntimeError("创建快捷方式失败：未生成 .lnk 文件")


def get_desktop_dir():
    # Prefer the real redirected desktop path, including OneDrive redirection.
    try:
        import winreg

        with winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Explorer\User Shell Folders",
        ) as key:
            desktop = winreg.QueryValueEx(key, "Desktop")[0]
        desktop = os.path.expandvars(desktop)
        if desktop and os.path.isdir(desktop):
            return desktop
    except Exception:
        pass

    return os.path.join(os.path.expanduser("~"), "Desktop")


def create_desktop_shortcut(shortcut_name: str):
    desktop = get_desktop_dir()
    os.makedirs(desktop, exist_ok=True)

    lnk_path = os.path.join(desktop, sanitize_shortcut_name(shortcut_name))
    target, args, workdir = get_launch_target_and_args()
    arg_line = f'"{args}"' if args else ""

    create_windows_shortcut(
        lnk_path=lnk_path,
        target=target,
        arguments=arg_line,
        working_dir=workdir,
        window_style=1,
    )


def create_startup_shortcut(autostart_flag: str):
    startup_dir = get_startup_dir()
    os.makedirs(startup_dir, exist_ok=True)

    lnk_path = get_startup_lnk()
    target, args, workdir = get_launch_target_and_args()
    arg_line = f'"{args}" {autostart_flag}' if args else autostart_flag

    create_windows_shortcut(
        lnk_path=lnk_path,
        target=target,
        arguments=arg_line,
        working_dir=workdir,
        window_style=1,
    )


def remove_startup_shortcut():
    lnk = get_startup_lnk()
    if os.path.exists(lnk):
        os.remove(lnk)


def is_autostart_enabled():
    return os.path.exists(get_startup_lnk())
