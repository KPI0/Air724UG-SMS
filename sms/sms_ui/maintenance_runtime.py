import threading
from dataclasses import dataclass

from sms_core.updates import (
    format_proxy_test_result,
    normalize_proxy_base,
    test_update_proxy_connectivity,
)
from sms_ui.update_proxy_dialog import open_update_proxy_dialog
from sms_ui.utility_dialogs import open_log_cleanup_dialog


def log_cleanup_status(days, interval_hours):
    return f"✅ 已启用自动日志清理：保留 {days} 天（每 {interval_hours} 小时执行一次）"


@dataclass
class AutoLogCleanupState:
    after_id: object = None


def normalized_retention_days(days):
    return days if days >= 0 else 0


def open_log_dir_runtime(log_dir, *, path_abspath, path_exists, open_path, show_message):
    log_path = path_abspath(log_dir)
    if not path_exists(log_path):
        show_message("warning", "提示", "日志目录不存在")
        return "missing"

    try:
        open_path(log_path)
        return "opened"
    except Exception as exc:
        show_message("error", "打开日志失败", f"无法打开日志目录：\n{exc}")
        return "error"


def post_to_ui_thread(callback, *, is_main_thread, ui_post):
    if is_main_thread():
        callback()
    else:
        ui_post(callback)


def run_auto_log_cleanup_tick_runtime(
    *,
    state,
    is_enabled,
    retention_days,
    interval_hours,
    cleanup_old_logs,
    system_ui,
    tk_alive,
    root_after,
    tick_callback,
    is_main_thread=lambda: threading.current_thread() is threading.main_thread(),
    ui_post,
):
    def run_tick():
        if not is_enabled():
            state.after_id = None
            return

        days = normalized_retention_days(retention_days())
        try:
            deleted = cleanup_old_logs(days)
            system_ui(f"🧹 自动日志清理：已删除 {deleted} 个旧日志文件（保留 {days} 天）", "normal")
        except Exception as exc:
            system_ui(f"⚠️ 自动日志清理失败：{exc}")

        try:
            if tk_alive():
                state.after_id = root_after(interval_hours() * 3600 * 1000, tick_callback)
            else:
                state.after_id = None
        except Exception:
            state.after_id = None

    post_to_ui_thread(run_tick, is_main_thread=is_main_thread, ui_post=ui_post)


def schedule_auto_log_cleanup_runtime(
    *,
    state,
    restart=True,
    first_delay_sec=60,
    is_enabled,
    tk_alive,
    root_after,
    root_after_cancel,
    tick_callback,
    is_main_thread=lambda: threading.current_thread() is threading.main_thread(),
    ui_post,
):
    def schedule():
        if restart and state.after_id is not None:
            if tk_alive():
                try:
                    root_after_cancel(state.after_id)
                except Exception:
                    pass
            state.after_id = None

        if not is_enabled():
            return

        if not tk_alive():
            state.after_id = None
            return

        try:
            state.after_id = root_after(first_delay_sec * 1000, tick_callback)
        except Exception:
            state.after_id = None

    post_to_ui_thread(schedule, is_main_thread=is_main_thread, ui_post=ui_post)


def apply_log_cleanup_runtime(
    days,
    *,
    config,
    save_config,
    set_cleanup_state,
    system_ui,
    schedule_cleanup,
    interval_hours,
):
    set_cleanup_state(days, True)
    try:
        if not config.has_section("ui"):
            config["ui"] = {}
        config.set("ui", "auto_log_cleanup", "1")
        config.set("ui", "log_retention_days", str(days))
        save_config()
    except Exception:
        pass

    system_ui(log_cleanup_status(days, interval_hours), "normal")
    schedule_cleanup(restart=True, first_delay_sec=60)


def open_log_cleanup_dialog_runtime(
    parent,
    current_days,
    interval_hours,
    *,
    config,
    save_config,
    set_cleanup_state,
    system_ui,
    schedule_cleanup,
    center_window,
    open_dialog=open_log_cleanup_dialog,
):
    def apply_cleanup(days):
        apply_log_cleanup_runtime(
            days,
            config=config,
            save_config=save_config,
            set_cleanup_state=set_cleanup_state,
            system_ui=system_ui,
            schedule_cleanup=schedule_cleanup,
            interval_hours=interval_hours,
        )

    open_dialog(parent, current_days, interval_hours, apply_cleanup, center_window)


def open_log_cleanup_dialog_app_runtime(
    parent,
    *,
    get_retention_days,
    get_interval_hours,
    config,
    save_config,
    set_cleanup_state,
    system_ui,
    schedule_cleanup,
    center_window,
    open_dialog=open_log_cleanup_dialog,
):
    return open_log_cleanup_dialog_runtime(
        parent,
        get_retention_days(),
        get_interval_hours(),
        config=config,
        save_config=save_config,
        set_cleanup_state=set_cleanup_state,
        system_ui=system_ui,
        schedule_cleanup=schedule_cleanup,
        center_window=center_window,
        open_dialog=open_dialog,
    )


def run_auto_log_cleanup_tick_app_runtime(
    *,
    state,
    is_enabled,
    retention_days,
    interval_hours,
    cleanup_old_logs,
    system_ui,
    tk_alive,
    root_after,
    tick_callback,
    ui_post,
    run_tick_runtime=run_auto_log_cleanup_tick_runtime,
):
    return run_tick_runtime(
        state=state,
        is_enabled=is_enabled,
        retention_days=retention_days,
        interval_hours=interval_hours,
        cleanup_old_logs=cleanup_old_logs,
        system_ui=system_ui,
        tk_alive=tk_alive,
        root_after=root_after,
        tick_callback=tick_callback,
        ui_post=ui_post,
    )


def schedule_auto_log_cleanup_app_runtime(
    *,
    state,
    restart,
    first_delay_sec,
    is_enabled,
    tk_alive,
    root_after,
    root_after_cancel,
    tick_callback,
    ui_post,
    schedule_runtime=schedule_auto_log_cleanup_runtime,
):
    return schedule_runtime(
        state=state,
        restart=restart,
        first_delay_sec=first_delay_sec,
        is_enabled=is_enabled,
        tk_alive=tk_alive,
        root_after=root_after,
        root_after_cancel=root_after_cancel,
        tick_callback=tick_callback,
        ui_post=ui_post,
    )


def ensure_update_config(config, defaults=None):
    values = defaults or {
        "api_proxy_base": "https://github-api.daybyday.top/",
        "proxy_base": "https://gh-proxy.com/",
    }
    if not config.has_section("update"):
        config["update"] = dict(values)
        return True
    return False


def save_update_proxy_config(config, api_proxy_base, proxy_base, save_config):
    ensure_update_config(config)
    config.set("update", "proxy_base", normalize_proxy_base(proxy_base))
    config.set("update", "api_proxy_base", normalize_proxy_base(api_proxy_base))
    save_config()


def test_update_proxy_async(
    owner,
    repo,
    api_proxy_base,
    proxy_base,
    on_success,
    on_error,
    ui_post,
    *,
    connectivity_func=test_update_proxy_connectivity,
    formatter=format_proxy_test_result,
    thread_factory=threading.Thread,
):
    def worker():
        try:
            result = connectivity_func(owner, repo, api_proxy_base, proxy_base)
            ui_post(lambda: on_success(formatter(result)))
        except Exception as exc:
            ui_post(lambda exc=exc: on_error(str(exc)))

    thread = thread_factory(target=worker, daemon=True)
    thread.start()
    return thread


def open_update_proxy_dialog_runtime(
    parent,
    *,
    config,
    owner,
    repo,
    save_config,
    ui_post,
    center_window,
    open_dialog=open_update_proxy_dialog,
    test_async=test_update_proxy_async,
):
    ensure_update_config(config)

    def save(api_proxy_base, proxy_base, _win):
        save_update_proxy_config(config, api_proxy_base, proxy_base, save_config)

    def test_connection(api_proxy_base, proxy_base, on_success, on_error):
        test_async(owner, repo, api_proxy_base, proxy_base, on_success, on_error, ui_post)

    open_dialog(
        parent,
        config.get("update", "api_proxy_base", fallback=""),
        config.get("update", "proxy_base", fallback=""),
        save,
        test_connection,
        center_window,
    )
