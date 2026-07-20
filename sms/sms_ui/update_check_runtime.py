import threading

from sms_core.threading_runtime import start_registered_daemon_thread
from sms_core.updates import check_latest_release, normalize_config_bases
from sms_ui.thread_runtime import post_ui_if_running_runtime


def read_update_config_runtime(config):
    proxy_base = config.get("update", "proxy_base", fallback="https://gh-proxy.com/").strip()
    api_proxy_base = config.get("update", "api_proxy_base", fallback="").strip()
    return normalize_config_bases(proxy_base, api_proxy_base)


def handle_update_check_plan(
    plan,
    current_version,
    *,
    show_info,
    show_warning,
    ask_open_download,
    open_url,
):
    if plan.kind == "latest":
        show_info("检测更新", f"当前已是最新版本：v{current_version}")
        return "latest"

    if plan.kind == "no_zip":
        show_warning("检测更新", f"发现新版本：{plan.latest_tag}\n但 Release 里没有 .zip 附件。")
        return "no_zip"

    ok = ask_open_download(
        "发现新版本",
        f"当前：v{current_version}\n最新：{plan.latest_tag}\n\n是否打开下载链接？（如已配置下载代理，将优先使用）",
    )
    if ok:
        try:
            open_url(plan.download_url)
        except Exception as exc:
            # If the browser fails to open, swallowing the error leaves the
            # user waiting with nothing. Surface the URL so they can copy it
            # and download manually, plus the reason it could not auto-open.
            show_warning(
                "无法打开下载链接",
                f"自动打开浏览器失败：{exc}\n\n请手动复制以下链接下载：\n{plan.download_url}",
            )
    return "update"


def check_update_and_prompt_runtime(
    *,
    owner,
    repo,
    current_version,
    get_update_config,
    ui_post,
    show_info,
    show_warning,
    show_error,
    ask_open_download,
    open_url,
    check_latest=check_latest_release,
    thread_factory=threading.Thread,
    log_error=None,
    is_stopping=lambda: False,
    thread_registry=None,
    task_state=None,
):
    def release_task():
        if task_state is not None:
            task_state.release()

    def post_completion(callback):
        def run_and_release():
            try:
                return callback()
            finally:
                release_task()

        return post_ui_if_running_runtime(
            ui_post,
            run_and_release,
            is_stopping,
            on_skipped=release_task,
        )

    def worker():
        if is_stopping():
            release_task()
            return None
        callback_handed_off = False
        try:
            proxy_base, api_proxy_base = get_update_config()
            plan = check_latest(owner, repo, current_version, proxy_base, api_proxy_base)
            callback = lambda: handle_update_check_plan(
                plan,
                current_version,
                show_info=show_info,
                show_warning=show_warning,
                ask_open_download=ask_open_download,
                open_url=open_url,
            )
        except Exception as exc:
            callback = lambda exc=exc: show_error("检测更新失败", str(exc))
        try:
            callback_handed_off = post_completion(callback)
        finally:
            if not callback_handed_off:
                release_task()

    if is_stopping():
        return None
    if task_state is not None and not task_state.try_acquire():
        return None
    if is_stopping():
        release_task()
        return None
    try:
        return start_registered_daemon_thread(
            "update_check",
            worker,
            thread_registry=thread_registry,
            log_error=log_error,
            thread_factory=thread_factory,
        )
    except Exception:
        release_task()
        raise
