import threading

from sms_core.threading_runtime import start_daemon_thread
from sms_core.updates import check_latest_release, normalize_config_bases


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
        except Exception:
            pass
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
):
    def worker():
        try:
            proxy_base, api_proxy_base = get_update_config()
            plan = check_latest(owner, repo, current_version, proxy_base, api_proxy_base)
            ui_post(
                lambda: handle_update_check_plan(
                    plan,
                    current_version,
                    show_info=show_info,
                    show_warning=show_warning,
                    ask_open_download=ask_open_download,
                    open_url=open_url,
                )
            )
        except Exception as exc:
            ui_post(lambda exc=exc: show_error("检测更新失败", str(exc)))

    return start_daemon_thread(
        "update_check",
        worker,
        log_error=log_error,
        thread_factory=thread_factory,
    )
