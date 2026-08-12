import asyncio
import json
import random
import ssl
import threading
from urllib.parse import urlsplit

from sms_core.cloud_messages import is_cloud_auth_ack_type, parse_cloud_message
from sms_core.cloud_protocol import auth_status_from_ack
from sms_core.cloud_runtime import (
    cloud_backoff_sleep_ticks,
    cloud_stopped_status,
    next_cloud_backoff,
)


def cloud_thread_main_runtime(
    url,
    reconnect_interval,
    *,
    lock,
    set_loop,
    get_thread,
    set_thread,
    run_main,
    log,
    new_event_loop=asyncio.new_event_loop,
    set_event_loop=asyncio.set_event_loop,
    current_thread=threading.current_thread,
):
    loop = new_event_loop()
    with lock:
        set_loop(loop)

    try:
        set_event_loop(loop)
        loop.run_until_complete(run_main(url, reconnect_interval))
        return "stopped"
    except Exception as exc:
        log(f"云端控制线程异常：{exc}")
        return "error"
    finally:
        with lock:
            set_loop(None)
            if current_thread() is get_thread():
                set_thread(None)
        try:
            loop.close()
        except Exception:
            pass


def base_cloud_backoff(reconnect_interval):
    try:
        return max(1.0, float(reconnect_interval))
    except Exception:
        return 1.0


def build_cloud_ssl_context(
    url,
    *,
    create_default_context=ssl.create_default_context,
    certifi_where=None,
):
    """为 WSS 合并系统信任库和 certifi CA；普通 WS 不创建 TLS 上下文。"""
    scheme = urlsplit(str(url or "").strip()).scheme.lower()
    if scheme != "wss":
        return None

    if certifi_where is None:
        try:
            from certifi import where as certifi_where
        except Exception as exc:
            raise RuntimeError("WSS 连接缺少 certifi 根证书库，请重新安装客户端依赖") from exc

    context = create_default_context()
    try:
        context.load_verify_locations(cafile=certifi_where())
    except Exception as exc:
        raise RuntimeError("WSS 根证书库加载失败，请重新安装客户端") from exc
    return context


def cloud_websocket_connect_kwargs(url, *, ssl_context_factory=build_cloud_ssl_context):
    kwargs = {
        "ping_interval": 30,
        "ping_timeout": 30,
    }
    ssl_context = ssl_context_factory(url)
    if ssl_context is not None:
        kwargs["ssl"] = ssl_context
    return kwargs


def current_cloud_control_enabled(value):
    try:
        return bool(value()) if callable(value) else bool(value)
    except Exception:
        return False


async def wait_for_cloud_imei(
    *,
    stop_event,
    runtime_imei,
    set_cloud_status,
    request_cloud_device_imei,
    last_imei_request,
    monotonic,
    sleep,
    request_interval=5.0,
):
    while not stop_event.is_set() and not runtime_imei():
        set_cloud_status("🌐 等待读取IMEI", "#b26a00")
        now = monotonic()
        if now - last_imei_request >= request_interval:
            request_cloud_device_imei()
            last_imei_request = now
        await sleep(0.5)
    return last_imei_request


async def wait_cloud_login_ack_runtime(
    ws,
    *,
    stop_event,
    set_authorized,
    set_auth_status_from_ack,
    log,
    safe_preview,
    timeout=8.0,
    monotonic,
    wait_for=asyncio.wait_for,
):
    deadline = monotonic() + float(timeout)
    while not stop_event.is_set():
        remain = deadline - monotonic()
        if remain <= 0:
            set_authorized(False)
            log("设备登录确认超时，已停止本次云端连接", show_main=True)
            return False

        try:
            msg = await wait_for(ws.recv(), timeout=min(0.5, remain))
        except asyncio.TimeoutError:
            continue

        incoming, _error = parse_cloud_message(msg)
        data = incoming.data
        if data is None:
            continue

        if not is_cloud_auth_ack_type(incoming.msg_type):
            log(f"登录确认前已忽略云端消息：{safe_preview(json.dumps(data, ensure_ascii=False))}")
            continue

        status = auth_status_from_ack(data)
        set_authorized(status == "authorized")
        set_auth_status_from_ack(data)
        if status == "authorized":
            log(str(data.get("message") or "服务端已确认设备密码"), show_main=True)
            return True

        if status == "waiting":
            log(
                str(data.get("message") or "设备正在等待网页端绑定"),
                show_main=True,
            )
            return True

        log(str(data.get("message") or "服务端未授权设备登录，请先在网页端添加正确 IMEI 和控制密码"), show_main=True)
        # 明确的密码/设备授权拒绝不是网络故障。返回专用结果，调用方
        # 必须停止自动重连，避免错误密码形成无限重试和服务端限流。
        return "auth_failed"

    return False


async def wait_for_cloud_auth_retry_runtime(
    ws,
    *,
    stop_event,
    send_register,
    wait_login_ack,
    set_connection_state,
    set_cloud_status,
    reset_serial_log_state,
    log,
    wait_for=asyncio.wait_for,
):
    """Hold a rejected socket until the web console asks for a retry.

    This is deliberately event-driven: no password-error reconnect timer is
    created.  A retry can only begin after the server has durably accepted a
    web-side device-password update and sends ``device_auth_retry``.
    """
    while not stop_event.is_set():
        try:
            message = await wait_for(ws.recv(), timeout=0.5)
        except asyncio.TimeoutError:
            continue

        incoming, _error = parse_cloud_message(message)
        if incoming.msg_type != "device_auth_retry":
            continue

        set_connection_state(ws, connected=True, authorized=False)
        set_cloud_status("🌐 等待授权", "#b26a00")
        log("收到网页端控制密码更新通知，正在重新尝试设备授权", show_main=True)
        await send_register(ws)
        result = await wait_login_ack(ws)
        if result is True:
            return True

        # Another explicit rejection keeps the same socket in the hold state;
        # it never falls through to the normal reconnect backoff.
        set_connection_state(ws, connected=False, authorized=False)
        reset_serial_log_state()
        set_cloud_status("🌐 授权失败", "#cc0000")
        log("设备授权仍未通过，等待网页端再次更新控制密码", show_main=True)

    return False


async def cloud_ws_main_runtime(
    url,
    reconnect_interval,
    *,
    stop_event,
    runtime_imei,
    request_cloud_device_imei,
    set_cloud_status,
    log,
    connect,
    set_connection_state,
    reset_serial_log_state,
    send_register,
    wait_login_ack,
    handle_message,
    cloud_control_enabled,
    monotonic,
    sleep=asyncio.sleep,
    wait_for=asyncio.wait_for,
    jitter=random.uniform,
    schedule_pending_sms_events=None,
):
    schedule_pending_sms_events = schedule_pending_sms_events or (lambda: None)
    last_imei_request = 0.0
    current_backoff = base_cloud_backoff(reconnect_interval)
    auth_blocked = False

    while not stop_event.is_set():
        try:
            last_imei_request = await wait_for_cloud_imei(
                stop_event=stop_event,
                runtime_imei=runtime_imei,
                set_cloud_status=set_cloud_status,
                request_cloud_device_imei=request_cloud_device_imei,
                last_imei_request=last_imei_request,
                monotonic=monotonic,
                sleep=sleep,
            )

            if stop_event.is_set():
                break

            set_cloud_status("🌐 连接中", "#b26a00")
            log(f"正在连接：{url}", show_main=True)

            async with connect(url, **cloud_websocket_connect_kwargs(url)) as ws:
                set_connection_state(ws, connected=True, authorized=False)
                set_cloud_status("🌐 等待授权", "#b26a00")
                log(f"已连接：{url}", show_main=True)
                await send_register(ws)
                login_result = await wait_login_ack(ws)
                if login_result == "auth_failed":
                    # 先设置终态，再执行清理。即使 websocket 已被服务端
                    # 关闭、close() 抛出异常，也不能回到通用重连分支。
                    auth_blocked = True
                    # Keep the socket reference so the manual Disconnect or
                    # the web-triggered restart can close the held channel.
                    set_connection_state(ws, connected=False, authorized=False)
                    reset_serial_log_state()
                    set_cloud_status("🌐 授权失败", "#cc0000")
                    try:
                        retry_authorized = await wait_for_cloud_auth_retry_runtime(
                            ws,
                            stop_event=stop_event,
                            send_register=send_register,
                            wait_login_ack=wait_login_ack,
                            set_connection_state=set_connection_state,
                            set_cloud_status=set_cloud_status,
                            reset_serial_log_state=reset_serial_log_state,
                            log=log,
                            wait_for=wait_for,
                        )
                    except asyncio.CancelledError:
                        raise
                    except Exception:
                        # The password rejection remains valid, but the held
                        # transport has disappeared.  Reconnect with the normal
                        # network backoff so a later web-side password update
                        # can still reach this client.  The server keeps an
                        # intact rejected connection in the event-driven hold,
                        # so this does not restore password-error polling.
                        auth_blocked = False
                        raise
                    if retry_authorized:
                        auth_blocked = False
                        login_result = True
                        current_backoff = base_cloud_backoff(reconnect_interval)
                        schedule_pending_sms_events()
                    else:
                        break
                if not login_result:
                    set_connection_state(None, connected=False, authorized=False)
                    reset_serial_log_state()
                    set_cloud_status("🌐 授权失败", "#cc0000")
                    await ws.close()
                    raise RuntimeError("设备登录确认超时，等待网络恢复后重试")
                current_backoff = base_cloud_backoff(reconnect_interval)
                schedule_pending_sms_events()

                while not stop_event.is_set():
                    try:
                        msg = await wait_for(ws.recv(), timeout=0.5)
                    except asyncio.TimeoutError:
                        continue
                    message_result = await handle_message(ws, msg)
                    schedule_pending_sms_events()
                    if message_result == "auth_failed":
                        auth_blocked = True
                        set_connection_state(None, connected=False, authorized=False)
                        reset_serial_log_state()
                        set_cloud_status("🌐 授权失败", "#cc0000")
                        try:
                            await ws.close()
                        except Exception:
                            pass
                        break

            if auth_blocked:
                break

        except asyncio.CancelledError:
            break
        except Exception as e:
            if stop_event.is_set() or auth_blocked:
                break
            set_connection_state(None, connected=False, authorized=False)
            reset_serial_log_state()
            err = str(e).strip() or e.__class__.__name__
            set_cloud_status("🌐 重连中", "#b26a00")
            log(f"连接异常：{err}")

            for _ in range(cloud_backoff_sleep_ticks(current_backoff)):
                if stop_event.is_set():
                    break
                await sleep(0.1)

            if not stop_event.is_set():
                await sleep(jitter(0, 0.5))
                current_backoff = next_cloud_backoff(current_backoff)

    set_connection_state(None, connected=False, authorized=False)
    reset_serial_log_state()
    if auth_blocked:
        # 保持“授权失败”可见。正常情况下服务端会在网页端保存新密码
        # 后通知这个等待中的连接；只有服务端或客户端主动关闭连接时才
        # 会走到这里，此时不应伪装成普通网络重连。
        set_cloud_status("🌐 授权失败", "#cc0000")
        log("授权失败，已停止自动重连；请在网页端更新控制密码后等待自动重试")
    else:
        set_cloud_status(
            cloud_stopped_status(current_cloud_control_enabled(cloud_control_enabled)),
            "#666666",
        )
        log("连接已停止")


async def cloud_ws_main_app_runtime(
    url,
    reconnect_interval,
    *,
    stop_event,
    runtime_imei,
    request_cloud_device_imei,
    set_cloud_status,
    log,
    connect,
    set_ws,
    set_connected,
    set_authorized,
    reset_serial_log_state,
    send_register,
    wait_login_ack,
    handle_message,
    cloud_control_enabled,
    monotonic,
    run_ws_main=cloud_ws_main_runtime,
    schedule_pending_sms_events=None,
):
    def set_connection_state(ws, connected, authorized):
        set_ws(ws)
        set_connected(bool(connected))
        set_authorized(bool(authorized))

    await run_ws_main(
        url,
        reconnect_interval,
        stop_event=stop_event,
        runtime_imei=runtime_imei,
        request_cloud_device_imei=request_cloud_device_imei,
        set_cloud_status=set_cloud_status,
        log=log,
        connect=connect,
        set_connection_state=set_connection_state,
        reset_serial_log_state=reset_serial_log_state,
        send_register=send_register,
        wait_login_ack=wait_login_ack,
        handle_message=handle_message,
        cloud_control_enabled=cloud_control_enabled,
        monotonic=monotonic,
        schedule_pending_sms_events=schedule_pending_sms_events,
    )
