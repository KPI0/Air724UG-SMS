import asyncio
import json
import random
import threading

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

        log(str(data.get("message") or "服务端未授权设备登录，请先在网页端添加正确 IMEI 和控制密码"), show_main=True)
        return False

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
):
    last_imei_request = 0.0
    current_backoff = base_cloud_backoff(reconnect_interval)

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

            async with connect(url, ping_interval=30, ping_timeout=30) as ws:
                set_connection_state(ws, connected=True, authorized=False)
                set_cloud_status("🌐 等待授权", "#b26a00")
                log(f"已连接：{url}", show_main=True)
                await send_register(ws)
                if not await wait_login_ack(ws):
                    set_connection_state(None, connected=False, authorized=False)
                    reset_serial_log_state()
                    set_cloud_status("🌐 授权失败", "#cc0000")
                    await ws.close()
                    raise RuntimeError("设备登录未通过服务端确认")
                current_backoff = base_cloud_backoff(reconnect_interval)

                while not stop_event.is_set():
                    try:
                        msg = await wait_for(ws.recv(), timeout=0.5)
                    except asyncio.TimeoutError:
                        continue
                    message_result = await handle_message(ws, msg)
                    if message_result == "auth_failed":
                        set_connection_state(None, connected=False, authorized=False)
                        reset_serial_log_state()
                        set_cloud_status("🌐 授权失败", "#cc0000")
                        await ws.close()
                        raise RuntimeError("设备会话授权已失效")

        except asyncio.CancelledError:
            break
        except Exception as e:
            if stop_event.is_set():
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
    set_cloud_status(cloud_stopped_status(cloud_control_enabled), "#666666")
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
    )
