import hashlib
import json
import re
import time
from dataclasses import dataclass

from sms_core.cloud_command_security import (
    cloud_command_batch_error,
    cloud_command_control_char_error,
    sensitive_cloud_command_decision,
)


SECRET_KEYS = {"secret", "device_secret", "password", "pwd", "token"}
SMS_META_KEYS = {"sms_phone", "sms_message"}
SMS_COMMAND_KEYS = {"cmd", "command", "data"}
HIDDEN_SMS_COMMAND = "短信PDU命令（已隐藏）"
HIDDEN_SMS_META = "短信元数据（已隐藏）"
HIDDEN_SENSITIVE_COMMAND = "云端敏感指令（内容已隐藏）"
SECRET_TEXT_RE = re.compile(
    r"(?i)"
    r"([\"']?(?:secret|device_secret|password|pwd|token)[\"']?"
    r"\s*(?:=|:)\s*)"
    r"([\"']?)"
    r"([^&\s,;\"'}]+)"
    r"([\"']?)"
)


@dataclass(frozen=True)
class CloudReplayCheckResult:
    ok: bool
    log_message: str = ""
    payload: dict | None = None


def safe_preview(raw: str, limit: int = 500) -> str:
    """Mask common secret fields in JSON text and truncate long previews."""
    def _mask(obj):
        if isinstance(obj, dict):
            sms_log = str(obj.get("sms_log") or "").strip().lower()
            command_kind = str(obj.get("command_kind") or "").strip().lower()
            command_value = next(
                (
                    obj.get(key)
                    for key in ("command", "data", "cmd")
                    if isinstance(obj.get(key), (str, bytes)) and str(obj.get(key)).strip()
                ),
                "",
            )
            sensitive_decision = sensitive_cloud_command_decision(command_value, obj)
            sensitive_reason = sensitive_decision.reason
            hide_sensitive_command = bool(sensitive_reason)
            command_is_sms = sensitive_decision.category == "sms"
            hide_sms_command = command_is_sms and sms_log == "suppress"
            hide_sms_meta = command_is_sms and (
                sms_log in ("summary", "suppress") or command_kind == "send_sms"
            )
            hide_batched_command = bool(cloud_command_batch_error(command_value))
            hide_control_command = bool(cloud_command_control_char_error(command_value))
            masked = {}
            for key, value in obj.items():
                key_lower = str(key).lower()
                if key_lower in SECRET_KEYS:
                    masked[key] = "***"
                elif hide_sms_command and key_lower in SMS_COMMAND_KEYS:
                    masked[key] = HIDDEN_SMS_COMMAND
                elif hide_sensitive_command and key_lower in SMS_COMMAND_KEYS:
                    masked[key] = HIDDEN_SENSITIVE_COMMAND
                elif hide_batched_command and key_lower in SMS_COMMAND_KEYS:
                    masked[key] = "云端 AT 指令（已隐藏：包含多条指令）"
                elif hide_control_command and key_lower in SMS_COMMAND_KEYS:
                    masked[key] = "云端 AT 指令（已隐藏：包含不支持的控制字符）"
                elif hide_sms_meta and key_lower in SMS_META_KEYS:
                    masked[key] = HIDDEN_SMS_META
                else:
                    masked[key] = _mask(value)
            return masked
        if isinstance(obj, list):
            return [_mask(item) for item in obj]
        return obj

    try:
        data = json.loads(str(raw))
        text = json.dumps(_mask(data), ensure_ascii=False)
    except Exception:
        text = SECRET_TEXT_RE.sub(r"\1\2***\4", str(raw))
    return text if len(text) <= limit else text[:limit] + "..."


def read_unix_timestamp(data: dict):
    # Use explicit None check to avoid treating 0 timestamp as falsy
    raw_ts = data.get("timestamp")
    if raw_ts is None:
        raw_ts = data.get("ts")
    if raw_ts is None:
        raw_ts = data.get("unix_time")
    if raw_ts is None:
        raw_ts = data.get("time")

    if raw_ts is None or raw_ts == "":
        return None, None

    try:
        ts = int(float(str(raw_ts).strip()))
    except Exception:
        return None, raw_ts

    if ts > 10_000_000_000:
        ts = ts // 1000
    return ts, raw_ts


def replay_key(data: dict, ts: int) -> str:
    nonce = str(data.get("nonce") or data.get("request_id") or data.get("rid") or "").strip()
    if nonce:
        return "nonce:" + nonce
    command = str(data.get("command") or data.get("data") or data.get("cmd") or "").strip()
    action = str(data.get("type") or data.get("action") or "").strip().lower()
    target = str(data.get("target_imei") or data.get("imei") or data.get("device_imei") or "").strip()
    raw = json.dumps(
        {
            "ts": ts,
            "action": action,
            "target": target,
            "command": command,
        },
        ensure_ascii=False,
        sort_keys=True,
        separators=(",", ":"),
    )
    return "fp:" + hashlib.sha256(raw.encode("utf-8")).hexdigest()


def replay_error_payload(message, task_id=""):
    payload = {
        "type": "error",
        "ok": False,
        "message": message,
    }
    task_id = str(task_id or "").strip()
    if task_id:
        payload = {
            **payload,
            "task_id": task_id,
            "command_task_id": task_id,
        }
    return payload


def check_replay_window(
    data: dict,
    cache: dict,
    now_ts: int,
    window_seconds: int,
    max_size: int,
    mark_seen: bool = True,
) -> CloudReplayCheckResult:
    task_id = str(data.get("task_id") or data.get("command_task_id") or "").strip()
    ts, raw_ts = read_unix_timestamp(data)
    if ts is None:
        return CloudReplayCheckResult(
            ok=False,
            log_message="已拒绝云端指令：缺少 Unix 时间戳字段 timestamp/ts",
            payload=replay_error_payload(
                "安全拦截：缺少 Unix 时间戳，请使用 timestamp 或 ts 秒级时间戳",
                task_id,
            ),
        )

    delta = abs(int(now_ts) - ts)
    if delta > int(window_seconds):
        return CloudReplayCheckResult(
            ok=False,
            log_message=f"已拒绝云端指令：时间戳超时或疑似重放攻击 (timestamp={raw_ts}, delta={delta}s)",
            payload=replay_error_payload(
                "安全拦截：指令已过期，请检查服务器和本机时钟是否同步",
                task_id,
            ),
        )

    if mark_seen:
        prune_replay_cache(cache, now_ts, window_seconds, max_size)
        key = replay_key(data, ts)
        if key in cache:
            return CloudReplayCheckResult(
                ok=False,
                log_message=f"已拒绝云端指令：检测到重复 nonce/指令指纹 (timestamp={raw_ts})",
                payload=replay_error_payload(
                    "安全拦截：检测到重复指令，疑似重放攻击",
                    task_id,
                ),
            )
        cache[key] = now_ts

    return CloudReplayCheckResult(ok=True)


def prune_replay_cache(cache: dict, now_ts: int, window_seconds: int, max_size: int):
    """Prune an in-memory replay cache in place."""
    try:
        expired = [
            key for key, seen_ts in cache.items()
            if now_ts - int(seen_ts) > window_seconds
        ]
        for key in expired:
            cache.pop(key, None)
        while len(cache) > max_size:
            oldest = min(cache, key=cache.get)
            cache.pop(oldest, None)
    except Exception:
        cache.clear()


def repeat_filter(
    state: dict,
    msg: str,
    reset_seconds: float,
    repeat_limit: int,
    now=None,
):
    current = time.monotonic() if now is None else float(now)
    last_seen, count = state.get(msg, (0.0, 0))
    if current - last_seen > reset_seconds:
        count = 0
    count += 1
    state[msg] = (current, count)

    if count < repeat_limit:
        return msg
    if count == repeat_limit:
        return f"{msg}（后续同类消息已忽略）"
    return None
