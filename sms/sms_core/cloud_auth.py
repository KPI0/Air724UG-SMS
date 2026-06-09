import hmac

from sms_core.cloud_protocol import normalize_imei


def command_action(data: dict):
    action = str(data.get("type") or data.get("action") or "").strip().lower()
    if not action and data.get("cmd"):
        action = "cmd"
    return action


def command_text(data: dict):
    return data.get("command") or data.get("data") or data.get("cmd") or ""


def incoming_secret(data: dict):
    return str(
        data.get("secret")
        or data.get("device_secret")
        or data.get("password")
        or data.get("pwd")
        or data.get("token")
        or ""
    ).strip()


def target_value(data: dict):
    return (
        data.get("target_imei")
        or data.get("imei")
        or data.get("device_imei")
        or data.get("target")
        or data.get("device")
    )


def target_imeis(data: dict):
    target = target_value(data)
    raw_targets = list(target) if isinstance(target, (list, tuple, set)) else [target]
    return target, [normalize_imei(item) for item in raw_targets]


def secret_match_result(data: dict, expected_secret: str):
    expected = str(expected_secret or "").strip()
    incoming = incoming_secret(data)

    if not expected:
        return False, "已拒绝云端指令：本机云端控制密码为空"
    if not incoming:
        return False, "已拒绝云端指令：缺少密码"
    if not hmac.compare_digest(incoming, expected):
        return False, "已拒绝云端指令：密码错误"
    return True, ""


def target_match_result(data: dict, local_imei: str):
    target, targets = target_imeis(data)

    if target is None or target == "":
        return False, "已拒绝云端指令：缺少目标IMEI"

    normalized_local = normalize_imei(local_imei)
    if not normalized_local:
        return False, f"已拒绝云端指令：本机IMEI未知，目标={target}"

    if normalized_local not in targets:
        return False, f"已忽略非本机指令：本机IMEI={normalized_local}，目标={target}"
    return True, ""


def auth_match_result(data: dict, local_imei: str, expected_secret: str):
    target_ok, target_reason = target_match_result(data, local_imei)
    if not target_ok:
        return False, target_reason
    return secret_match_result(data, expected_secret)
