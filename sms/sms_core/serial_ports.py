from dataclasses import dataclass


LUAT_EXCLUDE_DESC_TOKENS = (
    "DIAG",
    "NPI",
    "MOS",
    "DEBUG",
    "DOWNLOAD",
    "CP ",
    "CP_",
    "AP ",
    "AP_",
)


@dataclass(frozen=True)
class SerialPortInfo:
    device: str
    description: str = ""
    hwid: str = ""


@dataclass(frozen=True)
class ManualRebindCandidate:
    device: str = ""
    description: str = ""

    @property
    def found(self) -> bool:
        return bool(self.device)


def serial_port_info(port) -> SerialPortInfo:
    return SerialPortInfo(
        device=str(getattr(port, "device", "") or ""),
        description=str(getattr(port, "description", "") or ""),
        hwid=str(getattr(port, "hwid", "") or ""),
    )


def is_luat_modem_candidate(port_info, remembered_port="") -> bool:
    dev = str(port_info.device or "")
    desc_u = str(port_info.description or "").upper()
    hwid_u = str(port_info.hwid or "").upper()

    if "LUAT" not in desc_u and "LUAT" not in hwid_u:
        return False
    if any(token in desc_u for token in LUAT_EXCLUDE_DESC_TOKENS):
        return False
    if " AT" in desc_u or desc_u.endswith("AT"):
        return False
    if "MODEM" not in desc_u and dev != remembered_port:
        return False
    return True


def luat_modem_score(port_info, remembered_port="") -> int:
    score = 0
    desc_u = str(port_info.description or "").upper()
    if "MODEM" in desc_u:
        score += 100
    if "USB DEVICE 0" in desc_u:
        score += 10
    if port_info.device == remembered_port:
        score += 1000
    return score


def choose_luat_modem_port(ports, remembered_port="", is_locked=None):
    is_locked = is_locked or (lambda _device: False)
    candidates = []
    for port in ports:
        info = serial_port_info(port)
        if not info.device or is_locked(info.device):
            continue
        if not is_luat_modem_candidate(info, remembered_port):
            continue
        candidates.append((luat_modem_score(info, remembered_port), info))

    if not candidates:
        return None, None

    candidates.sort(reverse=True, key=lambda item: item[0])
    best = candidates[0][1]
    return best.device, best.description


def unlocked_ports(ports, is_locked=None):
    is_locked = is_locked or (lambda _device: False)
    return [port for port in ports if not is_locked(serial_port_info(port).device)]


def choose_manual_rebind_candidate(luat_device, luat_description, all_ports, current_port=""):
    device = str(luat_device or "")
    description = str(luat_description or "")

    if not device:
        ports = list(all_ports or [])
        if len(ports) == 1:
            info = serial_port_info(ports[0])
            device = info.device
            description = info.description

    if not device or device == current_port:
        return ManualRebindCandidate()
    return ManualRebindCandidate(device, description)


def manual_rebind_hint(old_port, new_port, description="", reason=""):
    hint = f"🔁 手动模式端口失效，已自动重绑：{old_port} -> {new_port}"
    if description:
        hint += f"（{description}）"
    if reason:
        hint += f"；原因：{reason}"
    return hint
