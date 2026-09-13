"""Read-only Air724UG call-forwarding capability probe.

The probe only sends AT command capability/status queries. It never enables,
registers, disables, or erases call forwarding.
"""

from __future__ import annotations

import argparse
import re
import sys
import time
from dataclasses import dataclass

import serial
from serial.tools import list_ports


TERMINAL_RESPONSE_RE = re.compile(
    r"^(?:OK|ERROR|NO CARRIER|NO ANSWER|BUSY|\+CME ERROR:.*|\+CMS ERROR:.*)$",
    re.IGNORECASE,
)
UNSUPPORTED_RESPONSE_RE = re.compile(
    r"^(?:ERROR|\+CME ERROR:\s*(?:3|4|50|100|operation not supported|unknown).*)$",
    re.IGNORECASE,
)
LUAT_EXCLUDED_DESCRIPTIONS = ("DIAG", "NPI", "MOS", "DEBUG", "DOWNLOAD", "CP ", "AP ")
QUERY_COMMANDS = (
    ("命令能力", "AT+CCFC=?"),
    ("无条件转移", "AT+CCFC=0,2"),
    ("遇忙转移", "AT+CCFC=1,2"),
    ("无应答转移", "AT+CCFC=2,2"),
    ("不可达转移", "AT+CCFC=3,2"),
)


@dataclass(frozen=True)
class CommandResult:
    command: str
    lines: tuple[str, ...]
    terminal: str

    @property
    def succeeded(self) -> bool:
        return self.terminal.upper() == "OK"

    @property
    def explicitly_unsupported(self) -> bool:
        return bool(UNSUPPORTED_RESPONSE_RE.fullmatch(self.terminal.strip()))


def is_luat_modem_port(port) -> bool:
    description = str(getattr(port, "description", "") or "")
    hwid = str(getattr(port, "hwid", "") or "")
    haystack = (description + " " + hwid).upper()
    if "LUAT" not in haystack:
        return False
    if any(token in description.upper() for token in LUAT_EXCLUDED_DESCRIPTIONS):
        return False
    return "MODEM" in description.upper()


def select_port(explicit_port: str = "") -> str:
    if explicit_port:
        return explicit_port.strip()
    candidates = [port for port in list_ports.comports() if is_luat_modem_port(port)]
    if not candidates:
        raise RuntimeError("未找到 LUAT USB Modem 串口，请使用 --port COMx 指定")
    if len(candidates) > 1:
        devices = "、".join(str(port.device) for port in candidates)
        raise RuntimeError(f"发现多个 LUAT Modem 串口（{devices}），请使用 --port 指定")
    return str(candidates[0].device)


def read_command_response(connection, command: str, timeout: float) -> CommandResult:
    connection.reset_input_buffer()
    connection.write((command + "\r\n").encode("ascii"))
    connection.flush()
    deadline = time.monotonic() + timeout
    lines: list[str] = []
    terminal = ""
    while time.monotonic() < deadline:
        raw = connection.readline()
        if not raw:
            continue
        line = raw.decode("utf-8", "replace").strip()
        if not line or line.upper() == command.upper():
            continue
        lines.append(line)
        if TERMINAL_RESPONSE_RE.fullmatch(line):
            terminal = line
            break
    return CommandResult(command, tuple(lines), terminal or "TIMEOUT")


def classify_results(results: list[CommandResult]) -> tuple[int, str]:
    capability = results[0]
    queries = results[1:]
    if queries and all(result.succeeded for result in queries):
        return 0, "Air724UG 当前固件支持 AT+CCFC，四类呼叫转移状态均可查询。"
    if any(result.succeeded for result in queries):
        return 1, "Air724UG 支持 AT+CCFC，但部分网络状态查询失败；可能受 SIM 或运营商限制。"
    if results and all(result.explicitly_unsupported for result in results):
        return 2, "Air724UG 当前固件明确不支持 AT+CCFC 呼叫转移命令。"
    if capability.succeeded:
        return 1, "Air724UG 识别 AT+CCFC，但状态查询均失败；请检查 SIM、网络注册和运营商支持。"
    if all(result.terminal == "TIMEOUT" for result in results):
        return 3, "能力查询超时，无法判断；请确认端口未被客户端占用后重试。"
    return 3, "AT+CCFC 返回异常，当前结果不足以确认支持情况。"


def run_probe(port: str, baud: int, timeout: float) -> int:
    print(f"串口：{port} @ {baud}")
    print("安全模式：仅查询，不修改呼叫转移设置")
    results: list[CommandResult] = []
    try:
        with serial.Serial(port, baud, timeout=0.2, write_timeout=1) as connection:
            communication = read_command_response(connection, "AT", timeout)
            print(f"\n[通信测试] AT -> {communication.terminal}")
            if not communication.succeeded:
                print("结论：串口通信失败，未继续检测。")
                return 3
            for label, command in QUERY_COMMANDS:
                result = read_command_response(connection, command, timeout)
                results.append(result)
                print(f"\n[{label}] {command}")
                for line in result.lines:
                    print("  " + line)
                if not result.lines:
                    print("  " + result.terminal)
    except serial.SerialException as exc:
        print(f"无法打开或使用串口：{exc}", file=sys.stderr)
        print("请先关闭占用该 Modem 串口的客户端或其他串口工具。", file=sys.stderr)
        return 3

    exit_code, conclusion = classify_results(results)
    print("\n结论：" + conclusion)
    return exit_code


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="只读检测 Air724UG 是否支持呼叫转移")
    parser.add_argument("--port", default="", help="Modem 串口，例如 COM210；留空则自动识别")
    parser.add_argument("--baud", type=int, default=115200, help="串口波特率，默认 115200")
    parser.add_argument("--timeout", type=float, default=12.0, help="每条命令等待秒数，默认 12")
    return parser


def main() -> int:
    args = build_parser().parse_args()
    if args.baud <= 0:
        print("波特率必须大于 0", file=sys.stderr)
        return 3
    if args.timeout <= 0 or args.timeout > 120:
        print("超时时间必须在 0-120 秒之间", file=sys.stderr)
        return 3
    try:
        port = select_port(args.port)
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        return 3
    return run_probe(port, args.baud, args.timeout)


if __name__ == "__main__":
    raise SystemExit(main())
