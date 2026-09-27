"""Local, offline firmware updates using the existing USB OTA transport."""
from dataclasses import dataclass
import hashlib
from pathlib import Path
import re
import threading
import time
import uuid

from sms_core.cloud_firmware_ota import RelayError, UsbOtaTransport
from sms_core.firmware_package import (
    CAPACITY, CORE, MAX_PACKAGE_BYTES, PackageError, parse_upload_package, version_tuple,
)
from sms_core.threading_runtime import start_registered_daemon_thread


class UpdateError(ValueError):
    """Fixed user-facing message; never expose serial frames or package contents."""


class DeviceRejected(UpdateError):
    pass


class UpdateCancelled(Exception):
    pass


REASONS = {
    "busy": "设备正在通话、收发短信、导出录音或更新，请结束后重试",
    "auth": "设备拒绝更新授权，请检查固件是否支持本地 USB 更新",
    "config_mismatch": "设备仍使用限制配置变更的旧固件，请先用原配置升级接收端，或使用 Luatools 烧录",
    "integrity": "设备校验升级包失败，请重新生成固件",
    "storage": "设备空间不足或写入失败，更新已停止",
    "unsupported": "当前设备不支持本地更新，请先用 Luatools 烧录支持 USB OTA 的固件",
    "usb_unavailable": "无法读取设备 USB AT 接口，请检查驱动、端口占用和固件版本",
    "connection_changed": "串口连接或设备身份已变化，请重新读取设备后再试",
    "wrong_device": "USB 接口的设备身份不一致，已停止更新",
    "usb_timeout": "设备回复超时，更新已停止，请检查 USB 连接",
    "usb_write": "USB 写入失败，更新已停止",
    "frame_size": "设备通信数据超出允许长度，更新已停止",
}


@dataclass(frozen=True)
class LocalSerialContext:
    serial: object
    generation: int
    imei: str


def capture_local_context(namespace):
    with namespace["serial_lock"]:
        serial = namespace.get("serial_obj")
        imei = str(namespace["_cloud_runtime_imei"]() or "").strip()
        if serial is None or not getattr(serial, "is_open", False):
            raise UpdateError("请先连接设备的 Modem 串口")
        if not re.fullmatch(r"\d{14,17}", imei):
            raise UpdateError("尚未识别设备 IMEI，请等待串口连接完成后重新读取")
        return LocalSerialContext(serial, namespace.get("serial_connection_generation", 0), imei)


def check_capability(capability):
    if not isinstance(capability, dict) or capability.get("protocol") != 1 or capability.get("supported") is not True:
        raise UpdateError(REASONS["unsupported"])
    if capability.get("core") != CORE or capability.get("capacity") != CAPACITY:
        raise UpdateError("设备 CORE 不匹配，仅支持 V4029 / RFTIPMSTSVT / 0x70000")
    for key in ("config_md5", "main_md5"):
        if not re.fullmatch(r"[0-9a-f]{32}", str(capability.get(key) or "")):
            raise UpdateError("设备未提供有效的脚本校验信息，请重新读取")
    version_tuple(capability.get("version"))
    return capability


def check_package(package, capability):
    check_capability(capability)
    if version_tuple(package.manifest["version"]) <= version_tuple(capability["version"]):
        raise UpdateError("目标版本必须高于当前版本；不支持同版本更新或降级")
    return package.manifest["config_md5"] != capability["config_md5"]


class LocalFirmwareUpdater:
    def __init__(self, namespace, *, transport_factory=UsbOtaTransport,
                 clock=time.monotonic, confirmation_timeout=150, start_worker=None):
        self.namespace = namespace
        self.transport_factory, self.clock = transport_factory, clock
        self.confirmation_timeout = confirmation_timeout
        self.start_worker = start_worker or self._start_worker
        self.lock = threading.RLock()
        self.cancel_event = threading.Event()
        self.busy = False
        self.package = None
        self.context = None
        self.state = dict(phase="idle", message="连接设备后读取固件，或选择本地升级包", progress=0,
                          filename="", current_version="", target_version="", device="", config_changed=False)

    def _stopping(self):
        event = self.namespace.get("TK_SHUTDOWN")
        return bool(self.namespace.get("is_exiting") or (event is not None and event.is_set()))

    def _current(self, context):
        if self._stopping():
            return False
        try:
            return capture_local_context(self.namespace) == context
        except UpdateError:
            return False

    def _start_worker(self, target):
        return start_registered_daemon_thread("local_firmware_update", target,
            thread_registry=self.namespace.get("UPDATE_THREAD_REGISTRY"))

    def snapshot(self):
        with self.lock:
            return dict(self.state, busy=self.busy,
                        can_start=not self.busy and self.state["phase"] == "ready" and self.package is not None,
                        can_cancel=self.busy and self.state["phase"] in {"checking", "reading", "transferring"})

    def _set(self, **changes):
        with self.lock:
            self.state.update(changes)

    def _check_cancel(self):
        if self.cancel_event.is_set() or self._stopping():
            raise UpdateCancelled()

    def _launch(self, operation, phase, message):
        with self.lock:
            if self.busy or self._stopping():
                return False
            self.busy = True
            self.cancel_event.clear()
            self.state.update(phase=phase, message=message, progress=0)

        def worker():
            try:
                self._check_cancel()
                operation()
            except UpdateCancelled:
                self._set(phase="cancelled", message="更新已取消，未提交安装")
            except (UpdateError, PackageError) as exc:
                self._set(phase="failed", message=str(exc))
            except RelayError as exc:
                self._set(phase="failed", message=REASONS.get(str(exc), "设备通信失败，请检查连接后重试"))
            except Exception:
                self._set(phase="failed", message="固件操作失败，请检查文件权限和 USB 连接后重试")
            finally:
                with self.lock:
                    self.busy = False
        try:
            self.start_worker(worker)
        except Exception:
            with self.lock:
                self.busy = False
                self.state.update(phase="failed", message="无法启动固件更新任务，请稍后重试")
            return False
        return True

    def cancel(self):
        with self.lock:
            if not self.snapshot()["can_cancel"]:
                return False
            self.cancel_event.set()
            self.state["message"] = "正在取消，等待当前 USB 操作结束…"
            return True

    def _transport(self, context):
        return self.transport_factory(context, lambda: self._current(context))

    def _read_device(self, context):
        transport = self._transport(context)
        try:
            reply = transport.exchange(dict(type="firmware_ota", action="info", request_id=uuid.uuid4().hex))
            self._check_cancel()
            if not self._current(context):
                raise UpdateError(REASONS["connection_changed"])
            if reply.get("ok") is not True:
                raise UpdateError(REASONS.get(reply.get("reason"), "读取固件信息失败"))
            return check_capability(reply.get("ota"))
        finally:
            transport.close()

    def prepare(self, path=None):
        def operation():
            with self.lock:
                self.package = None
                self.state.update(filename="", target_version="", config_changed=False)
                if path is None or not self._current(self.context):
                    self.context = None
                    self.state.update(current_version="", device="")
            try:
                context = capture_local_context(self.namespace)
                package = None
                if path is not None:
                    try:
                        with open(path, "rb") as stream:
                            data = stream.read(MAX_PACKAGE_BYTES + 1)
                    except OSError:
                        raise UpdateError("无法读取升级文件，请检查文件是否存在及读取权限") from None
                    package = parse_upload_package(data)
                    self._check_cancel()
                capability = self._read_device(context)
                with self.lock:
                    self._check_cancel()
                    if not self._current(context):
                        raise UpdateError(REASONS["connection_changed"])
                    # Device information remains valid when only the selected
                    # package fails its version or compatibility check.
                    self.context = context
                    self.state.update(current_version=capability["version"],
                        device=f"{getattr(context.serial, 'port', '')} · IMEI 尾号 {context.imei[-6:]}")
                changed = check_package(package, capability) if package else False
                with self.lock:
                    self._check_cancel()
                    self.package = package
                    self.state.update(phase="ready" if package else "idle", filename=Path(path).name if path else "",
                        target_version=package.manifest["version"] if package else "", config_changed=changed,
                        message=("预检通过，包内配置与设备不同，将随固件更新" if changed else "预检通过，可以开始更新")
                                if package else "设备支持本地更新，请选择升级包")
            finally:
                with self.lock:
                    if self.context is not None and not self._current(self.context):
                        self.context = self.package = None
                        self.state.update(current_version="", device="")
                        self._check_cancel()
                        raise UpdateError(REASONS["connection_changed"])
        return self._launch(operation, "checking" if path else "reading", "正在校验升级包并读取设备…" if path else "正在读取设备固件…")

    def _exchange(self, transport, action, job_id, *, attempts=3, **fields):
        request = dict(type="firmware_ota", action=action, job_id=job_id, request_id=uuid.uuid4().hex, **fields)
        for attempt in range(attempts):
            if action != "commit":
                self._check_cancel()
            try:
                reply = transport.exchange(request)
            except RelayError as exc:
                if str(exc) != "usb_timeout" or attempt + 1 == attempts:
                    raise
                continue
            if reply.get("ok") is not True:
                raise DeviceRejected(REASONS.get(reply.get("reason"), "设备拒绝更新或安装失败"))
            if action != "commit":
                self._check_cancel()
            return reply, request

    def _confirm_install(self, transport, request, context, package, job_id):
        deadline = self.clock() + self.confirmation_timeout
        try:
            result = transport.wait_install_result(request)
            if result.get("phase") == "failed":
                raise DeviceRejected(REASONS.get(result.get("reason"), "设备报告安装失败，请检查固件"))
            if result.get("phase") == "rebooted":
                return True
        except DeviceRejected:
            raise
        except Exception:
            # A disconnect is expected during reboot; only fresh USB evidence
            # for this IMEI, job and script can prove completion.
            pass
        transport.close()
        while self.clock() < deadline and not self._stopping():
            try:
                fresh_context = capture_local_context(self.namespace)
                if fresh_context.imei != context.imei:
                    return False
                capability = self._read_device(fresh_context)
                if (capability["last_job"] == job_id and capability["version"] == package.manifest["version"]
                        and capability["main_md5"] == package.manifest["main_md5"]):
                    transport.commit_sent = False
                    transport.close()
                    return True
            except (UpdateError, PackageError, RelayError, OSError):
                pass
            shutdown = self.namespace.get("TK_SHUTDOWN")
            if shutdown is not None:
                shutdown.wait(1)
            else:
                time.sleep(1)
        return False

    def start(self):
        with self.lock:
            if not self.snapshot()["can_start"]:
                return False
            context, package = self.context, self.package

        def operation():
            if not self._current(context):
                raise UpdateError(REASONS["connection_changed"])
            transport = self._transport(context)
            job_id = uuid.uuid4().hex
            commit_attempted = False
            try:
                reply, _ = self._exchange(transport, "info", job_id)
                check_package(package, reply.get("ota"))
                self._exchange(transport, "begin", job_id, manifest=package.manifest)
                started = self.clock()
                for offset in range(0, len(package.payload), 2048):
                    if self.clock() - started > 900:
                        raise UpdateError("传输超过 15 分钟，已停止更新")
                    block = package.payload[offset:offset + 2048]
                    reply, _ = self._exchange(transport, "chunk", job_id, offset=offset,
                                             data=block.hex(), sha256=hashlib.sha256(block).hexdigest())
                    if reply.get("offset") != offset + len(block):
                        raise UpdateError("设备确认的分块位置不一致，已停止更新")
                    self._set(progress=round((offset + len(block)) * 100 / len(package.payload)),
                              message="正在传输固件，请保持 USB 连接和设备供电")
                self._exchange(transport, "verify", job_id)
                with self.lock:
                    self._check_cancel()
                    self.state.update(phase="installing", message="正在安装并等待重启确认，请保持供电；此阶段不能取消")
                    commit_attempted = True
                request = dict(type="firmware_ota", action="commit", job_id=job_id, request_id=uuid.uuid4().hex)
                try:
                    reply = transport.exchange(request)  # Never retry commit.
                    if reply.get("ok") is not True:
                        raise DeviceRejected(REASONS.get(reply.get("reason"), "设备拒绝安装固件"))
                except DeviceRejected:
                    raise
                except Exception:
                    # The write may have reached the device even if its ACK was lost.
                    self._set(message="安装指令结果待确认，正在读取重启后的设备；不会重复提交安装")
                if self._confirm_install(transport, request, context, package, job_id):
                    self._set(phase="success", current_version=package.manifest["version"], progress=100,
                              message="固件更新成功，已确认设备重启后的版本和脚本校验值")
                else:
                    self._set(phase="unconfirmed", message="尚未确认更新结果，请恢复原设备 USB 连接后读取当前版本，勿反复提交安装")
            except DeviceRejected:
                raise
            except Exception:
                if commit_attempted:
                    self._set(phase="unconfirmed", message="安装结果未确认，请检查设备当前版本；不会自动重复安装")
                else:
                    raise
            finally:
                transport.abort()
                transport.close()
        return self._launch(operation, "transferring", "正在重新核对设备并准备传输…")
