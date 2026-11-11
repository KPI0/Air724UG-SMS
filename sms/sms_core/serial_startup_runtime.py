import time
import threading

from sms_core.serial_ports import is_single_port_fallback_candidate, unlocked_ports
from sms_core.serial_reconnect import is_serial_open_denied, serial_open_denied_repeat_key
from sms_core.serial_sender import (
    DEFAULT_SERIAL_TRANSACTION_LOCK, DEFAULT_SERIAL_WRITE_LOCK, _write_bytes_locked,
    write_serial_command_result,
)


class SerialInitializationSession:
    """Settle startup commands on the existing reader while owning its transaction."""

    commands = ("AT+CLIP=1", "AT+CGSN", "AT+CNUM")

    def __init__(self, serial_obj, *, serial_lock, transaction_lock, response_coordinator,
                 get_serial, get_generation, expected_generation, set_imei_deadline,
                 is_stopping, report, monotonic, response_timeout=5.0, on_finished=None):
        self.serial = serial_obj
        self.serial_lock = serial_lock
        self.transaction_lock = transaction_lock
        self.coordinator = response_coordinator
        self.get_serial = get_serial
        self.get_generation = get_generation
        self.generation = expected_generation
        self.set_imei_deadline = set_imei_deadline
        self.is_stopping = is_stopping
        self.report = report
        self.monotonic = monotonic
        self.timeout = response_timeout
        self.on_finished = on_finished
        self.has_command_errors = False
        self.owner_thread = threading.get_ident()
        self.cancelled = threading.Event()
        self.pending = True
        self.quarantined = False
        self.waiter = None
        self.index = 0
        self.attempted = False
        self.deadline = 0.0
        self.transaction_lock.acquire()
        self.lock_held = True
        try:
            serial_obj._serial_initialization_session = self
        except BaseException:
            self.transaction_lock.release()
            self.lock_held = False
            raise

    def _current(self):
        return (self.get_serial() is self.serial and self.get_generation() == self.generation
                and not self.cancelled.is_set() and not self.is_stopping() and bool(self.serial.is_open))

    def _report(self, message):
        try:
            self.report(message)
        except Exception:
            pass

    def _finish_waiter(self):
        waiter, self.waiter = self.waiter, None
        if waiter is not None:
            try:
                if not waiter.done():
                    waiter.cancel("串口初始化已收尾")
                self.coordinator.finish(waiter)
            except Exception as exc:
                self._report(f"串口初始化应答收尾异常（{type(exc).__name__}）")
                return False
        return True

    def _finish(self):
        try:
            self._finish_waiter()
        except Exception as exc:
            self._report(f"串口初始化应答收尾异常（{type(exc).__name__}）")
        finally:
            self.pending = False
            try:
                if not self.quarantined and getattr(self.serial, "_serial_initialization_session", None) is self:
                    self.serial._serial_initialization_session = None
            except Exception as exc:
                self._report(f"串口初始化属性收尾异常（{type(exc).__name__}）")
            finally:
                if self.lock_held:
                    self.transaction_lock.release()
                    self.lock_held = False
            if self.on_finished is not None:
                try:
                    self.on_finished(self)
                except Exception as exc:
                    self._report(f"串口初始化状态提示失败（{type(exc).__name__}）")

    def _fail(self, message):
        try:
            if self.attempted:
                self.quarantined = True
                with self.serial_lock, DEFAULT_SERIAL_WRITE_LOCK:
                    self.serial._sms_recovery_error = (
                        "串口初始化结果未获确认，已暂停此串口发送，请重启设备后重新连接"
                    )
        except Exception as exc:
            self._report(f"串口初始化隔离标记异常（{type(exc).__name__}）")
        finally:
            self._report(message)
            self._finish()

    def close(self):
        if not self.pending:
            return
        self.cancelled.set()
        if threading.get_ident() != self.owner_thread:
            waiter = self.waiter
            if waiter is not None:
                try:
                    waiter.cancel("串口初始化已取消")
                except Exception as exc:
                    self._report(f"串口初始化取消应答失败（{type(exc).__name__}）")
            return
        self._fail("串口初始化已取消，未继续提交启动指令")

    def observe_after_line(self):
        if not self.pending:
            return
        if threading.get_ident() != self.owner_thread:
            self.close()
            return
        if not self._current():
            self._fail("串口连接已变化，原初始化已取消")
            return
        if self.waiter is not None and self.waiter.done():
            response = self.waiter.wait(0)
            if not response.line:
                self._fail("串口初始化未获明确应答，已停止发送，请重启设备后重新连接")
                return
            command = self.commands[self.index]
            if not self._finish_waiter():
                self._fail("串口初始化等待器收尾未完成，请重启设备后重新连接")
                return
            self.index += 1
            self.attempted = False
            if not response.ok:
                self.has_command_errors = True
                self._report(f"串口初始化 {command} 收到明确错误，应答已收尾")

    def read_boundary(self):
        if not self.pending:
            return
        if threading.get_ident() != self.owner_thread:
            self.close()
            return
        try:
            if not self._current():
                self._fail("串口连接已变化，原初始化已取消")
                return
            if self.waiter is not None:
                self.observe_after_line()
                if self.pending and self.waiter is not None and self.monotonic() >= self.deadline:
                    self._fail("串口初始化应答超时，已停止发送，请重启设备后重新连接")
                return
            if self.index == len(self.commands):
                self._finish()
                return
            command = self.commands[self.index]
            waiter = self.coordinator.begin(command=command, connection=self.serial)
            self.waiter = waiter

            def before_write():
                if not self._current() or waiter.mark_write_started() is False:
                    return False
                if command == "AT+CGSN":
                    self.set_imei_deadline(self.monotonic() + 6.0)
                self.deadline = self.monotonic() + self.timeout
                self.attempted = True
                return True

            result = _write_bytes_locked(
                self.serial_lock, lambda: self.serial if self._current() else None,
                (command + "\r\n").encode("ascii"), expected_serial=self.serial,
                before_write=before_write, initialization_owner=self,
            )
            if not result.ok:
                response = waiter.wait(0)
                if response.line:
                    self.observe_after_line()
                else:
                    self._fail("串口初始化提交未获确认，已停止发送，请检查串口状态")
        except Exception as exc:
            self._fail(f"串口初始化执行失败（{type(exc).__name__}），请检查串口状态")


def resolve_serial_target_port_runtime(
    *,
    mode,
    current_port,
    reconnect_interval,
    find_luat_best_port,
    list_ports,
    is_port_locked,
    auto_connect_ui,
    set_status,
    wakeup_wait,
    wakeup_clear,
    sleep=time.sleep,
    unlocked_ports_func=unlocked_ports,
):
    if mode == "Auto":
        dev, desc = find_luat_best_port()
        if dev:
            auto_connect_ui(f"🔌 检测到 LUAT Modem 标识，自动连接：{dev}（{desc}）")
            return dev

        available_ports = unlocked_ports_func(list_ports(), is_port_locked)
        if len(available_ports) == 1 and is_single_port_fallback_candidate(available_ports[0]):
            single = available_ports[0]
            auto_connect_ui(f"🔌 未检测到 LUAT Modem 标识，但仅发现单一串口，自动连接：{single.device}")
            return single.device

        set_status("🔍 扫描 LUAT Modem 中…", "orange")
        wakeup_wait(timeout=reconnect_interval)
        wakeup_clear()
        return None

    if not current_port:
        set_status("🔒 手动模式：未指定串口", "red")
        sleep(reconnect_interval)
        return None

    return current_port


def open_and_initialize_serial_runtime(
    *,
    target_port,
    baud,
    mode,
    serial_lock,
    open_serial,
    set_serial_obj,
    set_port,
    lock_port_mutex,
    set_cloud_imei_query_deadline,
    serial_error_ui,
    set_status,
    write_command=write_serial_command_result,
    transaction_lock=DEFAULT_SERIAL_TRANSACTION_LOCK,
    is_open_denied=is_serial_open_denied,
    open_denied_repeat_key=serial_open_denied_repeat_key,
    monotonic=time.monotonic,
    response_coordinator=None,
    get_serial_obj=None,
    get_generation=None,
    is_stopping=lambda: False,
    on_initialized=None,
):
    with transaction_lock:
        session = None
        try:
            with serial_lock:
                serial_obj = open_serial(target_port, baud)
                if response_coordinator is not None:
                    session = SerialInitializationSession(
                        serial_obj, serial_lock=serial_lock, transaction_lock=transaction_lock,
                        response_coordinator=response_coordinator, get_serial=get_serial_obj,
                        get_generation=get_generation, expected_generation=get_generation() + 1,
                        set_imei_deadline=set_cloud_imei_query_deadline, is_stopping=is_stopping,
                        report=serial_error_ui, monotonic=monotonic,
                        on_finished=on_initialized,
                    )
                set_serial_obj(serial_obj)
        except Exception as exc:
            if session is not None:
                session.close()
            if is_open_denied(str(exc)):
                serial_error_ui(
                    f"⚠️ 端口占用：{target_port} 已被其他程序或本软件其他实例占用。",
                    repeat_key=open_denied_repeat_key(target_port),
                )
                set_status("🔴 端口占用，等待释放…", "red")
            raise

        try:
            lock_port_mutex(target_port)
        except BaseException:
            if session is not None:
                session.close()
            raise
        if response_coordinator is None:
            write_command(serial_obj, "AT+CLIP=1")
            try:
                set_cloud_imei_query_deadline(monotonic() + 6.0)
                write_command(serial_obj, "AT+CGSN")
            except Exception:
                pass
            try:
                write_command(serial_obj, "AT+CNUM")
            except Exception:
                pass

        try:
            if mode == "Auto":
                set_port(target_port)
        except BaseException:
            if session is not None:
                session.close()
            raise

        return serial_obj
