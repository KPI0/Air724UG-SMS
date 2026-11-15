import queue
import re
import threading
from contextlib import ExitStack
from types import SimpleNamespace

from sms_core.cloud_payloads import identity_payload
from sms_core.cloud_protocol import normalize_imei, parse_sms_callback_head
from sms_core.cloud_state_namespace_runtime import (
    cloud_identity_payload_namespace_runtime, cloud_runtime_imei_namespace_runtime,
    maybe_capture_cloud_device_imei_namespace_runtime, set_cloud_device_imei_namespace_runtime,
)
from sms_core.serial_app_runtime import run_serial_reader_namespace_runtime
from sms_core.serial_io_runtime import read_serial_line_safely_runtime
from sms_core.serial_namespace_runtime import (
    open_and_initialize_serial_namespace_runtime, schedule_delayed_connected_log_namespace_runtime,
)
from sms_ui.app_infrastructure_namespace_runtime import safe_close_serial_namespace_runtime
from sms_ui.call_popup_namespace_runtime import (
    get_serial_call_state_namespace_runtime, merge_serial_call_state_namespace_runtime,
    set_serial_call_state_namespace_runtime,
)
from tests import test_local_call_command_isolation as call_fixtures


class InitializationHarness:
    """Production reader with fake serial I/O and no startup-session contract."""
    def __init__(self, *, hold="AT+CNUM", stage="normal", connected_error=False,
                 error_message="synthetic-private-startup-error"):
        self.fixture = call_fixtures.LocalDialButtonIsolationTests("runTest")
        self.fixture.setUp()
        self.ns = namespace = self.fixture.ns
        self.stop = threading.Event()
        self.ready = threading.Event()
        self.clock = 100.0
        self.errors, self.messages, self.lines, self.recording_lines = [], [], [], []
        self.statuses = []
        self.on_line = None
        self.connected_error = connected_error
        self.reads = queue.Queue()
        self.hold, self.stage = hold, stage
        self.wrote = {command: threading.Event() for command in ("AT+CLIP=1", "AT+CGSN", "AT+CNUM", "ATD10086;")}
        harness = self

        class Modem:
            is_open = True

            def __init__(self):
                self.writes = []

            @property
            def in_waiting(self):
                return 4096

            def write(self, payload):
                self.writes.append(payload)
                command = payload.decode("ascii").strip()
                if command in harness.wrote:
                    harness.wrote[command].set()
                if command.startswith("AT+"):
                    if harness.stage == "write-error":
                        raise OSError(error_message)
                    if harness.stage == "partial":
                        return 1
                    if command != harness.hold:
                        harness.terminal(command)
                return len(payload)

            def flush(self):
                if harness.stage == "flush-error":
                    raise OSError(error_message)

            def read(self, _size):
                try:
                    return harness.reads.get(timeout=0.02)
                except queue.Empty:
                    return b""

            def close(self):
                self.is_open = False

        self.modem = Modem()
        namespace.update(
            serial_obj=None, serial_read_lock=threading.Lock(), serial_stop_event=self.stop,
            serial_wakeup_event=threading.Event(), MODE="Manual", PORT="TEST", BAUD=115200,
            serial=SimpleNamespace(Serial=lambda *args, **kwargs: self.modem, SerialException=OSError),
            KEYWORDS=[], LOG_UNMATCHED_SMS=False, LOG_DIR="unused", LOG_PREFIX="TEST",
            ERROR_REPEAT_LIMIT=3, CALL_FILTER_MODE="Disabled", CALL_WHITELIST=[], CALL_BLACKLIST=[],
            FILE_LOG_Q=queue.Queue(), current_call_popup=None, _sms_ignore_repeat_state={},
            serial_running=True, RECONNECT_INTERVAL=1,
            CLOUD_DEVICE_IMEI="", cloud_imei_verified=False, cloud_imei_query_deadline=0.0,
            IMEI_REGEX=re.compile(r"\b(\d{14,17})\b"), _normalize_imei=normalize_imei,
            _cloud_identity_payload_core=identity_payload, APP_VERSION="4.0.2",
            _auto_connect_notice=SimpleNamespace(reset=lambda: None),
            _serial_error_notice=SimpleNamespace(clear=lambda: None),
            root=SimpleNamespace(after=lambda *args: None), ui_post=lambda callback: None,
            status_var=SimpleNamespace(get=lambda: "connecting"), APP_START_MONO=0.0, START_UI_DELAY=0.0,
            time=SimpleNamespace(monotonic=lambda: self.clock),
        )
        for name in (
            "enqueue_third_push", "_cloud_send_sms_event", "play_alert", "show_sms_popup", "system_ui",
            "set_temperature", "set_signal", "close_call_popup", "send_call_hangup_command",
            "show_call_popup", "serial_error_ui", "lock_port_mutex", "unlock_port_mutex",
            "_notify_cloud_identity_changed", "notify_cloud_channel_status", "_cloud_log", "set_status",
        ):
            namespace[name] = lambda *args, **kwargs: None
        namespace["port_ui"] = lambda text, *args: self.messages.append(text)
        namespace["set_status"] = lambda text, color: self.statuses.append((text, color))
        namespace["serial_error_ui"] = lambda text, *args, **kwargs: self.messages.append(text)
        namespace["_push_serial_debug"] = lambda text: self.lines.append(text)
        namespace["_cloud_send_serial_log"] = self.observe_line
        namespace["consume_call_recording_line"] = lambda line: self.recording_lines.append(line) or line.startswith("@@CALL_RECORD")
        namespace["_cloud_runtime_imei"] = lambda: cloud_runtime_imei_namespace_runtime(namespace)
        namespace["_cloud_identity_payload"] = lambda: cloud_identity_payload_namespace_runtime(namespace)
        namespace["_set_cloud_device_imei"] = lambda value, **kwargs: set_cloud_device_imei_namespace_runtime(namespace, value, **kwargs)
        namespace["_maybe_capture_cloud_device_imei"] = lambda line: maybe_capture_cloud_device_imei_namespace_runtime(namespace, line)
        namespace["set_serial_log_prefix"] = lambda value: namespace.update(LOG_PREFIX=value)
        namespace["get_serial_call_state"] = lambda: get_serial_call_state_namespace_runtime(namespace)
        namespace["set_serial_call_state"] = lambda *args: set_serial_call_state_namespace_runtime(namespace, *args)
        namespace["merge_serial_call_state"] = lambda *args: merge_serial_call_state_namespace_runtime(namespace, *args)
        namespace["resolve_serial_target_port"] = lambda: "TEST"
        namespace["open_and_initialize_serial"] = lambda port: open_and_initialize_serial_namespace_runtime(namespace, port)
        namespace["read_serial_line_safely"] = lambda **kwargs: read_serial_line_safely_runtime(
            namespace["serial_lock"], lambda: namespace["serial_obj"], OSError,
            read_lock=namespace["serial_read_lock"],
            **kwargs,
        )
        namespace["safe_close_serial"] = lambda: safe_close_serial_namespace_runtime(namespace)
        namespace["try_manual_rebind_after_error"] = lambda error: False
        namespace["schedule_delayed_connected_log"] = lambda port, baud: schedule_delayed_connected_log_namespace_runtime(
            namespace, port, baud, start_runtime=self.connected,
        )

    def connected(self, *args, **kwargs):
        self.connected_kwargs = kwargs
        self.ready.set()
        if self.connected_error:
            self.stop.set()
            raise RuntimeError("synthetic connected callback failure")

    def terminal(self, command, *, response="OK", extra=""):
        body = "[I]-[ril.proatc] 000000000000001\r\n" if command == "AT+CGSN" else ""
        self.reads.put((body + "[I]-[ril.proatc] " + response + "\r\n" + extra).encode("ascii"))

    def emit(self, text):
        self.reads.put(text.encode("ascii"))

    def observe_line(self, text, *args):
        self.lines.append(text)
        if self.on_line is not None:
            self.on_line(text)

    def __enter__(self):
        self.stack = ExitStack()
        self.fixture.open_window(self.stack)

        def reader():
            try:
                run_serial_reader_namespace_runtime(
                    self.ns, parse_callback_head=parse_sms_callback_head,
                    apply_disconnect_effects=lambda *args: None,
                )
            except Exception as exc:
                self.errors.append(type(exc).__name__)

        self.thread = threading.Thread(target=reader)
        self.thread.start()
        if not self.ready.wait(2):
            self.__exit__(None, None, None)
            raise RuntimeError("formal reader did not establish the connection")
        return self

    def __exit__(self, *_args):
        self.stop.set()
        self.ns["SERIAL_COMMAND_RESPONSE_COORDINATOR"].cancel_active("synthetic stop")
        self.thread.join(2)
        for worker in self.ns["SERIAL_COMMAND_THREAD_REGISTRY"].snapshot():
            worker.join(2)
        self.stack.close()
        if self.thread.is_alive():
            raise RuntimeError("formal reader did not stop")
