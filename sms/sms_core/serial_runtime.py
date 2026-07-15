from dataclasses import dataclass
import codecs
import time

from sms_core.call_effects import apply_call_decision, apply_ring_timeout_expired
from sms_core.call_events import CallState, handle_call_line, ring_timeout_expired
from sms_core.long_sms_assembler import LongSmsAssembler
from sms_core.serial_line_effects import apply_serial_line_effects, push_serial_debug_insights
from sms_core.serial_sms import SmsPendingCollector, flush_pending_sms, handle_sms_collector_line
from sms_core.serial_sms_pdu_cache import SmsPduCorrectionCache
from sms_core.sms_receive_pipeline import SmsReceivePipeline


@dataclass
class SerialRuntimeState:
    sms_collector: SmsPendingCollector
    sms_pdu_cache: SmsPduCorrectionCache
    long_sms_assembler: LongSmsAssembler
    sms_pipeline: SmsReceivePipeline
    call_state: CallState

    @classmethod
    def create(cls, parse_callback_head):
        sms_pdu_cache = SmsPduCorrectionCache()
        long_sms_assembler = LongSmsAssembler(parse_callback_head)
        sms_pipeline = SmsReceivePipeline(parse_callback_head, sms_pdu_cache, long_sms_assembler)
        return cls(
            sms_collector=SmsPendingCollector(parse_callback_head),
            sms_pdu_cache=sms_pdu_cache,
            long_sms_assembler=long_sms_assembler,
            sms_pipeline=sms_pipeline,
            call_state=CallState(),
        )

    def sync_call_state(self, ring_timeout_target, current_dial_num):
        self.call_state.ring_timeout_target = ring_timeout_target
        self.call_state.current_dial_num = current_dial_num

    def reset_call_state(self):
        self.call_state = CallState()

    def reset_sms_state(self):
        self.sms_collector.reset()
        self.sms_pipeline.reset()


@dataclass(frozen=True)
class SerialRuntimeResult:
    continue_read: bool = False


@dataclass(frozen=True)
class SerialRuntimeConfig:
    keywords: list
    log_unmatched_sms: bool
    log_dir: str
    log_prefix: str
    error_repeat_limit: int
    call_filter_mode: str
    call_whitelist: list
    call_blacklist: list


@dataclass(frozen=True)
class SerialRuntimeCallbacks:
    enqueue_third_push: object
    send_cloud_sms_event: object
    port_ui: object
    play_alert: object
    show_sms_popup: object
    file_log: object
    system_ui: object
    push_serial_debug: object
    send_cloud_serial_log: object
    capture_cloud_device_imei: object
    set_temperature: object
    set_signal: object
    set_status: object
    close_call_popup: object
    send_call_hangup: object
    show_call_popup: object
    set_local_number: object = lambda *_args: None
    observe_sms_send_line: object = lambda *_args: None


def build_sms_diagnostic_log(config, callbacks, now_func=None):
    def log(message, tag="normal"):
        _push_sms_diagnostic_debug(callbacks.push_serial_debug, message)

    return log


def _push_sms_diagnostic_debug(push_serial_debug, message):
    try:
        push_serial_debug(message)
    except Exception:
        pass


class SerialLineDecoder:
    def __init__(self, encoding="utf-8"):
        self.encoding = encoding
        self.decoder = codecs.getincrementaldecoder(encoding)("replace")
        self.text_buffer = ""

    def _ends_with_incomplete_character(self, raw):
        if not raw:
            return False
        try:
            raw.decode(self.encoding, "strict")
            return False
        except UnicodeDecodeError as exc:
            return exc.reason == "unexpected end of data" and exc.end == len(raw)

    def _drop_artificial_newline_after_incomplete_character(self, raw):
        for suffix in (b"\r\n", b"\n", b"\r"):
            if raw.endswith(suffix):
                body = raw[:-len(suffix)]
                if self._ends_with_incomplete_character(body):
                    return body
                break
        return raw

    def feed(self, raw):
        if not raw:
            if self.text_buffer or self.decoder.getstate()[0]:
                return []
            return [""]

        raw = self._drop_artificial_newline_after_incomplete_character(bytes(raw))
        self.text_buffer += self.decoder.decode(raw, final=False)
        if self.text_buffer.strip() == ">":
            self.text_buffer = ""
            return [">"]
        lines = []
        pending = []
        for part in self.text_buffer.splitlines(keepends=True):
            if part.endswith(("\r", "\n")):
                lines.append(part.rstrip("\r\n").strip())
            else:
                pending.append(part)
        self.text_buffer = "".join(pending)
        return lines


def flush_runtime_pending_sms(state, config, callbacks, ignore_repeat_state, now=None):
    sms_diagnostic_log = build_sms_diagnostic_log(config, callbacks)
    return flush_pending_sms(
        state.sms_collector,
        config.keywords,
        config.log_unmatched_sms,
        config.log_dir,
        config.log_prefix,
        ignore_repeat_state,
        config.error_repeat_limit,
        callbacks.enqueue_third_push,
        callbacks.send_cloud_sms_event,
        callbacks.port_ui,
        callbacks.play_alert,
        callbacks.show_sms_popup,
        callbacks.file_log,
        callbacks.system_ui,
        assembler=state.sms_pipeline,
        now=now,
        concat_log=sms_diagnostic_log,
    )


def handle_serial_runtime_line(
    state,
    line,
    now,
    current_port,
    popup_active,
    config,
    callbacks,
    ignore_repeat_state,
):
    if ring_timeout_expired(state.call_state.ring_timeout_target, now):
        state.call_state.ring_timeout_target = 0.0
        state.call_state.last_clip_num = ""
        apply_ring_timeout_expired(
            current_port,
            callbacks.port_ui,
            callbacks.set_status,
            callbacks.close_call_popup,
        )

    sms_diagnostic_log = build_sms_diagnostic_log(config, callbacks)
    state.sms_pipeline.observe_line(line, now, log=sms_diagnostic_log)
    try:
        callbacks.observe_sms_send_line(line)
    except Exception:
        pass

    if not line:
        if state.sms_collector.expired(now):
            flush_runtime_pending_sms(state, config, callbacks, ignore_repeat_state, now=now)
        return SerialRuntimeResult(continue_read=True)

    apply_serial_line_effects(
        line,
        callbacks.push_serial_debug,
        callbacks.send_cloud_serial_log,
        callbacks.capture_cloud_device_imei,
        callbacks.set_temperature,
        callbacks.set_signal,
        callbacks.set_local_number,
    )

    call_decision = handle_call_line(
        line,
        state.call_state,
        now,
        config.call_filter_mode,
        config.call_whitelist,
        config.call_blacklist,
        popup_active,
    )
    state.call_state = call_decision.state

    call_effect = apply_call_decision(
        call_decision,
        current_port,
        callbacks.send_call_hangup,
        callbacks.enqueue_third_push,
        callbacks.port_ui,
        callbacks.set_status,
        callbacks.show_call_popup,
        callbacks.close_call_popup,
    )
    if call_effect.stop_processing:
        return SerialRuntimeResult(continue_read=True)

    push_serial_debug_insights(line, callbacks.push_serial_debug)

    sms_decision = handle_sms_collector_line(
        state.sms_collector,
        line,
        now,
        lambda: flush_runtime_pending_sms(state, config, callbacks, ignore_repeat_state, now=now),
    )
    return SerialRuntimeResult(continue_read=sms_decision.continue_read)


def decode_serial_line(raw):
    return raw.decode("utf-8", "replace").strip()


def run_serial_thread_loop(
    *,
    should_continue,
    get_target_port,
    resolve_target_port,
    set_connecting_status,
    open_and_initialize_serial,
    on_connected_port,
    read_serial_line,
    handle_line,
    handle_error,
    wait_before_retry,
    safe_close_serial,
    monotonic=time.monotonic,
    empty_target_min_delay=0.05,
):
    while should_continue():
        target_port = get_target_port()
        try:
            resolve_started = monotonic()
            target_port = resolve_target_port()
            if not target_port:
                try:
                    resolve_elapsed = monotonic() - resolve_started
                except Exception:
                    resolve_elapsed = 0.0
                if resolve_elapsed < float(empty_target_min_delay):
                    wait_before_retry()
                continue

            set_connecting_status(target_port)
            open_and_initialize_serial(target_port)
            on_connected_port(target_port)
            line_decoder = SerialLineDecoder()

            while should_continue():
                for line in line_decoder.feed(read_serial_line()):
                    handle_line(line)

        except Exception as e:
            if handle_error(e, target_port):
                continue
            wait_before_retry()

    safe_close_serial()


def run_serial_runtime_thread(
    *,
    parse_callback_head,
    get_runtime_config,
    callbacks,
    get_call_state,
    set_call_state,
    popup_active,
    ignore_repeat_state,
    should_continue,
    get_target_port,
    resolve_target_port,
    set_connecting_status,
    open_and_initialize_serial,
    on_connected_port,
    read_serial_line,
    handle_disconnect,
    wait_before_retry,
    safe_close_serial,
    clock=None,
    run_loop=run_serial_thread_loop,
    handle_runtime_line=handle_serial_runtime_line,
):
    state = SerialRuntimeState.create(parse_callback_head)
    now = clock or time.monotonic
    set_call_state(0.0, "")
    config = get_runtime_config()

    def handle_connected_port(target_port):
        nonlocal config
        on_connected_port(target_port)
        config = get_runtime_config()

    def sync_app_call_state():
        set_call_state(
            state.call_state.ring_timeout_target,
            state.call_state.current_dial_num,
        )

    def handle_line(line):
        nonlocal config
        config = get_runtime_config()
        ring_timeout_target, current_dial_num = get_call_state()
        state.sync_call_state(ring_timeout_target, current_dial_num)
        handle_runtime_line(
            state,
            line,
            now(),
            get_target_port(),
            popup_active(),
            config,
            callbacks,
            ignore_repeat_state,
        )
        sync_app_call_state()

    def handle_error(error, target_port):
        state.reset_sms_state()
        state.reset_call_state()
        sync_app_call_state()
        return handle_disconnect(error, target_port)

    run_loop(
        should_continue=should_continue,
        get_target_port=get_target_port,
        resolve_target_port=resolve_target_port,
        set_connecting_status=set_connecting_status,
        open_and_initialize_serial=open_and_initialize_serial,
        on_connected_port=handle_connected_port,
        read_serial_line=read_serial_line,
        handle_line=handle_line,
        handle_error=handle_error,
        wait_before_retry=wait_before_retry,
        safe_close_serial=safe_close_serial,
    )
    return state
