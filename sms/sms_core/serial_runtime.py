from contextlib import nullcontext
from dataclasses import dataclass, replace
import codecs
import inspect
import re
import time

from sms_core.call_effects import apply_call_decision, apply_ring_timeout_expired
from sms_core.call_command_context import SerialCallStateSnapshot, call_with_state_snapshot
from sms_core.call_events import (
    CallLineDecision,
    CallState,
    handle_call_line,
    line_confirms_call_presence,
    ring_timeout_expired,
)
from sms_core.long_sms_assembler import LongSmsAssembler
from sms_core.serial_line_effects import apply_serial_line_effects, push_serial_debug_insights
from sms_core.serial_sms import (
    SmsPendingCollector,
    flush_pending_sms,
    handle_sms_collector_line,
    process_pending_sms_items,
)
from sms_core.serial_sms_pdu_cache import SmsPduCorrectionCache
from sms_core.sms_receive_pipeline import SmsReceivePipeline


@dataclass
class SerialRuntimeState:
    sms_collector: SmsPendingCollector
    sms_pdu_cache: SmsPduCorrectionCache
    long_sms_assembler: LongSmsAssembler
    sms_pipeline: SmsReceivePipeline
    call_state: CallState
    call_state_snapshot: object = None
    received_call_snapshot: object = None
    dial_generation: object = None

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

    def sync_call_snapshot(self, snapshot):
        if isinstance(snapshot, SerialCallStateSnapshot):
            if snapshot.generation is not self.dial_generation:
                if self.call_state.call_session_id.startswith("outgoing:"):
                    self.call_state.call_session_id = ""
                self.call_state.outgoing_connected_session_id = ""
                self.dial_generation = snapshot.generation
            self.call_state_snapshot = snapshot
        else:
            self.call_state_snapshot = None
        self.sync_call_state(*snapshot)

    def reset_call_state(self):
        self.call_state = CallState()

    def reset_sms_state(self):
        self.sms_collector.reset()
        self.sms_pipeline.reset()


@dataclass(frozen=True)
class SerialRuntimeResult:
    continue_read: bool = False
    incoming_state_committed: bool = False
    call_state_committed: bool = False


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
    send_cloud_call_event: object = lambda *_args, **_kwargs: None
    send_cloud_call_state: object = lambda *_args, **_kwargs: None
    set_local_number: object = lambda *_args: None
    observe_sms_send_line: object = lambda *_args: None
    start_incoming_call: object = lambda _caller_num: True
    finish_incoming_call: object = lambda: None
    reset_incoming_call: object = lambda: None
    show_missed_call_popup: object = lambda _missed_call: None
    consume_call_recording_line: object = lambda _line: False
    mark_dial_connected: object = lambda: None
    mark_call_connected: object = lambda: None
    finish_dial_call: object = lambda _message="": None
    consume_remote_incoming_end: object = lambda _session_id: False
    call_state_guard: object = nullcontext
    observe_call_command_line: object = lambda _line: None
    start_serial_incoming_call: object = None


def build_sms_diagnostic_log(config, callbacks, now_func=None):
    def log(message, tag="normal"):
        _push_sms_diagnostic_debug(callbacks.push_serial_debug, message)

    return log


def _push_sms_diagnostic_debug(push_serial_debug, message):
    try:
        push_serial_debug(message)
    except Exception:
        pass


def _mark_call_connected(callback, call_session_id=""):
    """Notify the UI and keep compatibility with older zero-argument bindings."""
    session_id = str(call_session_id or "").strip()
    if not session_id:
        return callback()
    try:
        signature = inspect.signature(callback)
        signature.bind(session_id)
    except (TypeError, ValueError):
        return callback()
    return callback(session_id)


MAX_SERIAL_LINE_CHARS = 256 * 1024


class SerialLineDecoder:
    def __init__(self, encoding="utf-8", *, max_line_chars=MAX_SERIAL_LINE_CHARS):
        self.encoding = encoding
        self.decoder = codecs.getincrementaldecoder(encoding)("replace")
        self.text_buffer = ""
        self.max_line_chars = max(1, int(max_line_chars))
        self._skip_lf = False
        self.call_state_snapshot = None

    def feed(self, raw):
        snapshot = getattr(raw, "call_state_snapshot", None)

        def decoded_line(text, captured):
            return SerialReceivedLine(text, captured) if captured is not None else text

        if not raw:
            # A lone > at a chunk boundary may still be SMS body text. Wait
            # for one idle read before treating an unterminated > as a prompt.
            if self.text_buffer.strip() == ">" and not self.decoder.getstate()[0]:
                self.text_buffer = ""
                captured = self.call_state_snapshot
                self.call_state_snapshot = None
                return [decoded_line(">", captured)]
            if self.text_buffer or self.decoder.getstate()[0]:
                return []
            return [decoded_line("", snapshot)]

        lines = []
        # Chunk boundaries are unrelated to frame boundaries. Preserve split
        # CRLF and the firmware's artificial newline inside a UTF-8 character.
        for index, part in enumerate(re.split(rb"(\r\n|[\r\n])", bytes(raw))):
            if index % 2 == 0:
                if not part:
                    continue
                if not self.text_buffer and not self.decoder.getstate()[0]:
                    self.call_state_snapshot = snapshot
                self._skip_lf = False
                self.text_buffer += self.decoder.decode(part, final=False)
                if len(self.text_buffer) > self.max_line_chars:
                    self.text_buffer = ""
                    self.decoder.reset()
                    raise ValueError("串口单行数据过长，已停止当前读取，请检查串口和波特率")
            else:
                if part == b"\n" and self._skip_lf:
                    self._skip_lf = False
                    continue
                self._skip_lf = part == b"\r"
                if not self.decoder.getstate()[0]:
                    lines.append(decoded_line(self.text_buffer.strip(), self.call_state_snapshot or snapshot))
                    self.text_buffer = ""
                    self.call_state_snapshot = None
        return lines


class SerialReadChunk(bytes):
    def __new__(cls, raw, snapshot):
        result = super().__new__(cls, raw or b"")
        result.call_state_snapshot = snapshot
        return result


class SerialReceivedLine(str):
    def __new__(cls, line, snapshot):
        result = super().__new__(cls, line)
        result.call_state_snapshot = snapshot
        return result


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


def process_runtime_sms_items(pending, config, callbacks, ignore_repeat_state):
    return process_pending_sms_items(
        pending,
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
    )


def _consume_remote_incoming_end(state, callbacks):
    call_state = state.call_state
    consume_remote_end = getattr(callbacks, "consume_remote_incoming_end", None)
    if (
        call_state.call_session_id and not call_state.current_dial_num
        and not call_state.call_session_id.startswith("outgoing:")
        and callable(consume_remote_end)
        and consume_remote_end(call_state.call_session_id)
    ):
        # Only the serial reader owns CallState. A confirmed peer terminal
        # names the generation to clear before a new RING/CLIP is parsed.
        call_state.ring_timeout_target = 0.0
        call_state.last_clip_num = ""
        call_state.last_clip_time = 0.0
        call_state.call_session_id = ""
        call_state.pending_incoming_connected_until = 0.0
        call_state.incoming_connected_session_id = ""
        return True
    return False


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
    snapshot = state.call_state_snapshot
    call_with_state_snapshot(
        getattr(callbacks, "observe_call_command_line", lambda _line: None), line,
        call_state_snapshot=snapshot,
    )
    if _consume_remote_incoming_end(state, callbacks):
        popup_active = False
    try:
        if callbacks.consume_call_recording_line(line):
            return SerialRuntimeResult(continue_read=True)
    except Exception:
        return SerialRuntimeResult(continue_read=True)

    if (
        ring_timeout_expired(state.call_state.ring_timeout_target, now)
        and not line_confirms_call_presence(line)
    ):
        timed_out_number = state.call_state.last_clip_num
        timed_out_session_id = state.call_state.call_session_id
        state.call_state.ring_timeout_target = 0.0
        state.call_state.last_clip_num = ""
        state.call_state.call_session_id = ""
        state.call_state.pending_incoming_connected_until = 0.0
        state.call_state.incoming_connected_session_id = ""
        if timed_out_number:
            try:
                kwargs = {"direction": "incoming"}
                if timed_out_session_id:
                    kwargs["call_session_id"] = timed_out_session_id
                callbacks.send_cloud_call_state(timed_out_number, "no_answer", "RING TIMEOUT", **kwargs)
            except Exception:
                pass
        apply_ring_timeout_expired(
            current_port,
            callbacks.port_ui,
            callbacks.set_status,
            callbacks.close_call_popup,
            callbacks.finish_incoming_call,
            callbacks.show_missed_call_popup,
        )
        popup_active = False

    sms_diagnostic_log = build_sms_diagnostic_log(config, callbacks)
    ready_sms = state.sms_pipeline.observe_line(line, now, log=sms_diagnostic_log)
    if ready_sms is not None:
        process_runtime_sms_items(ready_sms, config, callbacks, ignore_repeat_state)
    try:
        call_with_state_snapshot(
            callbacks.observe_sms_send_line, line,
            call_state_snapshot=state.received_call_snapshot or snapshot, check_current=False,
        )
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

    incoming_state_committed = False
    call_state_committed = False
    start_result = True
    with getattr(callbacks, "call_state_guard", nullcontext)():
        # A peer terminal can arrive during SMS/log callbacks above. Consume
        # its generation and commit the next pure call transition atomically.
        if _consume_remote_incoming_end(state, callbacks):
            popup_active = False
        if snapshot is not None and not snapshot.is_current():
            call_decision = CallLineDecision(state=state.call_state)
            call_state_committed = True
        else:
            decision_state = state.call_state
            if snapshot is not None and snapshot.pending:
                decision_state = replace(decision_state, current_dial_num="")
            call_decision = handle_call_line(
                line, decision_state, now, config.call_filter_mode,
                config.call_whitelist, config.call_blacklist, popup_active,
            )
        state.call_state = call_decision.state
        if call_decision.incoming_number:
            start_serial_incoming = getattr(callbacks, "start_serial_incoming_call", None)
            if callable(start_serial_incoming):
                start_result = start_serial_incoming(
                    call_decision.incoming_number, call_decision.call_session_id,
                    state.call_state.ring_timeout_target,
                )
                incoming_state_committed = bool(start_result)
            else:
                start_result = callbacks.start_incoming_call(call_decision.incoming_number)
        elif snapshot is not None and snapshot[1]:
            # Commit the original outgoing transition while its generation is
            # still protected; later effects may yield to a same-number redial.
            if not call_state_committed:
                snapshot.commit(state.call_state.ring_timeout_target, state.call_state.current_dial_num)
            call_state_committed = True

    call_effect = apply_call_decision(
        call_decision,
        current_port,
        callbacks.send_call_hangup,
        callbacks.enqueue_third_push,
        callbacks.port_ui,
        lambda text, color: call_with_state_snapshot(
            callbacks.set_status, text, color,
            call_state_snapshot=snapshot if snapshot is not None and snapshot[1] else None),
        callbacks.show_call_popup,
        callbacks.close_call_popup,
        lambda _number: start_result,
        callbacks.finish_incoming_call,
        callbacks.show_missed_call_popup,
        callbacks.send_cloud_call_event,
        callbacks.send_cloud_call_state,
        lambda message: call_with_state_snapshot(
            callbacks.finish_dial_call, message, call_state_snapshot=snapshot, check_incoming=False),
    )
    if call_effect.stop_processing:
        return SerialRuntimeResult(
            continue_read=True, incoming_state_committed=incoming_state_committed,
            call_state_committed=call_state_committed)

    if call_decision.connected_number:
        try:
            call_with_state_snapshot(callbacks.mark_dial_connected, call_state_snapshot=snapshot)
        except Exception:
            pass
    if call_decision.incoming_connected_number:
        try:
            _mark_call_connected(
                callbacks.mark_call_connected,
                call_decision.call_session_id,
            )
        except Exception:
            pass

    push_serial_debug_insights(line, callbacks.push_serial_debug)

    sms_decision = handle_sms_collector_line(
        state.sms_collector,
        line,
        now,
        lambda: flush_runtime_pending_sms(state, config, callbacks, ignore_repeat_state, now=now),
    )
    return SerialRuntimeResult(
        continue_read=sms_decision.continue_read,
        incoming_state_committed=incoming_state_committed,
        call_state_committed=call_state_committed,
    )


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
    is_stopping=lambda: False,
):
    try:
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
                if is_stopping():
                    break
                if handle_error(e, target_port):
                    continue
                wait_before_retry()
    finally:
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
    settle_wait=time.sleep,
    merge_call_state=None,
):
    state = SerialRuntimeState.create(parse_callback_head)
    now = clock or time.monotonic
    set_call_state(0.0, "")
    callbacks.reset_incoming_call()
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
        received_state = getattr(line, "call_state_snapshot", None)
        line = str(line)
        previous_state = get_call_state()
        if isinstance(received_state, SerialCallStateSnapshot):
            # Earlier lines from the same read may have advanced this call.
            # Refresh values only within the originally received generation.
            if not received_state.same_generation(previous_state):
                previous_state = received_state
        config = get_runtime_config()
        state.sync_call_snapshot(previous_state)
        state.received_call_snapshot = (
            received_state if isinstance(received_state, SerialCallStateSnapshot)
            else state.call_state_snapshot)
        result = handle_runtime_line(
            state,
            line,
            now(),
            get_target_port(),
            popup_active(),
            config,
            callbacks,
            ignore_repeat_state,
        )
        if (getattr(result, "incoming_state_committed", False)
                or getattr(result, "call_state_committed", False)):
            # The new tracker lifetime and shared values were committed under
            # the peer lock. Preserve any answer/terminal received afterwards.
            state.sync_call_snapshot(get_call_state())
        elif merge_call_state is None:
            sync_app_call_state()
        else:
            # UI/cloud callbacks may update the shared state while this line
            # is being processed. Commit only against the snapshot we read.
            merged = merge_call_state(
                previous_state, state.call_state.ring_timeout_target,
                state.call_state.current_dial_num,
            )
            state.sync_call_snapshot(merged)

    def read_call_scoped_chunk():
        raw = read_serial_line()
        return SerialReadChunk(raw, get_call_state())

    def finish_sms_session():
        # Flush while the original device identity and log prefix still apply.
        flush_runtime_pending_sms(state, config, callbacks, ignore_repeat_state, now=now())
        ready = state.long_sms_assembler.finish_session(
            now=now(), wait=settle_wait, log=build_sms_diagnostic_log(config, callbacks),
        )
        process_runtime_sms_items(ready, config, callbacks, ignore_repeat_state)
        state.reset_sms_state()

    def handle_error(error, target_port):
        finish_sms_session()
        state.reset_call_state()
        callbacks.reset_incoming_call()
        sync_app_call_state()
        return handle_disconnect(error, target_port)

    try:
        run_loop(
            should_continue=should_continue,
            get_target_port=get_target_port,
            resolve_target_port=resolve_target_port,
            set_connecting_status=set_connecting_status,
            open_and_initialize_serial=open_and_initialize_serial,
            on_connected_port=handle_connected_port,
            read_serial_line=read_call_scoped_chunk,
            handle_line=handle_line,
            handle_error=handle_error,
            wait_before_retry=wait_before_retry,
            safe_close_serial=safe_close_serial,
            is_stopping=lambda: not should_continue(),
        )
    finally:
        finish_sms_session()
    return state
