from dataclasses import dataclass
import time

from sms_core.call_effects import apply_call_decision, apply_ring_timeout_expired
from sms_core.call_events import CallState, handle_call_line, ring_timeout_expired
from sms_core.serial_line_effects import apply_serial_line_effects, push_serial_debug_insights
from sms_core.serial_sms import SmsPendingCollector, flush_pending_sms, handle_sms_collector_line


@dataclass
class SerialRuntimeState:
    sms_collector: SmsPendingCollector
    call_state: CallState

    @classmethod
    def create(cls, parse_callback_head):
        return cls(
            sms_collector=SmsPendingCollector(parse_callback_head),
            call_state=CallState(),
        )

    def sync_call_state(self, ring_timeout_target, current_dial_num):
        self.call_state.ring_timeout_target = ring_timeout_target
        self.call_state.current_dial_num = current_dial_num

    def reset_call_state(self):
        self.call_state = CallState()


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


def flush_runtime_pending_sms(state, config, callbacks, ignore_repeat_state):
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

    if not line:
        if state.sms_collector.expired(now):
            flush_runtime_pending_sms(state, config, callbacks, ignore_repeat_state)
        return SerialRuntimeResult(continue_read=True)

    apply_serial_line_effects(
        line,
        callbacks.push_serial_debug,
        callbacks.send_cloud_serial_log,
        callbacks.capture_cloud_device_imei,
        callbacks.set_temperature,
        callbacks.set_signal,
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
        lambda: flush_runtime_pending_sms(state, config, callbacks, ignore_repeat_state),
    )
    return SerialRuntimeResult(continue_read=sms_decision.continue_read)


def decode_serial_line(raw):
    return raw.decode("utf-8", "ignore").strip()


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
):
    while should_continue():
        target_port = get_target_port()
        try:
            target_port = resolve_target_port()
            if not target_port:
                continue

            set_connecting_status(target_port)
            open_and_initialize_serial(target_port)
            on_connected_port(target_port)

            while should_continue():
                handle_line(decode_serial_line(read_serial_line()))

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
