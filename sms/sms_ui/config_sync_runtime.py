import os
import time
from dataclasses import dataclass


@dataclass
class ConfigFileWatchState:
    signature: object = None
    after_id: object = None
    generation: int = 0


@dataclass
class ConfigReloadFailureLogState:
    consecutive_failures: int = 0
    suppressed_since_log: int = 0
    last_log_at: object = None


def _safe_emit_log(log_error, message):
    if log_error is None:
        return False
    try:
        log_error(message)
        return True
    except Exception:
        return False


def report_config_reload_failure_runtime(
    state,
    error_category,
    *,
    log_error=None,
    monotonic=time.monotonic,
    log_interval_seconds=60.0,
):
    state.consecutive_failures += 1
    now = monotonic()
    if state.last_log_at is None:
        state.last_log_at = now
        state.suppressed_since_log = 0
        return _safe_emit_log(log_error, f"Reload shared UI config failed ({error_category})")

    state.suppressed_since_log += 1
    if now - state.last_log_at < max(1.0, float(log_interval_seconds)):
        return False

    suppressed = state.suppressed_since_log
    state.last_log_at = now
    state.suppressed_since_log = 0
    return _safe_emit_log(
        log_error,
        f"Reload shared UI config still failing ({error_category}); "
        f"suppressed {suppressed} repeated attempts",
    )


def clear_config_reload_failure_runtime(state, *, log_error=None):
    failures = state.consecutive_failures
    if failures:
        _safe_emit_log(
            log_error,
            f"Reload shared UI config recovered after {failures} failed attempts",
        )
    state.consecutive_failures = 0
    state.suppressed_since_log = 0
    state.last_log_at = None
    return failures


def config_file_signature_runtime(config_file, *, stat_file=os.stat):
    try:
        stat_result = stat_file(config_file)
    except OSError:
        return None
    mtime_ns = getattr(stat_result, "st_mtime_ns", None)
    if mtime_ns is None:
        mtime_ns = int(stat_result.st_mtime * 1_000_000_000)
    return (
        mtime_ns,
        stat_result.st_size,
        getattr(stat_result, "st_ino", None),
    )


def schedule_config_file_watch_runtime(
    *,
    state,
    config_file,
    interval_ms,
    root_after,
    root_after_cancel,
    tk_alive,
    is_stopping,
    on_change,
    signature_func=config_file_signature_runtime,
    log_error=None,
):
    state.generation += 1
    generation = state.generation

    if state.after_id is not None:
        try:
            root_after_cancel(state.after_id)
        except Exception:
            pass
        state.after_id = None

    state.signature = signature_func(config_file)

    def safe_log(message):
        if log_error is None:
            return
        try:
            log_error(message)
        except Exception:
            pass

    def schedule_next():
        if generation != state.generation or is_stopping() or not tk_alive():
            state.after_id = None
            return
        try:
            state.after_id = root_after(interval_ms, tick)
        except Exception as exc:
            state.after_id = None
            safe_log(f"Schedule config file watch failed: {exc!r}")

    def tick():
        if generation != state.generation or is_stopping() or not tk_alive():
            state.after_id = None
            return
        state.after_id = None
        try:
            next_signature = signature_func(config_file)
            if next_signature != state.signature:
                if next_signature is None:
                    state.signature = None
                else:
                    result = on_change()
                    if result is not False:
                        state.signature = next_signature
        except Exception as exc:
            safe_log(f"Config file watch tick failed: {exc!r}")
        finally:
            schedule_next()

    schedule_next()
    return state.after_id


def stop_config_file_watch_runtime(state, *, root_after_cancel=None):
    state.generation += 1
    after_id = state.after_id
    state.after_id = None
    if after_id is not None and root_after_cancel is not None:
        try:
            root_after_cancel(after_id)
        except Exception:
            pass
    return after_id
