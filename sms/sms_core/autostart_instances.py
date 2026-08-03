import hashlib
import json
import os
import threading
from dataclasses import dataclass

from sms_core.app_launch import (
    get_clean_restart_env,
    get_launch_target_and_args,
    launch_detached_process,
)
from sms_core.windows_runtime import (
    acquire_mutex_with_error,
    acquire_named_mutex_lock,
    close_windows_handle,
    is_existing_instance_error,
    release_named_mutex_lock,
)


AUTOSTART_STATE_FILENAME = "autostart_instances.json"
AUTOSTART_STATE_MUTEX_PREFIX = "Air724UG_SMS_Autostart_State_V1"
AUTOSTART_LAUNCHER_MUTEX_PREFIX = "Air724UG_SMS_Autostart_Launcher_V1"
AUTOSTART_CHILD_FLAG = "--autostart-child"
# Keep automatic restoration aligned with the per-process instance-number
# boundary.  This is a technical guard against invalid state, not a smaller
# product limit on how many running instances may be restored after login.
MAX_AUTOSTART_INSTANCES = 9999
AUTOSTART_LAUNCH_INTERVAL_SECONDS = 1.0
AUTOSTART_LAUNCH_RETRY_COUNT = 3


@dataclass(frozen=True)
class AutostartRegistrationResult:
    desired_count: int
    registered: bool
    active_count: int = 1


def get_autostart_state_path(app_dir):
    return os.path.join(os.path.abspath(str(app_dir)), AUTOSTART_STATE_FILENAME)


def autostart_state_mutex_name(app_dir):
    normalized = os.path.normcase(os.path.abspath(str(app_dir)))
    digest = hashlib.sha256(normalized.encode("utf-8", "surrogatepass")).hexdigest()[:16]
    return f"{AUTOSTART_STATE_MUTEX_PREFIX}_{digest}"


def autostart_launcher_mutex_name(app_dir):
    normalized = os.path.normcase(os.path.abspath(str(app_dir)))
    digest = hashlib.sha256(normalized.encode("utf-8", "surrogatepass")).hexdigest()[:16]
    return f"{AUTOSTART_LAUNCHER_MUTEX_PREFIX}_{digest}"


def claim_autostart_launcher(
    app_dir,
    *,
    acquire_mutex=acquire_mutex_with_error,
    close_handle=close_windows_handle,
    existing_error=is_existing_instance_error,
    log_error=None,
):
    try:
        handle, last_error = acquire_mutex(autostart_launcher_mutex_name(app_dir))
    except Exception as exc:
        _safe_log(log_error, f"Claim autostart launcher mutex failed: {exc!r}")
        return None

    if existing_error(last_error):
        if handle:
            try:
                close_handle(handle)
            except Exception as exc:
                _safe_log(log_error, f"Close occupied autostart launcher mutex failed: {exc!r}")
        return None

    if not handle:
        _safe_log(
            log_error,
            f"Claim autostart launcher mutex failed with error code {last_error}",
        )
        return None
    return handle


def normalize_desired_count(value, *, maximum=MAX_AUTOSTART_INSTANCES):
    try:
        count = int(value)
    except (TypeError, ValueError):
        count = 1
    return min(max(1, count), max(1, int(maximum)))


def _safe_log(log_error, message):
    if log_error is None:
        return
    try:
        log_error(message)
    except Exception:
        pass


def _load_state(state_path, *, open_file=open, log_error=None):
    try:
        with open_file(state_path, "r", encoding="utf-8") as file_obj:
            raw = json.load(file_obj)
    except FileNotFoundError:
        return {"version": 1, "desired_count": 1, "active_instances": []}
    except Exception as exc:
        _safe_log(log_error, f"Load autostart instance state failed: {exc!r}")
        return {"version": 1, "desired_count": 1, "active_instances": []}

    if not isinstance(raw, dict):
        return {"version": 1, "desired_count": 1, "active_instances": []}

    raw_active = raw.get("active_instances", [])
    if not isinstance(raw_active, (list, tuple)):
        raw_active = []

    active = []
    for value in raw_active:
        try:
            number = int(value)
        except (TypeError, ValueError):
            continue
        if 1 <= number <= 9999 and number not in active:
            active.append(number)
        if len(active) >= MAX_AUTOSTART_INSTANCES:
            break
    return {
        "version": 1,
        "desired_count": normalize_desired_count(raw.get("desired_count", 1)),
        "active_instances": sorted(active),
    }


def _save_state(
    state_path,
    state,
    *,
    open_file=open,
    replace_file=os.replace,
    make_dirs=os.makedirs,
    path_exists=os.path.exists,
    remove_file=os.remove,
):
    state_dir = os.path.dirname(os.path.abspath(state_path))
    make_dirs(state_dir, exist_ok=True)
    tmp_path = f"{state_path}.{os.getpid()}.{threading.get_ident()}.tmp"
    try:
        with open_file(tmp_path, "w", encoding="utf-8") as file_obj:
            json.dump(state, file_obj, ensure_ascii=True, sort_keys=True)
        replace_file(tmp_path, state_path)
    finally:
        try:
            if path_exists(tmp_path):
                remove_file(tmp_path)
        except Exception:
            pass


def _reconcile_active_instances(active_instances, is_instance_active, *, log_error=None):
    reconciled = []
    for number in active_instances:
        try:
            active = bool(is_instance_active(number))
        except Exception as exc:
            _safe_log(log_error, f"Probe app instance {number} failed: {exc!r}")
            active = True
        if active:
            reconciled.append(number)
    return reconciled


def _with_state_lock(
    app_dir,
    callback,
    *,
    acquire_lock=acquire_named_mutex_lock,
    release_lock=release_named_mutex_lock,
    lock_timeout_ms=10000,
):
    handle, result = acquire_lock(
        autostart_state_mutex_name(app_dir),
        timeout_ms=lock_timeout_ms,
    )
    if not handle:
        raise RuntimeError(f"autostart state lock acquisition failed: {result}")
    try:
        return callback()
    finally:
        release_lock(handle)


def register_autostart_instance(
    *,
    app_dir,
    state_path,
    instance_number,
    allow_multi_instance,
    is_instance_active,
    log_error=None,
    load_state=_load_state,
    save_state=_save_state,
    with_state_lock=_with_state_lock,
):
    number = max(1, int(instance_number or 1))

    def update():
        state = load_state(state_path, log_error=log_error)
        previous_desired = normalize_desired_count(state.get("desired_count", 1))
        active = _reconcile_active_instances(
            state.get("active_instances", []),
            is_instance_active,
            log_error=log_error,
        )
        if number not in active:
            active.append(number)
            active.sort()

        if not allow_multi_instance:
            desired = 1
        else:
            # Starting a process must never lower the remembered target. Only
            # a clean unregister represents an intentional instance removal.
            desired = normalize_desired_count(max(previous_desired, len(active)))

        save_state(
            state_path,
            {
                "version": 1,
                "desired_count": desired,
                "active_instances": active,
            },
        )
        return AutostartRegistrationResult(desired, True, len(active))

    try:
        return with_state_lock(app_dir, update)
    except Exception as exc:
        _safe_log(log_error, f"Register autostart instance state failed: {exc!r}")
        try:
            fallback = load_state(state_path, log_error=log_error)
            desired = normalize_desired_count(fallback.get("desired_count", 1))
        except Exception:
            desired = 1
        return AutostartRegistrationResult(desired, False, 1)


def unregister_autostart_instance(
    *,
    app_dir,
    state_path,
    instance_number,
    is_instance_active,
    log_error=None,
    load_state=_load_state,
    save_state=_save_state,
    with_state_lock=_with_state_lock,
):
    number = max(1, int(instance_number or 1))

    def update():
        state = load_state(state_path, log_error=log_error)
        remaining = [
            value
            for value in state.get("active_instances", [])
            if value != number
        ]
        active = _reconcile_active_instances(
            remaining,
            is_instance_active,
            log_error=log_error,
        )
        desired = normalize_desired_count(len(active))
        save_state(
            state_path,
            {
                "version": 1,
                "desired_count": desired,
                "active_instances": active,
            },
        )
        return desired

    try:
        return with_state_lock(app_dir, update)
    except Exception as exc:
        _safe_log(log_error, f"Unregister autostart instance state failed: {exc!r}")
        return None


def get_active_autostart_instance_count(
    *,
    app_dir,
    state_path,
    is_instance_active,
    log_error=None,
    load_state=_load_state,
    with_state_lock=_with_state_lock,
):
    def read():
        state = load_state(state_path, log_error=log_error)
        active = _reconcile_active_instances(
            state.get("active_instances", []),
            is_instance_active,
            log_error=log_error,
        )
        return max(1, len(active))

    try:
        return with_state_lock(app_dir, read)
    except Exception as exc:
        _safe_log(log_error, f"Read active autostart instance count failed: {exc!r}")
        return 1


def build_autostart_child_command(
    autostart_flag,
    child_flag=AUTOSTART_CHILD_FLAG,
    *,
    launch_target_func=get_launch_target_and_args,
):
    target, script_arg, workdir = launch_target_func()
    command = [target]
    if script_arg:
        command.append(script_arg)
    command.extend((autostart_flag, child_flag))
    return command, workdir


def launch_autostart_companions(
    *,
    desired_count,
    allow_multi_instance,
    is_leader,
    autostart_flag,
    child_flag=AUTOSTART_CHILD_FLAG,
    active_instance_count=1,
    wait_before_launch=lambda seconds: False,
    interval_seconds=AUTOSTART_LAUNCH_INTERVAL_SECONDS,
    retry_count=AUTOSTART_LAUNCH_RETRY_COUNT,
    prepare_launch=build_autostart_child_command,
    launch_process=launch_detached_process,
    clean_env=get_clean_restart_env,
    log_error=None,
):
    target_count = normalize_desired_count(desired_count)
    if not allow_multi_instance or not is_leader:
        return 0

    def read_active_count(fallback):
        try:
            value = active_instance_count() if callable(active_instance_count) else active_instance_count
            return max(1, int(value or 1))
        except Exception as exc:
            _safe_log(log_error, f"Read active instance count for autostart failed: {exc!r}")
            return max(1, int(fallback or 1))

    accounted_count = read_active_count(1)
    launch_slot_count = max(0, target_count - accounted_count)
    if launch_slot_count <= 0:
        return 0

    try:
        attempts_per_instance = max(1, int(retry_count))
    except (TypeError, ValueError):
        attempts_per_instance = AUTOSTART_LAUNCH_RETRY_COUNT

    try:
        command, workdir = prepare_launch(autostart_flag, child_flag)
    except Exception as exc:
        _safe_log(log_error, f"Prepare autostart companion launch failed: {exc!r}")
        return 0

    launched = 0
    for companion_index in range(launch_slot_count):
        for attempt_index in range(attempts_per_instance):
            if wait_before_launch(max(0.0, float(interval_seconds))):
                return launched
            accounted_count = max(
                accounted_count,
                read_active_count(accounted_count),
            )
            if accounted_count >= target_count:
                return launched
            try:
                launch_process(
                    list(command),
                    env=clean_env(),
                    cwd=workdir,
                )
                launched += 1
                accounted_count += 1
                break
            except Exception as exc:
                _safe_log(
                    log_error,
                    "Launch autostart companion "
                    f"{companion_index + 1}/{launch_slot_count} failed "
                    f"(attempt {attempt_index + 1}/{attempts_per_instance}): {exc!r}",
                )
    return launched
