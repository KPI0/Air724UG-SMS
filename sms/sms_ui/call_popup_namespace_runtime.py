import time
from contextlib import contextmanager

from sms_core.status_text import format_connected_status
from sms_core.call_command_context import (
    SerialCallStateSnapshot, call_state_lock, invalidate_call_commands, set_local_dial_popup,
)

from sms_ui.call_popup_runtime import (
    close_call_popup_app_runtime,
    close_call_popup_runtime,
    show_call_popup_app_runtime,
)
from sms_ui.missed_call_popup_runtime import show_missed_call_popup_app_runtime


PEER_CALL_BINDING_TTL_SECONDS = 15.0
RETIRED_PEER_CALL_TTL_SECONDS = 120.0
RETIRED_PEER_CALL_MAX_ENTRIES = 100
PEER_CALL_MAX_PENDING_ENTRIES = 100
_PEER_PENDING_KEYS = (
    "call_popup_pending_peer_incoming",
    "call_popup_pending_peer_connected",
    "call_popup_pending_peer_terminal",
)


def _peer_state_lock(namespace):
    return call_state_lock(namespace)


def _call_snapshot(namespace):
    snapshot = getattr(namespace.get("INCOMING_CALL_SESSION"), "snapshot", None)
    return snapshot() if callable(snapshot) else None


def _local_call_session_id(namespace):
    snapshot = _call_snapshot(namespace)
    if snapshot is not None and hasattr(snapshot, "session_token"):
        pair = namespace.get("call_popup_local_session")
        if (
            isinstance(pair, tuple) and len(pair) == 2
            and pair[0] is not None
            and getattr(snapshot, "session_token", None) is pair[0]
        ):
            return str(pair[1] or "").strip()
        return ""
    # Older integrations without a lifecycle-aware tracker retain their
    # explicit local generation; production never derives it from a number.
    return str(namespace.get("call_popup_active_session_id") or "").strip()


def _same_local_call(namespace, local_session_id, snapshot):
    current = _call_snapshot(namespace)
    token = getattr(snapshot, "session_token", None)
    if token is not None and getattr(current, "session_token", None) is not token:
        return False
    if local_session_id and _local_call_session_id(namespace) != local_session_id:
        return False
    if token is None and getattr(current, "started_at", None) != getattr(snapshot, "started_at", None):
        return False
    return True


def _pending_peer_entries(namespace, key):
    with _peer_state_lock(namespace):
        pending = namespace.get(key)
        if not isinstance(pending, dict):
            return {}
        if isinstance(pending.get("session_id"), str):
            pending = {pending["session_id"]: pending}
        now = _peer_now(namespace)
        entries = {}
        for peer_id, entry in list(pending.items()):
            if not isinstance(entry, dict) or not peer_id:
                continue
            try:
                expires_at = float(entry.get("expires_at") or 0.0)
            except (TypeError, ValueError):
                continue
            if expires_at > now and not _is_retired_peer_call(namespace, peer_id):
                entries[peer_id] = entry
        entries = dict(list(entries.items())[-PEER_CALL_MAX_PENDING_ENTRIES:])
        namespace[key] = entries or None
        return entries


def _queue_peer_entry(namespace, key, entry):
    peer_id = str(entry.get("session_id") or "").strip()
    if not peer_id:
        return False
    with _peer_state_lock(namespace):
        if _is_retired_peer_call(namespace, peer_id):
            return False
        entries = _pending_peer_entries(namespace, key)
        previous = entries.get(peer_id)
        if previous and previous.get("local_session_id") and (
            previous["local_session_id"] != entry.get("local_session_id")
        ):
            return False
        entries.pop(peer_id, None)
        entries[peer_id] = dict(entry, expires_at=_peer_now(namespace) + PEER_CALL_BINDING_TTL_SECONDS)
        namespace[key] = dict(list(entries.items())[-PEER_CALL_MAX_PENDING_ENTRIES:])
    return True


def _take_pending_peer_entry(namespace, key, peer_id):
    with _peer_state_lock(namespace):
        entries = _pending_peer_entries(namespace, key)
        entry = entries.pop(peer_id, None)
        namespace[key] = entries or None
        return entry


def _remember_serial_incoming_end(namespace, local_session_id):
    if not local_session_id:
        return
    with _peer_state_lock(namespace):
        now = _peer_now(namespace)
        entries = {
            key: expires_at
            for key, expires_at in namespace.get("call_popup_pending_serial_ends", {}).items()
            if expires_at > now and key != local_session_id
        }
        entries[local_session_id] = now + RETIRED_PEER_CALL_TTL_SECONDS
        namespace["call_popup_pending_serial_ends"] = dict(
            list(entries.items())[-RETIRED_PEER_CALL_MAX_ENTRIES:]
        )


def consume_remote_incoming_end_namespace_runtime(namespace, local_session_id):
    with _peer_state_lock(namespace):
        now = _peer_now(namespace)
        entries = {
            key: expires_at
            for key, expires_at in namespace.get("call_popup_pending_serial_ends", {}).items()
            if expires_at > now
        }
        matched = entries.pop(str(local_session_id or ""), None) is not None
        namespace["call_popup_pending_serial_ends"] = entries
        return matched


def _peer_now(namespace):
    clock = namespace.get("_call_peer_monotonic")
    if not callable(clock):
        clock = time.monotonic
    try:
        return float(clock())
    except Exception:
        return time.monotonic()


def _is_retired_peer_call(namespace, session_id):
    with _peer_state_lock(namespace):
        if not session_id:
            return False
        retired = namespace.get("call_popup_retired_peer_sessions", {})
        return retired.get(session_id, 0.0) > _peer_now(namespace)


def _retire_peer_call(namespace):
    with _peer_state_lock(namespace):
        peer_id = str(namespace.get("call_popup_peer_session_id") or "").strip()
        if not peer_id:
            return
        now = _peer_now(namespace)
        # Keep completed generations beyond the binding lifetime: a reliable
        # peer event can be replayed while the same caller is already ringing again.
        retired = {
            key: expires_at
            for key, expires_at in namespace.get("call_popup_retired_peer_sessions", {}).items()
            if expires_at > now and key != peer_id
        }
        retired[peer_id] = now + RETIRED_PEER_CALL_TTL_SECONDS
        namespace.__setitem__(
            "call_popup_retired_peer_sessions",
            dict(list(retired.items())[-RETIRED_PEER_CALL_MAX_ENTRIES:]),
        )


def _clear_peer_call_state(namespace, *, retire_current=False, preserve_next=False):
    with _peer_state_lock(namespace):
        peer_id = namespace.get("call_popup_peer_session_id")
        if retire_current:
            _retire_peer_call(namespace)
        for key in _PEER_PENDING_KEYS:
            entries = _pending_peer_entries(namespace, key) if preserve_next else {}
            entries.pop(peer_id, None)
            namespace[key] = entries or None
        for key in ("call_popup_peer_session_id", "call_popup_peer_caller_num",
                    "call_popup_peer_local_session_id", "call_popup_peer_expires_at"):
            namespace[key] = "" if key.endswith("_id") or key.endswith("_num") else None


def _pending_peer_entry(namespace, key, peer_id):
    return _pending_peer_entries(namespace, key).get(peer_id)


def _get_peer_call_binding(namespace):
    with _peer_state_lock(namespace):
        peer_id = str(namespace.get("call_popup_peer_session_id") or "").strip()
        if not peer_id:
            return None
        try:
            expires_at = float(namespace.get("call_popup_peer_expires_at") or 0.0)
        except (TypeError, ValueError):
            expires_at = 0.0
        binding = {
            "peer_session_id": peer_id,
            "caller_num": str(namespace.get("call_popup_peer_caller_num") or "").strip(),
            "local_session_id": str(
                namespace.get("call_popup_peer_local_session_id") or ""
            ).strip(),
        }
        if expires_at <= _peer_now(namespace):
            active_session = _local_call_session_id(namespace)
            # The short TTL only limits an unbound, out-of-order peer frame.  Once
            # a firmware generation is bound to the active local generation it
            # must survive the full call, including calls that ring for longer
            # than the pending window.  Normal call finish/reset clears it.
            if binding["local_session_id"] and binding["local_session_id"] == active_session:
                return binding
            _clear_peer_call_state(namespace, preserve_next=True)
            return None
        return binding


def _set_peer_call_binding(namespace, session_id, caller_num, local_session_id=""):
    with _peer_state_lock(namespace):
        peer_id = str(session_id or "").strip()
        if not peer_id or _is_retired_peer_call(namespace, peer_id):
            return False
        namespace.__setitem__("call_popup_peer_session_id", peer_id)
        namespace.__setitem__("call_popup_peer_caller_num", str(caller_num or "").strip())
        namespace.__setitem__(
            "call_popup_peer_local_session_id", str(local_session_id or "").strip()
        )
        namespace.__setitem__(
            "call_popup_peer_expires_at",
            _peer_now(namespace) + PEER_CALL_BINDING_TTL_SECONDS,
        )
        return True


def register_peer_incoming_call_namespace_runtime(
    namespace, session_id="", caller_num="", local_session_id="",
):
    """Remember the server-confirmed peer/local pair, never infer it by phone."""
    peer_id = str(session_id or "").strip()
    local_id = str(local_session_id or "").strip()
    caller_text = str(caller_num or "").strip()
    if not peer_id or not local_id or peer_id == local_id or _is_retired_peer_call(namespace, peer_id):
        return False
    binding = _get_peer_call_binding(namespace)
    if binding is not None and binding["peer_session_id"] == peer_id:
        if binding["local_session_id"] != local_id:
            return False
        _set_peer_call_binding(namespace, peer_id, binding["caller_num"] or caller_text, local_id)
        _consume_pending_peer_connected(namespace)
        _consume_pending_peer_terminal(namespace)
        return True
    active_session = _local_call_session_id(namespace)
    snapshot = _call_snapshot(namespace)
    popup = namespace.get("current_call_popup")
    current_caller = str(
        getattr(snapshot, "caller_num", "")
        or getattr(popup, "_call_popup_caller_num", "") or ""
    ).strip()
    if local_id == active_session and caller_text and current_caller and caller_text != current_caller:
        return False
    if binding is not None and binding["local_session_id"] == local_id == active_session:
        return False
    if not _queue_peer_entry(namespace, "call_popup_pending_peer_incoming", {
        "session_id": peer_id, "local_session_id": local_id, "caller_num": caller_text,
    }):
        return False
    if local_id == active_session and current_caller:
        _bind_peer_to_local_session(namespace, local_id, current_caller)
        _consume_pending_peer_connected(namespace)
        _consume_pending_peer_terminal(namespace)
    return True


def _bind_peer_to_local_session(namespace, local_session_id, caller_num):
    local_id = str(local_session_id or "").strip()
    caller_text = str(caller_num or "").strip()
    # Serial start/finish must not replace the lifetime between selecting a
    # pending peer and committing it. UI work remains outside this state lock.
    with _peer_state_lock(namespace):
        snapshot = _call_snapshot(namespace)
        if not local_id or local_id != _local_call_session_id(namespace):
            return None
        binding = _get_peer_call_binding(namespace)
        if binding is not None and binding["local_session_id"] == local_id:
            if binding["caller_num"] and caller_text and binding["caller_num"] != caller_text:
                return None
            return binding
        for peer_id, pending in _pending_peer_entries(namespace, "call_popup_pending_peer_incoming").items():
            if pending.get("local_session_id") != local_id:
                continue
            pending_caller = str(pending.get("caller_num") or "").strip()
            if pending_caller and caller_text and pending_caller != caller_text:
                continue
            if not _same_local_call(namespace, local_id, snapshot):
                return None
            _take_pending_peer_entry(namespace, "call_popup_pending_peer_incoming", peer_id)
            if binding is not None:
                _clear_peer_call_state(namespace, retire_current=True, preserve_next=True)
            if not _set_peer_call_binding(namespace, peer_id, pending_caller or caller_text, local_id):
                return None
            return _get_peer_call_binding(namespace)
    return None


def _consume_pending_peer_terminal(namespace):
    binding = _get_peer_call_binding(namespace)
    if binding is None or binding["local_session_id"] != _local_call_session_id(namespace):
        return False
    pending = _pending_peer_entry(
        namespace, "call_popup_pending_peer_terminal", binding["peer_session_id"],
    )
    if pending is None:
        return False
    pending_caller = str(pending.get("caller_num") or "").strip()
    if pending_caller and binding["caller_num"] and pending_caller != binding["caller_num"]:
        return False
    _take_pending_peer_entry(namespace, "call_popup_pending_peer_terminal", binding["peer_session_id"])
    return finish_remote_incoming_call_namespace_runtime(
        namespace, pending.get("phase") or "ended", pending.get("reason") or "",
        session_id=binding["peer_session_id"],
        caller_num=pending_caller or binding["caller_num"],
    )


def _consume_pending_peer_connected(namespace):
    binding = _get_peer_call_binding(namespace)
    if binding is None or binding["local_session_id"] != _local_call_session_id(namespace):
        return False
    pending = _pending_peer_entry(
        namespace, "call_popup_pending_peer_connected", binding["peer_session_id"],
    )
    if pending is None:
        return False
    pending_caller = str(pending.get("caller_num") or "").strip()
    if pending_caller and binding["caller_num"] and pending_caller != binding["caller_num"]:
        return False
    _take_pending_peer_entry(namespace, "call_popup_pending_peer_connected", binding["peer_session_id"])
    return mark_peer_incoming_call_connected_namespace_runtime(
        namespace, binding["peer_session_id"], pending_caller or binding["caller_num"],
    )


def mark_peer_incoming_call_connected_namespace_runtime(
    namespace, session_id="", caller_num="",
):
    peer_id = str(session_id or "").strip()
    caller_text = str(caller_num or "").strip()
    if _is_retired_peer_call(namespace, peer_id):
        return False
    binding = _get_peer_call_binding(namespace)
    local_id = _local_call_session_id(namespace)
    if peer_id:
        if (
            binding is None or binding["peer_session_id"] != peer_id
            or not local_id or binding["local_session_id"] != local_id
        ):
            return _queue_peer_entry(namespace, "call_popup_pending_peer_connected", {
                "session_id": peer_id, "caller_num": caller_text,
            })
        if caller_text and binding["caller_num"] and caller_text != binding["caller_num"]:
            return False
    elif binding is not None:
        return False
    if not local_id:
        return False
    return mark_call_popup_connected_namespace_runtime(namespace, local_id, caller_text)


def _sync_phone_popup_stacking(namespace, window=None):
    try:
        if window is not None:
            bind = getattr(window, "bind", None)
            exists = getattr(window, "winfo_exists", None)
            if (
                callable(bind)
                and callable(exists)
                and exists()
                and not getattr(window, "_call_popup_stacking_bound", False)
            ):
                def sync_on_visibility(event):
                    if event.widget is window:
                        # A delayed event from an old window must use the
                        # currently registered incoming window's visibility.
                        _sync_phone_popup_stacking(namespace)

                for event_name in ("<Map>", "<Unmap>", "<Destroy>"):
                    bind(event_name, sync_on_visibility, add="+")
                window._call_popup_stacking_bound = True

        missed = namespace.get("current_missed_call_popup")
        missed_exists = getattr(missed, "winfo_exists", None)
        attributes = getattr(missed, "attributes", None)
        if not callable(missed_exists) or not callable(attributes) or not missed_exists():
            return
        incoming = namespace.get("current_call_popup")
        incoming_exists = getattr(incoming, "winfo_exists", None)
        is_mapped = getattr(incoming, "winfo_ismapped", None)
        incoming_visible = (
            callable(incoming_exists)
            and callable(is_mapped)
            and incoming_exists()
            and is_mapped()
        )
        # Two topmost windows can cover each other when focus changes.
        # Keep the reminder below a visible live call without taking focus.
        topmost = not incoming_visible
        if bool(attributes("-topmost")) != topmost:
            attributes("-topmost", topmost)
    except Exception as exc:
        log_error = namespace.get("log_file_only")
        if callable(log_error):
            try:
                log_error(f"Sync phone popup stacking failed: {exc!r}")
            except Exception:
                pass


def set_call_popup_namespace_runtime(namespace, window, *, call_session_id=None):
    namespace.__setitem__("current_call_popup", window)
    _sync_phone_popup_stacking(namespace, window)
    if window is None:
        return

    active_session_id = str(
        (namespace.get("call_popup_active_session_id", "")
         if call_session_id is None else call_session_id) or ""
    ).strip()
    if active_session_id:
        try:
            setattr(window, "_call_popup_session_id", active_session_id)
        except Exception:
            pass
        caller_num = str(getattr(window, "_call_popup_caller_num", "") or "").strip()
        if caller_num:
            _bind_peer_to_local_session(namespace, active_session_id, caller_num)
            _consume_pending_peer_connected(namespace)
            if _consume_pending_peer_terminal(namespace):
                # Consuming a terminal state can synchronously destroy this
                # window and clear the active generation.  Do not continue
                # applying a queued connected marker to that stale object.
                return

    pending_session_id = str(
        namespace.get("call_popup_pending_connected_session_id", "") or ""
    ).strip()
    if not pending_session_id:
        return
    if not active_session_id:
        namespace.__setitem__("call_popup_pending_connected_session_id", "")
        return
    if active_session_id and pending_session_id != active_session_id:
        namespace.__setitem__("call_popup_pending_connected_session_id", "")
        return
    marker = getattr(window, "_call_popup_mark_connected", None)
    if not callable(marker):
        return
    try:
        if not window.winfo_exists():
            return
        marker()
        handled = namespace.get("mark_incoming_call_handled")
        if callable(handled):
            handled()
        namespace.__setitem__("call_popup_pending_connected_session_id", "")
    except Exception as exc:
        log_error = namespace.get("log_file_only")
        if callable(log_error):
            try:
                log_error(f"Apply pending incoming call connected state failed: {exc!r}")
            except Exception:
                pass


def set_dial_popup_namespace_runtime(namespace, window):
    set_local_dial_popup(namespace, window)


def mark_dial_popup_connected_namespace_runtime(namespace, *, call_state_snapshot=None):
    popup = call_state_snapshot.popup if call_state_snapshot is not None else namespace.get("current_dial_popup")

    def mark_on_ui():
        if (popup is None or namespace.get("current_dial_popup") is not popup
                or (call_state_snapshot is not None and not call_state_snapshot.is_current())):
            return False
        try:
            if not popup.winfo_exists():
                return False
        except Exception:
            return False
        marker = getattr(popup, "_call_popup_mark_connected", None)
        if not callable(marker):
            return False
        try:
            marker()
            return True
        except Exception as exc:
            log_error = namespace.get("log_file_only")
            if callable(log_error):
                try:
                    log_error(f"Mark dial popup connected failed: {exc!r}")
                except Exception:
                    pass
            return False

    return namespace["run_on_ui_thread"](mark_on_ui, namespace["ui_post"])


def mark_call_popup_connected_namespace_runtime(
    namespace, call_session_id="", caller_num="", *, call_guard=None,
):
    with _peer_state_lock(namespace):
        if call_guard is not None and not call_guard():
            return False
        local_id = _local_call_session_id(namespace)
        snapshot = _call_snapshot(namespace)
        requested_session_id = str(call_session_id or "").strip()
        caller_text = str(caller_num or "").strip()
        popup = namespace.get("current_call_popup")
        if snapshot is not None and hasattr(snapshot, "session_token") and snapshot.session_token is None:
            return False
        if not local_id and popup is None:
            return False
        popup_caller = str(getattr(popup, "_call_popup_caller_num", "") or "").strip()
        current_caller = str(getattr(snapshot, "caller_num", "") or popup_caller).strip()
        if caller_text and current_caller and caller_text != current_caller:
            return False
        if requested_session_id and local_id and requested_session_id != local_id:
            binding = _get_peer_call_binding(namespace)
            if not (
                binding is not None and binding["peer_session_id"] == requested_session_id
                and binding["local_session_id"] == local_id
            ):
                return False
            requested_session_id = local_id
        if requested_session_id and not local_id:
            return False
        if requested_session_id:
            namespace["call_popup_pending_connected_session_id"] = requested_session_id

    def mark_on_ui():
        with _peer_state_lock(namespace):
            if call_guard is not None and not call_guard():
                return False
            if not _same_local_call(namespace, local_id, snapshot):
                return False
            tracker = namespace.get("INCOMING_CALL_SESSION")
            handled = getattr(tracker, "mark_handled", None)
            if not callable(handled):
                handled = namespace.get("mark_incoming_call_handled")
            if callable(handled):
                handled()
            update_status = not namespace.get("current_dial_num")
            if update_status:
                namespace["ring_timeout_target"] = -1.0
        if update_status:
            set_status = namespace.get("set_status")
            if callable(set_status):
                set_status(f"📞 通话中：{current_caller or caller_text}", "blue")
        current_popup = namespace.get("current_call_popup")
        if current_popup is None:
            return True
        try:
            if not current_popup.winfo_exists():
                return True
            popup_session = str(getattr(current_popup, "_call_popup_session_id", local_id) or "").strip()
            if local_id and popup_session and popup_session != local_id:
                return False
            marker = getattr(current_popup, "_call_popup_mark_connected", None)
            if not callable(marker):
                return True
            marker()
            namespace["call_popup_pending_connected_session_id"] = ""
            return True
        except Exception as exc:
            log_error = namespace.get("log_file_only")
            if callable(log_error):
                try:
                    log_error(f"Mark incoming call popup connected failed: {exc!r}")
                except Exception:
                    pass
            return False

    return namespace["run_on_ui_thread"](mark_on_ui, namespace["ui_post"])


def finish_dial_popup_namespace_runtime(namespace, message="", *, call_state_snapshot=None):
    popup = call_state_snapshot.popup if call_state_snapshot is not None else namespace.get("current_dial_popup")

    def finish_on_ui():
        if (popup is None or namespace.get("current_dial_popup") is not popup
                or (call_state_snapshot is not None and not call_state_snapshot.is_current(check_incoming=False))):
            return False
        try:
            if not popup.winfo_exists():
                return False
        except Exception:
            return False
        marker = getattr(popup, "_call_popup_mark_ended", None)
        if not callable(marker):
            return False
        try:
            marker(message)
            return True
        except Exception as exc:
            log_error = namespace.get("log_file_only")
            if callable(log_error):
                try:
                    log_error(f"Mark dial popup ended failed: {exc!r}")
                except Exception:
                    pass
            return False

    return namespace["run_on_ui_thread"](finish_on_ui, namespace["ui_post"])


def set_missed_call_popup_namespace_runtime(namespace, window):
    namespace.__setitem__("current_missed_call_popup", window)
    _sync_phone_popup_stacking(namespace)


def start_incoming_call_session_namespace_runtime(namespace, caller_num):
    with _peer_state_lock(namespace):
        result = namespace["INCOMING_CALL_SESSION"].start(caller_num)
        if result:
            invalidate_call_commands(namespace)
        return result


def serial_call_state_guard_namespace_runtime(namespace):
    return _peer_state_lock(namespace)


def start_serial_incoming_call_namespace_runtime(namespace, caller_num, session_id, ring_timeout):
    """Commit a serial-created lifetime before dispatching its UI/IO effects."""
    with _peer_state_lock(namespace):
        result = namespace["INCOMING_CALL_SESSION"].start(caller_num)
        if result:
            invalidate_call_commands(namespace)
            snapshot = _call_snapshot(namespace)
            namespace["call_popup_local_session"] = (
                getattr(snapshot, "session_token", None), str(session_id or "").strip(),
            )
            namespace["ring_timeout_target"] = ring_timeout
            namespace["current_dial_num"] = ""
        return result


def mark_incoming_call_handled_namespace_runtime(namespace):
    return namespace["INCOMING_CALL_SESSION"].mark_handled()


def _has_local_incoming_call(namespace):
    if namespace.get("call_popup_peer_local_session_id"):
        return True
    snapshot = getattr(namespace["INCOMING_CALL_SESSION"], "snapshot", None)
    return callable(snapshot) and bool(snapshot().caller_num)


def finish_incoming_call_session_namespace_runtime(namespace):
    with _peer_state_lock(namespace):
        had_local_call = _has_local_incoming_call(namespace)
        result = namespace["INCOMING_CALL_SESSION"].finish()
        namespace["call_popup_local_session"] = None
        # Duplicate serial termination must preserve future peer generations.
        if had_local_call:
            invalidate_call_commands(namespace)
            _clear_peer_call_state(namespace, retire_current=True, preserve_next=True)
        return result


def finish_remote_incoming_call_namespace_runtime(
    namespace, phase="ended", reason="", *, session_id="", caller_num="",
):
    """Apply an authenticated peer terminal only to its known local lifetime."""
    peer_id = str(session_id or "").strip()
    caller_text = str(caller_num or "").strip()
    with _peer_state_lock(namespace):
        if _is_retired_peer_call(namespace, peer_id):
            return False
        local_id = _local_call_session_id(namespace)
        snapshot = _call_snapshot(namespace)
        current_popup = namespace.get("current_call_popup")
        current_caller = str(
            getattr(snapshot, "caller_num", "")
            or getattr(current_popup, "_call_popup_caller_num", "") or ""
        ).strip()
        binding = _get_peer_call_binding(namespace)
        if peer_id:
            if (
                binding is None or binding["peer_session_id"] != peer_id
                or not local_id or binding["local_session_id"] != local_id
            ):
                return _queue_peer_entry(namespace, "call_popup_pending_peer_terminal", {
                    "session_id": peer_id, "caller_num": caller_text,
                    "phase": str(phase or "ended"), "reason": str(reason or ""),
                })
        elif binding is not None:
            return False
        if snapshot is not None and hasattr(snapshot, "session_token") and snapshot.session_token is None:
            return False
        if not current_caller or (caller_text and caller_text != current_caller):
            return False

        # Commit modem lifetime state before waiting for Tk. A queued terminal
        # must not leave the next same-number CLIP looking like a duplicate.
        if local_id and namespace.get("call_popup_pending_connected_session_id") == local_id:
            # Its accepted connected transition may still be in the UI queue.
            mark_handled = getattr(namespace["INCOMING_CALL_SESSION"], "mark_handled", None)
            if callable(mark_handled):
                mark_handled()
        missed_call = namespace["INCOMING_CALL_SESSION"].finish()
        invalidate_call_commands(namespace)
        namespace["call_popup_local_session"] = None
        namespace["call_popup_active_session_id"] = ""
        namespace["call_popup_request_token"] = None
        namespace["call_popup_pending_connected_session_id"] = ""
        _remember_serial_incoming_end(namespace, local_id)
        _clear_peer_call_state(namespace, retire_current=True, preserve_next=True)
        if not namespace.get("current_dial_num"):
            namespace["ring_timeout_target"] = 0.0

    def finish_on_ui():
        with _peer_state_lock(namespace):
            update_status = (
                not _has_local_incoming_call(namespace)
                and not namespace.get("current_dial_num")
            )
        if update_status:
            set_status = namespace.get("set_status")
            if callable(set_status):
                set_status(format_connected_status(namespace.get("PORT", "")), "green")

        def clear_if_current(window):
            if namespace.get("current_call_popup") is current_popup:
                set_call_popup_namespace_runtime(namespace, window)

        # A peer inbound terminal must not close an independent outbound window.
        close_call_popup_runtime(
            current_popup, clear_if_current, log_error=namespace.get("log_file_only"),
        )
        if missed_call is not None:
            show_missed = namespace.get("show_missed_call_popup")
            if callable(show_missed):
                show_missed(missed_call)
        return True

    return namespace["run_on_ui_thread"](finish_on_ui, namespace["ui_post"])


def reset_incoming_call_session_namespace_runtime(namespace):
    with _peer_state_lock(namespace):
        had_local_call = _has_local_incoming_call(namespace)
        result = namespace["INCOMING_CALL_SESSION"].reset()
        invalidate_call_commands(namespace)
        namespace["call_popup_local_session"] = None
        namespace["call_popup_pending_serial_ends"] = {}
        _clear_peer_call_state(namespace, retire_current=had_local_call)
        return result


def close_call_popup_namespace_runtime(namespace, *, close_app_runtime=close_call_popup_app_runtime):
    # Invalidate the session before queueing window destruction.  A late
    # connected callback must not update a popup belonging to a newer call.
    namespace.__setitem__("call_popup_active_session_id", "")
    namespace.__setitem__("call_popup_request_token", None)
    namespace.__setitem__("call_popup_pending_connected_session_id", "")
    result = close_app_runtime(
        get_popup=lambda: namespace["current_call_popup"],
        set_popup=lambda window: set_call_popup_namespace_runtime(namespace, window),
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
        log_error=namespace.get("log_file_only"),
    )
    if namespace.get("current_dial_popup") is not None:
        dial_result = close_app_runtime(
            get_popup=lambda: namespace.get("current_dial_popup"),
            set_popup=lambda window: set_dial_popup_namespace_runtime(namespace, window),
            run_on_ui_thread=namespace["run_on_ui_thread"],
            ui_post=namespace["ui_post"],
            log_error=namespace.get("log_file_only"),
        )
        if result is None:
            result = dial_result
    return result


def close_missed_call_popup_namespace_runtime(namespace, *, close_app_runtime=close_call_popup_app_runtime):
    return close_app_runtime(
        get_popup=lambda: namespace["current_missed_call_popup"],
        set_popup=lambda window: set_missed_call_popup_namespace_runtime(namespace, window),
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
        log_error=namespace.get("log_file_only"),
    )


def close_phone_popups_namespace_runtime(namespace):
    for callback_name in ("close_call_popup", "close_missed_call_popup"):
        callback = namespace.get(callback_name)
        if not callable(callback):
            continue
        try:
            callback()
        except Exception as exc:
            log_error = namespace.get("log_file_only")
            if callable(log_error):
                try:
                    log_error(f"Close {callback_name} after disabling phone popups failed: {exc!r}")
                except Exception:
                    pass


def show_call_popup_namespace_runtime(
    namespace,
    caller_num,
    call_session_id="",
    *,
    show_app_runtime=show_call_popup_app_runtime,
):
    session_id = str(call_session_id or "").strip()
    with _peer_state_lock(namespace):
        initial_call = _call_snapshot(namespace)
        call_serial = namespace.get("serial_obj")
        serial_generation = namespace.get("serial_connection_generation", 0)
        token = getattr(initial_call, "session_token", None)
        if initial_call is not None and hasattr(initial_call, "session_token") and token is None:
            return "disabled"
        if token is not None:
            namespace["call_popup_local_session"] = (token, session_id)

    def same_initial_lifetime():
        if initial_call is None or not hasattr(initial_call, "session_token"):
            return True
        return getattr(_call_snapshot(namespace), "session_token", None) is token

    popup_enabled = namespace.get("CALL_POPUP_ENABLED", True)
    request_token = object()
    with _peer_state_lock(namespace):
        # A peer terminal or redial can arrive after the local pair was
        # recorded. Do not overwrite a newer request before checking its life.
        if not same_initial_lifetime():
            return "disabled"
        if popup_enabled:
            namespace.__setitem__("call_popup_request_token", request_token)
            pending_session_id = str(
                namespace.get("call_popup_pending_connected_session_id", "") or ""
            ).strip()
            namespace.__setitem__("call_popup_active_session_id", session_id)
    if not popup_enabled:
        if session_id:
            _bind_peer_to_local_session(namespace, session_id, caller_num)
            _consume_pending_peer_connected(namespace)
            _consume_pending_peer_terminal(namespace)
        return "disabled"
    if session_id:
        _bind_peer_to_local_session(namespace, session_id, caller_num)
    # A connected indication can be queued before this UI task gets a chance
    # to create the popup.  Keep a matching pending marker so set_popup() can
    # apply it immediately after the window is registered; discard only a
    # marker belonging to a different call generation.
    if pending_session_id and pending_session_id != session_id:
        namespace.__setitem__("call_popup_pending_connected_session_id", "")
    popup_holder = {}

    def is_current():
        return (
            namespace.get("CALL_POPUP_ENABLED", True)
            and same_initial_lifetime()
            and namespace.get("call_popup_request_token") is request_token
            and namespace.get("call_popup_active_session_id", "") == session_id
        )

    def is_same_call():
        session_token = getattr(initial_call, "session_token", None)
        if session_token is not None:
            return same_initial_lifetime()
        return is_current()

    def is_same_connection():
        return (
            namespace.get("serial_obj") is call_serial
            and namespace.get("serial_connection_generation", 0) == serial_generation
            and namespace.get("serial_running", True) is not False
            and not any(event is not None and event.is_set() for event in (
                namespace.get("serial_stop_event"), namespace.get("TK_SHUTDOWN"),
            ))
        )

    @contextmanager
    def write_guard():
        # The sender already owns the transaction, connection and write locks.
        # Keep the call lifetime stable until the command bytes are committed.
        with _peer_state_lock(namespace):
            yield is_same_call() and is_same_connection()

    def mark_handled():
        with _peer_state_lock(namespace):
            if not is_same_call() or not is_same_connection():
                return False
            return mark_incoming_call_handled_namespace_runtime(namespace)

    def dispatch_call_result(callback):
        def apply_on_ui():
            # Hiding a window does not end its modem call. Use the tracked
            # call lifetime, and recheck after any wait in the UI queue.
            if is_same_call():
                return callback()
            return False
        return namespace["run_on_ui_thread"](apply_on_ui, namespace["ui_post"])

    def set_ring_timeout(value):
        with _peer_state_lock(namespace):
            if is_same_call():
                namespace.__setitem__("ring_timeout_target", value)

    def discard_popup(window):
        def clear_if_current(value):
            if namespace.get("current_call_popup") is window:
                namespace.__setitem__("current_call_popup", value)
        close_call_popup_runtime(window, clear_if_current, log_error=namespace.get("log_file_only"))

    def get_popup():
        window = namespace.get("current_call_popup")
        if session_id and window is not None and (
            getattr(window, "_call_popup_session_id", "") != session_id
        ):
            discard_popup(window)
            return None
        return window

    def set_popup(window):
        popup_holder["window"] = window
        # Tk construction/mapping can yield to terminal or new-call work.
        # Validate again after construction and never label an old window
        # with the generation that happens to be active when it returns.
        if not is_current():
            discard_popup(window)
            return
        set_call_popup_namespace_runtime(namespace, window, call_session_id=session_id)
        if not is_current():
            discard_popup(window)

    def close_popup():
        def close_on_ui():
            if namespace.get("current_call_popup") is not popup_holder.get("window"):
                return False
            if session_id and namespace.get("call_popup_active_session_id") != session_id:
                return False
            return namespace["close_call_popup"]()

        return namespace["run_on_ui_thread"](close_on_ui, namespace["ui_post"])

    return show_app_runtime(
        parent=namespace["root"],
        caller_num=caller_num,
        get_popup=get_popup,
        set_popup=set_popup,
        center_window=namespace["center_window"],
        serial_lock=namespace["serial_lock"],
        get_serial=lambda: call_serial,
        port_ui=namespace["port_ui"],
        set_status=namespace["set_status"],
        ui_post=namespace["ui_post"],
        close_popup=close_popup,
        set_ring_timeout=set_ring_timeout,
        run_on_ui_thread=namespace["run_on_ui_thread"],
        is_enabled=is_current,
        mark_call_handled=mark_handled,
        log_error=namespace.get("log_file_only"),
        dispatch_call_result=dispatch_call_result,
        write_guard=write_guard,
    )


def show_missed_call_popup_namespace_runtime(
    namespace,
    missed_call,
    *,
    show_app_runtime=show_missed_call_popup_app_runtime,
):
    if not namespace.get("CALL_POPUP_ENABLED", True):
        return "disabled"
    return show_app_runtime(
        missed_call=missed_call,
        parent=namespace["root"],
        get_popup=lambda: namespace["current_missed_call_popup"],
        set_popup=lambda window: set_missed_call_popup_namespace_runtime(namespace, window),
        center_window=namespace["center_window"],
        show_window=namespace["show_window"],
        get_active_call_popup=lambda: namespace.get("current_call_popup"),
        log_error=namespace.get("log_file_only"),
        run_on_ui_thread=namespace["run_on_ui_thread"],
        ui_post=namespace["ui_post"],
        is_enabled=lambda: bool(namespace.get("CALL_POPUP_ENABLED", True)),
    )


def get_serial_call_state_namespace_runtime(namespace):
    return SerialCallStateSnapshot(namespace)


def set_serial_call_state_namespace_runtime(namespace, next_ring_timeout, next_dial_num):
    with _peer_state_lock(namespace):
        namespace.__setitem__("ring_timeout_target", next_ring_timeout)
        namespace.__setitem__("current_dial_num", next_dial_num)


def merge_serial_call_state_namespace_runtime(namespace, previous_state, next_ring_timeout, next_dial_num):
    with _peer_state_lock(namespace):
        if isinstance(previous_state, SerialCallStateSnapshot) and not previous_state.is_current():
            return get_serial_call_state_namespace_runtime(namespace)
        previous_ring_timeout, previous_dial_num = previous_state
        current_ring_timeout, current_dial_num = get_serial_call_state_namespace_runtime(namespace)
        snapshot = _call_snapshot(namespace)
        incoming_ended = (
            next_ring_timeout == 0.0 and not current_dial_num
            and not getattr(snapshot, "caller_num", "")
        )
        if incoming_ended or current_ring_timeout == previous_ring_timeout:
            namespace["ring_timeout_target"] = next_ring_timeout
        if (not getattr(previous_state, "pending", False)
                and current_dial_num == previous_dial_num and next_dial_num != previous_dial_num):
            namespace["current_dial_num"] = next_dial_num
        return get_serial_call_state_namespace_runtime(namespace)
