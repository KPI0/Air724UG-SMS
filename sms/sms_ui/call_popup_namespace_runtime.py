import time

from sms_ui.call_popup_runtime import (
    close_call_popup_app_runtime,
    close_call_popup_runtime,
    show_call_popup_app_runtime,
)
from sms_ui.missed_call_popup_runtime import show_missed_call_popup_app_runtime


PEER_CALL_BINDING_TTL_SECONDS = 15.0
RETIRED_PEER_CALL_TTL_SECONDS = 120.0
RETIRED_PEER_CALL_MAX_ENTRIES = 100


def _peer_now(namespace):
    clock = namespace.get("_call_peer_monotonic")
    if not callable(clock):
        clock = time.monotonic
    try:
        return float(clock())
    except Exception:
        return time.monotonic()


def _is_retired_peer_call(namespace, session_id):
    if not session_id:
        return False
    retired = namespace.get("call_popup_retired_peer_sessions", {})
    return retired.get(session_id, 0.0) > _peer_now(namespace)


def _retire_peer_call(namespace):
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
    next_entries = {}
    if preserve_next:
        pending = _pending_peer_entry(namespace, "call_popup_pending_peer_incoming")
        if pending and pending.get("session_id") != namespace.get("call_popup_peer_session_id"):
            next_entries["call_popup_pending_peer_incoming"] = pending
            for key in ("call_popup_pending_peer_connected", "call_popup_pending_peer_terminal"):
                entry = _pending_peer_entry(namespace, key)
                if entry and entry.get("session_id") == pending.get("session_id"):
                    next_entries[key] = entry
    if retire_current:
        _retire_peer_call(namespace)
    for key in (
        "call_popup_peer_session_id",
        "call_popup_peer_caller_num",
        "call_popup_peer_local_session_id",
        "call_popup_peer_expires_at",
        "call_popup_pending_peer_connected",
        "call_popup_pending_peer_terminal",
        "call_popup_pending_peer_incoming",
    ):
        namespace.__setitem__(key, "" if key.endswith("_id") or key.endswith("_num") else None)
    namespace.update(next_entries)


def _pending_peer_entry(namespace, key):
    pending = namespace.get(key)
    if not isinstance(pending, dict):
        return None
    try:
        expires_at = float(pending.get("expires_at") or 0.0)
    except (TypeError, ValueError):
        expires_at = 0.0
    if expires_at <= _peer_now(namespace):
        namespace.__setitem__(key, None)
        return None
    return pending


def _get_peer_call_binding(namespace):
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
        active_session = str(namespace.get("call_popup_active_session_id") or "").strip()
        # The short TTL only limits an unbound, out-of-order peer frame.  Once
        # a firmware generation is bound to the active local generation it
        # must survive the full call, including calls that ring for longer
        # than the pending window.  Normal call finish/reset clears it.
        if binding["local_session_id"] and binding["local_session_id"] == active_session:
            return binding
        _clear_peer_call_state(namespace)
        return None
    return binding


def _set_peer_call_binding(namespace, session_id, caller_num, local_session_id=""):
    peer_id = str(session_id or "").strip()
    if not peer_id:
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
    namespace,
    session_id="",
    caller_num="",
):
    """Remember a firmware call generation before local serial UI catches up."""
    peer_id = str(session_id or "").strip()
    caller_text = str(caller_num or "").strip()
    if not peer_id or _is_retired_peer_call(namespace, peer_id):
        return False
    binding = _get_peer_call_binding(namespace)
    active_session = str(namespace.get("call_popup_active_session_id") or "").strip()
    popup = namespace.get("current_call_popup")
    popup_caller = str(getattr(popup, "_call_popup_caller_num", "") or "").strip()
    tracker = namespace.get("INCOMING_CALL_SESSION")
    tracked_caller = ""
    if tracker is not None and callable(getattr(tracker, "snapshot", None)):
        try:
            tracked_caller = str(tracker.snapshot().caller_num or "").strip()
        except Exception:
            tracked_caller = ""
    current_caller = popup_caller or tracked_caller
    if caller_text and current_caller and caller_text != current_caller:
        return False
    if binding is not None:
        if binding["peer_session_id"] == peer_id:
            _set_peer_call_binding(
                namespace,
                peer_id,
                binding["caller_num"] or caller_text,
                binding["local_session_id"],
            )
            return True
        if active_session or current_caller:
            # The peer has announced the next generation before the serial
            # reader has replaced the old local popup.  Keep it separately;
            # the next local session will consume it by caller and TTL.
            namespace.__setitem__(
                "call_popup_pending_peer_incoming",
                {
                    "session_id": peer_id,
                    "caller_num": caller_text,
                    "expires_at": _peer_now(namespace) + PEER_CALL_BINDING_TTL_SECONDS,
                },
            )
            return True
        _clear_peer_call_state(namespace)
    local_session_id = active_session if current_caller else ""
    if local_session_id and caller_text and current_caller != caller_text:
        return False
    _set_peer_call_binding(namespace, peer_id, caller_text or current_caller, local_session_id)
    _consume_pending_peer_terminal(namespace)
    return True


def _bind_peer_to_local_session(namespace, local_session_id, caller_num):
    local_id = str(local_session_id or "").strip()
    caller_text = str(caller_num or "").strip()
    pending = _pending_peer_entry(namespace, "call_popup_pending_peer_incoming")
    binding = _get_peer_call_binding(namespace)
    if pending is not None and (binding is None or binding["local_session_id"] != local_id):
        pending_caller = str(pending.get("caller_num") or "").strip()
        if not pending_caller or not caller_text or pending_caller == caller_text:
            _set_peer_call_binding(
                namespace,
                pending.get("session_id"),
                pending_caller or caller_text,
                local_id,
            )
            namespace.__setitem__("call_popup_pending_peer_incoming", None)
            return _get_peer_call_binding(namespace)
    if binding is None:
        return None
    if binding["caller_num"] and caller_text and binding["caller_num"] != caller_text:
        return None
    if not binding["local_session_id"]:
        _set_peer_call_binding(
            namespace,
            binding["peer_session_id"],
            binding["caller_num"] or caller_text,
            local_id,
        )
        return _get_peer_call_binding(namespace)
    return binding


def _consume_pending_peer_terminal(namespace):
    pending = _pending_peer_entry(namespace, "call_popup_pending_peer_terminal")
    if pending is None:
        return False
    binding = _get_peer_call_binding(namespace)
    active_session = str(namespace.get("call_popup_active_session_id") or "").strip()
    if (
        binding is None
        or not active_session
        or binding["local_session_id"] != active_session
        or binding["peer_session_id"] != str(pending.get("session_id") or "").strip()
    ):
        return False
    popup = namespace.get("current_call_popup")
    popup_caller = str(getattr(popup, "_call_popup_caller_num", "") or "").strip()
    pending_caller = str(pending.get("caller_num") or "").strip()
    if pending_caller and popup_caller and pending_caller != popup_caller:
        return False
    namespace.__setitem__("call_popup_pending_peer_terminal", None)
    return finish_remote_incoming_call_namespace_runtime(
        namespace,
        pending.get("phase") or "ended",
        pending.get("reason") or "",
        session_id=binding["peer_session_id"],
        caller_num=pending_caller or binding["caller_num"],
    )


def _consume_pending_peer_connected(namespace):
    pending = _pending_peer_entry(namespace, "call_popup_pending_peer_connected")
    if pending is None:
        return False
    binding = _get_peer_call_binding(namespace)
    if (
        binding is None
        or binding["peer_session_id"] != str(pending.get("session_id") or "").strip()
    ):
        return False
    namespace.__setitem__("call_popup_pending_peer_connected", None)
    return mark_peer_incoming_call_connected_namespace_runtime(
        namespace,
        binding["peer_session_id"],
        pending.get("caller_num") or binding["caller_num"],
    )


def mark_peer_incoming_call_connected_namespace_runtime(
    namespace,
    session_id="",
    caller_num="",
):
    peer_id = str(session_id or "").strip()
    caller_text = str(caller_num or "").strip()
    if _is_retired_peer_call(namespace, peer_id):
        return False
    binding = _get_peer_call_binding(namespace)
    if peer_id:
        if binding is None:
            namespace.__setitem__(
                "call_popup_pending_peer_connected",
                {
                    "session_id": peer_id,
                    "caller_num": caller_text,
                    "expires_at": _peer_now(namespace) + PEER_CALL_BINDING_TTL_SECONDS,
                },
            )
            return True
        if binding["peer_session_id"] != peer_id:
            return False
    elif binding is not None:
        return False
    active_session = str(namespace.get("call_popup_active_session_id") or "").strip()
    popup = namespace.get("current_call_popup")
    popup_caller = str(getattr(popup, "_call_popup_caller_num", "") or "").strip()
    if caller_text and popup_caller and caller_text != popup_caller:
        return False
    if not active_session:
        if peer_id:
            namespace.__setitem__(
                "call_popup_pending_peer_connected",
                {
                    "session_id": peer_id,
                    "caller_num": caller_text,
                    "expires_at": _peer_now(namespace) + PEER_CALL_BINDING_TTL_SECONDS,
                },
            )
            return True
        return False
    if peer_id:
        binding = _bind_peer_to_local_session(namespace, active_session, caller_text)
        if binding is None or binding["peer_session_id"] != peer_id:
            return False
        if binding["local_session_id"] != active_session:
            return False
    return mark_call_popup_connected_namespace_runtime(
        namespace,
        active_session,
        caller_text,
    )


def set_call_popup_namespace_runtime(namespace, window, *, call_session_id=None):
    namespace.__setitem__("current_call_popup", window)
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
    namespace.__setitem__("current_dial_popup", window)


def mark_dial_popup_connected_namespace_runtime(namespace):
    def mark_on_ui():
        popup = namespace.get("current_dial_popup")
        if popup is None:
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
    namespace,
    call_session_id="",
    caller_num="",
):
    requested_session_id = str(call_session_id or "").strip()
    caller_text = str(caller_num or "").strip()
    popup = namespace.get("current_call_popup")
    popup_caller = str(
        getattr(popup, "_call_popup_caller_num", "") or ""
    ).strip() if popup is not None else ""
    active_session_id = str(
        namespace.get("call_popup_active_session_id", "") or ""
    ).strip()
    if requested_session_id and active_session_id and requested_session_id != active_session_id:
        binding = _get_peer_call_binding(namespace)
        if (
            binding is None
            or binding["peer_session_id"] != requested_session_id
            or binding["local_session_id"] not in ("", active_session_id)
        ):
            # The ATA command path predates peer generation binding.  It is
            # safe to retain its bounded caller match because it is an
            # explicit answer operation, not a passive terminal frame.
            if not caller_text or not popup_caller or caller_text != popup_caller:
                return False
            requested_session_id = active_session_id
        else:
            requested_session_id = active_session_id
    if requested_session_id and not active_session_id:
        # No active generation means this state is late.  Never retain it for
        # an unrelated future popup.
        return False
    # Record the indication before posting the UI callback.  This closes a
    # race where the serial thread observes CALL=1 while the popup creation
    # task is still queued, and also lets popup creation recover if the UI
    # queue is temporarily saturated.
    if requested_session_id:
        namespace.__setitem__(
            "call_popup_pending_connected_session_id",
            requested_session_id,
        )

    def mark_on_ui():
        current_active_session_id = str(
            namespace.get("call_popup_active_session_id", "") or ""
        ).strip()
        session_id = requested_session_id or current_active_session_id
        if (
            requested_session_id
            and current_active_session_id
            and requested_session_id != current_active_session_id
        ):
            return False
        popup = namespace.get("current_call_popup")
        if popup is None:
            if session_id:
                namespace.__setitem__(
                    "call_popup_pending_connected_session_id", session_id
                )
                return True
            return False
        try:
            if not popup.winfo_exists():
                if session_id:
                    namespace.__setitem__(
                        "call_popup_pending_connected_session_id", session_id
                    )
                    return True
                return False
        except Exception:
            return False
        popup_session_id = str(
            getattr(popup, "_call_popup_session_id", current_active_session_id) or ""
        ).strip()
        if session_id and popup_session_id and session_id != popup_session_id:
            return False
        marker = getattr(popup, "_call_popup_mark_connected", None)
        if not callable(marker):
            if session_id:
                namespace.__setitem__(
                    "call_popup_pending_connected_session_id", session_id
                )
                return True
            return False
        try:
            marker()
            handled = namespace.get("mark_incoming_call_handled")
            if callable(handled):
                handled()
            namespace.__setitem__("call_popup_pending_connected_session_id", "")
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


def finish_dial_popup_namespace_runtime(namespace, message=""):
    def finish_on_ui():
        popup = namespace.get("current_dial_popup")
        if popup is None:
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


def start_incoming_call_session_namespace_runtime(namespace, caller_num):
    return namespace["INCOMING_CALL_SESSION"].start(caller_num)


def mark_incoming_call_handled_namespace_runtime(namespace):
    return namespace["INCOMING_CALL_SESSION"].mark_handled()


def _has_local_incoming_call(namespace):
    if namespace.get("call_popup_peer_local_session_id"):
        return True
    snapshot = getattr(namespace["INCOMING_CALL_SESSION"], "snapshot", None)
    return callable(snapshot) and bool(snapshot().caller_num)


def finish_incoming_call_session_namespace_runtime(namespace):
    had_local_call = _has_local_incoming_call(namespace)
    result = namespace["INCOMING_CALL_SESSION"].finish()
    # A peer terminal may have already settled this call. Its delayed serial
    # terminal must leave a newly announced, unbound peer generation intact.
    if had_local_call:
        _clear_peer_call_state(namespace, retire_current=True, preserve_next=True)
    return result


def finish_remote_incoming_call_namespace_runtime(
    namespace,
    phase="ended",
    reason="",
    *,
    session_id="",
    caller_num="",
):
    """Settle a local incoming popup from a peer device channel.

    This only updates the desktop UI/session tracker.  It deliberately does
    not send ATH because the peer modem channel already observed the terminal
    state and the two transports share the same physical modem.
    """
    requested_session_id = str(session_id or "").strip()
    caller_text = str(caller_num or "").strip()
    if _is_retired_peer_call(namespace, requested_session_id):
        return False
    active_session_id = str(
        namespace.get("call_popup_active_session_id", "") or ""
    ).strip()
    popup = namespace.get("current_call_popup")
    popup_session_id = str(
        getattr(popup, "_call_popup_session_id", "") or ""
    ).strip() if popup is not None else ""
    popup_caller = str(
        getattr(popup, "_call_popup_caller_num", "") or ""
    ).strip() if popup is not None else ""
    snapshot = namespace["INCOMING_CALL_SESSION"].snapshot()
    tracked_caller = str(snapshot.caller_num or "").strip()
    current_caller = popup_caller or tracked_caller
    peer_binding = _get_peer_call_binding(namespace)
    if not current_caller:
        if requested_session_id and caller_text:
            namespace.__setitem__(
                "call_popup_pending_peer_terminal",
                {
                    "phase": str(phase or "ended"),
                    "reason": str(reason or ""),
                    "session_id": requested_session_id,
                    "caller_num": caller_text,
                    "expires_at": _peer_now(namespace) + PEER_CALL_BINDING_TTL_SECONDS,
                },
            )
            return True
        return False
    if caller_text and caller_text != current_caller:
        return False
    if requested_session_id:
        if peer_binding is None:
            namespace.__setitem__(
                "call_popup_pending_peer_terminal",
                {
                    "phase": str(phase or "ended"),
                    "reason": str(reason or ""),
                    "session_id": requested_session_id,
                    "caller_num": caller_text,
                    "expires_at": _peer_now(namespace) + PEER_CALL_BINDING_TTL_SECONDS,
                },
            )
            return True
        if peer_binding["peer_session_id"] != requested_session_id:
            return False
        bound_local = peer_binding["local_session_id"]
        if bound_local and active_session_id and bound_local != active_session_id:
            return False
        if not bound_local and active_session_id:
            _set_peer_call_binding(
                namespace,
                requested_session_id,
                peer_binding["caller_num"] or caller_text,
                active_session_id,
            )
    elif peer_binding is not None:
        # Do not accept a legacy no-ID terminal while the current call has a
        # known peer generation.  Both sides must omit IDs for the old
        # same-number compatibility path to apply.
        return False
    if requested_session_id and not active_session_id and popup_session_id:
        if peer_binding is None or peer_binding["local_session_id"] not in ("", popup_session_id):
            return False

    expected_local_session = active_session_id or popup_session_id

    def finish_on_ui():
        current_popup = namespace.get("current_call_popup")
        current_active = str(
            namespace.get("call_popup_active_session_id", "") or ""
        ).strip()
        current_popup_session = str(
            getattr(current_popup, "_call_popup_session_id", "") or ""
        ).strip() if current_popup is not None else ""
        current_popup_caller = str(
            getattr(current_popup, "_call_popup_caller_num", "") or ""
        ).strip() if current_popup is not None else ""
        live_snapshot = namespace["INCOMING_CALL_SESSION"].snapshot()
        # A terminal frame can wait in the UI queue while the same caller redials.
        if expected_local_session and (current_active or current_popup_session) != expected_local_session:
            return False
        if getattr(live_snapshot, "started_at", None) != getattr(snapshot, "started_at", None):
            return False
        live_caller = current_popup_caller or str(live_snapshot.caller_num or "").strip()
        if not live_caller or (caller_text and caller_text != live_caller):
            return False
        if requested_session_id:
            current_binding = _get_peer_call_binding(namespace)
            if current_binding is None or current_binding["peer_session_id"] != requested_session_id:
                return False
            if current_binding["local_session_id"] and current_active and current_binding["local_session_id"] != current_active:
                return False
        elif _get_peer_call_binding(namespace) is not None:
            return False
        missed_call = namespace["INCOMING_CALL_SESSION"].finish()
        namespace.__setitem__("call_popup_active_session_id", "")
        namespace.__setitem__("call_popup_request_token", None)
        namespace.__setitem__("call_popup_pending_connected_session_id", "")
        _clear_peer_call_state(namespace, retire_current=True, preserve_next=True)
        # Close only the incoming popup.  The normal close_call_popup binding
        # also tears down an outgoing dial popup, which is a separate session
        # and must not be touched by a peer incoming terminal frame.
        close_call_popup_runtime(
            current_popup,
            lambda window: namespace.__setitem__("current_call_popup", window),
            log_error=namespace.get("log_file_only"),
        )
        if missed_call is not None:
            show_missed = namespace.get("show_missed_call_popup")
            if callable(show_missed):
                show_missed(missed_call)
        return True

    return namespace["run_on_ui_thread"](finish_on_ui, namespace["ui_post"])


def reset_incoming_call_session_namespace_runtime(namespace):
    had_local_call = _has_local_incoming_call(namespace)
    result = namespace["INCOMING_CALL_SESSION"].reset()
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
    if not namespace.get("CALL_POPUP_ENABLED", True):
        return "disabled"
    session_id = str(call_session_id or "").strip()
    request_token = object()
    namespace.__setitem__("call_popup_request_token", request_token)
    pending_session_id = str(
        namespace.get("call_popup_pending_connected_session_id", "") or ""
    ).strip()
    namespace.__setitem__("call_popup_active_session_id", session_id)
    if session_id:
        _bind_peer_to_local_session(namespace, session_id, caller_num)
    # A connected indication can be queued before this UI task gets a chance
    # to create the popup.  Keep a matching pending marker so set_popup() can
    # apply it immediately after the window is registered; discard only a
    # marker belonging to a different call generation.
    if pending_session_id and pending_session_id != session_id:
        namespace.__setitem__("call_popup_pending_connected_session_id", "")
    popup_holder = {}
    snapshot_call = getattr(namespace.get("INCOMING_CALL_SESSION"), "snapshot", None)
    initial_call = snapshot_call() if callable(snapshot_call) else None

    def is_current():
        return (
            namespace.get("CALL_POPUP_ENABLED", True)
            and namespace.get("call_popup_request_token") is request_token
            and namespace.get("call_popup_active_session_id", "") == session_id
        )

    def is_same_call():
        if initial_call is not None and getattr(initial_call, "started_at", None) is not None:
            current_call = snapshot_call()
            return bool(initial_call.caller_num) and (
                current_call.caller_num == initial_call.caller_num
                and current_call.started_at == initial_call.started_at
            )
        return is_current()

    def dispatch_call_result(callback):
        def apply_on_ui():
            # Hiding a window does not end its modem call. Use the tracked
            # call lifetime, and recheck after any wait in the UI queue.
            if is_same_call():
                return callback()
            return False
        return namespace["run_on_ui_thread"](apply_on_ui, namespace["ui_post"])

    def set_ring_timeout(value):
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
        get_serial=lambda: namespace["serial_obj"],
        port_ui=namespace["port_ui"],
        set_status=namespace["set_status"],
        ui_post=namespace["ui_post"],
        close_popup=close_popup,
        set_ring_timeout=set_ring_timeout,
        run_on_ui_thread=namespace["run_on_ui_thread"],
        is_enabled=is_current,
        mark_call_handled=lambda: mark_incoming_call_handled_namespace_runtime(namespace),
        log_error=namespace.get("log_file_only"),
        dispatch_call_result=dispatch_call_result,
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
    return namespace["ring_timeout_target"], namespace["current_dial_num"]


def set_serial_call_state_namespace_runtime(namespace, next_ring_timeout, next_dial_num):
    namespace.__setitem__("ring_timeout_target", next_ring_timeout)
    namespace.__setitem__("current_dial_num", next_dial_num)
