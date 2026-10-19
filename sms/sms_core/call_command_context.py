"""Bind queued modem controls to a call lifetime, independently of its UI."""
import re
import inspect
import threading
from contextlib import contextmanager
from dataclasses import dataclass, field

from sms_core.serial_parsers import is_clip_line, is_hangup_event, is_ring_line


VOICE_DIAL_RE = re.compile(r"ATD(\+?[0-9]+);?", re.IGNORECASE)
CALL_METADATA_KEYS = (
    "call_control_guard", "call_direction", "call_session_id", "call_dial_id",
    "call_dial_task_id", "call_phone", "local_call_session_id", "call_session_ids",
)
CALL_CHANGED_MESSAGE = "通话已结束或已变化，已取消旧通话指令"


def call_state_lock(namespace):
    lock = namespace.get("_call_peer_state_lock")
    return lock if lock is not None else namespace.setdefault("_call_peer_state_lock", threading.RLock())


def _incoming_token(namespace):
    snapshot = getattr(namespace.get("INCOMING_CALL_SESSION"), "snapshot", None)
    return getattr(snapshot(), "session_token", None) if callable(snapshot) else None


def invalidate_call_commands(namespace):
    """Caller may already hold the reentrant lifetime lock."""
    with call_state_lock(namespace):
        namespace["_call_command_epoch"] = object()
        namespace["_cloud_dial_lifetime"] = None
        namespace["_call_command_pending_incoming"] = False
        pending = namespace.pop("_cloud_pending_dial", None)
        if pending is not None:
            pending.ready.set()


def observe_call_command_line(namespace, line, *, call_state_snapshot=None):
    # Run before response waiters and UI/log callbacks can yield to a sender.
    with call_state_lock(namespace):
        if call_state_snapshot is not None and not call_state_snapshot.is_current():
            return
        outgoing = namespace.get("_cloud_dial_lifetime")
        if is_hangup_event(line):
            invalidate_call_commands(namespace)
        elif is_ring_line(line) or is_clip_line(line):
            if outgoing or (_incoming_token(namespace) is None
                            and not namespace.get("_call_command_pending_incoming")):
                invalidate_call_commands(namespace)
                namespace["_call_command_pending_incoming"] = True


@dataclass(repr=False)
class LocalDialRequest:
    number: str
    popup: object = None
    queued: bool = False
    submitted: bool = False
    finished: bool = False


def set_local_dial_number(namespace, number):
    with call_state_lock(namespace):
        request = LocalDialRequest(str(number)) if number else None
        namespace["_local_dial_request"] = request
        namespace["_dial_state_generation"] = request if request is not None else object()
        namespace["current_dial_num"] = number


def set_local_dial_popup(namespace, popup):
    with call_state_lock(namespace):
        namespace["current_dial_popup"] = popup
        request = namespace.get("_local_dial_request")
        if request is namespace.get("_dial_state_generation") and request is not None:
            request.popup = popup


class SerialCallStateSnapshot(tuple):
    """Keep the legacy two values while retaining the originating call identity."""

    def __new__(cls, namespace):
        with call_state_lock(namespace):
            result = super().__new__(cls, (
                namespace["ring_timeout_target"], namespace["current_dial_num"],
            ))
            result.namespace = namespace
            result.generation = namespace.get("_dial_state_generation")
            result.incoming_token = _incoming_token(namespace)
            request = namespace.get("_local_dial_request")
            result.request_phase = (
                (request.submitted, request.finished)
                if request is not None and request is result.generation else None)
            result.pending = bool(
                request is not None and request is result.generation
                and not request.submitted and not request.finished)
            result.popup = (
                request.popup if request is not None and request is result.generation
                and not result.pending else None)
            if result.generation is None:
                result.popup = namespace.get("current_dial_popup")
            return result

    def is_current(self, *, check_incoming=True):
        with call_state_lock(self.namespace):
            request = self.namespace.get("_local_dial_request")
            phase = ((request.submitted, request.finished)
                     if request is not None and request is self.generation else None)
            return (self.namespace.get("_dial_state_generation") is self.generation
                    and phase == self.request_phase
                    and (not check_incoming or not self[1]
                         or _incoming_token(self.namespace) is self.incoming_token))

    def same_generation(self, other):
        return (isinstance(other, SerialCallStateSnapshot)
                and self.generation is other.generation
                and self.request_phase == other.request_phase)

    def is_response_current(self):
        with call_state_lock(self.namespace):
            return (self.is_current(check_incoming=False)
                    and _incoming_token(self.namespace) is self.incoming_token)

    def commit(self, ring_timeout, dial_number):
        with call_state_lock(self.namespace):
            if not self.is_current() or self.pending:
                return False
            self.namespace["ring_timeout_target"] = ring_timeout
            self.namespace["current_dial_num"] = dial_number
            return True


def call_with_state_snapshot(
    callback, *args, call_state_snapshot=None, check_current=True, check_incoming=True, **kwargs,
):
    if call_state_snapshot is not None:
        if check_current and not call_state_snapshot.is_current(check_incoming=check_incoming):
            return None
        try:
            inspect.signature(callback).bind_partial(*args, **kwargs, call_state_snapshot=call_state_snapshot)
        except (TypeError, ValueError):
            pass
        else:
            return callback(*args, **kwargs, call_state_snapshot=call_state_snapshot)
    return callback(*args, **kwargs)


@dataclass(repr=False)
class PendingDial:
    lifetime: tuple
    committed_epoch: object = None
    committed_generation: object = None
    ready: object = field(default_factory=threading.Event)


@dataclass(repr=False)
class CallCommandContext:
    namespace: dict
    command: str
    epoch: object
    incoming_token: object
    outgoing_lifetime: object
    direction: str = ""
    session_id: str = ""
    dial_id: str = ""
    dial_number: str = ""
    error: str = ""
    written: bool = False
    pending_dial: object = None
    dial_generation: object = None
    local_dial_request: object = None

    def rejection_reason(self):
        namespace = self.namespace
        if self.error:
            return self.error
        epoch, outgoing = self.epoch, self.outgoing_lifetime
        generation = self.dial_generation
        if self.pending_dial is not None:
            if self.pending_dial.committed_epoch is not None:
                epoch = self.pending_dial.committed_epoch
                outgoing = self.pending_dial.lifetime
                generation = self.pending_dial.committed_generation
            elif namespace.get("_cloud_pending_dial") is not self.pending_dial:
                return CALL_CHANGED_MESSAGE
        if (namespace.get("_call_command_epoch") is not epoch
                or namespace.get("_dial_state_generation") is not generation
                or _incoming_token(namespace) is not self.incoming_token
                or namespace.get("_cloud_dial_lifetime") is not outgoing):
            return CALL_CHANGED_MESSAGE
        if self.direction == "incoming":
            pair = namespace.get("call_popup_local_session")
            if (not self.incoming_token or pair != (self.incoming_token, self.session_id)
                    or namespace.get("current_dial_num")):
                return CALL_CHANGED_MESSAGE
        if self.direction == "outgoing" and not self.dial_number:
            pending = self.pending_dial is not None and self.pending_dial.committed_epoch is None
            if not pending and (not outgoing or outgoing[0] != self.dial_id):
                return CALL_CHANGED_MESSAGE
        return ""

    def admit(self):
        # Only authenticated commands entering the worker queue may reserve a
        # dial. Capture before replay/auth awaits never reserves a dial.
        with call_state_lock(self.namespace):
            reason = self.rejection_reason()
            if reason:
                return reason
            if self.direction == "outgoing" and self.dial_number and self.pending_dial is None:
                if self.namespace.get("_cloud_pending_dial") is not None:
                    return "已有拨号指令正在排队，请等待其完成"
                self.pending_dial = PendingDial((self.dial_id, object()))
                self.namespace["_cloud_pending_dial"] = self.pending_dial
            return ""

    def wait_for_dial(self, timeout):
        if self.direction == "outgoing" and not self.dial_number and self.pending_dial is not None:
            if not self.pending_dial.ready.wait(timeout):
                return "原拨号尚未提交，挂断等待超时，请重试"
        with call_state_lock(self.namespace):
            return self.rejection_reason()

    def is_current(self):
        with call_state_lock(self.namespace):
            return not self.rejection_reason()

    @contextmanager
    def write_guard(self):
        with call_state_lock(self.namespace):
            pending_hangup = (self.direction == "outgoing" and not self.dial_number
                              and self.pending_dial is not None
                              and self.pending_dial.committed_epoch is None)
            yield not pending_hangup and not self.rejection_reason()

    def before_write(self):
        # Invoked only after all send locks, OTA checks and waiter readiness.
        if self.dial_number:
            pending = self.namespace.pop("_cloud_pending_dial", None)
            if pending is not None and pending is not self.pending_dial:
                pending.ready.set()
            self.outgoing_lifetime = (
                self.pending_dial.lifetime if self.pending_dial is not None else (self.dial_id, object()))
            self.epoch = object()
            self.namespace["_cloud_dial_lifetime"] = self.outgoing_lifetime
            self.namespace["_call_command_epoch"] = self.epoch
            self.namespace["_call_command_pending_incoming"] = False
            self.dial_generation = self.local_dial_request or object()
            self.namespace["_dial_state_generation"] = self.dial_generation
            if self.local_dial_request is not None:
                self.local_dial_request.submitted = True
            self.namespace["current_dial_num"] = self.dial_number
            if self.pending_dial is not None:
                self.pending_dial.committed_epoch = self.epoch
                self.pending_dial.committed_generation = self.dial_generation
                self.pending_dial.ready.set()
        self.written = True
        return True

    def finish(self, ok):
        if ok or not self.dial_number:
            return
        with call_state_lock(self.namespace):
            if (self.local_dial_request is not None
                    and self.namespace.get("_dial_state_generation") is self.local_dial_request):
                self.local_dial_request.finished = True
                self.namespace["current_dial_num"] = ""
            if (self.pending_dial is not None
                    and self.namespace.get("_cloud_pending_dial") is self.pending_dial):
                self.namespace.pop("_cloud_pending_dial", None)
                self.pending_dial.ready.set()
            if (self.written and self.namespace.get("_cloud_dial_lifetime") is self.outgoing_lifetime
                    and self.namespace.get("_dial_state_generation") is self.dial_generation):
                self.namespace["current_dial_num"] = ""
                invalidate_call_commands(self.namespace)


def capture_call_command_context(namespace, command, metadata=None):
    cmd = str(command or "").strip()
    meta = metadata if isinstance(metadata, dict) else {}
    scoped = any(key in meta for key in CALL_METADATA_KEYS)
    dial = VOICE_DIAL_RE.fullmatch(cmd)
    if not scoped and not dial and not re.fullmatch(r"AT[AH];?", cmd, re.IGNORECASE):
        return None
    with call_state_lock(namespace):
        context = CallCommandContext(
            namespace, cmd, namespace.setdefault("_call_command_epoch", object()),
            _incoming_token(namespace), namespace.get("_cloud_dial_lifetime"),
            dial_number=dial.group(1) if dial else "",
            dial_generation=namespace.get("_dial_state_generation"),
        )
        if not scoped:
            return context
        direction = str(meta.get("call_direction") or "").strip()
        context.direction = direction
        context.session_id = str(meta.get("call_session_id") or "").strip()
        context.dial_id = str(meta.get("call_dial_id") or "").strip()
        pending = namespace.get("_cloud_pending_dial")
        if direction == "outgoing" and not dial and pending is not None and pending.lifetime[0] == context.dial_id:
            context.pending_dial = pending
        if type(meta.get("call_control_guard")) is not int or meta["call_control_guard"] != 1:
            context.error = "通话控制元数据不完整，已拒绝执行"
        elif direction == "incoming":
            if (not context.session_id or context.dial_id
                    or not re.fullmatch(r"AT[AH];?", cmd, re.IGNORECASE)):
                context.error = "来电控制编号或指令无效"
        elif direction == "outgoing":
            if (not context.dial_id or context.session_id
                    or not (dial or re.fullmatch(r"ATH;?", cmd, re.IGNORECASE))):
                context.error = "呼出控制编号或指令无效"
            elif dial and (context.incoming_token or namespace.get("current_dial_num")
                           or namespace.get("_call_command_pending_incoming")):
                context.error = "设备已有通话，已取消新的拨号指令"
        else:
            context.error = "通话控制方向无效"
        context.error = context.rejection_reason()
        return context
