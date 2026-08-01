from dataclasses import dataclass
from datetime import datetime
from threading import Lock


@dataclass(frozen=True)
class MissedCall:
    caller_num: str
    started_at: datetime


@dataclass(frozen=True)
class IncomingCallSessionSnapshot:
    caller_num: str = ""
    started_at: object = None
    handled: bool = False


@dataclass(frozen=True)
class IncomingCallStartResult:
    started: bool
    replaced: bool = False
    replaced_missed_call: object = None

    def __bool__(self):
        return self.started


class IncomingCallSessionTracker:
    """Track one inbound call independently from Tk window lifetime."""

    def __init__(self, now_func=datetime.now):
        self._now_func = now_func
        self._lock = Lock()
        self._caller_num = ""
        self._started_at = None
        self._handled = False

    def start(self, caller_num):
        caller = str(caller_num or "").strip()
        if not caller:
            return IncomingCallStartResult(started=False)
        with self._lock:
            if self._caller_num == caller:
                return IncomingCallStartResult(started=False)

            replaced = bool(self._caller_num)
            replaced_missed_call = None
            if replaced and not self._handled:
                replaced_missed_call = MissedCall(
                    caller_num=self._caller_num,
                    started_at=self._started_at,
                )
            self._caller_num = caller
            self._started_at = self._now_func()
            self._handled = False
            return IncomingCallStartResult(
                started=True,
                replaced=replaced,
                replaced_missed_call=replaced_missed_call,
            )

    def mark_handled(self):
        with self._lock:
            if not self._caller_num:
                return False
            self._handled = True
            return True

    def finish(self):
        with self._lock:
            missed_call = None
            if self._caller_num and not self._handled:
                missed_call = MissedCall(
                    caller_num=self._caller_num,
                    started_at=self._started_at,
                )
            self._caller_num = ""
            self._started_at = None
            self._handled = False
            return missed_call

    def reset(self):
        with self._lock:
            self._caller_num = ""
            self._started_at = None
            self._handled = False

    def snapshot(self):
        with self._lock:
            return IncomingCallSessionSnapshot(
                caller_num=self._caller_num,
                started_at=self._started_at,
                handled=self._handled,
            )
