import threading


MODEM_RESPONSE_TIMEOUT_MARKERS = (
    "等待 modem 指令响应超时",
    "modem response timeout",
)


class CloudModemHealthState:
    """Track consecutive confirmed AT response timeouts."""

    def __init__(self, reconnect_threshold=3):
        self.reconnect_threshold = max(1, int(reconnect_threshold or 3))
        self._lock = threading.Lock()
        self._consecutive_timeouts = 0
        self._unresponsive = False
        self._reconnect_requested = False

    @staticmethod
    def is_response_timeout(message):
        text = str(message or "").strip().lower()
        return any(marker in text for marker in MODEM_RESPONSE_TIMEOUT_MARKERS)

    def record(self, ok, message=""):
        request_reconnect = False
        with self._lock:
            if ok:
                self._consecutive_timeouts = 0
                self._unresponsive = False
                self._reconnect_requested = False
            elif self.is_response_timeout(message):
                self._consecutive_timeouts += 1
                if self._consecutive_timeouts >= self.reconnect_threshold:
                    self._unresponsive = True
                    if not self._reconnect_requested:
                        self._reconnect_requested = True
                        request_reconnect = True
            else:
                self._consecutive_timeouts = 0
                self._unresponsive = False
                self._reconnect_requested = False
            return self._snapshot_unlocked(request_reconnect)

    def snapshot(self):
        with self._lock:
            return self._snapshot_unlocked(False)

    def _snapshot_unlocked(self, request_reconnect):
        return {
            "modem_unresponsive": bool(self._unresponsive),
            "consecutive_at_timeouts": int(self._consecutive_timeouts),
            "request_reconnect": bool(request_reconnect),
        }
