import time


class ConsecutiveRepeatNotice:
    def __init__(self, *, limit, suffix):
        self.limit = limit
        self.suffix = suffix
        self.last_message = None
        self.count = 0

    def reset(self):
        self.last_message = None
        self.count = 0

    def next_message(self, message, **_kwargs):
        if self.last_message == message:
            self.count += 1
        else:
            self.last_message = message
            self.count = 1
        return _limited_message(message, self.count, self.limit, self.suffix)


class TimedRepeatNotice:
    def __init__(self, *, limit, reset_seconds, suffix, monotonic=time.monotonic):
        self.limit = limit
        self.reset_seconds = reset_seconds
        self.suffix = suffix
        self.monotonic = monotonic
        self.state = {}

    def clear(self):
        self.state.clear()

    def next_message(self, message, *, repeat_key=""):
        key = str(repeat_key or message)
        now = self.monotonic()
        last_seen, count = self.state.get(key, (0.0, 0))
        if now - last_seen > self.reset_seconds:
            count = 0
        count += 1
        self.state[key] = (now, count)
        return _limited_message(message, count, self.limit, self.suffix)


def _limited_message(message, count, limit, suffix):
    if count < limit:
        return message
    if count == limit:
        return f"{message}{suffix}"
    return None


def emit_repeat_notice(notice, message, system_ui, *, level="normal", **kwargs):
    try:
        text = notice.next_message(message, **kwargs)
        if text is not None:
            system_ui(text, level)
    except Exception:
        try:
            system_ui(message, level)
        except Exception:
            pass
