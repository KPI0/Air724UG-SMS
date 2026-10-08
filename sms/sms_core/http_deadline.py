"""urllib transport with one read deadline across headers, redirects and body."""
import http.client
import io
import math
import time
import urllib.error
import urllib.parse
import urllib.request


class _Deadline:
    def __init__(self, timeout):
        seconds = float(timeout)
        if not math.isfinite(seconds) or seconds <= 0:
            raise ValueError("请求超时必须为正数")
        self.end = time.monotonic() + seconds

    def remaining(self):
        remaining = self.end - time.monotonic()
        if remaining <= 0:
            raise TimeoutError("推送响应超过总等待时间，无法确认是否送达")
        return remaining


class _DeadlineReader(io.RawIOBase):
    def __init__(self, sock, deadline):
        self.sock = sock
        self.deadline = deadline
        self.raw = None
        self.raw = sock.makefile("rb", buffering=0)

    def readable(self):
        return True

    def readinto(self, buffer):
        self.sock.settimeout(self.deadline.remaining())
        return self.raw.readinto(buffer)

    def close(self):
        try:
            if self.raw is not None:
                self.raw.close()
        finally:
            super().close()


class _ResponseSocket:
    def __init__(self, sock, deadline):
        self.sock = sock
        self.deadline = deadline

    def makefile(self, mode):
        if mode != "rb":
            raise ValueError("HTTP response must use a binary reader")
        return io.BufferedReader(_DeadlineReader(self.sock, self.deadline))


class _DeadlineConnection:
    def __init__(self, host, *, deadline, **kwargs):
        self.deadline = deadline
        kwargs["timeout"] = deadline.remaining()
        super().__init__(host, **kwargs)
        self.response_class = lambda sock, **options: http.client.HTTPResponse(
            _ResponseSocket(sock, deadline), **options
        )

    def connect(self):
        super().connect()
        try:
            self.sock.settimeout(self.deadline.remaining())
        except BaseException:
            self.close()
            raise

    def send(self, data):
        remaining = self.deadline.remaining()
        if self.sock is not None:
            self.sock.settimeout(remaining)
        return super().send(data)


class _HTTPConnection(_DeadlineConnection, http.client.HTTPConnection):
    pass


class _HTTPSConnection(_DeadlineConnection, http.client.HTTPSConnection):
    pass


class _HTTPHandler(urllib.request.HTTPHandler):
    def __init__(self, deadline):
        super().__init__()
        self.deadline = deadline

    def http_open(self, request):
        return self.do_open(
            lambda host, **kwargs: _HTTPConnection(host, deadline=self.deadline, **kwargs), request
        )


class _HTTPSHandler(urllib.request.HTTPSHandler):
    def __init__(self, deadline):
        super().__init__()
        self.deadline = deadline

    def https_open(self, request):
        return self.do_open(
            lambda host, **kwargs: _HTTPSConnection(host, deadline=self.deadline, **kwargs),
            request, context=self._context,
        )


class _HTTPRedirectHandler(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, request, fp, code, msg, headers, newurl):
        if urllib.parse.urlsplit(newurl).scheme.lower() not in ("http", "https"):
            raise urllib.error.HTTPError(newurl, code, "不支持的推送重定向协议", headers, fp)
        return super().redirect_request(request, fp, code, msg, headers, newurl)


def urlopen_with_deadline(request, *, timeout=15):
    deadline = _Deadline(timeout)
    # Keep urllib's platform proxies and HTTPS certificate verification. Each
    # redirect reuses this deadline; no helper request is left running on timeout.
    opener = urllib.request.build_opener(
        _HTTPHandler(deadline), _HTTPSHandler(deadline), _HTTPRedirectHandler()
    )
    response = opener.open(request, timeout=deadline.remaining())
    try:
        deadline.remaining()
    except BaseException:
        response.close()
        raise
    return response
