"""Standard SMTP transports sharing one deadline across the whole session."""
import io
import smtplib

from sms_core.http_deadline import _Deadline, _DeadlineReader, create_connection_with_deadline


def _preserve_timeout(exc):
    if isinstance(exc.__cause__, TimeoutError) or isinstance(exc.__context__, TimeoutError):
        raise TimeoutError("SMTP 会话超过总等待时间") from None


class _SmtpDeadlineMixin:
    def __init__(self, *args, timeout=15, **kwargs):
        # Reuse the same per-read deadline implementation as HTTP, including
        # partial lines; a server cannot extend the session by dripping bytes.
        self.deadline = _Deadline(timeout, response_label="SMTP 会话")
        try:
            super().__init__(*args, timeout=self.deadline.remaining(), **kwargs)
        except BaseException:
            self.close()
            raise

    def _get_socket(self, host, port, timeout):
        sock = create_connection_with_deadline((host, port), self.deadline, self.source_address)
        try:
            sock.settimeout(self.deadline.remaining())
            return sock
        except BaseException:
            sock.close()
            raise

    def getreply(self):
        if self.file is None:
            self.file = io.BufferedReader(_DeadlineReader(self.sock, self.deadline))
        try:
            result = super().getreply()
        except smtplib.SMTPServerDisconnected as exc:
            _preserve_timeout(exc)
            self.deadline.remaining()
            raise
        self.sock.settimeout(self.deadline.remaining())
        return result

    def send(self, data):
        remaining = self.deadline.remaining()
        if self.sock is not None:
            self.sock.settimeout(remaining)
        try:
            return super().send(data)
        except smtplib.SMTPServerDisconnected as exc:
            _preserve_timeout(exc)
            raise

    def starttls(self, *args, **kwargs):
        result = super().starttls(*args, **kwargs)
        self.sock.settimeout(self.deadline.remaining())
        return result


class DeadlineSMTP(_SmtpDeadlineMixin, smtplib.SMTP):
    pass


class DeadlineSMTP_SSL(_SmtpDeadlineMixin, smtplib.SMTP_SSL):
    def _get_socket(self, host, port, timeout):
        sock = create_connection_with_deadline((host, port), self.deadline, self.source_address)
        try:
            # TCP and TLS share the budget; do not give the handshake the
            # timeout originally selected before the TCP connection finished.
            sock.settimeout(self.deadline.remaining())
            sock = self.context.wrap_socket(sock, server_hostname=self._host)
            sock.settimeout(self.deadline.remaining())
            return sock
        except BaseException:
            sock.close()
            raise
