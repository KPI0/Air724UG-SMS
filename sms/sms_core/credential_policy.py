"""Shared credential checks; values are validated without trimming or rewriting."""
import re
from urllib.parse import unquote, unquote_plus


CREDENTIAL_RULE = "只能包含数字、英文字母和英文符号，不能包含中文、emoji、空白或控制字符"
PUSH_CREDENTIAL_FIELDS = frozenset((
    'pushdeer_key', 'bark_key', 'dingtalk_secret', 'pushover_api_token',
    'pushover_user_key', 'gotify_token', 'wxpusher_app_token',
    'email_password', 'next_smtp_proxy_password',
))
PUSH_URL_FIELDS = {
    'custom_post_url': '', 'telegram_api': 'telegram', 'pushdeer_api': '',
    'bark_api': '', 'dingtalk_webhook': '', 'feishu_webhook': 'feishu',
    'wecom_webhook': '', 'inotify_api': '', 'next_smtp_proxy_api': '',
    'gotify_api': '', 'serverchan_api': 'serverchan',
}
URL_CREDENTIAL_NAMES = frozenset((
    'access_token', 'auth_token', 'token', 'key', 'secret', 'password', 'api_key', 'apikey', 'app_key', 'appkey',
))
URL_PARTS = re.compile(r'^\s*[A-Za-z][A-Za-z0-9+.-]*://([^/?#]*)([^?#]*)(?:\?([^#]*))?(?:#.*)?$', re.DOTALL)
PROVIDER_TOKEN_PATHS = {
    'telegram': re.compile(r'^/bot([^/]+)(?:/|$)'),
    'feishu': re.compile(r'^/open-apis/bot/v2/hook/([^/]+)(?:/|$)'),
    'serverchan': re.compile(r'^/([^/]+)\.send$'),
}


class CredentialValidationError(ValueError):
    pass


def credential_error(value, label, allow_empty=True):
    text = str(value or '')
    if not text and allow_empty:
        return ''
    if not text or any(not 33 <= ord(char) <= 126 for char in text):
        return f'{label}{CREDENTIAL_RULE}'
    return ''


def validate_credential(value, label, allow_empty=True):
    error = credential_error(value, label, allow_empty)
    if error:
        raise CredentialValidationError(error)


def credential_url_error(value, label, provider=''):
    matched = URL_PARTS.match(str(value or ''))
    if matched is None:
        return ''  # The existing URL policy handles malformed addresses.
    authority, path, query = matched.groups()
    values = []
    if '@' in authority:
        userinfo = authority.rsplit('@', 1)[0]
        if ':' in userinfo:
            values.append((userinfo.split(':', 1)[1], False))
    for field in (query or '').split('&'):
        key, separator, raw = field.partition('=')
        if separator and unquote_plus(key).lower() in URL_CREDENTIAL_NAMES:
            values.append((raw, True))
    token_pattern = PROVIDER_TOKEN_PATHS.get(provider)
    token = token_pattern.match(path) if token_pattern is not None else None
    if token:
        values.append((token.group(1), False))
    for raw, is_query in values:
        try:
            decoded = (unquote_plus if is_query else unquote)(raw, errors='strict')
        except UnicodeError:
            return f'{label}中的密钥或密码{CREDENTIAL_RULE}'
        error = credential_error(decoded, f'{label}中的密钥或密码')
        if error:
            return error
    return ''


def push_credential_errors(settings):
    errors = [credential_error(settings.get(key, ''), key.upper())
              for key in sorted(PUSH_CREDENTIAL_FIELDS)]
    errors.extend(credential_url_error(settings.get(key, ''), key.upper(), provider)
                  for key, provider in PUSH_URL_FIELDS.items())
    return [error for error in errors if error]


def validate_push_credentials(settings):
    errors = push_credential_errors(settings)
    if errors:
        raise CredentialValidationError(errors[0])


def validate_config_credentials(snapshot):
    cloud = snapshot.get('cloud_control', {})
    validate_credential(cloud.get('device_secret', ''), '云端控制密码')
    error = credential_url_error(cloud.get('url', ''), 'WebSocket地址')
    if error:
        raise CredentialValidationError(error)
    validate_push_credentials(snapshot.get('third_push', {}))


def _split_at_fields(text, separators, with_separator=False):
    start, quoted, escaped = 0, False, False
    for index, char in enumerate(text):
        if escaped:
            escaped = False
        elif char == '\\' and quoted:
            escaped = True
        elif char == '"':
            quoted = not quoted
        elif char in separators and not quoted:
            yield (text[start:index], char) if with_separator else text[start:index]
            start = index + 1
    yield (text[start:], '') if with_separator else text[start:]


def _at_parameter_value(raw):
    text = raw.lstrip(' \t\r\n')  # A quoted tail may still belong to the password.
    if not text.startswith('"'):
        return text.rstrip(' \t\r\n')
    escaped = False
    for index in range(1, len(text)):
        char = text[index]
        if escaped:
            escaped = False
        elif char == '\\':
            escaped = True
        elif char == '"':
            return text[1:index] if not text[index + 1:].strip(' \t\r\n') else text
    return text


def serial_credential_error(command):
    at_context = False
    for statement, separator in _split_at_fields(str(command or ''), ';\r\n', with_separator=True):
        if re.match(r'^[ \t]*AT', statement, re.IGNORECASE):
            at_context = True
        elif at_context and re.match(r'^[ \t]*\+', statement):
            statement = 'AT' + statement.lstrip(' \t')
        else:
            at_context = False
        matched = re.match(r'^[ \t\r\n]*AT[ \t\r\n]*\+[ \t\r\n]*(CPIN2?|CPWD|CLCK|CPNETAPN)[ \t\r\n]*=(.*)$',
                           statement, re.IGNORECASE | re.DOTALL)
        at_context = at_context and separator == ';'
        if matched is None or matched.group(2).strip(' \t\r\n') == '?':
            continue
        name, payload = matched.groups()
        fields = list(_split_at_fields(payload, ','))
        indexes = {'CPIN': (0, 1), 'CPIN2': (0, 1), 'CPWD': (1, 2), 'CLCK': (2,), 'CPNETAPN': (3,)}[name.upper()]
        for index in indexes:
            if index < len(fields):
                error = credential_error(_at_parameter_value(fields[index]), 'AT指令中的密码')
                if error:
                    return error
    return ''
