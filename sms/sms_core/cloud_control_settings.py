from dataclasses import dataclass

from sms_core.cloud_protocol import normalize_cloud_ws_url
from sms_core.credential_policy import CredentialValidationError, credential_url_error, validate_credential


@dataclass(frozen=True)
class CloudControlFormValues:
    enabled: bool
    url: str
    reconnect_interval: object
    device_secret: str
    auto_upload: bool


@dataclass(frozen=True)
class CloudControlSettings:
    enabled: bool = False
    url: str = ""
    reconnect_interval: int = 5
    device_secret: str = ""
    auto_upload: bool = False


def cloud_control_state(enabled, auto_upload, url, device_secret, reconnect_interval):
    return {
        "enabled": bool(enabled),
        "auto_upload": bool(auto_upload),
        "url": str(url or ""),
        "secret": str(device_secret or ""),
        "reconnect_interval": reconnect_interval,
    }


def cloud_control_form_values(values, enabled_override=None):
    enabled, url, interval, secret, auto_upload = values
    if enabled_override is not None:
        enabled = bool(enabled_override)
    return CloudControlFormValues(
        bool(enabled),
        str(url or ""),
        interval,
        str(secret or ""),
        bool(auto_upload),
    )


def cloud_control_save_kwargs(values, enabled_override=None):
    form = cloud_control_form_values(values, enabled_override=enabled_override)
    return {
        "enabled": form.enabled,
        "url": form.url,
        "reconnect_interval": form.reconnect_interval,
        "device_secret": form.device_secret,
        "auto_upload": form.auto_upload,
    }


def normalize_reconnect_interval(value, default=5):
    try:
        return max(1, int(value))
    except Exception:
        return int(default)


def read_cloud_control_settings(config):
    device_secret = config.get("cloud_control", "device_secret", fallback="", raw=True)
    validate_credential(device_secret, '云端控制密码')
    raw_url = config.get("cloud_control", "url", fallback="", raw=True)
    error = credential_url_error(raw_url, 'WebSocket地址')
    if error:
        raise CredentialValidationError(error)
    try:
        enabled = config.getboolean("cloud_control", "enabled", fallback=False)
    except Exception:
        enabled = False

    try:
        url = normalize_cloud_ws_url(raw_url)
    except Exception:
        url = ""

    try:
        reconnect_interval = normalize_reconnect_interval(
            config.getint("cloud_control", "reconnect_interval", fallback=5)
        )
    except Exception:
        reconnect_interval = 5

    try:
        auto_upload = config.getboolean("cloud_control", "auto_upload", fallback=False)
    except Exception:
        auto_upload = False

    return CloudControlSettings(
        enabled=enabled,
        url=url,
        reconnect_interval=reconnect_interval,
        device_secret=device_secret,
        auto_upload=auto_upload,
    )


def update_cloud_control_settings(
    current,
    *,
    enabled=None,
    url=None,
    reconnect_interval=None,
    device_secret=None,
    auto_upload=None,
):
    secret = current.device_secret if device_secret is None else str(device_secret)
    validate_credential(secret, '云端控制密码')
    next_url = current.url if url is None else url
    error = credential_url_error(next_url, 'WebSocket地址')
    if error:
        raise CredentialValidationError(error)
    return CloudControlSettings(
        enabled=current.enabled if enabled is None else bool(enabled),
        url=current.url if url is None else normalize_cloud_ws_url(url),
        reconnect_interval=(
            current.reconnect_interval
            if reconnect_interval is None
            else normalize_reconnect_interval(reconnect_interval)
        ),
        device_secret=secret,
        auto_upload=current.auto_upload if auto_upload is None else bool(auto_upload),
    )


def write_cloud_control_settings(config, settings):
    validate_credential(settings.device_secret, '云端控制密码')
    error = credential_url_error(settings.url, 'WebSocket地址')
    if error:
        raise CredentialValidationError(error)
    if "cloud_control" not in config:
        config["cloud_control"] = {}
    config["cloud_control"]["enabled"] = "1" if settings.enabled else "0"
    config["cloud_control"]["url"] = settings.url
    if config.has_option("cloud_control", "device_imei"):
        config.remove_option("cloud_control", "device_imei")
    config["cloud_control"]["device_secret"] = settings.device_secret
    config["cloud_control"]["reconnect_interval"] = str(settings.reconnect_interval)
    config["cloud_control"]["auto_upload"] = "1" if settings.auto_upload else "0"
