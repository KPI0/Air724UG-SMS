from dataclasses import dataclass
from sms_core.call_command_context import call_state_lock, capture_call_command_context


_CURRENT_WEBSOCKET = object()


def set_cloud_command_authorized(namespace, value):
    authorized = bool(value)
    with call_state_lock(namespace):
        if not authorized:
            namespace["_cloud_authorization_epoch"] = object()
            pending = namespace.pop("_cloud_pending_dial", None)
            if pending is not None:
                pending.ready.set()
        namespace["cloud_device_authorized"] = authorized


@dataclass(frozen=True, repr=False)
class CloudSerialCommandContext:
    """The connection that admitted a cloud request, never a moving target."""

    serial: object
    generation: int
    imei: str
    websocket: object
    authorized: bool
    secret: str
    ws_url: str
    call_context: object = None
    authorization_epoch: object = None

    def rejection_reason(self, namespace):
        # The caller holds serial_lock, also used by each actual serial write.
        if (
            self.serial is None
            or namespace.get("serial_obj") is not self.serial
            or namespace.get("serial_connection_generation", 0) != self.generation
            or not self.imei
            or str(namespace["_cloud_runtime_imei"]() or "") != self.imei
        ):
            return "串口连接或设备身份已变化，已取消旧云端指令，请重新提交"
        if (
            not self.authorized
            or namespace.get("_cloud_authorization_epoch") is not self.authorization_epoch
            or not namespace.get("cloud_device_authorized")
            or not namespace.get("cloud_connected")
            or not namespace.get("CLOUD_CONTROL_ENABLED", True)
            or self.websocket is None
            or namespace.get("cloud_ws_conn") is not self.websocket
            or namespace.get("CLOUD_DEVICE_SECRET", "") != self.secret
            or namespace.get("CLOUD_WS_URL", "") != self.ws_url
        ):
            return "云端连接或授权已变化，已取消旧云端指令，请重新提交"
        return ""


def capture_cloud_serial_command_context(
    namespace, *, websocket=_CURRENT_WEBSOCKET, command="", command_data=None,
):
    with namespace["serial_lock"], call_state_lock(namespace):
        return CloudSerialCommandContext(
            serial=namespace.get("serial_obj"),
            generation=namespace.get("serial_connection_generation", 0),
            imei=str(namespace["_cloud_runtime_imei"]() or ""),
            websocket=namespace.get("cloud_ws_conn") if websocket is _CURRENT_WEBSOCKET else websocket,
            authorized=bool(namespace.get("cloud_device_authorized")),
            secret=namespace.get("CLOUD_DEVICE_SECRET", ""),
            ws_url=namespace.get("CLOUD_WS_URL", ""),
            call_context=capture_call_command_context(namespace, command, command_data),
            authorization_epoch=namespace.setdefault("_cloud_authorization_epoch", object()),
        )
