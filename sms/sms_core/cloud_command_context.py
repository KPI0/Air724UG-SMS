from dataclasses import dataclass


_CURRENT_WEBSOCKET = object()


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


def capture_cloud_serial_command_context(namespace, *, websocket=_CURRENT_WEBSOCKET):
    with namespace["serial_lock"]:
        return CloudSerialCommandContext(
            serial=namespace.get("serial_obj"),
            generation=namespace.get("serial_connection_generation", 0),
            imei=str(namespace["_cloud_runtime_imei"]() or ""),
            websocket=namespace.get("cloud_ws_conn") if websocket is _CURRENT_WEBSOCKET else websocket,
            authorized=bool(namespace.get("cloud_device_authorized")),
            secret=namespace.get("CLOUD_DEVICE_SECRET", ""),
            ws_url=namespace.get("CLOUD_WS_URL", ""),
        )
