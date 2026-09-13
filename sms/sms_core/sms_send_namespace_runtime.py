"""Bind a manual SMS send and its cloud record to the selected serial device."""
import uuid

from sms_core.cloud_payloads import build_sent_sms_event_payload
from sms_core.cloud_sms_event_runtime import enqueue_cloud_sms_event_runtime
from sms_core.serial_sender import send_text_sms_pdu_async


def send_manual_sms_namespace_runtime(
    namespace, phone, message, *, send_runtime=send_text_sms_pdu_async,
):
    with namespace["serial_lock"]:
        serial = namespace.get("serial_obj")
        generation = namespace.get("serial_connection_generation", 0)
        identity = dict(namespace["_cloud_identity_payload"]())
        imei = str(identity.get("imei") or "").strip()
    source_id = "desktop_sms_" + uuid.uuid4().hex

    def get_captured_serial():
        # The sender invokes this getter under serial_lock before each write.
        # A queued UI action must never move to a newly selected device.
        if (
            namespace.get("serial_obj") is not serial
            or namespace.get("serial_connection_generation", 0) != generation
            or (imei and str(namespace["_cloud_runtime_imei"]() or "") != imei)
        ):
            raise RuntimeError("串口连接或设备身份已变化，已取消短信发送，请重新提交")
        return serial

    def record_sent():
        if not namespace.get("CLOUD_CONTROL_ENABLED", True):
            return "disabled"
        if not imei:
            namespace["port_ui"]("短信已发送，但设备身份尚未就绪，未上传发送记录", "normal")
            return "missing_imei"
        payload = build_sent_sms_event_payload(
            phone, message, namespace["_cloud_now_ts"](), identity, source_id,
        )
        loop, ws = namespace.get("cloud_ws_loop"), namespace.get("cloud_ws_conn")
        can_send = bool(
            loop is not None and loop.is_running() and ws is not None
            and namespace.get("cloud_connected")
            and namespace.get("cloud_device_authorized")
            and namespace["_cloud_runtime_imei"]() == imei
        )
        return enqueue_cloud_sms_event_runtime(
            payload,
            event_queue=namespace["CLOUD_SMS_EVENT_Q"],
            can_send=can_send,
            loop=loop,
            ws=ws,
            schedule_drain=namespace["_schedule_cloud_sms_event_drain"],
            log_error=namespace.get("log_file_only"),
            state=namespace.get("CLOUD_SMS_EVENT_DRAIN_STATE"),
            is_enabled=lambda: namespace.get("CLOUD_CONTROL_ENABLED", True),
        )

    return send_runtime(
        namespace["serial_lock"], get_captured_serial, phone, message,
        push_debug=namespace["_push_serial_debug"],
        port_ui=namespace["port_ui"],
        log_error=namespace.get("log_file_only"),
        on_sent=record_sent,
    )
