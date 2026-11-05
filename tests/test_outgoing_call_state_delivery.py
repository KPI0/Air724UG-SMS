from dataclasses import replace
import asyncio
import unittest

from sms_core.call_command_context import SerialCallStateSnapshot, set_local_dial_number
from sms_core.call_session import IncomingCallSessionTracker
from sms_core.call_events import CallLineDecision, CallState
from sms_core.cloud_message_namespace_runtime import send_cloud_call_state_namespace_runtime
from sms_core.serial_runtime import SerialLineDecoder, SerialRuntimeState, handle_serial_runtime_line
from sms_core.cloud_protocol import parse_sms_callback_head
from tests.test_serial_runtime import runtime_callbacks, runtime_config


class OutgoingCallStateDeliveryTests(unittest.TestCase):
    def test_legacy_positional_decision_fields_keep_their_original_order(self):
        decision = CallLineDecision(CallState(), "", "", "", "", "", False, False,
            "outgoing-number", "incoming-number", False, "end-message", "incoming", "end-number",
            "ended", "reason", "original-session", False)
        self.assertEqual(decision.connected_number, "outgoing-number")
        self.assertEqual(decision.incoming_connected_number, "incoming-number")
        self.assertEqual(decision.call_session_id, "original-session")
        self.assertFalse(decision.stop_processing)
        self.assertEqual(decision.outgoing_call_phase, "")

    def run_lines(self, lines):
        state = SerialRuntimeState.create(parse_sms_callback_head)
        state.call_state.current_dial_num = "00000000000"
        events = []
        callbacks = replace(runtime_callbacks([]),
            send_cloud_call_state=lambda *args, **kwargs: events.append((args, kwargs)))
        decoder = SerialLineDecoder()
        for raw in lines:
            for line in decoder.feed((raw + "\r\n").encode()):
                handle_serial_runtime_line(state, line, 100, "synthetic-port", False,
                    runtime_config(), callbacks, {})
        return events

    def test_outgoing_connect_is_structured_once_with_original_session_id(self):
        events = self.run_lines(["CONNECT"])
        self.assertEqual(len(events), 1)
        self.assertEqual(events[0][0], ("00000000000", "connected", ""))
        self.assertEqual(events[0][1]["direction"], "outgoing")
        self.assertTrue(events[0][1]["call_session_id"].startswith("outgoing:"))

    def test_every_current_clcc_presence_survives_repeat_probe(self):
        events = self.run_lines([
            '+CLCC: 1,0,2,0,0,"00000000000",129',
            '+CLCC: 1,0,3,0,0,"00000000000",129',
            '+CLCC: 1,0,3,0,0,"00000000000",129',
            '+CLCC: 1,0,0,0,0,"00000000000",129',
        ])
        self.assertEqual([event[0][1] for event in events], ["dialing", "alerting", "alerting", "connected"])
        self.assertEqual(len({event[1]["call_session_id"] for event in events}), 1)

    def test_other_number_or_incoming_clcc_cannot_supply_outgoing_presence(self):
        self.assertEqual(self.run_lines([
            '+CLCC: 1,0,3,0,0,"11111111111",129',
            '+CLCC: 1,1,3,0,0,"00000000000",129',
        ]), [])

    def test_outgoing_snapshot_rejects_terminal_redial_and_serial_replacement(self):
        for change in ("terminal", "redial", "incoming", "serial", "generation"):
            with self.subTest(change=change):
                namespace = {"ring_timeout_target": 0.0, "current_dial_num": "",
                    "INCOMING_CALL_SESSION": IncomingCallSessionTracker(),
                    "serial_obj": object(), "serial_connection_generation": 1}
                set_local_dial_number(namespace, "00000000000")
                namespace["_local_dial_request"].submitted = True
                snapshot = SerialCallStateSnapshot(namespace)
                if change == "terminal":
                    namespace["current_dial_num"] = ""
                elif change == "redial":
                    set_local_dial_number(namespace, "00000000000")
                elif change == "incoming":
                    namespace["INCOMING_CALL_SESSION"].start("00000000000")
                elif change == "serial":
                    namespace["serial_obj"] = object()
                else:
                    namespace["serial_connection_generation"] += 1
                self.assertFalse(snapshot.is_outgoing_state_current())

    def test_scheduled_presence_rechecks_snapshot_and_terminal_retains_original_id(self):
        for change in ("terminal", "redial", "incoming", "serial", "generation"):
            with self.subTest(change=change):
                wires, pending = [], []

                async def send(_ws, payload):
                    wires.append(payload)
                    return "sent"

                namespace = {"ring_timeout_target": 0.0, "current_dial_num": "",
                    "INCOMING_CALL_SESSION": IncomingCallSessionTracker(),
                    "serial_obj": object(), "serial_connection_generation": 1,
                    "cloud_device_authorized": True, "cloud_ws_loop": None, "cloud_ws_conn": object(),
                    "cloud_connected": True, "_cloud_runtime_imei": lambda: "synthetic-device",
                    "_cloud_build_call_state_payload": lambda *_args, **_kwargs: None,
                    "_cloud_send_payload": send, "_cloud_now_ts": lambda: 100,
                    "_cloud_identity_payload": lambda: {"imei": "synthetic-device"}}
                set_local_dial_number(namespace, "00000000000")
                namespace["_local_dial_request"].submitted = True
                snapshot = SerialCallStateSnapshot(namespace)

                def schedule(phone, phase, reason, **kwargs):
                    pending.append(kwargs["send_payload"](namespace["cloud_ws_conn"],
                        {"phone": phone, "phase": phase, "call_session_id": kwargs["call_session_id"]}))
                    return "scheduled"

                self.assertEqual(send_cloud_call_state_namespace_runtime(namespace, "00000000000", "alerting",
                    direction="outgoing", call_session_id="outgoing-original", call_state_snapshot=snapshot,
                    send_runtime=schedule), "scheduled")
                if change == "terminal":
                    namespace["current_dial_num"] = ""
                elif change == "redial":
                    set_local_dial_number(namespace, "00000000000")
                elif change == "incoming":
                    namespace["INCOMING_CALL_SESSION"].start("00000000000")
                elif change == "serial":
                    namespace["serial_obj"] = object()
                else:
                    namespace["serial_connection_generation"] += 1
                self.assertEqual(asyncio.run(pending.pop()), "stale_call")
                self.assertEqual(wires, [])
                self.assertEqual(send_cloud_call_state_namespace_runtime(namespace, "00000000000", "ended",
                    direction="outgoing", call_session_id="outgoing-original", call_state_snapshot=snapshot,
                    send_runtime=schedule), "scheduled")
                self.assertEqual(asyncio.run(pending.pop()), "sent")
                self.assertEqual(wires[0]["call_session_id"], "outgoing-original")
