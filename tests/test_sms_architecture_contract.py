import unittest
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def _read(relative_path):
    return (ROOT / relative_path).read_text(encoding="utf-8")


class SmsArchitectureContractTests(unittest.TestCase):
    def test_pipeline_stays_blind_forwarder(self):
        source = _read("sms/sms_core/sms_receive_pipeline.py")

        forbidden = [
            "correct_callback_text",
            "concat_part_for_callback",
            "segments_for_concat",
            "PendingSms",
            "full_msg",
        ]
        for token in forbidden:
            with self.subTest(token=token):
                self.assertNotIn(token, source)

    def test_pdu_cache_does_not_reintroduce_multipart_assembler(self):
        source = _read("sms/sms_core/serial_sms_pdu_cache.py")

        forbidden = [
            "_multipart",
            "complete_metadata_for_concat_part",
            "_lookup_assembled",
            "body_parts",
            "_multipart_timestamps",
            "CachedSmsPduMessage",
            "message_trace_id",
            'body = "".join',
            "parts.setdefault",
        ]
        for token in forbidden:
            with self.subTest(token=token):
                self.assertNotIn(token, source)

    def test_collected_event_adapter_does_not_decide_multipart_completion(self):
        source = _read("sms/sms_core/sms_collected_event.py")

        forbidden = [
            "_completed",
            "is_complete",
            "complete_body",
            'body = "".join',
            "parts_seen",
        ]
        for token in forbidden:
            with self.subTest(token=token):
                self.assertNotIn(token, source)
        self.assertIsNone(re.search(r"self\._pending|_pending\s*=", source))


if __name__ == "__main__":
    unittest.main()
