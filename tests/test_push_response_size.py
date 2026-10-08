import io
import json
import unittest
from unittest.mock import patch

from sms_core import third_push_sender as sender


class Response(io.BytesIO):
    def __init__(self, data):
        super().__init__(data)
        self.read_sizes = []

    def getcode(self):
        return 200

    def read(self, size=-1):
        self.read_sizes.append(size)
        return super().read(size)


class PushResponseSizeTests(unittest.TestCase):
    def test_wxpusher_success_for_one_hundred_recipients_is_not_truncated(self):
        rows = [{"uid": f"UID_test_{index:03d}", "code": 1000, "messageId": index,
                 "status": "ok", "topicId": None} for index in range(100)]
        body = json.dumps({"code": 1000, "msg": "success", "data": rows}).encode()
        self.assertGreater(len(body), 4096)
        response = Response(body)
        with patch.object(sender.urllib.request, "urlopen", return_value=response):
            ok, info = sender.send_wxpusher("synthetic", {
                "wxpusher_app_token": "AT_test", "wxpusher_uids": ",".join(row["uid"] for row in rows),
            })
        self.assertTrue(ok, info)
        self.assertTrue(response.closed)

    def test_response_limit_is_bounded_and_oversize_is_explicit(self):
        for extra in (0, 1, 1024):
            response = Response(b" " * (sender.MAX_PUSH_RESPONSE_BYTES + extra))
            with self.subTest(extra=extra), patch.object(sender.urllib.request, "urlopen", return_value=response):
                result = sender.http_request("https://example.test/")
            self.assertEqual(result[0], extra == 0)
            self.assertTrue(response.closed)
            self.assertEqual(response.read_sizes, [sender.MAX_PUSH_RESPONSE_BYTES + 1])
            if extra:
                ok, info = sender.api_ok("wxpusher", *result)
                self.assertFalse(ok)
                self.assertIn("响应超过", info)

    def test_malformed_or_business_failure_is_still_a_failure(self):
        for body in (b'{"code":1000,', b'{"code":1001}', b''):
            with self.subTest(body=body), patch.object(sender.urllib.request, "urlopen", return_value=Response(body)):
                ok, _info = sender.send_wxpusher("synthetic", {
                    "wxpusher_app_token": "AT_test", "wxpusher_uids": "UID_test",
                })
            self.assertFalse(ok)


if __name__ == "__main__":
    unittest.main()
