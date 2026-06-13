import configparser
import json
import queue
import unittest

from sms_core.third_push import (
    PushDispatchResult,
    dispatch_push_item,
    push_result_status_message,
    third_push_save_kwargs,
    third_push_saved_status,
    third_push_state,
)
from sms_core.third_push_runtime import (
    build_third_push_payload,
    enqueue_third_push_runtime,
    resolve_push_channels,
    third_push_worker_runtime,
)
from sms_core.third_push_config import (
    ThirdPushSettings,
    ensure_third_push_config_values,
    read_third_push_settings,
    update_third_push_settings,
    write_third_push_settings,
)
from sms_core.third_push_sender import ALLOWED_PUSH_SCHEMES, http_request


class ThirdPushDispatchTests(unittest.TestCase):
    def test_third_push_window_state_helpers(self):
        state = third_push_state(True, False, True, ["dingtalk"], {"token": "t"})

        self.assertEqual(state, {
            "enabled": True,
            "sms_enabled": False,
            "call_enabled": True,
            "channels": ["dingtalk"],
            "settings": {"token": "t"},
        })

        self.assertEqual(
            third_push_save_kwargs((1, 0, True, ("dingtalk",), {"token": "t"})),
            {
                "enabled": True,
                "sms_enabled": False,
                "call_enabled": True,
                "notify_type": ["dingtalk"],
                "settings": {"token": "t"},
            },
        )

        self.assertEqual(third_push_saved_status(True, ["dingtalk"]), "📡 三方推送：已开启，通道：dingtalk")
        self.assertEqual(third_push_saved_status(False, []), "📡 三方推送：已关闭，通道：未选择")

    def test_dispatch_push_item_collects_successes_and_failures(self):
        sent = []

        def send_channel(channel, message, settings):
            sent.append((channel, message, settings["token"]))
            if channel == "dingtalk":
                return True, "ok"
            return False, "bad config"

        result = dispatch_push_item(
            {
                "message": "raw",
                "channels": ["dingtalk", "wecom"],
                "settings": {"token": "t"},
                "template": "hello {msg}",
            },
            send_channel,
            format_message_func=lambda raw, template: template.replace("{msg}", raw),
        )

        self.assertEqual(result.ok_channels, ["钉钉"])
        self.assertEqual(result.fail_infos, ["企业微信: bad config"])
        self.assertEqual(sent[0], ("dingtalk", "hello raw", "t"))

    def test_dispatch_push_item_converts_sender_exception_to_failure(self):
        result = dispatch_push_item(
            {"message": "raw", "channels": ["telegram"], "settings": {}},
            lambda *_args: (_ for _ in ()).throw(RuntimeError("boom")),
        )

        self.assertEqual(result.ok_channels, [])
        self.assertEqual(result.fail_infos, ["Telegram: boom"])

    def test_push_result_status_message_formats_result_cases(self):
        partial = PushDispatchResult(["钉钉"], ["企业微信: bad"])
        failed = PushDispatchResult([], ["企业微信: bad"])
        success = PushDispatchResult(["钉钉"], [], show_success=True)
        quiet = PushDispatchResult(["钉钉"], [], show_success=False)

        self.assertIn("部分成功", push_result_status_message(partial))
        self.assertIn("推送失败", push_result_status_message(failed))
        self.assertIn("测试成功", push_result_status_message(success))
        self.assertEqual(push_result_status_message(quiet), "")

    def test_resolve_push_channels_checks_feature_flags(self):
        base = {
            "enabled": True,
            "sms_enabled": True,
            "call_enabled": True,
            "configured_channels": ["dingtalk", "bark"],
            "valid_channels": {"dingtalk": "Ding", "bark": "Bark"},
        }

        self.assertEqual(resolve_push_channels(None, **base), ["dingtalk", "bark"])
        self.assertEqual(resolve_push_channels(None, **{**base, "enabled": False}), [])
        self.assertEqual(resolve_push_channels(None, **{**base, "sms_enabled": False}), [])
        self.assertEqual(
            resolve_push_channels(None, **{**base, "call_enabled": False, "event_type": "call"}),
            [],
        )
        self.assertEqual(resolve_push_channels(["bark", "missing"], **base), ["bark"])

    def test_build_third_push_payload_uses_call_template(self):
        payload = build_third_push_payload(
            "hello",
            channels=["dingtalk"],
            settings={"token": "t"},
            event_type="call",
            show_success=True,
            show_result=True,
            sms_template="sms:{msg}",
            call_template="call:{msg}",
        )

        self.assertEqual(payload["message"], "hello")
        self.assertEqual(payload["channels"], ["dingtalk"])
        self.assertEqual(payload["settings"], {"token": "t"})
        self.assertEqual(payload["template"], "call:{msg}")
        self.assertTrue(payload["show_success"])
        self.assertTrue(payload["show_result"])

    def test_enqueue_third_push_runtime_queues_payload(self):
        push_queue = queue.Queue()

        result = enqueue_third_push_runtime(
            "msg",
            push_queue=push_queue,
            enabled=True,
            sms_enabled=True,
            call_enabled=True,
            configured_channels=["dingtalk"],
            current_settings={"token": "current"},
            show_success=True,
            valid_channels={"dingtalk": "Ding"},
        )

        self.assertTrue(result)
        self.assertEqual(push_queue.get_nowait()["settings"], {"token": "current"})

    def test_enqueue_third_push_runtime_handles_disabled_and_full_queue(self):
        messages = []
        push_queue = queue.Queue(maxsize=1)
        push_queue.put_nowait({"old": True})

        self.assertFalse(enqueue_third_push_runtime(
            "msg",
            push_queue=queue.Queue(),
            enabled=False,
            sms_enabled=True,
            call_enabled=True,
            configured_channels=["dingtalk"],
            current_settings={},
            valid_channels={"dingtalk": "Ding"},
        ))

        self.assertFalse(enqueue_third_push_runtime(
            "msg",
            push_queue=push_queue,
            enabled=True,
            sms_enabled=True,
            call_enabled=True,
            configured_channels=["dingtalk"],
            current_settings={},
            system_ui=lambda msg, tag="normal": messages.append((msg, tag)),
            valid_channels={"dingtalk": "Ding"},
        ))
        self.assertEqual(messages[0][1], "normal")

    def test_third_push_worker_runtime_dispatches_and_marks_done(self):
        class StopAfterOne:
            def __init__(self):
                self.calls = 0

            def is_set(self):
                self.calls += 1
                return self.calls > 1

        push_queue = queue.Queue()
        push_queue.put_nowait({"message": "raw", "show_result": True})
        ui_messages = []
        results = []

        third_push_worker_runtime(
            stop_event=StopAfterOne(),
            push_queue=push_queue,
            send_channel_func=lambda *_args: (True, "ok"),
            system_ui=lambda msg, tag="normal": ui_messages.append((msg, tag)),
            show_result=lambda ok, fail: results.append((ok, fail)),
            dispatch_func=lambda item, send, format_message_func=None: PushDispatchResult(
                ["Ding"],
                [],
                show_success=True,
                show_result=item.get("show_result"),
            ),
            poll_timeout=0,
        )

        self.assertTrue(push_queue.empty())
        self.assertIn("Ding", ui_messages[0][0])
        self.assertEqual(results, [(["Ding"], [])])

    def test_third_push_config_read_update_and_write(self):
        config = configparser.ConfigParser()
        config["third_push"] = {
            "enabled": "1",
            "sms_enabled": "0",
            "call_enabled": "1",
            "notify_type": '["dingtalk", "bad", "telegram"]',
            "message_template": "old",
            "dingtalk_webhook": "https://example.test/ding",
        }

        self.assertTrue(ensure_third_push_config_values(config))
        self.assertFalse(config.has_option("third_push", "message_template"))

        settings = read_third_push_settings(config)
        self.assertTrue(settings.enabled)
        self.assertFalse(settings.sms_enabled)
        self.assertTrue(settings.call_enabled)
        self.assertEqual(settings.channels, ["dingtalk", "telegram"])
        self.assertEqual(settings.settings["dingtalk_webhook"], "https://example.test/ding")

        updated = update_third_push_settings(
            settings,
            enabled=False,
            sms_enabled=True,
            call_enabled=False,
            notify_type=["wecom", "missing", "bark"],
            settings={"wecom_webhook": "https://example.test/wecom"},
        )
        self.assertIsInstance(updated, ThirdPushSettings)
        self.assertFalse(updated.enabled)
        self.assertTrue(updated.sms_enabled)
        self.assertFalse(updated.call_enabled)
        self.assertEqual(updated.channels, ["wecom", "bark"])
        self.assertEqual(updated.settings["wecom_webhook"], "https://example.test/wecom")
        self.assertEqual(updated.settings["dingtalk_webhook"], "")

        write_third_push_settings(config, updated)
        self.assertEqual(config.get("third_push", "enabled"), "0")
        self.assertEqual(config.get("third_push", "sms_enabled"), "1")
        self.assertEqual(config.get("third_push", "call_enabled"), "0")
        self.assertEqual(json.loads(config.get("third_push", "notify_type")), ["wecom", "bark"])
        self.assertEqual(config.get("third_push", "wecom_webhook"), "https://example.test/wecom")


class HttpRequestSchemeTests(unittest.TestCase):
    def test_rejects_non_http_schemes_without_opening(self):
        for url in (
            "file:///etc/passwd",
            "ftp://example.test/data",
            "gopher://example.test/",
            "data:text/plain,boom",
            "C:\\Windows\\System32\\drivers\\etc\\hosts",
            "",
        ):
            ok, code, body = http_request(url)
            self.assertFalse(ok, f"scheme should be rejected: {url!r}")
            self.assertIsNone(code)
            self.assertIn("不支持的 URL 协议", body)

    def test_allows_http_and_https_schemes(self):
        opened = []

        class FakeResp:
            def __init__(self, url):
                self._url = url

            def __enter__(self):
                return self

            def __exit__(self, *exc):
                return False

            def read(self, _n):
                return b"ok"

            def getcode(self):
                return 200

        def fake_urlopen(req, timeout=None):
            opened.append(req.full_url)
            return FakeResp(req.full_url)

        import urllib.request as _u

        original = _u.urlopen
        _u.urlopen = fake_urlopen
        try:
            for url in ("http://example.test/hook", "https://example.test/hook"):
                ok, code, body = http_request(url, method="GET")
                self.assertTrue(ok, f"scheme should be allowed: {url!r}")
                self.assertEqual(code, 200)
        finally:
            _u.urlopen = original

        self.assertEqual(opened, ["http://example.test/hook", "https://example.test/hook"])

    def test_allowed_schemes_constant(self):
        self.assertEqual(ALLOWED_PUSH_SCHEMES, ("http", "https"))


if __name__ == "__main__":
    unittest.main()
