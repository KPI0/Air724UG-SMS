import asyncio
import unittest

from sms_core.namespace_binding import bind_async_namespace_runtime, bind_namespace_runtime


class NamespaceBindingTests(unittest.TestCase):
    def test_sync_binding_uses_latest_module_global(self):
        calls = []

        def runtime(namespace, value, *, flag=False):
            return ("old", namespace, value, flag)

        module_globals = {"runtime": runtime}
        bound = bind_namespace_runtime({"name": "ns"}, module_globals, "runtime")

        def patched_runtime(namespace, value, *, flag=False):
            calls.append((namespace, value, flag))
            return "patched"

        module_globals["runtime"] = patched_runtime

        self.assertEqual(bound("value", flag=True), "patched")
        self.assertEqual(calls, [({"name": "ns"}, "value", True)])

    def test_sync_binding_maps_legacy_positional_arguments_to_keywords(self):
        def runtime(namespace, required, *, delay=0, enabled=True):
            return namespace, required, delay, enabled

        bound = bind_namespace_runtime(
            "ns",
            {"runtime": runtime},
            "runtime",
            positional_keywords=("delay", "enabled"),
            positional_prefix_count=1,
        )

        self.assertEqual(bound("work", 3, False), ("ns", "work", 3, False))
        self.assertEqual(bound("work", enabled=False), ("ns", "work", 0, False))

        with self.assertRaises(TypeError):
            bound("work", 3, delay=4)

    def test_async_binding_uses_latest_module_global(self):
        async def runtime(namespace, value, *, timeout=1):
            return ("old", namespace, value, timeout)

        module_globals = {"runtime": runtime}
        bound = bind_async_namespace_runtime(
            "ns",
            module_globals,
            "runtime",
            positional_keywords=("timeout",),
            positional_prefix_count=1,
        )

        async def patched_runtime(namespace, value, *, timeout=1):
            return ("patched", namespace, value, timeout)

        module_globals["runtime"] = patched_runtime

        self.assertEqual(asyncio.run(bound("value", 5)), ("patched", "ns", "value", 5))


if __name__ == "__main__":
    unittest.main()
