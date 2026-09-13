import importlib.util
import os
import sys
import unittest


TOOL_PATH = os.path.join(
    os.path.dirname(os.path.dirname(__file__)),
    "tools",
    "check_call_forwarding.py",
)
SPEC = importlib.util.spec_from_file_location("check_call_forwarding", TOOL_PATH)
MODULE = importlib.util.module_from_spec(SPEC)
sys.modules[SPEC.name] = MODULE
SPEC.loader.exec_module(MODULE)


class CheckCallForwardingToolTests(unittest.TestCase):
    def test_probe_commands_are_read_only_queries(self):
        self.assertEqual(
            [command for _label, command in MODULE.QUERY_COMMANDS],
            [
                "AT+CCFC=?",
                "AT+CCFC=0,2",
                "AT+CCFC=1,2",
                "AT+CCFC=2,2",
                "AT+CCFC=3,2",
            ],
        )

    def test_classifies_full_query_support(self):
        results = [
            MODULE.CommandResult(command, ("OK",), "OK")
            for _label, command in MODULE.QUERY_COMMANDS
        ]
        code, message = MODULE.classify_results(results)
        self.assertEqual(code, 0)
        self.assertIn("支持", message)

    def test_classifies_explicit_unsupported_response(self):
        results = [
            MODULE.CommandResult(command, ("ERROR",), "ERROR")
            for _label, command in MODULE.QUERY_COMMANDS
        ]
        code, message = MODULE.classify_results(results)
        self.assertEqual(code, 2)
        self.assertIn("不支持", message)

    def test_specific_queries_override_unsupported_test_command(self):
        results = [MODULE.CommandResult("AT+CCFC=?", ("ERROR",), "ERROR")]
        results.extend(
            MODULE.CommandResult(command, ("OK",), "OK")
            for _label, command in MODULE.QUERY_COMMANDS[1:]
        )
        code, message = MODULE.classify_results(results)
        self.assertEqual(code, 0)
        self.assertIn("支持", message)


if __name__ == "__main__":
    unittest.main()
