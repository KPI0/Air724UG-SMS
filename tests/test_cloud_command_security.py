import configparser
import unittest

from sms_core.cloud_command_security import (
    CLOUD_COMMAND_PERMISSION_SPECS,
    CLOUD_SEND_SMS_TRANSACTION_COMMAND,
    CLOUD_SET_OWN_NUMBER_TRANSACTION_COMMAND,
    cloud_command_batch_error,
    cloud_command_control_char_error,
    cloud_command_has_line_break,
    cloud_command_has_chained_separator,
    cloud_sensitive_command_block_message,
    is_sensitive_cloud_command_allowed,
    read_cloud_command_permissions,
    read_cloud_sensitive_commands_enabled,
    sensitive_cloud_command_decision,
    sensitive_cloud_command_reason,
)


class CloudCommandSecurityTests(unittest.TestCase):
    def test_cloud_command_line_break_detection_uses_original_text(self):
        self.assertTrue(cloud_command_has_line_break("AT+CSQ\r\nAT+RESET"))
        self.assertTrue(cloud_command_has_line_break("AT+CSQ\nAT+CMGD=1"))
        self.assertTrue(cloud_command_has_line_break("AT+CSQ\rAT+RESET"))
        self.assertFalse(cloud_command_has_line_break("  AT+CSQ  "))

    def test_cloud_command_chained_separator_allows_only_terminal_separators(self):
        self.assertTrue(cloud_command_has_chained_separator("AT+CSQ;AT+RESET"))
        self.assertTrue(cloud_command_has_chained_separator("001122\x1aAT+RESET"))
        self.assertFalse(cloud_command_has_chained_separator("ATD10086;"))
        self.assertFalse(cloud_command_has_chained_separator("001122\x1a"))
        self.assertFalse(cloud_command_has_chained_separator('AT+CUSD=1,"*100;#"'))
        self.assertIn("每次只发送一条", cloud_command_batch_error("AT+CSQ;AT+RESET"))

    def test_cloud_command_control_char_detection_matches_server_policy(self):
        rejected_codes = (
            *range(0x00, 0x09),
            0x0B,
            0x0C,
            *range(0x0E, 0x1A),
            *range(0x1B, 0x20),
            0x7F,
        )
        for code in rejected_codes:
            command = "AT+CSQ" + chr(code) + "AT+RESET"
            with self.subTest(code=hex(code)):
                self.assertIn("控制字符", cloud_command_control_char_error(command))

        self.assertEqual(cloud_command_control_char_error("AT\t"), "")
        self.assertEqual(cloud_command_control_char_error("001122\x1a"), "")

    def test_config_defaults_to_disabled_and_reads_enabled(self):
        config = configparser.ConfigParser()
        self.assertFalse(read_cloud_sensitive_commands_enabled(config))

        config["cloud_control"] = {"allow_sensitive_commands": "1"}
        self.assertTrue(read_cloud_sensitive_commands_enabled(config))
        self.assertTrue(all(read_cloud_command_permissions(config).values()))

    def test_reads_independent_permission_switches(self):
        config = configparser.ConfigParser()
        config["cloud_control"] = {
            "allow_sensitive_commands": "0",
            "allow_sensitive_sms": "1",
            "allow_sensitive_call": "0",
        }

        permissions = read_cloud_command_permissions(config)

        self.assertTrue(permissions["sms"])
        self.assertFalse(permissions["call"])
        self.assertEqual(len(permissions), len(CLOUD_COMMAND_PERMISSION_SPECS))

    def test_migrates_previous_pin_and_other_permissions(self):
        config = configparser.ConfigParser()
        config["cloud_control"] = {
            "allow_sensitive_commands": "0",
            "allow_sensitive_sim_security": "1",
            "allow_sensitive_other": "1",
        }

        permissions = read_cloud_command_permissions(config)

        self.assertTrue(permissions["pin"])
        self.assertTrue(permissions["puk"])
        self.assertNotIn("pin_lock", permissions)
        self.assertTrue(permissions["phone_number"])
        self.assertTrue(permissions["sn"])

    def test_migrates_previous_pin_lock_permission_into_pin_operations(self):
        config = configparser.ConfigParser()
        config["cloud_control"] = {
            "allow_sensitive_pin": "0",
            "allow_sensitive_pin_lock": "1",
        }

        permissions = read_cloud_command_permissions(config)

        self.assertTrue(permissions["pin"])
        self.assertNotIn("pin_lock", permissions)

    def test_allows_known_safe_read_only_commands(self):
        for command in (
            "AT",
            "ATI",
            "AT+CSQ",
            "AT+CESQ",
            "AT+CFUN?",
            "AT+COPS=?",
            "AT+CUSD?",
            "AT+CSCA?",
        ):
            with self.subTest(command=command):
                self.assertEqual(sensitive_cloud_command_reason(command), "")

    def test_blocks_required_sensitive_command_groups(self):
        cases = {
            CLOUD_SEND_SMS_TRANSACTION_COMMAND: "短信",
            CLOUD_SET_OWN_NUMBER_TRANSACTION_COMMAND: "本机号码",
            'AT+CMGS=23': "短信",
            'AT+CMGC=23': "短信",
            'ATD10086;': "电话",
            'AT+CPIN="1234"': "PIN",
            'AT+CPIN="12345678","1234"': "PUK",
            'AT+CLCK="SC",1,"1234"': "PIN 码操作",
            'AT+CLCK="SC",0,"1234"': "PIN 码操作",
            'AT+CPBS="ON"': "本机号码",
            'AT+CPBW=1,"+8610010",145': "本机号码",
            'AT+WISN=ABC123': "SN",
            'AT+EEMGINFO?': "基站定位",
            'AT+CUSD=1,"*100#"': "USSD",
            'ATD*100#;': "USSD",
            'AT+CCFC=0,3': "呼叫转移",
            'AT+CLCK="AO",1,"1234"': "呼叫限制",
            'AT+CPWD="AO","1234","5678"': "呼叫限制",
            'ATD**21*13800138000#;': "呼叫转移",
            'ATD#33*1234#;': "呼叫限制",
            'AT+CSCA="+8613800100500"': "信息中心",
            'AT+CMGD=1': "删除设备数据",
            'AT+CPBW=1': "删除设备数据",
            'AT+REBOOT': "重置或关闭设备",
            'AT+CFUN=0': "重置或关闭设备",
            'AT+CPOWD=1': "重置或关闭设备",
        }
        for command, expected in cases.items():
            with self.subTest(command=command):
                self.assertIn(expected, sensitive_cloud_command_reason(command))

    def test_pin_permission_uses_operation_label(self):
        pin_spec = next(
            spec for spec in CLOUD_COMMAND_PERMISSION_SPECS
            if spec.category == "pin"
        )

        self.assertEqual(pin_spec.label, "PIN 码操作")
        self.assertEqual(
            sensitive_cloud_command_reason('AT+CPIN="1234"'),
            "PIN 码操作",
        )
        self.assertEqual(
            sensitive_cloud_command_decision('AT+CLCK="SC",1,"1234"').category,
            "pin",
        )
        self.assertEqual(
            sensitive_cloud_command_decision(
                'AT+CPWD="SC","1234","5678"'
            ).category,
            "pin",
        )
        self.assertEqual(
            sensitive_cloud_command_decision(
                'AT+CPWD="AO","1234","5678"'
            ).category,
            "call_control",
        )

    def test_puk_and_phone_number_permissions_use_current_labels(self):
        specs = {
            spec.category: spec
            for spec in CLOUD_COMMAND_PERMISSION_SPECS
        }

        self.assertEqual(specs["puk"].label, "PUK 码操作")
        self.assertEqual(specs["phone_number"].label, "修改本机号码")
        self.assertEqual(
            sensitive_cloud_command_reason('AT+CPIN="12345678","1234"'),
            "PUK 码操作",
        )
        self.assertEqual(
            sensitive_cloud_command_reason('AT+CPBW=1,"+8610010",145'),
            "修改本机号码",
        )

    def test_information_center_permission_uses_current_label(self):
        spec = next(
            item for item in CLOUD_COMMAND_PERMISSION_SPECS
            if item.category == "sms_center"
        )

        self.assertEqual(spec.label, "修改信息中心号码")
        self.assertEqual(
            sensitive_cloud_command_reason('AT+CSCA="+8613800100500"'),
            "修改信息中心号码",
        )

    def test_other_commands_are_not_controlled_by_the_simplified_setting(self):
        for command in (
            "AT+CMGR=1",
            "AT+CGSN",
            "0011000D916831",
            "AT+RESET",
            "AT+UNKNOWN",
        ):
            with self.subTest(command=command):
                self.assertEqual(sensitive_cloud_command_reason(command), "")

    def test_new_permissions_use_independent_categories(self):
        cases = {
            'AT+CUSD=1,"*100#"': "ussd",
            'AT+CCFC=0,3': "call_control",
            'AT+CSCA="+8613800100500"': "sms_center",
            'AT+CMGD=1': "delete_data",
            'AT+FSFORMAT=1': "delete_data",
            'AT+REBOOT': "device_power",
        }

        for command, category in cases.items():
            with self.subTest(command=command):
                decision = sensitive_cloud_command_decision(command)
                self.assertEqual(decision.category, category)
                self.assertFalse(
                    is_sensitive_cloud_command_allowed(decision, {category: False})
                )
                self.assertTrue(
                    is_sensitive_cloud_command_allowed(decision, {category: True})
                )

    def test_sms_center_and_phonebook_delete_do_not_use_older_permissions(self):
        sms_center = sensitive_cloud_command_decision(
            'AT+CSCA="+8613800100500"'
        )
        phonebook_delete = sensitive_cloud_command_decision("AT+CPBW=1")
        phone_number_write = sensitive_cloud_command_decision(
            'AT+CPBW=1,"+8610010",145'
        )

        self.assertEqual(sms_center.category, "sms_center")
        self.assertEqual(phonebook_delete.category, "delete_data")
        self.assertEqual(phone_number_write.category, "phone_number")

    def test_mmi_security_codes_cannot_bypass_pin_puk_or_call_control(self):
        cases = {
            'ATD**04*1234*5678*5678#;': "pin",
            'ATD**05*12345678*5678*5678#;': "puk",
            'ATD**21*13800138000#;': "call_control",
            'ATD##21#;': "call_control",
            'ATD*33*1234#;': "call_control",
            'ATD*100#;': "ussd",
        }

        for command, category in cases.items():
            with self.subTest(command=command):
                self.assertEqual(
                    sensitive_cloud_command_decision(command).category,
                    category,
                )

    def test_send_sms_metadata_is_sensitive_even_when_command_looks_generic(self):
        reason = sensitive_cloud_command_reason(
            "AT+CMGF=0",
            {"command_kind": "send_sms", "sms_log": "summary"},
        )
        self.assertEqual(reason, "发送短信")

    def test_untrusted_metadata_cannot_override_real_sensitive_category(self):
        cases = (
            (
                "AT+REBOOT",
                {"command_kind": "send_sms", "sms_log": "suppress"},
                "device_power",
            ),
            (
                'AT+CPIN="1234"',
                {"command_kind": "send_sms", "sms_log": "summary"},
                "pin",
            ),
        )

        for command, metadata, expected_category in cases:
            with self.subTest(command=command):
                self.assertEqual(
                    sensitive_cloud_command_decision(command, metadata).category,
                    expected_category,
                )

    def test_permission_check_uses_only_matching_category(self):
        sms = sensitive_cloud_command_decision("AT+CMGS=23")
        pin = sensitive_cloud_command_decision('AT+CPIN="1234"')
        permissions = {"sms": True, "pin": False}

        self.assertTrue(is_sensitive_cloud_command_allowed(sms, permissions))
        self.assertFalse(is_sensitive_cloud_command_allowed(pin, permissions))

    def test_block_message_points_to_client_setting(self):
        message = cloud_sensitive_command_block_message("拨打电话")
        self.assertIn("安全设置", message)
        self.assertIn("拨打电话", message)


if __name__ == "__main__":
    unittest.main()
