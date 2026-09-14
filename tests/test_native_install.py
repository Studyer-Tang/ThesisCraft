import struct
import unittest
from unittest.mock import patch

from word_formatter.entry import main
from word_formatter.native_install import validate_binary


class NativeInstallTests(unittest.TestCase):
    def test_rejects_truncated_and_wrong_architecture_dlls(self):
        data = bytearray(128)
        data[:2] = b'MZ'
        struct.pack_into('<I', data, 60, 80)
        data[80:84] = b'PE\0\0'
        struct.pack_into('<H', data, 84, 0x8664)
        validate_binary(data, 0x8664)
        for bad, arch in [(data[:64], 0x8664), (data, 0x14c), (b'not a DLL', 0x8664)]:
            with self.assertRaises(ValueError):
                validate_binary(bad, arch)

    def test_install_and_uninstall_routes_do_not_start_formatting(self):
        for option, remove in [('--install-native', False), ('--uninstall-native', True)]:
            with patch('word_formatter.native_install.main', return_value=0) as install:
                self.assertEqual(0, main([option]))
                install.assert_called_once_with(remove=remove)
