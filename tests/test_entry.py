"""Installed launches must work without the repository's compatibility scripts."""
import subprocess
import sys
import tempfile
import unittest

from word_formatter.academic.plugin import command


class EntryTests(unittest.TestCase):
    def test_plugin_uses_installed_module(self):
        args = command(['--version'])
        self.assertEqual(args[1:], ['-m', 'word_formatter', '--version'])

    def test_module_version_outside_checkout(self):
        with tempfile.TemporaryDirectory() as cwd:
            result = subprocess.run(
                [sys.executable, '-m', 'word_formatter', '--version'],
                cwd=cwd, capture_output=True, text=True, check=True,
            )
        self.assertEqual(result.stdout.strip(), '4.0.0')
