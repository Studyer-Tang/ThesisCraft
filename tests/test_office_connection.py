"""Regression coverage for secondary Office instances and disconnected cleanup."""
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest
from unittest.mock import Mock, patch

from word_formatter.office_connection import choose_application, is_busy_error
from word_formatter.office_toolbar import OfficeToolbar


class Documents(list):
    @property
    def Count(self):
        return len(self)


def application(hwnd, path='example.docx', name='Microsoft Word', visible=True):
    return SimpleNamespace(
        Name=name, Visible=visible,
        Documents=Documents([SimpleNamespace(FullName=str(Path(path).resolve()))]),
        Windows=[SimpleNamespace(Hwnd=hwnd)],
    )


class OfficeConnectionTests(unittest.TestCase):
    def test_front_document_wins_over_global_active_instance(self):
        global_app, front_app = application(10), application(20)
        self.assertIs(choose_application([global_app, front_app], 'word', [99, 20, 10]), front_app)

    def test_path_limits_candidates(self):
        front_app, requested_app = application(10, 'front.docx'), application(20, 'target.docx')
        self.assertIs(choose_application([front_app, requested_app], 'word', [10, 20],
                                         'target.docx'), requested_app)

    def test_skips_hidden_empty_wrong_host_stale_and_invisible_windows(self):
        empty = application(2)
        empty.Documents.clear()
        stale = Mock()
        type(stale).Name = property(lambda _: (_ for _ in ()).throw(RuntimeError('disconnected')))
        valid = application(6)
        candidates = [application(1, visible=False), empty, application(3, name='WPS Writer'),
                      stale, application(5), valid]
        self.assertIs(choose_application(candidates, 'word', [1, 2, 3, 4, 6]), valid)

    def test_wps_is_selected_independently(self):
        word, wps = application(10), application(20, name='WPS 文字')
        self.assertIs(choose_application([word, wps], 'wps', [10, 20]), wps)

    def test_signed_window_handles_are_normalized(self):
        app = application(-1)
        self.assertIs(choose_application([app], 'word', [0xffffffff]), app)

    def test_wps_word_compatibility_name_and_child_window(self):
        with tempfile.TemporaryDirectory() as folder:
            (Path(folder) / 'wps.exe').touch()
            wps = application(123, name='Microsoft Word')
            wps.Path = folder
            self.assertIs(choose_application([wps], 'wps', [999],
                                             window_root=lambda hwnd: 999), wps)
            with self.assertRaises(RuntimeError):
                choose_application([wps], 'word', [999], window_root=lambda hwnd: 999)

    def test_no_document_has_actionable_message(self):
        with self.assertRaisesRegex(RuntimeError, '连接当前文档'):
            choose_application([], 'word', [])

    def test_busy_errors_are_distinct_from_rpc_disconnect(self):
        self.assertTrue(is_busy_error(Exception(-2147418111, 'rejected')))
        self.assertTrue(is_busy_error(Exception(-2147417846, 'retry')))
        self.assertFalse(is_busy_error(Exception(-2147023174, 'RPC unavailable')))
        self.assertFalse(is_busy_error(Exception()))


class OfficeCleanupTests(unittest.TestCase):
    def test_academic_result_opens_document_without_report_browser(self):
        app, notify = Mock(), Mock()
        with patch.object(OfficeToolbar, '_attach_buttons'):
            toolbar = OfficeToolbar(app, 'word', notify)
        toolbar.buttons = [SimpleNamespace(Enabled=False, Caption='busy')]
        payload = {'output': str(Path('result.docx').resolve()), 'report': None, 'warnings': []}
        toolbar.events.put(('academic', payload))
        with patch('webbrowser.open') as browser:
            toolbar.poll()
        app.Documents.Open.assert_called_once_with(payload['output'])
        browser.assert_not_called()
        notify.assert_not_called()

    def test_unadvise_failure_does_not_abort_remaining_cleanup(self):
        with patch.object(OfficeToolbar, '_attach_buttons'):
            toolbar = OfficeToolbar(None, 'word')
        failed, remaining, bar = Mock(), Mock(), Mock()
        failed.close.side_effect = RuntimeError('RPC server unavailable')
        bar.Delete.side_effect = RuntimeError('Office already exited')
        toolbar.handlers = [failed, remaining]
        toolbar.bar = bar
        toolbar.buttons = [Mock()]
        toolbar.close()
        toolbar.close()
        failed.close.assert_called_once()
        remaining.close.assert_called_once()
        bar.Delete.assert_called_once()
        self.assertEqual(toolbar.handlers, [])
        self.assertEqual(toolbar.buttons, [])
        self.assertIsNone(toolbar.bar)

    def test_partial_attach_failure_preserves_original_error(self):
        bar = Mock()
        bar.Delete.side_effect = RuntimeError('cleanup failed')
        def fail_attach(toolbar):
            toolbar.bar = bar
            raise ValueError('attach failed')
        with patch.object(OfficeToolbar, '_attach_buttons', fail_attach):
            with self.assertRaisesRegex(ValueError, 'attach failed'):
                OfficeToolbar(None, 'word')
        bar.Delete.assert_called_once()

    def test_panel_actions_survive_native_toolbar_failure(self):
        with patch.object(OfficeToolbar, '_attach_buttons', side_effect=RuntimeError('unsupported')):
            with self.assertLogs('word_formatter.office_toolbar', level='ERROR'):
                toolbar = OfficeToolbar(None, 'wps', allow_panel_fallback=True)
        self.assertFalse(toolbar.closed)
        self.assertEqual(len(toolbar.buttons), len(toolbar.actions()))
        self.assertTrue(all(button.Enabled for button in toolbar.buttons))
        toolbar.close()
