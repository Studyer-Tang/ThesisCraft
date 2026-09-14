"""Output selection preserves source-adjacent defaults and no-clobber behavior."""

from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from docx import Document

from word_formatter.academic.templates import load_template
from word_formatter.academic.workflow import run
from word_formatter.office_toolbar import OfficeToolbar
from word_formatter.ui_common import output_directory


class OutputLocationTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name).resolve()
        self.source = self.root / '论文.docx'
        doc = Document()
        doc.add_paragraph('这是需要保留的原文。')
        doc.save(self.source)
        self.folder = self.root / '导出 文件夹'
        self.folder.mkdir()

    def test_custom_folder_creates_only_new_docx_and_preserves_existing_files(self):
        existing = self.folder / '论文_排版.docx'
        existing.write_bytes(b'previous result')
        before = self.source.read_bytes()
        template = load_template('master')
        template['enabled'] = []
        result = run(self.source, template, output_dir=output_directory(self.folder))
        self.assertEqual(self.folder / '论文_排版_2.docx', Path(result['output']))
        self.assertEqual({existing, Path(result['output'])}, set(self.folder.iterdir()))
        self.assertEqual(before, self.source.read_bytes())
        self.assertEqual(b'previous result', existing.read_bytes())
        self.assertIsNone(result['report'])

    def test_empty_selection_defaults_to_source_but_missing_selection_is_an_error(self):
        self.assertIsNone(output_directory(''))
        self.assertIsNone(output_directory(None))
        for bad in [self.root / 'missing', self.source]:
            with self.assertRaisesRegex(ValueError, '重新选择'):
                output_directory(bad)
        template = load_template('master')
        template['enabled'] = []
        result = run(self.source, template, output_dir=output_directory(''))
        self.assertEqual(self.root, Path(result['output']).parent)

    def test_batch_default_saves_beside_each_source(self):
        from word_formatter.jobs import build_jobs
        other = self.folder / self.source.name
        other.write_bytes(self.source.read_bytes())
        jobs = build_jobs([self.source, other], beside_sources=True)
        self.assertEqual([self.source.parent, other.parent], [job.output.parent for job in jobs])
        self.assertEqual(2, len({job.output for job in jobs}))

    def test_office_quick_format_and_settings_use_selected_folder(self):
        app, notify = Mock(), Mock()
        app.Documents.Count = 1
        app.ActiveDocument = SimpleNamespace(Path=str(self.root), Saved=True, FullName=str(self.source))
        selected = str(self.folder)
        with patch.object(OfficeToolbar, '_attach_buttons'):
            toolbar = OfficeToolbar(app, 'word', notify, output_getter=lambda: selected)
        toolbar.buttons = [SimpleNamespace(Enabled=True, Caption='')]
        with patch('word_formatter.academic.workflow.run', return_value={'output': None}) as workflow:
            toolbar.format_thesis(pdf=True, export_only=True)
            toolbar.worker.join(timeout=5)
            self.assertFalse(toolbar.worker.is_alive())
            self.assertEqual(selected, workflow.call_args.kwargs['output_dir'])
            self.assertTrue(workflow.call_args.kwargs['pdf'])
            self.assertEqual([], workflow.call_args.args[1]['enabled'])
        toolbar.busy = False
        with patch('word_formatter.academic.plugin.launch') as launch:
            toolbar.open_academic()
            self.assertEqual(['--academic', str(self.source), '--host', 'word', '--output-dir', selected], launch.call_args.args[0])
        notify.assert_not_called()

    def test_office_general_format_routes_jobs_to_selected_folder(self):
        app, notify = Mock(), Mock()
        app.Documents.Count = 1
        app.ActiveDocument = SimpleNamespace(Path=str(self.root), Saved=True, FullName=str(self.source))
        with patch.object(OfficeToolbar, '_attach_buttons'):
            toolbar = OfficeToolbar(app, 'word', notify, output_getter=lambda: str(self.folder))
        toolbar.buttons = [SimpleNamespace(Enabled=True, Caption='')]
        with patch('word_formatter.office_toolbar.user_config_path', return_value=self.root/'no-config.json'), patch('word_formatter.office_toolbar.run_jobs') as worker:
            toolbar.format_current()
            toolbar.worker.join(timeout=5)
            self.assertFalse(toolbar.worker.is_alive())
            self.assertEqual(self.folder, worker.call_args.args[1][0].output.parent)
        notify.assert_not_called()
