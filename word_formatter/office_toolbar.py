"""Live COM toolbar for Windows Word/WPS, without changing Office trust settings."""

import argparse
import json
import os
from pathlib import Path
import queue
import subprocess
import sys
import threading
import time

from .config import normalize_config
from .engine import WordProcessor
from .jobs import build_jobs
from .service import run_jobs
from .storage import user_config_path
from .version import __version__


class OfficeToolbar:
    def __init__(self, application, host, notify=None):
        import win32com.client

        self.application = application
        self.host = host
        self.events = queue.Queue()
        self.busy = False
        self.worker = None
        self.closed = False
        self.last_result = None
        self.notify = notify or self._notify
        self.bar = application.CommandBars.Add(
            Name=f"学研排版 {os.getpid()}", Position=1, Temporary=True
        )
        self.buttons = []
        self.handlers = []
        for caption, action in (
            ("排版为新副本", self.format_current),
            ("快速论文排版", self.format_thesis),
            ("论文设置 / 检查", self.open_academic),
            ("插入编号 / 引用", self.insert_academic),
            ("更新目录与引用", self.update_academic),
            ("导出论文 PDF", self.export_academic),
            ("排版设置 / 桌面版", self.open_settings),
            ("关于 Study-Tang", self.about),
            ("关闭排版插件", self.request_close),
        ):
            button = self.bar.Controls.Add(Type=1, Temporary=True)
            button = win32com.client.CastTo(button, "_CommandBarButton")
            button.Caption = caption
            button.Style = 2  # msoButtonCaption; independent of bitmap resources.
            button.TooltipText = caption

            class ClickHandler:
                def OnClick(self, control, cancel_default):
                    self.callback()
                    return True

            event = win32com.client.WithEvents(button, ClickHandler)
            event.callback = action
            self.buttons.append(button)
            self.handlers.append(event)
        self.bar.Visible = True

    @staticmethod
    def _notify(message):
        import win32api

        win32api.MessageBox(0, message, "学研排版", 0x40)

    def format_current(self):
        if self.busy:
            return
        try:
            if self.application.Documents.Count == 0:
                raise ValueError("请先打开需要排版的文档。")
            doc = self.application.ActiveDocument
            if not doc.Path or not doc.Saved:
                raise ValueError("请先按 Ctrl+S 保存当前修改，再排版为新副本。")
            source = Path(doc.FullName)
            if not source.is_file():
                raise ValueError("请先将在线文档另存到本地。")
            config_file = user_config_path()
            config = normalize_config(
                json.loads(config_file.read_text(encoding="utf-8"))
                if config_file.exists()
                else {}
            )
            config["preprocess_office"] = False
            jobs = build_jobs([source])
        except Exception as exc:
            self.notify(str(exc))
            return
        self.busy = True
        self.buttons[0].Enabled = False
        self.buttons[0].Caption = "正在排版，请稍候…"

        def work():
            from .conversion import (
                _initialize_com_for_thread,
                _uninitialize_com_for_thread,
            )

            initialized = _initialize_com_for_thread()
            try:
                result = run_jobs(WordProcessor(config), jobs)
                self.events.put(("result", result))
            except Exception as exc:
                self.events.put(("error", str(exc)))
            finally:
                _uninitialize_com_for_thread(initialized)

        self.worker = threading.Thread(target=work, daemon=False)
        self.worker.start()

    def poll(self):
        try:
            kind, payload = self.events.get_nowait()
        except queue.Empty:
            return
        self.busy = False
        try:
            for button in self.buttons:
                button.Enabled = True
            self.buttons[0].Enabled = True
            self.buttons[0].Caption = "排版为新副本"
            if kind == "error":
                self.notify(payload)
                return
            if kind == 'academic':
                self.last_result = payload
                if payload.get('output'):
                    self.application.Documents.Open(payload['output'])
                import webbrowser
                webbrowser.open(Path(payload['report']).as_uri())
                return
            self.last_result = payload
            if payload.outputs:
                self.application.Documents.Open(payload.outputs[0])
            else:
                self.notify(
                    "未生成结果："
                    + json.dumps(
                        payload.failures or payload.skipped, ensure_ascii=False
                    )
                )
        except Exception as exc:
            self.notify("处理已结束，但无法在 Office 中打开结果：" + str(exc))

    def open_settings(self):
        command = (
            [sys.executable]
            if getattr(sys, "frozen", False)
            else [sys.executable, "-m", "word_formatter"]
        )
        subprocess.Popen(command, creationflags=subprocess.CREATE_NO_WINDOW)

    def saved_source(self):
        if self.application.Documents.Count == 0:
            raise ValueError('请先打开论文。')
        doc = self.application.ActiveDocument
        if not doc.Path or not doc.Saved:
            raise ValueError('请先按 Ctrl+S 保存当前修改。')
        source = Path(doc.FullName)
        if source.suffix.lower() != '.docx' or not source.is_file():
            raise ValueError('论文模式请先另存为本地 DOCX。')
        return source

    def open_academic(self):
        try:
            source = self.saved_source()
            from .academic.plugin import launch
            launch(['--academic', str(source), '--host', self.host])
        except Exception as exc:
            self.notify(str(exc))

    def format_thesis(self, pdf=False, export_only=False):
        if self.busy:
            return
        try:
            source = self.saved_source()
            from .academic.templates import load_template
            template = load_template()
            if export_only:
                template['enabled'] = []
        except Exception as exc:
            self.notify(str(exc))
            return
        self.busy = True
        for button in self.buttons:
            button.Enabled = False
        self.buttons[0].Caption = '正在处理论文…'
        def work():
            try:
                from .academic.workflow import run
                result = run(source, template, reference_path=template['reference_library'] or None, host=self.host, pdf=pdf)
                self.events.put(('academic', result))
            except Exception as exc:
                self.events.put(('error', str(exc)))
        self.worker = threading.Thread(target=work, daemon=False)
        self.worker.start()

    def export_academic(self):
        self.format_thesis(pdf=True, export_only=True)

    def insert_academic(self):
        if not self.busy:
            from .academic.plugin import insert_dialog
            insert_dialog(self.application)

    def update_academic(self):
        try:
            if not self.application.Documents.Count:
                raise ValueError('请先打开论文。')
            from .academic.office_io import update_fields
            result = update_fields(self.application.ActiveDocument)
            self.notify('目录与引用已更新，请保存文档。' if not result['errors'] else '部分字段更新失败：' + str(result['errors'][0]))
        except Exception as exc:
            self.notify(str(exc))

    def about(self):
        self.notify(
            f"ThesisCraft {__version__}\n作者：Study-Tang\n"
            "当前文档需先保存。排版始终生成新文件。\n"
            "在桌面版“保存为默认”后，插件会在下一次排版时读取新配置。"
        )

    def request_close(self):
        if self.busy:
            self.notify("正在保存排版副本，请等待完成后再关闭插件。")
            return
        self.closed = True

    def close(self):
        if self.worker and self.worker.is_alive():
            self.worker.join()
        for handler in self.handlers:
            handler.close()
        self.handlers.clear()
        try:
            self.bar.Delete()
        except Exception:
            pass


def main(argv=None):
    parser = argparse.ArgumentParser(description="启动 Word/WPS 排版工具栏插件")
    parser.add_argument("host", choices=["word", "wps"])
    args = parser.parse_args(argv)
    if os.name != "nt":
        parser.error("此插件入口需要 Windows；桌面版和 CLI 支持其他平台。")
    import pythoncom
    import win32api
    import win32com.client
    import win32event
    import winerror

    mutex = win32event.CreateMutex(
        None, False, "Local\\StudyTangWordFormatter_" + args.host
    )
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        OfficeToolbar._notify("此 Office 插件已启动，请查看现有排版工具栏。")
        win32api.CloseHandle(mutex)
        return 0
    pythoncom.CoInitialize()
    toolbar = None
    try:
        program = "Word.Application" if args.host == "word" else "KWPS.Application"
        try:
            application = win32com.client.GetActiveObject(program)
        except pythoncom.com_error:
            application = win32com.client.DispatchEx(program)
        application.Visible = True
        if application.Documents.Count == 0:
            application.Documents.Add()
        toolbar = OfficeToolbar(application, args.host)
        while not toolbar.closed:
            if pythoncom.PumpWaitingMessages():
                break
            toolbar.poll()
            try:
                application.Name
            except pythoncom.com_error:
                break
            time.sleep(0.05)
    except Exception as exc:
        OfficeToolbar._notify(
            "无法启动插件，请确认已安装桌面版 Word/WPS：\n" + str(exc)
        )
        return 1
    finally:
        if toolbar:
            toolbar.close()
        pythoncom.CoUninitialize()
        win32api.CloseHandle(mutex)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
