"""Live COM toolbar for Windows Word/WPS, without changing Office trust settings."""

import argparse
import json
import logging
import os
from pathlib import Path
import queue
import subprocess
import sys
import threading
import time
from types import SimpleNamespace

from .config import normalize_config
from .engine import WordProcessor
from .jobs import build_jobs
from .service import run_jobs
from .storage import user_config_path
from .version import __version__

log = logging.getLogger(__name__)


class OfficeToolbar:
    def __init__(self, application, host, notify=None, allow_panel_fallback=False, dialog_parent=None):
        self.application = application
        self.host = host
        self.events = queue.Queue()
        self.busy = False
        self.worker = None
        self.closed = False
        self.last_result = None
        self.notify = notify or self._notify
        self.dialog_parent = dialog_parent
        self.bar = None
        self.buttons = []
        self.handlers = []
        try:
            self._attach_buttons()
        except Exception:
            self.close()
            if not allow_panel_fallback:
                raise
            log.exception('Office toolbar unavailable; using the companion panel')
            self.buttons = [SimpleNamespace(Enabled=True, Caption=caption)
                            for caption, _ in self.actions()]

    def actions(self):
        return (
            ("排版为新副本", self.format_current),
            ("快速论文排版", self.format_thesis),
            ("论文设置 / 检查", self.open_academic),
            ("插入编号 / 引用", self.insert_academic),
            ("更新目录与引用", self.update_academic),
            ("导出论文 PDF", self.export_academic),
            ("排版设置 / 桌面版", self.open_settings),
            ("关于 Study-Tang", self.about),
            ("关闭排版插件", self.request_close),
        )

    def _attach_buttons(self):
        import win32com.client

        self.bar = self.application.CommandBars.Add(
            Name=f"学研排版 {os.getpid()}", Position=1, Temporary=True
        )
        for caption, action in self.actions():
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
                warnings = payload.get('warnings', [])
                if warnings:
                    self.notify('副本已生成。请注意：\n' + '\n'.join(
                        str(item['message']) for item in warnings[:3]))
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
            insert_dialog(self.application, parent=self.dialog_parent)

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
        handlers, self.handlers = self.handlers, []
        for handler in handlers:
            try:
                handler.close()
            except Exception:
                # Word may have exited before COM event unsubscription runs.
                # Cleanup must neither abort remaining cleanup nor mask errors.
                log.debug('Office event source disconnected during cleanup', exc_info=True)
        bar, self.bar = self.bar, None
        self.buttons.clear()
        try:
            if bar is not None:
                bar.Delete()
        except Exception:
            log.debug('Office toolbar already disconnected', exc_info=True)


class OfficePanel:
    """Visible actions when an Office version hides legacy CommandBars."""

    def __init__(self, root, host):
        import tkinter as tk
        from tkinter import ttk

        self.root, self.host, self.toolbar = root, host, None
        self.closing = False
        self.last_probe = 0
        label = 'Word' if host == 'word' else 'WPS'
        root.title(f'ThesisCraft · {label}')
        root.geometry('620x290')
        root.minsize(560, 280)
        root.attributes('-topmost', True)
        frame = ttk.Frame(root, padding=16)
        frame.pack(fill='both', expand=True)
        ttk.Label(frame, text='学研排版', font=('Microsoft YaHei', 18, 'bold')).pack(anchor='w')
        self.status = tk.StringVar(value='正在连接当前文档…')
        ttk.Label(frame, textvariable=self.status, wraplength=580).pack(anchor='w', pady=(4, 10))
        grid = ttk.Frame(frame)
        grid.pack(fill='x')
        self.action_buttons = []
        for index, (caption, method) in enumerate([
            ('论文设置 / 检查', 'open_academic'), ('快速论文排版', 'format_thesis'),
            ('插入编号 / 引用', 'insert_academic'), ('更新目录与引用', 'update_academic'),
            ('导出论文 PDF', 'export_academic'), ('通用文档排版', 'format_current'),
        ]):
            button = ttk.Button(grid, text=caption, command=lambda m=method: self.invoke(m))
            button.grid(row=index // 3, column=index % 3, padx=3, pady=4, sticky='ew')
            self.action_buttons.append(button)
        for column in range(3):
            grid.columnconfigure(column, weight=1)
        bottom = ttk.Frame(frame)
        bottom.pack(fill='x', pady=(12, 0))
        self.connect_button = ttk.Button(bottom, text='连接当前文档', command=self.connect)
        self.connect_button.pack(side='left')
        self.pin = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom, text='保持置顶', variable=self.pin,
                        command=lambda: root.attributes('-topmost', self.pin.get())).pack(side='left', padx=12)
        ttk.Button(bottom, text='关闭插件', command=self.close).pack(side='right')
        root.protocol('WM_DELETE_WINDOW', self.close)
        self.connect()
        root.after(80, self.tick)

    def notify(self, message):
        from tkinter import messagebox
        messagebox.showinfo('学研排版', message, parent=self.root)

    def connect(self):
        from .office_connection import connect_application
        if self.toolbar and self.toolbar.busy:
            self.status.set('正在处理文档，请等待完成后再切换。')
            return
        if self.toolbar:
            self.toolbar.close()
            self.toolbar = None
        try:
            application = connect_application(self.host)
            self.toolbar = OfficeToolbar(application, self.host, self.notify,
                                         allow_panel_fallback=True, dialog_parent=self.root)
            self.status.set('已连接：' + str(application.ActiveDocument.Name))
        except Exception as exc:
            if self.toolbar:
                self.toolbar.close()
                self.toolbar = None
            self.status.set(str(exc))
            log.exception('Unable to connect Office document')
        self.update_buttons()

    def update_buttons(self):
        enabled = self.toolbar is not None and not self.toolbar.busy
        for button in self.action_buttons:
            button.configure(state='normal' if enabled else 'disabled')
        self.connect_button.configure(state='disabled' if self.toolbar and self.toolbar.busy else 'normal')

    def invoke(self, method):
        if not self.toolbar or self.toolbar.busy:
            return
        try:
            getattr(self.toolbar, method)()
        except Exception as exc:
            self.notify(str(exc))
        self.update_buttons()

    def tick(self):
        import pythoncom
        from .office_connection import is_busy_error
        if self.closing:
            return
        pythoncom.PumpWaitingMessages()
        toolbar = self.toolbar
        if toolbar:
            toolbar.poll()
            if toolbar.closed:
                self.close()
                return
            if time.monotonic() - self.last_probe > 0.75:
                self.last_probe = time.monotonic()
                try:
                    name = str(toolbar.application.ActiveDocument.Name)
                    self.status.set(('正在处理：' if toolbar.busy else '已连接：') + name)
                except Exception as exc:
                    if is_busy_error(exc):
                        self.status.set('Word/WPS 正忙，请完成对话框操作后继续。')
                    elif toolbar.busy:
                        self.status.set('文档窗口已关闭，正在等待当前排版任务完成。')
                    else:
                        toolbar.close()
                        self.toolbar = None
                        self.status.set('文档连接已断开。打开文档后，点击“连接当前文档”。')
            self.update_buttons()
        self.root.after(80, self.tick)

    def close(self):
        if self.toolbar and self.toolbar.busy:
            self.notify('正在处理文档，请等待完成后再关闭插件。')
            return
        self.closing = True
        if self.toolbar:
            self.toolbar.close()
        self.root.destroy()


def main(argv=None):
    parser = argparse.ArgumentParser(description="启动 Word/WPS 排版工具栏插件")
    parser.add_argument("host", choices=["word", "wps"])
    args = parser.parse_args(argv)
    if os.name != "nt":
        parser.error("此插件入口需要 Windows；桌面版和 CLI 支持其他平台。")
    import pythoncom
    import win32api
    import win32event
    import winerror

    mutex = win32event.CreateMutex(
        None, False, "Local\\StudyTangWordFormatter_" + args.host
    )
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        import win32gui
        label = 'Word' if args.host == 'word' else 'WPS'
        hwnd = win32gui.FindWindow(None, f'ThesisCraft · {label}')
        if hwnd:
            win32gui.ShowWindow(hwnd, 9)
            try:
                win32gui.SetForegroundWindow(hwnd)
            except Exception:
                win32gui.FlashWindow(hwnd, True)
        else:
            OfficeToolbar._notify("插件已在运行。若有旧版错误提示，请先关闭提示，再重新启动插件。")
        win32api.CloseHandle(mutex)
        return 0
    pythoncom.CoInitialize()
    folder = user_config_path().parent
    folder.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(filename=str(folder / 'office-plugin.log'), encoding='utf-8', level=logging.INFO)
    panel = None
    try:
        from .gui import _create_root
        root, _ = _create_root()
        panel = OfficePanel(root, args.host)
        root.mainloop()
    except Exception as exc:
        OfficeToolbar._notify(
            "无法启动插件，请确认已安装桌面版 Word/WPS：\n" + str(exc)
        )
        return 1
    finally:
        try:
            if panel and panel.toolbar:
                panel.toolbar.close()
        finally:
            pythoncom.CoUninitialize()
            win32api.CloseHandle(mutex)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
