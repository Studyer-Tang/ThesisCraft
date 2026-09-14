"""Shared, lightweight desktop styling and explicit output-folder selection."""

import os
from pathlib import Path
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, ttk

BG = '#f3f5f8'
INK = '#172b44'
MUTED = '#5c6c80'
ACCENT = '#246b68'


def apply_theme(root):
    from tkinter import font
    families = set(font.families(root))
    family = next((name for name in ('Microsoft YaHei UI', 'Microsoft YaHei', 'PingFang SC', 'Noto Sans CJK SC') if name in families), 'Arial')
    style = ttk.Style(root)
    style.theme_use('clam')
    if 'vista' in style.theme_names():
        style.element_create('Clean.Checkbutton.indicator', 'from', 'vista', 'Checkbutton.indicator')
        style.layout('TCheckbutton', [('Checkbutton.padding', {'sticky': 'nswe', 'children': [
            ('Clean.Checkbutton.indicator', {'side': 'left', 'sticky': ''}),
            ('Checkbutton.focus', {'side': 'left', 'sticky': 'w', 'children': [('Checkbutton.label', {'sticky': 'nswe'})]})
        ]})])
    root.configure(background=BG)
    style.configure('.', font=(family, 10), background=BG, foreground=INK)
    style.configure('TFrame', background=BG)
    style.configure('TLabel', background=BG)
    style.configure('Muted.TLabel', foreground=MUTED)
    style.configure('Title.TLabel', font=(family, 21, 'bold'))
    style.configure('Card.TFrame', background='white')
    style.configure('Card.TLabel', background='white')
    style.configure('TButton', padding=(12, 7), background='white', bordercolor='#dce3eb', focusthickness=1, focuscolor=ACCENT)
    style.map('TButton', background=[('active', '#e7efef'), ('disabled', BG)], foreground=[('disabled', '#96a0af')])
    style.configure('Primary.TButton', background=ACCENT, foreground='white', font=(family, 10, 'bold'), bordercolor=ACCENT)
    style.map('Primary.TButton', background=[('disabled', '#b3c9c7'), ('active', '#195654')], foreground=[('disabled', 'white')])
    style.configure('TEntry', padding=6, fieldbackground='white', bordercolor='#dce3eb')
    style.configure('TCombobox', padding=5, fieldbackground='white', bordercolor='#dce3eb', arrowsize=14)
    style.map('TCombobox', fieldbackground=[('readonly', 'white')], foreground=[('readonly', INK)])
    style.configure('TNotebook', background=BG, borderwidth=0)
    style.configure('TNotebook.Tab', padding=(13, 10), background=BG, borderwidth=0)
    style.map('TNotebook.Tab', background=[('selected', 'white')], foreground=[('selected', ACCENT)])
    style.configure('TLabelframe', bordercolor='#dce3eb', relief='solid')
    style.configure('TLabelframe.Label', foreground=MUTED)
    style.configure('Treeview', rowheight=30, fieldbackground='white', background='white', bordercolor='#dce3eb')
    style.configure('Treeview.Heading', padding=7, background='#e9eef4')
    style.map('Treeview', background=[('selected', '#dcefeb')], foreground=[('selected', INK)])
    style.configure('Horizontal.TProgressbar', background=ACCENT, troughcolor='#e3e9ef', borderwidth=0, thickness=4)


def output_directory(value):
    """An empty selection means beside the source; invalid selections never fall back."""
    value = str(value).strip() if value is not None else ''
    if not value:
        return None
    path = Path(value).expanduser().resolve()
    if not path.is_dir():
        raise ValueError('所选保存文件夹不存在，请重新选择，或点击“恢复默认”。')
    return str(path)


def open_folder(path):
    path = str(Path(path).resolve())
    if os.name == 'nt':
        os.startfile(path)
    else:
        subprocess.Popen(['open' if sys.platform == 'darwin' else 'xdg-open', path])


class OutputFolder(ttk.Frame):
    def __init__(self, parent, value='', source=None):
        super().__init__(parent)
        self.value = tk.StringVar(self, value=value)
        self.display = tk.StringVar(self)
        self.source = source
        self.choosing = False
        ttk.Label(self, text='保存位置', style='Muted.TLabel').pack(side='left', padx=(0, 10))
        self.entry = ttk.Entry(self, textvariable=self.display, state='readonly', width=16)
        self.entry.pack(side='left', fill='x', expand=True)
        self.choose_button = ttk.Button(self, text='选择文件夹…', command=self.choose)
        self.choose_button.pack(side='left', padx=(8, 4))
        self.reset_button = ttk.Button(self, text='恢复默认', command=lambda: self.value.set(''))
        self.reset_button.pack(side='left')
        self.value.trace_add('write', self.refresh)
        self.refresh()

    def refresh(self, *_args):
        self.display.set(self.value.get() or '与原文件相同')
        self.entry.xview_moveto(1 if self.value.get() else 0)

    def choose(self):
        initial = self.value.get()
        if not initial and self.source:
            try:
                source = self.source()
                if source:
                    initial = str(Path(source).resolve().parent)
            except Exception:
                pass
        if not initial or not Path(initial).is_dir():
            initial = str(Path.home())
        self.choosing = True
        try:
            chosen = filedialog.askdirectory(parent=self.winfo_toplevel(), title='选择排版结果的保存文件夹', initialdir=initial, mustexist=True)
        finally:
            self.choosing = False
        if chosen:
            self.value.set(chosen)

    def get(self):
        return output_directory(self.value.get())

    def set_busy(self, busy):
        for button in (self.choose_button, self.reset_button):
            button.configure(state='disabled' if busy else 'normal')
