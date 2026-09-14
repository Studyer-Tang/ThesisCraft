"""Convenience launchers and opt-in login-time Office connection."""

import os
from pathlib import Path
import subprocess
import sys
import time


def command(arguments):
    if getattr(sys, "frozen", False):
        return [sys.executable, *arguments]
    python = Path(sys.executable)
    if os.name == "nt" and python.with_name("pythonw.exe").exists():
        python = python.with_name("pythonw.exe")
    return [str(python), "-m", "word_formatter", *arguments]


def launch(arguments):
    return subprocess.Popen(
        command(arguments), creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0)
    )


def set_autostart(enabled):
    if os.name != "nt":
        raise RuntimeError("自动连接插件仅支持 Windows。")
    import win32com.client

    shell = win32com.client.Dispatch("WScript.Shell")
    path = Path(shell.SpecialFolders("Startup")) / "Study-Tang Paper Studio.lnk"
    # Exact user Startup shortcut only; no host trust policies are changed.
    if enabled:
        args = command(["--office-watch"])
        shortcut = shell.CreateShortcut(str(path))
        shortcut.TargetPath = args[0]
        shortcut.Arguments = subprocess.list2cmdline(args[1:])
        shortcut.WorkingDirectory = str(Path(args[0]).parent)
        shortcut.WindowStyle = 7
        shortcut.Description = "等待打开的 Word/WPS 并连接学研排版工具栏"
        shortcut.Save()
    elif path.exists():
        path.unlink()
    marker = (
        Path(os.environ["APPDATA"])
        / "Study-Tang"
        / "ThesisCraft"
        / "PaperStudio"
        / "watch-enabled"
    )
    marker.parent.mkdir(parents=True, exist_ok=True)
    if enabled:
        marker.write_text("enabled", encoding="ascii")
    elif marker.exists():
        marker.unlink()
    if enabled:
        launch(["--office-watch"])


def insert_dialog(application):
    """A compact modal chooser; Office owns the actual selected document and undo action."""
    from ..gui import _create_root
    import tkinter as tk
    from tkinter import ttk, messagebox

    root, _ = _create_root()
    root.title("插入编号与引用 · 学研排版")
    root.geometry("610x340")
    kind = tk.StringVar(value="图片题注")
    key = tk.StringVar()
    text = tk.StringVar()
    ttk.Label(root, text="在 Word/WPS 中先将光标放到目标位置。", padding=12).pack(
        anchor="w"
    )
    ttk.Combobox(
        root,
        textvariable=kind,
        state="readonly",
        values=[
            "图片题注",
            "表格题注",
            "公式编号",
            "交叉引用",
            "文献引用",
            "脚注",
            "尾注",
        ],
        width=26,
    ).pack(anchor="w", padx=12)
    ttk.Label(root, text="标识（如 figure1 / ref1；脚注尾注无需填写）").pack(
        anchor="w", padx=12, pady=(12, 0)
    )
    ttk.Entry(root, textvariable=key, width=60).pack(anchor="w", padx=12)
    ttk.Label(root, text="题注文字或注释内容").pack(anchor="w", padx=12, pady=(12, 0))
    ttk.Entry(root, textvariable=text, width=60).pack(anchor="w", padx=12)
    targets = []
    if application.Documents.Count:
        for bm in application.ActiveDocument.Bookmarks:
            if bm.Name.startswith("PS_"):
                targets.append((str(bm.Range.Text)[:65], str(bm.Name)))
    chosen = tk.StringVar()
    labels = [f"{i + 1}. {t[0]}" for i, t in enumerate(targets)]
    ttk.Combobox(
        root, textvariable=chosen, values=labels, width=65, state="readonly"
    ).pack(anchor="w", padx=12, pady=8)

    def insert():
        try:
            import re

            if not application.Documents.Count:
                raise ValueError("请先打开文档。")
            doc = application.ActiveDocument
            selection = application.Selection
            if kind.get() in ("脚注", "尾注"):
                if not text.get().strip():
                    raise ValueError("请输入注释内容。")
                collection = doc.Footnotes if kind.get() == "脚注" else doc.Endnotes
                collection.Add(Range=selection.Range, Text=text.get())
            elif kind.get() == "交叉引用":
                if chosen.get() not in labels:
                    raise ValueError("请在下方选择一个已编号目标。")
                name = targets[labels.index(chosen.get())][1]
                doc.Fields.Add(
                    Range=selection.Range,
                    Type=-1,
                    Text=f"REF {name} \\h",
                    PreserveFormatting=True,
                )
            else:
                if not re.fullmatch(r"[\w.-]+", key.get()):
                    raise ValueError("标识只能包含文字、数字、点、短横线或下划线。")
                token = {
                    "图片题注": "fig",
                    "表格题注": "table",
                    "公式编号": "eq",
                    "文献引用": "cite",
                }[kind.get()]
                marker = (
                    "{{"
                    + token
                    + ":"
                    + key.get()
                    + "}}"
                    + (" " + text.get() if text.get() else "")
                )
                selection.TypeText(marker)
            root.destroy()
        except Exception as exc:
            messagebox.showerror("插入失败", str(exc), parent=root)

    ttk.Button(root, text="插入到当前光标位置", command=insert).pack(
        anchor="w", padx=12, pady=8
    )
    root.mainloop()


def watch():
    if os.name != "nt":
        return 1
    import pythoncom, win32com.client, win32event, win32api, winerror

    mutex = win32event.CreateMutex(None, False, "Local\\StudyTangPaperStudioWatcher")
    if win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        win32api.CloseHandle(mutex)
        return 0
    pythoncom.CoInitialize()
    marker = (
        Path(os.environ["APPDATA"])
        / "Study-Tang"
        / "ThesisCraft"
        / "PaperStudio"
        / "watch-enabled"
    )
    connected = set()
    try:
        while marker.exists():
            for host, program in [
                ("word", "Word.Application"),
                ("wps", "KWPS.Application"),
            ]:
                try:
                    app = win32com.client.GetActiveObject(program)
                    if not app.Documents.Count:
                        continue
                    identity = (host, str(app.Hwnd))
                    if identity not in connected:
                        launch(["--office", host])
                        connected.add(identity)
                except Exception:
                    connected = {c for c in connected if c[0] != host}
            time.sleep(3)
    finally:
        pythoncom.CoUninitialize()
        win32api.CloseHandle(mutex)
    return 0
