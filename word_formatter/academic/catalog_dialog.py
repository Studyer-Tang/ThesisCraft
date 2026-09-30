"""Searchable offline library; source files are exported without conversion."""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webbrowser

from . import catalog


class CatalogDialog(tk.Toplevel):
    def __init__(self, parent, on_apply, on_document):
        super().__init__(parent)
        self.title("高校规范与 Word 模板")
        self.geometry("980x680")
        self.minsize(780, 560)
        self.transient(parent)
        self.on_apply, self.on_document = on_apply, on_document
        self.query, self.category = tk.StringVar(), tk.StringVar(value="全部类型")
        self.rows = {}
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame, text="排版配置可离线使用；官方原件通过来源链接获取，或选择已下载的样稿。"
        ).pack(anchor="w")
        filters = ttk.Frame(frame)
        filters.pack(fill="x", pady=10)
        entry = ttk.Entry(filters, textvariable=self.query)
        entry.pack(side="left", fill="x", expand=True)
        combo = ttk.Combobox(
            filters,
            textvariable=self.category,
            values=["全部类型", *catalog.CATEGORIES.values()],
            state="readonly",
            width=14,
        )
        combo.pack(side="left", padx=(8, 0))
        self.query.trace_add("write", self.search)
        combo.bind("<<ComboboxSelected>>", self.search)
        self.tree = ttk.Treeview(
            frame,
            columns=("school", "category", "version", "department"),
            show="headings",
            height=9,
            selectmode="browse",
        )
        for key, label, width in [
            ("school", "学校", 240),
            ("category", "类型", 80),
            ("version", "规范版本", 180),
            ("department", "适用院系", 240),
        ]:
            self.tree.heading(key, text=label)
            self.tree.column(key, width=width, minwidth=60)
        self.tree.pack(fill="both", expand=True)
        self.tree.bind("<<TreeviewSelect>>", self.show_details)
        self.count = ttk.Label(frame)
        self.count.pack(anchor="w", pady=4)
        details_frame = ttk.Frame(frame)
        details_frame.pack(fill="both", expand=True)
        self.details = tk.Text(
            details_frame, height=11, wrap="word", state="disabled", padx=8, pady=8
        )
        scroll = ttk.Scrollbar(details_frame, command=self.details.yview)
        self.details.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.details.pack(fill="both", expand=True)
        actions = ttk.Frame(frame)
        actions.pack(fill="x", pady=(12, 0))
        self.apply_button = ttk.Button(actions, text="使用排版设置", command=self.apply)
        self.apply_button.pack(side="left")
        self.word_button = ttk.Button(actions, text="新建 Word 文档…", command=self.new)
        self.word_button.pack(side="left", padx=8)
        self.bundle_button = ttk.Button(
            actions, text="导出配置与来源…", command=self.export
        )
        self.bundle_button.pack(side="left")
        self.import_button = ttk.Button(actions, text="导入官方 DOCX…", command=self.import_original)
        self.import_button.pack(side="left", padx=8)
        self.link_button = ttk.Button(
            actions, text="查看官方来源", command=self.open_source
        )
        self.link_button.pack(side="right")
        self.search()
        self.grab_set()
        entry.focus_set()

    def selected(self):
        selection = self.tree.selection()
        return self.rows.get(selection[0]) if selection else None

    def search(self, *_):
        category = next(
            (k for k, v in catalog.CATEGORIES.items() if v == self.category.get()), ""
        )
        self.rows = {p["id"]: p for p in catalog.profiles(self.query.get(), category)}
        self.tree.delete(*self.tree.get_children())
        for key, p in self.rows.items():
            self.tree.insert(
                "",
                "end",
                iid=key,
                values=(
                    p["school"],
                    catalog.CATEGORIES[p["category"]],
                    p["version"],
                    p["department"] or "校级",
                ),
            )
        self.count.configure(
            text=f"找到 {len(self.rows)} 项 · 收录核对日期 2026-09-30 · 以学校最新通知为准"
        )
        if self.rows:
            self.tree.selection_set(next(iter(self.rows)))
        self.show_details()

    def show_details(self, *_):
        p = self.selected()
        for button in (
            self.apply_button,
            self.word_button,
            self.bundle_button,
            self.link_button,
            self.import_button,
        ):
            button.configure(state="normal" if p else "disabled")
        text = "没有匹配的模板，请缩短关键词或更改类型。"
        if p:
            t = catalog.profile_template(p)
            if not t["enabled"]:
                self.apply_button.configure(state="disabled")
            self.word_button.configure(text="新建可编辑 DOCX…")
            self.import_button.configure(state="normal" if p.get("word_source") else "disabled")
            text = p["name"] + "\n\n" + t["scope"] + "\n\n需要核对：\n"
            text += "\n".join("• " + item for item in t["manual_checks"])
            text += "\n\n原始资料：\n" + "\n".join(
                f"• {s['title']}（{s['version']}）\n  {s.get('page') or s['url']}"
                for s in catalog.sources(p)
            )
            text += "\n\n新建 DOCX 为本软件生成的起草文件，不含官方封面。目录域请在 Word/WPS 中更新。原件未随包分发。"
        self.details.configure(state="normal")
        self.details.delete("1.0", "end")
        self.details.insert("1.0", text)
        self.details.configure(state="disabled")

    def apply(self):
        if p := self.selected():
            self.on_apply(catalog.profile_template(p))
            self.destroy()

    def import_original(self):
        path = filedialog.askopenfilename(parent=self, title="选择从学校官网下载的 DOCX",
                                         filetypes=[("Word 文档", "*.docx")])
        if path:
            self.new(original_path=path)

    def new(self, original_path=None):
        if not (p := self.selected()):
            return
        path = filedialog.asksaveasfilename(
            parent=self,
            initialfile=p["id"] + ".docx",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
        )
        if path:
            try:
                catalog.new_document(p, path, overwrite=True, original_path=original_path)
                self.on_apply(catalog.profile_template(p))
                self.on_document(path)
                self.destroy()
            except Exception as exc:
                messagebox.showerror("无法新建", str(exc), parent=self)

    def export(self):
        if not (p := self.selected()):
            return
        path = filedialog.asksaveasfilename(
            parent=self,
            initialfile=p["id"] + ".zip",
            defaultextension=".zip",
            filetypes=[("资料包", "*.zip")],
        )
        if path:
            try:
                catalog.export_bundle(p, path, overwrite=True)
                messagebox.showinfo(
                    "已保存",
                    "已导出可导入排版设置和官方来源记录，不含第三方原件。\n" + path,
                    parent=self,
                )
            except Exception as exc:
                messagebox.showerror("无法保存", str(exc), parent=self)

    def open_source(self):
        if p := self.selected():
            webbrowser.open(catalog.profile_template(p)["sources"][0])
