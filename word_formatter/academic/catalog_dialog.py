"""Searchable offline library; source files are exported without conversion."""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webbrowser

from . import catalog


def _public_links(profile):
    """Community references may have provenance metadata without a public URL."""
    return [
        url for url in catalog.profile_template(profile)["sources"]
        if isinstance(url, str) and url.startswith(("https://", "http://"))
    ]


class CatalogDialog(tk.Toplevel):
    def __init__(self, parent, on_apply, on_document):
        super().__init__(parent)
        self.title("高校规范与课程参考模板")
        self.geometry("980x680")
        self.minsize(780, 560)
        self.transient(parent)
        self.on_apply, self.on_document = on_apply, on_document
        self.query, self.category = tk.StringVar(), tk.StringVar(value="全部类型")
        self.rows = {}
        frame = ttk.Frame(self, padding=16)
        frame.pack(fill="both", expand=True)
        ttk.Label(
            frame, text="排版配置可离线使用；公开规范提供来源链接，社区参考格式单独说明来历。"
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
            actions, text="查看公开来源", command=self.open_source
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
            text=f"找到 {len(self.rows)} 项 · 具体来源及核对日期见详情 · 以课程或学校要求为准"
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
            self.link_button.configure(state="normal" if _public_links(p) else "disabled")
            text = p["name"] + "\n\n" + t["scope"] + "\n\n需要核对：\n"
            text += "\n".join("• " + item for item in t["manual_checks"])
            records = []
            for source in catalog.sources(p):
                record = f"• {source['title']}（{source['version']}）"
                if source.get("kind") == "user-provided":
                    record += "\n  用户提供的社区参考资料；无已核验公开链接。"
                    record += "\n  " + source["summary"]
                    record += "\n  核对日期：" + source["retrieved"]
                    record += "\n  SHA-256：" + source["sha256"]
                    record += "\n  仅保存配置摘要与校验信息，原件未随包分发。"
                else:
                    record += "\n  " + (source.get("page") or source.get("url") or "无公开链接")
                records.append(record)
            text += "\n\n来源与来历：\n" + "\n".join(records)
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
                    "已导出可导入排版设置和来源记录，不含第三方原件。\n" + path,
                    parent=self,
                )
            except Exception as exc:
                messagebox.showerror("无法保存", str(exc), parent=self)

    def open_source(self):
        if p := self.selected():
            links = _public_links(p)
            if links:
                webbrowser.open(links[0])
