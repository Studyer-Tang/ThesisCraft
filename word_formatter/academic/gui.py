"""Task-oriented academic workspace with editable styles and a shared Office workflow."""

import argparse
from copy import deepcopy
from pathlib import Path
import queue
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import webbrowser

from .templates import (
    load_template,
    save_template,
    validate_template,
    STYLE_LABELS,
    GROUPS,
    NUMBERING,
    import_word_styles,
)
from .structure import scan
from .workflow import run

PRESETS = {
    "北京大学硕士（研究生指南）": "pku-master",
    "北京大学博士（研究生指南）": "pku-doctor",
    "本科通用": "bachelor",
    "硕士通用": "master",
    "博士通用": "doctor",
}
CHOICES = {
    "degree": {"本科": "bachelor", "硕士": "master", "博士": "doctor"},
    "scheme": {v: k for k, v in NUMBERING.items()},
    "align": {
        "左对齐": "left",
        "居中": "center",
        "右对齐": "right",
        "两端对齐": "justify",
    },
    "spacing_unit": {"倍数": "multiple", "固定磅值": "pt"},
    "front_number": {
        "I, II, III": "upperRoman",
        "i, ii, iii": "lowerRoman",
        "1, 2, 3": "decimal",
    },
    "number_position": {"居中": "center", "奇偶页外侧": "outside"},
    "appendix_format": {"A, B, C": "upperLetter", "1, 2, 3": "decimal"},
    "style": {
        "GB/T 7714-2015 顺序编码": "gb7714-numeric",
        "APA 作者年份": "apa",
        "IEEE 数字引用": "ieee",
    },
    "number_format": {
        "1, 2, 3": "decimal",
        "①②③": "decimalEnclosedCircle",
        "i, ii, iii": "lowerRoman",
        "I, II, III": "upperRoman",
        "a, b, c": "lowerLetter",
    },
    "restart": {"连续": "continuous", "每页重置": "eachPage", "每节重置": "eachSect"},
}
LABELS = {
    "font": "中文字体",
    "latin": "英文/数字字体",
    "size": "字号（磅）",
    "bold": "加粗",
    "align": "对齐",
    "spacing": "行距",
    "spacing_unit": "行距单位",
    "before": "段前（磅）",
    "after": "段后（磅）",
    "indent": "首行缩进（字）",
    "color": "文字颜色（如000000）",
    "scheme": "章编号样式",
    "east_asia": "编号中文字体",
    "separator": "图表章号分隔符",
    "by_chapter": "图表公式按章编号",
    "equation_brackets": "公式括号",
    "appendix_format": "附录编号",
    "match_heading": "标题编号字体跟随对应标题",
    "match_caption": "题注编号字体跟随对应题注",
    "top": "上边距 cm",
    "bottom": "下边距 cm",
    "left": "左边距 cm",
    "right": "右边距 cm",
    "gutter": "装订线 cm",
    "header_distance": "页眉距离 cm",
    "footer_distance": "页脚距离 cm",
    "front_number": "前置页页码",
    "number_position": "页码位置",
    "odd_even": "奇偶页不同",
    "chapter_recto": "章从奇数页开始",
    "chapter_header": "奇数页显示章节标题",
    "header_text": "固定页眉（可用{school}/{degree}）",
    "preserve_landscape": "保留横向页面",
    "mirror_margins": "左右对称页边距",
    "header_line_pt": "页眉横线（磅，0关闭）",
    "three_line": "三线表",
    "repeat_header": "跨页重复表头",
    "prevent_row_split": "禁止单行跨页",
    "outer_pt": "表格外线（磅）",
    "inner_pt": "表头下线（磅）",
    "landscape_wide": "宽表独立横向页",
    "rows_per_part": "续表每部分数据行（0不拆分）",
    "max_width_cm": "图片最大宽度 cm",
    "max_height_cm": "图片最大高度 cm（0按页面）",
    "center": "独立图片居中",
    "before_pt": "图片段前间距（磅）",
    "after_pt": "图片段后间距（磅）",
    "min_dpi": "最低图片清晰度 DPI",
    "keep_caption": "图表与题注保持相邻",
    "toc": "正文目录",
    "figures": "图目录",
    "tables": "表目录",
    "depth": "目录级数",
    "hanging_cm": "文献悬挂缩进 cm",
    "superscript": "数字引用使用上标",
    "csl_path": "自定义CSL路径（空白用内置）",
    "number_format": "脚注序号",
    "restart": "脚注重新编号",
    "title": "中文题目",
    "title_en": "英文题目",
    "author": "姓名",
    "student_id": "学号",
    "supervisor": "导师",
    "date": "完成年月",
    "keywords": "中文关键词",
    "keywords_en": "英文关键词",
}


class AcademicWindow:
    def __init__(self, root, source="", host="none", output_dir=""):
        self.root = root
        from ..ui_common import apply_theme
        apply_theme(root)
        self.initial_output = output_dir
        self.root.title("学研排版 · 论文工作台 | Study-Tang")
        self.root.geometry("1120x760")
        self.root.minsize(980, 640)
        self.template = load_template()
        self.events, self.cancel = queue.Queue(), threading.Event()
        self.worker = None
        self.report = None
        self.vars, self.style_vars = {}, {}
        self.source = tk.StringVar(value=source)
        self.host = tk.StringVar(
            value={"word": "Word", "wps": "WPS"}.get(host, "稍后手动更新")
        )
        self.pdf = tk.BooleanVar(value=False)
        self.refs = tk.StringVar()
        self.preset = tk.StringVar(value=self.template["name"])
        self.style_key = tk.StringVar(value="正文")
        self.status = tk.StringVar(value="选择论文和模板，即可开始排版。原文档保留。")
        self.current_style = "body"
        self.blocks = []
        self.structure_source = None
        self.scroll_canvases = {}
        self._build()
        self.refresh()
        self.root.after(120, self.poll)
        self.root.protocol("WM_DELETE_WINDOW", self.close)
        self.root.bind("<MouseWheel>", self.scroll_active_tab)

    def _build(self):
        outer = ttk.Frame(self.root, padding=16)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="论文工作台", style="Title.TLabel").pack(
            anchor="w"
        )
        ttk.Label(
            outer, text="把时间留给研究。选好模板，让论文格式保持一致。", style="Muted.TLabel"
        ).pack(anchor="w", pady=(2, 12))
        top = ttk.Frame(outer)
        top.pack(fill="x")
        ttk.Label(top, text="论文文件", style="Muted.TLabel").pack(side="left", padx=(0, 12))
        ttk.Entry(top, textvariable=self.source).pack(
            side="left", fill="x", expand=True
        )
        ttk.Button(top, text="选择论文…", command=self.select_source).pack(
            side="left", padx=6
        )
        ttk.Button(top, text="新建论文骨架", command=self.skeleton).pack(side="left")
        presets = ttk.Frame(outer)
        presets.pack(fill="x", pady=10)
        ttk.Label(presets, text="论文模板", style="Muted.TLabel").pack(side="left", padx=(0, 12))
        combo = ttk.Combobox(
            presets,
            textvariable=self.preset,
            values=list(PRESETS),
            state="readonly",
            width=32,
        )
        combo.pack(side="left")
        combo.bind("<<ComboboxSelected>>", self.choose_preset)
        ttk.Button(presets, text="设为插件默认", command=self.save_default).pack(side="left", padx=8)
        more = ttk.Menubutton(presets, text="模板管理 ▾")
        menu = tk.Menu(more, tearoff=False)
        for title, action in [("导入模板", self.import_template), ("从 Word 样稿读取", self.import_styles), ("导出模板", self.export_template)]:
            menu.add_command(label=title, command=action)
        more.configure(menu=menu)
        more.pack(side="left")
        self.scope = ttk.Label(outer, wraplength=1060, foreground="#596779")
        self.scope.pack(anchor="w", pady=(0, 8))
        footer_area = ttk.Frame(outer)
        footer_area.pack(side="bottom", fill="x")
        self.notebook = ttk.Notebook(outer)
        self.notebook.pack(fill="both", expand=True)
        main = self.scroll_tab("1 排版范围")
        self.group_vars = {key: tk.BooleanVar() for key in GROUPS}
        for row, (key, title) in enumerate(GROUPS.items()):
            ttk.Checkbutton(main, text=title, variable=self.group_vars[key]).grid(
                row=row // 2, column=row % 2, sticky="w", padx=25, pady=8
            )
        info = ttk.LabelFrame(main, text="学校和模板版本", padding=10)
        info.grid(row=5, column=0, columnspan=2, sticky="ew", pady=10)
        for i, (key, label) in enumerate(
            [
                ("name", "模板名称"),
                ("school", "学校"),
                ("department", "学院/专业"),
                ("degree", "学位"),
                ("year", "适用年份"),
                ("template_version", "模板版本"),
            ]
        ):
            self.control(
                info, self.vars, key, label, self.template[key], i // 2, (i % 2) * 2
            )
        typography = self.scroll_tab("2 字体与段落")
        chooser = ttk.Combobox(
            typography,
            textvariable=self.style_key,
            values=list(STYLE_LABELS.values()),
            state="readonly",
            width=28,
        )
        chooser.grid(row=0, column=0, columnspan=4, sticky="w", pady=10)
        chooser.bind("<<ComboboxSelected>>", self.change_style)
        for i, (key, value) in enumerate(self.template["styles"]["body"].items()):
            self.control(
                typography,
                self.style_vars,
                key,
                LABELS[key],
                value,
                i // 2 + 1,
                (i % 2) * 2,
            )
        self.preview = ttk.Label(
            typography,
            text="第一章 绪论 / Chapter 1\n学研排版 AaBb 0123456789",
            padding=18,
        )
        self.preview.grid(row=8, column=0, columnspan=4, sticky="ew")
        ttk.Button(typography, text="预览字体与编号", command=self.preview_style).grid(
            row=9, column=0, columnspan=2, sticky="w"
        )
        ttk.Label(
            typography,
            text="字号可自由输入，例如 10.5、11、12、13、14、16。字体可选择已安装字体或手动输入。",
            wraplength=850,
        ).grid(row=10, column=0, columnspan=4, sticky="w", pady=12)
        self.settings_tab(
            "3 编号与页码", [("numbering", "编号"), ("page", "页面"), ("notes", "脚注")]
        )
        self.settings_tab(
            "4 图表与目录",
            [
                ("tables", "表格"),
                ("figures", "图片"),
                ("contents", "目录"),
                ("references", "参考文献"),
            ],
        )
        self.settings_tab("5 封面摘要", [("metadata", "填写后用于新建论文骨架")])
        structure = self.tab("6 结构与文献")
        buttons = ttk.Frame(structure)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="读取文档结构", command=self.read_structure).pack(
            side="left"
        )
        self.kind_var = tk.StringVar(value="正文")
        self.kinds = {
            **{v: k for k, v in STYLE_LABELS.items()},
            "保留不处理": "ignore",
            "参考文献标题": "bibliography_heading",
            "致谢标题": "acknowledgements_heading",
            "符号表标题": "symbols_heading",
            "学术成果标题": "publications_heading",
        }
        ttk.Combobox(
            buttons,
            values=list(self.kinds),
            textvariable=self.kind_var,
            state="readonly",
            width=22,
        ).pack(side="left", padx=8)
        ttk.Button(buttons, text="将选中段落设为此类型", command=self.assign_kind).pack(
            side="left"
        )
        self.tree = ttk.Treeview(
            structure,
            columns=("index", "kind", "text"),
            show="headings",
            height=11,
            selectmode="extended",
        )
        for key, label, width in [
            ("index", "段落", 60),
            ("kind", "识别类型", 150),
            ("text", "内容", 720),
        ]:
            self.tree.heading(key, text=label)
            self.tree.column(key, width=width)
        self.tree.pack(fill="both", expand=True, pady=8)
        ttk.Button(
            structure, text="导入 RIS / BibTeX / JSON 文献库", command=self.select_refs
        ).pack(anchor="w")
        ttk.Label(structure, textvariable=self.refs).pack(anchor="w")
        ttk.Label(
            structure,
            text="在正文写 {{cite:文献key}}；题注可用 {{fig:图片key}} 标题，引用用 {{ref:fig:图片key}}。\n插件提供插入按钮，无需记忆标记格式。原有 Zotero / EndNote 域会保留。",
            wraplength=950,
        ).pack(anchor="w", pady=8)
        tools = self.scroll_tab("7 章节与插件")
        for title, key, pattern in [
            ("选择正式封面/声明 DOCX…", "front_matter_path", "*.docx"),
            ("导入符号/缩略语 CSV…", "glossary_path", "*.csv"),
            ("选择自定义 CSL 文献样式…", "references.csl_path", "*.csl"),
        ]:
            ttk.Button(
                tools,
                text=title,
                command=lambda k=key, p=pattern: self.select_asset(k, p),
            ).pack(anchor="w", pady=4)
        ttk.Button(
            tools,
            text="查看北大官方规范与模板",
            command=lambda: webbrowser.open(
                "https://grs.pku.edu.cn/xwgz11/xwsy11/bsxw111/clxz09/index.htm"
            ),
        ).pack(anchor="w", pady=4)
        for title, action in [
            ("按顺序合并章节…", self.merge),
            ("拆分为章节副本…", self.split),
            ("连接 Word 工具栏", lambda: self.launch_host("word")),
            ("连接 WPS 工具栏", lambda: self.launch_host("wps")),
            ("启用登录后自动连接插件", lambda: self.autostart(True)),
            ("关闭插件自动连接", lambda: self.autostart(False)),
        ]:
            ttk.Button(tools, text=title, command=action).pack(anchor="w", pady=5)
        ttk.Label(
            tools,
            text="自动连接只等待已经打开的 Word/WPS，不会主动打开办公软件。\n合并以第一个文件的页眉页脚为基础；请在合并后应用模板并检查。拆分后跨章引用需在完整论文中更新。",
            wraplength=950,
        ).pack(anchor="w", pady=14)
        from ..ui_common import OutputFolder
        self.output = OutputFolder(footer_area, value=self.initial_output, source=self.source.get)
        self.output.pack(fill="x", pady=(14, 0))
        footer = ttk.Frame(footer_area)
        footer.pack(fill="x", pady=(12, 5))
        ttk.Label(footer, text="更新目录", style="Muted.TLabel").pack(side="left")
        ttk.Combobox(
            footer,
            textvariable=self.host,
            values=["Word", "WPS", "稍后手动更新"],
            state="readonly",
            width=17,
        ).pack(side="left")
        ttk.Checkbutton(footer, text="同时生成 PDF", variable=self.pdf).pack(
            side="left", padx=10
        )
        self.action_buttons = []
        for title, action in [
            ("只检查", lambda: self.process(True)),
            ("排版并生成副本", lambda: self.process(False)),
        ]:
            button = ttk.Button(footer, text=title, command=action, style="Primary.TButton" if title == "排版并生成副本" else "TButton")
            button.pack(side="left", padx=3)
            self.action_buttons.append(button)
        ttk.Button(footer, text="取消", command=self.cancel.set).pack(
            side="left", padx=3
        )
        self.folder_button = ttk.Button(footer, text="结果文件夹", command=self.open_result_folder, state="disabled")
        self.folder_button.pack(side="right", padx=(6, 0))
        ttk.Button(footer, text="打开结果", command=self.open_report).pack(
            side="right"
        )
        track = ttk.Frame(footer_area, height=4)
        track.pack(fill="x", pady=6)
        track.pack_propagate(False)
        self.progress = ttk.Progressbar(track, mode="indeterminate")
        self.progress.pack(fill="both", expand=True)
        ttk.Label(footer_area, textvariable=self.status, wraplength=1050).pack(anchor="w")

    def tab(self, title):
        frame = ttk.Frame(self.notebook, padding=12)
        self.notebook.add(frame, text=title)
        return frame

    def scroll_tab(self, title):
        outer = self.tab(title)
        canvas = tk.Canvas(outer, highlightthickness=0, background="#f3f5f8")
        scroll = ttk.Scrollbar(outer, orient="vertical", command=canvas.yview)
        scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.configure(yscrollcommand=scroll.set)
        content = ttk.Frame(canvas)
        window = canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind(
            "<Configure>", lambda _e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        canvas.bind(
            "<Configure>", lambda e: canvas.itemconfigure(window, width=e.width)
        )
        self.scroll_canvases[str(outer)] = canvas
        return content

    def scroll_active_tab(self, event):
        canvas = self.scroll_canvases.get(self.notebook.select())
        if canvas is not None:
            canvas.yview_scroll(-int(event.delta / 120), "units")

    def settings_tab(self, title, sections):
        content = self.scroll_tab(title)
        for section, label in sections:
            box = ttk.LabelFrame(content, text=label, padding=10)
            box.pack(fill="x", pady=6)
            for i, (key, val) in enumerate(
                (k, v) for k, v in self.template[section].items() if k != "body_number"
            ):
                self.control(
                    box,
                    self.vars,
                    section + "." + key,
                    LABELS.get(key, key),
                    val,
                    i // 2,
                    (i % 2) * 2,
                    choice_key=key,
                )

    def control(self, parent, store, key, label, value, row, col, choice_key=None):
        choice_key = choice_key or key
        var = (
            tk.BooleanVar(value=value)
            if isinstance(value, bool)
            else tk.StringVar(value=str(value))
        )
        store[key] = (var, type(value), CHOICES.get(choice_key))
        if isinstance(value, bool):
            ttk.Checkbutton(parent, text=label, variable=var).grid(
                row=row, column=col, columnspan=2, sticky="w", padx=8, pady=5
            )
            return
        ttk.Label(parent, text=label).grid(
            row=row, column=col, sticky="w", padx=8, pady=5
        )
        values = list(CHOICES.get(choice_key, {}))
        if choice_key in ("font", "latin", "east_asia"):
            from tkinter.font import families

            values = sorted(set(families(self.root)))
        elif choice_key == "size":
            values = [
                "9",
                "10.5",
                "11",
                "12",
                "13",
                "14",
                "15",
                "16",
                "18",
                "22",
                "24",
                "26",
            ]
        elif choice_key == "separator":
            values = ["-", ".", "–"]
        elif choice_key == "equation_brackets":
            values = ["()", "（）", "[]"]
        widget = (
            ttk.Combobox(
                parent,
                textvariable=var,
                values=values,
                width=25,
                state="readonly" if choice_key in CHOICES else "normal",
            )
            if values
            else ttk.Entry(parent, textvariable=var, width=27)
        )
        widget.grid(row=row, column=col + 1, sticky="ew", padx=8, pady=5)
        parent.columnconfigure(col + 1, weight=1)

    @staticmethod
    def get_value(binding):
        var, typ, choices = binding
        value = var.get()
        if choices:
            return choices.get(value, value)
        if typ in (int, float):
            return typ(value)
        return value

    @staticmethod
    def set_value(binding, value):
        var, _typ, choices = binding
        if choices:
            value = next((k for k, v in choices.items() if v == value), value)
        var.set(value)

    def refresh(self):
        self.refs.set(self.template["reference_library"])
        for key, binding in self.vars.items():
            value = self.template
            for part in key.split("."):
                value = value[part]
            self.set_value(binding, value)
        for key, var in self.group_vars.items():
            var.set(key in self.template["enabled"])
        for key, binding in self.style_vars.items():
            self.set_value(binding, self.template["styles"][self.current_style][key])
        self.scope.configure(text=self.template["scope"])
        self.preview_style()

    def collect(self):
        result = deepcopy(self.template)
        if (
            self.structure_source is not None
            and self.source.get() != self.structure_source
        ):
            result["structure_overrides"] = {}
        for key, binding in self.vars.items():
            owner = result
            parts = key.split(".")
            for part in parts[:-1]:
                owner = owner[part]
            owner[parts[-1]] = self.get_value(binding)
        result["styles"][self.current_style] = {
            k: self.get_value(v) for k, v in self.style_vars.items()
        }
        result["enabled"] = [k for k, v in self.group_vars.items() if v.get()]
        result["reference_library"] = self.refs.get()
        return validate_template(result)

    def change_style(self, _event=None):
        try:
            self.template = self.collect()
            self.current_style = next(
                k for k, v in STYLE_LABELS.items() if v == self.style_key.get()
            )
            self.refresh()
        except Exception as exc:
            self.error(exc)

    def preview_style(self):
        try:
            spec = {k: self.get_value(v) for k, v in self.style_vars.items()}
            self.preview.configure(
                font=(
                    spec["font"],
                    int(spec["size"]),
                    "bold" if spec["bold"] else "normal",
                ),
                foreground="#" + spec["color"],
            )
        except Exception:
            pass

    def choose_preset(self, _event=None):
        self.template = load_template(PRESETS[self.preset.get()])
        self.refresh()

    def select_source(self):
        path = filedialog.askopenfilename(
            parent=self.root, filetypes=[("Word论文", "*.docx")]
        )
        if path:
            self.source.set(path)
            self.template["structure_overrides"] = {}
            self.read_structure()

    def import_template(self):
        path = filedialog.askopenfilename(
            parent=self.root, filetypes=[("模板", "*.json")]
        )
        if path:
            try:
                self.template = load_template(path)
                self.preset.set(self.template["name"])
                self.refresh()
            except Exception as exc:
                self.error(exc)

    def import_styles(self):
        path = filedialog.askopenfilename(
            parent=self.root, filetypes=[("Word样稿", "*.docx")]
        )
        if path:
            try:
                self.template, imported = import_word_styles(path, self.collect())
                self.preset.set(self.template["name"])
                self.refresh()
                self.status.set(
                    "已读取："
                    + "、".join(imported)
                    + "。编号、院系规则和适用年份请在设置中确认。"
                )
            except Exception as exc:
                self.error(exc)

    def export_template(self):
        path = filedialog.asksaveasfilename(
            parent=self.root, defaultextension=".json", filetypes=[("模板", "*.json")]
        )
        if path:
            try:
                template = self.collect()
                template["structure_overrides"] = {}
                save_template(template, path)
                self.status.set("模板已导出：" + path)
            except Exception as exc:
                self.error(exc)

    def save_default(self):
        try:
            template = self.collect()
            template["structure_overrides"] = {}
            save_template(template)
            self.status.set("已保存。Word/WPS“快速论文排版”下次使用这些设置。")
        except Exception as exc:
            self.error(exc)

    def skeleton(self):
        path = filedialog.asksaveasfilename(
            parent=self.root,
            defaultextension=".docx",
            filetypes=[("Word论文", "*.docx")],
        )
        if path:
            try:
                from .layout import create_skeleton

                create_skeleton(self.collect(), path)
                self.source.set(path)
                self.status.set("论文骨架已生成；请填写内容并替换学校正式封面和声明。")
            except Exception as exc:
                self.error(exc)

    def read_structure(self):
        try:
            from docx import Document

            if (
                self.structure_source is not None
                and self.source.get() != self.structure_source
            ):
                self.template["structure_overrides"] = {}
            self.structure_source = self.source.get()

            self.blocks = scan(
                Document(self.source.get()), self.template["structure_overrides"]
            )
            self.tree.delete(*self.tree.get_children())
            for block in self.blocks:
                if block.text:
                    self.tree.insert(
                        "",
                        "end",
                        iid=str(block.index),
                        values=(
                            block.index + 1,
                            STYLE_LABELS.get(block.kind, block.kind),
                            block.text[:200],
                        ),
                    )
            self.status.set(f"已读取 {len(self.blocks)} 个段落；可批量选择并校正结构。")
        except Exception as exc:
            self.error(exc)

    def assign_kind(self):
        for ident in self.tree.selection():
            self.template["structure_overrides"][ident] = self.kinds[
                self.kind_var.get()
            ]
        self.read_structure()

    def select_refs(self):
        path = filedialog.askopenfilename(
            parent=self.root, filetypes=[("文献库", "*.ris *.bib *.json")]
        )
        if path:
            try:
                from .bibliography import load_references

                rows = load_references(path)
                self.refs.set(path)
                self.status.set(
                    f"已读取 {len(rows)} 条文献。引用 key 示例："
                    + ", ".join(r["key"] for r in rows[:5])
                )
            except Exception as exc:
                self.error(exc)

    def process(self, check_only):
        if self.worker and self.worker.is_alive():
            return
        try:
            template = self.collect()
            output_dir = self.output.get()
        except Exception as exc:
            self.error(exc)
            return
        source, refs = self.source.get(), self.refs.get() or None
        host = {"Word": "word", "WPS": "wps"}.get(self.host.get())
        pdf = self.pdf.get()
        self.cancel.clear()
        self.progress.start()
        self.output.set_busy(True)
        self.status.set("正在检查论文…" if check_only else "正在生成排版副本…")
        for b in self.action_buttons:
            b.configure(state="disabled")

        def work():
            try:
                self.events.put(
                    (
                        "result",
                        run(
                            source,
                            template,
                            output_dir=output_dir,
                            check_only=check_only,
                            reference_path=refs,
                            host=host,
                            pdf=pdf,
                            cancel=self.cancel,
                            progress=lambda s: self.events.put(("progress", s)),
                        ),
                    )
                )
            except Exception as exc:
                self.events.put(("error", str(exc)))

        self.worker = threading.Thread(target=work, daemon=False)
        self.worker.start()

    def poll(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "progress":
                    self.status.set(payload)
                    continue
                self.progress.stop()
                self.output.set_busy(False)
                for b in self.action_buttons:
                    b.configure(state="normal")
                if kind == "error":
                    self.error(payload)
                else:
                    self.report = payload
                    self.folder_button.configure(state="normal")
                    self.status.set("完成：" + (payload.get("output") or payload["report"]))
                    self.open_report()
                    if payload.get('output') and payload.get('warnings'):
                        messagebox.showwarning('副本已生成，请注意', '\n'.join(
                            str(item['message']) for item in payload['warnings'][:3]), parent=self.root)
        except queue.Empty:
            pass
        self.root.after(120, self.poll)

    def open_result_folder(self):
        if self.report:
            from ..ui_common import open_folder
            path = self.report.get('output') or self.report.get('report')
            if path:
                try:
                    open_folder(Path(path).parent)
                except Exception as exc:
                    self.error(exc)

    def open_report(self):
        if self.report:
            output = self.report.get('output')
            if output:
                from .plugin import open_document
                try:
                    open_document(output, {"Word": "word", "WPS": "wps"}.get(self.host.get()))
                except Exception as exc:
                    self.status.set('副本已保存：' + output + '；自动打开失败：' + str(exc))
            elif self.report.get('report'):
                webbrowser.open(Path(self.report['report']).as_uri())

    def merge(self):
        paths = filedialog.askopenfilenames(
            parent=self.root, filetypes=[("章节", "*.docx")]
        )
        if not paths:
            return
        win = tk.Toplevel(self.root)
        win.title("确认章节顺序")
        win.geometry("720x400")
        box = tk.Listbox(win, selectmode="single")
        box.pack(fill="both", expand=True, padx=12, pady=12)
        for path in paths:
            box.insert("end", path)

        def move(delta):
            selection = box.curselection()
            if not selection:
                return
            i = selection[0]
            j = max(0, min(box.size() - 1, i + delta))
            value = box.get(i)
            box.delete(i)
            box.insert(j, value)
            box.selection_set(j)

        def save():
            out = filedialog.asksaveasfilename(parent=win, defaultextension=".docx")
            if out:
                try:
                    from .documents import merge_chapters

                    merge_chapters(box.get(0, "end"), out)
                    self.source.set(out)
                    win.destroy()
                    self.status.set("合并完成，请检查分节和引用。")
                except Exception as exc:
                    self.error(exc)

        for label, action in [
            ("上移", lambda: move(-1)),
            ("下移", lambda: move(1)),
            ("合并为新文件", save),
        ]:
            ttk.Button(win, text=label, command=action).pack(
                side="left", padx=8, pady=8
            )

    def split(self):
        folder = filedialog.askdirectory(parent=self.root)
        if folder:
            try:
                from .documents import split_chapters

                paths = split_chapters(self.source.get(), folder)
                self.status.set(
                    f"已拆分 {len(paths)} 个副本；跨章引用保留原域，须在完整论文中更新。"
                )
            except Exception as exc:
                self.error(exc)

    def launch_host(self, host):
        from .plugin import launch

        launch(["--office", host])

    def select_asset(self, key, pattern):
        path = filedialog.askopenfilename(
            parent=self.root, filetypes=[("资料文件", pattern)]
        )
        if path:
            if key in self.vars:
                self.set_value(self.vars[key], path)
            else:
                self.template[key] = path
            self.status.set("已选择：" + path)

    def autostart(self, enabled):
        try:
            from .plugin import set_autostart

            set_autostart(enabled)
            self.status.set("插件自动连接已" + ("启用" if enabled else "关闭"))
        except Exception as exc:
            self.error(exc)

    def error(self, exc):
        self.status.set(str(exc))
        messagebox.showerror("学研排版", str(exc), parent=self.root)

    def close(self):
        if self.worker and self.worker.is_alive():
            self.cancel.set()
            self.status.set("已请求取消，请等待当前保存或 Office 更新结束。")
            return
        self.root.destroy()


def main(argv=None):
    parser = argparse.ArgumentParser(description="论文工作台")
    parser.add_argument("source", nargs="?", default="")
    parser.add_argument("--host", choices=["word", "wps", "none"], default="none")
    parser.add_argument("--output-dir", default="")
    args = parser.parse_args(argv)
    from ..gui import _create_root

    root, _ = _create_root()
    AcademicWindow(root, args.source, args.host, args.output_dir)
    root.mainloop()


if __name__ == "__main__":
    main()
