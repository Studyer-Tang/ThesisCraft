"""Desktop widgets; batch work and persistence belong to the controller."""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, font as tkfont
from .constants import BLANK_LINE_MODE_OPTIONS

UNIT_DISPLAY_TO_CONFIG = {
    "磅值": "pt",
    "倍数": "multiple",
    "厘米": "cm",
    "字符": "chars",
}
UNIT_CONFIG_TO_DISPLAY = {
    "pt": "磅值",
    "multiple": "倍数",
    "cm": "厘米",
    "chars": "字符",
}
UNIT_VALUE_SUFFIX = {
    "pt": "磅",
    "multiple": "倍",
    "cm": "厘米",
    "chars": "字符",
}


class FormatterView:
    def _get_installed_fonts(self):
        try:
            fonts = tkfont.families(self.master)
        except tk.TclError:
            return []

        unique_fonts = {
            font.strip()
            for font in fonts
            if font and font.strip() and not font.strip().startswith("@")
        }
        return sorted(unique_fonts, key=str.casefold)

    def _with_installed_fonts(self, preset_fonts):
        options = []
        seen = set()
        for font in preset_fonts:
            normalized = font.casefold()
            if normalized not in seen:
                options.append(font)
                seen.add(normalized)

        installed_fonts = [
            font for font in self.installed_fonts if font.casefold() not in seen
        ]
        if installed_fonts:
            options.append(self.font_separator)
            options.extend(installed_fonts)
        return options

    def _update_english_font_state(self):
        combo = self.entries.get("english_font")
        if combo:
            combo.configure(
                state="normal" if self.use_custom_english_font_var.get() else "disabled"
            )

    def _set_widgets_enabled(self, widgets, enabled):
        for widget in widgets:
            if not hasattr(widget, "_enabled_state"):
                try:
                    enabled_state = widget.cget("state") or "normal"
                    widget._enabled_state = (
                        "normal" if enabled_state == "disabled" else enabled_state
                    )
                except tk.TclError:
                    widget._enabled_state = "normal"
            try:
                widget.configure(state=widget._enabled_state if enabled else "disabled")
            except tk.TclError:
                pass

    def _update_attachment_state(self):
        self._set_widgets_enabled(
            self.attachment_option_widgets, self.enable_attachment_var.get()
        )

    def _update_table_state(self):
        self._set_widgets_enabled(
            self.table_option_widgets, self.enable_table_var.get()
        )

    def _config_unit_value(self, widget, default):
        value = widget.get().strip()
        return UNIT_DISPLAY_TO_CONFIG.get(value, value or default)

    def _refresh_spacing_controls(self):
        for control in self.spacing_controls:
            unit = self._config_unit_value(control["unit"], "pt")
            show_multiple = unit == "multiple"
            control["suffix"].configure(
                text=UNIT_VALUE_SUFFIX.get(unit, UNIT_VALUE_SUFFIX["pt"])
            )
            if show_multiple:
                if not control["multiple"].get().strip():
                    control["multiple"].insert(0, "1.0")
                control["pt"].grid_remove()
                control["multiple"].grid()
            else:
                control["multiple"].grid_remove()
                control["pt"].grid()

    def _refresh_indent_controls(self):
        for control in self.indent_controls:
            unit = self._config_unit_value(control["unit"], "cm")
            show_chars = unit == "chars"
            for label in control["suffix_labels"]:
                label.configure(
                    text=UNIT_VALUE_SUFFIX.get(unit, UNIT_VALUE_SUFFIX["cm"])
                )
            for widget in control["cm_widgets"]:
                widget.grid_remove() if show_chars else widget.grid()
            for widget in control["char_widgets"]:
                widget.grid() if show_chars else widget.grid_remove()

    def _enable_dependent_widgets_for_config_load(self):
        self._set_widgets_enabled(self.attachment_option_widgets, True)
        self._set_widgets_enabled(self.table_option_widgets, True)
        english_font_combo = self.entries.get("english_font")
        if english_font_combo:
            english_font_combo.configure(state="normal")

    def _set_widget_value(self, widget, value, is_size=False):
        previous_state = None
        try:
            previous_state = widget.cget("state")
            if previous_state == "disabled":
                widget.configure(state=getattr(widget, "_enabled_state", "normal"))
        except tk.TclError:
            previous_state = None

        try:
            if is_size and isinstance(widget, ttk.Combobox):
                display_val = self.font_size_map_rev.get(value, str(value))
                widget.set(display_val)
            elif isinstance(widget, ttk.Combobox):
                display_val = UNIT_CONFIG_TO_DISPLAY.get(str(value).strip(), value)
                widget.set(display_val)
                if display_val != self.font_separator:
                    widget._last_valid_value = display_val
            else:
                widget.delete(0, tk.END)
                widget.insert(0, str(value))
        finally:
            if previous_state == "disabled":
                try:
                    widget.configure(state="disabled")
                except tk.TclError:
                    pass

    def _show_help_tooltip(self, title, message):
        messagebox.showinfo(title, message, parent=self.master)

    def _create_help_label(self, parent, text, row, col):
        help_label = ttk.Label(parent, text="(?)", foreground="blue", cursor="hand2")
        help_label.grid(row=row, column=col, sticky="W", padx=(0, 5))
        help_label.bind(
            "<Button-1>", lambda e: self._show_help_tooltip("识别规则说明", text)
        )

    def create_widgets(self):
        main_pane = ttk.PanedWindow(self.master, orient=tk.HORIZONTAL)
        main_pane.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.main_pane = main_pane
        main_pane.bind("<Configure>", self._ensure_pane_widths, add="+")

        left_frame = ttk.Frame(main_pane, padding=5, width=420)
        main_pane.add(left_frame, weight=2)

        notebook = ttk.Notebook(left_frame)
        notebook.pack(fill=tk.BOTH, expand=True)
        self.notebook = notebook

        file_tab = ttk.Frame(notebook)
        notebook.add(file_tab, text=" 文件批量处理 ")

        list_frame = ttk.LabelFrame(
            file_tab, text="待处理文件列表（可拖拽文件或文件夹）"
        )
        list_frame.pack(fill=tk.BOTH, expand=True, pady=5)

        v_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL)
        h_scrollbar = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL)
        self.file_listbox = tk.Listbox(
            list_frame,
            yscrollcommand=v_scrollbar.set,
            xscrollcommand=h_scrollbar.set,
            selectmode=tk.EXTENDED,
        )
        v_scrollbar.config(command=self.file_listbox.yview)
        h_scrollbar.config(command=self.file_listbox.xview)
        self.file_listbox.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")
        list_frame.rowconfigure(0, weight=1)
        list_frame.columnconfigure(0, weight=1)

        if self.dnd_available and hasattr(self.file_listbox, "drop_target_register"):
            try:
                self.file_listbox.drop_target_register(self.dnd_files)
                self.file_listbox.dnd_bind("<<Drop>>", self.handle_drop)
            except Exception as e:
                self.log_to_debug_window(
                    f"拖拽组件初始化失败，已改为按钮添加文件/文件夹：{e}"
                )
        else:
            self.log_to_debug_window("拖拽添加不可用，已改为按钮添加文件/文件夹。")
        dnd_enabled = self.dnd_available and hasattr(
            self.file_listbox, "drop_target_register"
        )
        placeholder_text = (
            "可以拖拽文件或文件夹到这里"
            if dnd_enabled
            else "请使用下方按钮添加文件或文件夹"
        )
        self.placeholder_label = ttk.Label(
            self.file_listbox, text=placeholder_text, foreground="grey"
        )

        file_button_frame = ttk.Frame(file_tab)
        file_button_frame.pack(fill=tk.X, pady=5)
        ttk.Button(file_button_frame, text="添加文件", command=self.add_files).grid(
            row=0, column=0, sticky="ew", padx=2, pady=2
        )
        ttk.Button(file_button_frame, text="添加文件夹", command=self.add_folder).grid(
            row=0, column=1, sticky="ew", padx=2, pady=2
        )
        ttk.Button(file_button_frame, text="移除文件", command=self.remove_files).grid(
            row=1, column=0, sticky="ew", padx=2, pady=2
        )
        ttk.Button(file_button_frame, text="清空列表", command=self.clear_list).grid(
            row=1, column=1, sticky="ew", padx=2, pady=2
        )
        file_button_frame.columnconfigure(0, weight=1)
        file_button_frame.columnconfigure(1, weight=1)

        text_tab = ttk.Frame(notebook)
        notebook.add(text_tab, text=" 直接输入文本 ")
        text_frame = ttk.LabelFrame(text_tab, text="在此处输入或粘贴文本")
        text_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        self.direct_text_input = scrolledtext.ScrolledText(
            text_frame, height=10, wrap=tk.WORD
        )
        self.direct_text_input.pack(fill=tk.BOTH, expand=True)

        style = ttk.Style()
        style.configure(
            "Success.TButton", font=("Helvetica", 10, "bold"), foreground="green"
        )

        left_action_frame = ttk.Frame(left_frame)
        left_action_frame.pack(fill=tk.X, pady=(5, 0))
        self.start_btn = ttk.Button(
            left_action_frame,
            text="开始排版",
            style="Success.TButton",
            command=self.start_processing,
        )
        self.start_btn.pack(fill=tk.X, ipady=8)

        progress_frame = ttk.Frame(left_frame)
        progress_frame.pack(fill=tk.X, pady=(5, 0))
        self.progressbar = ttk.Progressbar(
            progress_frame, mode="determinate", variable=self.progress_var, maximum=100
        )
        self.progressbar.pack(fill=tk.X)
        ttk.Label(
            progress_frame, textvariable=self.progress_text_var, foreground="grey"
        ).pack(anchor=tk.W)

        log_frame = ttk.LabelFrame(left_frame, text="调试日志")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))
        self.debug_text = scrolledtext.ScrolledText(
            log_frame, height=10, state="disabled", wrap=tk.WORD
        )
        self.debug_text.pack(fill=tk.BOTH, expand=True)

        right_frame = ttk.Frame(main_pane, padding=4, width=660)
        main_pane.add(right_frame, weight=4)

        config_area = ttk.Frame(right_frame)
        config_canvas = tk.Canvas(config_area, highlightthickness=0, width=620)
        config_scrollbar = ttk.Scrollbar(
            config_area, orient=tk.VERTICAL, command=config_canvas.yview
        )
        config_canvas.configure(yscrollcommand=config_scrollbar.set)
        config_frame = ttk.Frame(config_canvas, padding=(6, 4))
        config_window = config_canvas.create_window(
            (0, 0), window=config_frame, anchor="nw"
        )
        for column in (1, 3):
            config_frame.columnconfigure(column, weight=1)

        def refresh_config_scrollregion(_event=None):
            config_canvas.configure(scrollregion=config_canvas.bbox("all"))

        def fit_config_width(event):
            config_canvas.itemconfig(config_window, width=event.width)

        config_frame.bind("<Configure>", refresh_config_scrollregion)
        config_canvas.bind("<Configure>", fit_config_width)
        config_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        config_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        def event_from_right_config(widget):
            while widget is not None:
                if widget == right_frame:
                    return True
                widget = getattr(widget, "master", None)
            return False

        def on_config_mousewheel(event):
            if not event_from_right_config(event.widget):
                return
            try:
                if not config_canvas.winfo_exists():
                    return
                if getattr(event, "num", None) == 4:
                    config_canvas.yview_scroll(-3, "units")
                elif getattr(event, "num", None) == 5:
                    config_canvas.yview_scroll(3, "units")
                else:
                    config_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except tk.TclError:
                pass

        right_frame.bind_all("<MouseWheel>", on_config_mousewheel, add="+")
        right_frame.bind_all("<Button-4>", on_config_mousewheel, add="+")
        right_frame.bind_all("<Button-5>", on_config_mousewheel, add="+")

        def create_entry(parent, label, var_name, r, c, width=8):
            ttk.Label(parent, text=label).grid(
                row=r, column=c, sticky=tk.W, padx=3, pady=1
            )
            entry = ttk.Entry(parent, width=width)
            entry.grid(row=r, column=c + 1, sticky=tk.EW, padx=3, pady=1)
            self.entries[var_name] = entry
            return entry

        def create_combo(parent, label, var_name, opts, r, c, readonly=True, width=12):
            ttk.Label(parent, text=label).grid(
                row=r, column=c, sticky=tk.W, padx=3, pady=1
            )
            state = "readonly" if readonly else "normal"
            combo = ttk.Combobox(parent, values=opts, state=state, width=width)
            combo.grid(row=r, column=c + 1, sticky=tk.EW, padx=3, pady=1)
            if self.font_separator in opts:
                combo._last_valid_value = ""

                def remember_font_value(event, combo=combo):
                    current = combo.get().strip()
                    if current and current != self.font_separator:
                        combo._last_valid_value = current

                def reject_font_separator(event, combo=combo):
                    if combo.get() == self.font_separator:
                        combo.set(getattr(combo, "_last_valid_value", ""))
                    else:
                        remember_font_value(event, combo)

                combo.bind("<FocusIn>", remember_font_value, add="+")
                combo.bind("<<ComboboxSelected>>", reject_font_separator, add="+")
            self.entries[var_name] = combo
            return combo

        def create_font_size_combo(parent, label, var_name, r, c):
            ttk.Label(parent, text=label).grid(
                row=r, column=c, sticky=tk.W, padx=3, pady=1
            )
            combo = ttk.Combobox(
                parent, values=list(self.font_size_map.keys()), width=12
            )
            combo.grid(row=r, column=c + 1, sticky=tk.EW, padx=3, pady=1)
            self.entries[var_name] = combo
            return combo

        def create_unit_combo(parent, var_name, values, r, c):
            combo = ttk.Combobox(parent, values=values, state="readonly", width=6)
            combo.grid(row=r, column=c, sticky=tk.EW, padx=3, pady=1)
            combo._is_unit_combo = True
            self.entries[var_name] = combo
            return combo

        def create_section_header(parent, text, help_text, r):
            header_frame = ttk.Frame(parent)
            header_frame.grid(row=r, column=0, columnspan=4, sticky="ew", pady=(6, 1))
            ttk.Label(header_frame, text=text, font=("Helvetica", 9, "bold")).pack(
                side=tk.LEFT
            )
            if help_text:
                help_label = ttk.Label(
                    header_frame, text="(?)", foreground="blue", cursor="hand2"
                )
                help_label.pack(side=tk.LEFT, padx=(2, 0))
                help_label.bind(
                    "<Button-1>",
                    lambda e, t=text, m=help_text: self._show_help_tooltip(
                        f"{t} - 识别规则", m
                    ),
                )
            ttk.Separator(parent, orient="horizontal").grid(
                row=r + 1, column=0, columnspan=4, sticky="ew"
            )
            return r + 2

        def create_spacing_control(parent, label, pt_key, unit_key, multiple_key, r):
            ttk.Label(parent, text=label).grid(
                row=r, column=0, sticky=tk.W, padx=3, pady=1
            )
            value_frame = ttk.Frame(parent)
            value_frame.grid(row=r, column=1, sticky=tk.EW, padx=3, pady=1)
            value_frame.columnconfigure(0, weight=1)
            pt_entry = ttk.Entry(value_frame, width=8)
            multiple_entry = ttk.Entry(value_frame, width=8)
            pt_entry.grid(row=0, column=0, sticky=tk.EW, padx=(0, 3))
            multiple_entry.grid(row=0, column=0, sticky=tk.EW, padx=(0, 3))
            suffix_label = ttk.Label(value_frame, text=UNIT_VALUE_SUFFIX["pt"], width=4)
            suffix_label.grid(row=0, column=1, sticky=tk.W)
            self.entries[pt_key] = pt_entry
            self.entries[multiple_key] = multiple_entry
            ttk.Label(parent, text="单位").grid(
                row=r, column=2, sticky=tk.W, padx=3, pady=1
            )
            unit_combo = create_unit_combo(parent, unit_key, ["磅值", "倍数"], r, 3)
            control = {
                "unit": unit_combo,
                "pt": pt_entry,
                "multiple": multiple_entry,
                "suffix": suffix_label,
            }
            self.spacing_controls.append(control)
            unit_combo.bind(
                "<<ComboboxSelected>>",
                lambda _event: self._refresh_spacing_controls(),
                add="+",
            )
            return [pt_entry, multiple_entry, unit_combo]

        def create_indent_controls(parent, r):
            ttk.Label(parent, text="缩进单位").grid(
                row=r, column=0, sticky=tk.W, padx=3, pady=1
            )
            unit_combo = create_unit_combo(
                parent, "paragraph_indent_unit", ["厘米", "字符"], r, 1
            )
            unit_combo.bind(
                "<<ComboboxSelected>>",
                lambda _event: self._refresh_indent_controls(),
                add="+",
            )
            r += 1
            ttk.Label(parent, text="左缩进").grid(
                row=r, column=0, sticky=tk.W, padx=3, pady=1
            )
            left_frame = ttk.Frame(parent)
            left_frame.grid(row=r, column=1, sticky=tk.EW, padx=3, pady=1)
            left_frame.columnconfigure(0, weight=1)
            left_cm = ttk.Entry(left_frame, width=8)
            left_chars = ttk.Entry(left_frame, width=8)
            left_cm.grid(row=0, column=0, sticky=tk.EW, padx=(0, 3))
            left_chars.grid(row=0, column=0, sticky=tk.EW, padx=(0, 3))
            left_suffix = ttk.Label(left_frame, text=UNIT_VALUE_SUFFIX["cm"], width=4)
            left_suffix.grid(row=0, column=1, sticky=tk.W)
            self.entries["left_indent_cm"] = left_cm
            self.entries["left_indent_chars"] = left_chars
            ttk.Label(parent, text="右缩进").grid(
                row=r, column=2, sticky=tk.W, padx=3, pady=1
            )
            right_frame_inner = ttk.Frame(parent)
            right_frame_inner.grid(row=r, column=3, sticky=tk.EW, padx=3, pady=1)
            right_frame_inner.columnconfigure(0, weight=1)
            right_cm = ttk.Entry(right_frame_inner, width=8)
            right_chars = ttk.Entry(right_frame_inner, width=8)
            right_cm.grid(row=0, column=0, sticky=tk.EW, padx=(0, 3))
            right_chars.grid(row=0, column=0, sticky=tk.EW, padx=(0, 3))
            right_suffix = ttk.Label(
                right_frame_inner, text=UNIT_VALUE_SUFFIX["cm"], width=4
            )
            right_suffix.grid(row=0, column=1, sticky=tk.W)
            self.entries["right_indent_cm"] = right_cm
            self.entries["right_indent_chars"] = right_chars
            r += 1
            ttk.Label(parent, text="首行缩进").grid(
                row=r, column=0, sticky=tk.W, padx=3, pady=1
            )
            first_line_frame = ttk.Frame(parent)
            first_line_frame.grid(row=r, column=1, sticky=tk.EW, padx=3, pady=1)
            first_line_frame.columnconfigure(0, weight=1)
            first_line_chars = ttk.Entry(first_line_frame, width=8)
            first_line_chars.grid(row=0, column=0, sticky=tk.EW, padx=(0, 3))
            ttk.Label(first_line_frame, text=UNIT_VALUE_SUFFIX["chars"], width=4).grid(
                row=0, column=1, sticky=tk.W
            )
            self.entries["first_line_indent_chars"] = first_line_chars
            self.indent_controls.append(
                {
                    "unit": unit_combo,
                    "cm_widgets": [left_cm, right_cm],
                    "char_widgets": [left_chars, right_chars],
                    "suffix_labels": [left_suffix, right_suffix],
                }
            )
            return r + 1

        page_frame = config_frame
        title_frame = config_frame
        body_frame = config_frame
        table_frame = config_frame
        advanced_frame = config_frame

        row = create_section_header(page_frame, "页面设置", None, 0)
        create_entry(page_frame, "上边距(cm)", "margin_top", row, 0)
        create_entry(page_frame, "下边距(cm)", "margin_bottom", row, 2)
        row += 1
        create_entry(page_frame, "左边距(cm)", "margin_left", row, 0)
        create_entry(page_frame, "右边距(cm)", "margin_right", row, 2)
        row += 1
        create_entry(page_frame, "页脚距(cm)", "footer_distance", row, 0)
        ttk.Checkbutton(
            page_frame, text="强制设置为A4纸张", variable=self.force_a4_var
        ).grid(row=row, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1)
        row += 1
        create_combo(
            page_frame, "页码对齐", "page_number_align", ["奇偶分页", "居中"], row, 0
        )
        row += 1
        create_combo(
            page_frame,
            "页码字体",
            "page_number_font",
            self.font_options["page_number"],
            row,
            0,
            readonly=False,
        )
        create_font_size_combo(page_frame, "页码字号", "page_number_size", row, 2)
        row += 1

        title_help = "• 主标题: 识别文档开头的连续【居中】且【字体字号相同】的段落。\n• 副标题: 主标题下方，同样【居中】但【字体字号与主标题不同】的段落。\n• TXT文件: 会将首个非层级标题的段落视为题目。"
        row = create_section_header(title_frame, "文章标题", title_help, row)
        create_combo(
            title_frame,
            "题目字体",
            "title_font",
            self.font_options["title"],
            row,
            0,
            readonly=False,
        )
        create_font_size_combo(title_frame, "题目字号", "title_size", row, 2)
        row += 1
        create_spacing_control(
            title_frame,
            "题目行距",
            "title_line_spacing",
            "title_line_spacing_unit",
            "title_line_spacing_multiple",
            row,
        )
        row += 1
        create_combo(
            title_frame,
            "副标题字体",
            "subtitle_font",
            self.font_options["subtitle"],
            row,
            0,
            readonly=False,
        )
        create_font_size_combo(title_frame, "副标题字号", "subtitle_size", row, 2)
        row += 1
        create_spacing_control(
            title_frame,
            "副标题行距",
            "subtitle_line_spacing",
            "subtitle_line_spacing_unit",
            "subtitle_line_spacing_multiple",
            row,
        )
        row += 1
        ttk.Checkbutton(
            title_frame, text="题目加粗", variable=self.title_bold_var
        ).grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        row += 1
        headings_help = '• 一级标题: "一、", "二、" ...\n• 二级标题: "（一）", "（二）" ...\n• 三级标题: "1.", "2." ...\n• 四级标题: "(1)", "(2)" ...\n\n注：正文、三级、四级标题共用一套字体字号。'
        row = create_section_header(title_frame, "层级标题", headings_help, row)
        create_combo(
            title_frame,
            "一级标题字体",
            "h1_font",
            self.font_options["h1"],
            row,
            0,
            readonly=False,
        )
        create_font_size_combo(title_frame, "一级标题字号", "h1_size", row, 2)
        row += 1
        create_combo(
            title_frame,
            "二级标题字体",
            "h2_font",
            self.font_options["h2"],
            row,
            0,
            readonly=False,
        )
        create_font_size_combo(title_frame, "二级标题字号", "h2_size", row, 2)
        row += 1
        ttk.Checkbutton(
            title_frame, text="一级标题加粗", variable=self.h1_bold_var
        ).grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        ttk.Checkbutton(
            title_frame, text="二级标题加粗", variable=self.h2_bold_var
        ).grid(row=row, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1)
        row += 1

        row = create_section_header(body_frame, "正文设置", None, row)
        create_combo(
            body_frame,
            "正文/三四级字体",
            "body_font",
            self.font_options["body"],
            row,
            0,
            readonly=False,
        )
        create_font_size_combo(body_frame, "正文/三四级字号", "body_size", row, 2)
        row += 1
        create_spacing_control(
            body_frame,
            "正文行距",
            "line_spacing",
            "line_spacing_unit",
            "line_spacing_multiple",
            row,
        )
        row += 1
        row = create_section_header(body_frame, "段落缩进", None, row)
        row = create_indent_controls(body_frame, row)
        row += 1

        table_help = (
            "• 默认不启用表格自动调整，启用后才会调整表头/内容字体、字号、行距、行高、列宽和边框。\n"
            "• 默认保留单元格原始对齐方式；勾选智能对齐后，表头/序号/短文本居中，数字靠右，长文本靠左。"
        )
        row = create_section_header(table_frame, "表格内容", table_help, row)
        ttk.Checkbutton(
            table_frame,
            text="启用表格自动调整",
            variable=self.enable_table_var,
            command=self._update_table_state,
        ).grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        table_auto_col_width_check = ttk.Checkbutton(
            table_frame, text="自动调整列宽", variable=self.table_auto_col_width_var
        )
        table_auto_col_width_check.grid(
            row=row, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1
        )
        row += 1
        table_unified_borders_check = ttk.Checkbutton(
            table_frame, text="统一表格边框", variable=self.table_unified_borders_var
        )
        table_unified_borders_check.grid(
            row=row, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1
        )
        table_header_bold_check = ttk.Checkbutton(
            table_frame, text="表头行加粗", variable=self.table_header_bold_var
        )
        table_header_bold_check.grid(
            row=row, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1
        )
        row += 1
        table_header_font_combo = create_combo(
            table_frame,
            "表头字体",
            "table_header_font",
            self.font_options["table"],
            row,
            0,
            readonly=False,
        )
        table_font_combo = create_combo(
            table_frame,
            "表格字体",
            "table_font",
            self.font_options["table"],
            row,
            2,
            readonly=False,
        )
        row += 1
        table_size_combo = create_font_size_combo(
            table_frame, "表格字号", "table_size", row, 0
        )
        table_smart_align_check = ttk.Checkbutton(
            table_frame, text="智能调整单元格对齐", variable=self.table_smart_align_var
        )
        table_smart_align_check.grid(
            row=row, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1
        )
        row += 1
        table_spacing_widgets = create_spacing_control(
            table_frame,
            "表格行距",
            "table_line_spacing",
            "table_line_spacing_unit",
            "table_line_spacing_multiple",
            row,
        )
        row += 1
        table_row_height_entry = create_entry(
            table_frame, "表格行高(cm)", "table_row_height_cm", row, 0
        )
        table_width_percent_entry = create_entry(
            table_frame, "表格宽度(%)", "table_width_percent", row, 2
        )
        row += 1
        table_border_size_entry = create_entry(
            table_frame, "边框粗细(pt)", "table_border_size_pt", row, 0
        )
        self.table_option_widgets = [
            table_auto_col_width_check,
            table_unified_borders_check,
            table_header_font_combo,
            table_font_combo,
            table_size_combo,
            *table_spacing_widgets,
            table_row_height_entry,
            table_width_percent_entry,
            table_border_size_entry,
            table_header_bold_check,
            table_smart_align_check,
        ]
        self._update_table_state()
        row += 1
        other_help = '• 图/表标题: 自动查找图片或表格【上方或下方】最近的、居中的、以"图"或"表"开头的段落。'
        row = create_section_header(table_frame, "图表标题", other_help, row)
        create_combo(
            table_frame,
            "表格标题字体",
            "table_caption_font",
            self.font_options["table_caption"],
            row,
            0,
            readonly=False,
        )
        create_font_size_combo(
            table_frame, "表格标题字号", "table_caption_size", row, 2
        )
        row += 1
        create_combo(
            table_frame,
            "图形标题字体",
            "figure_caption_font",
            self.font_options["figure_caption"],
            row,
            0,
            readonly=False,
        )
        create_font_size_combo(
            table_frame, "图形标题字号", "figure_caption_size", row, 2
        )
        row += 1

        attachment_help = '• 附件标识: 识别"附件1"、"附件："等独立段落。启用后将自动【段前分页】并按主副标题规则识别其自身标题。'
        row = create_section_header(advanced_frame, "附件与增强", attachment_help, row)
        ttk.Checkbutton(
            advanced_frame,
            text="启用附件格式化",
            variable=self.enable_attachment_var,
            command=self._update_attachment_state,
        ).grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        attachment_font_combo = create_combo(
            advanced_frame,
            "附件标识字体",
            "attachment_font",
            self.font_options["attachment"],
            row,
            2,
            readonly=False,
        )
        row += 1
        attachment_size_combo = create_font_size_combo(
            advanced_frame, "附件标识字号", "attachment_size", row, 0
        )
        self.attachment_option_widgets = [attachment_font_combo, attachment_size_combo]
        self._update_attachment_state()
        row += 1
        row = create_section_header(advanced_frame, "全局选项", None, row)
        ttk.Checkbutton(
            advanced_frame, text="自动设置大纲级别", variable=self.set_outline_var
        ).grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        ttk.Checkbutton(
            advanced_frame,
            text="启用符号标准化",
            variable=self.normalize_punctuation_var,
        ).grid(row=row, column=2, columnspan=2, sticky=tk.W, padx=3, pady=1)
        row += 1
        ttk.Checkbutton(
            advanced_frame,
            text="自定义数字和字母字体",
            variable=self.use_custom_english_font_var,
            command=self._update_english_font_state,
        ).grid(row=row, column=0, columnspan=2, sticky=tk.W, padx=3, pady=1)
        create_combo(
            advanced_frame,
            "数字和字母字体",
            "english_font",
            self.font_options["english"],
            row,
            2,
            readonly=False,
        )
        self._update_english_font_state()
        row += 1
        blank_line_combo = create_combo(
            advanced_frame,
            "TXT/MD空行处理",
            "blank_line_mode",
            BLANK_LINE_MODE_OPTIONS,
            row,
            0,
            width=34,
        )
        blank_line_combo.grid_configure(columnspan=3)

        self.safety_vars = {}
        for key, label in (
            ("force_black_text", "统一为黑色文字（默认保留原色）"),
            ("preprocess_office", "Office 预处理：接受修订并转为手动编号"),
            ("page_numbers", "设置页码（保留其他页脚文字）"),
            ("preserve_pagination", "保留原有分页和标题同页设置"),
        ):
            row += 1
            variable = tk.BooleanVar(value=self.default_params[key])
            self.safety_vars[key] = variable
            ttk.Checkbutton(advanced_frame, text=label, variable=variable).grid(
                row=row, column=0, columnspan=4, sticky=tk.W, padx=3, pady=1
            )

        self._refresh_spacing_controls()
        self._refresh_indent_controls()

        config_buttons = ttk.Frame(right_frame)
        config_buttons.pack(side=tk.BOTTOM, fill=tk.X, pady=(4, 0))
        ttk.Button(config_buttons, text="加载配置", command=self.load_config).grid(
            row=0, column=0, sticky="ew", padx=2, pady=2
        )
        ttk.Button(config_buttons, text="保存配置", command=self.save_config).grid(
            row=0, column=1, sticky="ew", padx=2, pady=2
        )
        ttk.Button(
            config_buttons, text="保存为默认", command=self.save_default_config
        ).grid(row=1, column=0, sticky="ew", padx=2, pady=2)
        ttk.Button(
            config_buttons, text="恢复内置默认", command=self.load_defaults
        ).grid(row=1, column=1, sticky="ew", padx=2, pady=2)
        config_buttons.columnconfigure(0, weight=1)
        config_buttons.columnconfigure(1, weight=1)
        config_area.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self._update_listbox_placeholder()
