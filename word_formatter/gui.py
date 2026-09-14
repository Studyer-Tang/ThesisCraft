# -*- coding: utf-8 -*-
"""Tkinter GUI for ThesisCraft."""

import json
import logging
import os
from pathlib import Path
import queue
import sys
import subprocess
from .conversion import IS_WINDOWS
import tempfile
import threading


def _configure_tcl_tk_paths():
    if getattr(sys, "frozen", False):
        return
    prefix = getattr(sys, "prefix", "")
    dll_dir = os.path.join(prefix, "Library", "bin")
    if os.path.isdir(dll_dir) and hasattr(os, "add_dll_directory"):
        try:
            os.add_dll_directory(dll_dir)
        except OSError:
            pass
    preferred_libraries = None
    if os.name == "nt":
        try:
            import ctypes

            library_tcl_dll = os.path.join(prefix, "Library", "bin", "tcl86t.dll")
            library_tk_dll = os.path.join(prefix, "Library", "bin", "tk86t.dll")
            dlls_tcl_dll = os.path.join(prefix, "DLLs", "tcl86t.dll")
            dlls_tk_dll = os.path.join(prefix, "DLLs", "tk86t.dll")
            if os.path.exists(library_tcl_dll):
                tcl_dll = ctypes.CDLL(library_tcl_dll)
                if os.path.exists(library_tk_dll):
                    ctypes.CDLL(library_tk_dll)
                preferred_libraries = (
                    os.path.join(prefix, "Library", "lib", "tcl8.6"),
                    os.path.join(prefix, "Library", "lib", "tk8.6"),
                )
            elif os.path.exists(dlls_tcl_dll):
                tcl_dll = ctypes.CDLL(dlls_tcl_dll)
                if os.path.exists(dlls_tk_dll):
                    ctypes.CDLL(dlls_tk_dll)
                preferred_libraries = (
                    os.path.join(prefix, "tcl", "tcl8.6"),
                    os.path.join(prefix, "tcl", "tk8.6"),
                )
            else:
                tcl_dll = None
            if tcl_dll is not None:
                tcl_dll.Tcl_FindExecutable.argtypes = [ctypes.c_char_p]
                executable = sys.executable.replace("\\", "/").encode("utf-8")
                tcl_dll.Tcl_FindExecutable(executable)
        except Exception:
            preferred_libraries = None
    library_candidates = []
    if preferred_libraries:
        library_candidates.append(preferred_libraries)
    library_candidates.extend(
        [
            (
                os.path.join(prefix, "Library", "lib", "tcl8.6"),
                os.path.join(prefix, "Library", "lib", "tk8.6"),
            ),
            (
                os.path.join(prefix, "tcl", "tcl8.6"),
                os.path.join(prefix, "tcl", "tk8.6"),
            ),
            (
                os.path.join(prefix, "Lib", "tcl8.6"),
                os.path.join(prefix, "Lib", "tk8.6"),
            ),
        ]
    )
    for tcl_library, tk_library in library_candidates:
        if os.path.exists(os.path.join(tcl_library, "init.tcl")) and os.path.exists(
            os.path.join(tk_library, "tk.tcl")
        ):
            os.environ.setdefault("TCL_LIBRARY", tcl_library)
            os.environ.setdefault("TK_LIBRARY", tk_library)
            break


_configure_tcl_tk_paths()

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext, Menu

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    TKDND_AVAILABLE = True
except Exception:
    DND_FILES = None
    TkinterDnD = None
    TKDND_AVAILABLE = False

from .config import DEFAULT_CONFIG, FONT_SIZE_MAP, PRESET_FONT_OPTIONS, normalize_config
from .jobs import build_jobs
from .service import run_jobs
from .storage import user_config_path
from .compat import (
    BLANK_LINE_MODE_DELETE_SINGLE,
    BLANK_LINE_MODE_KEEP_SINGLE,
    LARGE_FOLDER_FILE_CONFIRM_THRESHOLD,
    SUPPORTED_FILE_EXTENSIONS,
    WPSAppManager,
    WordProcessor,
    _initialize_com_for_thread,
    _uninitialize_com_for_thread,
)
from .version import APP_TITLE, __version__, AUTHOR

from .gui_view import FormatterView, UNIT_DISPLAY_TO_CONFIG


class WordFormatterGUI(FormatterView):
    def __init__(self, master):
        from .ui_common import apply_theme
        apply_theme(master)
        self.master = master
        self.dnd_available = TKDND_AVAILABLE
        self.dnd_files = DND_FILES
        master.title(f"{APP_TITLE} v{__version__} · {AUTHOR}")
        self.left_pane_min_width = 360
        self.right_pane_min_width = 520
        self.preferred_left_ratio = 0.38
        self.main_pane = None
        self._configure_window_geometry()
        self.log_queue = queue.Queue()
        self.ui_queue = queue.Queue()
        self.cancel_event = threading.Event()
        self.worker_thread = None
        self.close_pending = False
        self.extra_config = {}
        self.is_processing = False

        self.font_size_map = FONT_SIZE_MAP.copy()
        self.font_size_map_rev = {v: k for k, v in self.font_size_map.items()}
        self.default_params = DEFAULT_CONFIG.copy()
        self.font_separator = "── 已安装字体 ──"
        self.installed_fonts = self._get_installed_fonts()
        self.font_options = {
            key: self._with_installed_fonts(options)
            for key, options in PRESET_FONT_OPTIONS.items()
        }
        self.set_outline_var = tk.BooleanVar(value=self.default_params["set_outline"])
        self.enable_attachment_var = tk.BooleanVar(
            value=self.default_params["enable_attachment_formatting"]
        )
        self.force_a4_var = tk.BooleanVar(value=self.default_params["force_a4"])
        self.use_custom_english_font_var = tk.BooleanVar(
            value=self.default_params["use_custom_english_font"]
        )
        self.normalize_punctuation_var = tk.BooleanVar(
            value=self.default_params["normalize_punctuation"]
        )
        self.enable_table_var = tk.BooleanVar(
            value=self.default_params["enable_table_formatting"]
        )
        self.table_auto_col_width_var = tk.BooleanVar(
            value=self.default_params["table_auto_col_width"]
        )
        self.table_header_bold_var = tk.BooleanVar(
            value=self.default_params["table_header_bold"]
        )
        self.table_smart_align_var = tk.BooleanVar(
            value=self.default_params["table_smart_align"]
        )
        self.table_unified_borders_var = tk.BooleanVar(
            value=self.default_params["table_unified_borders"]
        )
        self.title_bold_var = tk.BooleanVar(
            value=self.default_params.get("title_bold", False)
        )
        self.h1_bold_var = tk.BooleanVar(
            value=self.default_params.get("h1_bold", False)
        )
        self.h2_bold_var = tk.BooleanVar(
            value=self.default_params.get("h2_bold", False)
        )
        self.progress_var = tk.DoubleVar(value=0.0)
        self.progress_text_var = tk.StringVar(value="")
        self.entries = {}
        self.attachment_option_widgets = []
        self.table_option_widgets = []
        self.spacing_controls = []
        self.indent_controls = []
        self._active_config_canvas = None

        self.default_config_path = str(user_config_path())

        self.create_menu()
        self.create_widgets()
        self.load_initial_config()

        self.master.protocol("WM_DELETE_WINDOW", self._on_close)
        self.master.after(250, self.set_initial_pane_position)
        self.master.after(100, self._check_log_queue)

    def _configure_window_geometry(self):
        screen_width = max(self.master.winfo_screenwidth(), 1024)
        screen_height = max(self.master.winfo_screenheight(), 720)
        width = min(1200, max(900, screen_width - 80))
        height = min(860, max(640, screen_height - 100))
        min_width = min(1000, max(860, screen_width - 120))
        min_height = min(720, max(600, screen_height - 160))
        x = max(0, (screen_width - width) // 2)
        y = max(0, (screen_height - height) // 2)
        self.master.geometry(f"{width}x{height}+{x}+{y}")
        self.master.minsize(min_width, min_height)

    def _clamped_left_width(self, total_width, preferred=None):
        total_width = max(int(total_width), 1)
        min_left = self.left_pane_min_width
        min_right = self.right_pane_min_width
        if total_width < min_left + min_right:
            min_left = max(280, int(total_width * 0.42))
            min_right = max(320, total_width - min_left)
        preferred = int(
            preferred
            if preferred is not None
            else total_width * self.preferred_left_ratio
        )
        return max(min_left, min(preferred, max(min_left, total_width - min_right)))

    def _ensure_pane_widths(self, event=None):
        pane = self.main_pane
        if pane is None:
            return
        try:
            total_width = pane.winfo_width()
            if total_width <= 100:
                return
            current_left = pane.sashpos(0)
            desired_left = self._clamped_left_width(total_width, preferred=current_left)
            if desired_left != current_left:
                pane.sashpos(0, desired_left)
        except tk.TclError:
            pass

    def set_initial_pane_position(self):
        pane = self.main_pane
        if pane is None:
            return
        try:
            total_width = pane.winfo_width()
            if total_width > 100:
                left_width = self._clamped_left_width(total_width)
                pane.sashpos(0, left_width)
        except tk.TclError:
            pass

    def create_menu(self):
        menubar = Menu(self.master)
        menubar.add_command(label="论文工作台", command=self._launch_academic)
        plugin_menu = Menu(menubar, tearoff=0)
        plugin_menu.add_command(
            label="启动 Word 插件", command=lambda: self._launch_office_plugin("word")
        )
        plugin_menu.add_command(
            label="启动 WPS 插件", command=lambda: self._launch_office_plugin("wps")
        )
        menubar.add_cascade(label="插件", menu=plugin_menu)
        help_menu = Menu(menubar, tearoff=0)
        help_menu.add_command(label="使用说明", command=self.show_help_window)
        help_menu.add_command(label="重置界面布局", command=self.reset_layout)
        menubar.add_cascade(label="帮助", menu=help_menu)
        self.master.config(menu=menubar)

    def _launch_academic(self):
        from .academic.plugin import launch
        launch(['--academic'])

    def _launch_office_plugin(self, host):
        if not IS_WINDOWS:
            messagebox.showinfo(
                "插件", "Word/WPS 工具栏插件需要 Windows 桌面版。", parent=self.master
            )
            return
        command = (
            [sys.executable, "--office", host]
            if getattr(sys, "frozen", False)
            else [sys.executable, "-m", "word_formatter.office_toolbar", host]
        )
        subprocess.Popen(command, creationflags=subprocess.CREATE_NO_WINDOW)

    def reset_layout(self):
        self.set_initial_pane_position()
        self.log_to_debug_window("已重置界面布局。")

    def _append_log_message(self, message):
        try:
            self.debug_text.config(state="normal")
            self.debug_text.insert(tk.END, message + "\n")
            self.debug_text.config(state="disabled")
            self.debug_text.see(tk.END)
        except tk.TclError:
            pass

    def _check_log_queue(self):
        try:
            while True:
                callback, args = self.ui_queue.get_nowait()
                callback(*args)
        except queue.Empty:
            pass
        self._drain_log_queue()
        try:
            self.master.after(100, self._check_log_queue)
        except tk.TclError:
            pass

    def log_to_debug_window(self, message):
        self.log_queue.put(message)

    def _drain_log_queue(self):
        try:
            while True:
                self._append_log_message(self.log_queue.get_nowait())
        except queue.Empty:
            pass

    def _clear_debug_log(self):
        self._drain_log_queue()
        try:
            self.debug_text.config(state="normal")
            self.debug_text.delete("1.0", tk.END)
            self.debug_text.config(state="disabled")
        except tk.TclError:
            pass

    def _run_on_main(self, callback, *args):
        self.ui_queue.put((callback, args))

    def _set_progress(self, value, text=""):
        def update():
            try:
                self.progress_var.set(value)
                self.progress_text_var.set(text)
            except tk.TclError:
                pass

        self._run_on_main(update)

    def load_initial_config(self):
        if not os.path.exists(self.default_config_path) and os.path.exists(
            "default_config.json"
        ):
            self.default_config_path = os.path.abspath("default_config.json")
        if os.path.exists(self.default_config_path):
            try:
                with open(self.default_config_path, "r", encoding="utf-8") as f:
                    config = json.load(f)
                self._apply_config(config)
                self.log_to_debug_window(
                    f"已加载默认配置文件: {self.default_config_path}"
                )
            except Exception as e:
                self.log_to_debug_window(
                    f"加载默认配置 '{self.default_config_path}' 失败: {e}。将使用内置默认值。"
                )
                self.load_defaults()
        else:
            self.log_to_debug_window("未找到默认配置文件，将使用内置默认值。")
            self.load_defaults()

    @staticmethod
    def _legacy_blank_line_mode(remove_blank_lines):
        return (
            BLANK_LINE_MODE_DELETE_SINGLE
            if remove_blank_lines
            else BLANK_LINE_MODE_KEEP_SINGLE
        )

    def _apply_config(self, loaded_config):
        loaded_config = normalize_config(loaded_config)
        self.extra_config = loaded_config.copy()
        for key, variable in self.safety_vars.items():
            variable.set(loaded_config[key])
        self.set_outline_var.set(loaded_config.get("set_outline", True))
        self.enable_attachment_var.set(
            loaded_config.get("enable_attachment_formatting", True)
        )
        self.force_a4_var.set(loaded_config.get("force_a4", False))
        self.use_custom_english_font_var.set(
            loaded_config.get("use_custom_english_font", False)
        )
        self.normalize_punctuation_var.set(
            loaded_config.get("normalize_punctuation", False)
        )
        self.enable_table_var.set(loaded_config.get("enable_table_formatting", False))
        self.table_auto_col_width_var.set(
            loaded_config.get("table_auto_col_width", True)
        )
        self.table_header_bold_var.set(loaded_config.get("table_header_bold", True))
        self.table_smart_align_var.set(loaded_config.get("table_smart_align", False))
        self.table_unified_borders_var.set(
            loaded_config.get("table_unified_borders", True)
        )
        self.title_bold_var.set(loaded_config.get("title_bold", False))
        self.h1_bold_var.set(loaded_config.get("h1_bold", False))
        self.h2_bold_var.set(loaded_config.get("h2_bold", False))
        boolean_keys = [
            "set_outline",
            "enable_attachment_formatting",
            "force_a4",
            "use_custom_english_font",
            "use_times_new_roman",
            "remove_blank_lines",
            "normalize_punctuation",
            "enable_table_formatting",
            "table_auto_col_width",
            "table_header_bold",
            "table_smart_align",
            "table_unified_borders",
            "title_bold",
            "h1_bold",
            "h2_bold",
        ]
        self._enable_dependent_widgets_for_config_load()
        for key, value in loaded_config.items():
            if key in boolean_keys:
                continue
            widget = self.entries.get(key)
            if widget:
                self._set_widget_value(widget, value, is_size=("_size" in key))
        self._update_english_font_state()
        self._update_attachment_state()
        self._update_table_state()
        self._refresh_spacing_controls()
        self._refresh_indent_controls()

    def load_defaults(self):
        self._apply_config(self.default_params)

    def collect_config(self):
        config = self.extra_config.copy()
        for key, widget in self.entries.items():
            value = widget.get().strip()
            if isinstance(widget, ttk.Combobox) and value == self.font_separator:
                value = getattr(widget, "_last_valid_value", "").strip()
            if isinstance(widget, ttk.Combobox) and getattr(
                widget, "_is_unit_combo", False
            ):
                value = UNIT_DISPLAY_TO_CONFIG.get(value, value)
            if value == "" and key in self.default_params:
                config[key] = self.default_params[key]
                continue
            if "_size" in key and isinstance(widget, ttk.Combobox):
                if value in self.font_size_map:
                    config[key] = self.font_size_map[value]
                else:
                    try:
                        config[key] = float(value)
                    except (ValueError, TypeError):
                        self.log_to_debug_window(
                            f"警告: 无效的字号值 '{value}' for '{key}'. 使用默认值 16pt。"
                        )
                        config[key] = 16
            else:
                try:
                    config[key] = float(value) if "." in value else int(value)
                except (ValueError, TypeError):
                    config[key] = value
        config["set_outline"] = self.set_outline_var.get()
        config["enable_attachment_formatting"] = self.enable_attachment_var.get()
        config["force_a4"] = self.force_a4_var.get()
        config["use_custom_english_font"] = self.use_custom_english_font_var.get()
        config["normalize_punctuation"] = self.normalize_punctuation_var.get()
        config["enable_table_formatting"] = self.enable_table_var.get()
        config["table_auto_col_width"] = self.table_auto_col_width_var.get()
        config["table_header_bold"] = self.table_header_bold_var.get()
        config["table_smart_align"] = self.table_smart_align_var.get()
        config["table_unified_borders"] = self.table_unified_borders_var.get()
        config["title_bold"] = self.title_bold_var.get()
        config["h1_bold"] = self.h1_bold_var.get()
        config["h2_bold"] = self.h2_bold_var.get()
        config.update(
            {key: variable.get() for key, variable in self.safety_vars.items()}
        )
        return normalize_config(config)

    def save_config(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".json", filetypes=[("JSON files", "*.json")]
        )
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                json.dump(self.collect_config(), f, ensure_ascii=False, indent=4)
            messagebox.showinfo("成功", f"配置已保存至 {file_path}")

    def save_default_config(self):
        try:
            self.default_config_path = str(user_config_path())
            os.makedirs(os.path.dirname(self.default_config_path), exist_ok=True)
            with open(self.default_config_path, "w", encoding="utf-8") as f:
                json.dump(self.collect_config(), f, ensure_ascii=False, indent=4)
            messagebox.showinfo(
                "成功", "当前配置已保存为默认配置。\n下次启动软件时将自动加载。"
            )
        except Exception as e:
            messagebox.showerror("错误", f"保存默认配置失败: {e}")

    def load_config(self):
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    loaded_config = json.load(f)
                self._apply_config(loaded_config)
                messagebox.showinfo("成功", "配置已加载")
            except Exception as e:
                messagebox.showerror("错误", f"加载配置文件失败: {e}")

    def _update_listbox_placeholder(self):
        if self.file_listbox.size() == 0:
            self.placeholder_label.place(
                in_=self.file_listbox, relx=0.5, rely=0.5, anchor=tk.CENTER
            )
        else:
            self.placeholder_label.place_forget()

    def handle_drop(self, event):
        paths = self.master.tk.splitlist(event.data)
        self._add_paths_to_listbox(paths)

    def _should_scan_folder(self, folder_path):
        file_count = 0
        for _, _, files in os.walk(folder_path):
            file_count += len(files)
            if file_count > LARGE_FOLDER_FILE_CONFIRM_THRESHOLD:
                folder_name = (
                    os.path.basename(os.path.normpath(folder_path)) or folder_path
                )
                return messagebox.askyesno(
                    "确认",
                    f"文件夹“{folder_name}”包含超过 {LARGE_FOLDER_FILE_CONFIRM_THRESHOLD} 个文件，继续扫描可能需要较长时间。\n\n确定继续扫描吗？",
                    parent=self.master,
                )
        return True

    def _add_paths_to_listbox(self, paths):
        current_files = set(self.file_listbox.get(0, tk.END))
        added_count = 0
        skipped_dirs = 0

        for path in paths:
            if os.path.isdir(path):
                if not self._should_scan_folder(path):
                    skipped_dirs += 1
                    self.log_to_debug_window(f"已跳过文件夹: {path}")
                    continue

                for root, _, files in os.walk(path):
                    for f in files:
                        if f.lower().endswith(SUPPORTED_FILE_EXTENSIONS):
                            full_path = os.path.join(root, f)
                            if full_path not in current_files:
                                self.file_listbox.insert(tk.END, full_path)
                                current_files.add(full_path)
                                added_count += 1
            elif os.path.isfile(path):
                if path.lower().endswith(SUPPORTED_FILE_EXTENSIONS):
                    if path not in current_files:
                        self.file_listbox.insert(tk.END, path)
                        current_files.add(path)
                        added_count += 1

        if added_count > 0:
            self.log_to_debug_window(f"通过按钮或拖拽添加了 {added_count} 个新文件。")
        if skipped_dirs > 0:
            self.log_to_debug_window(f"已跳过 {skipped_dirs} 个大文件夹。")

        self._update_listbox_placeholder()

    def add_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[
                ("所有支持的文件", "*.docx;*.doc;*.wps;*.txt;*.md"),
                ("Word 文档", "*.docx;*.doc"),
                ("WPS 文档", "*.wps"),
                ("纯文本", "*.txt"),
                ("Markdown", "*.md"),
            ]
        )
        if files:
            self._add_paths_to_listbox(files)

    def add_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self._add_paths_to_listbox([folder])

    def remove_files(self):
        selected_indices = self.file_listbox.curselection()
        if not selected_indices:
            messagebox.showinfo("提示", "请先在列表中选择要移除的文件。")
            return
        for index in sorted(selected_indices, reverse=True):
            self.file_listbox.delete(index)
        self._update_listbox_placeholder()

    def clear_list(self):
        self.file_listbox.delete(0, tk.END)
        self._update_listbox_placeholder()

    def show_help_window(self):
        help_win = tk.Toplevel(self.master)
        help_win.title("使用说明")
        help_win.geometry("600x600")
        help_text_widget = scrolledtext.ScrolledText(
            help_win, wrap=tk.WORD, state="disabled"
        )
        help_text_widget.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        help_content = f"""
{APP_TITLE} v{__version__} - 使用说明

本工具旨在提供一键式的专业文档排版体验，支持批量处理和高度自定义。

【核心功能模式】
1. 文件批量处理：可拖拽或添加 .docx, .doc, .wps, .txt, .md 文件。
2. 直接输入文本：直接粘贴文本进行排版（自动强制使用A4纸张）。

【操作流程】
1. 选择模式并添加内容。
2. （可选）在"参数设置"区调整格式，可点击各分区旁的 (?) 图标查看具体识别规则。
3. 点击"开始排版"，并选择输出位置。

【智能识别规则详解】
- 主标题与副标题:
  • 主标题: 识别文档开头的连续【居中】且【字体字号相同】的段落。
  • 副标题: 主标题下方，同样【居中】但【字体字号与主标题不同】的段落。
  • TXT/MD文件: 会将首个非层级标题的段落视为题目。

- 正文与层级标题:
  • 一级标题: “一、”, “二、” ...
  • 二级标题: “（一）”, “（二）” ...
  • 三级标题: “1.”, “2.” ...
  • 四级标题: “(1)”, “(2)” ...
  • 注：正文、三级、四级标题默认共用一套字体字号。

- 其他元素:
  • 图/表标题: 自动查找图片或表格【上方或下方】最近的、居中的、以“图”或“表”开头的段落。
  • 附件标识: 识别“附件1”、“附件：”等独立段落。启用附件格式化后，将自动【段前分页】并按主副标题规则识别其自身标题。
  • 表格内容: 默认不启用表格自动调整。启用后可分别设置表头/内容字体，并统一字号、行距、行高、列宽、边框，可选智能对齐。

【其他特性】
- 纸张设置：直接输入文本默认使用A4纸。文件处理默认保持原样，可勾选“强制设置为A4纸张”进行修改。
- 保留原文格式：统一格式时，会保留【加粗、斜体、下划线、字体颜色】等。
- 二级标题智能拆分：若二级标题后紧跟正文（如"（一）标题。正文..."），会自动在【同一个段落内】为标题和正文应用不同格式。
- 豁免内容：图片、嵌入对象等内容会自动跳过格式化；表格仅在勾选“启用表格自动调整”后处理。
- 参数自定义：所有核心参数均可在界面调整。配置方案可【保存】和【加载】。
- Markdown 支持：.md 文件会自动清理 Markdown 标记（标题#、粗体**、链接[]()、图片![]()等）后转为纯文本进行排版。
- 空行处理：TXT/MD 支持三种模式：不改动任何空行；删除单个空行且多个空行保留至1个空行；保留单个空行且多个空行保留至1个空行。默认使用“删除单个空行，多个空行保留至1个空行”。
- 跨平台说明：Windows 下优先使用 WPS/Word 处理 .doc/.wps 和自动编号转文本；macOS/Kylin/Linux 下不调用 WPS/Word，不执行自动编号转文本。.doc/.wps 会尝试使用 LibreOffice 转换，未安装 LibreOffice 时会跳过这些旧格式文件，.docx/.txt/.md 仍可正常处理。

【安全提示】
本工具【绝对不会】修改您的任何原始文件。所有操作都在后台的临时副本上进行，确保源文件100%安全。
"""
        help_text_widget.config(state="normal")
        help_text_widget.insert("1.0", help_content.strip())
        help_text_widget.config(state="disabled")

    def start_processing(self):
        if self.is_processing:
            messagebox.showinfo("提示", "正在处理中，请稍候...", parent=self.master)
            return

        active_tab_index = self.notebook.index(self.notebook.select())
        try:
            collected_config = self.collect_config()
        except ValueError as exc:
            messagebox.showerror("配置错误", str(exc), parent=self.master)
            return
        file_list = []
        text_content = ""
        output_dir = None
        output_path = None

        if active_tab_index == 0:
            file_list = list(self.file_listbox.get(0, tk.END))
            if not file_list:
                messagebox.showwarning(
                    "警告", "文件列表为空，请先添加文件！", parent=self.master
                )
                return
            try:
                output_dir = self.output.get()
            except ValueError as exc:
                messagebox.showerror("保存位置不可用", str(exc), parent=self.master)
                return
        elif active_tab_index == 1:
            text_content = self.direct_text_input.get("1.0", tk.END).strip()
            if not text_content:
                messagebox.showwarning("警告", "文本框内容为空！", parent=self.master)
                return
            try:
                output_dir = self.output.get()
            except ValueError as exc:
                messagebox.showerror("保存位置不可用", str(exc), parent=self.master)
                return
            output_path = filedialog.asksaveasfilename(
                parent=self.master,
                initialdir=output_dir or str(Path.home()),
                defaultextension=".docx",
                filetypes=[("Word Document", "*.docx")],
                initialfile="formatted_document.docx",
            )
            if not output_path:
                return

        self._clear_debug_log()
        self.cancel_event.clear()
        self.is_processing = True
        self.output.set_busy(True)
        self.start_btn.config(state="disabled", text="排版中，请稍候...")
        self._set_progress(0, "开始处理...")

        def worker():
            com_initialized = _initialize_com_for_thread(self.log_to_debug_window)
            try:
                with WPSAppManager(self.log_to_debug_window) as com_mgr:
                    processor = WordProcessor(
                        collected_config, self.log_to_debug_window, com_manager=com_mgr
                    )
                    if active_tab_index == 0:
                        self._process_files(processor, file_list, output_dir)
                    else:
                        self._process_text(processor, text_content, output_path)
            except Exception as e:
                logging.error(f"处理过程中发生严重错误: {e}", exc_info=True)
                self.log_to_debug_window(f"\n❌ 处理过程中发生严重错误：\n{e}")
                self._set_progress(100, "处理失败")

                def show_error(err=e):
                    try:
                        messagebox.showerror(
                            "错误", f"处理过程中发生错误：\n{err}", parent=self.master
                        )
                    except tk.TclError:
                        pass

                self._run_on_main(show_error)
            finally:
                _uninitialize_com_for_thread(com_initialized, self.log_to_debug_window)
                self._run_on_main(self._restore_after_processing)

        self.worker_thread = threading.Thread(target=worker, daemon=False)
        self.worker_thread.start()

    def _process_files(self, processor, file_list, output_dir):
        jobs = build_jobs(file_list, output_dir, beside_sources=output_dir is None)

        def completed(index, total, job, state, detail):
            self._set_progress(
                index / total * 100, f"{index}/{total}: {job.source.name}"
            )
            self.log_to_debug_window(f"{state}: {job.source.name}: {detail}")

        result = run_jobs(
            processor, jobs, cancel_event=self.cancel_event, on_result=completed
        )
        summary = f"成功 {len(result.outputs)} / 跳过 {len(result.skipped)} / 失败 {len(result.failures)}"
        if result.cancelled:
            summary += "（已停止）"
        self._set_progress(100, summary)
        self.log_to_debug_window(summary)
        if not self.close_pending:
            self._run_on_main(
                lambda: messagebox.showinfo("处理结果", summary, parent=self.master)
            )

    def _process_text(self, processor, text_content, output_path):
        self._set_progress(20, "处理文本...")
        temp_file_path = None
        try:
            fd, temp_file_path = tempfile.mkstemp(suffix=".txt", text=True)
            with os.fdopen(fd, "w", encoding="utf-8") as tmp:
                tmp.write(text_content)

            self.log_to_debug_window("\n--- 开始处理输入的文本 ---")
            processor.format_document(temp_file_path, output_path, overwrite=True)
            self._set_progress(100, "完成")
            self.log_to_debug_window("\n🎉 排版全部完成！")

            def show_done(path=output_path):
                try:
                    messagebox.showinfo(
                        "完成",
                        f"文档排版成功！\n文件已保存至：\n{path}",
                        parent=self.master,
                    )
                except tk.TclError:
                    pass

            self._run_on_main(show_done)
        finally:
            processor._cleanup_temp_files()
            if temp_file_path and os.path.exists(temp_file_path):
                try:
                    os.remove(temp_file_path)
                    self.log_to_debug_window("  > 输入文本的临时文件已删除")
                except OSError:
                    pass

    def _restore_after_processing(self):
        self.output.set_busy(False)
        self.is_processing = False
        if self.close_pending:
            self.master.destroy()
            return
        try:
            self.start_btn.config(state="normal", text="开始排版")
        except tk.TclError:
            pass

    def _on_close(self):
        if self.is_processing:
            if messagebox.askyesno(
                "结束任务", "完成当前文件后停止并退出？", parent=self.master
            ):
                self.close_pending = True
                self.cancel_event.set()
                self.log_to_debug_window("正在等待当前文件安全保存，随后退出。")
            return
        self._drain_log_queue()
        self.master.destroy()


def _create_root():
    if sys.platform == "darwin":
        return tk.Tk(), "macOS 下已禁用拖拽组件，请使用按钮添加文件/文件夹。"
    if TKDND_AVAILABLE:
        previous_default_root = getattr(tk, "_default_root", None)
        try:
            return TkinterDnD.Tk(), None
        except Exception as e:
            failed_root = getattr(tk, "_default_root", None)
            if failed_root is not None and failed_root is not previous_default_root:
                try:
                    failed_root.destroy()
                except Exception:
                    pass
                try:
                    tk._default_root = previous_default_root
                except Exception:
                    pass
            return tk.Tk(), f"拖拽组件启动失败，已回退为普通 Tk 窗口：{e}"
    return tk.Tk(), None


def main():
    root, dnd_error = _create_root()
    app = WordFormatterGUI(root)
    if dnd_error:
        app.log_to_debug_window(dnd_error)
    root.mainloop()


if __name__ == "__main__":
    main()
