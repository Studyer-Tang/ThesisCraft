# -*- coding: utf-8 -*-
"""Shared configuration defaults for ThesisCraft."""

import math
from .constants import (
    DEFAULT_BLANK_LINE_MODE,
    BLANK_LINE_MODE_OPTIONS,
    BLANK_LINE_MODE_DELETE_SINGLE,
    BLANK_LINE_MODE_KEEP_SINGLE,
    BLANK_LINE_MODE_PRESERVE,
)

FONT_SIZE_MAP = {
    "一号 (26pt)": 26,
    "小一 (24pt)": 24,
    "二号 (22pt)": 22,
    "小二 (18pt)": 18,
    "三号 (16pt)": 16,
    "小三 (15pt)": 15,
    "四号 (14pt)": 14,
    "小四 (12pt)": 12,
    "五号 (10.5pt)": 10.5,
    "小五 (9pt)": 9,
}

DEFAULT_CONFIG = {
    "page_number_align": "奇偶分页",
    "footer_distance": 2.5,
    "line_spacing": 28,
    "margin_top": 3.7,
    "margin_bottom": 3.5,
    "margin_left": 2.8,
    "margin_right": 2.6,
    "title_font": "方正小标宋简体",
    "h1_font": "黑体",
    "h2_font": "楷体_GB2312",
    "body_font": "仿宋_GB2312",
    "page_number_font": "宋体",
    "table_caption_font": "黑体",
    "figure_caption_font": "黑体",
    "attachment_font": "黑体",
    "subtitle_font": "楷体_GB2312",
    "title_size": 22,
    "h1_size": 16,
    "h2_size": 16,
    "body_size": 16,
    "page_number_size": 14,
    "table_caption_size": 14,
    "figure_caption_size": 14,
    "attachment_size": 16,
    "subtitle_size": 16,
    "title_line_spacing": 33,
    "subtitle_line_spacing": 33,
    "line_spacing_unit": "pt",
    "line_spacing_multiple": 1.0,
    "title_line_spacing_unit": "pt",
    "title_line_spacing_multiple": 1.0,
    "subtitle_line_spacing_unit": "pt",
    "subtitle_line_spacing_multiple": 1.0,
    "left_indent_cm": 0.0,
    "right_indent_cm": 0.0,
    "paragraph_indent_unit": "cm",
    "left_indent_chars": 0.0,
    "right_indent_chars": 0.0,
    "first_line_indent_chars": 2.0,
    "title_bold": False,
    "h1_bold": False,
    "h2_bold": False,
    "set_outline": True,
    "enable_attachment_formatting": True,
    "force_a4": False,
    "use_custom_english_font": False,
    "english_font": "Times New Roman",
    "blank_line_mode": DEFAULT_BLANK_LINE_MODE,
    "normalize_punctuation": False,
    "enable_table_formatting": False,
    "table_header_font": "仿宋_GB2312",
    "table_font": "仿宋_GB2312",
    "table_size": 12,
    "table_line_spacing": 22,
    "table_line_spacing_unit": "pt",
    "table_line_spacing_multiple": 1.0,
    "table_row_height_cm": 0.7,
    "table_auto_col_width": True,
    "table_width_percent": 100,
    "table_header_bold": True,
    "table_smart_align": False,
    "table_unified_borders": True,
    "table_border_size_pt": 0.5,
}

PRESET_FONT_OPTIONS = {
    "title": ["方正小标宋简体", "方正小标宋_GBK", "华文中宋"],
    "h1": ["黑体", "方正黑体_GBK", "方正黑体简体", "华文黑体"],
    "h2": ["楷体_GB2312", "方正楷体_GBK", "楷体", "方正楷体简体", "华文楷体"],
    "body": ["仿宋_GB2312", "方正仿宋_GBK", "仿宋", "方正仿宋简体", "华文仿宋"],
    "page_number": ["宋体", "Times New Roman"],
    "table_caption": ["黑体", "宋体", "仿宋_GB2312"],
    "figure_caption": ["黑体", "宋体", "仿宋_GB2312"],
    "attachment": ["黑体", "宋体", "仿宋_GB2312"],
    "subtitle": ["楷体_GB2312", "方正楷体_GBK", "楷体", "方正楷体简体", "华文楷体"],
    "table": ["仿宋_GB2312", "宋体", "黑体", "楷体_GB2312", "方正仿宋_GBK", "仿宋"],
    "english": ["Times New Roman"],
}

DEFAULT_CONFIG.update(
    {
        "force_black_text": False,
        "preprocess_office": False,
        "page_numbers": True,
        "preserve_pagination": True,
        "table_col_min_pct": 8.0,
        "table_col_max_pct": 45.0,
        "table_short_text_len": 4,
    }
)


def normalize_config(config):
    """Migrate user fields before defaults, then validate every configuration."""
    if not isinstance(config, dict):
        raise ValueError("配置必须是 JSON 对象。")
    incoming = dict(config)
    if "blank_line_mode" not in incoming and "remove_blank_lines" in incoming:
        incoming["blank_line_mode"] = (
            BLANK_LINE_MODE_DELETE_SINGLE
            if incoming["remove_blank_lines"]
            else BLANK_LINE_MODE_KEEP_SINGLE
        )
    if "use_custom_english_font" not in incoming and "use_times_new_roman" in incoming:
        incoming["use_custom_english_font"] = incoming["use_times_new_roman"]
    incoming.pop("remove_blank_lines", None)
    incoming.pop("use_times_new_roman", None)
    unknown = set(incoming) - set(DEFAULT_CONFIG)
    if unknown:
        raise ValueError("未知配置项: " + ", ".join(sorted(unknown)))
    aliases = {
        "preserve": BLANK_LINE_MODE_PRESERVE,
        "none": BLANK_LINE_MODE_PRESERVE,
        "delete_single": BLANK_LINE_MODE_DELETE_SINGLE,
        "remove_single": BLANK_LINE_MODE_DELETE_SINGLE,
        "keep_single": BLANK_LINE_MODE_KEEP_SINGLE,
        "compress": BLANK_LINE_MODE_KEEP_SINGLE,
    }
    merged = dict(DEFAULT_CONFIG)
    for key, value in incoming.items():
        default = DEFAULT_CONFIG[key]
        if value is None or (isinstance(value, str) and not value.strip()):
            continue
        if isinstance(default, bool):
            if not isinstance(value, bool):
                raise ValueError(f"{key} 必须是 true 或 false，不能是字符串或数字。")
        elif isinstance(default, (int, float)):
            if (
                isinstance(value, bool)
                or not isinstance(value, (int, float))
                or not math.isfinite(value)
            ):
                raise ValueError(f"{key} 必须是有限数值。")
            allow_negative = key in (
                "left_indent_cm",
                "right_indent_cm",
                "left_indent_chars",
                "right_indent_chars",
            )
            positive = (
                "_size" in key
                or "line_spacing" in key
                or key
                in (
                    "table_width_percent",
                    "table_short_text_len",
                    "table_col_min_pct",
                    "table_col_max_pct",
                )
            )
            if (positive and value <= 0) or (not allow_negative and value < 0):
                raise ValueError(f"{key} 的数值范围不正确。")
            if key.endswith("_size") and value > 400:
                raise ValueError(f"{key} 不能超过 400 磅。")
        elif not isinstance(value, str):
            raise ValueError(f"{key} 必须是文本。")
        if isinstance(value, str):
            value = value.strip()
        if key == "blank_line_mode":
            value = aliases.get(value, value)
        merged[key] = value
    enums = {
        "page_number_align": ("居中", "奇偶分页"),
        "paragraph_indent_unit": ("cm", "chars"),
        "blank_line_mode": BLANK_LINE_MODE_OPTIONS,
    }
    enums.update(
        {key: ("pt", "multiple") for key in merged if key.endswith("line_spacing_unit")}
    )
    for key, choices in enums.items():
        if merged[key] not in choices:
            raise ValueError(f"{key} 必须是: {', '.join(choices)}")
    if not 0 < merged["table_width_percent"] <= 100:
        raise ValueError("table_width_percent 必须在 0 到 100 之间。")
    if not 0 < merged["table_col_min_pct"] <= merged["table_col_max_pct"] <= 100:
        raise ValueError("表格列宽上下限必须满足 0 < min <= max <= 100。")
    return merged
