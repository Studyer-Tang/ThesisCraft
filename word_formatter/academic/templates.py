"""Versioned, validated, editable thesis templates. Presets are not school approvals."""

from copy import deepcopy
import json
import math
from pathlib import Path

from ..storage import user_config_path

STYLE_LABELS = {
    "title": "论文题目",
    "subtitle": "英文题目",
    "h1": "章标题",
    "h2": "二级标题",
    "h3": "三级标题",
    "h4": "四级标题",
    "body": "正文",
    "abstract": "中文摘要",
    "abstract_en": "英文摘要",
    "keywords": "关键词",
    "caption_figure": "图题",
    "caption_table": "表题",
    "caption_en": "英文题注",
    "equation": "公式",
    "bibliography": "参考文献条目",
    "table": "表格正文",
    "table_header": "表头",
    "note": "图表注释",
    "footnote": "脚注",
    "endnote": "尾注",
    "appendix": "附录标题",
    "acknowledgements": "致谢正文",
    "symbols": "符号和缩略语",
    "publications": "学术成果",
    "header": "页眉",
    "footer": "页码",
    "front_heading": "前置/后置章节标题",
    "abstract_heading": "中文摘要标题",
    "abstract_en_heading": "英文摘要标题",
    "toc1": "目录一级",
    "toc2": "目录二级",
    "toc3": "目录三级",
    "cover_info": "封面信息",
}
STYLE_NAMES = {key: "PS " + key for key in STYLE_LABELS}
STYLE_NAMES.update(h1="Heading 1", h2="Heading 2", h3="Heading 3", h4="Heading 4")
STYLE_NAMES.update(toc1="TOC 1", toc2="TOC 2", toc3="TOC 3")
NUMBERING = {
    "chapter": "第1章 / 1.1",
    "chinese": "第一章 / 1.1",
    "decimal": "1 / 1.1",
    "english": "Chapter 1 / 1.1",
    "none": "保留原编号",
}
GROUPS = {
    "styles": "字体与段落",
    "pages": "分节与页眉页码",
    "numbering": "标题编号",
    "captions": "图表公式编号",
    "tables": "学术表格",
    "figures": "图片尺寸",
    "references": "引用与参考文献",
    "contents": "目录与图表目录",
}


def base_template(degree="master"):
    style = dict(
        font="宋体",
        latin="Times New Roman",
        size=12.0,
        bold=False,
        align="justify",
        spacing=1.5,
        spacing_unit="multiple",
        before=0.0,
        after=0.0,
        indent=2.0,
        color="000000",
    )
    styles = {key: dict(style) for key in STYLE_LABELS}
    for key, size in [
        ("title", 22),
        ("subtitle", 16),
        ("h1", 16),
        ("h2", 14),
        ("h3", 12),
        ("h4", 12),
        ("front_heading", 16),
        ("appendix", 16),
    ]:
        styles[key].update(
            font="黑体",
            size=float(size),
            bold=True,
            indent=0,
            before=12,
            after=12,
            align="center"
            if key in ("title", "subtitle", "h1", "front_heading", "appendix")
            else "left",
        )
    for key in (
        "caption_figure",
        "caption_table",
        "caption_en",
        "equation",
        "header",
        "footer",
    ):
        styles[key].update(indent=0, align="center", size=10.5, spacing=1.0)
    styles["equation"].update(size=12.0)
    for key in ("table", "table_header", "bibliography", "note", "footnote", "endnote"):
        styles[key].update(size=10.5, indent=0, spacing=1.0, align="left")
    styles["table_header"]["bold"] = True
    styles["abstract_heading"] = dict(styles["front_heading"])
    styles["abstract_en_heading"] = dict(styles["front_heading"])
    return dict(
        schema_version=1,
        name={"bachelor": "本科通用", "master": "硕士通用", "doctor": "博士通用"}[
            degree
        ],
        school="",
        department="",
        year="2026",
        degree=degree,
        template_version="1.0",
        sources=[],
        scope="通用预设，须按学校和院系当年规范调整。",
        styles=styles,
        page=dict(
            top=2.5,
            bottom=2.5,
            left=3.0,
            right=2.5,
            gutter=0.5,
            header_distance=1.5,
            footer_distance=1.5,
            front_number="upperRoman",
            body_number="decimal",
            number_position="outside",
            odd_even=True,
            chapter_recto=False,
            chapter_header=True,
            header_text="{school} {degree}学位论文",
            preserve_landscape=True,
            mirror_margins=False,
            header_line_pt=0.0,
        ),
        numbering=dict(
            scheme="chapter",
            font="Times New Roman",
            east_asia="宋体",
            size=12.0,
            bold=False,
            separator="-",
            by_chapter=True,
            equation_brackets="()",
            appendix_format="upperLetter",
            match_heading=True,
            match_caption=True,
        ),
        tables=dict(
            three_line=True,
            repeat_header=True,
            prevent_row_split=True,
            outer_pt=1.5,
            inner_pt=0.75,
            landscape_wide=False,
            rows_per_part=0,
        ),
        figures=dict(max_width_cm=14.5, min_dpi=150, keep_caption=True),
        contents=dict(toc=True, figures=True, tables=True, depth=3),
        references=dict(
            style="gb7714-numeric", hanging_cm=0.75, superscript=False, csl_path=""
        ),
        notes=dict(number_format="decimal", restart="continuous"),
        metadata=dict(
            title="",
            title_en="",
            author="",
            student_id="",
            supervisor="",
            date="",
            keywords="",
            keywords_en="",
        ),
        required_sections=[
            "abstract",
            "abstract_en",
            "bibliography",
            "acknowledgements",
        ],
        enabled=list(GROUPS),
        structure_overrides={},
        reference_library="",
        front_matter_path="",
        glossary_path="",
    )


def template_path():
    return user_config_path().parent / "PaperStudio" / "thesis-default.json"


def _merge(base, incoming, path=""):
    if not isinstance(incoming, dict):
        raise ValueError(f"{path or '模板'} 必须是对象。")
    for key, value in incoming.items():
        label = f"{path}.{key}" if path else key
        if key not in base:
            raise ValueError("未知模板设置：" + label)
        default = base[key]
        if isinstance(default, dict) and key != "structure_overrides":
            _merge(default, value, label)
        elif isinstance(default, bool):
            if not isinstance(value, bool):
                raise ValueError(label + " 必须为布尔值。")
            base[key] = value
        elif isinstance(default, (int, float)):
            if (
                isinstance(value, bool)
                or not isinstance(value, (int, float))
                or not math.isfinite(value)
            ):
                raise ValueError(label + " 必须为有限数字。")
            base[key] = value
        elif not isinstance(value, type(default)):
            raise ValueError(label + " 类型不正确。")
        else:
            base[key] = deepcopy(value)


def validate_template(data):
    result = base_template()
    _merge(result, data)
    if result["schema_version"] != 1:
        raise ValueError("不支持的模板 schema_version。")
    choices = {"degree": ("bachelor", "master", "doctor")}
    for key, values in choices.items():
        if result[key] not in values:
            raise ValueError(key + " 选项不正确。")
    for key, values in {
        "scheme": NUMBERING,
        "separator": ("-", ".", "–"),
        "equation_brackets": ("()", "（）", "[]"),
        "appendix_format": ("upperLetter", "decimal"),
    }.items():
        if result["numbering"][key] not in values:
            raise ValueError("编号选项不正确：" + key)
    for style in result["styles"].values():
        if style["align"] not in ("left", "center", "right", "justify") or style[
            "spacing_unit"
        ] not in ("multiple", "pt"):
            raise ValueError("段落对齐或行距单位不正确。")
        if not 5 <= style["size"] <= 96 or not 0 < style["spacing"] <= 200:
            raise ValueError("字号应在 5–96 磅，行距应大于 0 且不超过 200。")
        if any(not 0 <= style[k] <= 200 for k in ("before", "after", "indent")):
            raise ValueError("段前、段后或缩进范围不正确。")
        if not style["font"].strip() or not style["latin"].strip():
            raise ValueError("字体名称不能为空。")
        import re

        if not re.fullmatch("[0-9a-fA-F]{6}", style["color"]):
            raise ValueError("字体颜色须为六位十六进制颜色。")
    page = result["page"]
    for key in (
        "top",
        "bottom",
        "left",
        "right",
        "gutter",
        "header_distance",
        "footer_distance",
    ):
        if not 0 <= page[key] <= 10:
            raise ValueError("页面距离须在 0–10 厘米。")
    if (
        page["left"] + page["right"] + page["gutter"] >= 19
        or page["top"] + page["bottom"] >= 27
    ):
        raise ValueError("页边距过大，正文区域不足。")
    if (
        page["front_number"] not in ("upperRoman", "lowerRoman", "decimal")
        or page["body_number"] != "decimal"
    ):
        raise ValueError("页码格式不正确。")
    if page["number_position"] not in ("outside", "center"):
        raise ValueError("页码位置不正确。")
    if result["references"]["style"] not in ("gb7714-numeric", "apa", "ieee"):
        raise ValueError("参考文献样式不正确。")
    if not 0 <= result["references"]["hanging_cm"] <= 5:
        raise ValueError("文献悬挂缩进范围为 0–5 厘米。")
    if not 5 <= result["numbering"]["size"] <= 96:
        raise ValueError("编号字号范围为 5–96 磅。")
    if any(not 0 < result["tables"][k] <= 6 for k in ("outer_pt", "inner_pt")):
        raise ValueError("表格线宽范围为 0–6 磅。")
    if (
        not isinstance(result["tables"]["rows_per_part"], int)
        or not 0 <= result["tables"]["rows_per_part"] <= 1000
    ):
        raise ValueError("续表每部分行数应为 0–1000 的整数；0 表示不拆分。")
    if result["notes"]["number_format"] not in (
        "decimal",
        "decimalEnclosedCircle",
        "lowerRoman",
        "upperRoman",
        "lowerLetter",
    ) or result["notes"]["restart"] not in ("continuous", "eachPage", "eachSect"):
        raise ValueError("脚注编号或重新编号方式不正确。")
    if not 0 <= result["page"]["header_line_pt"] <= 6:
        raise ValueError("页眉线宽应为 0–6 磅。")
    if (
        not 1 <= result["figures"]["max_width_cm"] <= 40
        or not 36 <= result["figures"]["min_dpi"] <= 1200
    ):
        raise ValueError("图片宽度或清晰度阈值不正确。")
    if result["contents"]["depth"] not in (1, 2, 3, 4):
        raise ValueError("目录层级为 1–4。")
    if any(k not in GROUPS for k in result["enabled"]):
        raise ValueError("未知修复类别。")
    allowed = set(STYLE_LABELS) | {
        "ignore",
        "bibliography_heading",
        "abstract_heading",
        "abstract_en_heading",
        "acknowledgements_heading",
        "symbols_heading",
        "publications_heading",
    }
    if any(
        not str(k).isdigit() or v not in allowed
        for k, v in result["structure_overrides"].items()
    ):
        raise ValueError("段落结构标记无效。")
    if any(
        k
        not in (
            "abstract",
            "abstract_en",
            "bibliography",
            "acknowledgements",
            "symbols",
            "publications",
            "appendix",
        )
        for k in result["required_sections"]
    ):
        raise ValueError("未知必需章节。")
    return result


def pku_template(degree="master"):
    result = base_template(degree)
    result.update(
        name="北京大学" + ("博士" if degree == "doctor" else "硕士") + "（研究生指南）",
        school="北京大学",
        template_version="2014-guide+2024-cover",
        sources=[
            "https://grs.pku.edu.cn/docs/2019-05/20190524160158375113.pdf",
            "https://grs.pku.edu.cn/xwgz11/xwsy11/bsxw111/clxz09/50118yjsy346375.htm",
        ],
        scope="依据研究生院公开的2014年写作指南（2019年附件）及2024博士模板核对页边距；院系补充要求和当年正式封面需确认。本科不适用。",
    )
    result["page"].update(
        top=3.0,
        bottom=2.5,
        left=2.6,
        right=2.6,
        gutter=0,
        header_distance=2.0,
        footer_distance=1.75,
        number_position="center",
        mirror_margins=True,
        header_line_pt=0.75,
        header_text="北京大学{degree}学位论文",
    )
    result["numbering"].update(
        scheme="chinese", separator=".", equation_brackets="（）"
    )
    result["contents"].update(figures=False, tables=False)
    for key in (
        "body",
        "abstract",
        "abstract_en",
        "keywords",
        "acknowledgements",
        "symbols",
        "publications",
    ):
        result["styles"][key].update(
            size=12.0, spacing=20.0, spacing_unit="pt", before=0, after=0
        )
    for key in ("h1", "front_heading", "appendix", "abstract_heading"):
        result["styles"][key].update(
            size=16, spacing=1.0, spacing_unit="multiple", before=24, after=18
        )
    for key, size, before in [("h2", 14, 24), ("h3", 13, 12), ("h4", 12, 12)]:
        result["styles"][key].update(
            size=size, spacing=20, spacing_unit="pt", before=before, after=6
        )
    result["styles"]["title"].update(size=26)
    result["styles"]["subtitle"].update(
        latin="Arial", size=16, before=24, after=18, spacing=1.0
    )
    result["styles"]["cover_info"].update(font="仿宋", size=16, align="center")
    result["styles"]["abstract_en_heading"].update(
        latin="Arial", size=12, before=8, after=6, spacing_unit="pt", spacing=20
    )
    for key, before, after in [
        ("caption_figure", 6, 12),
        ("caption_table", 12, 6),
        ("caption_en", 6, 12),
    ]:
        result["styles"][key].update(font="宋体", size=11, before=before, after=after)
    for key in ("table", "table_header"):
        result["styles"][key].update(font="宋体", size=11, before=3, after=3)
    result["styles"]["equation"].update(size=12, before=6, after=6)
    result["styles"]["bibliography"].update(
        size=10.5, spacing=16, spacing_unit="pt", before=3, after=0
    )
    result["styles"]["footnote"].update(size=9, align="justify")
    result["notes"].update(number_format="decimalEnclosedCircle", restart="eachPage")
    for i in range(1, 4):
        result["styles"][f"toc{i}"].update(
            font="黑体" if i == 1 else "宋体",
            size=12,
            spacing=20,
            spacing_unit="pt",
            before=6 if i == 1 else 0,
            after=0,
            indent=0,
            align="left",
        )
    result["references"].update(superscript=True, hanging_cm=0.741)
    return validate_template(result)


def load_template(value=None):
    if value in ("pku-master", "pku-doctor"):
        return pku_template(value.split("-")[1])
    if value in ("bachelor", "master", "doctor"):
        return base_template(value)
    path = Path(value) if value else template_path()
    return (
        validate_template(json.loads(path.read_text(encoding="utf-8-sig")))
        if path.exists()
        else (pku_template() if value is None else _missing(path))
    )


def _missing(path):
    raise FileNotFoundError(path)


def save_template(data, path=None):
    from tempfile import NamedTemporaryFile
    import os

    path = Path(path) if path else template_path()
    data = validate_template(data)
    path.parent.mkdir(parents=True, exist_ok=True)
    with NamedTemporaryFile(
        "w", encoding="utf-8", dir=path.parent, delete=False
    ) as stream:
        json.dump(data, stream, ensure_ascii=False, indent=2)
        temp = stream.name
    try:
        os.replace(temp, path)
    finally:
        if os.path.exists(temp):
            os.unlink(temp)
    return path


def import_word_styles(path, template=None):
    """Import measurable Word style/page properties; return unresolved rules explicitly."""
    from docx import Document
    from docx.oxml.ns import qn

    doc = Document(path)
    result = deepcopy(template or base_template())
    imported = []
    for key, candidates in {
        "body": ["Normal", "正文"],
        "title": ["Title", "标题"],
        **{f"h{i}": [f"Heading {i}", f"标题 {i}"] for i in range(1, 5)},
        "caption_figure": ["Caption", "题注"],
        "footnote": ["Footnote Text"],
        "endnote": ["Endnote Text"],
    }.items():
        style = next(
            (doc.styles[name] for name in candidates if name in doc.styles), None
        )
        if style is None:
            continue
        target = result["styles"][key]
        chain, seen = [], set()
        while style is not None and style.style_id not in seen:
            seen.add(style.style_id)
            chain.append(style)
            style = style.base_style
        for style in reversed(chain):
            if style.font.name:
                target["latin"] = style.font.name
            fonts = style.element.find(".//" + qn("w:rFonts"))
            if fonts is not None and fonts.get(qn("w:eastAsia")):
                target["font"] = fonts.get(qn("w:eastAsia"))
            if style.font.size:
                target["size"] = style.font.size.pt
            if style.font.bold is not None:
                target["bold"] = style.font.bold
            fmt = style.paragraph_format
            if fmt.line_spacing is not None:
                target["spacing_unit"] = (
                    "pt" if hasattr(fmt.line_spacing, "pt") else "multiple"
                )
                target["spacing"] = (
                    fmt.line_spacing.pt
                    if hasattr(fmt.line_spacing, "pt")
                    else fmt.line_spacing
                )
            for field in ("before", "after"):
                val = getattr(fmt, "space_" + field)
                if val is not None:
                    target[field] = val.pt
        imported.append(STYLE_LABELS[key])
    section = doc.sections[0]
    for key in ("top", "bottom", "left", "right"):
        result["page"][key] = getattr(section, key + "_margin").cm
    result["name"] = Path(path).stem + "（导入待确认）"
    return validate_template(result), imported
