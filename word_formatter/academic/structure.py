"""Deterministic section classification with explicit, reviewable overrides."""

import re
from dataclasses import dataclass, asdict
from docx.oxml.ns import qn
from .xmlutil import visible_text

SECTION_NAMES = {
    "摘要": "abstract",
    "中文摘要": "abstract",
    "abstract": "abstract_en",
    "参考文献": "bibliography",
    "references": "bibliography",
    "bibliography": "bibliography",
    "致谢": "acknowledgements",
    "acknowledgements": "acknowledgements",
    "acknowledgments": "acknowledgements",
    "符号说明": "symbols",
    "符号表": "symbols",
    "主要符号对照表": "symbols",
    "缩略语表": "symbols",
    "缩略词表": "symbols",
    "攻读学位期间取得的研究成果": "publications",
    "学术成果": "publications",
    "研究成果": "publications",
}
CHAPTER = re.compile(
    r"^(?:第[一二三四五六七八九十百零〇\d]+[章节]|Chapter\s+\d+)\s*[、.．:]?\s*", re.I
)
DECIMAL = re.compile(r"^(\d+(?:[.．]\d+){0,3})[.．、]?\s+(?=\S)")
DECIMAL_CJK = re.compile(r"^(\d+(?:[.．]\d+){1,3})\s*(?=[\u4e00-\u9fffA-Za-z])")
CAPTION = re.compile(
    r"^(图|表|Figure|Fig\.|Table)\s*(\d+(?:[.\-–]\d+)*|[A-Z][.\-]\d+)\s*[：:.、]?\s*(.*)$",
    re.I,
)
MARKER = re.compile(r"^\{\{(fig|table|eq):([\w.-]+)\}\}\s*(.*)$")


@dataclass
class Block:
    index: int
    kind: str
    region: str
    text: str
    prefix: str = ""
    key: str = ""
    confidence: str = "rule"

    def to_dict(self):
        return asdict(self)


def scan(doc, overrides=None):
    overrides = overrides or {}
    region = "front"
    items = []
    first_text = True
    for index, paragraph in enumerate(doc.paragraphs):
        text = visible_text(paragraph).strip()
        compact = re.sub(r"\s+", "", text).lower()
        style = paragraph.style.name if paragraph.style is not None else ""
        kind, prefix, key, confidence = "body", "", "", "rule"
        if compact in SECTION_NAMES:
            region = SECTION_NAMES[compact]
            kind = region + "_heading"
        elif re.match(
            r"^(附录\s*[A-Z一二三四五六七八九十\d]|Appendix\s+[A-Z\d])", text, re.I
        ):
            region, kind = "appendix", "appendix"
            match = re.match(
                r"^(?:附录|Appendix)\s*[A-Z一二三四五六七八九十\d]+\s*[.、：:]?\s*",
                text,
                re.I,
            )
            prefix = match.group() if match else ""
        elif MARKER.match(text):
            marker, key, _ = MARKER.match(text).groups()
            kind = {
                "fig": "caption_figure",
                "table": "caption_table",
                "eq": "equation",
            }[marker]
        elif (
            CHAPTER.match(text)
            and len(text) < 100
            and not re.search(r"[。！？；]$", text)
        ):
            kind, prefix, region = "h1", CHAPTER.match(text).group(), "body"
        elif re.match(r"^(?:Heading|标题)\s*[1-4]$", style, re.I):
            kind, region = "h" + style[-1], "body"
            match = DECIMAL.match(text) or DECIMAL_CJK.match(text)
            prefix = match.group() if match else ""
        elif style == "PS appendix":
            kind, region = "appendix", "appendix"
        elif DECIMAL.match(text) or DECIMAL_CJK.match(text):
            match = DECIMAL.match(text) or DECIMAL_CJK.match(text)
            kind = "h" + str(match.group(1).replace("．", ".").count(".") + 1)
            prefix, region, confidence = match.group(), "body", "review"
        elif CAPTION.match(text):
            token = CAPTION.match(text).group(1).lower()
            kind = "caption_table" if token in ("表", "table") else "caption_figure"
            confidence = "review" if len(text) > 100 else "rule"
        elif style in ("PS caption_figure", "PS caption_table", "PS equation"):
            kind = style[3:]
        elif (
            next(paragraph._p.iter(qn("m:oMath")), None) is not None
            or next(paragraph._p.iter(qn("w:object")), None) is not None
        ):
            kind = "equation" if not text.strip() else "body"
        elif re.match(r"^(关键词|关键字|key\s*words)\s*[:：]", text, re.I):
            kind = "keywords"
        elif re.match(r"^(?:注|资料来源|说明|Notes?)\s*[:：]", text, re.I):
            kind = "note"
        elif compact in (
            "目录",
            "contents",
            "图目录",
            "表目录",
            "插图目录",
        ) or style.startswith("TOC"):
            kind = "ignore"
        elif first_text and region == "front" and text:
            kind = "title"
        elif region in (
            "abstract",
            "abstract_en",
            "bibliography",
            "acknowledgements",
            "symbols",
            "publications",
        ):
            kind = region
        elif region == "front":
            kind = "ignore"
        if str(index) in overrides:
            kind, confidence = overrides[str(index)], "manual"
            if kind.endswith("_heading"):
                region = kind.removesuffix("_heading")
            elif kind == "h1":
                region = "body"
        if text:
            first_text = False
        items.append(Block(index, kind, region, text, prefix, key, confidence))
    return items
