"""Input formats and precompiled recognition rules."""

import re

BLANK_LINE_MODE_PRESERVE = "不改动任何空行"
BLANK_LINE_MODE_DELETE_SINGLE = "删除单个空行，多个空行保留至1个空行"
BLANK_LINE_MODE_KEEP_SINGLE = "保留单个空行，多个空行保留至1个空行"
BLANK_LINE_MODE_OPTIONS = [
    BLANK_LINE_MODE_PRESERVE,
    BLANK_LINE_MODE_DELETE_SINGLE,
    BLANK_LINE_MODE_KEEP_SINGLE,
]
DEFAULT_BLANK_LINE_MODE = BLANK_LINE_MODE_DELETE_SINGLE
SUPPORTED_FILE_EXTENSIONS = (".docx", ".doc", ".wps", ".txt", ".md")
LARGE_FOLDER_FILE_CONFIRM_THRESHOLD = 1000

RE_SAFE_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]+')
RE_HAS_CHINESE = re.compile(r"[\u4e00-\u9fff]")
RE_MD_IMAGE = re.compile(r"!\[([^\]]*)\]\([^)]+\)")
RE_MD_LINK = re.compile(r"\[([^\]]+)\]\([^)]+\)")
RE_MD_HTML_TAG = re.compile(r"<[^>]+>")
RE_MD_INLINE_CODE = re.compile(r"`([^`]+)`")
RE_MD_BOLD_ASTERISK = re.compile(r"\*\*(.*?)\*\*")
RE_MD_BOLD_UNDERSCORE = re.compile(r"__(.*?)__")
RE_MD_HEADER = re.compile(r"^\s*#+\s+(.*)")
RE_MD_BLOCKQUOTE = re.compile(r"^\s*>\s?(.*)")
RE_MD_HORIZONTAL_RULE = re.compile(r"^\s*[-*_]{3,}\s*$")
RE_MD_UNORDERED_LIST_BLOCK = re.compile(r"^\s*[*+-]\s+")
RE_MD_UNORDERED_LIST = re.compile(r"^\s*[*+-]\s")
RE_MD_BULLET_WITH_CONTENT = re.compile(r"^(\s*[*+-]\s)(.*)")
RE_MD_EMPHASIS_ASTERISK = re.compile(r"(?<!\\)\*([^\s*][^*]*?)(?<!\\)\*")
RE_CURRENCY_PREFIX = re.compile(r"^[¥￥$]")
RE_CURRENCY_SUFFIX = re.compile(r"(元|万元|亿元)$")
RE_NUMERIC_TABLE_TEXT = re.compile(r"^[-+]?(?:\d+(?:\.\d+)?|\.\d+)%?$")
CHINESE_NUM_PATTERN = r"[一二三四五六七八九十百千万零]+"
RE_TITLE_H1 = re.compile(r"^" + CHINESE_NUM_PATTERN + r"\s*、")
RE_TITLE_H2 = re.compile(r"^[（\(]" + CHINESE_NUM_PATTERN + r"[）\)]")
RE_HEADING_H1 = re.compile(r"^[一二三四五六七八九十百千万零]+\s*、")
RE_HEADING_H2 = re.compile(r"^[（\(][一二三四五六七八九十百千万零]+[）\)]")
RE_HEADING_H3 = re.compile(r"^\d+\s*[\.．]")
RE_HEADING_H4 = re.compile(r"^[（\(]\d+[）\)]")
RE_ATTACHMENT = re.compile(r"^附件\s*(\d+|[一二三四五六七八九十百千万零]+)?\s*[:：]?$")
RE_H2_INLINE_TITLE = re.compile(r"^[（\(](.+?)[）\)](.*)", re.DOTALL)
