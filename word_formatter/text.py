"""Text operations for the shared document engine."""

from docx.oxml.ns import qn
from .constants import (
    BLANK_LINE_MODE_KEEP_SINGLE,
    BLANK_LINE_MODE_PRESERVE,
    RE_HAS_CHINESE,
    RE_MD_BLOCKQUOTE,
    RE_MD_BOLD_ASTERISK,
    RE_MD_BOLD_UNDERSCORE,
    RE_MD_BULLET_WITH_CONTENT,
    RE_MD_EMPHASIS_ASTERISK,
    RE_MD_HEADER,
    RE_MD_HORIZONTAL_RULE,
    RE_MD_HTML_TAG,
    RE_MD_IMAGE,
    RE_MD_INLINE_CODE,
    RE_MD_LINK,
    RE_MD_UNORDERED_LIST,
)
from .ooxml import paragraph_runs, replace_text_nodes, iter_tables


class TextMixin:
    @staticmethod
    def _has_chinese(text):
        return bool(RE_HAS_CHINESE.search(text or ""))

    @staticmethod
    def _is_digit_or_latin(ch):
        if not ch:
            return False
        code = ord(ch)
        return (
            ch.isdigit()
            or ("A" <= ch <= "Z")
            or ("a" <= ch <= "z")
            or (0xFF10 <= code <= 0xFF19)
            or (0xFF21 <= code <= 0xFF3A)
            or (0xFF41 <= code <= 0xFF5A)
        )

    @staticmethod
    def _prev_non_space(text, index):
        i = index - 1
        while i >= 0 and text[i].isspace():
            i -= 1
        return text[i] if i >= 0 else ""

    @staticmethod
    def _next_non_space(text, index):
        i = index + 1
        while i < len(text) and text[i].isspace():
            i += 1
        return text[i] if i < len(text) else ""

    @classmethod
    def _is_after_digit_or_latin(cls, text, index):
        return cls._is_digit_or_latin(cls._prev_non_space(text, index))

    @classmethod
    def _is_before_digit_or_latin(cls, text, index):
        return cls._is_digit_or_latin(cls._next_non_space(text, index))

    @classmethod
    def _normalize_ellipsis(cls, text):
        dot_chars = {".", "．", "。"}
        chars = []
        i = 0
        while i < len(text):
            ch = text[i]
            if ch not in dot_chars:
                chars.append(ch)
                i += 1
                continue

            j = i + 1
            while j < len(text) and text[j] in dot_chars:
                j += 1

            if j - i >= 3 and not cls._is_after_digit_or_latin(text, i):
                chars.append("……")
            else:
                chars.append(text[i:j])
            i = j

        return "".join(chars)

    @classmethod
    def _normalize_simple_punctuation(cls, text):
        replacements = {
            ",": "，",
            ".": "。",
            "．": "。",
            ";": "；",
            ":": "：",
            "?": "？",
            "!": "！",
        }
        chars = list(text)
        for i, ch in enumerate(chars):
            if ch not in replacements:
                continue
            if ch in {".", "．", "。"}:
                prev_is_dot = i > 0 and text[i - 1] in {".", "．", "。"}
                next_is_dot = i + 1 < len(text) and text[i + 1] in {".", "．", "。"}
                if prev_is_dot or next_is_dot:
                    continue
            if cls._is_after_digit_or_latin(text, i):
                continue
            chars[i] = replacements[ch]
        return "".join(chars)

    @classmethod
    def _normalize_bracket_pairs(cls, text):
        pair_sets = [
            ({"(": "（", "（": "（"}, {")": "）", "）": "）"}),
            ({"[": "［", "［": "［"}, {"]": "］", "］": "］"}),
        ]
        result = text

        for open_chars, close_chars in pair_sets:
            chars = list(result)
            stack = []
            for i, ch in enumerate(chars):
                if ch in open_chars:
                    stack.append(i)
                elif ch in close_chars and stack:
                    open_index = stack.pop()
                    content = "".join(chars[open_index + 1 : i])
                    if not cls._has_chinese(content):
                        continue
                    if cls._is_after_digit_or_latin(result, open_index):
                        continue
                    if cls._is_before_digit_or_latin(result, i):
                        continue
                    chars[open_index] = open_chars[chars[open_index]]
                    chars[i] = close_chars[ch]
            result = "".join(chars)

        return result

    @classmethod
    def _normalize_quote_pairs(
        cls, text, quote_chars, left_quote, right_quote, skip_inner_latin=False
    ):
        chars = list(text)
        open_index = None

        for i, ch in enumerate(chars):
            if ch not in quote_chars:
                continue

            prev_is_latin = cls._is_after_digit_or_latin(text, i)
            next_is_latin = cls._is_before_digit_or_latin(text, i)
            if skip_inner_latin and prev_is_latin and next_is_latin:
                continue

            if open_index is None:
                if prev_is_latin:
                    continue
                open_index = i
                continue

            content = "".join(chars[open_index + 1 : i])
            should_normalize = (
                cls._has_chinese(content)
                and not cls._is_after_digit_or_latin(text, open_index)
                and not cls._is_before_digit_or_latin(text, i)
            )
            if should_normalize:
                chars[open_index] = left_quote
                chars[i] = right_quote
            open_index = None

        return "".join(chars)

    @classmethod
    def _normalize_symbols_in_text(cls, text):
        if not text or not cls._has_chinese(text):
            return text

        result = cls._normalize_bracket_pairs(text)
        result = cls._normalize_ellipsis(result)
        result = cls._normalize_quote_pairs(
            result,
            {'"', "“", "”", "„", "‟", "「", "」"},
            "“",
            "”",
        )
        result = cls._normalize_quote_pairs(
            result,
            {"'", "‘", "’", "‚", "‛"},
            "‘",
            "’",
            skip_inner_latin=True,
        )
        result = cls._normalize_simple_punctuation(result)
        return result

    @staticmethod
    def _redistribute_text_to_runs(runs, new_full_text):
        replace_text_nodes(runs, new_full_text)

    def _normalize_paragraph_symbols(self, para):
        if (
            self._has_field_codes(para)
            or para._p.find(".//" + qn("w:fldSimple")) is not None
        ):
            return False
        runs = list(paragraph_runs(para))
        text = "".join(
            node.text or "" for run in runs for node in run._r.findall(qn("w:t"))
        )
        normalized = self._normalize_symbols_in_text(text)
        if normalized == text:
            return False
        replace_text_nodes(runs, normalized)
        return True

    def _normalize_document_symbols(self, doc):
        changes = 0

        for para in doc.paragraphs:
            if self._normalize_paragraph_symbols(para):
                changes += 1

        for table in iter_tables(doc):
            for row in table.rows:
                for cell in row.cells:
                    for para in cell.paragraphs:
                        if self._normalize_paragraph_symbols(para):
                            changes += 1

        return changes

    @staticmethod
    def _clean_markdown(text):
        """
        Clean Markdown content to plain text:
        1. Remove images, links, HTML, inline code markers
        2. Remove bold/italic markers (*, **, __)
        3. Remove heading markers (#)
        4. Remove blockquote markers (>)
        5. Remove horizontal rules (---)
        6. Preserve original ordered-list numbering from the source text
        """
        if not text:
            return ""

        # Global inline element replacements
        # Remove images: ![alt](url) -> alt
        text = RE_MD_IMAGE.sub(r"\1", text)
        # Remove links: [text](url) -> text
        text = RE_MD_LINK.sub(r"\1", text)
        # Remove HTML tags
        text = RE_MD_HTML_TAG.sub("", text)
        # Remove inline code: `code` -> code
        text = RE_MD_INLINE_CODE.sub(r"\1", text)
        # Remove bold/italic (** or __)
        text = RE_MD_BOLD_ASTERISK.sub(r"\1", text)
        text = RE_MD_BOLD_UNDERSCORE.sub(r"\1", text)

        lines = text.split("\n")
        new_lines = []

        for line in lines:
            cleaned_line = line

            # Remove heading markers: # Title -> Title
            header_match = RE_MD_HEADER.match(cleaned_line)
            if header_match:
                cleaned_line = header_match.group(1)

            # Remove blockquote markers: > Text -> Text
            blockquote_match = RE_MD_BLOCKQUOTE.match(cleaned_line)
            if blockquote_match:
                cleaned_line = blockquote_match.group(1)

            # Remove horizontal rules: ---, ***, ___
            if RE_MD_HORIZONTAL_RULE.match(cleaned_line):
                cleaned_line = ""

            # Ordered-list numbers are intentionally kept exactly as written.

            # Remove remaining * italic markers (careful not to break list markers)
            is_bullet = RE_MD_UNORDERED_LIST.match(cleaned_line)

            def remove_stars(m):
                return m.group(1)

            if is_bullet:
                bullet_match = RE_MD_BULLET_WITH_CONTENT.match(cleaned_line)
                if bullet_match:
                    marker = bullet_match.group(1)
                    content = bullet_match.group(2)
                    content = RE_MD_EMPHASIS_ASTERISK.sub(remove_stars, content)
                    cleaned_line = marker + content
            else:
                cleaned_line = RE_MD_EMPHASIS_ASTERISK.sub(remove_stars, cleaned_line)

            new_lines.append(cleaned_line)

        return "\n".join(new_lines)

    @staticmethod
    def _remove_blank_lines_from_text(text, keep_single_blank_lines=False):
        """
        Normalize blank lines from txt/md text:
        - 2+ consecutive blank lines: merge to 1
        - Single blank line: delete by default, or keep when requested
        """
        if not text:
            return text

        lines = text.split("\n")
        result = []
        blank_count = 0

        for line in lines:
            if line.strip() == "":
                blank_count += 1
            else:
                if blank_count >= 2 or (blank_count == 1 and keep_single_blank_lines):
                    result.append("")
                result.append(line)
                blank_count = 0

        if blank_count >= 2 or (blank_count == 1 and keep_single_blank_lines):
            result.append("")

        return "\n".join(result)

    def _normalize_text_blank_lines(self, text):
        if self.blank_line_mode == BLANK_LINE_MODE_PRESERVE:
            return text
        return self._remove_blank_lines_from_text(
            text,
            keep_single_blank_lines=self.blank_line_mode == BLANK_LINE_MODE_KEEP_SINGLE,
        )

    def _log_blank_line_mode(self, source_name):
        if self.blank_line_mode == BLANK_LINE_MODE_PRESERVE:
            self._log(f"  > 未改动 {source_name} 中的空行。")
        elif self.blank_line_mode == BLANK_LINE_MODE_KEEP_SINGLE:
            self._log(
                f"  > 已保留 {source_name} 中的单个空行，并将多个空行合并为 1 个。"
            )
        else:
            self._log(
                f"  > 已删除 {source_name} 中的单个空行，并将多个空行合并为 1 个。"
            )

    @staticmethod
    def _read_text_file(path):
        """Try multiple encodings to read text file"""
        for enc in ["utf-8-sig", "gb18030"]:
            try:
                with open(path, "r", encoding=enc) as f:
                    content = f.read()
                return content
            except UnicodeDecodeError:
                continue
        raise ValueError(f"无法无损读取文本编码，请将文件另存为 UTF-8: {path}")
