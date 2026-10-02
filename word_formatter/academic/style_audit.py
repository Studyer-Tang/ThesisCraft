"""Conservative, script-aware checks of effective DOCX text properties."""

from lxml import etree
from docx.oxml.ns import qn

from .preflight import ALIASES


def style_chain(style):
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style.element
        style = style.base_style


def _child(parent, name):
    return parent.find(qn(name)) if parent is not None else None


class EffectiveFormatting:
    def __init__(self, doc):
        self.doc = doc
        self.styles = {style.style_id: style for style in doc.styles}
        defaults = _child(doc.styles.element, "w:docDefaults")
        self.default_rpr = _child(_child(defaults, "w:rPrDefault"), "w:rPr")
        self.default_ppr = _child(_child(defaults, "w:pPrDefault"), "w:pPr")
        self.theme = None
        part = next((part for part in doc.part.package.parts
                     if str(part.partname).startswith("/word/theme/")), None)
        if part is not None:
            try:
                self.theme = etree.fromstring(part.blob)
            except etree.XMLSyntaxError:
                pass  # An unreadable theme needs review, not a guessed typeface.

    def run_properties(self, run, paragraph):
        chain = [_child(run, "w:rPr")]
        rstyle = _child(chain[0], "w:rStyle")
        if rstyle is not None:
            ident = rstyle.get(qn("w:val"))
            style = self.styles.get(ident)
            if style is not None:
                chain.extend(_child(s, "w:rPr") for s in style_chain(style))
        chain.extend(_child(s, "w:rPr") for s in style_chain(paragraph.style))
        chain.append(self.default_rpr)
        return [p for p in chain if p is not None]

    @staticmethod
    def value(chain, tag, attr="val"):
        for props in chain:
            child = _child(props, "w:" + tag)
            if child is not None and child.get(qn("w:" + attr)) is not None:
                return child.get(qn("w:" + attr))
        return None

    @staticmethod
    def flag(chain, tag):
        for props in chain:
            child = _child(props, "w:" + tag)
            if child is not None:
                return child.get(qn("w:val"), "1") not in {"0", "false", "off"}
        return False

    def _theme_font(self, token, slot, chain):
        if self.theme is None or token not in {
            "majorAscii", "majorHAnsi", "majorEastAsia",
            "minorAscii", "minorHAnsi", "minorEastAsia",
        }:
            return None
        group = "majorFont" if token.startswith("major") else "minorFont"
        fonts = self.theme.find(".//" + qn("a:fontScheme") + "/" + qn("a:" + group))
        if fonts is None:
            return None
        node = fonts.find(qn("a:ea" if slot == "eastAsia" else "a:latin"))
        face = node.get("typeface") if node is not None else None
        if face:
            return face
        if slot == "eastAsia":
            language = (self.value(chain, "lang", "eastAsia") or "").lower()
            script = ("Hans" if language in {"zh-cn", "zh-sg"} else
                      "Hant" if language in {"zh-tw", "zh-hk", "zh-mo"} else
                      "Jpan" if language.startswith("ja") else
                      "Hang" if language.startswith("ko") else None)
            supplemental = next((n for n in fonts.findall(qn("a:font"))
                                 if n.get("script") == script), None) if script else None
            return supplemental.get("typeface") if supplemental is not None else None
        return None

    def font(self, chain, slot):
        # Theme attributes take precedence over a face in the same rFonts element.
        for props in chain:
            fonts = _child(props, "w:rFonts")
            if fonts is None:
                continue
            theme = fonts.get(qn("w:" + slot + "Theme"))
            if theme:
                return self._theme_font(theme, slot, chain)
            face = fonts.get(qn("w:" + slot))
            if face:
                if face in {"+mj-lt", "+mn-lt", "+mj-ea", "+mn-ea"}:
                    token = ("major" if face.startswith("+mj") else "minor") + (
                        "EastAsia" if face.endswith("ea") else "Ascii")
                    return self._theme_font(token, slot, chain)
                return face
        return None

    def spacing(self, paragraph):
        chain = [paragraph._p.pPr]
        chain.extend(_child(s, "w:pPr") for s in style_chain(paragraph.style))
        chain.append(self.default_ppr)
        chain = [p for p in chain if p is not None]
        line = self.value(chain, "spacing", "line")
        rule = self.value(chain, "spacing", "lineRule") or "auto"
        # OOXML default: a single line when no explicit line height exists.
        if line is None:
            return ("auto", 1.0) if rule == "auto" else (rule, None)
        try:
            return rule, float(line) / (240 if rule == "auto" else 20)
        except ValueError:
            return rule, None


def _same_font(first, second):
    first, second = (s.casefold().lstrip("@") for s in (first, second))
    return first == second or any({first, second} <= aliases for aliases in ALIASES)


def _scripts(text):
    slots = set()
    unsupported = False
    for char in text:
        n = ord(char)
        if (0x2E80 <= n <= 0xA4CF or 0xAC00 <= n <= 0xD7AF
                or 0xF900 <= n <= 0xFAFF or 0xFF01 <= n <= 0xFF60
                or 0x20000 <= n <= 0x323AF):
            slots.add("eastAsia")
        elif char.isascii() and char.isalnum():
            slots.add("ascii")
        elif (0x00C0 <= n <= 0x024F or 0x1E00 <= n <= 0x1EFF) and char.isalpha():
            slots.add("hAnsi")
        elif char.isalpha():
            unsupported = True
        # Neutral punctuation/symbols alone do not determine Word's font slot.
    return slots, unsupported


def text_runs(paragraph, state):
    """Include hyperlink display runs; protect fields, math and revision formatting."""
    for node in paragraph._p.iter():
        if node.tag == qn("w:fldChar"):
            flag = node.get(qn("w:fldCharType"))
            if flag == "begin":
                state["depth"] += 1
            elif flag == "end":
                state["depth"] = max(0, state["depth"] - 1)
        if node.tag != qn("w:r"):
            continue
        ancestors = list(node.iterancestors())
        nearest = next((a for a in ancestors if a.tag == qn("w:p")), None)
        if nearest is not paragraph._p:
            continue
        text = "".join(t.text or "" for t in node.findall(qn("w:t")))
        if not text.strip():
            continue
        protected = {qn("w:fldSimple"), qn("w:ins"), qn("w:del"),
                     qn("m:oMath"), qn("m:oMathPara")}
        if state["depth"] or any(a.tag in protected for a in ancestors):
            state["skipped"] = True
            continue
        # A begin marker can share its run with text, before iteration reaches it.
        if (next(node.iter(qn("w:fldChar")), None) is not None
                or next(node.iter(qn("w:instrText")), None) is not None):
            state["skipped"] = True
            continue
        yield node, text


def check_paragraph(paragraph, spec, resolver, runs, issue, index):
    mismatches = {}
    unresolved = set()
    for run, text in runs:
        chain = resolver.run_properties(run, paragraph)
        slots, unsupported = _scripts(text)
        complex_text = ((unsupported and not slots)
                        or any(resolver.flag(chain, flag) for flag in ("cs", "rtl")))
        if unsupported or complex_text:
            unresolved.add("复杂文字或数学符号的字体")
        size = resolver.value(chain, "sz")
        alignment = resolver.value(chain, "vertAlign")
        math = resolver.flag(chain, "oMath")
        if not math and not complex_text and alignment not in {"superscript", "subscript"}:
            if size is None:
                unresolved.add("未显式定义的字号")
            else:
                try:
                    pt = float(size) / 2
                except ValueError:
                    unresolved.add("无法解析的字号")
                else:
                    if abs(pt - spec["size"]) > 0.1:
                        mismatches.setdefault("font-size", set()).add(f"{pt:g} 磅")
        if math or alignment in {"superscript", "subscript"} or complex_text:
            unresolved.add("上下标、数学或复杂文字的专用格式")
            continue
        for slot in slots:
            actual = resolver.font(chain, slot)
            expected = spec["font"] if slot == "eastAsia" else spec["latin"]
            if actual is None:
                unresolved.add("主题/未显式定义的字体")
            elif not _same_font(actual, expected):
                code = "font-eastasia" if slot == "eastAsia" else "font-latin"
                mismatches.setdefault(code, set()).add(actual)
    for code, values in mismatches.items():
        label, expected = {
            "font-size": ("文本字号", f"{spec['size']:g} 磅"),
            "font-eastasia": ("中日韩文字体", spec["font"]),
            "font-latin": ("西文/数字字体", spec["latin"]),
        }[code]
        issue(code, f"{label}包含 {'、'.join(sorted(values))}，模板要求 {expected}。", index)
    rule, spacing = resolver.spacing(paragraph)
    tall = any(next(paragraph._p.iter(qn(tag)), None) is not None
               for tag in ("wp:inline", "w:object", "w:pict", "m:oMath"))
    permitted = {"auto"} if spec["spacing_unit"] == "multiple" else (
        {"exact", "atLeast"} if tall else {"exact"})
    if spacing is None:
        unresolved.add("无法解析的行距")
    elif rule not in permitted or abs(spacing - spec["spacing"]) > 0.05:
        label = {"auto": "倍", "exact": "磅（固定）", "atLeast": "磅（最小）"}.get(rule, rule)
        expected = "倍" if spec["spacing_unit"] == "multiple" else "磅（固定；行内对象可用最小值）"
        issue("line-spacing", f"行距为 {spacing:g}{label}，模板要求 {spec['spacing']:g}{expected}。", index)
    if unresolved:
        issue("format-review", "需人工核对：" + "、".join(sorted(unresolved)) + "。", index,
              severity="info")
