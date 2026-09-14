"""Compliance findings, object-integrity evidence and offline review reports."""

from collections import Counter
import hashlib
import html
import json
from pathlib import Path
import re
from docx.oxml.ns import qn
from .structure import scan
from .templates import STYLE_NAMES, STYLE_LABELS


def inventory(doc):
    parts = doc.part.package.parts
    objects = {}
    for prefix in ("/word/media/", "/word/embeddings/", "/customXml/"):

        def digest(part):
            blob = part.blob
            if prefix == "/customXml/":
                from lxml import etree

                # Office may rewrite declarations and indentation without changing XML data.
                node = etree.fromstring(blob, etree.XMLParser(remove_blank_text=True))
                blob = etree.tostring(node, method="c14n", exclusive=True)
            return hashlib.sha256(blob).hexdigest()

        objects[prefix] = dict(
            Counter(
                digest(part)
                for part in parts
                if str(part.partname).startswith(prefix)
                and not (
                    prefix == "/customXml/"
                    and Path(str(part.partname)).name.startswith("itemProps")
                )
            )
        )
    for tag in ("m:oMath", "w:ins", "w:del"):
        from lxml import etree

        objects[tag] = dict(
            Counter(
                hashlib.sha256(etree.tostring(n, method="c14n")).hexdigest()
                for n in doc.element.iter(qn(tag))
            )
        )
    objects["bookmarks"] = [
        n.get(qn("w:name")) for n in doc.element.iter(qn("w:bookmarkStart"))
    ]
    objects["field_codes"] = [n.text or "" for n in doc.element.iter(qn("w:instrText"))]
    objects["simple_fields"] = [
        n.get(qn("w:instr"), "") for n in doc.element.iter(qn("w:fldSimple"))
    ]
    objects["text_count"] = sum(len(n.text or "") for n in doc.element.iter(qn("w:t")))
    objects["paragraphs"] = len(doc.paragraphs)
    return objects


def compare_inventory(before, after):
    losses = []
    for key in (
        "/word/media/",
        "/word/embeddings/",
        "/customXml/",
        "m:oMath",
        "w:ins",
        "w:del",
    ):
        if Counter(before[key]) - Counter(after[key]):
            losses.append(key)
    for key in ("bookmarks", "field_codes", "simple_fields"):
        if Counter(before[key]) - Counter(after[key]):
            losses.append(key)
    return dict(
        passed=not losses,
        lost_categories=losses,
        source_characters=before["text_count"],
        output_characters=after["text_count"],
        note="比较图片/嵌入对象/自定义XML、公式、修订节点、书签和原有域；文字数量包含新增目录和编号，不代表视觉分页验收。",
    )


def audit(doc, template):
    items = scan(doc, template["structure_overrides"])
    findings = []

    def issue(code, message, index=-1, severity="warning", group="styles"):
        findings.append(
            dict(
                code=code, message=message, index=index, severity=severity, group=group
            )
        )

    regions = {item.region for item in items}
    for region in template["required_sections"]:
        if region not in regions:
            issue(
                "missing-section",
                "未识别到必需部分：" + STYLE_LABELS.get(region, region),
                severity="error",
                group="structure",
            )
    last_level = 0
    paragraphs = list(doc.paragraphs)
    for item in items:
        p = paragraphs[item.index]
        key = (
            item.kind
            if item.kind in STYLE_NAMES
            else "front_heading"
            if item.kind.endswith("_heading")
            else None
        )
        if item.confidence == "review":
            issue(
                "structure-review",
                "该段由编号规则推断，请确认是否为标题/题注。",
                item.index,
                group="structure",
            )
        if item.kind in ("h1", "h2", "h3", "h4"):
            level = int(item.kind[-1])
            if level > last_level + 1:
                issue(
                    "heading-gap",
                    f"标题层级从 {last_level} 跳到 {level}。",
                    item.index,
                    group="numbering",
                )
            last_level = level
            if (
                p._p.find(qn("w:pPr") + "/" + qn("w:numPr")) is None
                and template["numbering"]["scheme"] != "none"
            ):
                issue(
                    "manual-number",
                    "标题尚未使用自动多级编号。",
                    item.index,
                    group="numbering",
                )
        if key and item.text:
            spec = template["styles"][key]
            style = p.style
            run = next((r for r in p.runs if r.text.strip()), None)
            if run is not None:
                chain = [run.font]
                seen = set()
                while style is not None and style.style_id not in seen:
                    seen.add(style.style_id)
                    chain.append(style.font)
                    style = style.base_style
                size = next((f.size.pt for f in chain if f.size is not None), None)
                if size is not None and abs(size - spec["size"]) > 0.1:
                    issue(
                        "font-size",
                        f"字号 {size:g} 磅，模板要求 {spec['size']:g} 磅。",
                        item.index,
                    )
            if p.style.name != STYLE_NAMES[key]:
                issue("style", "建议应用样式：" + STYLE_NAMES[key], item.index)
        if "{{" in item.text or "【" in item.text:
            issue(
                "placeholder",
                "仍有待填写内容或未解析标记。",
                item.index,
                severity="error",
                group="content",
            )
        if re.search(
            r"Error!|错误[!！]|未找到引用源|Bookmark not defined|Reference source not found",
            item.text,
            re.I,
        ):
            issue(
                "broken-field",
                "发现失效域或交叉引用错误。",
                item.index,
                severity="error",
                group="references",
            )
    for n, section in enumerate(doc.sections):
        for key in ("top", "bottom", "left", "right"):
            actual = getattr(section, key + "_margin")
            if actual is not None and abs(actual.cm - template["page"][key]) > 0.03:
                issue(
                    "margin",
                    f"第 {n + 1} 节{key}页边距为 {actual.cm:.2f}cm，要求 {template['page'][key]:g}cm。",
                    group="pages",
                )
    names = {n.get(qn("w:name")) for n in doc.element.iter(qn("w:bookmarkStart"))}
    codes = [n.text or "" for n in doc.element.iter(qn("w:instrText"))] + [
        n.get(qn("w:instr"), "") for n in doc.element.iter(qn("w:fldSimple"))
    ]
    for code in codes:
        match = re.match(r'\s*(?:REF|PAGEREF)\s+"?([^\s"\\]+)', code)
        if match and match.group(1) not in names:
            issue(
                "broken-bookmark",
                "找不到书签：" + match.group(1),
                severity="error",
                group="references",
            )
    if any(
        "ZOTERO" in c.upper() or "EN.CITE" in c.upper() or "EN.REFLIST" in c.upper()
        for c in codes
    ):
        issue(
            "citation-manager",
            "检测到文献管理器域，将保留；最终文献样式须在原管理器中更新。",
            severity="info",
            group="references",
        )
    if next(doc.element.iter(qn("w:ins")), None) is not None or next(doc.element.iter(qn("w:del")), None) is not None:
        issue(
            "revisions",
            "存在修订记录，排版保留修订；交稿前需自行处理。",
            group="content",
        )
    if any(doc.element.iter(qn("w:txbxContent"))):
        issue(
            "textboxes",
            "存在文本框，保留其内容与独立布局，需在打印预览确认。",
            group="content",
        )
    if next(doc.element.iter(qn('wp:anchor')), None) is not None:
        issue('floating-images', '浮动图片保留原有锚点、环绕与位置，请在打印预览检查遮挡。', group='figures')
    for index, shape in enumerate(doc.inline_shapes):
        if shape.width.cm > template["figures"]["max_width_cm"] + 0.01:
            issue(
                "image-width",
                f"图片 {index + 1} 宽度 {shape.width.cm:.2f}cm 超过模板限制。",
                group="figures",
            )
        blip = shape._inline.find(".//" + qn("a:blip"))
        if blip is not None and blip.get(qn("r:embed")):
            try:
                from docx.image.image import Image

                image = Image.from_blob(
                    doc.part.related_parts[blip.get(qn("r:embed"))].blob
                )
                dpi = image.px_width / (shape.width.cm / 2.54)
                if dpi < template["figures"]["min_dpi"]:
                    issue(
                        "image-dpi",
                        f"图片 {index + 1} 有效分辨率约 {dpi:.0f} DPI，低于 {template['figures']['min_dpi']}。",
                        group="figures",
                    )
            except Exception:
                issue(
                    "image-dpi-unknown",
                    f"图片 {index + 1} 格式无法测量清晰度。",
                    group="figures",
                )
    if template["school"] == "北京大学":
        issue("school-scope", template["scope"], severity="info", group="template")
        abstract = "".join(i.text for i in items if i.kind == "abstract")
        length = len(re.findall(r"[\u4e00-\u9fff]", abstract))
        if (template["degree"] == "doctor" and not 800 <= length <= 1000) or (
            template["degree"] == "master" and not 500 <= length <= 700
        ):
            issue(
                "abstract-length",
                f"中文摘要约 {length} 个汉字；指南建议硕士约600字、博士800–1000字，请结合内容确认。",
                group="content",
            )
    return findings, items


def write_report(directory, report):
    directory = Path(directory)
    json_path, html_path = directory / "检查报告.json", directory / "检查报告.html"
    json_path.write_text(
        json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
    )
    esc = lambda value: html.escape(str(value))
    issues = report.get("after_issues", report["before_issues"])
    counts = Counter(i.get("severity", "warning") for i in issues)
    levels = dict(error="错误", warning="待确认", info="提示")
    rows = "".join(
        f'<tr data-level="{esc(i.get("severity", "warning"))}"><td>{levels.get(i.get("severity", "warning"), "提示")}</td><td><a href="#p{i.get("index", -1)}">{i.get("index", -1) + 1 if i.get("index", -1) >= 0 else "全文"}</a></td><td>{esc(i["message"])}</td></tr>'
        for i in issues
    )
    structure = "".join(
        f'<tr id="p{i["index"]}"><td>{i["index"] + 1}</td><td>{esc(STYLE_LABELS.get(i["kind"], STYLE_LABELS.get(i["kind"].removesuffix("_heading"), "保留区域")))}</td><td>{esc(i["text"])}</td></tr>'
        for i in report["structure"]
    )
    changes = "".join(
        f"<tr><td>{esc(c.get('index', -1) + 1)}</td><td>{esc(c['action'])}</td><td>{esc(c.get('before', ''))}</td><td>{esc(c.get('after', ''))}</td></tr>"
        for c in report.get("changes", [])
    )
    source_link = Path(report["source"]).as_uri()
    result_link = (
        f'<a class="button" href="{esc(Path(report["output"]).as_uri())}">打开排版副本</a>'
        if report.get("output")
        else ""
    )
    integrity = report.get("integrity")
    integrity_text = (
        (
            "图片、公式、嵌入对象及原有引用保护检查："
            + ("通过" if integrity["passed"] else "未通过")
        )
        if integrity
        else "本次仅检查，文档未修改。"
    )
    office = report.get("office", {})
    office_text = (
        "目录和引用已更新。"
        if office.get("success")
        else "可使用插件“更新目录与引用”刷新页码。"
    )
    preview = ""
    if office.get("pdf") and Path(office["pdf"]).exists():
        uri = esc(Path(office["pdf"]).as_uri())
        result_link += f'<a class="button" href="{uri}">打开 PDF / 打印预览</a>'
        preview = f'<details><summary>PDF 分页预览</summary><object data="{uri}" type="application/pdf" width="100%" height="800"><a href="{uri}">使用系统阅读器查看 PDF</a></object></details>'
        check = office.get("pdf_check", {})
        office_text += f" PDF 共 {check.get('pages', office.get('pages', ''))} 页。"
        if check.get("possibly_blank_pages"):
            office_text += (
                " 发现空白页："
                + ", ".join(map(str, check["possibly_blank_pages"]))
                + "；请确认是否为双面打印留白。"
            )
        if check.get("fonts_without_embedded_data"):
            office_text += " 待检查字体嵌入：" + ", ".join(
                check["fonts_without_embedded_data"]
            )
    html_path.write_text(
        f'''<!doctype html><html lang="zh-CN"><meta charset="utf-8"><title>学研排版 · 检查报告</title>
<style>body{{font:16px/1.7 'Microsoft YaHei',sans-serif;background:#f5f7fb;color:#18263c;margin:0}}main{{max-width:1120px;margin:32px auto;padding:32px;background:white;border-radius:16px}}h1{{margin-top:0}}table{{border-collapse:collapse;width:100%;font-size:14px;table-layout:fixed}}td,th{{padding:10px;border-bottom:1px solid #ddd;text-align:left;overflow-wrap:anywhere}}th{{background:#edf2fa}}.button,button{{display:inline-block;padding:8px 14px;background:#174c87;color:white;border:0;border-radius:6px;margin:4px;text-decoration:none;cursor:pointer}}.status{{padding:18px;background:#eef5ff}}details{{margin:24px 0}}tr:target{{background:#fff2c4}}</style>
<main><h1>学研排版 · 检查报告</h1><p>{esc(report["template_name"])} · {esc(report.get("duration_seconds", 0))} 秒</p>
<p>{result_link}<a class="button" href="{esc(source_link)}">打开原件 / 回退</a></p>
<p class="status">错误 {counts["error"]} · 待确认 {counts["warning"]} · 提示 {counts["info"]}<br>原件始终保留。此报告检查结构与格式，最终分页请在 Word/WPS 或 PDF 预览中确认。</p>
<p>{esc(integrity_text)}</p><p>{esc(office_text)}</p>{preview}
<button onclick="filter('all')">全部</button><button onclick="filter('error')">只看错误</button><button onclick="filter('warning')">待确认</button>
<table id="issues"><thead><tr><th style="width:90px">级别</th><th style="width:70px">段落</th><th>问题与建议</th></tr></thead><tbody>{rows}</tbody></table>
<details><summary>排版前后变化（文字层面）</summary><table><tr><th>段落</th><th>操作</th><th>原内容</th><th>新内容</th></tr>{changes}</table></details>
<details open><summary>文档结构与段落定位</summary><table><tr><th style="width:70px">段落</th><th style="width:160px">识别类型</th><th>段落内容</th></tr>{structure}</table></details>
<p>ThesisCraft · Study-Tang · 本地处理</p></main><script>function filter(v){{document.querySelectorAll('#issues tbody tr').forEach(r=>r.hidden=v!=='all'&&r.dataset.level!==v)}}</script></html>''',
        encoding="utf-8",
    )
    return html_path
