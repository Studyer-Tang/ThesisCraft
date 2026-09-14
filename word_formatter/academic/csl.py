"""CSL rendering through citeproc-py, with upstream versioned styles."""

from pathlib import Path
import re


def render(records, keys, style_name, custom_path=""):
    from citeproc import (
        CitationStylesStyle,
        CitationStylesBibliography,
        Citation,
        CitationItem,
        formatter,
    )
    from citeproc.source.json import CiteProcJSON

    filename = {
        "gb7714-numeric": "china-national-standard-gb-t-7714-2015-numeric",
        "apa": "apa",
        "ieee": "ieee",
    }[style_name]
    path = (
        Path(custom_path)
        if custom_path
        else Path(__file__).parent / "styles" / (filename + ".csl")
    )
    types = dict(
        journal="article-journal",
        book="book",
        conference="paper-conference",
        thesis="thesis",
        web="webpage",
        report="report",
    )
    data = []
    for row in records:
        names = []
        for name in row["authors"]:
            if re.search("[\u4e00-\u9fff]", name):
                names.append({"literal": name})
            elif "," in name:
                family, given = name.split(",", 1)
                names.append(dict(family=family.strip(), given=given.strip()))
            elif " " in name:
                given, family = name.rsplit(" ", 1)
                names.append(dict(family=family, given=given))
            else:
                names.append(dict(literal=name))
        year = re.search(r"\d{4}", row["year"])
        item = dict(
            id=row["key"], type=types[row["type"]], title=row["title"], author=names
        )
        if year:
            item["issued"] = {"date-parts": [[int(year.group())]]}
        for key, target in [
            ("journal", "container-title"),
            ("booktitle", "container-title"),
            ("publisher", "publisher"),
            ("place", "publisher-place"),
            ("volume", "volume"),
            ("issue", "issue"),
            ("pages", "page"),
            ("doi", "DOI"),
            ("url", "URL"),
        ]:
            if row.get(key):
                item[target] = row[key]
        if row.get("accessed"):
            digits = [int(n) for n in re.findall(r"\d+", row["accessed"])[:3]]
            if digits:
                item["accessed"] = {"date-parts": [digits]}
        data.append(item)
    style = CitationStylesStyle(str(path), validate=False)
    bibliography = CitationStylesBibliography(
        style, CiteProcJSON(data), formatter.plain
    )
    for key in keys:
        bibliography.register(Citation([CitationItem(key)]))
    bibliography.sort()
    texts = {
        key: re.sub(r"^\s*\[\d+\]\s*", "", str(text)).strip()
        for key, text in zip(bibliography.keys, bibliography.bibliography())
    }

    def cite(cite_keys):
        citation = Citation([CitationItem(k) for k in cite_keys])
        bibliography.register(citation)
        return str(bibliography.cite(citation, lambda item: None))

    return texts, cite
