"""Offline, source-backed university profiles and lossless Word template copies."""

from copy import deepcopy
from functools import lru_cache
import hashlib
import json
from pathlib import Path
from tempfile import TemporaryDirectory
import zipfile

DATA = Path(__file__).with_name("catalog_data")
CATEGORIES = {
    "course": "课程论文",
    "bachelor": "本科",
    "master": "硕士",
    "doctor": "博士",
}


@lru_cache(maxsize=1)
def _catalog():
    return json.loads((DATA / "index.json").read_text(encoding="utf-8"))


def profiles(query="", category=""):
    words = query.casefold().split()
    return [
        deepcopy(p)
        for p in _catalog()["profiles"]
        if (not category or p["category"] == category)
        and all(
            word
            in " ".join(
                (p["id"], p["name"], p["school"], p["department"], p["keywords"])
            ).casefold()
            for word in words
        )
    ]


def get_profile(identifier):
    return next((p for p in profiles() if p["id"] == identifier), None)


def profile_template(profile):
    from .templates import validate_template

    return validate_template(profile["template"])


def sources(profile):
    return [deepcopy(_catalog()["sources"][key]) for key in profile["source_ids"]]


def source_bytes(source):
    """Validate bundled bytes before exporting; never fetch at app startup."""
    filename = source["file"]
    if Path(filename).name != filename or filename in ("", ".", ".."):
        raise ValueError("资料文件名无效。")
    data = (DATA / "sources" / filename).read_bytes()
    if hashlib.sha256(data).hexdigest() != source["sha256"]:
        raise ValueError("资料校验失败：" + filename)
    return data


def _publish(destination, writer, overwrite=False):
    from ..storage import publish_file

    destination = Path(destination).expanduser().resolve()
    if DATA.resolve() in destination.parents:
        raise ValueError("请保存到个人文件夹，不能覆盖内置资料。")
    destination.parent.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(prefix=".thesiscraft-", dir=destination.parent) as folder:
        temp = Path(folder) / destination.name
        writer(temp)
        publish_file(temp, destination, overwrite=overwrite)
    return destination


def new_document(profile, destination, overwrite=False):
    if Path(destination).suffix.lower() != ".docx":
        raise ValueError("新建 Word 文档请使用 .docx 扩展名。")
    original = profile.get("word_source")
    if original:
        data = source_bytes(_catalog()["sources"][original])
        return _publish(destination, lambda path: path.write_bytes(data), overwrite)
    from .layout import create_skeleton

    return _publish(
        destination,
        lambda path: create_skeleton(profile_template(profile), path),
        overwrite,
    )


def export_bundle(profile, destination, overwrite=False):
    """Export editable rules, provenance, and unmodified downloaded documents together."""
    if Path(destination).suffix.lower() != ".zip":
        raise ValueError("资料包请使用 .zip 扩展名。")
    records = sources(profile)
    assets = [(record["file"], source_bytes(record)) for record in records]

    def write(path):
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr(
                "template.json",
                json.dumps(profile_template(profile), ensure_ascii=False, indent=2),
            )
            archive.writestr(
                "sources.json", json.dumps(records, ensure_ascii=False, indent=2)
            )
            archive.writestr("NOTICE.md", (DATA / "NOTICE.md").read_bytes())
            for filename, data in assets:
                archive.writestr("sources/" + filename, data)

    return _publish(destination, write, overwrite)
