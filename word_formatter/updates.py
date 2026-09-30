"""User-initiated release discovery; no background traffic or executable downloads."""

import json
import re
from urllib.request import Request, urlopen

REPOSITORY = "Studyer-Tang/ThesisCraft"
RELEASES = f"https://github.com/{REPOSITORY}/releases"


def latest_release(preview=False):
    request = Request(f"https://api.github.com/repos/{REPOSITORY}/releases?per_page=30",
                      headers={"Accept": "application/vnd.github+json", "User-Agent": "ThesisCraft"})
    with urlopen(request, timeout=10) as response:
        raw = response.read(1024 * 1024 + 1)
    if len(raw) > 1024 * 1024:
        raise ValueError("更新信息过大，请直接查看项目发布页。")
    records = json.loads(raw)
    if not isinstance(records, list):
        raise ValueError("无法读取发布列表。")
    for release in records:
        tag = release.get("tag_name", "")
        if (release.get("draft") or (release.get("prerelease") and not preview)
                or not re.fullmatch(r"v\d+\.\d+\.\d+(?:-[A-Za-z0-9.]+)?", tag)):
            continue
        return {"tag": tag, "preview": bool(release.get("prerelease")),
                "url": RELEASES + "/tag/" + tag}
    return None
