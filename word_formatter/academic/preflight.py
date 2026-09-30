"""Cheap, explicit font checks. Never substitute a school's requested typeface."""

ALIASES = (
    {"宋体", "simsun"}, {"黑体", "simhei"}, {"楷体", "kaiti"},
    {"仿宋", "fangsong"}, {"微软雅黑", "microsoft yahei"},
    {"等线", "dengxian"},
)


def font_issues(template, available_fonts=None):
    if available_fonts is None:
        return [dict(code="fonts-unchecked", severity="info", index=-1,
                     message="本次未检测系统字体；请在实际编辑和打印的电脑上检查字体。")]
    available = {name.casefold().lstrip("@") for name in available_fonts}
    for aliases in ALIASES:
        if available & aliases:
            available.update(aliases)
    keys = set(template["style_keys"]) if "styles" in template["enabled"] else set()
    if "tables" in template["enabled"]:
        keys.update(("table", "table_header", "note"))
    required = {template["styles"][key][field] for key in keys for field in ("font", "latin")}
    if "numbering" in template["enabled"]:
        required.update(template["numbering"][field] for field in ("font", "east_asia"))
    missing = sorted(name for name in required if name.casefold() not in available)
    return [dict(code="missing-font", index=-1, severity="warning",
                 message="本机未检测到字体：" + "、".join(missing)
                 + "。已保留模板指定字体；请安装有授权的字体或在高级设置中调整，最终在 Word 中核对分页。")
            ] if missing else []
