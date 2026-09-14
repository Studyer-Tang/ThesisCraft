"""Academic commands available through both paper-studio and legacy wfp."""

import json
from .templates import load_template, save_template, validate_template


def config(args):
    template = load_template(getattr(args, "template", None))
    for setting in getattr(args, "set", None) or []:
        if "=" not in setting:
            raise ValueError("--set 格式为 path=value")
        path, text = setting.split("=", 1)
        try:
            value = json.loads(text)
        except json.JSONDecodeError:
            value = text
        owner = template
        parts = path.split(".")
        for part in parts[:-1]:
            owner = owner[part]
        if parts[-1] not in owner:
            raise ValueError("未知模板参数：" + path)
        owner[parts[-1]] = value
    if getattr(args, "only", None):
        template["enabled"] = args.only.split(",")
    return validate_template(template)


def execute(args):
    from .workflow import run

    result = run(
        args.input,
        config(args),
        output_dir=args.output_dir,
        check_only=args.action == "check",
        reference_path=args.references,
        host=args.host,
        pdf=args.pdf,
    )
    print(
        json.dumps(
            result
            if args.json
            else {
                k: result.get(k)
                for k in ("output", "report", "duration_seconds", "office")
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    issues = result.get("after_issues", result["before_issues"])
    return 2 if any(i.get("severity") == "error" for i in issues) else 0


def add_parser(subparsers):
    thesis = subparsers.add_parser("thesis", help="论文模板、检查、排版、章节和 PDF")
    commands = thesis.add_subparsers(dest="action", required=True)
    for action, label in [("check", "只检查"), ("format", "按模板生成排版副本")]:
        p = commands.add_parser(action, help=label)
        p.add_argument("-i", "--input", required=True)
        p.add_argument("--output-dir")
        p.add_argument(
            "--template",
            help="pku-master / pku-doctor / bachelor / master / doctor 或 JSON 路径",
        )
        p.add_argument(
            "--set",
            action="append",
            help="例如 styles.body.size=12 或 numbering.scheme=chinese",
        )
        p.add_argument(
            "--only",
            help="逗号分隔修复类别：styles,pages,numbering,captions,tables,figures,references,contents",
        )
        p.add_argument("--references", help="RIS、BibTeX 或 JSON 文献库")
        p.add_argument("--host", choices=["word", "wps"])
        p.add_argument("--pdf", action="store_true")
        p.add_argument("--json", action="store_true")
        p.set_defaults(func=execute)
    p = commands.add_parser("template", help="导出内置模板")
    p.add_argument("--template", default="pku-master")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--set", action="append")
    p.set_defaults(func=lambda a: print(save_template(config(a), a.output)) or 0)
    p = commands.add_parser("new", help="新建论文骨架")
    p.add_argument("--template", default="pku-master")
    p.add_argument("--set", action="append")
    p.add_argument("-o", "--output", required=True)

    def new(a):
        from .layout import create_skeleton

        print(create_skeleton(config(a), a.output))
        return 0

    p.set_defaults(func=new)
    p = commands.add_parser("merge", help="按给定顺序合并章节")
    p.add_argument("paths", nargs="+")
    p.add_argument("-o", "--output", required=True)

    def merge(a):
        from .documents import merge_chapters

        print(merge_chapters(a.paths, a.output))
        return 0

    p.set_defaults(func=merge)
    p = commands.add_parser("split", help="拆分章节副本")
    p.add_argument("-i", "--input", required=True)
    p.add_argument("-o", "--output", required=True)

    def split(a):
        from .documents import split_chapters

        print(json.dumps(split_chapters(a.input, a.output), ensure_ascii=False))
        return 0

    p.set_defaults(func=split)
