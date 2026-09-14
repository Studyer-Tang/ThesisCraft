"""Isolated, bounded Office field updates and accessible PDF export on generated copies."""

import argparse
import json
import os
from pathlib import Path
import subprocess
import sys

SAFE_FIELDS = {
    "TOC",
    "SEQ",
    "REF",
    "PAGEREF",
    "PAGE",
    "NUMPAGES",
    "STYLEREF",
    "SECTIONPAGES",
}


def update_fields(document):
    count, errors = 0, []
    for _ in range(2):
        for story in document.StoryRanges:
            current = story
            while current is not None:
                for item in current.Fields:
                    code = str(item.Code.Text).strip()
                    if code.split(" ", 1)[0].upper() in SAFE_FIELDS:
                        try:
                            item.Update()
                            count += 1
                        except Exception as exc:
                            errors.append(str(exc))
                try:
                    current = current.NextStoryRange
                except Exception:
                    current = None
        document.Repaginate()
    return dict(updated=count, errors=errors)


def finalize(path, host="word", pdf=False):
    path = Path(path).resolve()
    report = path.with_suffix(".office.json")
    command = (
        [sys.executable, "--academic-office"]
        if getattr(sys, "frozen", False)
        else [sys.executable, "-m", "word_formatter.academic.office_io"]
    )
    command += [str(path), "--host", host, "--report", str(report)]
    if pdf:
        command.append("--pdf")
    result = subprocess.run(
        command,
        timeout=180,
        capture_output=True,
        text=True,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    if report.exists():
        return json.loads(report.read_text(encoding="utf-8"))
    raise RuntimeError("Office 处理失败：" + result.stderr[-800:])


def main(argv=None):
    parser = argparse.ArgumentParser()
    parser.add_argument("path")
    parser.add_argument("--host", choices=["word", "wps"], default="word")
    parser.add_argument("--pdf", action="store_true")
    parser.add_argument("--report", required=True)
    args = parser.parse_args(argv)
    report = dict(host=args.host, success=False)
    app, doc = None, None
    try:
        if os.name != "nt":
            raise RuntimeError("自动更新与 PDF 导出需要 Windows Word/WPS。")
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        app = win32com.client.DispatchEx(
            "Word.Application" if args.host == "word" else "KWPS.Application"
        )
        app.Visible, app.DisplayAlerts = False, 0
        app.AutomationSecurity = 3
        doc = app.Documents.Open(
            str(Path(args.path).resolve()), ReadOnly=False, AddToRecentFiles=False
        )
        report.update(update_fields(doc))
        doc.Save()
        report["pages"] = doc.ComputeStatistics(2)
        if args.pdf:
            pdf_path = str(Path(args.path).resolve().with_suffix(".pdf"))
            try:
                doc.ExportAsFixedFormat(
                    OutputFileName=pdf_path,
                    ExportFormat=17,
                    OpenAfterExport=False,
                    CreateBookmarks=1,
                    DocStructureTags=True,
                )
            except Exception:
                # WPS editions differ in named-argument support.
                doc.ExportAsFixedFormat(pdf_path, 17)
            report["pdf"] = pdf_path
        report["success"] = not report["errors"]
    except Exception as exc:
        report["error"] = str(exc)
    finally:
        if doc is not None:
            doc.Close(False)
        if app is not None:
            app.Quit()
        if "pythoncom" in locals():
            pythoncom.CoUninitialize()
        Path(args.report).write_text(
            json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8"
        )
    return 0 if report["success"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
