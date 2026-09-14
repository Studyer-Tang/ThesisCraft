"""Input collection and output planning shared by GUI, CLI and Office add-ins."""

from dataclasses import dataclass
from pathlib import Path
from .constants import SUPPORTED_FILE_EXTENSIONS
from .storage import ensure_distinct_paths


@dataclass(frozen=True)
class InputRecord:
    source: Path
    relative: Path


@dataclass(frozen=True)
class Job:
    source: Path
    output: Path


def is_supported_file(path):
    path = Path(path)
    return (
        path.is_file()
        and not path.name.startswith(("~", ".wfp-"))
        and path.suffix.lower() in SUPPORTED_FILE_EXTENSIONS
    )


def collect_records(paths, recursive=True, exclude_dir=None):
    if not paths:
        raise ValueError("请提供至少一个输入文件或目录。")
    resolved = [Path(p).expanduser().resolve() for p in paths]
    records, seen = [], set()
    for source in resolved:
        if not source.exists():
            raise FileNotFoundError(f"输入路径不存在: {source}")
        if source.is_file() and not is_supported_file(source):
            raise ValueError(f"不支持的输入文件: {source}")
        files = (
            [source]
            if source.is_file()
            else sorted(source.rglob("*") if recursive else source.glob("*"))
        )
        for file in files:
            if not is_supported_file(file) or file.resolve() in seen:
                continue
            if (
                source.is_dir()
                and exclude_dir is not None
                and file.is_relative_to(exclude_dir)
            ):
                continue
            seen.add(file.resolve())
            relative = Path(file.name) if source.is_file() else file.relative_to(source)
            if source.is_dir() and len(resolved) > 1:
                relative = Path(source.name) / relative
            records.append(InputRecord(file.resolve(), relative))
    if not records:
        raise FileNotFoundError("未找到可处理的文件。")
    return records


def formatted_relative_path(relative):
    return relative.with_name(f"{relative.stem}_formatted.docx")


def unique_path(path, seen, overwrite=False):
    candidate, counter = path, 2
    while candidate in seen or (candidate.exists() and not overwrite):
        candidate = path.with_name(f"{path.stem}_{counter}{path.suffix}")
        counter += 1
    seen.add(candidate)
    return candidate


def build_jobs(input_paths, output_arg=None, recursive=True, overwrite=False, beside_sources=False):
    inputs = [Path(p).expanduser().resolve() for p in input_paths]
    if not inputs:
        raise ValueError("请提供输入文件或目录。")
    output = Path(output_arg).expanduser().resolve() if output_arg else None
    single_file = len(inputs) == 1 and inputs[0].is_file()
    if not single_file and output and output.suffix.lower() == ".docx":
        raise ValueError("目录或多文件输入必须使用输出目录。")
    exclude = (
        output
        if output and output.suffix.lower() != ".docx" and output not in inputs
        else None
    )
    records = collect_records(inputs, recursive, exclude_dir=exclude)
    explicit_file = single_file and output and output.suffix.lower() == ".docx"
    if explicit_file:
        ensure_distinct_paths(records[0].source, output)
        if output.exists() and not overwrite:
            raise FileExistsError(
                f"输出文件已存在，请改名或显式指定 --overwrite: {output}"
            )
        return [Job(records[0].source, output)]
    output_dir = output
    if output_dir is None:
        output_dir = (
            inputs[0].parent
            if single_file
            else inputs[0].parent / f"{inputs[0].name}_formatted"
            if len(inputs) == 1 and inputs[0].is_dir()
            else Path.cwd() / "wfp_formatted"
        )
    seen, jobs = set(), []
    for record in records:
        path = unique_path(
            record.source.with_name(f"{record.source.stem}_formatted.docx")
            if beside_sources and output_arg is None
            else output_dir / formatted_relative_path(record.relative), seen, overwrite
        )
        for source in inputs:
            if source.is_file():
                ensure_distinct_paths(source, path)
        ensure_distinct_paths(record.source, path)
        jobs.append(Job(record.source, path))
    return jobs
