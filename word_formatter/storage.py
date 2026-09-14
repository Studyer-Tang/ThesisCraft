"""Transactional saves with an explicit no-clobber default."""

import os
import shutil
from pathlib import Path
import tempfile


def publish_file(source, output, overwrite=False):
    """Publish exact bytes atomically, including from another drive's temp folder."""
    output = Path(output).resolve()
    fd, temporary = tempfile.mkstemp(prefix=".wfp-", dir=output.parent)
    os.close(fd)
    try:
        shutil.copyfile(source, temporary)
        if overwrite:
            os.replace(temporary, output)
        else:
            os.link(temporary, output)
    finally:
        Path(temporary).unlink(missing_ok=True)


def ensure_distinct_paths(source, output):
    source, output = Path(source).resolve(), Path(output).resolve()
    if source == output or (output.exists() and os.path.samefile(source, output)):
        raise ValueError("输入和输出不能是同一文件，请另选输出文件名。")


def save_document(document, output, overwrite=False):
    output = Path(output).resolve()
    output.parent.mkdir(parents=True, exist_ok=True)
    fd, temporary = tempfile.mkstemp(prefix=".wfp-", suffix=".docx", dir=output.parent)
    os.close(fd)
    try:
        document.save(temporary)
        if overwrite:
            os.replace(temporary, output)
        else:
            # Hard linking publishes the completed file atomically and fails if
            # another writer has created the destination since job planning.
            os.link(temporary, output)
    finally:
        if os.path.exists(temporary):
            os.unlink(temporary)


def user_config_path():
    if os.name == "nt":
        base = Path(os.environ.get("APPDATA", Path.home() / "AppData/Roaming"))
    else:
        base = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
    return base / "Study-Tang" / "ThesisCraft" / "config.json"
