"""Generate a standalone Skill from the single authoritative Python package."""

from pathlib import Path
import shutil
import zipfile
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from word_formatter.version import __version__  # noqa: E402


def main():
    target = ROOT / "dist/doc-format"
    if target.exists():
        if not target.resolve().is_relative_to((ROOT / "dist").resolve()):
            raise ValueError("Unsafe output directory")
        shutil.rmtree(target)
    shutil.copytree(ROOT / "skills/doc-format", target)
    shutil.copytree(
        ROOT / "word_formatter",
        target / "scripts/word_formatter",
        ignore=shutil.ignore_patterns("__pycache__", "*.pyc"),
    )
    shutil.copy2(ROOT / "LICENSE", target / "LICENSE")
    archive = ROOT / f"dist/doc-format-{__version__}.zip"
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED) as bundle:
        for file in target.rglob("*"):
            if file.is_file():
                bundle.write(file, file.relative_to(target.parent))
    print(archive)


if __name__ == "__main__":
    main()
