"""Bundle installed dependency versions and their shipped license notices."""

from importlib import metadata
import json
from pathlib import Path
import platform
import shutil
import sys


def collect(destination):
    destination.mkdir(parents=True, exist_ok=True)
    versions = {}
    for distribution in metadata.distributions():
        name = distribution.metadata["Name"]
        versions[name] = distribution.version
        for item in distribution.files or ():
            if item.name.lower().startswith(("license", "copying", "notice")):
                source = Path(distribution.locate_file(item))
                if source.is_file():
                    target = destination / name / Path(*[p for p in item.parts if p not in ("..", ".")])
                    target.parent.mkdir(parents=True, exist_ok=True)
                    shutil.copy2(source, target)
    for name in ("LICENSE.txt", "LICENSE"):
        source = Path(sys.base_prefix) / name
        if source.is_file():
            shutil.copy2(source, destination / ("Python-" + name))
    (destination / "build-environment.json").write_text(
        json.dumps(dict(python=platform.python_version(), system=platform.system(),
                        architecture=platform.machine(), dependencies=versions),
                   ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")


if __name__ == "__main__":
    collect(Path(sys.argv[1]))
