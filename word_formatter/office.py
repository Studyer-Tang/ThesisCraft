"""Office add-in boundary: saved local files, user configuration, new output."""

import sys
from .cli import main as cli_main
from .storage import user_config_path


def main(argv=None):
    argv = list(sys.argv[1:] if argv is None else argv)
    if len(argv) != 1:
        print("需要一个已保存的本地文档路径。", file=sys.stderr)
        return 1
    args = ["format", argv[0], "--json", "--set", "preprocess_office=false"]
    config = user_config_path()
    if config.exists():
        args += ["--config", str(config)]
    else:
        args += ["--config-json", "{}"]
    return cli_main(args)


if __name__ == "__main__":
    raise SystemExit(main())
