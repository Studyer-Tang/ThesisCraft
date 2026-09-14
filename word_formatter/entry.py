# -*- coding: utf-8 -*-
"""Compatibility entry point for ThesisCraft."""

import sys
import os

from word_formatter.version import __version__


def main(argv=None):
    if sys.stdout is None:
        sys.stdout = open(os.devnull, 'w', encoding='utf-8')
    if sys.stderr is None:
        sys.stderr = open(os.devnull, 'w', encoding='utf-8')
    argv = list(sys.argv[1:] if argv is None else argv)
    if argv and argv[0] == '--academic':
        from word_formatter.academic.gui import main as academic_main
        return academic_main(argv[1:])
    if argv and argv[0] == '--academic-office':
        from word_formatter.academic.office_io import main as academic_office_main
        return academic_office_main(argv[1:])
    if argv and argv[0] == '--office-watch':
        from word_formatter.academic.plugin import watch
        return watch()
    if argv and argv[0] == '--diagnose':
        from word_formatter.diagnostics import main as diagnostic_main
        return diagnostic_main(argv[1:])
    if argv and argv[0] == '--office':
        from word_formatter.office_toolbar import main as office_main
        return office_main(argv[1:])
    if argv and argv[0] == '--cli':
        from word_formatter.cli import main as cli_main
        return cli_main(argv[1:])
    if argv and argv[0] == "--test":
        from word_formatter.legacy_tests import main as test_main

        return test_main(argv[1:])
    if argv and argv[0] in ("--version", "-V"):
        print(__version__)
        return 0

    from word_formatter.gui import main as gui_main

    gui_main()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
