"""Backward compatible entry point; implementation lives in word_formatter.cli."""
from word_formatter.cli import *

if __name__ == "__main__":
    raise SystemExit(main())
