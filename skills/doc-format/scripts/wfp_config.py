"""Skill launcher. Release bundles contain an automatically generated package."""
from pathlib import Path
import sys
if not (Path(__file__).parent / "word_formatter").is_dir():
    sys.path.insert(0, str(Path(__file__).resolve().parents[3]))
from word_formatter.config import *
