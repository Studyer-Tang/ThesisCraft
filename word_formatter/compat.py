# ruff: noqa: F401,F403
"""Public compatibility facade for the original flat modules."""

from .constants import *
from .conversion import (
    IS_WINDOWS,
    IS_LINUX,
    WPSAppManager,
    SofficeConverter,
    LegacyConversionUnavailable,
    _initialize_com_for_thread,
    _uninitialize_com_for_thread,
)
from .engine import WordProcessor
