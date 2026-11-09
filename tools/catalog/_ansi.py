"""ANSI colour helpers shared by the catalog report scripts."""

from __future__ import annotations

import os
import sys


def _supports_color() -> bool:
    if os.environ.get("NO_COLOR"):
        return False
    return sys.stdout.isatty()


_ANSI = _supports_color()


def _c(code: str, text: str) -> str:
    if not _ANSI:
        return text
    return f"\033[{code}m{text}\033[0m"


def green(s: str) -> str:
    return _c("32", s)


def red(s: str) -> str:
    return _c("31", s)


def yellow(s: str) -> str:
    return _c("33", s)


def bold(s: str) -> str:
    return _c("1", s)


def dim(s: str) -> str:
    return _c("2", s)
