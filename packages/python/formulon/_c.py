# Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
"""Low-level ctypes bindings for libformulon.

This module is an implementation detail. Public users should import the
``Workbook`` / ``Value`` / ``FormulonError`` symbols re-exported by the
top-level :mod:`formulon` package instead.

The bindings here are intentionally a 1:1 mirror of ``src/c_api/formulon_c.h``:
every ``fm_*`` export is declared with explicit ``argtypes`` / ``restype``,
and the ``fm_value_t`` POD is re-expressed as a ctypes ``Structure`` whose
union layout matches the C side byte-for-byte.

The shared library is located via :func:`_load_library`, which prefers the
in-package ``_lib/`` location populated by ``packages/python/scripts/stage.py``
and falls back to the system loader path for development installs that
build ``libformulon`` out-of-tree.
"""

from __future__ import annotations

import ctypes
import os
import platform
import sys
from ctypes import (
    CDLL,
    POINTER,
    Structure,
    Union,
    c_char_p,
    c_double,
    c_int,
    c_int32,
    c_size_t,
    c_uint32,
    c_void_p,
)
from enum import IntEnum
from pathlib import Path
from typing import Optional

__all__ = [
    "LIB",
    "ValueKind",
    "fm_status_t",
    "fm_workbook_t_p",
    "fm_value_t",
]


# ---------------------------------------------------------------------------
# Library loading
# ---------------------------------------------------------------------------


def _candidate_lib_names() -> tuple[str, ...]:
    """Return the platform-specific shared-library file names to try."""
    system = platform.system()
    if system == "Darwin":
        return ("libformulon.dylib",)
    if system == "Windows":
        return ("formulon.dll", "libformulon.dll")
    # Treat everything else as ELF / POSIX.
    return ("libformulon.so",)


def _load_library() -> CDLL:
    """Locate and dlopen the libformulon shared library.

    Search order:
      1. ``packages/python/formulon/_lib/<name>`` (the package-data location
         populated by ``make python-package``).
      2. The platform's default loader path (so a development install that
         built ``libformulon`` somewhere on ``LD_LIBRARY_PATH`` /
         ``DYLD_LIBRARY_PATH`` / ``PATH`` still works).

    Raises:
      OSError: when neither location resolves to a loadable file. The error
        message includes the candidate names that were tried.
    """
    here = Path(__file__).resolve().parent
    lib_dir = here / "_lib"
    candidates = _candidate_lib_names()
    tried: list[str] = []

    for name in candidates:
        local = lib_dir / name
        if local.is_file():
            return CDLL(str(local))
        tried.append(str(local))

    # Fall back to the OS loader.
    for name in candidates:
        try:
            return CDLL(name)
        except OSError:
            tried.append(name)
            continue

    raise OSError(
        "formulon: failed to locate libformulon shared library. "
        f"Tried: {', '.join(tried)}. "
        "Run `make python-package` to stage the library, or install a "
        "wheel that ships it under formulon/_lib/."
    )


# ---------------------------------------------------------------------------
# Type definitions mirroring src/c_api/formulon_c.h
# ---------------------------------------------------------------------------


# fm_status_t: int32_t.
fm_status_t = c_int32

# fm_workbook_t*: opaque pointer.
fm_workbook_t_p = c_void_p


class ValueKind(IntEnum):
    """Mirror of ``fm_value_kind_t`` in ``formulon_c.h``.

    Numbering matches the C ABI ordinals exactly.
    """

    BLANK = 0
    NUMBER = 1
    BOOL = 2
    TEXT = 3
    ERROR = 4
    ARRAY = 5
    REF = 6
    LAMBDA = 7


class _FmValueU(Union):
    _fields_ = [
        ("number", c_double),
        ("boolean", c_int32),
        ("error_code", c_int32),
        ("text", c_char_p),
    ]


class fm_value_t(Structure):  # noqa: N801 -- mirrors C struct name
    """Mirror of ``fm_value_t``.

    The ``kind`` field is a C ``int`` (matching the underlying ``enum``).
    The active union member is selected by ``kind``; reading any other
    member is undefined per the C ABI contract.
    """

    _fields_ = [
        ("kind", c_int),
        ("u", _FmValueU),
    ]


# ---------------------------------------------------------------------------
# Signature setup
# ---------------------------------------------------------------------------


def _setup_signatures(lib: CDLL) -> None:
    """Bind argtypes / restype for every fm_* export.

    Called exactly once at module import time. The list mirrors the
    ordering of declarations in ``src/c_api/formulon_c.h``.
    """
    # -- Construction / lifecycle ------------------------------------------
    lib.fm_workbook_create.argtypes = [POINTER(fm_workbook_t_p)]
    lib.fm_workbook_create.restype = fm_status_t

    lib.fm_workbook_create_empty.argtypes = [POINTER(fm_workbook_t_p)]
    lib.fm_workbook_create_empty.restype = fm_status_t

    lib.fm_workbook_load.argtypes = [
        POINTER(ctypes.c_uint8),
        c_size_t,
        POINTER(fm_workbook_t_p),
    ]
    lib.fm_workbook_load.restype = fm_status_t

    lib.fm_workbook_destroy.argtypes = [fm_workbook_t_p]
    lib.fm_workbook_destroy.restype = None

    # -- Save --------------------------------------------------------------
    lib.fm_workbook_save.argtypes = [
        fm_workbook_t_p,
        POINTER(POINTER(ctypes.c_uint8)),
        POINTER(c_size_t),
    ]
    lib.fm_workbook_save.restype = fm_status_t

    lib.fm_buffer_free.argtypes = [POINTER(ctypes.c_uint8)]
    lib.fm_buffer_free.restype = None

    # -- Sheets ------------------------------------------------------------
    lib.fm_workbook_sheet_count.argtypes = [fm_workbook_t_p]
    lib.fm_workbook_sheet_count.restype = c_size_t

    lib.fm_workbook_sheet_name.argtypes = [
        fm_workbook_t_p,
        c_size_t,
        POINTER(c_char_p),
    ]
    lib.fm_workbook_sheet_name.restype = fm_status_t

    lib.fm_workbook_add_sheet.argtypes = [fm_workbook_t_p, c_char_p]
    lib.fm_workbook_add_sheet.restype = fm_status_t

    # -- Cell mutation -----------------------------------------------------
    lib.fm_workbook_set_number.argtypes = [
        fm_workbook_t_p,
        c_size_t,
        c_uint32,
        c_uint32,
        c_double,
    ]
    lib.fm_workbook_set_number.restype = fm_status_t

    lib.fm_workbook_set_bool.argtypes = [
        fm_workbook_t_p,
        c_size_t,
        c_uint32,
        c_uint32,
        c_int32,
    ]
    lib.fm_workbook_set_bool.restype = fm_status_t

    lib.fm_workbook_set_text.argtypes = [
        fm_workbook_t_p,
        c_size_t,
        c_uint32,
        c_uint32,
        c_char_p,
    ]
    lib.fm_workbook_set_text.restype = fm_status_t

    lib.fm_workbook_set_blank.argtypes = [
        fm_workbook_t_p,
        c_size_t,
        c_uint32,
        c_uint32,
    ]
    lib.fm_workbook_set_blank.restype = fm_status_t

    lib.fm_workbook_set_formula.argtypes = [
        fm_workbook_t_p,
        c_size_t,
        c_uint32,
        c_uint32,
        c_char_p,
    ]
    lib.fm_workbook_set_formula.restype = fm_status_t

    # -- Cell read ---------------------------------------------------------
    lib.fm_workbook_get_value.argtypes = [
        fm_workbook_t_p,
        c_size_t,
        c_uint32,
        c_uint32,
        POINTER(fm_value_t),
    ]
    lib.fm_workbook_get_value.restype = fm_status_t

    # -- Iteration / dump --------------------------------------------------
    lib.fm_workbook_cell_count.argtypes = [
        fm_workbook_t_p,
        c_size_t,
        POINTER(c_size_t),
    ]
    lib.fm_workbook_cell_count.restype = fm_status_t

    lib.fm_workbook_cell_at.argtypes = [
        fm_workbook_t_p,
        c_size_t,
        c_size_t,
        POINTER(c_uint32),
        POINTER(c_uint32),
        POINTER(c_char_p),
        POINTER(fm_value_t),
    ]
    lib.fm_workbook_cell_at.restype = fm_status_t

    lib.fm_workbook_defined_name_count.argtypes = [fm_workbook_t_p]
    lib.fm_workbook_defined_name_count.restype = c_size_t

    lib.fm_workbook_defined_name_at.argtypes = [
        fm_workbook_t_p,
        c_size_t,
        POINTER(c_char_p),
        POINTER(c_char_p),
    ]
    lib.fm_workbook_defined_name_at.restype = fm_status_t

    lib.fm_workbook_table_count.argtypes = [fm_workbook_t_p]
    lib.fm_workbook_table_count.restype = c_size_t

    lib.fm_workbook_table_at.argtypes = [
        fm_workbook_t_p,
        c_size_t,
        POINTER(c_char_p),
        POINTER(c_char_p),
        POINTER(c_char_p),
        POINTER(c_size_t),
    ]
    lib.fm_workbook_table_at.restype = fm_status_t

    lib.fm_workbook_passthrough_count.argtypes = [fm_workbook_t_p]
    lib.fm_workbook_passthrough_count.restype = c_size_t

    lib.fm_workbook_passthrough_at.argtypes = [
        fm_workbook_t_p,
        c_size_t,
        POINTER(c_char_p),
    ]
    lib.fm_workbook_passthrough_at.restype = fm_status_t

    # -- Recalc ------------------------------------------------------------
    lib.fm_workbook_recalc.argtypes = [fm_workbook_t_p]
    lib.fm_workbook_recalc.restype = fm_status_t

    lib.fm_workbook_set_iterative.argtypes = [
        fm_workbook_t_p,
        c_int32,
        c_int32,
        c_double,
    ]
    lib.fm_workbook_set_iterative.restype = fm_status_t

    # -- Diagnostics -------------------------------------------------------
    lib.fm_last_error_message.argtypes = []
    lib.fm_last_error_message.restype = c_char_p

    lib.fm_last_error_context.argtypes = []
    lib.fm_last_error_context.restype = c_char_p

    lib.fm_status_string.argtypes = [fm_status_t]
    lib.fm_status_string.restype = c_char_p

    # -- Version -----------------------------------------------------------
    lib.fm_version_string.argtypes = []
    lib.fm_version_string.restype = c_char_p


# Eagerly load + bind on import. Failure here is fatal: the package cannot
# function without libformulon.
LIB: CDLL = _load_library()
_setup_signatures(LIB)


# ---------------------------------------------------------------------------
# Small helpers used by the higher-level wrapper
# ---------------------------------------------------------------------------


def decode_cstr(ptr: Optional[bytes]) -> str:
    """Decode an optional C string (``c_char_p`` value) to ``str``.

    Returns the empty string when ``ptr`` is None. The C ABI guarantees
    that diagnostic / metadata pointers are non-NULL, but borrowed cell
    text pointers can theoretically be NULL for empty strings.
    """
    if ptr is None:
        return ""
    return ptr.decode("utf-8", errors="replace")


# Suppress the "unused import" warning when `os` / `sys` are only useful
# for downstream debugging.
_ = os, sys
