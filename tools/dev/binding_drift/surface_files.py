"""Repo locations of the hand-synced surface files, and the shared reader and skip ledger."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import List, Set

REPO_ROOT = Path(__file__).resolve().parents[3]

CAPI_HEADER = REPO_ROOT / "src" / "c_api" / "formulon_c.h"
CAPI_EXPORTS = REPO_ROOT / "tools" / "wasm" / "capi_exports.txt"
PYTHON_PKG_DIR = REPO_ROOT / "packages" / "python" / "formulon"
PYTHON_C_BINDING = PYTHON_PKG_DIR / "_c.py"

WASM_BINDINGS_CPP = REPO_ROOT / "src" / "wasm" / "parts" / "bindings_register.cpp"
WASM_DTS = REPO_ROOT / "src" / "wasm" / "formulon.d.ts"

NODE_WORKBOOK_CLASS_CC = REPO_ROOT / "src" / "node_addon" / "parts" / "workbook_class.cc"
NODE_ADDON_CC = REPO_ROOT / "src" / "node_addon" / "addon.cc"
NODE_DTS = REPO_ROOT / "packages" / "npm-native" / "index.d.ts"
NODE_README = REPO_ROOT / "packages" / "npm-native" / "README.md"
NODE_INDEX_MJS = REPO_ROOT / "packages" / "npm-native" / "index.mjs"
NODE_DIST_DIR = REPO_ROOT / "packages" / "npm-native" / "dist"

C_ABI_BASELINE = Path(__file__).resolve().parents[1] / "c_abi_baseline.txt"
C_ABI_BREAKS = Path(__file__).resolve().parents[1] / "c_abi_breaks.txt"
# The npm package's constants and helpers live in common.mjs; each published
# entry point (single-threaded and pthread) re-exports it wholesale.
NPM_COMMON_MJS = REPO_ROOT / "packages" / "npm" / "common.mjs"
NPM_ENTRY_MJS = (
    REPO_ROOT / "packages" / "npm" / "index.mjs",
    REPO_ROOT / "packages" / "npm" / "threads.mjs",
)

ERROR_H = REPO_ROOT / "src" / "utils" / "error.h"
VALUE_H = REPO_ROOT / "src" / "value.h"
CF_MATCH_H = REPO_ROOT / "src" / "cf" / "cf_match.h"
CALC_MODE_H = REPO_ROOT / "src" / "calc_settings.h"
EXTERNAL_LINKS_H = REPO_ROOT / "src" / "external_link.h"
PYTHON_STRUCTS = PYTHON_PKG_DIR / "_structs.py"

# A sub-check appends here when an input it needs is legitimately absent, as
# opposed to wrong. `main` prints these alongside the verdict so "not checked"
# is visible in the output and can never be read as "checked and clean" -- a
# guard that quietly passes when its input is missing is not a guard.
_SKIPPED: List[str] = []


def _read(path: Path) -> str:
    if not path.is_file():
        print(f"check_binding_drift: missing file: {path}", file=sys.stderr)
        sys.exit(2)
    return path.read_text(encoding="utf-8")


def _read_bytes(path: Path) -> bytes:
    """`_read` for a comparison that must not normalise line endings."""
    if not path.is_file():
        print(f"check_binding_drift: missing file: {path}", file=sys.stderr)
        sys.exit(2)
    return path.read_bytes()


def _format_diff(label_a: str, only_a: Set[str], label_b: str, only_b: Set[str]) -> List[str]:
    problems = []
    if only_a:
        problems.append(f"  in {label_a} but not {label_b}: {sorted(only_a)}")
    if only_b:
        problems.append(f"  in {label_b} but not {label_a}: {sorted(only_b)}")
    return problems


# ---------------------------------------------------------------------------
# Check 6: public style-record field sets <-> their C ABI struct.
#
# `FontRecord` and friends are declared independently in three public
# surfaces and each one claims to mirror the C struct. Method-level drift
# checking cannot see a field that one surface forgot, which is how a font's
# `vertAlign` reached the WASM `.d.ts` alone: reading a superscript font and
# writing it back through the other bindings silently demoted it.
# ---------------------------------------------------------------------------

PYTHON_WORKBOOK = PYTHON_PKG_DIR / "workbook.py"
