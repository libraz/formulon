#!/usr/bin/env python3
"""Compare a list of failed ctest test names against an allowlist.

Used by the native CI job to keep the build green for the small set of
pre-existing flakes (see ``tools/ci/expected_flakes.txt``) without hiding
new regressions.

Usage:
    check_test_failures.py --failed PATH --allowlist PATH

Exit codes:
    0 -- every failed test is in the allowlist (or no failures at all)
    1 -- at least one failed test is not in the allowlist
    2 -- argument or I/O error

The ``--failed`` file is one test name per line. ``ctest`` writes
``Testing/Temporary/LastTestsFailed.log`` with the format
``<index>:<name>``; both the bare name and the indexed form are accepted.

The ``--allowlist`` file ignores blank lines and lines starting with ``#``.
Inline ``# rationale`` comments after a name are stripped before matching.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path


def _strip_comment(line: str) -> str:
    """Return the part of ``line`` before any ``#`` comment, trimmed."""
    hash_at = line.find("#")
    if hash_at >= 0:
        line = line[:hash_at]
    return line.strip()


def _read_allowlist(path: Path) -> set[str]:
    names: set[str] = set()
    for raw in path.read_text(encoding="utf-8").splitlines():
        stripped = _strip_comment(raw)
        if stripped:
            names.add(stripped)
    return names


def _read_failed(path: Path) -> list[str]:
    failed: list[str] = []
    for raw in path.read_text(encoding="utf-8").splitlines():
        stripped = raw.strip()
        if not stripped:
            continue
        # ctest LastTestsFailed.log uses "<index>:<name>"; accept both forms.
        if ":" in stripped and stripped.split(":", 1)[0].isdigit():
            stripped = stripped.split(":", 1)[1]
        failed.append(stripped)
    return failed


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--failed", required=True, type=Path, help="path to a file listing failed test names")
    parser.add_argument("--allowlist", required=True, type=Path, help="path to the expected-flake allowlist")
    args = parser.parse_args(argv)

    if not args.failed.is_file():
        print(f"check_test_failures: --failed not found: {args.failed}", file=sys.stderr)
        return 2
    if not args.allowlist.is_file():
        print(f"check_test_failures: --allowlist not found: {args.allowlist}", file=sys.stderr)
        return 2

    allowlist = _read_allowlist(args.allowlist)
    failed = _read_failed(args.failed)

    if not failed:
        print("check_test_failures: no failures recorded; ok")
        return 0

    unexpected = [name for name in failed if name not in allowlist]
    expected_hits = [name for name in failed if name in allowlist]

    if expected_hits:
        print("check_test_failures: allowlisted failures (debt):")
        for name in expected_hits:
            print(f"  - {name}")
    if unexpected:
        print("check_test_failures: NEW failures (not in allowlist):", file=sys.stderr)
        for name in unexpected:
            print(f"  - {name}", file=sys.stderr)
        return 1
    print("check_test_failures: all failures allowlisted; ok")
    return 0


if __name__ == "__main__":
    sys.exit(main())
