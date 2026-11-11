#!/usr/bin/env python3
"""Reject references to local-only material in tracked files.

The design corpus and the measurement notes live in gitignored directories, so
a tracked file that cites them points readers at something they cannot open.
This check scans every tracked text file for the citation shapes that cannot be
anything else:

  backup/<...>          a path into the gitignored corpus
  README §A, README M-3 a section or run label of a notes README
  round 5, rounds 1-3   a measurement batch number ("rounds 10^15" is a verb
                        and is not matched)

Bare run labels such as R20, X6 or D3 are not checked: they are also valid A1
references, and the test suites use those everywhere.

Stdlib only; run from the repository root.
"""

from __future__ import annotations

import re
import subprocess
import sys

PATTERNS: list[tuple[str, re.Pattern[str]]] = [
    ("gitignored corpus path", re.compile(r"\bbackup/")),
    ("notes README section or label", re.compile(r"\bREADME *(?:§|[A-Z]{1,2}-?[0-9])")),
    ("measurement batch number", re.compile(r"\bround [0-9]+\b|\brounds [0-9]+ *- *[0-9]+\b")),
]


def tracked_files() -> list[str]:
    out = subprocess.run(["git", "ls-files", "-z"], check=True, capture_output=True).stdout
    skip = {".gitignore", "tools/dev/check_internal_refs.py"}
    return [p for p in out.decode("utf-8").split("\0") if p and p not in skip]


def main() -> int:
    hits: list[str] = []
    for path in tracked_files():
        try:
            with open(path, encoding="utf-8") as fh:
                lines = fh.readlines()
        except (UnicodeDecodeError, FileNotFoundError, IsADirectoryError):
            continue
        for number, line in enumerate(lines, 1):
            for label, pattern in PATTERNS:
                if pattern.search(line):
                    hits.append(f"{path}:{number}: {label}: {line.strip()}")
    if hits:
        print("check_internal_refs: tracked files cite local-only material:")
        for hit in hits:
            print(f"  {hit}")
        print("State the rule or measurement itself instead of citing where it was recorded.")
        return 1
    print("check_internal_refs: ok")
    return 0


if __name__ == "__main__":
    sys.exit(main())
