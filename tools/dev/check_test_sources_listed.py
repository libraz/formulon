#!/usr/bin/env python3
"""Reject a tracked `tests/**/*_test.cpp` that no CMakeLists.txt under tests/ lists.

A test source that drops out of its target's source list is neither compiled
nor run, and nothing else notices: the build stays green and the suite simply
shrinks. An edit to a shared list built from a stale base drops entries this
way. A file counts as listed when a non-comment line of a CMakeLists.txt under
tests/ names it relative to that file's directory, optionally behind
`${CMAKE_CURRENT_SOURCE_DIR}/`, or relative to the repository root behind
`${CMAKE_SOURCE_DIR}/`.

A file that is deliberately left unlisted goes in `UNLISTED` with the reason.

Stdlib only; run from the repository root.
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import PurePosixPath

# Test sources that are intentionally not compiled, each with the reason.
UNLISTED: dict[str, str] = {}

_SOURCE_TOKEN = re.compile(r"[\w./${}+-]+\.cpp")


def tracked(prefix: str) -> list[str]:
    out = subprocess.run(["git", "ls-files", "-z", prefix], check=True, capture_output=True).stdout
    return [p for p in out.decode("utf-8").split("\0") if p]


def listed_sources(cmake_files: list[str]) -> set[str]:
    listed: set[str] = set()
    for cmake in cmake_files:
        base = PurePosixPath(cmake).parent
        with open(cmake, encoding="utf-8") as fh:
            for line in fh:
                code = line.split("#", 1)[0]
                for token in _SOURCE_TOKEN.findall(code):
                    if token.startswith("${CMAKE_SOURCE_DIR}/"):
                        listed.add(token[len("${CMAKE_SOURCE_DIR}/") :])
                        continue
                    if token.startswith("${CMAKE_CURRENT_SOURCE_DIR}/"):
                        token = token[len("${CMAKE_CURRENT_SOURCE_DIR}/") :]
                    listed.add(str(base / token))
    return listed


def main() -> int:
    files = tracked("tests")
    tests = [p for p in files if p.endswith("_test.cpp")]
    listed = listed_sources([p for p in files if PurePosixPath(p).name == "CMakeLists.txt"])
    missing = [p for p in tests if p not in listed and p not in UNLISTED]
    stale = [p for p in UNLISTED if p not in tests or p in listed]
    if missing or stale:
        if missing:
            print("check_test_sources_listed: test sources no CMakeLists.txt under tests/ lists:")
            for path in missing:
                print(f"  {path}")
            print("Add each to its target's source list, or to UNLISTED with the reason it is not built.")
        if stale:
            print("check_test_sources_listed: UNLISTED entries that are missing or listed after all:")
            for path in stale:
                print(f"  {path}")
        return 1
    print(f"check_test_sources_listed: ok ({len(tests)} test sources listed)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
