#!/usr/bin/env python3
"""Reject gtest name generators whose lambda parameter is called `info`.

`INSTANTIATE_TEST_SUITE_P` expands its name generator inside a function that
already has a parameter named `info`, so a lambda parameter of the same name
shadows it. GCC reports that under `-Wshadow`, which the build turns into an
error; Clang's `-Wshadow` does not, so the defect only shows on GCC builds.
Name the parameter anything else (`param_info`), or use a named function.

Stdlib only; run from the repository root.
"""

from __future__ import annotations

import re
import subprocess
import sys

LAMBDA_INFO = re.compile(
    r"\[[^\]]*\]\s*\(\s*const\s+(?:::)?testing::TestParamInfo<[^)]*?>\s*&\s*info\s*\)",
    re.S,
)


def test_sources() -> list[str]:
    out = subprocess.run(["git", "ls-files", "-z", "tests"], check=True, capture_output=True).stdout
    return [p for p in out.decode("utf-8").split("\0") if p.endswith((".cpp", ".cc", ".h", ".hpp"))]


def main() -> int:
    hits: list[str] = []
    for path in test_sources():
        try:
            with open(path, encoding="utf-8") as fh:
                text = fh.read()
        except (UnicodeDecodeError, FileNotFoundError):
            continue
        for match in LAMBDA_INFO.finditer(text):
            line = text.count("\n", 0, match.start()) + 1
            hits.append(f"{path}:{line}")
    if hits:
        print("check_gtest_param_names: lambda name generators with a parameter named `info`:")
        for hit in hits:
            print(f"  {hit}")
        print("It shadows the macro's own `info` (GCC -Wshadow). Rename it, e.g. `param_info`.")
        return 1
    print("check_gtest_param_names: ok")
    return 0


if __name__ == "__main__":
    sys.exit(main())
