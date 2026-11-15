#!/usr/bin/env python3
"""Check that the released C ABI baseline is the newest published release.

``tools/dev/c_abi_baseline.txt`` names the tag it was generated from on its
first line; it must be the newest snapshot under ``tools/api/released/``. A
release that skips rebasing the baseline then fails the next build.

Exit 0 when the two agree; otherwise print both versions and exit 1.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import List, Optional

REPO_ROOT = Path(__file__).resolve().parents[2]
BASELINE = REPO_ROOT / "tools" / "dev" / "c_abi_baseline.txt"
RELEASED_DIR = REPO_ROOT / "tools" / "api" / "released"


def baseline_version(text: str) -> str:
    """The version in the baseline's first line, `# ... as shipped in vX.Y.Z.`."""
    first = text.splitlines()[0] if text else ""
    match = re.search(r"\bv(\d+\.\d+\.\d+)\b", first)
    if match is None:
        raise ValueError("the baseline's first line names no release tag")
    return match.group(1)


def newest_release(names: List[str]) -> str:
    """The highest `X.Y.Z` among snapshot file stems."""
    if not names:
        raise ValueError("no release snapshot")
    return max(names, key=lambda n: tuple(int(part) for part in n.split(".")))


def main(argv: Optional[List[str]] = None) -> int:
    del argv
    baseline = baseline_version(BASELINE.read_text(encoding="utf-8"))
    newest = newest_release([p.stem for p in RELEASED_DIR.glob("*.txt")])
    if baseline != newest:
        print(f"tools/dev/c_abi_baseline.txt is v{baseline}; the newest published release is v{newest}.")
        print("Rebase it with tools/dev/gen_c_abi_baseline.py v" + newest + " and update c_abi_breaks.txt.")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
