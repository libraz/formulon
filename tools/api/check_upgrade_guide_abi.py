#!/usr/bin/env python3
"""Check that the upgrade guide names every C ABI break a published consumer can meet.

Every name in ``tools/dev/c_abi_breaks.txt`` that the last published release
shipped (``tools/api/released/<version>.txt``, the newest version) must appear
as a whole word in the guide of the version being built,
``docs/releases/upgrading-<major>.<minor>.md``, where the version is the one in
the top-level ``CMakeLists.txt``. A ledger name absent from that release's list
is one no published consumer compiled against, so the guide owes it nothing.

Exit 0 when every such name is found; otherwise print one name per line and
exit 1.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path
from typing import Iterable, List, Optional, Set

REPO_ROOT = Path(__file__).resolve().parents[2]
BREAKS = REPO_ROOT / "tools" / "dev" / "c_abi_breaks.txt"
RELEASED_DIR = REPO_ROOT / "tools" / "api" / "released"
RELEASES_DIR = REPO_ROOT / "docs" / "releases"
CMAKE_LISTS = REPO_ROOT / "CMakeLists.txt"


def ledger_names(text: str) -> List[str]:
    """Field 2 of each `<kind> <name> :: <reason>` line, in file order."""
    names = []
    for line in text.splitlines():
        fields = line.split()
        if len(fields) >= 2 and not line.startswith("#"):
            names.append(fields[1])
    return names


def project_guide(cmake_text: str) -> Path:
    """The upgrade guide of the version `project(... VERSION x.y.z ...)` declares."""
    match = re.search(r"project\(\s*\w+\s+VERSION\s+(\d+)\.(\d+)\.\d+", cmake_text)
    if match is None:
        raise ValueError("CMakeLists.txt declares no project version")
    return RELEASES_DIR / f"upgrading-{match.group(1)}.{match.group(2)}.md"


def latest_released() -> Path:
    """The snapshot of the newest published release."""
    snapshots = sorted(RELEASED_DIR.glob("*.txt"), key=lambda p: tuple(int(part) for part in p.stem.split(".")))
    if not snapshots:
        raise ValueError(f"no release snapshot under {RELEASED_DIR}")
    return snapshots[-1]


def snapshot_names(text: str) -> Set[str]:
    return {line for line in text.splitlines() if line and not line.startswith("#")}


def missing_names(ledger: Iterable[str], released: Set[str], guide: str) -> List[str]:
    """Ledger names the release shipped and the guide never mentions as a whole word."""
    words = set(re.findall(r"[A-Za-z_][A-Za-z0-9_]*", guide))
    return [name for name in ledger if name in released and name not in words]


def main(argv: Optional[List[str]] = None) -> int:
    del argv
    guide_path = project_guide(CMAKE_LISTS.read_text(encoding="utf-8"))
    missing = missing_names(
        ledger_names(BREAKS.read_text(encoding="utf-8")),
        snapshot_names(latest_released().read_text(encoding="utf-8")),
        guide_path.read_text(encoding="utf-8"),
    )
    for name in missing:
        print(f"{guide_path.relative_to(REPO_ROOT)}: no mention of {name}")
    return 1 if missing else 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
