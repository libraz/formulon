"""Each sheet-scoped C function states in its body the sheet kinds its ledger entry says it serves.

`tools/api/names.yaml` gives every `scope: sheet` op with a C function a `sheet_kinds` value: `all` (no guard),
`worksheet_or_chartsheet` or `worksheet`. The body of the function under src/c_api/parts and src/c_api/generated
must call `check_sheet_use(wb, sheet_index, SheetUse::<kind>)` with the matching enumerator, and an `all` function
must not call it, so a hand-written entry cannot drift from its ledger value without running it.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional

from .api_conventions import load_ledger
from .surface_files import REPO_ROOT

_C_API_DIR = REPO_ROOT / "src" / "c_api"
_BODY_DIRS = ("parts", "generated")
_USE_ENUMERATORS = {"worksheet": "kWorksheet", "worksheet_or_chartsheet": "kWorksheetOrChartsheet"}
_USE_CALL = re.compile(r"\bcheck_sheet_use\(\s*wb\s*,\s*sheet_index\s*,\s*SheetUse::(\w+)\s*\)")


def function_bodies(c_api_dir: Path = _C_API_DIR) -> Dict[str, str]:
    """The text of each `extern "C" fm_status_t <name>(...) { ... }` under the parts and generated directories."""
    bodies: Dict[str, str] = {}
    for part in _BODY_DIRS:
        for path in sorted((c_api_dir / part).glob("*.cpp")):
            text = path.read_text(encoding="utf-8", errors="replace")
            for match in re.finditer(r'^extern "C" [\w:]+ (fm_[a-z0-9_]+)\(', text, re.M):
                end = text.find("\n}\n", match.end())
                bodies[match.group(1)] = text[match.start() : end if end >= 0 else len(text)]
    return bodies


def check_sheet_kind_guard(c_api_dir: Path = _C_API_DIR, entries: Optional[Dict[str, dict]] = None) -> List[str]:
    """One problem per sheet-scoped C function whose guard call disagrees with its ledger `sheet_kinds`."""
    if entries is None:
        entries = load_ledger().entries
    bodies = function_bodies(c_api_dir)
    problems: List[str] = []
    for key, entry in entries.items():
        if not isinstance(entry, dict) or entry.get("scope") != "sheet" or entry.get("kind") != "op":
            continue
        name = entry.get("c")
        if not isinstance(name, str) or (entry.get("drop") or {}).get("c") == name:
            continue
        kinds = entry.get("sheet_kinds")
        if kinds != "all" and kinds not in _USE_ENUMERATORS:
            problems.append(f"sheet-kind-guard: {key} ({name}) has no valid `sheet_kinds` in the ledger")
            continue
        body = bodies.get(name)
        if body is None:
            problems.append(f"sheet-kind-guard: {name} has no body under src/c_api/parts or src/c_api/generated")
            continue
        used = _USE_CALL.findall(body)
        if kinds == "all":
            if used:
                problems.append(f"sheet-kind-guard: {name} is `all` in the ledger but calls check_sheet_use")
        elif used != [_USE_ENUMERATORS[kinds]]:
            problems.append(
                f"sheet-kind-guard: {name} is `{kinds}` in the ledger; its body calls check_sheet_use with "
                f"{used or 'nothing'}, not SheetUse::{_USE_ENUMERATORS[kinds]}"
            )
    return problems
