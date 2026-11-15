"""Relative named imports in the JS test trees resolve to names their target module exports.

The api_*.test.mjs files are shared verbatim between the WASM and native
packages and import from each package's own smoke_support.mjs, so a helper
added to one support module and not the other leaves the other package's
copy dead at import time. Node reports that as one failed file, and the npm
suites are not part of the per-commit gate, so this static check stands in.
"""

from __future__ import annotations

import re
from pathlib import Path
from typing import Dict, List, Optional, Set

from binding_drift.surface_files import REPO_ROOT

TEST_DIRS = (
    REPO_ROOT / "packages" / "npm" / "test",
    REPO_ROOT / "packages" / "npm-native" / "test",
    REPO_ROOT / "tests" / "wasm",
)

_COMMENT = re.compile(r"/\*.*?\*/|//[^\n]*", re.S)
_NAMED_IMPORT = re.compile(r"\bimport\s*\{([^}]*)\}\s*from\s*'(\.{1,2}/[^']+\.mjs)'")
_EXPORT_LIST = re.compile(r"\bexport\s*\{([^}]*)\}(?!\s*from)")
_EXPORT_DECL = re.compile(r"\bexport\s+(?:async\s+)?(?:function\*?|const|let|var|class)\s+([A-Za-z_$][\w$]*)")
_EXPORT_STAR = re.compile(r"\bexport\s*\*")


def _source(path: Path) -> str:
    return _COMMENT.sub("", path.read_text(encoding="utf-8"))


def _names(clause: str, *, exported: bool) -> List[str]:
    """Names in an import or export clause; `a as b` counts as `b` on export, `a` on import."""
    out = []
    for part in clause.split(","):
        words = part.split()
        if not words:
            continue
        out.append(words[-1] if exported else words[0])
    return out


def _exports(path: Path) -> Optional[Set[str]]:
    """The module's export names, or None when a re-export makes them unknowable statically."""
    text = _source(path)
    if _EXPORT_STAR.search(text):
        return None
    names = set(_EXPORT_DECL.findall(text))
    for clause in _EXPORT_LIST.findall(text):
        names.update(_names(clause, exported=True))
    return names


def check_test_imports() -> List[str]:
    problems: List[str] = []
    cache: Dict[Path, Optional[Set[str]]] = {}
    for test_dir in TEST_DIRS:
        for module in sorted(test_dir.rglob("*.mjs")):
            for clause, spec in _NAMED_IMPORT.findall(_source(module)):
                target = (module.parent / spec).resolve()
                rel = module.relative_to(REPO_ROOT)
                # A staged copy under dist/ is a build output; staged-dist compares it with its source.
                if "dist" in target.relative_to(REPO_ROOT).parts:
                    continue
                if not target.is_file():
                    problems.append(f"{rel}: imports from missing module {spec}")
                    continue
                if target not in cache:
                    cache[target] = _exports(target)
                exported = cache[target]
                if exported is None:
                    continue
                missing = [n for n in _names(clause, exported=False) if n not in exported]
                if missing:
                    problems.append(f"{rel}: {spec} does not export {', '.join(missing)}")
    return problems
