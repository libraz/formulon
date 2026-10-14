"""README fact checks: every hand-written fact in the five READMEs against its source."""

from __future__ import annotations

import os
import re
import subprocess
import sys
from typing import Callable, Dict, List, Set, Tuple

from .surface_files import REPO_ROOT, _read

_ROOT_EN = REPO_ROOT / "README.md"
_ROOT_JA = REPO_ROOT / "README_ja.md"
_NPM = REPO_ROOT / "packages" / "npm" / "README.md"
_PYTHON = REPO_ROOT / "packages" / "python" / "README.md"
_NATIVE = REPO_ROOT / "packages" / "npm-native" / "README.md"

_EXCEL_PROFILE_H = REPO_ROOT / "src" / "excel_profile.h"
_SIZE_REPORT = REPO_ROOT / "tools" / "bench" / "wasm_size_report.sh"
_STATUS_PY = REPO_ROOT / "tools" / "catalog" / "status.py"
_RELEASE_YML = REPO_ROOT / ".github" / "workflows" / "release.yml"
_NPM_PACKAGE = REPO_ROOT / "packages" / "npm" / "package.json"
_PYPROJECT = REPO_ROOT / "packages" / "python" / "pyproject.toml"

_NUMBER_WORDS = {"twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15, "sixteen": 16}


def _function_counts() -> Tuple[str, str]:
    env = dict(os.environ, NO_COLOR="1")
    out = subprocess.run(
        [sys.executable, str(_STATUS_PY)], capture_output=True, text=True, check=True, cwd=REPO_ROOT, env=env
    ).stdout
    total = re.search(r"recognition: \d+ / (\d+) catalogued", out)
    real = re.search(r"Availability: (\d+)/\d+ real implementations", out)
    if total is None or real is None:
        raise RuntimeError(f"unrecognised {_STATUS_PY.name} header")
    return total.group(1), real.group(1)


def _profile_ids() -> Set[str]:
    table = _read(_EXCEL_PROFILE_H).split("kExcelProfileIds", 1)[1].split("};", 1)[0]
    return set(re.findall(r'"((?:mac|win)-365-[a-z]{2}_[A-Z]{2})"', table))


def _default_profile_id() -> str:
    text = _read(_EXCEL_PROFILE_H)
    fn = re.search(r"default_excel_profile\(\) noexcept \{\s*return (\w+)\(\);", text)
    if fn is None:
        raise RuntimeError("default_excel_profile() body not recognised")
    entry = re.search(rf'\{{{fn.group(1)}\(\), "([^"]+)"\}}', text)
    if entry is None:
        raise RuntimeError(f"{fn.group(1)}() missing from kExcelProfileIds")
    return entry.group(1)


def _size_ceilings() -> Tuple[str, str]:
    text = _read(_SIZE_REPORT)
    raw = int(re.search(r"^HARD_CEILING=(\d+)", text, re.M).group(1))
    brotli = int(re.search(r"^BROTLI_HARD_CEILING=(\d+)", text, re.M).group(1))
    return f"{raw / 1048576:g}", f"{brotli // 1024}"


def _cli_platforms() -> Set[str]:
    return set(re.findall(r"- platform-arch: (\S+)", _read(_RELEASE_YML)))


def _node_min() -> str:
    return re.search(r'"node": ">=(\d+)"', _read(_NPM_PACKAGE)).group(1)


def _python_min() -> str:
    return re.search(r'requires-python = ">=(3\.\d+)"', _read(_PYPROJECT)).group(1)


def _expand_profile_braces(match: re.Match) -> Set[str]:
    return {f"{h}-365-{loc.strip()}" for h in ("mac", "win") for loc in match.group(1).split(",")}


# (fact, readmes, pattern, extract-the-claim, source-of-truth label)
_Fact = Tuple[str, Tuple, re.Pattern, Callable[[re.Match], object], str]

_FACTS: List[_Fact] = [
    (
        "function-counts",
        (_ROOT_EN, _NPM, _PYTHON),
        re.compile(r"(\d+) Excel function names, (\d+) implemented locally"),
        lambda m: (m.group(1), m.group(2)),
        "tools/catalog/status.py",
    ),
    (
        "function-counts",
        (_ROOT_JA,),
        re.compile(r"Excel 関数名 (\d+) 件を認識し、(\d+) 件"),
        lambda m: (m.group(1), m.group(2)),
        "tools/catalog/status.py",
    ),
    (
        "size-ceilings",
        (_ROOT_EN,),
        re.compile(r"above ([\d.]+) MiB uncompressed or (\d+) KiB Brotli"),
        lambda m: (m.group(1), m.group(2)),
        "tools/bench/wasm_size_report.sh",
    ),
    (
        "size-ceilings",
        (_ROOT_JA,),
        re.compile(r"非圧縮 ([\d.]+) MiB、Brotli (\d+) KiB"),
        lambda m: (m.group(1), m.group(2)),
        "tools/bench/wasm_size_report.sh",
    ),
    (
        "profile-count",
        (_ROOT_EN, _NPM, _PYTHON),
        re.compile(r"(\d+) Excel behavior profiles across (\w+) locales"),
        lambda m: (int(m.group(1)), _NUMBER_WORDS.get(m.group(2))),
        "src/excel_profile.h kExcelProfileIds",
    ),
    (
        "profile-count",
        (_ROOT_JA,),
        re.compile(r"(\d+) ロケール × Mac / Windows ホストの、(\d+) 種類"),
        lambda m: (int(m.group(2)), int(m.group(1))),
        "src/excel_profile.h kExcelProfileIds",
    ),
    (
        "profile-count",
        (_NATIVE,),
        re.compile(r"any of the (\d+)"),
        lambda m: (int(m.group(1)), None),
        "src/excel_profile.h kExcelProfileIds",
    ),
    (
        "profile-ids",
        (_ROOT_EN, _ROOT_JA, _NPM, _PYTHON, _NATIVE),
        re.compile(r"`\{mac,win\}-365-\{([^}]*)\}`"),
        _expand_profile_braces,
        "src/excel_profile.h kExcelProfileIds",
    ),
    (
        "default-profile",
        (_ROOT_EN, _NPM, _PYTHON),
        re.compile(r"New workbooks use `([^`]+)`"),
        lambda m: m.group(1),
        "src/excel_profile.h default_excel_profile()",
    ),
    (
        "default-profile",
        (_ROOT_JA,),
        re.compile(r"新しいワークブックの既定は `([^`]+)`"),
        lambda m: m.group(1),
        "src/excel_profile.h default_excel_profile()",
    ),
    (
        "default-profile",
        (_NATIVE,),
        re.compile(r"default `([^`]+)` profile"),
        lambda m: m.group(1),
        "src/excel_profile.h default_excel_profile()",
    ),
    (
        "node-min",
        (_ROOT_EN, _ROOT_JA, _NPM),
        re.compile(r"Node (\d+)\+"),
        lambda m: m.group(1),
        "packages/npm/package.json engines",
    ),
    (
        "python-min",
        (_ROOT_EN, _ROOT_JA, _PYTHON),
        re.compile(r"Python (3\.\d+)\+"),
        lambda m: m.group(1),
        "packages/python/pyproject.toml requires-python",
    ),
    (
        "cli-platforms",
        (_ROOT_EN, _ROOT_JA),
        re.compile(r"((?:`[a-z]+-[a-z0-9]+`(?:, | and |、))+`[a-z]+-[a-z0-9]+`)"),
        lambda m: set(re.findall(r"`([a-z]+-[a-z0-9]+)`", m.group(1))),
        ".github/workflows/release.yml build-cli matrix",
    ),
]


def check_readme_facts() -> List[str]:
    ids = _profile_ids()
    locales = {i.split("-", 2)[2] for i in ids}
    truth: Dict[str, object] = {
        "function-counts": _function_counts(),
        "size-ceilings": _size_ceilings(),
        "profile-ids": ids,
        "default-profile": _default_profile_id(),
        "node-min": _node_min(),
        "python-min": _python_min(),
        "cli-platforms": _cli_platforms(),
    }
    problems: List[str] = []
    for fact, readmes, pattern, extract, source in _FACTS:
        for path in readmes:
            rel = path.relative_to(REPO_ROOT)
            match = pattern.search(_read(path))
            if match is None:
                problems.append(f"readme-facts: {rel}: {fact} sentence not found (pattern: {pattern.pattern!r})")
                continue
            claim = extract(match)
            if fact == "profile-count":
                count, n_locales = claim
                ok = count == len(ids) and n_locales in (None, len(locales))
                expected: object = f"{len(ids)} profiles / {len(locales)} locales"
            else:
                expected = truth[fact]
                ok = claim == expected
            if not ok:
                problems.append(f"readme-facts: {rel}: {fact} says {claim!r}, {source} says {expected!r}")
    return problems
