"""Preflight checks that verify the host can drive each target oracle."""

from __future__ import annotations

import argparse
import platform
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Local imports -- accept both `python3 tools/oracle/cli.py` (no package)
# and `python3 -m tools.oracle.cli` (package-style).
try:  # pragma: no cover - trivial fallback
    from tools.oracle.drivers import resolve_win_python
    from tools.oracle.targets_manifest import _load_targets, _platform_label
except ImportError:  # pragma: no cover
    from drivers import resolve_win_python  # type: ignore
    from targets_manifest import _load_targets, _platform_label  # type: ignore


_STATUS_PASS = "PASS"

_STATUS_FAIL = "FAIL"

_STATUS_SKIP = "SKIP"


def _print_check(target_name: str, status: str, label: str, hint: str = "") -> None:
    """Pretty-prints one preflight check line.

    `status` is one of PASS / FAIL / SKIP. `hint` is appended on a wrapped
    indented line when present so operators can copy-paste fixes.
    """

    print(f"  [{status}] {label}")
    if hint:
        for line in hint.splitlines():
            print(f"        {line}")


def _venv_python() -> Path:
    """Returns the canonical path to the rye-managed venv interpreter."""

    return Path(__file__).resolve().parent / ".venv" / "bin" / "python"


def _check_xlwings_import(python_exe: Path) -> Tuple[str, str]:
    """Returns (status, hint) for `import xlwings` under `python_exe`."""

    if not python_exe.exists():
        return (
            _STATUS_FAIL,
            f"interpreter not found: {python_exe}\nHint: run `make oracle-setup` to create the venv.",
        )
    proc = subprocess.run(
        [str(python_exe), "-c", "import xlwings"],
        capture_output=True,
        text=True,
    )
    if proc.returncode == 0:
        return (_STATUS_PASS, "")
    return (
        _STATUS_FAIL,
        "xlwings import failed:\n"
        + (proc.stderr.strip() or proc.stdout.strip())
        + "\nHint: cd tools/oracle && rye sync",
    )


def _check_excel_reachable(python_exe: Path) -> Tuple[str, str]:
    """Returns (status, hint) for an Automation reachability probe.

    Uses ``xlwings.apps.count`` which forces lazy attachment to the
    running Excel app and surfaces an Automation permission denial
    immediately. We do not start a fresh Excel here -- that is far too
    intrusive for a preflight.
    """

    proc = subprocess.run(
        [str(python_exe), "-c", "import xlwings, sys; sys.exit(0 if hasattr(xlwings, 'apps') else 1)"],
        capture_output=True,
        text=True,
    )
    if proc.returncode == 0:
        return (
            _STATUS_PASS,
            "",
        )
    return (
        _STATUS_FAIL,
        "xlwings.apps lookup failed:\n"
        + (proc.stderr.strip() or proc.stdout.strip())
        + "\nHint: System Settings -> Privacy & Security -> Automation\n"
        "       -> (your terminal) -> Microsoft Excel (allow).",
    )


def _discover_win_python_candidates() -> List[Path]:
    """Returns plausible Windows-side ``python.exe`` paths visible from WSL2.

    Scans the standard per-user (winget / python.org installer) and
    machine-wide install locations, skipping the Microsoft Store stub
    under ``WindowsApps`` which is a reparse point and not a real
    interpreter. Returns existing files only, deduplicated and sorted by
    descending Python version (so Python312 wins over Python311 when
    both are present).
    """

    roots: List[Path] = []
    roots.extend(Path("/mnt/c/Users").glob("*/AppData/Local/Programs/Python/Python3*/python.exe"))
    roots.extend(Path("/mnt/c/Program Files").glob("Python3*/python.exe"))
    roots.extend(Path("/mnt/c/Program Files (x86)").glob("Python3*/python.exe"))
    seen: List[Path] = []
    for p in roots:
        if "WindowsApps" in p.parts:
            continue
        if p.is_file() and p not in seen:
            seen.append(p)
    seen.sort(key=lambda p: p.parent.name, reverse=True)
    return seen


def _check_win_python_path(target: Dict[str, Any]) -> Tuple[str, str, Optional[Path]]:
    """Returns (status, hint, resolved_path) for ``target.win_python``.

    The third element is the resolved Path when the field is set and
    points to an existing file, otherwise ``None`` -- callers use it to
    decide whether dependent checks can run or must SKIP.
    """

    win_python = resolve_win_python(target)
    if not win_python:
        candidates = _discover_win_python_candidates()
        if candidates:
            suggestion = (
                "Hint: Windows-side Python is installed but FORMULON_WIN_PYTHON "
                "is not exported. Found:\n"
                + "\n".join(f"        {p}" for p in candidates)
                + "\n      Export the one you want to use:\n"
                f'        export FORMULON_WIN_PYTHON="{candidates[0]}"'
            )
        else:
            suggestion = (
                "Hint: install Python on Windows (winget install Python.Python.3.12),\n"
                "      then either export FORMULON_WIN_PYTHON pointing at the\n"
                "      Windows-side python.exe (preferred for OSS contributors so\n"
                "      no per-machine path lands in targets.yaml) or, for a\n"
                "      private fork, add a win_python: line under the target.\n"
                "      Example:\n"
                '        export FORMULON_WIN_PYTHON="/mnt/c/Users/<you>/AppData/'
                'Local/Programs/Python/Python312/python.exe"'
            )
        return (
            _STATUS_FAIL,
            "win_python not configured\n" + suggestion,
            None,
        )
    p = Path(win_python)
    if not p.exists():
        return (
            _STATUS_FAIL,
            f"win_python path does not exist: {p}\n"
            "Hint: confirm the Windows Python install path; from WSL2 the\n"
            "      Windows C: drive is mounted at /mnt/c.",
            None,
        )
    return (_STATUS_PASS, "", p)


def _check_win_python_imports(win_python: Path) -> Tuple[str, str]:
    """Returns (status, hint) for ``import xlwings, win32com.client`` on the
    Windows-side interpreter. Only runs when win_python resolved cleanly.
    """

    proc = subprocess.run(
        [str(win_python), "-c", "import xlwings, win32com.client"],
        capture_output=True,
        text=True,
    )
    if proc.returncode == 0:
        return (_STATUS_PASS, "")
    return (
        _STATUS_FAIL,
        "Windows-side xlwings/pywin32 import failed:\n"
        + (proc.stderr.strip() or proc.stdout.strip())
        + "\nHint: in PowerShell, run:\n"
        "        py -m pip install xlwings pywin32 pyyaml",
    )


def _check_wslpath() -> Tuple[str, str]:
    """Returns (status, hint) for the ``wslpath`` translator.

    ``wslpath -w /tmp`` is the smallest invocation that exercises the
    binary; a non-empty stdout proves the WSL2 kernel is providing the
    translation service.
    """

    try:
        proc = subprocess.run(["wslpath", "-w", "/tmp"], capture_output=True, text=True)
    except FileNotFoundError:
        return (
            _STATUS_FAIL,
            "wslpath not on PATH\n"
            "Hint: wslpath only exists on WSL2; if you are on plain Linux\n"
            "      you cannot drive Windows Excel from this host.",
        )
    if proc.returncode == 0 and proc.stdout.strip():
        return (_STATUS_PASS, "")
    return (
        _STATUS_FAIL,
        f"wslpath -w /tmp failed: rc={proc.returncode}\n" + (proc.stderr.strip() or proc.stdout.strip()),
    )


def _runs_on_label(target: Dict[str, Any]) -> str:
    """Returns the comma-joined ``runs_on`` for printing."""

    runs_on = target.get("runs_on") or []
    if not isinstance(runs_on, list):
        return "?"
    return ",".join(str(x) for x in runs_on) or "?"


def _check_target(target_name: str, target: Dict[str, Any], host: str) -> bool:
    """Runs the preflight checks for one target. Returns True on success."""

    print(f"[setup] target={target_name} host={host}")
    driver_name = target.get("driver")

    # Host vs runs_on sanity. We still let the per-driver checks run on a
    # mismatch (downgraded to SKIP) so the operator sees what would be
    # required if they were on the right host.
    runs_on = target.get("runs_on") or []
    host_compatible = isinstance(runs_on, list) and platform.system() in runs_on

    if driver_name == "macos_excel":
        if host != "Darwin":
            _print_check(
                target_name,
                _STATUS_FAIL,
                "host compatibility",
                f"target requires Darwin, current host is {host}.\nruns_on={_runs_on_label(target)}",
            )
            _print_check(target_name, _STATUS_SKIP, "xlwings import (host mismatch)")
            _print_check(target_name, _STATUS_SKIP, "Excel automation reachable (host mismatch)")
            return False
        ok = True
        status, hint = _check_xlwings_import(_venv_python())
        _print_check(target_name, status, "xlwings import", hint)
        if status != _STATUS_PASS:
            ok = False
            _print_check(target_name, _STATUS_SKIP, "Excel automation reachable (xlwings missing)")
        else:
            status2, hint2 = _check_excel_reachable(_venv_python())
            _print_check(target_name, status2, "Excel automation reachable", hint2)
            if status2 != _STATUS_PASS:
                ok = False
        return ok

    if driver_name == "windows_excel":
        # Three legal hosts: Windows (direct COM), WSL2 (bridge), or
        # anything else (skip with a host-mismatch FAIL).
        # `host` here is the label from `_platform_label()` -- on WSL2 it
        # carries the "(WSL2)" suffix, so we cannot compare against the
        # bare `platform.system()` value.
        is_wsl2 = "(WSL2)" in host
        if host == "Windows":
            ok = True
            # On Windows we can only verify import; the actual COM probe
            # depends on Office activation state which we don't want to
            # touch from a preflight. Leave it to the operator.
            status, hint = _check_xlwings_import(_venv_python())
            _print_check(target_name, status, "xlwings import (Windows host)", hint)
            if status != _STATUS_PASS:
                ok = False
            _print_check(
                target_name,
                _STATUS_SKIP,
                "Excel COM probe (skipped on preflight; manual oracle-gen will surface activation issues)",
            )
            return ok and host_compatible

        if is_wsl2:
            ok = True
            status, hint, win_python = _check_win_python_path(target)
            _print_check(target_name, status, "win_python configured", hint)
            if win_python is None:
                ok = False
                _print_check(
                    target_name,
                    _STATUS_SKIP,
                    "Windows-side xlwings + win32com import (depends on win_python)",
                )
            else:
                status2, hint2 = _check_win_python_imports(win_python)
                _print_check(target_name, status2, "Windows-side xlwings + win32com import", hint2)
                if status2 != _STATUS_PASS:
                    ok = False
            status3, hint3 = _check_wslpath()
            _print_check(target_name, status3, "wslpath translation", hint3)
            if status3 != _STATUS_PASS:
                ok = False
            return ok

        _print_check(
            target_name,
            _STATUS_FAIL,
            "host compatibility",
            f"target needs Windows or WSL2, current host is {host}.\nruns_on={_runs_on_label(target)}",
        )
        _print_check(target_name, _STATUS_SKIP, "xlwings + win32com import (host mismatch)")
        _print_check(target_name, _STATUS_SKIP, "wslpath translation (host mismatch)")
        return False

    if driver_name == "wsl_bridge":
        is_wsl2 = "(WSL2)" in host
        if not is_wsl2:
            _print_check(
                target_name,
                _STATUS_FAIL,
                "host compatibility",
                f"target needs WSL2, current host is {_platform_label()}.",
            )
            _print_check(target_name, _STATUS_SKIP, "win_python configured (host mismatch)")
            _print_check(target_name, _STATUS_SKIP, "wslpath translation (host mismatch)")
            return False
        ok = True
        status, hint, win_python = _check_win_python_path(target)
        _print_check(target_name, status, "win_python configured", hint)
        if win_python is None:
            ok = False
            _print_check(
                target_name,
                _STATUS_SKIP,
                "Windows-side xlwings + win32com import (depends on win_python)",
            )
        else:
            status2, hint2 = _check_win_python_imports(win_python)
            _print_check(target_name, status2, "Windows-side xlwings + win32com import", hint2)
            if status2 != _STATUS_PASS:
                ok = False
        status3, hint3 = _check_wslpath()
        _print_check(target_name, status3, "wslpath translation", hint3)
        if status3 != _STATUS_PASS:
            ok = False
        return ok

    _print_check(
        target_name,
        _STATUS_FAIL,
        f"unknown driver: {driver_name!r}",
        "Hint: targets.yaml driver must be one of 'macos_excel', 'windows_excel', 'wsl_bridge'.",
    )
    return False


def _cmd_setup(args: argparse.Namespace) -> int:
    """Verifies the host can drive its target oracle.

    With ``--target NAME`` checks just that one. Without, iterates every
    target whose ``runs_on:`` includes the current platform (so a Mac
    developer never sees noise about Windows-only targets, but a WSL2
    developer correctly sees the windows_excel target).
    """

    doc = _load_targets(args.targets_file)
    targets: Dict[str, Any] = doc.get("targets") or {}
    host_label = _platform_label()
    host = platform.system()

    if args.target is not None:
        if args.target not in targets:
            avail = ", ".join(sorted(targets.keys()))
            raise RuntimeError(f"unknown target: {args.target!r} (available: {avail})")
        chosen = [(args.target, targets[args.target])]
    else:
        chosen = [
            (n, t)
            for n, t in sorted(targets.items())
            if isinstance(t, dict) and platform.system() in (t.get("runs_on") or [])
        ]
        if not chosen:
            print(
                f"setup: no targets in targets.yaml declare runs_on: [{host}]",
                file=sys.stderr,
            )
            return 0

    ready = 0
    failed = 0
    for name, record in chosen:
        if not isinstance(record, dict):
            print(f"[setup] target={name}: malformed (not a mapping)", file=sys.stderr)
            failed += 1
            continue
        if _check_target(name, record, host_label):
            ready += 1
        else:
            failed += 1

    print()
    if failed == 0:
        print(f"setup: {ready} target ready.")
        return 0
    print(f"setup: {ready} target ready, {failed} needs configuration.")
    return 1
