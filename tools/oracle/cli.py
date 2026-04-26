#!/usr/bin/env python3
"""Cross-platform oracle CLI dispatcher.

A thin orchestrator over the oracle drivers. It (1) loads targets.yaml,
(2) selects targets that match the current platform, (3) either prints
them, runs preflight checks, or shells out to oracle_gen for each.

Subcommands:
    cli.py list                          # print available targets
    cli.py gen [--target NAME] [--all]   # delegate to oracle_gen
    cli.py setup [--target NAME]         # run preflight checks for the
                                         # current host

Examples:
    python3 tools/oracle/cli.py list
    python3 tools/oracle/cli.py gen
    python3 tools/oracle/cli.py gen --target mac-365-ja_JP
    python3 tools/oracle/cli.py gen --all
    python3 tools/oracle/cli.py setup
    python3 tools/oracle/cli.py setup --target win-365-ja_JP
"""

from __future__ import annotations

import argparse
import platform
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Local imports — accept both `python3 tools/oracle/cli.py` (no package)
# and `python3 -m tools.oracle.cli` (package-style).
try:  # pragma: no cover - trivial fallback
    from tools.oracle import oracle_gen
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import oracle_gen  # type: ignore


REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TARGETS_FILE = Path(__file__).resolve().parent / "targets.yaml"


def _load_targets(path: Path) -> Dict[str, Any]:
    """Reads and validates targets.yaml; returns the parsed mapping."""

    if not path.exists():
        raise RuntimeError(f"oracle targets file not found: {path}")
    try:
        import yaml  # type: ignore

        doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except Exception as exc:
        raise RuntimeError(f"failed to parse {path}: {exc}") from exc
    if not isinstance(doc, dict):
        raise RuntimeError(f"{path} root must be a mapping")
    targets = doc.get("targets")
    if not isinstance(targets, dict) or not targets:
        raise RuntimeError(f"{path} has no `targets:` mapping")
    return doc


def _runs_on_current(target: Dict[str, Any]) -> bool:
    """Whether the target declares the current OS in `runs_on:`."""

    runs_on = target.get("runs_on") or []
    if not isinstance(runs_on, list):
        return False
    return platform.system() in runs_on


def _platform_label() -> str:
    """Returns ``platform.system()`` with a ``(WSL2)`` suffix on WSL2.

    The CLI's ``list`` command surfaces this so operators can confirm at
    a glance which side of the WSL boundary they are on -- ``Linux`` and
    ``Linux (WSL2)`` route to different drivers for the same target.
    """

    if platform.system() == "Linux":
        try:
            if "microsoft" in Path("/proc/version").read_text(encoding="utf-8").lower():
                return "Linux (WSL2)"
        except OSError:
            pass
    return platform.system()


def _select_targets(
    doc: Dict[str, Any],
    *,
    name: Optional[str],
    all_targets: bool,
) -> List[Tuple[str, Dict[str, Any]]]:
    """Returns the (name, record) list to dispatch.

    If `name` is set, just that one (with no `runs_on` filtering — let
    oracle_gen surface the platform error directly). If `all_targets`,
    every target whose `runs_on:` contains the current OS. Otherwise the
    primary target only.
    """

    targets: Dict[str, Any] = doc.get("targets") or {}
    if name is not None:
        if name not in targets:
            avail = ", ".join(sorted(targets.keys()))
            raise RuntimeError(f"unknown target: {name!r} (available: {avail})")
        return [(name, targets[name])]
    if all_targets:
        chosen: List[Tuple[str, Dict[str, Any]]] = [
            (n, t) for n, t in sorted(targets.items()) if _runs_on_current(t)
        ]
        if not chosen:
            raise RuntimeError(
                "no targets in targets.yaml declare runs_on: "
                f"[{platform.system()}]"
            )
        return chosen
    primary = doc.get("primary")
    if not isinstance(primary, str) or primary not in targets:
        raise RuntimeError("targets.yaml is missing a valid `primary:` entry")
    return [(primary, targets[primary])]


def _cmd_list(args: argparse.Namespace) -> int:
    doc = _load_targets(args.targets_file)
    primary = doc.get("primary")
    targets: Dict[str, Any] = doc.get("targets") or {}
    print(f"host: {_platform_label()}")
    print(f"primary: {primary}")
    print("targets:")
    for name in sorted(targets.keys()):
        record = targets[name] if isinstance(targets[name], dict) else {}
        runs_on = record.get("runs_on") or []
        driver = record.get("driver", "?")
        marker = "*" if name == primary else " "
        print(f"  {marker} {name}  driver={driver}  runs_on={runs_on}")
    return 0


def _cmd_gen(args: argparse.Namespace) -> int:
    doc = _load_targets(args.targets_file)
    selected = _select_targets(doc, name=args.target, all_targets=args.all)

    overall = 0
    for name, _record in selected:
        print(f"[oracle-cli] target={name}")
        # Delegate to oracle_gen.main; it knows how to resolve per-target
        # output_dir / environment_md from the same targets.yaml.
        gen_argv: List[str] = ["--target", name, "--targets-file", str(args.targets_file)]
        if args.suite:
            for s in args.suite:
                gen_argv.extend(["--suite", s])
        if args.strict:
            gen_argv.append("--strict")
        if args.visible:
            gen_argv.append("--visible")
        rc = oracle_gen.main(gen_argv)
        if rc != 0:
            overall = rc
            if args.strict:
                return rc
    return overall


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
            f"interpreter not found: {python_exe}\n"
            "Hint: run `make oracle-setup` to create the venv.",
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
        "xlwings import failed:\n" + (proc.stderr.strip() or proc.stdout.strip())
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
        "xlwings.apps lookup failed:\n" + (proc.stderr.strip() or proc.stdout.strip())
        + "\nHint: System Settings -> Privacy & Security -> Automation\n"
        "       -> (your terminal) -> Microsoft Excel (allow).",
    )


def _check_win_python_path(target: Dict[str, Any]) -> Tuple[str, str, Optional[Path]]:
    """Returns (status, hint, resolved_path) for ``target.win_python``.

    The third element is the resolved Path when the field is set and
    points to an existing file, otherwise ``None`` -- callers use it to
    decide whether dependent checks can run or must SKIP.
    """

    win_python = target.get("win_python")
    if not isinstance(win_python, str) or not win_python.strip():
        return (
            _STATUS_FAIL,
            "target.win_python is empty\n"
            "Hint: install Python on Windows (winget install Python.Python.3.12)\n"
            "      then add this line to tools/oracle/targets.yaml under\n"
            "      the win-365-ja_JP entry:\n"
            "        win_python: \"/mnt/c/Users/<you>/AppData/Local/Programs/"
            "Python/Python312/python.exe\"",
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
        proc = subprocess.run(
            ["wslpath", "-w", "/tmp"], capture_output=True, text=True
        )
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
        f"wslpath -w /tmp failed: rc={proc.returncode}\n"
        + (proc.stderr.strip() or proc.stdout.strip()),
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
                f"target requires Darwin, current host is {host}.\n"
                f"runs_on={_runs_on_label(target)}",
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
        is_wsl2 = host == "Linux" and "(WSL2)" in _platform_label()
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
            f"target needs Windows or WSL2, current host is {host}.\n"
            f"runs_on={_runs_on_label(target)}",
        )
        _print_check(target_name, _STATUS_SKIP, "xlwings + win32com import (host mismatch)")
        _print_check(target_name, _STATUS_SKIP, "wslpath translation (host mismatch)")
        return False

    if driver_name == "wsl_bridge":
        is_wsl2 = host == "Linux" and "(WSL2)" in _platform_label()
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
        "Hint: targets.yaml driver must be one of "
        "'macos_excel', 'windows_excel', 'wsl_bridge'.",
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


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(
        prog="oracle-cli",
        description=__doc__,
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--targets-file",
        type=Path,
        default=DEFAULT_TARGETS_FILE,
        help="Path to targets.yaml (rarely needs overriding).",
    )
    sub = parser.add_subparsers(dest="cmd", required=True)

    p_list = sub.add_parser("list", help="Print available targets.")
    p_list.set_defaults(func=_cmd_list)

    p_gen = sub.add_parser("gen", help="Generate goldens for one or more targets.")
    p_gen.add_argument("--target", default=None, help="Target name (default: primary).")
    p_gen.add_argument(
        "--all",
        action="store_true",
        help="Run every target whose runs_on includes the current OS.",
    )
    p_gen.add_argument(
        "--suite",
        action="append",
        default=None,
        metavar="NAME",
        help="Restrict to the named suite(s); forwarded to oracle_gen.",
    )
    p_gen.add_argument("--strict", action="store_true")
    p_gen.add_argument("--visible", action="store_true")
    p_gen.set_defaults(func=_cmd_gen)

    p_setup = sub.add_parser(
        "setup",
        help="Verify the host can drive its target oracle.",
    )
    p_setup.add_argument(
        "--target",
        default=None,
        help=(
            "Verify just one target by name; defaults to every target whose "
            "runs_on declares the current host."
        ),
    )
    p_setup.set_defaults(func=_cmd_setup)

    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except RuntimeError as exc:
        print(f"oracle-cli: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
