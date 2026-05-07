#!/usr/bin/env python3
"""Cross-platform oracle CLI dispatcher.

A thin orchestrator over the oracle drivers. It (1) loads targets.yaml,
(2) selects targets that match the current platform, (3) either prints
them, runs preflight checks, generates goldens, or guides external
contributors through a one-command donation flow.

Subcommands:
    cli.py list                              # print available targets
    cli.py gen [--target NAME] [--all]       # delegate to oracle_gen
    cli.py setup [--target NAME]             # run preflight checks
    cli.py contribute [--target NAME]        # contributor onramp:
                                             #   banner + preflight + gen
                                             #   + push/PR instructions

Examples:
    python3 tools/oracle/cli.py list
    python3 tools/oracle/cli.py gen
    python3 tools/oracle/cli.py gen --target mac-365-ja_JP
    python3 tools/oracle/cli.py gen --all
    python3 tools/oracle/cli.py setup
    python3 tools/oracle/cli.py setup --target win-365-ja_JP
    python3 tools/oracle/cli.py contribute --target mac-365-en_US
"""

from __future__ import annotations

import argparse
import datetime as _dt
import json
import platform
import re
import subprocess
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Local imports — accept both `python3 tools/oracle/cli.py` (no package)
# and `python3 -m tools.oracle.cli` (package-style).
try:  # pragma: no cover - trivial fallback
    from tools.oracle import oracle_gen
    from tools.oracle.drivers import resolve_win_python
    from tools.oracle.drivers._locale import COUNTRY_CODE_TO_BCP47
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import oracle_gen  # type: ignore
    from drivers import resolve_win_python  # type: ignore
    from drivers._locale import COUNTRY_CODE_TO_BCP47  # type: ignore


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

    win_python = resolve_win_python(target)
    if not win_python:
        return (
            _STATUS_FAIL,
            "win_python not configured\n"
            "Hint: install Python on Windows (winget install Python.Python.3.12),\n"
            "      then either export FORMULON_WIN_PYTHON pointing at the\n"
            "      Windows-side python.exe (preferred for OSS contributors so\n"
            "      no per-machine path lands in targets.yaml) or, for a\n"
            "      private fork, add a win_python: line under the target.\n"
            "      Example:\n"
            "        export FORMULON_WIN_PYTHON=\"/mnt/c/Users/<you>/AppData/"
            "Local/Programs/Python/Python312/python.exe\"",
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
            f"target needs Windows or WSL2, current host is {host}.\n"
            f"runs_on={_runs_on_label(target)}",
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
        "Hint: targets.yaml driver must be one of "
        "'macos_excel', 'windows_excel', 'wsl_bridge'.",
    )
    return False


_CONTRIBUTE_BANNER = """\
========================================================================
  Formulon oracle contribution flow

  Thank you for taking the time to donate Excel oracle data.

  Why we need this:
    Formulon's 1-bit compatibility claim is anchored to ONE primary
    oracle (Mac Excel 365, ja-JP). Several function families behave
    differently across locales (BAHTTEXT, NUMBERSTRING, decimal-comma
    locales, localized TEXT format codes) and across the Mac/Windows
    ports. The maintainer team can't reproduce these from a single
    install -- Excel's locale is system-wide, not a runtime switch,
    and licenses are per-account/per-platform. Goldens captured by
    contributors who actually run Excel in that locale are how
    Formulon stays honest.

  What this command does:
    1. Confirms your host can drive the chosen target (preflight).
    2. Drives Excel to evaluate every YAML case under
       tests/oracle/cases/ and writes golden JSON.
    3. Records the Excel build + locale to ENVIRONMENT.md.
    4. Prints the exact git commands to push your branch and the
       PR URL with the oracle template prefilled.

  Tip: you can re-run safely. Goldens are rewritten from scratch.
========================================================================
"""


# Excel `Application.International(xlCountryCode)` (xlCountryCode = 1) maps
# the host's region setting to a phone-style country code. The shared
# `COUNTRY_CODE_TO_BCP47` map under tools/oracle/drivers/_locale.py is the
# single source of truth -- alias it locally so existing call sites and
# any external scripts that import `_COUNTRY_TO_LOCALE` keep working.
_COUNTRY_TO_LOCALE: Dict[int, str] = COUNTRY_CODE_TO_BCP47


def _short_host(host: str, label: Optional[str] = None) -> Optional[str]:
    """Maps platform.system() -> the prefix we use in target names.

    `Darwin` -> `mac`, Windows or WSL2 -> `win`, anything else -> None
    (which causes auto-detection to bail out and ask for --target).
    """

    if host == "Darwin":
        return "mac"
    if host == "Windows":
        return "win"
    if host == "Linux" and (label or _platform_label()).endswith("(WSL2)"):
        return "win"
    return None


def _version_tuple(s: str) -> Tuple[int, ...]:
    """Returns a comparable integer tuple from an Excel version string.

    Accepts ``"16.108.1"``, ``"16.84 (Build 24021522)"``, ``"16.84"``;
    extracts every leading-digit run separated by ``.`` until the first
    non-numeric chunk. Returns ``(0,)`` if no digits were found, so the
    caller can compare without crashing on garbage input.
    """

    parts: List[int] = []
    for chunk in re.split(r"[.\s(]+", s.strip()):
        m = re.match(r"\d+", chunk)
        if not m:
            break
        parts.append(int(m.group()))
    return tuple(parts) or (0,)


_PROBE_SCRIPT = """
import json
import sys

try:
    import xlwings  # noqa: F401
except Exception as exc:  # pragma: no cover
    print(json.dumps({"error": "xlwings missing: " + str(exc)}))
    sys.exit(2)

app = None
try:
    app = xlwings.App(visible=False, add_book=False)
    version = ""
    try:
        version = str(app.version) if app.version else ""
    except Exception:
        version = ""
    cc = None
    try:
        api = app.api
        # xlCountryCode = 1 in Application.International(...)
        v = api.international(1)
        cc = v() if callable(v) else v
        cc = int(cc)
    except Exception:
        cc = None
    if cc is None:
        # Mac AppleScript dictionary fallback: `country code` is a
        # property on the application object on some Excel builds.
        try:
            api = app.api
            v = getattr(api, "country_code", None)
            if v is not None:
                cc = int(v() if callable(v) else v)
        except Exception:
            cc = None
    print(json.dumps({"version": version, "country_code": cc}))
finally:
    if app is not None:
        try:
            app.quit()
        except Exception:
            pass
"""


def _probe_excel_brief(host: str) -> Optional[Dict[str, Any]]:
    """Briefly opens Excel via xlwings to read version + country code.

    Returns ``{"version": "16.x", "country_code": <int>, "locale": "xx-YY"}``
    or ``None`` if anything in the chain fails (xlwings not installed,
    Excel not reachable, country code couldn't be mapped). The caller
    treats ``None`` as "fall back to --target".
    """

    venv = _venv_python()
    if not venv.exists():
        return None
    try:
        proc = subprocess.run(
            [str(venv), "-c", _PROBE_SCRIPT],
            capture_output=True,
            text=True,
            timeout=120,
        )
    except (subprocess.TimeoutExpired, FileNotFoundError):
        return None
    if proc.returncode != 0 or not proc.stdout.strip():
        return None
    try:
        data = json.loads(proc.stdout.strip().splitlines()[-1])
    except json.JSONDecodeError:
        return None
    if "error" in data:
        return None
    version = data.get("version") or ""
    cc = data.get("country_code")
    if not version or cc is None:
        return None
    locale = _COUNTRY_TO_LOCALE.get(int(cc))
    if not locale:
        return {"version": version, "country_code": int(cc), "locale": None}
    return {"version": version, "country_code": int(cc), "locale": locale}


def _committed_excel_version_for(record: Dict[str, Any]) -> Optional[str]:
    """Reads the committed ENVIRONMENT.md for `record` and returns the
    Excel version string, or None if the file is missing / unparsable.
    """

    env_md_rel = record.get("environment_md")
    if not isinstance(env_md_rel, str) or not env_md_rel:
        return None
    env_md = REPO_ROOT / env_md_rel
    if not env_md.exists():
        return None
    pat = re.compile(r"-\s*\*\*Excel version\*\*:\s*`([^`]+)`")
    for line in env_md.read_text(encoding="utf-8").splitlines():
        m = pat.match(line.strip())
        if m:
            return m.group(1).strip()
    return None


def _print_unknown_target_help(
    probe: Dict[str, Any], derived_name: str, targets: Dict[str, Any]
) -> None:
    """Tells the contributor we don't track their locale yet.

    Prints to stderr -- the banner + this message together give a
    contributor enough to either choose an existing target or open an
    issue to add theirs.
    """

    locale = probe.get("locale") or "?"
    cc = probe.get("country_code")
    cc_str = f"country_code={cc}" if cc is not None else "country_code=?"
    msg_lines: List[str] = [
        "[contribute] detected Excel:",
        f"    version : {probe.get('version', '?')}",
        f"    locale  : {locale}  ({cc_str})",
        "",
        f"[contribute] derived target name: {derived_name}",
        "[contribute] this target isn't tracked in tools/oracle/targets.yaml.",
        "",
        "Two ways forward:",
        "  1) Open an issue so we can reserve the slot, then re-run:",
        f"     https://github.com/libraz/formulon/issues/new"
        f"?title=Oracle+target%3A+{derived_name}",
        "  2) Or pass --target=<existing-name> if you intend to map your",
        "     environment onto an already-tracked target.",
        "",
        "Existing targets compatible with this host:",
    ]
    host = platform.system()
    compatible = sorted(
        [
            (n, t)
            for n, t in targets.items()
            if isinstance(t, dict) and host in (t.get("runs_on") or [])
        ]
    )
    if not compatible:
        msg_lines.append("  (none)")
    else:
        for n, t in compatible:
            msg_lines.append(
                f"  {n}  locale={t.get('locale', '?')}  "
                f"status={t.get('status', '?')}"
            )
    print("\n".join(msg_lines), file=sys.stderr)


def _print_already_current_thanks(
    target_name: str,
    record: Dict[str, Any],
    committed_version: str,
    probed_version: str,
) -> None:
    """Closing message when no regeneration is needed.

    The repository's golden was generated against an Excel build at or
    above the version installed on this contributor's machine, so we
    skip oracle-gen entirely and just thank them for showing up.
    """

    print()
    print("========================================================================")
    print("  Thank you for trying to contribute oracle data!")
    print()
    print(
        f"  The repository already has goldens for `{target_name}` generated"
    )
    print(
        f"  against Excel {committed_version} -- that's at or above your"
    )
    print(
        f"  installed Excel {probed_version}, so there's nothing to refresh"
    )
    print("  right now. No PR is needed.")
    print()
    print("  When would a refresh help?")
    print("    * Your Excel updates and overtakes the committed version.")
    print(
        "    * A new oracle case is added to tests/oracle/cases/ that this"
    )
    print("      target hasn't covered yet.")
    print(
        "    * A divergence report against this target lands and we ask for"
    )
    print("      a re-run.")
    print()
    print("  Either way -- thanks for showing up. Run again any time:")
    print("    make oracle-contribute")
    print("========================================================================")


def _git_changed_paths(prefixes: List[str]) -> List[str]:
    """Returns repo-relative paths under any prefix that git considers
    modified, added, or untracked. Best-effort -- silently returns an
    empty list if git is unavailable.
    """

    try:
        proc = subprocess.run(
            ["git", "-C", str(REPO_ROOT), "status", "--porcelain"],
            capture_output=True,
            text=True,
        )
    except FileNotFoundError:
        return []
    if proc.returncode != 0:
        return []
    paths: List[str] = []
    for line in proc.stdout.splitlines():
        # porcelain format: "XY path" (status code is two chars + space).
        if len(line) < 4:
            continue
        path = line[3:].strip()
        # Handle rename "old -> new" -- keep the new path.
        if " -> " in path:
            path = path.split(" -> ", 1)[1]
        if any(path == p or path.startswith(p.rstrip("/") + "/") for p in prefixes):
            paths.append(path)
    return sorted(set(paths))


def _print_contribute_thanks(
    target_name: str, record: Dict[str, Any], targets_file: Path
) -> None:
    """Prints the closing thank-you message and git/PR instructions."""

    today = _dt.date.today().isoformat()
    branch = f"oracle/{target_name}-{today}"
    output_dir = str(record.get("output_dir") or "")
    env_md = str(record.get("environment_md") or "")
    targets_rel = str(targets_file.relative_to(REPO_ROOT)) if targets_file.is_relative_to(REPO_ROOT) else str(targets_file)

    prefixes = [p for p in (output_dir, env_md, targets_rel) if p]
    changed = _git_changed_paths(prefixes)

    print()
    print("========================================================================")
    print("  Done. Your oracle data is on disk -- thank you for the donation!")
    print()
    print(f"  Target:   {target_name}")
    print(f"  Locale:   {record.get('locale', '?')}")
    print(f"  Host:     {_platform_label()}")
    if env_md:
        print(f"  Env file: {env_md}")
    print()
    if changed:
        print("  Files git considers new or changed:")
        for path in changed[:25]:
            print(f"    {path}")
        if len(changed) > 25:
            print(f"    ... ({len(changed) - 25} more)")
    else:
        print("  No git changes detected. (Goldens may already match HEAD --")
        print("  re-running on an unchanged tree is a no-op.)")
    print()
    print("  Next: push your branch and open a PR.")
    print()
    print(f"    git checkout -b {branch}")
    if output_dir:
        print(f"    git add {output_dir}")
    if env_md:
        print(f"    git add {env_md}")
    if record.get("status") == "wanted":
        print(f"    git add {targets_rel}    # bump status: wanted -> scaffolded")
    print(f"    git commit -m 'test(oracle): contribute {target_name} goldens'")
    print(f"    git push -u origin {branch}")
    print()
    print("  Then open a PR (oracle template auto-loads):")
    print(
        f"    https://github.com/libraz/formulon/compare/main...{branch}"
        "?expand=1&template=oracle.md"
    )
    print()
    print("  Notes for reviewers and you:")
    print("    * Confirm ENVIRONMENT.md records the locale you intended.")
    print("    * For first-time targets, a second contributor will be asked")
    print("      to regenerate independently and diff -- see CONTRIBUTING.md")
    print("      for the two-person verify rule.")
    print("========================================================================")


def _cmd_contribute(args: argparse.Namespace) -> int:
    """Contributor onramp: probe Excel, decide if a refresh is needed,
    and (if so) run preflight + gen + push instructions.

    Each contributor has exactly one Excel install, so the flow is:
      1. Open Excel briefly and read version + country code (locale).
      2. Map (host_short, '365', locale) -> a target name in targets.yaml.
      3. If the committed ENVIRONMENT.md records a version >= the probed
         one, skip generation and just thank them.
      4. Otherwise run preflight + oracle_gen + closing instructions.

    The operator can override step 1 with --target if the auto-derived
    name doesn't fit (or the probe fails).
    """

    doc = _load_targets(args.targets_file)
    targets: Dict[str, Any] = doc.get("targets") or {}
    print(_CONTRIBUTE_BANNER)

    host_sys = platform.system()
    host_label = _platform_label()
    host_short = _short_host(host_sys, host_label)

    probe = _probe_excel_brief(host_sys)

    # Resolve the target. Explicit --target wins; otherwise we derive
    # from the probe.
    name: Optional[str] = None
    record: Optional[Dict[str, Any]] = None

    if args.target is not None:
        if args.target not in targets:
            avail = ", ".join(sorted(targets.keys()))
            print(
                f"unknown target: {args.target!r} (available: {avail})",
                file=sys.stderr,
            )
            return 2
        name, record = args.target, targets[args.target]
        if probe is not None and probe.get("version"):
            print(
                f"[contribute] detected Excel {probe['version']}"
                f" (locale={probe.get('locale') or '?'})"
            )
        print(f"[contribute] target: {name} (explicit)")
    else:
        if probe is None:
            print(
                "[contribute] could not auto-detect your Excel environment.",
                file=sys.stderr,
            )
            print(
                "    The probe needs xlwings + a working Excel install."
                " First-time setup:",
                file=sys.stderr,
            )
            print("      make oracle-setup", file=sys.stderr)
            print(
                "    Already set up? Re-run with TARGET=<name>;"
                " run `make oracle-contribute-list` to see the options.",
                file=sys.stderr,
            )
            return 2
        if not host_short or not probe.get("locale"):
            cc = probe.get("country_code")
            print(
                f"[contribute] detected Excel {probe['version']}"
                f" (country_code={cc}); locale could not be mapped.",
                file=sys.stderr,
            )
            print(
                "    Re-run with TARGET=<name> to pick a target manually,"
                " or open an issue with the country_code above so we can"
                " add the mapping.",
                file=sys.stderr,
            )
            return 2
        derived = f"{host_short}-365-{probe['locale'].replace('-', '_')}"
        if derived not in targets:
            _print_unknown_target_help(probe, derived, targets)
            return 2
        name, record = derived, targets[derived]
        print(
            f"[contribute] detected Excel {probe['version']}"
            f" (locale={probe['locale']})"
        )
        print(f"[contribute] target: {name} (auto-detected)")

    assert name is not None and record is not None  # for type-checkers

    # If the committed golden was generated against a version at or above
    # the contributor's Excel, there is nothing for them to refresh. Say
    # thanks and exit -- no PR needed.
    if probe is not None and probe.get("version"):
        committed = _committed_excel_version_for(record)
        if committed and _version_tuple(committed) >= _version_tuple(probe["version"]):
            _print_already_current_thanks(
                name, record, committed, probe["version"]
            )
            return 0

    print()
    print(f"[contribute] preflight: target={name}")
    if not _check_target(name, record, host_label):
        print()
        print(
            "[contribute] preflight failed. Address the FAIL lines above"
            " (the hints are copy-paste ready), then rerun:",
            file=sys.stderr,
        )
        print("    make oracle-contribute", file=sys.stderr)
        return 2

    print()
    print(f"[contribute] generating goldens for target={name}")
    rc = oracle_gen.main(
        ["--target", name, "--targets-file", str(args.targets_file)]
    )
    if rc != 0:
        print()
        print(
            f"[contribute] oracle_gen exited with code {rc}."
            " See the log above; common fixes:",
            file=sys.stderr,
        )
        print("  * Quit Excel and rerun (the driver opens a fresh app).", file=sys.stderr)
        print("  * Confirm Excel is signed in and Office is activated.", file=sys.stderr)
        print(
            "  * On macOS, re-grant Automation permission if it was reset"
            " by an OS update.",
            file=sys.stderr,
        )
        return rc

    _print_contribute_thanks(name, record, args.targets_file)
    return 0


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

    p_contrib = sub.add_parser(
        "contribute",
        help=(
            "Contributor onramp: thank-you banner + preflight + gen + "
            "push/PR instructions for one variant target."
        ),
    )
    p_contrib.add_argument(
        "--target",
        default=None,
        help=(
            "Target name to contribute for. If omitted and exactly one "
            "wanted target is compatible with this host, that one is "
            "auto-selected; otherwise the available targets are listed."
        ),
    )
    p_contrib.set_defaults(func=_cmd_contribute)

    args = parser.parse_args(argv)
    try:
        return args.func(args)
    except RuntimeError as exc:
        print(f"oracle-cli: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
