"""Contributor onramp: probe the local Excel, compare it with the committed goldens and print the donation steps."""

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

# Local imports -- accept both `python3 tools/oracle/cli.py` (no package)
# and `python3 -m tools.oracle.cli` (package-style).
try:  # pragma: no cover - trivial fallback
    from tools.oracle import oracle_gen
    from tools.oracle.drivers._locale import COUNTRY_CODE_TO_BCP47
    from tools.oracle.preflight import _check_target, _venv_python
    from tools.oracle.targets_manifest import _load_targets, _platform_label
except ImportError:  # pragma: no cover
    import oracle_gen  # type: ignore
    from drivers._locale import COUNTRY_CODE_TO_BCP47  # type: ignore
    from preflight import _check_target, _venv_python  # type: ignore
    from targets_manifest import _load_targets, _platform_label  # type: ignore


REPO_ROOT = Path(__file__).resolve().parents[2]

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


def _print_unknown_target_help(probe: Dict[str, Any], derived_name: str, targets: Dict[str, Any]) -> None:
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
        f"     https://github.com/libraz/formulon/issues/new?title=Oracle+target%3A+{derived_name}",
        "  2) Or pass --target=<existing-name> if you intend to map your",
        "     environment onto an already-tracked target.",
        "",
        "Existing targets compatible with this host:",
    ]
    host = platform.system()
    compatible = sorted(
        [(n, t) for n, t in targets.items() if isinstance(t, dict) and host in (t.get("runs_on") or [])]
    )
    if not compatible:
        msg_lines.append("  (none)")
    else:
        for n, t in compatible:
            msg_lines.append(f"  {n}  locale={t.get('locale', '?')}  status={t.get('status', '?')}")
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
    print(f"  The repository already has goldens for `{target_name}` generated")
    print(f"  against Excel {committed_version} -- that's at or above your")
    print(f"  installed Excel {probed_version}, so there's nothing to refresh")
    print("  right now. No PR is needed.")
    print()
    print("  When would a refresh help?")
    print("    * Your Excel updates and overtakes the committed version.")
    print("    * A new oracle case is added to tests/oracle/cases/ that this")
    print("      target hasn't covered yet.")
    print("    * A divergence report against this target lands and we ask for")
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


def _print_contribute_thanks(target_name: str, record: Dict[str, Any], targets_file: Path) -> None:
    """Prints the closing thank-you message and git/PR instructions."""

    today = _dt.date.today().isoformat()
    branch = f"oracle/{target_name}-{today}"
    output_dir = str(record.get("output_dir") or "")
    env_md = str(record.get("environment_md") or "")
    targets_rel = (
        str(targets_file.relative_to(REPO_ROOT)) if targets_file.is_relative_to(REPO_ROOT) else str(targets_file)
    )

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
    print(f"    https://github.com/libraz/formulon/compare/main...{branch}?expand=1&template=oracle.md")
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
            print(f"[contribute] detected Excel {probe['version']} (locale={probe.get('locale') or '?'})")
        print(f"[contribute] target: {name} (explicit)")
    else:
        if probe is None:
            print(
                "[contribute] could not auto-detect your Excel environment.",
                file=sys.stderr,
            )
            print(
                "    The probe needs xlwings + a working Excel install. First-time setup:",
                file=sys.stderr,
            )
            print("      make oracle-setup", file=sys.stderr)
            print(
                "    Already set up? Re-run with TARGET=<name>; run `make oracle-contribute-list` to see the options.",
                file=sys.stderr,
            )
            return 2
        if not host_short or not probe.get("locale"):
            cc = probe.get("country_code")
            print(
                f"[contribute] detected Excel {probe['version']} (country_code={cc}); locale could not be mapped.",
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
        print(f"[contribute] detected Excel {probe['version']} (locale={probe['locale']})")
        print(f"[contribute] target: {name} (auto-detected)")

    assert name is not None and record is not None  # for type-checkers

    # If the committed golden was generated against a version at or above
    # the contributor's Excel, there is nothing for them to refresh. Say
    # thanks and exit -- no PR needed.
    if probe is not None and probe.get("version"):
        committed = _committed_excel_version_for(record)
        if committed and _version_tuple(committed) >= _version_tuple(probe["version"]):
            _print_already_current_thanks(name, record, committed, probe["version"])
            return 0

    print()
    print(f"[contribute] preflight: target={name}")
    if not _check_target(name, record, host_label):
        print()
        print(
            "[contribute] preflight failed. Address the FAIL lines above (the hints are copy-paste ready), then rerun:",
            file=sys.stderr,
        )
        print("    make oracle-contribute", file=sys.stderr)
        return 2

    print()
    print(f"[contribute] generating goldens for target={name}")
    rc = oracle_gen.main(["--target", name, "--targets-file", str(args.targets_file)])
    if rc != 0:
        print()
        print(
            f"[contribute] oracle_gen exited with code {rc}. See the log above; common fixes:",
            file=sys.stderr,
        )
        print("  * Quit Excel and rerun (the driver opens a fresh app).", file=sys.stderr)
        print("  * Confirm Excel is signed in and Office is activated.", file=sys.stderr)
        print(
            "  * On macOS, re-grant Automation permission if it was reset by an OS update.",
            file=sys.stderr,
        )
        return rc

    _print_contribute_thanks(name, record, args.targets_file)
    return 0
