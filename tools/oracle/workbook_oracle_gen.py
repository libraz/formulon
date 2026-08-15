#!/usr/bin/env python3
"""Generates golden JSON for the workbook oracle track.

The workbook track covers workbook-level features that are NOT formula
results -- pivot tables and print areas. Cases are declarative mini-
workbook specs discovered from `tests/oracle/cases_wb/*.case.json`; the
golden JSON committed under `tests/oracle/golden_wb/` (or, for variants,
`tests/oracle/variants/<target>/golden_wb/`) is what the C++ verifier
diffs against.

Target resolution reads the `tracks.workbook` section of
`tools/oracle/targets.yaml`. The workbook track uses `win-365-ja_JP` as
its candidate because reliable PivotTable automation is only available
through Windows COM; eligibility is always read from the target manifest.

Golden JSON shape (one file per suite):

    {
      "suite": "<suite-name>",
      "kind": "workbook",
      "environment": { "excel_version": ..., "excel_locale": ..., ... },
      "cases": [
        {
          "id": "<case-id>",
          "spec": { ...the declarative mini-workbook spec... },
          "expect": { ...observed pivot / print result... }
        }
      ]
    }

This entry point wires target resolution, suite discovery, provenance-safe
staging, and golden write-out around the concrete driver implementations.

Usage:
    python3 tools/oracle/workbook_oracle_gen.py [--target NAME]
                                                [--suite NAME ...]
                                                [--cases-dir P]
                                                [--golden-dir P]
"""

from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import json
import os
import platform
import sys
import tempfile
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

# Local imports -- accept both `python3 tools/oracle/workbook_oracle_gen.py`
# (no package) and `python3 -m tools.oracle.workbook_oracle_gen`.
try:  # pragma: no cover - trivial fallback
    from tools.oracle.drivers import select_driver
    from tools.oracle.oracle_gen import _load_divergence_skips
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    from drivers import select_driver  # type: ignore
    from oracle_gen import _load_divergence_skips  # type: ignore
try:
    from tools.oracle import workbook_case_schema
except ImportError:  # pragma: no cover
    import workbook_case_schema  # type: ignore


REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_CASES_DIR = REPO_ROOT / "tests" / "oracle" / "cases_wb"
DEFAULT_GOLDEN_DIR = REPO_ROOT / "tests" / "oracle" / "golden_wb"
DEFAULT_TARGETS_FILE = Path(__file__).resolve().parent / "targets.yaml"
DEFAULT_DIVERGENCE = REPO_ROOT / "tests" / "divergence.yaml"

# Where a `status: wanted` target stages its capture when the caller does
# not name a directory. A wanted target has no established provenance yet,
# so its goldens must not land in the repository straight off the driver --
# but making the operator invent a path for every run turned the ordinary
# capture into a bespoke command line. A stable per-user cache keeps the
# staging out of the tree, survives a reboot (unlike /tmp) so a capture can
# be reviewed the next day, and gives `promote_capture.py` a default source.
STAGING_ROOT = Path.home() / ".cache" / "formulon" / "oracle-capture"


def staging_dir_for(track: str, target_name: str) -> Path:
    """Returns the default staging directory for one track / target pair."""

    return STAGING_ROOT / track / target_name


def _load_targets(path: Path) -> Dict[str, Any]:
    """Loads `targets.yaml`. Raises RuntimeError on any read / parse error."""

    if not path.exists():
        raise RuntimeError(f"oracle targets file not found: {path}")
    try:
        import yaml  # type: ignore

        doc = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except Exception as exc:  # pragma: no cover - yaml import / parse guard
        raise RuntimeError(f"failed to parse {path}: {exc}") from exc
    if not isinstance(doc, dict):
        raise RuntimeError(f"{path} root must be a mapping")
    targets = doc.get("targets")
    if not isinstance(targets, dict) or not targets:
        raise RuntimeError(f"{path} has no `targets:` mapping")
    return doc


def _workbook_primary(targets_doc: Dict[str, Any]) -> str:
    """Return the manifest-declared workbook primary, failing closed."""

    tracks = targets_doc.get("tracks")
    if isinstance(tracks, dict):
        workbook = tracks.get("workbook")
        if isinstance(workbook, dict):
            primary = workbook.get("primary")
            if isinstance(primary, str) and primary:
                return primary
    raise RuntimeError("targets.yaml: tracks.workbook.primary is required")


def _workbook_track_targets(targets_doc: Dict[str, Any]) -> List[str]:
    """Returns the workbook track's `[primary, *variants]` target names."""

    names: List[str] = [_workbook_primary(targets_doc)]
    tracks = targets_doc.get("tracks")
    if isinstance(tracks, dict):
        workbook = tracks.get("workbook")
        if isinstance(workbook, dict):
            variants = workbook.get("variants")
            if isinstance(variants, list):
                names.extend(str(v) for v in variants)
    return names


def _autodetect_target(targets_doc: Dict[str, Any]) -> str:
    """Picks the workbook-track target whose `runs_on` matches this host.

    The host OS decides which Excel install can be driven, so the common
    case needs no `--target`: a Windows / WSL2 host (`platform.system()`
    is `Windows` / `Linux`) resolves to the win primary, a macOS host
    (`Darwin`) to the mac variant. When no track target matches the host,
    falls back to the primary so the driver factory surfaces a clear
    host-mismatch error rather than this function guessing.
    """

    host = platform.system()
    targets = targets_doc.get("targets") or {}
    for name in _workbook_track_targets(targets_doc):
        record = targets.get(name)
        if isinstance(record, dict) and host in (record.get("runs_on") or []):
            return name
    return _workbook_primary(targets_doc)


def _resolve_target(targets_doc: Dict[str, Any], name: Optional[str]) -> Dict[str, Any]:
    """Returns the target record for `name` (auto-detected when None)."""

    targets = targets_doc.get("targets") or {}
    if name is None:
        name = _autodetect_target(targets_doc)
    if name not in targets:
        avail = ", ".join(sorted(targets.keys()))
        raise RuntimeError(f"unknown oracle target: {name!r} (available: {avail})")
    record = targets[name]
    if not isinstance(record, dict):
        raise RuntimeError(f"target {name!r} must be a mapping")
    record = dict(record)
    record["_name"] = name
    return record


def _golden_dir_for_target(targets_doc: Dict[str, Any], target: Dict[str, Any]) -> Path:
    """Returns the golden_wb directory for `target`.

    The workbook primary writes to `tests/oracle/golden_wb/`; any other
    target is treated as a variant and writes under
    `tests/oracle/variants/<target>/golden_wb/`.
    """

    if target["_name"] == _workbook_primary(targets_doc):
        return DEFAULT_GOLDEN_DIR
    return REPO_ROOT / "tests" / "oracle" / "variants" / target["_name"] / "golden_wb"


def _resolve_skips(
    targets_doc: Dict[str, Any],
    target: Dict[str, Any],
    divergence_path: Path,
) -> Dict[str, str]:
    """Returns the case-id -> reason skip map for the workbook track.

    Reuses `oracle_gen._load_divergence_skips` so the workbook track honors
    `tests/divergence.yaml` exactly the way the formula track does: only
    `mode: skip-oracle` entries are collected, and `applies_to` scoping is
    respected. For a non-primary (variant) target the per-variant override
    file `tests/oracle/variants/<tag>/divergence.yaml` is merged on top --
    entries there win on key collision because they are more specific.
    """

    skips = _load_divergence_skips(divergence_path, target["_name"])
    if target["_name"] != _workbook_primary(targets_doc):
        variant_div = REPO_ROOT / "tests" / "oracle" / "variants" / target["_name"] / "divergence.yaml"
        if variant_div.exists():
            skips.update(_load_divergence_skips(variant_div, target["_name"]))
    return skips


def _discover_workbook_suites(cases_dir: Path) -> List[Tuple[Path, Dict[str, Any]]]:
    """Loads every `*.case.json` under `cases_dir` (non-recursive).

    Returns `(path, case_doc)` pairs sorted by path so generation output
    is deterministic regardless of filesystem order.
    """

    if not cases_dir.exists() or not cases_dir.is_dir():
        return []
    out: List[Tuple[Path, Dict[str, Any]]] = []
    for path in sorted(cases_dir.iterdir()):
        if not path.name.endswith(".case.json"):
            continue
        with path.open("r", encoding="utf-8") as f:
            doc = json.load(f)
        if not isinstance(doc, dict):
            raise RuntimeError(f"{path}: top-level JSON must be an object")
        out.append((path, doc))
    return out


def _suite_name(case_doc: Dict[str, Any], path: Path) -> str:
    """Returns the suite name from the case doc, or the file stem."""

    suite = case_doc.get("suite")
    if isinstance(suite, str) and suite:
        return suite
    # Strip the ".case" left over after ".case.json" -> ".case" stem.
    stem = path.stem
    return stem[:-5] if stem.endswith(".case") else stem


def _env_to_json(env: Any, iso_now: str, capture_id: str) -> Dict[str, Any]:
    """Shapes an EnvironmentInfo snapshot into the golden's environment block."""

    return {
        "excel_version": getattr(env, "excel_version", ""),
        "excel_locale": getattr(env, "excel_locale", ""),
        "date1904": getattr(env, "date1904", False),
        "iterative": getattr(env, "iterative", False),
        "generated_at": iso_now,
        "capture_id": capture_id,
    }


def _display_path(path: Path) -> str:
    """Render repository paths compactly while supporting external staging."""

    try:
        return str(path.resolve().relative_to(REPO_ROOT))
    except ValueError:
        return str(path)


def _write_golden(
    out_path: Path,
    suite_name: str,
    env_json: Dict[str, Any],
    cases: List[Dict[str, Any]],
) -> None:
    """Writes one `<suite>.golden.json` file for the workbook track."""

    doc = {
        "suite": suite_name,
        "kind": "workbook",
        "environment": env_json,
        "cases": cases,
    }
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(
        json.dumps(doc, indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--target",
        default=None,
        metavar="NAME",
        help=(
            "Oracle target from tools/oracle/targets.yaml. When omitted, the "
            "target is auto-detected from the host OS: a Windows / WSL2 host "
            "resolves to the manifest-declared workbook primary, "
            "a macOS host to the mac variant."
        ),
    )
    parser.add_argument(
        "--targets-file",
        type=Path,
        default=DEFAULT_TARGETS_FILE,
        help="Path to targets.yaml (rarely needs overriding).",
    )
    parser.add_argument(
        "--suite",
        action="append",
        default=None,
        metavar="NAME",
        help="Run only the named suite(s); defaults to all *.case.json files.",
    )
    parser.add_argument(
        "--cases-dir",
        type=Path,
        default=DEFAULT_CASES_DIR,
        help="Directory of *.case.json workbook case files.",
    )
    parser.add_argument(
        "--golden-dir",
        type=Path,
        default=None,
        help=("Directory to write *.golden.json files to. Overrides the per-target golden_wb path."),
    )
    parser.add_argument(
        "--divergence",
        type=Path,
        default=DEFAULT_DIVERGENCE,
        help="YAML listing cases to skip; see tests/divergence.yaml.",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="Show the Excel window during generation (debug aid).",
    )
    args = parser.parse_args(argv)

    # Resolve target metadata. Errors here are fatal -- refuse to start
    # rather than write goldens to a stale path on a typo.
    try:
        targets_doc = _load_targets(args.targets_file)
        target = _resolve_target(targets_doc, args.target)
    except RuntimeError as exc:
        print(f"workbook-oracle-gen: {exc}", file=sys.stderr)
        return 2

    staged = False
    if target.get("status") == "wanted" and args.golden_dir is None:
        # No established provenance yet, so the capture stages outside the
        # tree rather than failing: the operator gets a complete run and a
        # `PROVENANCE.candidate.json`, and `promote_capture.py` is what
        # decides whether it earns a place in the repository.
        args.golden_dir = staging_dir_for("workbook", target["_name"])
        staged = True
    if args.golden_dir is not None:
        resolved = args.golden_dir.resolve()
        try:
            resolved.relative_to(REPO_ROOT)
            inside_repo = True
        except ValueError:
            inside_repo = False
        if target.get("status") == "wanted" and inside_repo:
            print(
                "workbook-oracle-gen: wanted target staging directory must be outside the repository", file=sys.stderr
            )
            return 2
        golden_dir = args.golden_dir
    else:
        golden_dir = _golden_dir_for_target(targets_doc, target)
    if staged:
        print(
            f"workbook-oracle-gen: target {target['_name']!r} has status=wanted; "
            f"staging outside the repository at {golden_dir}"
        )
        print(f"  review, then land it with: make oracle-promote TRACK=workbook TARGET={target['_name']}")

    # Cases marked `mode: skip-oracle` in tests/divergence.yaml (plus any
    # per-variant override) are excluded from generation. A typo in an
    # `applies_to` list raises here -- fail fast rather than silently
    # masking entries.
    try:
        skips = _resolve_skips(targets_doc, target, args.divergence)
    except RuntimeError as exc:
        print(f"workbook-oracle-gen: {exc}", file=sys.stderr)
        return 2

    try:
        suites = _discover_workbook_suites(args.cases_dir)
    except (RuntimeError, json.JSONDecodeError) as exc:
        print(f"workbook-oracle-gen: {exc}", file=sys.stderr)
        return 2
    required_suites = sorted(_suite_name(doc, path) for path, doc in suites)
    if args.suite:
        wanted = set(args.suite)
        suites = [(p, d) for (p, d) in suites if _suite_name(d, p) in wanted]
    if not suites:
        print(f"workbook-oracle-gen: no *.case.json suites found in {args.cases_dir}")
        return 0

    # Driver factory errors (wrong host OS, missing config) are fatal.
    try:
        oracle_cm = select_driver(target, visible=args.visible)
    except RuntimeError as exc:
        print(f"workbook-oracle-gen: {exc}", file=sys.stderr)
        return 2

    iso_now = _dt.datetime.utcnow().replace(microsecond=0).isoformat() + "Z"

    with oracle_cm as oracle:
        # Sentinel: refuse to start when Excel is pre-M365. Mirrors the
        # check in oracle_gen.py; see OracleDriver.assert_m365_or_abort
        # for the why (win-2019 archive incident).
        try:
            oracle.assert_m365_or_abort()
        except RuntimeError as exc:
            print(f"workbook-oracle-gen: {exc}", file=sys.stderr)
            return 2

        env = oracle.probe_environment()
        declared_locale = target.get("locale")
        if not isinstance(declared_locale, str) or not declared_locale or env.excel_locale != declared_locale:
            print(
                f"workbook-oracle-gen: target locale {declared_locale!r} does not match Excel locale {env.excel_locale!r}",
                file=sys.stderr,
            )
            return 2
        capture_id = f"{target['_name']}:{env.excel_version}:{env.excel_locale}:{iso_now}"
        env_json = _env_to_json(env, iso_now, capture_id)
        golden_dir.mkdir(parents=True, exist_ok=True)
        exit_code = 0
        inventory: List[Dict[str, Any]] = []
        for path, case_doc in suites:
            suite_name = _suite_name(case_doc, path)
            raw_cases = case_doc.get("cases") or []
            print(f"[workbook-oracle-gen] {suite_name}  ({len(raw_cases)} cases)")
            try:
                out_cases: List[Dict[str, Any]] = []
                skipped_here = 0
                for case in raw_cases:
                    if not isinstance(case, dict):
                        raise RuntimeError(f"{path}: case entry is not an object")
                    cid = case.get("id")
                    # A divergence.yaml skip-oracle entry excludes this
                    # case from Excel automation; the golden records the
                    # documented reason in place of `expect` so the C++
                    # verifier can recognise an intentional gap.
                    if isinstance(cid, str) and cid in skips:
                        out_cases.append(
                            {
                                "id": cid,
                                "spec": case,
                                "skipped": skips[cid],
                            }
                        )
                        skipped_here += 1
                        continue
                    # The driver evaluates the declarative workbook spec
                    # and returns the observed pivot / print result.
                    expect = oracle.run_workbook_case(case)
                    out_cases.append(
                        {
                            "id": cid,
                            "spec": case,
                            "expect": expect,
                        }
                    )
                if skipped_here:
                    print(f"  ! {skipped_here} case(s) skipped by divergence.yaml (see golden 'skipped' fields)")
                out_path = golden_dir / f"{suite_name}.golden.json"
                _write_golden(out_path, suite_name, env_json, out_cases)
                case_ids, _ = workbook_case_schema.validate_pair(path, out_path)
                if len(case_ids) != len(out_cases) or not out_cases:
                    raise RuntimeError(f"{suite_name}: generated golden is empty or incomplete")
                inventory.append(
                    {
                        "suite": suite_name,
                        "case_count": len(out_cases),
                        "sha256": hashlib.sha256(out_path.read_bytes()).hexdigest(),
                        "capture_id": capture_id,
                    }
                )
                print(f"  -> {_display_path(out_path)}")
            except NotImplementedError:
                exit_code = 1
                print(
                    f"  ! workbook driver unavailable for target {target['_name']!r}",
                    file=sys.stderr,
                )
            except Exception as exc:  # pragma: no cover - per-suite guard
                exit_code = 1
                print(f"  ! failed: {exc}", file=sys.stderr)
        captured_suites = sorted(item["suite"] for item in inventory)
        if exit_code == 0 and captured_suites == required_suites:
            candidate = {
                "target": target["_name"],
                "status": target.get("status"),
                "classification": "candidate",
                "verified": True,
                "active_ctest": False,
                "capture_id": capture_id,
                "all_suites_same_capture": True,
                "product": env.excel_version,
                "locale": env.excel_locale,
                "m365_sentinel": "ARRAYTOTEXT(1) == text '1'",
                "required_suites": required_suites,
                "captured_suites": captured_suites,
                "suite_inventory": inventory,
            }
            candidate_path = golden_dir / "PROVENANCE.candidate.json"
            with tempfile.NamedTemporaryFile(
                "w", encoding="utf-8", dir=golden_dir, prefix=".PROVENANCE.candidate.", delete=False
            ) as temp:
                temp.write(json.dumps(candidate, indent=2, ensure_ascii=False) + "\n")
                temp_path = Path(temp.name)
            os.replace(temp_path, candidate_path)
            print(f"  -> {_display_path(candidate_path)} (verified candidate)")
        elif exit_code == 0:
            print("  ! partial suite capture; no provenance candidate emitted", file=sys.stderr)
    return exit_code


if __name__ == "__main__":
    sys.exit(main())
