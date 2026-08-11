#!/usr/bin/env python3
"""Oracle target status and golden provenance checks.

The target manifest is intentionally small and human-readable, but a target
name alone is not evidence that the files below it came from the claimed
Excel product.  This module keeps that distinction explicit:

* ``status: wanted`` is a reserved target and is never an active CTest
  variant;
* a ``reference-only`` capture may remain on disk for historical comparison,
  but it cannot be reported as a verified Microsoft 365 capture; and
* active variant directories opt in through a small ``PROVENANCE.json``
  record.  CMake uses the same record when it discovers variant goldens.

The command line is deliberately dependency-light and is useful in CI and
in review scripts::

    python3 tools/oracle/provenance.py check
    python3 tools/oracle/provenance.py active-variants
    python3 tools/oracle/provenance.py workbook-active
    python3 tools/oracle/provenance.py cf-active

``check`` exits non-zero on a stale or contradictory manifest.  It does not
delete or rewrite any golden data.
"""

from __future__ import annotations

import argparse
import hashlib
import json
import sys
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional, Sequence

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TARGETS = Path(__file__).resolve().parent / "targets.yaml"
_VALID_STATUSES = {"primary", "scaffolded", "wanted"}
_ACTIVE_STATUSES = {"primary", "scaffolded"}
_M365_SENTINEL = "ARRAYTOTEXT(1) == text '1'"


def _load_yaml(path: Path) -> Dict[str, Any]:
    try:
        import yaml  # type: ignore
    except ImportError as exc:  # pragma: no cover - setup failure
        raise RuntimeError("PyYAML is required to inspect oracle targets") from exc
    if not path.exists():
        raise RuntimeError(f"targets file does not exist: {path}")
    try:
        raw = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    except Exception as exc:  # pragma: no cover - parser-specific
        raise RuntimeError(f"failed to parse {path}: {exc}") from exc
    if not isinstance(raw, dict):
        raise RuntimeError(f"{path}: root must be a mapping")
    targets = raw.get("targets")
    if not isinstance(targets, dict) or not targets:
        raise RuntimeError(f"{path}: missing non-empty targets mapping")
    return raw


def _record_status(record: Mapping[str, Any]) -> str:
    status = record.get("status")
    return status if isinstance(status, str) else ""


def _provenance_path(repo_root: Path, target: str, record: Mapping[str, Any]) -> Path:
    env = record.get("environment_md")
    if isinstance(env, str) and env:
        return repo_root / Path(env).parent / "PROVENANCE.json"
    output = record.get("output_dir")
    if isinstance(output, str) and output:
        return repo_root / output / ".." / "PROVENANCE.json"
    return repo_root / "tests" / "oracle" / "variants" / target / "PROVENANCE.json"


def load_provenance(repo_root: Path, target: str, record: Mapping[str, Any]) -> Optional[Dict[str, Any]]:
    """Load a target's optional directory provenance record."""

    path = _provenance_path(repo_root, target, record).resolve()
    if not path.exists():
        return None
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"{path}: invalid provenance JSON: {exc}") from exc
    if not isinstance(raw, dict):
        raise RuntimeError(f"{path}: provenance root must be an object")
    return raw


def _capture_marker_ready(provenance: Mapping[str, Any], target: str, record: Mapping[str, Any]) -> bool:
    """Check the fields that make a directory marker eligible for coverage."""

    locale = record.get("locale")
    inventory = provenance.get("suite_inventory")
    required = provenance.get("required_suites")
    captured = provenance.get("captured_suites")
    if not isinstance(locale, str) or not locale:
        return False
    if not isinstance(provenance.get("target"), str) or provenance["target"] != target:
        return False
    if not isinstance(provenance.get("product"), str) or not provenance["product"].strip():
        return False
    if provenance.get("product") == "unknown Excel 16.0 capture":
        return False
    if provenance.get("locale") != locale or provenance.get("m365_sentinel") != _M365_SENTINEL:
        return False
    if not isinstance(inventory, list) or not inventory:
        return False
    if not isinstance(provenance.get("capture_id"), str) or not provenance["capture_id"]:
        return False
    if not isinstance(required, list) or not required or not all(isinstance(v, str) for v in required):
        return False
    if not isinstance(captured, list) or sorted(captured) != sorted(required):
        return False
    capture_ids = set()
    for item in inventory:
        if not isinstance(item, dict):
            return False
        if not isinstance(item.get("suite"), str) or item["suite"] not in required:
            return False
        if not isinstance(item.get("case_count"), int) or item["case_count"] <= 0:
            return False
        if not isinstance(item.get("sha256"), str) or len(item["sha256"]) != 64:
            return False
        if not isinstance(item.get("capture_id"), str) or not item["capture_id"]:
            return False
        capture_ids.add(item["capture_id"])
    if len({item.get("suite") for item in inventory}) != len(required):
        return False
    return capture_ids == {provenance["capture_id"]}


def _variant_inventory_ready(repo_root: Path, provenance: Mapping[str, Any], record: Mapping[str, Any]) -> bool:
    """Verify every real variant golden belongs to one recorded capture."""

    output_dir = record.get("output_dir")
    if not isinstance(output_dir, str) or not output_dir:
        return False
    golden_dir = repo_root / output_dir
    actual_paths = sorted(golden_dir.glob("*.golden.json")) if golden_dir.is_dir() else []
    actual_suites = sorted(path.name[: -len(".golden.json")] for path in actual_paths)
    required = provenance.get("required_suites")
    if not isinstance(required, list) or sorted(required) != actual_suites or not actual_suites:
        return False
    inventory = {
        item["suite"]: item
        for item in provenance.get("suite_inventory", [])
        if isinstance(item, dict) and isinstance(item.get("suite"), str)
    }
    if set(inventory) != set(actual_suites):
        return False
    import hashlib

    for path in actual_paths:
        suite = path.name[: -len(".golden.json")]
        item = inventory[suite]
        try:
            document = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return False
        environment = document.get("environment") if isinstance(document, dict) else None
        cases = document.get("cases") if isinstance(document, dict) else None
        if not isinstance(document, dict) or document.get("suite") != suite:
            return False
        if not isinstance(environment, dict) or environment.get("capture_id") != provenance.get("capture_id"):
            return False
        if environment.get("excel_version") != provenance.get("product") or environment.get(
            "excel_locale"
        ) != provenance.get("locale"):
            return False
        if not isinstance(cases, list) or not cases or item.get("case_count") != len(cases):
            return False
        if item.get("capture_id") != provenance.get("capture_id"):
            return False
        if item.get("sha256") != hashlib.sha256(path.read_bytes()).hexdigest():
            return False
    return True


def active_variant_names(doc: Mapping[str, Any], repo_root: Path = REPO_ROOT) -> List[str]:
    """Return variant target names eligible for default CTest coverage.

    ``primary`` is not a variant.  A wanted target is excluded even if its
    directory contains old goldens.  For a scaffolded target, the directory
    record must explicitly say ``active``; this prevents an accidental stale
    directory from becoming a CI surface merely by being copied into the
    tree.
    """

    targets = doc.get("targets")
    if not isinstance(targets, dict):
        return []
    names: List[str] = []
    for name, raw in sorted(targets.items()):
        if not isinstance(raw, dict) or name == doc.get("primary"):
            continue
        if _record_status(raw) not in _ACTIVE_STATUSES:
            continue
        provenance = load_provenance(repo_root, str(name), raw)
        if provenance is None:
            continue
        if not _capture_marker_ready(provenance, str(name), raw):
            continue
        if provenance.get("status") != _record_status(raw):
            continue
        if provenance.get("classification") not in {"active", "verified"}:
            continue
        if provenance.get("verified") is not True:
            continue
        if provenance.get("active_ctest") is not True:
            continue
        if not isinstance(provenance.get("capture_id"), str) or not provenance["capture_id"]:
            continue
        if provenance.get("all_suites_same_capture") is not True:
            continue
        if not _variant_inventory_ready(repo_root, provenance, raw):
            continue
        names.append(str(name))
    return names


def validate_targets(doc: Mapping[str, Any], repo_root: Path = REPO_ROOT) -> List[str]:
    """Validate status/provenance invariants and return human errors."""

    errors: List[str] = []
    targets = doc.get("targets")
    if not isinstance(targets, dict) or not targets:
        return ["targets: missing non-empty mapping"]
    primary = doc.get("primary")
    if not isinstance(primary, str) or primary not in targets:
        errors.append("primary: must name a target")

    tracks = doc.get("tracks")
    if isinstance(tracks, Mapping) and "cf" in tracks:
        cf_track = tracks.get("cf")
        if not isinstance(cf_track, Mapping):
            errors.append("tracks.cf: expected a mapping")
        else:
            cf_primary = cf_track.get("primary")
            if not isinstance(cf_primary, str) or cf_primary not in targets:
                errors.append("tracks.cf.primary: must name a target")
            else:
                # Structural target errors are always reported.  A missing
                # marker is allowed while a capture is pending, but an
                # existing marker must be internally consistent.
                errors.extend(_cf_provenance_errors(doc, repo_root))

    for name, raw in sorted(targets.items()):
        if not isinstance(raw, dict):
            errors.append(f"targets.{name}: expected mapping")
            continue
        status = _record_status(raw)
        if status not in _VALID_STATUSES:
            errors.append(f"targets.{name}.status: expected one of {sorted(_VALID_STATUSES)}, got {status!r}")
            continue
        if name == primary and status != "primary":
            errors.append(f"targets.{name}: manifest primary must have status='primary'")
        provenance = load_provenance(repo_root, str(name), raw)
        if status == "wanted":
            # Wanted targets are reservations, never an assertion that any
            # output directory is a live capture.  If a directory record is
            # present, it must say the same thing explicitly.
            if provenance is not None:
                if provenance.get("classification") != "reference-only":
                    errors.append(f"targets.{name}: wanted target provenance must be reference-only")
                if provenance.get("verified") is True:
                    errors.append(f"targets.{name}: wanted target cannot have verified=true")
            continue
        if status == "scaffolded":
            # Scaffolded targets may have partial output, but they are not
            # active coverage until a verified directory record is supplied.
            continue
        # The manifest's primary target is maintained separately from the
        # variant provenance protocol.  A provenance file is still allowed,
        # but it must not downgrade the primary to reference-only.
        if provenance is not None and provenance.get("classification") == "reference-only":
            errors.append(f"targets.{name}: primary target cannot be reference-only")
    return errors


def workbook_active(doc: Mapping[str, Any], repo_root: Path = REPO_ROOT) -> bool:
    """Whether the workbook track has a manifest-approved active capture."""

    tracks = doc.get("tracks")
    workbook = tracks.get("workbook") if isinstance(tracks, dict) else None
    target = workbook.get("primary") if isinstance(workbook, dict) else None
    targets = doc.get("targets")
    record = targets.get(target) if isinstance(targets, dict) and isinstance(target, str) else None
    if not isinstance(record, dict) or _record_status(record) not in {"primary", "scaffolded"}:
        return False
    marker = repo_root / "tests" / "oracle" / "golden_wb" / "PROVENANCE.json"
    if not marker.exists():
        return False
    try:
        provenance = json.loads(marker.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False
    if not isinstance(provenance, dict) or not _capture_marker_ready(provenance, target, record):
        return False
    if provenance.get("status") != record.get("status"):
        return False
    if provenance.get("classification") not in {"active", "verified"}:
        return False
    if provenance.get("verified") is not True or provenance.get("active_ctest") is not True:
        return False
    capture_id = provenance.get("capture_id")
    if not isinstance(capture_id, str) or not capture_id or provenance.get("all_suites_same_capture") is not True:
        return False
    cases_dir = repo_root / "tests" / "oracle" / "cases_wb"
    required = sorted(path.name[: -len(".case.json")] for path in cases_dir.glob("*.case.json"))
    if sorted(provenance.get("required_suites", [])) != required:
        return False
    inventory = {item["suite"]: item for item in provenance["suite_inventory"]}
    import hashlib

    for suite in required:
        golden = repo_root / "tests" / "oracle" / "golden_wb" / f"{suite}.golden.json"
        if not golden.exists():
            return False
        try:
            doc_json = json.loads(golden.read_text(encoding="utf-8"))
            env = doc_json.get("environment")
            cases = doc_json.get("cases")
            case_doc = json.loads((cases_dir / f"{suite}.case.json").read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError):
            return False
        if not isinstance(env, dict) or env.get("capture_id") != capture_id:
            return False
        if env.get("excel_locale") != record.get("locale") or env.get("excel_version") != provenance.get("product"):
            return False
        if not isinstance(cases, list) or not cases:
            return False
        if not isinstance(case_doc, dict):
            return False
        expected_ids = {case.get("id") for case in case_doc.get("cases", []) if isinstance(case, dict)}
        golden_ids = {case.get("id") for case in cases if isinstance(case, dict)}
        if not expected_ids or expected_ids != golden_ids:
            return False
        item = inventory.get(suite)
        if item is None or item.get("case_count") != len(cases):
            return False
        if hashlib.sha256(golden.read_bytes()).hexdigest() != item.get("sha256"):
            return False
    return True


def _cf_primary(doc: Mapping[str, Any]) -> Optional[str]:
    tracks = doc.get("tracks")
    cf_track = tracks.get("cf") if isinstance(tracks, Mapping) else None
    primary = cf_track.get("primary") if isinstance(cf_track, Mapping) else None
    return primary if isinstance(primary, str) and primary else None


def _cf_provenance_errors(doc: Mapping[str, Any], repo_root: Path) -> List[str]:
    """Validate the CF track's capture marker and every golden hash."""

    primary = _cf_primary(doc)
    if primary is None:
        return ["tracks.cf.primary: must name a target"]
    targets = doc.get("targets")
    record = targets.get(primary) if isinstance(targets, Mapping) else None
    if not isinstance(record, Mapping):
        return [f"tracks.cf.primary: unknown target {primary!r}"]
    errors: List[str] = []
    if record.get("status") not in {"primary", "scaffolded"}:
        errors.append(f"tracks.cf.primary: target {primary!r} has non-generating status {record.get('status')!r}")
    if record.get("driver") != "macos_excel":
        errors.append(f"targets.{primary}: CF track requires driver='macos_excel'")
    if "Darwin" not in (record.get("runs_on") or []):
        errors.append(f"targets.{primary}: CF track requires Darwin in runs_on")
    locale = record.get("locale")
    if not isinstance(locale, str) or not locale:
        errors.append(f"targets.{primary}: CF track requires a locale")

    golden_dir = (
        (repo_root / "tests" / "oracle" / "golden_cf")
        if primary == doc.get("primary")
        else (repo_root / "tests" / "oracle" / "variants" / primary / "golden_cf")
    )
    marker_path = golden_dir / "PROVENANCE.json"
    if not marker_path.exists():
        return errors
    try:
        marker = json.loads(marker_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        errors.append(f"tracks.cf: invalid {marker_path}: {exc}")
        return errors
    if not isinstance(marker, Mapping):
        return errors + [f"tracks.cf: {marker_path} must contain an object"]

    required_pairs = {
        "track": "cf",
        "target": primary,
        "status": record.get("status"),
        "locale": locale,
        "m365_sentinel": _M365_SENTINEL,
    }
    for key, expected in required_pairs.items():
        if marker.get(key) != expected:
            errors.append(f"tracks.cf: provenance {key!r} must be {expected!r}")
    classification = marker.get("classification")
    if classification not in {"active", "verified", "reference-only"}:
        errors.append("tracks.cf: provenance classification must be active, verified, or reference-only")
    if classification == "reference-only":
        if marker.get("verified") is True or marker.get("active_ctest") is True:
            errors.append("tracks.cf: reference-only provenance cannot be active or verified")
    elif marker.get("verified") is not True or marker.get("active_ctest") is not True:
        errors.append("tracks.cf: active provenance must set verified=true and active_ctest=true")
    product = marker.get("product")
    if (
        not isinstance(product, str)
        or not product.strip()
        or (classification != "reference-only" and "unknown" in product.lower())
    ):
        errors.append("tracks.cf: provenance product must identify the captured Excel build")
    capture_id = marker.get("capture_id")
    if not isinstance(capture_id, str) or not capture_id:
        errors.append("tracks.cf: provenance capture_id must be non-empty")
    if marker.get("all_suites_same_capture") is not True:
        errors.append("tracks.cf: provenance all_suites_same_capture must be true")

    actual_paths = sorted(golden_dir.glob("*.golden.json")) if golden_dir.is_dir() else []
    actual_suites = sorted(path.name[: -len(".golden.json")] for path in actual_paths)
    required = marker.get("required_suites")
    captured = marker.get("captured_suites")
    if not isinstance(required, list) or not required or not all(isinstance(v, str) and v for v in required):
        errors.append("tracks.cf: provenance required_suites must be a non-empty string list")
        required = []
    if sorted(required) != actual_suites:
        errors.append(f"tracks.cf: required_suites {sorted(required)!r} != on-disk suites {actual_suites!r}")
    if not isinstance(captured, list) or sorted(captured) != sorted(required):
        errors.append("tracks.cf: captured_suites must exactly match required_suites")
    inventory = marker.get("suite_inventory")
    if not isinstance(inventory, list):
        errors.append("tracks.cf: provenance suite_inventory must be a list")
        inventory = []
    by_suite: Dict[str, Mapping[str, Any]] = {}
    for item in inventory:
        if not isinstance(item, Mapping) or not isinstance(item.get("suite"), str):
            errors.append("tracks.cf: suite_inventory entries require a string suite")
            continue
        suite = str(item["suite"])
        if suite in by_suite:
            errors.append(f"tracks.cf: duplicate suite_inventory entry {suite!r}")
        by_suite[suite] = item
    if set(by_suite) != set(actual_suites):
        errors.append("tracks.cf: suite_inventory must cover every on-disk golden exactly once")
    total_cases = 0
    for path in actual_paths:
        suite = path.name[: -len(".golden.json")]
        item = by_suite.get(suite)
        if item is None:
            continue
        try:
            golden = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, json.JSONDecodeError) as exc:
            errors.append(f"tracks.cf: invalid golden {path}: {exc}")
            continue
        cases = golden.get("cases") if isinstance(golden, Mapping) else None
        count = len(cases) if isinstance(cases, list) else 0
        total_cases += count
        if not isinstance(golden, Mapping) or golden.get("name") != suite:
            errors.append(f"tracks.cf: golden {suite!r} has an unexpected name")
        if count <= 0 or item.get("case_count") != count:
            errors.append(f"tracks.cf: {suite!r} case_count does not match the golden")
        if item.get("capture_id") != capture_id:
            errors.append(f"tracks.cf: {suite!r} capture_id does not match the marker")
        digest = hashlib.sha256(path.read_bytes()).hexdigest()
        if item.get("sha256") != digest:
            errors.append(f"tracks.cf: {suite!r} sha256 does not match the golden")
    if marker.get("case_count") != total_cases:
        errors.append("tracks.cf: aggregate case_count does not match the golden files")
    return errors


def cf_active(doc: Mapping[str, Any], repo_root: Path = REPO_ROOT) -> bool:
    """Whether the manifest-selected CF capture is complete and verified."""

    primary = _cf_primary(doc)
    if not primary:
        return False
    marker_dir = (
        repo_root / "tests" / "oracle" / "golden_cf"
        if primary == doc.get("primary")
        else repo_root / "tests" / "oracle" / "variants" / primary / "golden_cf"
    )
    marker_path = marker_dir / "PROVENANCE.json"
    if not marker_path.exists():
        return False
    try:
        marker = json.loads(marker_path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return False
    return (
        isinstance(marker, Mapping)
        and marker.get("classification") in {"active", "verified"}
        and marker.get("verified") is True
        and marker.get("active_ctest") is True
        and not _cf_provenance_errors(doc, repo_root)
    )


def _cmd_check(args: argparse.Namespace) -> int:
    try:
        doc = _load_yaml(args.targets_file)
        errors = validate_targets(doc, args.repo_root)
    except RuntimeError as exc:
        print(f"oracle-provenance: {exc}", file=sys.stderr)
        return 2
    if errors:
        for error in errors:
            print(f"FAIL {error}")
        return 1
    print("PASS oracle target status/provenance")
    return 0


def _cmd_active(args: argparse.Namespace) -> int:
    try:
        doc = _load_yaml(args.targets_file)
    except RuntimeError as exc:
        print(f"oracle-provenance: {exc}", file=sys.stderr)
        return 2
    for name in active_variant_names(doc, args.repo_root):
        print(name)
    return 0


def _cmd_workbook_active(args: argparse.Namespace) -> int:
    try:
        doc = _load_yaml(args.targets_file)
    except RuntimeError as exc:
        print(f"oracle-provenance: {exc}", file=sys.stderr)
        return 2
    if workbook_active(doc, args.repo_root):
        print("active")
    return 0


def _cmd_cf_active(args: argparse.Namespace) -> int:
    try:
        doc = _load_yaml(args.targets_file)
    except RuntimeError as exc:
        print(f"oracle-provenance: {exc}", file=sys.stderr)
        return 2
    if cf_active(doc, args.repo_root):
        print("active")
        return 0
    errors = _cf_provenance_errors(doc, args.repo_root)
    for error in errors:
        print(f"FAIL {error}")
    return 1


def main(argv: Optional[Sequence[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("command", choices=("check", "active-variants", "workbook-active", "cf-active"))
    parser.add_argument("--targets-file", type=Path, default=DEFAULT_TARGETS)
    parser.add_argument("--repo-root", type=Path, default=REPO_ROOT)
    args = parser.parse_args(argv)
    if args.command == "check":
        return _cmd_check(args)
    if args.command == "workbook-active":
        return _cmd_workbook_active(args)
    if args.command == "cf-active":
        return _cmd_cf_active(args)
    return _cmd_active(args)


if __name__ == "__main__":  # pragma: no cover - CLI shim
    raise SystemExit(main())
