"""Lands a staged oracle capture in the repository.

A target whose manifest `status:` is still `wanted` has no established
provenance, so `oracle_gen` / `workbook_oracle_gen` write its goldens to a
staging directory outside the tree (see `staging_dir_for`). This module is
the second half of that flow: it takes a staged capture, checks that the
capture actually earns repository coverage, copies the goldens into the
target's declared directory, writes the directory's `PROVENANCE.json`, and
flips the manifest `status:` to `scaffolded`.

Promotion is deliberately a separate command rather than a flag on the
generator. The generator knows what Excel said; only a human knows whether
the host that said it is the one the project wants to be measured against,
and the previous Windows capture entered the tree precisely because that
question was never asked out loud.

What is checked before anything is copied:

  * the staged directory carries a `PROVENANCE.candidate.json`, which the
    generator writes only after a complete, single-capture run whose Excel
    passed the M365 sentinel;
  * every suite the case directory declares is present in the capture;
  * each staged golden's recorded `capture_id` / product / locale agree
    with the candidate, and its SHA-256 matches the inventory;
  * the target's declared locale matches what Excel reported.

`--dry-run` performs every check and prints the plan without writing.

Usage:
    python3 tools/oracle/promote_capture.py --track workbook
                                            [--target NAME]
                                            [--from DIR]
                                            [--dry-run]
"""

from __future__ import annotations

import argparse
import hashlib
import json
import shutil
import sys
from pathlib import Path
from typing import Any, Dict, List, Mapping, Optional

try:  # pragma: no cover - trivial fallback
    from tools.oracle import provenance as provenance_mod
    from tools.oracle.workbook_oracle_gen import staging_dir_for
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import provenance as provenance_mod  # type: ignore
    from workbook_oracle_gen import staging_dir_for  # type: ignore

import yaml

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TARGETS_FILE = Path(__file__).resolve().parent / "targets.yaml"
CANDIDATE_NAME = "PROVENANCE.candidate.json"
PROVENANCE_NAME = "PROVENANCE.json"

# Where each track's goldens live once promoted, and which case directory
# defines the suite set that capture must cover.
TRACKS = {
    "workbook": {
        "golden_key": "golden_wb",
        "default_golden_dir": REPO_ROOT / "tests" / "oracle" / "golden_wb",
    },
}


def _load_targets(path: Path) -> Dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        doc = yaml.safe_load(handle) or {}
    if not isinstance(doc, dict) or not isinstance(doc.get("targets"), dict):
        raise RuntimeError(f"{path}: missing a `targets:` mapping")
    return doc


def _resolve_target(doc: Mapping[str, Any], track: str, name: Optional[str]) -> str:
    """Returns the target name, defaulting to the track's declared primary."""

    if name:
        if name not in doc["targets"]:
            known = ", ".join(sorted(doc["targets"]))
            raise RuntimeError(f"unknown target {name!r}; known targets: {known}")
        return name
    tracks = doc.get("tracks")
    if isinstance(tracks, dict):
        record = tracks.get(track)
        if isinstance(record, dict) and isinstance(record.get("primary"), str):
            return record["primary"]
    primary = doc.get("primary")
    if isinstance(primary, str):
        return primary
    raise RuntimeError(f"no primary target declared for track {track!r}")


def _read_json(path: Path) -> Any:
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except OSError as exc:
        raise RuntimeError(f"{path}: {exc}") from exc
    except json.JSONDecodeError as exc:
        raise RuntimeError(f"{path}: not valid JSON ({exc})") from exc


def check_capture(staged_dir: Path, target_name: str, target_record: Mapping[str, Any]) -> Dict[str, Any]:
    """Validates a staged capture and returns its candidate document.

    Raises `RuntimeError` with a single actionable message on the first
    problem found. Nothing here writes.
    """

    if not staged_dir.is_dir():
        raise RuntimeError(f"no staged capture at {staged_dir} (run the generator for this target first)")
    candidate_path = staged_dir / CANDIDATE_NAME
    if not candidate_path.is_file():
        raise RuntimeError(
            f"{staged_dir}: no {CANDIDATE_NAME}. The generator writes it only after a complete run whose "
            "Excel passed the M365 sentinel, so this capture is partial or came from a pre-M365 host."
        )
    candidate = _read_json(candidate_path)
    if not isinstance(candidate, dict):
        raise RuntimeError(f"{candidate_path}: expected a JSON object")

    if candidate.get("target") != target_name:
        raise RuntimeError(f"{candidate_path}: capture is for target {candidate.get('target')!r}, not {target_name!r}")
    if candidate.get("verified") is not True:
        raise RuntimeError(f"{candidate_path}: capture is not marked verified")
    if candidate.get("m365_sentinel") != provenance_mod._M365_SENTINEL:
        raise RuntimeError(
            f"{candidate_path}: M365 sentinel is {candidate.get('m365_sentinel')!r}, "
            f"expected {provenance_mod._M365_SENTINEL!r}"
        )
    product = candidate.get("product")
    if not isinstance(product, str) or not product.strip():
        raise RuntimeError(f"{candidate_path}: missing the Excel product/version string")
    if "unknown" in product.lower():
        raise RuntimeError(
            f"{candidate_path}: product is {product!r}. A capture that cannot name its Excel build is "
            "exactly the shape of the historical Office 2019 mix-up and is refused."
        )
    declared_locale = target_record.get("locale")
    if candidate.get("locale") != declared_locale:
        raise RuntimeError(
            f"{candidate_path}: Excel reported locale {candidate.get('locale')!r} but target "
            f"{target_name!r} declares {declared_locale!r}. Capture under the matching target."
        )

    required = candidate.get("required_suites")
    captured = candidate.get("captured_suites")
    if not isinstance(required, list) or not required:
        raise RuntimeError(f"{candidate_path}: missing required_suites")
    if not isinstance(captured, list) or sorted(captured) != sorted(required):
        missing = sorted(set(required) - set(captured or []))
        raise RuntimeError(f"{candidate_path}: capture is missing suite(s): {', '.join(missing) or 'unknown'}")

    inventory = candidate.get("suite_inventory")
    if not isinstance(inventory, list) or not inventory:
        raise RuntimeError(f"{candidate_path}: missing suite_inventory")
    capture_id = candidate.get("capture_id")
    if not isinstance(capture_id, str) or not capture_id:
        raise RuntimeError(f"{candidate_path}: missing capture_id")

    # The inventory is the capture's own account of itself; re-derive it
    # from the files so a hand-edited candidate cannot promote a golden it
    # does not describe.
    for item in inventory:
        if not isinstance(item, dict) or not isinstance(item.get("suite"), str):
            raise RuntimeError(f"{candidate_path}: malformed suite_inventory entry")
        suite = item["suite"]
        golden_path = staged_dir / f"{suite}.golden.json"
        if not golden_path.is_file():
            raise RuntimeError(f"{staged_dir}: inventory names {suite} but {golden_path.name} is absent")
        digest = hashlib.sha256(golden_path.read_bytes()).hexdigest()
        if item.get("sha256") != digest:
            raise RuntimeError(f"{golden_path.name}: SHA-256 disagrees with the inventory (edited after capture?)")
        document = _read_json(golden_path)
        environment = document.get("environment") if isinstance(document, dict) else None
        if not isinstance(environment, dict) or environment.get("capture_id") != capture_id:
            raise RuntimeError(f"{golden_path.name}: belongs to a different capture than the candidate")
        if environment.get("excel_version") != product or environment.get("excel_locale") != candidate["locale"]:
            raise RuntimeError(f"{golden_path.name}: environment disagrees with the candidate's product/locale")

    return candidate


def _promoted_provenance(candidate: Mapping[str, Any]) -> Dict[str, Any]:
    """Rewrites a verified candidate as the directory's active marker."""

    promoted = dict(candidate)
    promoted["classification"] = "active"
    promoted["verified"] = True
    promoted["active_ctest"] = True
    promoted["status"] = "scaffolded"
    promoted.pop("reason", None)
    return promoted


def _set_target_status(targets_file: Path, target_name: str, status: str) -> bool:
    """Rewrites one target's `status:` line in place. Returns True if changed.

    Edits the text rather than round-tripping the YAML: the manifest is a
    heavily commented document and a dump would discard every comment that
    explains why the entries are shaped the way they are.
    """

    lines = targets_file.read_text(encoding="utf-8").splitlines(keepends=True)
    in_target = False
    for index, line in enumerate(lines):
        stripped = line.strip()
        if stripped.startswith(f"{target_name}:"):
            in_target = True
            continue
        if in_target:
            # A new top-level or sibling key at two-space indent ends the block.
            if stripped and not line.startswith("    ") and not stripped.startswith("#"):
                break
            if stripped.startswith("status:"):
                indent = line[: len(line) - len(line.lstrip())]
                if stripped == f"status: {status}":
                    return False
                lines[index] = f"{indent}status: {status}\n"
                targets_file.write_text("".join(lines), encoding="utf-8")
                return True
    raise RuntimeError(f"{targets_file}: no `status:` line found under target {target_name!r}")


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    parser.add_argument("--track", default="workbook", choices=sorted(TRACKS), help="Oracle track to promote.")
    parser.add_argument("--target", default=None, metavar="NAME", help="Target name (default: the track's primary).")
    parser.add_argument(
        "--from",
        dest="staged_dir",
        type=Path,
        default=None,
        metavar="DIR",
        help="Staged capture directory (default: the generator's staging path for this track/target).",
    )
    parser.add_argument("--targets-file", type=Path, default=DEFAULT_TARGETS_FILE)
    parser.add_argument("--dry-run", action="store_true", help="Check and print the plan; write nothing.")
    args = parser.parse_args(argv)

    try:
        targets_doc = _load_targets(args.targets_file)
        target_name = _resolve_target(targets_doc, args.track, args.target)
    except RuntimeError as exc:
        print(f"oracle-promote: {exc}", file=sys.stderr)
        return 2
    target_record = targets_doc["targets"][target_name]

    staged_dir = args.staged_dir or staging_dir_for(args.track, target_name)
    try:
        candidate = check_capture(staged_dir, target_name, target_record)
    except RuntimeError as exc:
        print(f"oracle-promote: {exc}", file=sys.stderr)
        return 1

    golden_dir = TRACKS[args.track]["default_golden_dir"]
    suites = sorted(candidate["captured_suites"])
    print(f"oracle-promote: track={args.track} target={target_name}")
    print(f"  from:    {staged_dir}")
    print(f"  into:    {golden_dir.relative_to(REPO_ROOT)}")
    print(f"  Excel:   {candidate['product']} ({candidate['locale']})")
    print(f"  capture: {candidate['capture_id']}")
    print(f"  suites:  {len(suites)} ({', '.join(suites)})")
    if args.dry_run:
        print("  dry-run: nothing written")
        return 0

    golden_dir.mkdir(parents=True, exist_ok=True)
    for suite in suites:
        name = f"{suite}.golden.json"
        shutil.copyfile(staged_dir / name, golden_dir / name)
        print(f"  -> {(golden_dir / name).relative_to(REPO_ROOT)}")
    provenance_path = golden_dir / PROVENANCE_NAME
    provenance_path.write_text(
        json.dumps(_promoted_provenance(candidate), indent=2, ensure_ascii=False) + "\n",
        encoding="utf-8",
    )
    print(f"  -> {provenance_path.relative_to(REPO_ROOT)}")

    if _set_target_status(args.targets_file, target_name, "scaffolded"):
        print(f"  -> {args.targets_file.relative_to(REPO_ROOT)} (status: wanted -> scaffolded)")

    print("")
    print("Next: verify the promotion is coverage-eligible, then review the diff.")
    print("  tools/oracle/.venv/bin/python tools/oracle/provenance.py check")
    print(f"  tools/oracle/.venv/bin/python tools/oracle/provenance.py {args.track}-active")
    print("  make oracle-verify")
    return 0


if __name__ == "__main__":
    sys.exit(main())
