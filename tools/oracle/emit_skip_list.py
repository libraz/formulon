#!/usr/bin/env python3
"""Emits a track's `skip-oracle` case ids as JSON for the C++ verifiers.

`mode: skip-oracle` is a *verification-time* policy: it says "we know we
differ here, do not assert it". The generators used to bake it into each
golden's `skipped` field, which made the policy retroactive only through a
fresh Excel capture -- registering a newly adjudicated divergence left the
verifier red until someone with Excel re-ran the suite. That is the same
deadlock the capture side had, from the other end.

The C++ verifiers cannot read YAML (the tree carries pugixml, not a YAML
parser), so the registry is projected to JSON at configure time and read
back through the JSON reader the oracle runners already use. Goldens keep
their own `skipped` field for captures that predate this.

Each oracle track resolves the registry against its own primary target,
because `applies_to` scoping is per target and the formula and workbook
tracks do not share one. `--track` picks which primary to resolve against;
`--target` overrides it outright, and may be repeated when one binary
loads cases from several variant tags at once.

Usage:
    python3 tools/oracle/emit_skip_list.py --out PATH [--track NAME] [--target NAME ...]
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Dict

try:  # pragma: no cover - trivial fallback
    from tools.oracle import case_schema
    from tools.oracle.oracle_gen import _load_divergence_skips
    from tools.oracle.workbook_oracle_gen import _load_targets, _workbook_primary
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import case_schema  # type: ignore
    from oracle_gen import _load_divergence_skips  # type: ignore
    from workbook_oracle_gen import _load_targets, _workbook_primary  # type: ignore

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DIVERGENCE = REPO_ROOT / "tests" / "divergence.yaml"
FORMULA_CASES_DIR = REPO_ROOT / "tests" / "oracle" / "cases"

TRACKS = ("formula", "workbook")


def track_primary(targets_doc: dict, track: str) -> str:
    """Return the primary target `track` resolves `applies_to` against.

    A track may declare its own primary in `targets.yaml`; the formula
    track does not, and inherits the manifest's global `primary:` -- which
    is what "the formula track stays on the Mac primary" means there.
    """

    if track == "workbook":
        return _workbook_primary(targets_doc)
    tracks = targets_doc.get("tracks")
    if isinstance(tracks, dict) and isinstance(tracks.get(track), dict):
        primary = tracks[track].get("primary")
        if isinstance(primary, str) and primary:
            return primary
    primary = targets_doc.get("primary")
    if isinstance(primary, str) and primary:
        return primary
    raise RuntimeError(f"targets.yaml: no primary target for the {track} track")


def _resolve_target_suites(track: str):
    # A `suite:` selector removes every case in that suite, and only the
    # formula track has the case files to expand one -- the workbook
    # registry selects by id. Leaving the suites out would silently emit
    # a shorter list than the registry asks for.
    if track == "formula" and FORMULA_CASES_DIR.is_dir():
        return case_schema.discover_suites(FORMULA_CASES_DIR)
    return None


def _skips_for_target(
    divergence: Path,
    target: str,
    *,
    track: str,
    targets_doc: dict | None,
    suites,
) -> Dict[str, str]:
    """Returns the case-id -> reason skip map for one target.

    When `targets_doc` is given and `target` is not that track's own
    primary, `target` is a variant: its per-variant override file
    (`tests/oracle/targets/<target>/divergence.yaml`) is merged on top of
    the primary registry, entries there winning on key collision, mirroring
    `workbook_oracle_gen._resolve_skips`. Omitting `targets_doc` (the
    default) keeps a primary-target call resolving from the primary
    registry alone.
    """

    skips = _load_divergence_skips(divergence, target, suites=suites)
    if targets_doc is not None and target != track_primary(targets_doc, track):
        variant_div = REPO_ROOT / "tests" / "oracle" / "targets" / target / "divergence.yaml"
        if variant_div.exists():
            skips.update(_load_divergence_skips(variant_div, target, suites=suites))
    return skips


def emit(
    divergence: Path,
    target: str,
    out: Path,
    *,
    track: str = "workbook",
    targets_doc: dict | None = None,
) -> int:
    """Writes `{case_id: reason}` for `target`; returns the entry count."""

    suites = _resolve_target_suites(track)
    skips = _skips_for_target(divergence, target, track=track, targets_doc=targets_doc, suites=suites)
    out.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps({"target": target, "skips": skips}, indent=2, ensure_ascii=False) + "\n"
    # Only rewrite on a real change so CMake's configure dependency does not
    # retrigger a build every time it runs.
    if not out.is_file() or out.read_text(encoding="utf-8") != payload:
        out.write_text(payload, encoding="utf-8")
    return len(skips)


def emit_many(
    divergence: Path,
    targets: list[str],
    out: Path,
    *,
    track: str = "workbook",
    targets_doc: dict | None = None,
) -> int:
    """Writes the union of `targets`' skip maps to one file.

    A binary that loads cases from several variant tags at once (the
    formula-track variant oracle) still reads a single
    `FORMULON_ORACLE_SKIP_FILE`, so each tag's own resolution (primary
    registry + that tag's `applies_to` scoping and divergence override) has
    to land in one JSON payload. A case id shared by two tags is a
    pre-existing ambiguity the registry does not scope past target names;
    the later target in `targets` wins.
    """

    suites = _resolve_target_suites(track)
    merged: Dict[str, str] = {}
    for target in targets:
        merged.update(_skips_for_target(divergence, target, track=track, targets_doc=targets_doc, suites=suites))
    out.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps({"targets": list(targets), "skips": merged}, indent=2, ensure_ascii=False) + "\n"
    if not out.is_file() or out.read_text(encoding="utf-8") != payload:
        out.write_text(payload, encoding="utf-8")
    return len(merged)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--divergence", type=Path, default=DEFAULT_DIVERGENCE)
    parser.add_argument(
        "--track",
        default="workbook",
        choices=TRACKS,
        help="Oracle track whose primary target the `applies_to` scoping resolves against.",
    )
    parser.add_argument(
        "--target",
        action="append",
        default=None,
        help=(
            "Target the `applies_to` scoping resolves against; defaults to the track's "
            "primary. Repeatable for a binary that loads several variant tags at once, "
            "in which case their skip maps are unioned into one file."
        ),
    )
    parser.add_argument("--targets-file", type=Path, default=Path(__file__).resolve().parent / "targets.yaml")
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        targets_doc = _load_targets(args.targets_file)
        targets = args.target or [track_primary(targets_doc, args.track)]
        if len(targets) == 1:
            count = emit(args.divergence, targets[0], args.out, track=args.track, targets_doc=targets_doc)
        else:
            count = emit_many(args.divergence, targets, args.out, track=args.track, targets_doc=targets_doc)
    except (OSError, RuntimeError) as exc:
        print(f"emit-skip-list: {exc}", file=sys.stderr)
        return 1
    print(f"emit-skip-list: {count} skip-oracle case(s) for {', '.join(targets)} ({args.track} track)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
