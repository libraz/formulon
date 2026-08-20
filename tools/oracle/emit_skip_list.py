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
`--target` overrides it outright.

Usage:
    python3 tools/oracle/emit_skip_list.py --out PATH [--track NAME] [--target NAME]
"""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

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


def emit(divergence: Path, target: str, out: Path, *, track: str = "workbook") -> int:
    """Writes `{case_id: reason}` for `target`; returns the entry count."""

    # A `suite:` selector removes every case in that suite, and only the
    # formula track has the case files to expand one -- the workbook
    # registry selects by id. Leaving the suites out would silently emit
    # a shorter list than the registry asks for.
    suites = None
    if track == "formula" and FORMULA_CASES_DIR.is_dir():
        suites = case_schema.discover_suites(FORMULA_CASES_DIR)
    skips = _load_divergence_skips(divergence, target, suites=suites)
    out.parent.mkdir(parents=True, exist_ok=True)
    payload = json.dumps({"target": target, "skips": skips}, indent=2, ensure_ascii=False) + "\n"
    # Only rewrite on a real change so CMake's configure dependency does not
    # retrigger a build every time it runs.
    if not out.is_file() or out.read_text(encoding="utf-8") != payload:
        out.write_text(payload, encoding="utf-8")
    return len(skips)


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
        default=None,
        help="Target the `applies_to` scoping resolves against; defaults to the track's primary.",
    )
    parser.add_argument("--targets-file", type=Path, default=Path(__file__).resolve().parent / "targets.yaml")
    parser.add_argument("--out", type=Path, required=True)
    args = parser.parse_args(argv)
    try:
        target = args.target or track_primary(_load_targets(args.targets_file), args.track)
        count = emit(args.divergence, target, args.out, track=args.track)
    except (OSError, RuntimeError) as exc:
        print(f"emit-skip-list: {exc}", file=sys.stderr)
        return 1
    print(f"emit-skip-list: {count} skip-oracle case(s) for {target} ({args.track} track)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
