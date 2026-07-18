#!/usr/bin/env python3
"""Builds a temporary divergence file with selected cases unskipped.

This is for oracle re-probes: keep the committed `tests/divergence.yaml`
unchanged, but run a suite with one or more historical skip-oracle entries
enabled again to see whether a newer Excel / xlwings build can now capture
them.
"""

from __future__ import annotations

import argparse
from pathlib import Path
from typing import List, Optional

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_DIVERGENCE = REPO_ROOT / "tests" / "divergence.yaml"


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--input",
        type=Path,
        default=DEFAULT_DIVERGENCE,
        help="Source divergence.yaml; defaults to tests/divergence.yaml.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        required=True,
        help="Temporary divergence YAML to write.",
    )
    parser.add_argument(
        "--allow-case",
        action="append",
        default=[],
        metavar="ID",
        help="Case id to remove from the temporary divergence file.",
    )
    args = parser.parse_args(argv)

    try:
        import yaml  # type: ignore
    except ImportError as exc:
        raise SystemExit("PyYAML is required; run make oracle-setup first") from exc

    doc = yaml.safe_load(args.input.read_text(encoding="utf-8")) or {}
    entries = doc.get("entries") or []
    if not isinstance(entries, list):
        raise SystemExit(f"{args.input}: `entries` must be a list")

    allowed = set(args.allow_case)
    kept = [entry for entry in entries if not (isinstance(entry, dict) and entry.get("id") in allowed)]
    removed = len(entries) - len(kept)
    doc["entries"] = kept
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(yaml.safe_dump(doc, sort_keys=False, allow_unicode=True), encoding="utf-8")
    print(f"wrote {args.output} ({removed} case(s) unskipped)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
