#!/usr/bin/env python3
"""Cross-platform oracle CLI dispatcher.

A thin orchestrator over the oracle drivers. It (1) loads targets.yaml,
(2) selects targets that match the current platform, (3) either prints
them, runs preflight checks, generates goldens, or guides external
contributors through a one-command donation flow.

Subcommands:
    cli.py list                              # print available targets
    cli.py gen [--target NAME] [--all]       # delegate to oracle_gen
    cli.py workbook [--target NAME]          # workbook track (pivot / print)
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
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional

# Local imports — accept both `python3 tools/oracle/cli.py` (no package)
# and `python3 -m tools.oracle.cli` (package-style).
try:  # pragma: no cover - trivial fallback
    from tools.oracle import oracle_gen, workbook_oracle_gen
    from tools.oracle.contribute import _cmd_contribute
    from tools.oracle.preflight import _cmd_setup
    from tools.oracle.targets_manifest import DEFAULT_TARGETS_FILE, _load_targets, _platform_label, _select_targets
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import oracle_gen  # type: ignore
    import workbook_oracle_gen  # type: ignore
    from contribute import _cmd_contribute  # type: ignore
    from preflight import _cmd_setup  # type: ignore
    from targets_manifest import DEFAULT_TARGETS_FILE, _load_targets, _platform_label, _select_targets  # type: ignore


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


def _cmd_workbook(args: argparse.Namespace) -> int:
    """Generates goldens for the workbook oracle track.

    Delegates to `workbook_oracle_gen.main`. When `--target` is omitted
    that entry point auto-detects the target from the host OS (a Windows
    / WSL2 host -> the win-365-ja_JP primary, a macOS host -> the
    mac-365-ja_JP variant), reading the `tracks.workbook` section of
    targets.yaml. The subcommand exists so the workbook track is
    reachable through the same CLI as the formula `gen` flow.
    """

    gen_argv: List[str] = ["--targets-file", str(args.targets_file)]
    if args.target:
        gen_argv.extend(["--target", args.target])
    if args.suite:
        for s in args.suite:
            gen_argv.extend(["--suite", s])
    if getattr(args, "golden_dir", None):
        gen_argv.extend(["--golden-dir", str(args.golden_dir)])
    if args.visible:
        gen_argv.append("--visible")
    print("[oracle-cli] track=workbook")
    return workbook_oracle_gen.main(gen_argv)


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

    p_wb = sub.add_parser(
        "workbook",
        help="Generate goldens for the workbook oracle track (pivot / print).",
    )
    p_wb.add_argument(
        "--target",
        default=None,
        help="Target name (default: auto-detected from the host OS).",
    )
    p_wb.add_argument(
        "--suite",
        action="append",
        default=None,
        metavar="NAME",
        help="Restrict to the named suite(s); forwarded to workbook_oracle_gen.",
    )
    p_wb.add_argument(
        "--golden-dir",
        default=None,
        metavar="DIR",
        help=(
            "Write goldens here instead of the per-target path. A target "
            "with status=wanted always stages outside the repository; "
            "omit this to accept the default staging directory."
        ),
    )
    p_wb.add_argument("--visible", action="store_true")
    p_wb.set_defaults(func=_cmd_workbook)

    p_setup = sub.add_parser(
        "setup",
        help="Verify the host can drive its target oracle.",
    )
    p_setup.add_argument(
        "--target",
        default=None,
        help=("Verify just one target by name; defaults to every target whose runs_on declares the current host."),
    )
    p_setup.set_defaults(func=_cmd_setup)

    p_contrib = sub.add_parser(
        "contribute",
        help=("Contributor onramp: thank-you banner + preflight + gen + push/PR instructions for one variant target."),
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
