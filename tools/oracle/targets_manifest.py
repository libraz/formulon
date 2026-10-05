"""Loading and selection of oracle targets from targets.yaml, and the host label."""

from __future__ import annotations

import platform
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

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
        chosen: List[Tuple[str, Dict[str, Any]]] = [(n, t) for n, t in sorted(targets.items()) if _runs_on_current(t)]
        if not chosen:
            raise RuntimeError(f"no targets in targets.yaml declare runs_on: [{platform.system()}]")
        return chosen
    primary = doc.get("primary")
    if not isinstance(primary, str) or primary not in targets:
        raise RuntimeError("targets.yaml is missing a valid `primary:` entry")
    return [(primary, targets[primary])]
