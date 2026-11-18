"""Suite names of the external corpus fixtures, with no spreadsheet library.

The importer names a fixture sheet's suite with `suite_name`, and the
divergence registry check derives the same name for every sheet it reads, so
both import it from here and the check runs with PyYAML alone.
"""

from __future__ import annotations

import re
from pathlib import Path

_FILENAME_SCRUB = re.compile(r"[^A-Za-z0-9._-]+")


def scrub_filename(part: str) -> str:
    """Folds a path component into filesystem-safe characters."""

    return _FILENAME_SCRUB.sub("_", part).strip("._") or "sheet"


def suite_name(rel_source: str, sheet_title: str) -> str:
    """The suite of one fixture sheet; `<suite>.<addr>` is a case's id.

    `tests/ironcalc_divergence.yaml` keys its entries by that id, and the
    loaded-file gate derives the same name for every sheet it reads.
    """

    xlsx_stem = scrub_filename(Path(rel_source).stem)
    sheet_slug = scrub_filename(sheet_title)
    rel_dir = scrub_filename(str(Path(rel_source).parent).replace("/", "__"))
    if rel_dir in ("", "."):
        return f"ironcalc_{xlsx_stem}_{sheet_slug}"
    return f"ironcalc_{rel_dir}_{xlsx_stem}_{sheet_slug}"
