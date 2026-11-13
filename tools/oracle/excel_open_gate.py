#!/usr/bin/env python3
"""Converts Excel-made workbooks with formulon_cli and checks that Mac Excel opens every output intact.

Each base listed in ``excel_open_gate.txt`` (paths relative to ``tests/fixtures``) is loaded and saved
by ``formulon_cli recalc`` as ``.xlsx`` and as ``.xlsb``, and each output is opened in Excel with alerts
off. An output fails the gate when Excel refuses it, when Excel repairs it (repair marker in the
window title, or a sheet count differing from the source), or when Excel sees less of a feature than
it sees in the source base: the readings below are taken from the source and from the output and
must match.

Readings come from Excel's object model where the Mac object model exposes the feature (page setup,
autofilter mode, tables, shapes, charts). The rest (notes, PivotTables, threaded comments, sparklines,
protected ranges, slicers, page breaks, header/footer) is read by letting Excel re-save the opened book
as .xlsx and counting the parts or elements Excel wrote.

A base listed as ``values:<path>`` also has its cell values compared after Excel recalculates the whole
book, so a formula Excel reads differently from the source shows as a different number. Only bases
without volatile functions are listed that way.

Lines of the form ``control:refused <path>`` or ``control:repaired <path>`` name files that are opened
as they are and must produce that outcome, so the gate proves it can fail.

The outputs go to a folder inside Excel's sandbox container, so opening them raises no file-access
prompt. Hold the shared Excel lock while this runs. The script never brings Excel to the front.

Usage: tools/oracle/.venv/bin/python tools/oracle/excel_open_gate.py [--cli build/bin/formulon_cli]
Exit status: 0 when every output opens cleanly with the source's readings and every control behaves.
"""

from __future__ import annotations

import argparse
import re
import shutil
import subprocess
import sys
import zipfile
from pathlib import Path

import xlwings as xw

sys.path.insert(0, str(Path(__file__).resolve().parent))
from drivers.excel_alerts import alerts_off  # noqa: E402

ROOT = Path(__file__).resolve().parents[2]
MANIFEST = Path(__file__).with_name("excel_open_gate.txt")
CONTAINER = Path.home() / "Library/Containers/com.microsoft.Excel/Data/formulon_open_gate"
REPAIR_MARKERS = ("Repaired", "修復")
XLSX_FORMAT = 51
# Parts and elements counted in an Excel re-save, for features the object model does not expose.
RESAVE_PATTERNS = {
    "notes": (r"xl/comments\d+\.xml$", rb"<comment "),
    "tables": (r"xl/tables/.*\.xml$", rb"<table "),
    "pivots": (r"xl/pivotTables/.*\.xml$", rb"<pivotTableDefinition "),
    "threaded_comments": (r"xl/threadedComments/.*\.xml$", rb"<threadedComment "),
    "sparkline_groups": (r"xl/worksheets/sheet\d+\.xml$", rb"<x14:sparklineGroup "),
    "protected_ranges": (r"xl/worksheets/sheet\d+\.xml$", rb"<protectedRange "),
    "slicers": (r"xl/slicers/.*\.xml$", rb"<slicer "),
    "page_breaks": (r"xl/worksheets/sheet\d+\.xml$", rb"<brk "),
    "header_footer": (r"xl/worksheets/sheet\d+\.xml$", rb"<headerFooter"),
}


def manifest() -> tuple[list[tuple[Path, bool]], list[tuple[str, Path]]]:
    """(bases, controls): each base with whether its recalculated values are compared."""

    bases, controls = [], []
    for line in MANIFEST.read_text(encoding="utf-8").splitlines():
        line = line.split("#", 1)[0].strip()
        if not line:
            continue
        if line.startswith("control:"):
            outcome, path = line[len("control:") :].split(None, 1)
            controls.append((outcome, ROOT / "tests" / "fixtures" / path.strip()))
        elif line.startswith("values:"):
            bases.append((ROOT / "tests" / "fixtures" / line[len("values:") :].strip(), True))
        else:
            bases.append((ROOT / "tests" / "fixtures" / line, False))
    return bases, controls


def source_sheet_count(path: Path) -> int:
    with zipfile.ZipFile(path) as z:
        return z.read("xl/workbook.xml").count(b"<sheet ")


def close(book_name: str) -> None:
    script = 'tell application "Microsoft Excel" to close workbook "%s" saving no' % book_name
    subprocess.run(["osascript", "-e", script], check=False)


def safe(read):
    try:
        value = read()
        return value if isinstance(value, (int, float, str, bool, type(None))) else str(value)
    except Exception:  # an object-model property that raises is recorded, not fatal
        return "n/a"


def om_readings(book) -> dict:
    """Object-model readings of the features Excel exposes, per sheet."""

    out = {}
    for sheet in book.sheets:
        api = sheet.api
        ps = safe(lambda: api.page_setup_object)
        reading = {
            "autofilter_mode": safe(lambda: api.autofilter_mode.get()),
            "tables": safe(lambda: len(api.list_objects.get())),
            "shapes": safe(lambda: len(sheet.shapes)),
            "charts": safe(lambda: len(sheet.charts)),
        }
        if ps != "n/a":
            p = api.page_setup_object
            for name in (
                "left_margin",
                "right_margin",
                "top_margin",
                "bottom_margin",
                "header_margin",
                "footer_margin",
                "orientation",
                "zoom",
                "fit_to_pages_wide",
                "fit_to_pages_tall",
                "first_page_number",
                "center_horizontally",
                "print_gridlines",
                "print_headings",
                "center_header",
                "left_footer",
            ):
                reading["page_" + name] = safe(lambda n=name: getattr(p, n).get())
        out[sheet.name] = reading
    return out


def resave_readings(book, stem: str) -> dict:
    """Counts of parts and elements in Excel's own .xlsx re-save of the opened book."""

    target = CONTAINER / (stem + ".resave.xlsx")
    if target.exists():
        target.unlink()
    try:
        book.api.save_workbook_as(filename=str(target), file_format=XLSX_FORMAT, timeout=120)
    except Exception as e:
        return {"resave": "failed: " + str(e).splitlines()[0][:60]}
    counts = {}
    with zipfile.ZipFile(target) as z:
        for key, (part_pattern, needle) in RESAVE_PATTERNS.items():
            counts[key] = sum(z.read(n).count(needle) for n in z.namelist() if re.match(part_pattern, n))
    return counts


def value_readings(app, book) -> dict:
    """Every sheet's used-range values after Excel recalculates the whole book."""

    app.api.calculate_full()
    return {sheet.name: safe(lambda s=sheet: s.used_range.value) for sheet in book.sheets}


def open_and_read(app, path: Path, sheets: int | None, stem: str, values: bool = False):
    """(error, readings): error is None when Excel opens `path` cleanly."""

    before = {b.name for b in app.books}
    error = None
    readings = {}
    with alerts_off(app):
        try:
            app.books.open(str(path))
        except Exception as e:  # xlwings reports a refusal as an AppleScript error
            error = "refused: " + str(e).splitlines()[0][:80]
    opened = [b for b in app.books if b.name not in before]
    for book in opened:
        if error is None:
            if any(marker in book.name for marker in REPAIR_MARKERS):
                error = "repaired: " + book.name
            elif sheets is not None and len(book.sheets) != sheets:
                error = "repaired: %d sheets, source has %d" % (len(book.sheets), sheets)
            else:
                readings = {"om": om_readings(book)}
                if values:
                    with alerts_off(app):
                        readings["values"] = value_readings(app, book)
                with alerts_off(app):
                    readings["resave"] = resave_readings(book, stem)
        for b in [b for b in app.books if b.name not in before]:
            close(b.name)
    if error is None and not opened:
        error = "refused: no workbook opened"
    return error, readings


def diff(source: dict, output: dict) -> list[str]:
    lost = []
    for sheet, reading in source.get("om", {}).items():
        theirs = output.get("om", {}).get(sheet, {})
        for key, value in reading.items():
            if theirs.get(key) != value:
                lost.append("%s.%s: source %r, output %r" % (sheet, key, value, theirs.get(key)))
    for key, value in source.get("resave", {}).items():
        if output.get("resave", {}).get(key) != value:
            lost.append("resave.%s: source %r, output %r" % (key, value, output.get("resave", {}).get(key)))
    for sheet, value in source.get("values", {}).items():
        if output.get("values", {}).get(sheet) != value:
            lost.append("values.%s: source %r, output %r" % (sheet, value, output.get("values", {}).get(sheet)))
    return lost


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--cli", default=str(ROOT / "build" / "bin" / "formulon_cli"))
    args = parser.parse_args()
    if CONTAINER.exists():
        shutil.rmtree(CONTAINER)
    CONTAINER.mkdir(parents=True)
    app = xw.apps.active
    bases, controls = manifest()
    failures = 0
    for outcome, control in controls:
        staged = CONTAINER / control.name
        shutil.copy(control, staged)
        error, _ = open_and_read(app, staged, None, "control_" + control.stem)
        got = error.split(":", 1)[0] if error else "opened"
        ok = got == outcome
        print("%s control %s: expected %s, got %s" % ("ok  " if ok else "FAIL", control.name, outcome, error or got))
        failures += 0 if ok else 1
    for base, values in bases:
        sheets = source_sheet_count(base)
        staged = CONTAINER / ("source_" + base.name)
        shutil.copy(base, staged)
        error, source = open_and_read(app, staged, sheets, "source_" + base.stem, values)
        if error:
            print("FAIL %s: the source itself does not open cleanly: %s" % (base.relative_to(ROOT), error))
            failures += 1
            continue
        for ext in ("xlsx", "xlsb"):
            out = CONTAINER / ("%s.%s" % (base.stem, ext))
            r = subprocess.run(
                [args.cli, "recalc", "--quiet", "--format", ext, str(base), "-o", str(out)],
                capture_output=True,
                text=True,
            )
            if r.returncode != 0:
                print("FAIL %s -> %s: formulon_cli: %s" % (base.relative_to(ROOT), ext, r.stderr.strip()))
                failures += 1
                continue
            error, output = open_and_read(app, out, sheets, "%s_%s" % (base.stem, ext), values)
            problems = [error] if error else ["feature lost: " + d for d in diff(source, output)]
            status = "FAIL" if problems else "ok  "
            print("%s %s -> %s%s" % (status, base.relative_to(ROOT), ext, "".join("\n       " + p for p in problems)))
            failures += 1 if problems else 0
    if not app.display_alerts:
        print("FAIL Excel alerts were left off")
        failures += 1
    shutil.rmtree(CONTAINER)
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
