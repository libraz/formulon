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

A base listed as ``values:<path>`` also has its cell values, and whether each cell passes its data
validation, compared after Excel recalculates the whole book, so a formula Excel reads differently from
the source shows as a different number or verdict. Only bases
without volatile functions are listed that way. Excel keeps a book's cached values unless told otherwise,
so each such file is opened from a copy with fullCalcOnLoad set and every numeric formula cache replaced
by a sentinel, then recalculated per book and per sheet; a sentinel still read back fails the gate.

Lines of the form ``control:refused <path>`` or ``control:repaired <path>`` name files that are opened
as they are and must produce that outcome, so the gate proves it can fail.

The outputs go to a folder inside Excel's sandbox container, so opening them raises no file-access
prompt. Hold the shared Excel lock while this runs. The script never brings Excel to the front.

Usage: tools/oracle/.venv/bin/python tools/oracle/excel_open_gate.py [--cli build/bin/formulon_cli]
       [--source-root <checkout the CLI was built from>]
The run stops before Excel opens when the CLI is older than a file under the source root's src/.
Exit status: 0 when every output opens cleanly with the source's readings and every control behaves.
"""

from __future__ import annotations

import argparse
import re
import shutil
import struct
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
# Written over every numeric formula cache of a `values:` file; a recalculation that reaches the cell replaces it.
SENTINEL = 987654321.0
BRT_FMLA_NUM = 9
BRT_CALC_PROP = 157
# BrtCalcProp: recalcID, fAutoRecalc, cCalcCount, xnumDelta, cUserThreadCount, then the flag word.
CALC_PROP_FLAGS_AT = 24
CALC_FULL_CALC_ON_LOAD = 0x0001
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


def split_records(data: bytes):
    """(type, payload start, payload end) of each record of an .xlsb part."""

    at = 0
    while at < len(data):
        rtype = data[at] & 0x7F
        if data[at] & 0x80:
            rtype |= (data[at + 1] & 0x7F) << 7
            at += 1
        at += 1
        size, shift = 0, 0
        while True:
            byte = data[at]
            at += 1
            size |= (byte & 0x7F) << shift
            shift += 7
            if not byte & 0x80:
                break
        yield rtype, at, at + size
        at += size


def plant_xlsb_part(name: str, data: bytes) -> tuple[bytes, int]:
    out = bytearray(data)
    planted = 0
    for rtype, start, end in split_records(data):
        if name == "xl/workbook.bin" and rtype == BRT_CALC_PROP and end - start >= CALC_PROP_FLAGS_AT + 2:
            (flags,) = struct.unpack_from("<H", out, start + CALC_PROP_FLAGS_AT)
            struct.pack_into("<H", out, start + CALC_PROP_FLAGS_AT, flags | CALC_FULL_CALC_ON_LOAD)
        elif name.startswith("xl/worksheets/") and rtype == BRT_FMLA_NUM and end - start >= 16:
            struct.pack_into("<d", out, start + 8, SENTINEL)
            planted += 1
    return bytes(out), planted


def plant_xlsx_part(name: str, data: bytes) -> tuple[bytes, int]:
    text = data.decode("utf-8")
    if name == "xl/workbook.xml":
        if "<calcPr" in text:
            text = re.sub(r"<calcPr\b(?![^>]*fullCalcOnLoad)", '<calcPr fullCalcOnLoad="1"', text, count=1)
        else:
            text = text.replace("</workbook>", '<calcPr fullCalcOnLoad="1"/></workbook>', 1)
        return text.encode("utf-8"), 0
    if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
        return data, 0
    planted = 0

    def cell(m: re.Match) -> str:
        nonlocal planted
        kind = re.search(r'\bt="([^"]*)"', m.group(1))
        if (kind and kind.group(1) != "n") or "<f" not in m.group(2) or "<v>" not in m.group(2):
            return m.group(0)
        planted += 1
        return m.group(1) + re.sub(r"<v>[^<]*</v>", "<v>%r</v>" % SENTINEL, m.group(2), count=1)

    text = re.sub(r"(<c\b[^>]*(?<!/)>)(.*?</c>)", cell, text, flags=re.S)
    return text.encode("utf-8"), planted


def plant(source: Path, dest: Path) -> int:
    """Copies `source` to `dest` with fullCalcOnLoad set and every numeric formula cache replaced by SENTINEL;
    returns how many caches were replaced."""

    planted = 0
    xlsb = source.suffix == ".xlsb"
    with zipfile.ZipFile(source) as zin, zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            data = zin.read(info.filename)
            data, n = (plant_xlsb_part if xlsb else plant_xlsx_part)(info.filename, data)
            planted += n
            zout.writestr(info, data)
    return planted


def holds_sentinel(value) -> bool:
    if isinstance(value, (list, tuple)):
        return any(holds_sentinel(v) for v in value)
    return value == SENTINEL


def value_readings(app, book) -> dict:
    """Every sheet's used-range values after Excel recalculates the whole book and each sheet, and whether each
    used cell passes its data validation (Excel reports True for a cell without one)."""

    app.api.calculate_full()
    for sheet in book.sheets:
        sheet.api.calculate()
    out = {}
    for sheet in book.sheets:
        out[sheet.name] = safe(lambda s=sheet: s.used_range.value)
        valid = {}
        for cell in sheet.used_range:
            valid[cell.address] = safe(lambda c=cell: c.api.validation.value.get())
        out[sheet.name + ".validation"] = valid
    return out


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


def stale_cli(cli: Path, source_root: Path) -> str | None:
    """Why `cli` cannot vouch for the sources it was built from: it is missing, or older than the
    newest file under `source_root`/src."""

    if not cli.is_file():
        return "formulon_cli not found: %s" % cli
    files = [p for p in (source_root / "src").rglob("*") if p.is_file()]
    if not files:
        return "no sources under %s" % (source_root / "src")
    newest = max(files, key=lambda p: p.stat().st_mtime)
    if cli.stat().st_mtime < newest.stat().st_mtime:
        return "formulon_cli %s is older than %s; rebuild it first" % (cli, newest)
    return None


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument("--cli", default=str(ROOT / "build" / "bin" / "formulon_cli"))
    parser.add_argument(
        "--source-root",
        default=str(ROOT),
        help="the checkout the CLI was built from; its src/ must not be newer than the CLI",
    )
    args = parser.parse_args()
    print("checking formulon_cli against %s" % (Path(args.source_root) / "src"))
    stale = stale_cli(Path(args.cli), Path(args.source_root))
    if stale:
        print("FAIL " + stale)
        return 1
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
        if values:
            plant(base, staged)
        else:
            shutil.copy(base, staged)
        error, source = open_and_read(app, staged, sheets, "source_" + base.stem, values)
        if error is None and holds_sentinel(list(source.get("values", {}).values())):
            error = "a planted cache survived the recalculation"
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
            opened = out
            if values:
                opened = CONTAINER / ("planted_%s.%s" % (base.stem, ext))
                plant(out, opened)
            error, output = open_and_read(app, opened, sheets, "%s_%s" % (base.stem, ext), values)
            if error is None and holds_sentinel(list(output.get("values", {}).values())):
                error = "a planted cache survived the recalculation"
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
