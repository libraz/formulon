#!/usr/bin/env python3
"""Captures chart-track goldens by driving Mac Excel, and checks that Excel
opens a file's charts intact.

Usage::

    chart_oracle_gen.py [capture] [--suite NAME ...] [--golden-dir DIR]
    chart_oracle_gen.py open-check FILE.xlsx [FILE.xlsx ...]

``capture`` (the default) reads ``tests/oracle/cases_chart/<suite>.yaml``
and writes ``golden_chart/<suite>.golden.json`` plus ``PROVENANCE.json``.
How a case is captured follows its shape:

* built case (``sheets`` + ``charts``): Excel builds the cells and charts
  through its object model (a chart that lists ``series`` is created over a
  scratch block and pointed at its references with SERIES formulas), runs
  the case's ``ops`` and saves; the golden is the model of Excel's file.
* edit case (``fixture`` + ``edits``): Formulon applies the edits and saves;
  Excel opens that file and re-saves it; the golden is the model of the
  re-saved file.

A suite with ``capture_formulon: true`` also passes Formulon's own file
through the open check, and a failed check fails the capture. The open
check re-saves a file in Excel and compares the two packages: (a) every
chart part is present and referenced from a drawing, (b) every reference
site is unchanged, (c) series count and plot shape are unchanged, (d) Excel
reports as many chart objects as the file holds, and (e) every formatting
value reads back as written. Excel repairs a damaged file without a dialog
under automation, so the check judges content, not dialogs.

Every model is read by ``chart_model_dump`` (built from
tests/oracle/chart_model_dump.cpp): there is no second chart reader here.

macOS + Microsoft 365 Excel only. Excel saves only inside its sandbox
container, so files are staged there.
"""

from __future__ import annotations

import argparse
import datetime as _dt
import hashlib
import json
import os
import platform
import posixpath
import shutil
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import Any, Dict, Iterable, List, Mapping, Optional, Tuple

try:  # pragma: no cover - trivial fallback
    from tools.oracle import chart_case_schema as schema
    from tools.oracle.oracle_gen import _load_divergence_reprobes, _load_divergence_skips
except ImportError:  # pragma: no cover
    sys.path.insert(0, str(Path(__file__).resolve().parent))
    import chart_case_schema as schema  # type: ignore
    from oracle_gen import _load_divergence_reprobes, _load_divergence_skips  # type: ignore

REPO_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_TARGETS_FILE = REPO_ROOT / "tools" / "oracle" / "targets.yaml"
DEFAULT_DUMP = REPO_ROOT / "build" / "tests" / "oracle" / "chart_model_dump"
CONTAINER_DIR = Path.home() / "Library/Containers/com.microsoft.Excel/Data/formulon_chart_oracle"
_M365_SENTINEL = "ARRAYTOTEXT(1) == text '1'"
_TRACK = "chart"

_REL_NS = "{http://schemas.openxmlformats.org/package/2006/relationships}"
_CHART_RELS = ("/chart", "/chartEx")

# Excel chart types for (kind, grouping); line / area "clustered" is Standard.
_CHART_TYPES = {
    ("column", "clustered"): "column_clustered",
    ("column", "stacked"): "column_stacked",
    ("column", "percentStacked"): "column_stacked_100",
    ("bar", "clustered"): "bar_clustered",
    ("bar", "stacked"): "bar_stacked",
    ("bar", "percentStacked"): "bar_stacked_100",
    ("line", "clustered"): "line_chart",
    ("line", "standard"): "line_chart",
    ("line", "stacked"): "line_stacked",
    ("line", "percentStacked"): "line_stacked_100",
    ("area", "clustered"): "area_chart",
    ("area", "standard"): "area_chart",
    ("area", "stacked"): "area_stacked",
    ("area", "percentStacked"): "area_stacked_100",
    ("pie", "clustered"): "pie_chart",
    ("scatter", "clustered"): "xyscatter",
    ("doughnut", "clustered"): "doughnut",
}
_BLANKS = {"gap": "not_plotted", "zero": "zero", "span": "interpolated"}
# Formulon places the n-th chart of a case at this row (0-based), column U.
_ANCHOR_ROW_STEP = 16
_ANCHOR_COL = 21


class CaptureError(RuntimeError):
    """A case Excel could not build, or a file that failed the open check."""


# ---------------------------------------------------------------------------
# Manifest and environment
# ---------------------------------------------------------------------------


def _iso_now() -> str:
    return _dt.datetime.now(_dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def _resolve_target(targets_file: Path) -> Dict[str, Any]:
    import yaml

    doc = yaml.safe_load(targets_file.read_text(encoding="utf-8")) or {}
    track = (doc.get("tracks") or {}).get(_TRACK) or {}
    name = track.get("primary")
    record = (doc.get("targets") or {}).get(name)
    if not isinstance(name, str) or not isinstance(record, dict):
        raise CaptureError("targets.yaml: tracks.chart.primary must name a target")
    if record.get("status") not in {"primary", "scaffolded"}:
        raise CaptureError(f"chart target {name!r} has non-generating status {record.get('status')!r}")
    if record.get("driver") != "macos_excel" or platform.system() not in (record.get("runs_on") or []):
        raise CaptureError(f"chart target {name!r} needs the macos_excel driver on this host")
    output = record.get("output_dir")
    root = (
        REPO_ROOT / Path(output).parent
        if isinstance(output, str) and output
        else REPO_ROOT / "tests/oracle/targets" / name
    )
    return {"_name": name, "_golden_dir": root / "golden_chart", **record}


def _assert_m365() -> Dict[str, str]:
    try:
        from tools.oracle.cf_oracle_gen import _assert_m365_or_abort
    except ImportError:  # pragma: no cover
        from cf_oracle_gen import _assert_m365_or_abort  # type: ignore
    return _assert_m365_or_abort()


# ---------------------------------------------------------------------------
# chart_model_dump
# ---------------------------------------------------------------------------


class Dump:
    """Runs the C++ chart model reader / case builder."""

    def __init__(self, binary: Path) -> None:
        if not binary.is_file():
            raise CaptureError(f"{binary} not found; run `cmake --build build --target chart_model_dump`")
        self.binary = binary

    def model(self, path: Path, views: Iterable[str] = ("all",)) -> Dict[str, Any]:
        out = subprocess.run(
            [str(self.binary), "--view", ",".join(views), str(path)], capture_output=True, text=True, check=False
        )
        if out.returncode != 0:
            raise CaptureError(f"chart_model_dump {path.name}: {out.stderr.strip()}")
        return json.loads(out.stdout)

    def build(self, case_id: str, spec: Mapping[str, Any], out_path: Path) -> Path:
        with tempfile.NamedTemporaryFile("w", suffix=".json", delete=False, encoding="utf-8") as f:
            json.dump({"cases": [{"id": case_id, "spec": spec}]}, f, ensure_ascii=False)
            cases_path = f.name
        try:
            out = subprocess.run(
                [str(self.binary), "--build", cases_path, "--case", case_id, "--out", str(out_path)],
                capture_output=True,
                text=True,
                check=False,
            )
        finally:
            os.unlink(cases_path)
        if out.returncode != 0:
            raise CaptureError(f"Formulon cannot build {case_id}: {out.stderr.strip()}")
        return out_path


# ---------------------------------------------------------------------------
# Excel
# ---------------------------------------------------------------------------


class Excel:
    """One attached Excel session; books are staged in the sandbox container."""

    def __init__(self) -> None:
        from appscript import k  # type: ignore

        self.k = k
        self._attach()
        CONTAINER_DIR.mkdir(parents=True, exist_ok=True)

    def _attach(self) -> None:
        import xlwings as xw  # type: ignore

        self.app = xw.apps.active if xw.apps else xw.App(visible=True, add_book=False)
        self.app.display_alerts = False
        self.keep = {b.name for b in self.app.books}

    def restart(self) -> None:
        """Replaces an Excel that stopped answering Apple events (-1712 / -50)."""

        try:
            self.app.kill()
        except Exception as exc:  # the process may already be gone
            print(f"[chart-oracle] kill Excel: {exc}", file=sys.stderr)
        time.sleep(5.0)
        self._attach()

    def close_new(self) -> None:
        # A book saved under a new name keeps its old handle name, so close by
        # diffing against the books that were open before.
        for name in [b.name for b in self.app.books]:
            if name not in self.keep:
                try:
                    self.app.api.workbooks[name].close(saving=self.k.no)
                except Exception as exc:  # appscript raises its own error types
                    print(f"[chart-oracle] close {name}: {exc}", file=sys.stderr)

    def save(self, book: Any, name: str, dest_dir: Path) -> Path:
        staged = CONTAINER_DIR / name
        if staged.exists():
            staged.unlink()
        book.save(str(staged))
        time.sleep(0.5)
        self.close_new()
        dest = dest_dir / name
        shutil.copyfile(staged, dest)
        return dest

    def open(self, path: Path) -> Any:
        staged = CONTAINER_DIR / ("in_" + path.name)
        shutil.copyfile(path, staged)
        book = self.app.books.open(str(staged))
        # HasTitle and series reads settle a moment after opening.
        time.sleep(2.0)
        return book

    def chart_object_count(self, book: Any) -> int:
        count = 0
        for sheet in book.sheets:
            count += len(sheet.api.chart_objects.get() or [])
        try:
            count += len(book.api.charts.get() or [])
        except Exception:  # a book without chartsheets may not answer
            pass
        return count

    # -- building a case ------------------------------------------------------

    def build(self, spec: Mapping[str, Any], name: str, dest_dir: Path) -> Path:
        book = self.app.books.add()
        try:
            sheets = self._sheets(book, spec["sheets"])
            for n in spec.get("names", []):
                scope = n.get("scope")
                (sheets[scope] if scope else book).names.add(n["name"], n["refers_to"])
            series_widths = [len(c["series"]) + (c["kind"] == "scatter") for c in spec["charts"] if "series" in c]
            scratch = None
            if series_widths:
                scratch = book.sheets.add(name="Scratch", after=book.sheets[-1])
                width = max(series_widths)
                scratch.range("A1").value = [[1.0 + i for i in range(width)], [2.0 + i for i in range(width)]]
            for i, chart in enumerate(spec["charts"]):
                self._chart(book, sheets, scratch, chart, i)
            if scratch is not None:
                scratch.delete()
            for sheet in spec["sheets"]:
                sh = sheets[sheet["name"]]
                for row in sheet.get("hidden_rows", []):
                    sh.range(f"A{row}").api.entire_row.hidden.set(True)
                for col in sheet.get("hidden_cols", []):
                    sh.range(f"{col}1").api.entire_column.hidden.set(True)
                if sheet.get("hidden"):
                    sh.api.visible.set(False)
            for op in spec.get("ops", []):
                self._op(book, op)
            time.sleep(1.0)
            return self.save(book, name, dest_dir)
        except Exception:
            self.close_new()
            raise

    def _sheets(self, book: Any, specs: List[Mapping[str, Any]]) -> Dict[str, Any]:
        out: Dict[str, Any] = {}
        for i, spec in enumerate(specs):
            sh = book.sheets[0] if i == 0 else book.sheets.add(after=book.sheets[-1])
            sh.name = spec["name"]
            for addr, cell in spec.get("cells", {}).items():
                rng = sh.range(addr)
                value = cell
                if isinstance(cell, dict):
                    if "format" in cell:
                        rng.number_format = cell["format"]
                    value = cell.get("value")
                if isinstance(value, str) and value.startswith("="):
                    rng.formula = value
                elif value is not None:
                    rng.value = value
            out[spec["name"]] = sh
        return out

    def _chart(self, book: Any, sheets: Dict[str, Any], scratch: Any, chart: Mapping[str, Any], ordinal: int) -> None:
        k = self.k
        host = sheets[chart.get("host", next(iter(sheets)))]
        anchor = host.range((1 + _ANCHOR_ROW_STEP * ordinal + 1, _ANCHOR_COL))
        co = host.api.make(
            at=host.api,
            new=k.chart_object,
            with_properties={k.left_position: anchor.left, k.top: anchor.top, k.width: 360, k.height: 216},
        )
        co.name.set(chart["id"])
        c = co.chart
        kind = chart["kind"]
        grouping = chart.get("grouping", "clustered")
        c.chart_type.set(getattr(k, _CHART_TYPES[(kind, grouping)]))
        if "series" in chart:
            n = len(chart["series"])
            width = n + (kind == "scatter")
            c.set_source_data(source=scratch.range((1, 1), (2, width)).api, plot_by=k.by_columns)
            made = c.series_collection.get() or []
            if len(made) != n:
                raise CaptureError(f"chart {chart['id']}: scratch source made {len(made)} series, want {n}")
            for order, (s, src) in enumerate(zip(made, chart["series"]), 1):
                cat = src.get("x" if kind == "scatter" else "cat", "")
                val = src["y" if kind == "scatter" else "val"]
                s.formula.set(self._series_formula(book, src.get("name", ""), cat, val, order))
        else:
            source = host.range(chart["source"]).api
            direction = chart.get("direction", "auto")
            if direction == "auto":
                c.set_source_data(source=source)
            else:
                c.set_source_data(source=source, plot_by=k.by_rows if direction == "rows" else k.by_columns)
        if "title_ref" in chart:
            c.has_title.set(True)
            formula = "=" + self._book_names(book, chart["title_ref"])
            try:
                c.chart_title.formula.set(formula)
            except Exception:  # older object models expose the link through the caption only
                c.chart_title.caption.set(formula)
        if "plot_vis_only" in chart:
            c.plot_visible_only.set(chart["plot_vis_only"])
        if "blanks_as" in chart:
            c.display_blanks_as.set(getattr(k, _BLANKS[chart["blanks_as"]]))

    @staticmethod
    def _book_names(book: Any, ref: str) -> str:
        # The file spells a workbook-scope name `[0]!Name`; Excel's formula bar wants the book name.
        return ref.replace("[0]!", f"'{book.name}'!")

    def _series_formula(self, book: Any, name: str, cat: str, val: str, order: int) -> str:
        parts = [self._book_names(book, p) for p in (name, cat, val)]
        return f"=SERIES({parts[0]},{parts[1]},{parts[2]},{order})"

    def _op(self, book: Any, op: Mapping[str, Any]) -> None:
        sh = book.sheets[op["sheet"]]
        kind = op["op"]
        count = op.get("count", 1)
        if kind == "rename_sheet":
            sh.name = op["to"]
        elif kind == "delete_sheet":
            sh.delete()
        elif kind.endswith("_rows"):
            rows = sh.range(f"{op['at']}:{op['at'] + count - 1}")
            if kind == "insert_rows":
                rows.insert(shift="down")
            else:
                rows.delete(shift="up")
        else:
            first = _column_index(op["at"])
            cols = sh.range(f"{op['at']}:{_column_letters(first + count - 1)}")
            if kind == "insert_cols":
                cols.insert(shift="right")
            else:
                cols.delete(shift="left")


def _column_index(letters: str) -> int:
    n = 0
    for ch in letters:
        n = n * 26 + ord(ch) - 64
    return n - 1


def _column_letters(index: int) -> str:
    out = ""
    index += 1
    while index:
        index, r = divmod(index - 1, 26)
        out = chr(65 + r) + out
    return out


# ---------------------------------------------------------------------------
# Open check
# ---------------------------------------------------------------------------


def _rels(z: zipfile.ZipFile, part: str) -> List[Tuple[str, str]]:
    """(type suffix, resolved target) of `part`'s relationships."""

    folder, base = posixpath.split(part)
    rels_path = posixpath.join(folder, "_rels", base + ".rels")
    if rels_path not in z.namelist():
        return []
    out = []
    for rel in ET.fromstring(z.read(rels_path)).iter(f"{_REL_NS}Relationship"):
        target = rel.get("Target", "")
        resolved = target[1:] if target.startswith("/") else posixpath.normpath(posixpath.join(folder, target))
        out.append(("/" + rel.get("Type", "").rsplit("/", 1)[-1], resolved))
    return out


def package_chart_parts(path: Path) -> Tuple[int, List[str]]:
    """Charts referenced from sheet drawings, and the referenced parts that are missing."""

    with zipfile.ZipFile(path) as z:
        names = set(z.namelist())
        referenced = 0
        missing: List[str] = []
        for kind, sheet in _rels(z, "xl/workbook.xml"):
            if kind not in ("/worksheet", "/chartsheet"):
                continue
            for kind2, drawing in _rels(z, sheet):
                if kind2 != "/drawing":
                    continue
                for kind3, chart in _rels(z, drawing):
                    if kind3 in _CHART_RELS:
                        referenced += 1
                        if chart not in names:
                            missing.append(chart)
        return referenced, missing


def _charts(model: Mapping[str, Any]) -> List[Mapping[str, Any]]:
    return list(model.get("charts") or [])


def compare_models(before: Mapping[str, Any], after: Mapping[str, Any], *, check_format: bool) -> List[str]:
    """Checks (b), (c) and optionally (e) between a file and its re-save."""

    problems: List[str] = []
    a, b = _charts(before), _charts(after)
    if len(a) != len(b):
        return [f"chart count {len(a)} -> {len(b)}"]
    for x, y in zip(a, b):
        label = f"{x.get('sheet')}/{x.get('name')}"
        if x.get("refs") != y.get("refs"):
            problems.append(
                f"(b) {label}: references changed: {json.dumps(x.get('refs'))} -> {json.dumps(y.get('refs'))}"
            )
        if x.get("shape") != y.get("shape"):
            problems.append(
                f"(c) {label}: plot shape changed: {json.dumps(x.get('shape'))} -> {json.dumps(y.get('shape'))}"
            )
        if check_format and x.get("format") != y.get("format"):
            problems.append(f"(e) {label}: formatting changed: {_first_difference(x.get('format'), y.get('format'))}")
    return problems


def _first_difference(a: Any, b: Any, path: str = "") -> str:
    if isinstance(a, dict) and isinstance(b, dict):
        for key in sorted(set(a) | set(b)):
            if a.get(key) != b.get(key):
                return _first_difference(a.get(key), b.get(key), f"{path}.{key}")
    if isinstance(a, list) and isinstance(b, list) and len(a) == len(b):
        for i, (x, y) in enumerate(zip(a, b)):
            if x != y:
                return _first_difference(x, y, f"{path}[{i}]")
    return f"{path or '.'}: {json.dumps(a, ensure_ascii=False)} -> {json.dumps(b, ensure_ascii=False)}"


def open_check(
    excel: Excel, dump: Dump, path: Path, work: Path, *, check_format: bool = True
) -> Tuple[Path, List[str]]:
    """Re-saves `path` in Excel and returns the re-saved file and every failed check."""

    before = dump.model(path)
    book = excel.open(path)
    try:
        opened = excel.chart_object_count(book)
    except Exception:
        excel.close_new()
        raise
    resaved = excel.save(book, "resaved_" + path.name, work)
    after = dump.model(resaved)
    problems: List[str] = []
    referenced, missing = package_chart_parts(resaved)
    if missing:
        problems.append(f"(a) chart parts referenced but missing: {missing}")
    if referenced != len(_charts(before)):
        problems.append(f"(a) {referenced} chart parts referenced after re-save, file holds {len(_charts(before))}")
    if opened != len(_charts(before)):
        problems.append(f"(d) Excel opened {opened} chart objects, file holds {len(_charts(before))}")
    problems.extend(compare_models(before, after, check_format=check_format))
    return resaved, problems


# ---------------------------------------------------------------------------
# Capture
# ---------------------------------------------------------------------------


def _capture_case(
    excel: Excel, dump: Dump, cid: str, spec: Mapping[str, Any], views: List[str], capture_formulon: bool, work: Path
) -> Tuple[Dict[str, Any], List[str]]:
    """Excel's model of one case, and every failed open check on the way."""

    if "fixture" in spec:
        ours = dump.build(cid, spec, work / f"{cid}.xlsx")
        saved, problems = open_check(excel, dump, ours, work, check_format=True)
    else:
        saved = excel.build(spec, f"{cid}.xlsx", work)
        problems = []
        if capture_formulon:
            ours = dump.build(cid, spec, work / f"formulon_{cid}.xlsx")
            _, problems = open_check(excel, dump, ours, work, check_format=False)
    return dump.model(saved, views), problems


def _capture_retrying(
    excel: Excel, dump: Dump, cid: str, spec: Mapping[str, Any], views: List[str], capture_formulon: bool, work: Path
) -> Tuple[Dict[str, Any], List[str]]:
    """`_capture_case`, run once more on a fresh Excel when the session stops answering."""

    try:
        return _capture_case(excel, dump, cid, spec, views, capture_formulon, work)
    except CaptureError:
        raise
    except Exception as exc:  # appscript reports a hung session as its own error types
        print(f"[chart-oracle] {cid}: Excel stopped answering ({exc.__class__.__name__}); restarting", file=sys.stderr)
        excel.restart()
    try:
        return _capture_case(excel, dump, cid, spec, views, capture_formulon, work)
    except CaptureError:
        raise
    except Exception as exc:
        raise CaptureError(f"Excel failed twice: {exc}") from exc


def capture_suite(
    excel: Excel,
    dump: Dump,
    stem: str,
    suite: Mapping[str, Any],
    work: Path,
    skips: Mapping[str, str],
    reprobes: Mapping[str, str],
) -> Tuple[List[Dict[str, Any]], List[str]]:
    """Captures one suite. A skip-oracle case records its reason in place of
    `expect`; one whose stamp is still pending is captured anyway and parked
    under `observed` for review."""

    views = suite["views"]
    capture_formulon = bool(suite.get("capture_formulon"))
    cases: List[Dict[str, Any]] = []
    failures: List[str] = []
    for case in suite["cases"]:
        cid = case["id"]
        spec = schema.case_spec(case)
        entry: Dict[str, Any] = {"id": cid}
        if "description" in case:
            entry["description"] = case["description"]
        entry["spec"] = json.loads(json.dumps(spec))
        print(f"[chart-oracle] {stem}/{cid}")
        if cid in skips:
            entry["skipped"] = skips[cid]
            if cid in reprobes:
                try:
                    observed, problems = _capture_retrying(excel, dump, cid, spec, views, capture_formulon, work)
                    entry["observed"] = observed
                    if problems:
                        entry["observed_error"] = "; ".join(problems)
                except CaptureError as exc:
                    entry["observed_error"] = str(exc)
            cases.append(entry)
            continue
        try:
            expect, problems = _capture_retrying(excel, dump, cid, spec, views, capture_formulon, work)
        except CaptureError as exc:
            failures.append(f"{cid}: {exc}")
            continue
        if problems:
            failures.extend(f"{cid}: {p}" for p in problems)
            continue
        entry["expect"] = expect
        cases.append(entry)
    return cases, failures


def write_provenance(
    golden_dir: Path,
    target: Mapping[str, Any],
    environment: Mapping[str, str],
    capture_id: str,
    generated_at: str,
    required: List[str],
    captured: List[str],
    counts: Mapping[str, int],
) -> None:
    """Writes the capture marker; a partial run is recorded as inactive."""

    inventory = [
        {
            "suite": suite,
            "case_count": counts[suite],
            "sha256": hashlib.sha256((golden_dir / f"{suite}.golden.json").read_bytes()).hexdigest(),
            "capture_id": capture_id,
        }
        for suite in sorted(captured)
    ]
    complete = sorted(required) == sorted(captured)
    marker = {
        "track": _TRACK,
        "target": target["_name"],
        "status": target["status"],
        "classification": "active" if complete else "scaffolded",
        "verified": complete,
        "active_ctest": complete,
        "locale": environment.get("excel_locale") or target.get("locale"),
        "product": environment["excel_version"],
        "m365_sentinel": _M365_SENTINEL,
        "capture_id": capture_id,
        "generated_at": generated_at,
        "required_suites": sorted(required),
        "captured_suites": sorted(captured),
        "case_count": sum(item["case_count"] for item in inventory),
        "all_suites_same_capture": True,
        "suite_inventory": inventory,
        "reason": "Chart goldens captured by the manifest-selected Mac Excel 365 target.",
        "source": "tools/oracle/chart_oracle_gen.py",
    }
    (golden_dir / "PROVENANCE.json").write_text(
        json.dumps(marker, indent=2, ensure_ascii=False) + "\n", encoding="utf-8"
    )


def _cmd_capture(args: argparse.Namespace) -> int:
    target = _resolve_target(args.targets_file)
    dump = Dump(args.dump)
    schema.validate_all(args.cases_dir, None)
    suites = schema.load_suites(args.cases_dir)
    wanted = args.suite or sorted(suites)
    unknown = sorted(set(wanted) - set(suites))
    if unknown:
        raise CaptureError(f"unknown suite(s): {unknown}")
    skips = _load_divergence_skips(args.divergence, target["_name"])
    reprobes = _load_divergence_reprobes(args.divergence, target["_name"])
    environment = _assert_m365()
    if environment.get("excel_locale") and environment["excel_locale"] != target.get("locale"):
        raise CaptureError(f"Excel locale {environment['excel_locale']!r} != target locale {target.get('locale')!r}")
    golden_dir: Path = args.golden_dir or target["_golden_dir"]
    golden_dir.mkdir(parents=True, exist_ok=True)
    generated_at = _iso_now()
    capture_id = f"{target['_name']}:{environment['excel_version']}:{target.get('locale')}:{generated_at}"
    excel = Excel()
    counts: Dict[str, int] = {}
    failed = False
    with tempfile.TemporaryDirectory(prefix="formulon-chart-oracle-") as tmp:
        for stem in wanted:
            suite = suites[stem]
            cases, failures = capture_suite(excel, dump, stem, suite, Path(tmp), skips, reprobes)
            if failures:
                failed = True
                print(f"[chart-oracle] {stem}: {len(failures)} failure(s); golden not written", file=sys.stderr)
                for line in failures:
                    print(f"  {line}", file=sys.stderr)
                continue
            golden = {
                "name": stem,
                "track": _TRACK,
                "description": suite.get("description", ""),
                "views": suite["views"],
                "environment": {**environment, "generated_at": generated_at},
                "cases": cases,
            }
            path = golden_dir / f"{stem}.golden.json"
            path.write_text(json.dumps(golden, indent=1, ensure_ascii=False) + "\n", encoding="utf-8")
            counts[stem] = len(cases)
            print(f"[chart-oracle] wrote {path} ({len(cases)} cases)")
    if counts:
        write_provenance(
            golden_dir, target, environment, capture_id, generated_at, sorted(suites), sorted(counts), counts
        )
    return 1 if failed else 0


def _cmd_open_check(args: argparse.Namespace) -> int:
    dump = Dump(args.dump)
    excel = Excel()
    failed = False
    with tempfile.TemporaryDirectory(prefix="formulon-chart-open-check-") as tmp:
        for path in args.files:
            try:
                _, problems = open_check(excel, dump, Path(path).resolve(), Path(tmp))
            except CaptureError as exc:
                problems = [str(exc)]
            if problems:
                failed = True
                print(f"FAIL {path}")
                for problem in problems:
                    print(f"  {problem}")
            else:
                print(f"PASS {path}")
    return 1 if failed else 0


def main(argv: Optional[List[str]] = None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)
    if not argv or argv[0] not in ("capture", "open-check"):
        argv.insert(0, "capture")
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = parser.add_subparsers(dest="command", required=True)
    dump_default = Path(os.environ.get("FORMULON_CHART_MODEL_DUMP", DEFAULT_DUMP))
    p_cap = sub.add_parser("capture", help="Capture goldens (default).")
    p_cap.add_argument("--suite", action="append", default=None, metavar="NAME")
    p_cap.add_argument("--cases-dir", type=Path, default=schema.CASES_DIR)
    p_cap.add_argument("--golden-dir", type=Path, default=None)
    p_cap.add_argument("--targets-file", type=Path, default=DEFAULT_TARGETS_FILE)
    p_cap.add_argument("--divergence", type=Path, default=REPO_ROOT / "tests" / "divergence.yaml")
    p_open = sub.add_parser("open-check", help="Re-save files in Excel and check their charts survive.")
    p_open.add_argument("files", nargs="+")
    for p in (p_cap, p_open):
        p.add_argument("--dump", type=Path, default=dump_default, help="chart_model_dump binary")
    args = parser.parse_args(argv)
    if platform.system() != "Darwin":
        print("chart-oracle: macOS + Excel only", file=sys.stderr)
        return 2
    try:
        return _cmd_open_check(args) if args.command == "open-check" else _cmd_capture(args)
    except (CaptureError, schema.ValidationError) as exc:
        print(f"chart-oracle: {exc}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    sys.exit(main())
