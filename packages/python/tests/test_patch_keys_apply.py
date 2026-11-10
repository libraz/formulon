"""Every key of every patch record applies: each key is set through its public ``update_*`` method to a value that
differs from the setup state and read back through the public getter. A meta-test ties the table to the API ledger
(``tools/api/names.yaml``) and the C structs, so a new patch key cannot ship without a case."""

from __future__ import annotations

import re
import unittest
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional

import formulon
from formulon import (
    BlanksAs,
    BorderStyle,
    ChartAxisCrosses,
    ChartAxisKind,
    ChartAxisPosition,
    ChartDataSource,
    ChartInput,
    ChartLabelPosition,
    ChartLineDash,
    ChartMarkerSymbol,
    ChartPlotType,
    ChartSourceKind,
    ChartTickLabelPosition,
    Color,
    ColorKind,
    FillPattern,
    FontFamily,
    FontScheme,
    FontVerticalAlign,
    GradientFill,
    GradientStop,
    HorizontalAlign,
    ReadingOrder,
    UnderlineStyle,
    VerticalAlign,
    Workbook,
)

REPO = Path(__file__).resolve().parents[3]
SURFACE = "py"

#: Set the opposite of what the getter reads before the patch.
FLIP = object()
_UNSET = object()


@dataclass
class Case:
    """One key: set ``v`` (or flip the pre-read bool), expect ``want``; ``clear`` also sends ``None``."""

    v: Any
    get: Optional[Callable[[Any], Any]] = None
    want: Any = _UNSET
    patch: Optional[Callable[[Any], Dict[str, Any]]] = None
    clear: bool = False
    cleared: Any = None
    clear_patch: Optional[Dict[str, Any]] = None
    make: Optional[Callable[[Any, Any], Any]] = None


@dataclass
class PatchType:
    """A patch record: how to build a target, apply a patch to it, read the patched record back, and its cases."""

    make: Callable[[Any, Any], Any]
    update: Callable[..., None]
    read: Callable[[Any], Any]
    cases: Dict[str, Case] = field(default_factory=dict)


def _argb(color: Any) -> Optional[int]:
    return None if color is None else color.argb & 0xFFFFFFFF


def _rgb(argb: int) -> Color:
    return Color(kind=ColorKind.RGB, argb=argb)


def _grid(ws: Any) -> None:
    """Sheet1: B1:D1 series names, A2:A5 categories, B2:D5 numbers."""
    for c, name in enumerate(("S1", "S2", "S3")):
        ws.set_text(0, c + 1, name)
    for r, category in enumerate(("a", "b", "c", "d")):
        ws.set_text(r + 1, 0, category)
        for c in range(3):
            ws.set_number(r + 1, c + 1, float((r + 1) * (c + 2)))


def _chart(ws: Any, kind: ChartPlotType = ChartPlotType.BAR) -> Any:
    _grid(ws)
    return ws.add_chart(ChartInput(type=kind, range="Sheet1!A1:D5"))


def _first_series(chart: Any) -> Any:
    return chart.get_properties().plots[0].series[0]


def _literal(*points: Any) -> ChartDataSource:
    return ChartDataSource(kind=ChartSourceKind.LITERAL, literal=list(points))


def _points(source: Any) -> List[Any]:
    return list(source.literal)


def _source_case(*points: Any, **extra: Any) -> Case:
    return Case(_literal(*points), want=list(points), **extra)


# ---- Chart sub-records, reached through a series, a point or a marker ---------------------------------------------


def _line_chart(ws: Any, wb: Any) -> Any:
    return _chart(ws, ChartPlotType.LINE)


def _paint(slot: Callable[[Any], Any], name: str) -> Dict[str, Any]:
    return {
        "make": _line_chart,
        "update": lambda chart, **patch: chart.update_series(0, 0, **{name: patch}),
        "read": lambda chart: slot(_first_series(chart)),
    }


def _bar_paint(slot: Callable[[Any], Any], name: str) -> Dict[str, Any]:
    return {**_paint(slot, name), "make": lambda ws, wb: _chart(ws)}


CHART_FILL = PatchType(
    **_bar_paint(lambda s: s.fill, "fill"),
    cases={
        "color": Case("#FF0000", get=lambda r: _argb(r and r.color), want=0xFFFF0000, clear=True),
        "transparency_pct": Case(
            40,
            get=lambda r: r and r.transparency_pct,
            patch=lambda v: {"color": "#FF0000", "transparency_pct": v},
            clear=True,
        ),
        "none": Case(FLIP, get=lambda r: bool(r and r.none)),
    },
)

CHART_LINE = PatchType(
    **_paint(lambda s: s.line, "line"),
    cases={
        "color": Case("#00FF00", get=lambda r: _argb(r and r.color), want=0xFF00FF00, clear=True),
        "transparency_pct": Case(
            25,
            get=lambda r: r and r.transparency_pct,
            patch=lambda v: {"color": "#00FF00", "transparency_pct": v},
            clear=True,
        ),
        "width_pt": Case(3.5, get=lambda r: r and r.width_pt, clear=True),
        "dash": Case(ChartLineDash.DOT, get=lambda r: r and r.dash, clear=True),
        "none": Case(FLIP, get=lambda r: bool(r and r.none)),
    },
)

CHART_MARKER = PatchType(
    **_paint(lambda s: s.marker, "marker"),
    cases={
        "symbol": Case(ChartMarkerSymbol.SQUARE, get=lambda r: r and r.symbol),
        "size_pt": Case(
            11,
            get=lambda r: r and r.size_pt,
            patch=lambda v: {"symbol": ChartMarkerSymbol.CIRCLE, "size_pt": v},
            clear=True,
            cleared=5,
        ),
    },
)


def _labels_chart(ws: Any, wb: Any) -> Any:
    chart = _chart(ws, ChartPlotType.LINE)
    chart.update_series(0, 0, labels={"show_value": True})
    return chart


CHART_LABELS = PatchType(
    make=_labels_chart,
    update=lambda chart, **patch: chart.update_series(0, 0, labels=patch),
    read=lambda chart: _first_series(chart).labels,
    cases={
        "show_value": Case(FLIP),
        "show_category": Case(FLIP),
        "show_series": Case(FLIP),
        "show_percent": Case(FLIP),
        "show_legend_key": Case(FLIP),
        "separator": Case("; ", clear=True),
        "position": Case(ChartLabelPosition.ABOVE, clear=True),
        "number_format": Case("0.00", clear=True),
        "font_color": Case("#0000FF", get=lambda r: _argb(r.font_color), want=0xFF0000FF, clear=True),
        "font_size_pt": Case(14, clear=True),
        "source_linked": Case(FLIP),
        "deleted": Case(FLIP),
    },
)


def _scatter_chart(ws: Any, wb: Any) -> Any:
    return _chart(ws, ChartPlotType.SCATTER)


CHART_SERIES = PatchType(
    make=_line_chart,
    update=lambda chart, **patch: chart.update_series(0, 0, **patch),
    read=_first_series,
    cases={
        "name": _source_case("Fixed", get=lambda r: _points(r.name), clear=True, cleared=[]),
        "categories": _source_case("p", "q", "r", "s", get=lambda r: _points(r.categories), clear=True, cleared=[]),
        "values": _source_case(9.0, 8.0, 7.0, 6.0, get=lambda r: _points(r.values)),
        "x": _source_case(5.0, 6.0, 7.0, 8.0, get=lambda r: _points(r.x), make=_scatter_chart, clear=True, cleared=[]),
        "y": _source_case(4.0, 3.0, 2.0, 1.0, get=lambda r: _points(r.y), make=_scatter_chart),
        "marker": Case(
            {"symbol": ChartMarkerSymbol.DIAMOND},
            get=lambda r: r.marker and r.marker.symbol,
            want=ChartMarkerSymbol.DIAMOND,
            clear=True,
        ),
        "smooth": Case(FLIP, get=lambda r: bool(r.smooth)),
        "labels": Case(
            {"show_value": True},
            get=lambda r: bool(r.labels and r.labels.show_value),
            want=True,
            clear=True,
            cleared=False,
        ),
    },
)

CHART_POINT = PatchType(
    make=_line_chart,
    update=lambda chart, **patch: chart.update_point(0, 0, 1, **patch),
    read=lambda chart: next((p for p in _first_series(chart).points if p.index == 1), None),
    cases={
        "marker": Case(
            {"symbol": ChartMarkerSymbol.STAR},
            get=lambda r: r and r.marker and r.marker.symbol,
            want=ChartMarkerSymbol.STAR,
            clear=True,
        ),
    },
)

CHART_PLOT = PatchType(
    make=lambda ws, wb: _chart(ws),
    update=lambda chart, **patch: chart.update_plot(0, **patch),
    read=lambda chart: chart.get_properties().plots[0],
    cases={
        "gap_width_pct": Case(80),
        "overlap_pct": Case(25),
        "vary_colors": Case(FLIP),
        "first_slice_angle_deg": Case(90, make=lambda ws, wb: _chart(ws, ChartPlotType.PIE)),
        "hole_size_pct": Case(70, make=lambda ws, wb: _chart(ws, ChartPlotType.DOUGHNUT)),
    },
)


def _value_axis(chart: Any) -> Any:
    return next(a for a in chart.get_properties().axes if a.kind == ChartAxisKind.VALUE)


CHART_AXIS = PatchType(
    make=lambda ws, wb: _chart(ws),
    update=lambda chart, **patch: chart.update_axis(_value_axis(chart).id, **patch),
    read=_value_axis,
    cases={
        "visible": Case(FLIP),
        "title": Case("Amount", get=lambda r: r.title and r.title.text, clear=True),
        "min": Case(1.0, clear=True),
        "max": Case(100.0, clear=True),
        "major_unit": Case(5.0, clear=True),
        "minor_unit": Case(0.5, clear=True),
        "number_format": Case("0.00", clear=True),
        "source_linked": Case(FLIP),
        "major_gridlines": Case(FLIP),
        "minor_gridlines": Case(FLIP),
        "position": Case(ChartAxisPosition.RIGHT),
        "reversed": Case(FLIP),
        "log_base": Case(10.0, clear=True),
        "crosses": Case(ChartAxisCrosses.MAX),
        "crosses_at": Case(3.0),
        "tick_label_position": Case(ChartTickLabelPosition.LOW, clear=True, cleared=ChartTickLabelPosition.NEXT_TO),
    },
)

CHART_DISPLAY = PatchType(
    make=lambda ws, wb: _chart(ws),
    update=lambda chart, **patch: chart.update_display(**patch),
    read=lambda chart: chart.get_properties(),
    cases={
        "blanks_as": Case(BlanksAs.SPAN, clear=True),
        "plot_visible_only": Case(
            FLIP, get=lambda r: True if r.plot_visible_only is None else r.plot_visible_only, clear=True, cleared=True
        ),
        "na_as_blank": Case(FLIP, get=lambda r: bool(r.na_as_blank), clear=True, cleared=False),
        "style": Case(26, clear=True),
    },
)


# ---- Cell format --------------------------------------------------------------------------------------------------


def _format_type(slot: str, cases: Dict[str, Case]) -> PatchType:
    """A cell-format sub-patch type: ``slot`` names the CellFormat group the keys land in."""
    return PatchType(
        make=lambda ws, wb: ws,
        update=lambda ws, **patch: ws.update_cell_format("B2", **{slot: patch}),
        read=lambda ws: getattr(ws.get_cell_format("B2"), slot),
        cases=cases,
    )


FONT = _format_type(
    "font",
    {
        "name": Case("Courier New"),
        "size_pt": Case(17.0, clear=True, cleared=11.0),
        "bold": Case(FLIP, clear=True),
        "italic": Case(FLIP, clear=True),
        "strike": Case(FLIP, clear=True),
        "outline": Case(FLIP, clear=True),
        "shadow": Case(FLIP, clear=True),
        "condense": Case(FLIP, clear=True),
        "extend": Case(FLIP, clear=True),
        "underline": Case(UnderlineStyle.DOUBLE),
        "vert_align": Case(FontVerticalAlign.SUPERSCRIPT),
        "family": Case(FontFamily.MODERN, clear=True),
        "charset": Case(128, clear=True),
        "scheme": Case(FontScheme.MINOR),
        "color": Case("#123456", get=lambda r: _argb(r.color), want=0xFF123456),
    },
)

FILL = _format_type(
    "fill",
    {
        "pattern": Case(FillPattern.LIGHT_GRID),
        "foreground": Case(
            "#112233",
            get=lambda r: _argb(r.foreground),
            want=0xFF112233,
            patch=lambda v: {"pattern": FillPattern.SOLID, "foreground": v},
        ),
        "background": Case(
            "#445566",
            get=lambda r: _argb(r.background),
            want=0xFF445566,
            patch=lambda v: {"pattern": FillPattern.SOLID, "background": v},
        ),
        "gradient": Case(
            GradientFill(
                angle_deg=45.0,
                stops=[GradientStop(0.0, _rgb(0xFFFF0000)), GradientStop(1.0, _rgb(0xFF0000FF))],
            ),
            get=lambda r: r.gradient and r.gradient.angle_deg,
            want=45.0,
            clear=True,
        ),
    },
)

BORDER_SIDE = PatchType(
    make=lambda ws, wb: ws,
    update=lambda ws, **patch: ws.update_cell_format("B2", border={"left": patch}),
    read=lambda ws: ws.get_cell_format("B2").border.left,
    cases={
        "style": Case(BorderStyle.THICK),
        "color": Case(
            "#AA0000",
            get=lambda r: _argb(r.color),
            want=0xFFAA0000,
            patch=lambda v: {"style": BorderStyle.THIN, "color": v},
        ),
    },
)


def _side_case(side: str) -> Case:
    return Case({"style": BorderStyle.MEDIUM}, get=lambda r: getattr(r, side).style, want=BorderStyle.MEDIUM)


BORDER = _format_type(
    "border",
    {
        "left": _side_case("left"),
        "right": _side_case("right"),
        "top": _side_case("top"),
        "bottom": _side_case("bottom"),
        "diagonal": _side_case("diagonal"),
        "diagonal_up": Case(FLIP),
        "diagonal_down": Case(FLIP),
    },
)

ALIGNMENT = _format_type(
    "alignment",
    {
        "horizontal": Case(HorizontalAlign.CENTER),
        "vertical": Case(VerticalAlign.TOP),
        "wrap_text": Case(FLIP),
        "shrink_to_fit": Case(FLIP),
        "indent": Case(3, patch=lambda v: {"horizontal": HorizontalAlign.LEFT, "indent": v}),
        "text_rotation": Case(45),
        "reading_order": Case(ReadingOrder.RIGHT_TO_LEFT),
        "justify_last_line": Case(
            FLIP, patch=lambda v: {"horizontal": HorizontalAlign.JUSTIFY, "justify_last_line": v}
        ),
        "relative_indent": Case(2),
    },
)

PROTECTION = _format_type("protection", {"locked": Case(FLIP), "hidden": Case(FLIP)})


def _column_table(sheet: Any, wb: Any) -> Any:
    """Table Tbl over A1:B4 with a totals row: headers Item / Qty, data rows 2..3."""
    sheet.set_text("A1", "Item")
    sheet.set_text("B1", "Qty")
    for row, qty in ((2, 1.0), (3, 2.0)):
        sheet.set_text(f"A{row}", f"r{row}")
        sheet.set_number(f"B{row}", qty)
    return sheet.add_table(formulon.TableInput(range="A1:B4", name="Tbl", totals_row=True))


TABLE_COLUMN = PatchType(
    make=_column_table,
    update=lambda table, **kw: table.update(columns=[formulon.TableColumnPatch(), formulon.TableColumnPatch(**kw)]),
    read=lambda table: table.get_properties().columns[1],
    cases={
        # The setup's totals row is blank, so its last column already defaults to SUM; patch a different function.
        "totals_row_function": Case(
            formulon.TableTotalsRowFunction.MAX, clear=True, cleared=formulon.TableTotalsRowFunction.NONE
        ),
        "totals_row_label": Case("Total", clear=True),
        "totals_row_formula": Case(
            "=5",
            patch=lambda v: {"totals_row_function": formulon.TableTotalsRowFunction.CUSTOM, "totals_row_formula": v},
            clear=True,
            clear_patch={"totals_row_function": formulon.TableTotalsRowFunction.MAX, "totals_row_formula": None},
        ),
        "calculated_formula": Case("=[@Qty]*10", clear=True),
    },
)


TYPES: Dict[str, PatchType] = {
    "ChartFillPatch": CHART_FILL,
    "ChartLinePatch": CHART_LINE,
    "ChartMarkerPatch": CHART_MARKER,
    "ChartLabelsPatch": CHART_LABELS,
    "ChartSeriesPatch": CHART_SERIES,
    "ChartPointPatch": CHART_POINT,
    "ChartPlotPatch": CHART_PLOT,
    "ChartAxisPatch": CHART_AXIS,
    "ChartDisplayPatch": CHART_DISPLAY,
    "FontPatch": FONT,
    "FillPatch": FILL,
    "BorderSidePatch": BORDER_SIDE,
    "BorderPatch": BORDER,
    "AlignmentPatch": ALIGNMENT,
    "CellProtectionPatch": PROTECTION,
    "TableColumnPatch": TABLE_COLUMN,
}

#: Keys with no public readback, by type; each value is the reason.
EXEMPT: Dict[str, Dict[str, str]] = {
    "ChartSeriesPatch": {
        "size": "a bubble plot cannot be created through the public API, so no series accepts a size",
    },
}


def _apply(test: unittest.TestCase, type_name: str, key: str) -> None:
    spec = TYPES[type_name]
    case = spec.cases[key]
    with Workbook.create_default() as wb:
        target = (case.make or spec.make)(wb.sheet(0), wb)
        get = case.get or (lambda record: getattr(record, key))
        before = get(spec.read(target))
        value = (not before) if case.v is FLIP else case.v
        want = value if case.want is _UNSET else case.want
        test.assertNotEqual(before, want, f"{key}: the value must differ from the setup state")
        patch = case.patch(value) if case.patch else {key: value}
        spec.update(target, **patch)
        test.assertEqual(get(spec.read(target)), want, f"{type_name}.{key} was not applied")
        if case.clear:
            spec.update(target, **(case.clear_patch if case.clear_patch is not None else {key: None}))
            test.assertEqual(get(spec.read(target)), case.cleared, f"{type_name}.{key} was not cleared")


class PatchKeysApplyTests(unittest.TestCase):
    """One generated test per patch key."""


def _attach() -> None:
    for type_name, spec in TYPES.items():
        for key in spec.cases:
            name = "test_" + re.sub(r"(?<!^)(?=[A-Z])", "_", type_name).lower() + "__" + key
            setattr(PatchKeysApplyTests, name, lambda self, t=type_name, k=key: _apply(self, t, k))


_attach()


# ---- Coverage of the ledger ---------------------------------------------------------------------------------------


def _ledger_patches() -> List[Dict[str, Optional[str]]]:
    """The ``role: patch`` records of the ledger that exist on the Python surface."""
    text = (REPO / "tools/api/names.yaml").read_text(encoding="utf-8")
    out: List[Dict[str, Optional[str]]] = []
    for m in re.finditer(r"^  (record\.\w+):\n((?:    .*\n)+)", text, re.M):
        body = m.group(2)
        if "role: patch" not in body:
            continue
        surfaces = re.search(r"surfaces: \[(.*?)\]", body)
        if not surfaces or SURFACE not in [s.strip() for s in surfaces.group(1).split(",")]:
            continue
        c = re.search(r"\bc: (\w+)", body)
        py = re.search(r"\bpy: (\w+)", body)
        out.append({"id": m.group(1), "c": c.group(1) if c else None, "py": py.group(1) if py else None})
    return out


def _struct_keys(header: str, c_name: str) -> List[str]:
    """The fields of a C patch struct that carry a sibling ``<field>_op``."""
    end = header.index(f"}} {c_name};")
    body = header[header.rindex("typedef struct", 0, end) : end]
    return re.findall(r"\b(\w+)_op;", body)


class PatchKeyCoverageTests(unittest.TestCase):
    def test_ledger_lists_patch_records(self) -> None:
        self.assertGreaterEqual(len(_ledger_patches()), 15)

    def test_every_key_of_every_patch_record_has_a_case(self) -> None:
        header = (REPO / "src/c_api/formulon_c.h").read_text(encoding="utf-8")
        for rec in _ledger_patches():
            name = rec["py"]
            with self.subTest(record=rec["id"]):
                self.assertTrue(name, f"{rec['id']} has no Python type name")
                self.assertIn(name, TYPES, f"{name} has no cases")
                if rec["c"]:
                    keys = set(_struct_keys(header, rec["c"]))
                else:
                    keys = set(getattr(formulon, name).__annotations__)
                self.assertTrue(keys, f"{name} has no keys")
                covered = set(TYPES[name].cases) | set(EXEMPT.get(name, {}))
                self.assertEqual(sorted(keys - covered), [], f"{name} keys without a case")
                self.assertEqual(sorted(set(TYPES[name].cases) - keys), [], f"{name} cases that are not keys")


if __name__ == "__main__":
    unittest.main()
