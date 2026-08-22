# formulon

[![PyPI](https://img.shields.io/pypi/v/formulon)](https://pypi.org/project/formulon/)
[![npm](https://img.shields.io/npm/v/@libraz/formulon)](https://www.npmjs.com/package/@libraz/formulon)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![Python](https://img.shields.io/pypi/pyversions/formulon)](https://pypi.org/project/formulon/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net)

Excel 365 calculation engine, exposed as a pure-Python binding driven by
WebAssembly. Evaluates formulas; loads and saves `.xlsx` workbooks; and
edits cells, the row/column matrix, styles, merges, comments,
hyperlinks, data validations, conditional formats, defined names, and
PivotTables. It also exposes recalc (full and partial), dependency-graph
tracing (precedents / dependents), dynamic-array spill info, the function
catalog, and per-sheet view / protection settings. The goal is 1-bit
compatibility with the default `win-365-ja_JP` behavior profile; callers can
select the separately supported `mac-365-ja_JP` profile when required.

## Install

```sh
pip install formulon
```

Requires Python 3.9 or newer. The wheel is `py3-none-any`: it contains a
single `formulon_capi.wasm` (a standalone reactor-style WebAssembly
module exporting the engine's C ABI) plus a thin Python wrapper that
drives it through [`wasmtime`](https://pypi.org/project/wasmtime/). One
wheel works on every platform `wasmtime` supports: Linux x86_64 /
aarch64, macOS x86_64 / arm64, Windows x86_64.

`wasmtime` is the only runtime dependency, declared in the wheel's
metadata; `pip` resolves the right platform-specific `wasmtime` build at
install time.

## Quick start

```python
import formulon

# One-shot formula evaluation against a fresh workbook.
v = formulon.eval_formula("=SUM(1,2,3)")
print(v.to_python())  # 6.0

# Cell-level Excel errors surface as Value(kind=ValueKind.ERROR), not
# Python exceptions.
v = formulon.eval_formula("=1/0")
print(v.kind)  # ValueKind.ERROR
print(v.error_code)  # 1 (ErrorCode::Div0)
```

## Workbook example

```python
from formulon import Workbook

with Workbook.create_default() as wb:
    wb.set_number(0, 0, 0, 21.0)  # Sheet1!A1 = 21
    wb.set_formula(0, 0, 1, "=A1*2")  # Sheet1!B1 = =A1*2
    wb.recalc()

    print(wb.get_value(0, 0, 1).to_python())  # 42.0

    # Serialise back to an in-memory .xlsx.
    blob = wb.save()
    with open("output.xlsx", "wb") as f:
        f.write(blob)

# Load back from disk.
with open("output.xlsx", "rb") as f:
    blob = f.read()

with Workbook.load(blob) as wb:
    wb.recalc()
    for cell in wb.iter_cells(0):
        print(cell)
```

## API surface

`Workbook` mirrors the C ABI surface that the npm bindings expose, minus
the exclusions listed under [Not exposed](#not-exposed). Every public
method carries a docstring (`help(Workbook.set_number)`), and the
signatures live in the hand-rolled type stubs in `formulon/__init__.pyi`.
The groups, with one example each:

**Core** -- `create_default()` / `create_empty()` / `load(bytes)` factories,
`set_number/set_bool/set_text/set_blank/set_formula`, `get_value`,
`recalc()`, `save()`, and the `iter_cells/iter_defined_names/iter_tables/
iter_passthrough` iterators.

**Tables** -- create, retarget, and drop `xl/tables/tableN.xml` entries:

```python
wb.set_text(0, 0, 0, "Region")
wb.set_text(0, 0, 1, "Amount")
index = wb.table_create(0, "A1:B3", "Table1", "Table1", ["Region", "Amount"])
wb.table_update(index, "A1:B9")  # retarget; None preserves style / flags
wb.table_remove(index)
```

**Phonetic guides** -- the OOXML `<rPh>` furigana attached to a cell string:

```python
wb.set_text(0, 0, 0, "日本語")
wb.set_phonetic(0, 0, 0, "ニホンゴ")
wb.get_phonetic(0, 0, 0)  # -> 'ニホンゴ'; '' when the cell has none
```

A reading that covers only part of the text needs one run per span. The
spans are observable through `PHONETIC`, which substitutes each annotated
span and passes the rest through, so a partially annotated cell has to be
read and written as runs -- `set_phonetic` annotates the whole cell and
would collapse them:

```python
wb.set_text(0, 0, 0, "東京都")
wb.set_phonetic_runs(0, 0, 0, [PhoneticRun(0, 2, "トウキョウ"), PhoneticRun(2, 3, "ト")])
wb.get_phonetic_runs(0, 0, 0)  # -> [PhoneticRun(0, 2, 'トウキョウ'), PhoneticRun(2, 3, 'ト')]
wb.get_phonetic(0, 0, 0)  # -> 'トウキョウト' (the readings concatenated)
```

Offsets are UTF-16 code units, and the runs must be an ordered partition:
each needs `sb <= eb` and must start at or after the previous run's `eb`.

**AutoFilter** -- the raw `<autoFilter>` fragment, preserved verbatim so
filter criteria and extensions survive a round trip:

```python
wb.set_auto_filter_xml(0, '<autoFilter ref="A1:B3"/>')
wb.get_auto_filter_xml(0)
wb.set_auto_filter_xml(0, "")  # removes it
```

**Sheets & matrix edits**

```python
wb.add_sheet("Data")
wb.rename_sheet(1, "Numbers")
wb.move_sheet(1, 0)
wb.set_number(0, 0, 0, 11.0)  # A1 = 11
wb.insert_rows(0, 0, 1)  # A1 shifts down to A2
wb.delete_cols(0, 5, 1)  # delete column F
```

**Defined names**

```python
wb.set_defined_name("TaxRate", "Sheet1!$A$1")  # set / replace
wb.set_defined_name("TaxRate", "")  # empty formula removes it
```

**Partial recalc** -- recompute only the closure feeding a viewport:

```python
recomputed = wb.partial_recalc(sheet=0, first_row=0, last_row=0, first_col=0, last_col=1)
```

**Merges / comments / hyperlinks / validations**

```python
from formulon import MergeRange, DataValidationInput

wb.add_merge(0, MergeRange(0, 0, 1, 1))
wb.set_comment(0, 0, 0, author="alice", text="see note")
wb.add_hyperlink(0, 0, 0, "https://example.com", "Example", "tooltip")
wb.add_validation(0, DataValidationInput(type=3, ranges=[MergeRange(0, 0, 4, 0)], formula1='"a,b,c"'))
```

**Styles & number formats**

```python
from formulon import FontRecord, FillRecord, CellXf

fi = wb.add_font(FontRecord(name="Calibri", size=12.0, bold=True))
fill = wb.add_fill(FillRecord(pattern=1, fg_argb=0xFFFFFF00))
border = wb.add_border({"left": {"style": 1, "color_argb": 0xFF000000}})
nf = wb.add_num_fmt("0.00")
xf = wb.add_cell_xf(
    CellXf(
        font_index=fi,
        fill_index=fill,
        border_index=border,
        num_fmt_id=nf,
        horizontal_align=0,
        vertical_align=0,
        wrap_text=False,
    )
)
wb.set_cell_xf_index(0, 0, 0, xf)
```

`add_font` always appends beside font 0, the record an unstyled cell
resolves to. To change what a never-styled cell is saved as -- Calibri 11
in a fresh workbook -- declare the default instead:

```python
wb.set_default_font(FontRecord(name="游ゴシック", size=11.0, has_charset=True, charset=128))
wb.get_font(0).name  # -> '游ゴシック'
```

**Conditional formatting**

```python
from formulon import ConditionalFormatInput, MergeRange

wb.add_conditional_format(
    0,
    ConditionalFormatInput(
        sqref=[MergeRange(0, 0, 9, 0)],
        type=1,  # cellIs
        op_engaged=True,
        op=4,
        formula1="100",
    ),
)  # greaterThan 100
matches = wb.evaluate_cf_range(0, 0, 0, 9, 0)  # per-cell resolved matches
```

**PivotTables** -- build a cache, project a layout:

```python
from formulon import PivotFieldSpec, PivotDataFieldSpec, PivotAxis, PivotAggregation

cache = wb.pivot_cache_create()
wb.pivot_cache_field_add(cache, "Region")
wb.pivot_cache_field_add(cache, "Amount")
rec = wb.pivot_cache_record_add(cache)
wb.pivot_cache_record_set_text(cache, rec, 0, "East")
wb.pivot_cache_record_set_number(cache, rec, 1, 10.0)
pivot = wb.pivot_create(0, "Pivot1", cache, anchor_row=0, anchor_col=4)
region = wb.pivot_field_add(0, pivot, PivotFieldSpec("Region", axis=PivotAxis.ROW))
amount = wb.pivot_field_add(0, pivot, PivotFieldSpec("Amount", axis=PivotAxis.VALUE))
wb.pivot_data_field_add(0, pivot, PivotDataFieldSpec("Sum of Amount", amount, aggregation=PivotAggregation.SUM))
layout = wb.pivot_layout(0, pivot)  # -> PivotLayout(top, left, rows, cols, cells)
```

**Dependency trace & spill**

```python
wb.set_number(0, 0, 0, 1.0)
wb.set_formula(0, 0, 1, "=A1")
wb.recalc()
wb.precedents(0, 0, 1)  # -> [CellNode(sheet=0, row=0, col=0)]
wb.dependents(0, 0, 0)  # -> [CellNode(sheet=0, row=0, col=1)]
wb.set_formula(0, 5, 0, "=SEQUENCE(3)")
wb.recalc()
wb.spill_info(0, 5, 0)  # -> SpillInfo(engaged=True, rows=3, cols=1, ...)
```

**Function catalog** (static; needs no workbook handle)

```python
Workbook.function_count()  # number of registered functions
Workbook.function_metadata("SUM", 0)  # FunctionMetadata or None
```

**Sheet view / protection / calc policy** -- `get_sheet_view` /
`set_sheet_zoom` / `set_sheet_freeze` / `set_sheet_tab_hidden`,
`get_sheet_columns` / `set_column_width` / `set_row_height` (and the
hidden / outline variants), `get_sheet_protection` /
`set_sheet_protection`, `calc_mode` / `set_calc_mode`, `excel_profile_id`
/ `set_excel_profile_id`, plus `get_external_links()`.

**Print layout & pagination** -- author page setup, print area/titles and
manual page breaks, then resolve them into printed pages:

```python
wb.set_print_area(0, "A1:F20")
wb.set_print_titles(0, repeat_rows="1:2")
wb.add_row_break(0, 20)
wb.set_print_options(0, grid_lines=True)
result = wb.paginate(0)  # -> PaginationResult(page_count, print_area, ...)
```

`get_print_area` / `get_print_titles`, `add_col_break` /
`remove_row_break` / `remove_col_break` / `clear_breaks`,
`get_row_breaks` / `get_col_breaks` and `set_header_footer` round out the
group. `set_range_xf_index(sheet, first_row, first_col, last_row,
last_col, xf_index)` stores one style index across a rectangle of cells
in a single call, which is what rules a report's border without one call
per cell.

### Values and errors

`Value` exposes `kind`, `number`, `boolean`, `text`, `error_code`, plus
`to_python()` which converts to the natural Python type
(`None` / `float` / `bool` / `str`) or returns the `Value` itself for
errors and reserved kinds.

`FormulonError` is raised only for host-side problems (NULL handle,
parser crash inside `Workbook.load`, out-of-range index, OOM). Excel cell
errors travel inside `Value(kind=ValueKind.ERROR)`. The one absent
lookup is `get_comment`, which returns `None` (not an error) when no
comment is anchored at the requested cell.

### Not exposed

The scalar `fm_workbook_evaluate_formula` (`evaluateFormulaText` in the
npm bindings) is not bound. `evaluate_formula_array` covers the same
evaluate-without-mutating use case; a scalar result comes back as a 1x1
array, so index `[0][0]` in place of the reduced top-left element the
scalar call would have returned.

`recalc()` is always serial in this no-pthread WASM wheel: the parallel
scheduler requires a pthread runtime that wasmtime does not provide.
Thread-capable native surfaces can opt in separately (the CLI uses
`formulon recalc --threads N`), while Python intentionally exposes no
parallel-recalc method. Result fidelity is identical.

The iterative-solver progress callback (`fm_workbook_set_iterative_progress`
in the C ABI) is intentionally **not** bound -- it takes a C function
pointer that the host cannot synthesise into the WebAssembly module's
function table through `wasmtime`. Configure iterative calculation via
`set_iterative(enabled, max_iterations, max_change)` instead; only the
per-sweep callback is unavailable.

The structured-log **sink** (`fm_set_log_sink`) is unavailable for the same
reason: it too takes a C function pointer. The threshold half of that
surface is bound as the module-level `set_log_min_level(level)`, which is
process-wide state rather than a `Workbook` method. The default is
`LogLevel.OFF`, so a caller that never touches it sees no engine output;
raising it to `LogLevel.WARN` or below sends one JSON record per line to
stderr.

Worksheet XML is read through the DOM parser only. The native CLI
switches to a streaming parser for worksheets past 256 KiB; that
implementation costs binary size the WASM budget does not have, so
loading a workbook here needs memory proportional to the largest single
worksheet's XML rather than a fixed window. Sheets are read one at a
time, so the peak is per worksheet, and the practical ceiling is the
32-bit WASM address space. Results are identical either way.

## Building from source

```sh
# From the repository root:
make wasm-capi         # builds build-wasm-capi/formulon_capi.wasm (Emscripten)
make python-package    # stages the wasm into packages/python/formulon/_wasm/
make python-test       # runs the smoke tests against the staged package
make python-wheel      # produces a py3-none-any build-py/dist/formulon-*.whl
```

`packages/python/scripts/stage.py` is the entry point; it just copies
the pre-built `formulon_capi.wasm` into the package data directory --
no compilation happens inside the Python build.

The wheel is intentionally not built with `pip install` from source:
that would require Emscripten on the user's machine. CI builds the wheel
once on Linux and publishes it to PyPI as `py3-none-any`.

## Project

Documentation — guides, compatibility notes, and the per-runtime API
reference — is at <https://formulon.libraz.net>. Source, design notes,
and the oracle test suite live at <https://github.com/libraz/formulon>.

## License

Apache License 2.0. See [LICENSE](./LICENSE) and [NOTICE](./NOTICE).
