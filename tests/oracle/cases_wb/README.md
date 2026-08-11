# tests/oracle/cases_wb

Workbook oracle scaffold. Pins Formulon's workbook-level behaviour --
pivot tables and print areas -- against committed JSON goldens, so
changes to those features surface as a CI failure rather than as a
silent regression.

This pipeline is **separate** from the three existing oracle tracks
(the function oracle under `cases/` + `golden/`, the conditional-
formatting oracle under `cases_cf/` + `golden_cf/`, and the IronCalc
variant oracle). The function oracle fixates `=FORMULA` evaluation
results; a workbook case does not fit that shape. A workbook case is a
declarative mini-workbook spec -- cell data across sheets, column widths
and row heights, plus optional `pivot` and `print` feature blocks -- and
the golden records the observed pivot / print result. The data model is
large enough to warrant its own schema and its own test binary.

## Overview

```
 cases_wb/<suite>.yaml         human-authored source of truth
 cases_wb/<suite>.case.json    JSON mirror (validator + workbook driver consume this)
 golden_wb/<suite>.golden.json Excel-actual golden (generated on a Windows + Excel host)
 tests/oracle/workbook_oracle_test.cpp
                               parameterized C++ verifier; registers zero cases
                               when no goldens exist, so the build stays green
 tools/oracle/workbook_case_schema.py
                               Python validator for the case + golden JSON files
 tools/oracle/workbook_oracle_gen.py
                               golden generator entry point
```

## Schema

### Workbook case JSON (`cases_wb/<suite>.case.json`)

```json
{
  "suite": "example_smoke",
  "kind": "workbook",
  "description": "...",
  "cases": [
    {
      "id": "two_cells_one_sheet",
      "description": "...",
      "sheets": {
        "Sheet1": {
          "A1": {"kind": "number", "value": 10.0},
          "B1": {"kind": "text", "value": "hello"}
        }
      }
    }
  ]
}
```

Per-case field reference:

- `id` -- unique within the suite; doubles as the gtest parameter name.
- `description` -- optional free text.
- `sheets` -- mapping of sheet-name to a mapping of A1-address to cell
  value. Cell values follow the same shorthand / `{kind, value}`
  normalisation the function oracle uses (`tools/oracle/case_schema.py`).
- `column_widths` -- optional mapping of a column letter or span
  (`"A"`, `"A:D"`) to a positive width number.
- `row_heights` -- optional mapping of a row number (`"1"`) to a
  positive height number.
- `pivot` -- optional feature block. `row_fields`, `col_fields`, and
  `page_fields` name source fields. `formula_probes` is an optional
  post-build list of `{id, cell, formula}` records; the Windows driver
  writes those formulas after `RefreshTable` and records their scalar
  result. See "Pivot suites".
- `print` -- optional feature block with a checked shape. See "Print
  suites".

### Workbook golden JSON (`golden_wb/<suite>.golden.json`)

```json
{
  "suite": "example_smoke",
  "kind": "workbook",
  "environment": { "excel_version": "...", "excel_locale": "..." },
  "cases": [
    {
      "id": "two_cells_one_sheet",
      "spec": { "...the declarative case spec..." },
      "expect": { "...the observed pivot / print result..." }
    }
  ]
}
```

## Target

Reliable PivotTable COM automation is only available through Windows Excel,
so the workbook track is prepared for `win-365-ja_JP` but remains
external-pending while that target is `wanted`. The old workbook goldens are
marked `reference-only` and excluded from active CTest/coverage; they must
not be described as Microsoft 365 verified. `mac-365-ja_JP` remains the
formula-track primary.

## Adding a case

1. Author the case in `cases_wb/<suite>.yaml`.
2. Mirror the same case as JSON in `cases_wb/<suite>.case.json` (the
   YAML is for humans; the JSON is for the validator and the driver).
3. Validate the on-disk files:
   ```
   python3 tools/oracle/workbook_case_schema.py \
       tests/oracle/cases_wb/<suite>.case.json \
       [tests/oracle/golden_wb/<suite>.golden.json]
   ```
   The golden argument is optional -- omit it until the golden exists.
4. On a product-verified Microsoft 365 Windows + Excel host, generate the
   golden:
   ```
   python3 tools/oracle/cli.py workbook --suite <suite>
   ```
5. Build and run the test binary:
   ```
   cmake --build build --target formulon_workbook_oracle_tests --parallel
   cd build && ctest -R WorkbookOracle --output-on-failure
   ```

## Pivot suites

The pivot-table verifier is implemented. Three authored suites exercise
it:

- `pivot_basic` -- single / multi row fields, row + column fields, grand
  totals on/off, anchor offset.
- `pivot_aggregations` -- COUNT / AVERAGE / MAX / MIN / PRODUCT /
  CountNumbers / StdDev / VarP and multi data-field pivots.
- `pivot_layout` -- Compact / Tabular / Outline layouts, multi-level
  subtotals, and manual item filters that hide row / column items.

Each pivot case carries a `pivot` block in its `spec`:

```yaml
pivot:
  source: Data!A1:C13          # sheet-qualified A1 range; row 0 is headers
  anchor: Report!A1            # sheet-qualified A1 top-left of the pivot
  row_fields: [Region]         # source header names on the row axis
  col_fields: [Product]        # source header names on the column axis
  data_fields:                 # value-axis aggregations
    - {field: Amount, agg: Sum}
  layout: Compact              # Compact | Tabular | Outline (optional)
  grand_totals: {rows: true, cols: true}   # optional, defaults true/true
  filters:                     # optional manual item filters
    - {field: Region, hide: [South]}
```

`page_fields` places source fields on Excel's report-filter/page axis. The
optional `formula_probes` list contains `{id, cell, formula}` records. The
Windows driver writes those formulas after `RefreshTable` and records their
scalar results, allowing GETPIVOTDATA page/data-axis routing to be checked
independently of the rendered pivot grid. The C++ verifier consumes the same
probe schema once an external, product-verified golden exists.

`agg` accepts: `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`,
`CountNumbers`, `StdDev`, `StdDevP`, `Var`, `VarP`.

The C++ verifier (`workbook_oracle_test.cpp`) rebuilds the pivot from the
declarative `spec` via `tests/oracle/workbook_builder.{h,cpp}`, evaluates
and lays it out, then diffs the rendered grid against
`expect.pivot.grid` from the golden. `tests/oracle/workbook_builder_test.cpp`
covers the builder directly so the pipeline has local verification even
before any golden exists.

## Print suites

The print-area / pagination verifier is implemented. Three authored
suites exercise it:

- `print_basic` -- explicit and absent (used-range fallback) print
  areas, A4 vs Letter paper, portrait vs landscape, print titles, and a
  multi-area print range.
- `print_pagination` -- wide tables forcing vertical breaks (via set
  column widths), tall tables forcing horizontal breaks (via set row
  heights), and manual row / column page breaks.
- `print_fit` -- explicit scale percentages vs fit-to-width /
  fit-to-height collapse.

Each print case carries a `print` block in its `spec`:

```yaml
print:
  sheet: Sheet1                       # REQUIRED: sheet to paginate
  print_area: "A1:H80"                # optional; A1 range, may be
                                      # comma-separated multi-area.
                                      # Absent => used-range fallback
  print_titles: {rows: "1:1", cols: "A:A"}   # optional repeat tracks
  page_setup:                         # optional
    orientation: portrait             # portrait | landscape | default
    paper: 9                          # OOXML paperSize code (9 == A4)
    scale: 100                        # percentage
    fit_to_width: 1                   # pages; non-zero => fit-to-page
    fit_to_height: 0                  # pages; non-zero => fit-to-page
  manual_breaks:                      # optional manual page breaks
    rows: [40]                        # 1-based Excel row numbers
    cols: ["D"]                       # column letters
```

Row / column units: `print_titles.rows` and `manual_breaks.rows` use
1-based Excel row numbers; `print_titles.cols` and `manual_breaks.cols`
use column letters. A `manual_breaks` entry places a break *before*
that track (Excel's "insert page break" semantics). When
`fit_to_width` or `fit_to_height` is non-zero the pagination engine
derives a shrink factor and ignores `scale`.

Case-level `column_widths` (`"A:D"` span -> char-unit width) and
`row_heights` (`"3"` 1-based row -> point height) feed the sheet layout
so the pagination geometry is deterministic.

The C++ verifier (`workbook_oracle_test.cpp`) rebuilds the workbook
from the declarative `spec` via `tests/oracle/workbook_builder.{h,cpp}`,
paginates the named sheet, and diffs the result against `expect.print`:

```json
{ "print": {
    "print_area": "A1:H80",   // resolved area as an A1 string (exact)
    "h_breaks":   [40, 80],   // 0-based row indices each h-break precedes
    "v_breaks":   [4],        // 0-based col indices each v-break precedes
    "pages":      6           // total physical page count
} }
```

Pagination 1-bit parity with Excel is best-effort -- font-metric
rounding can shift a break by one track -- so the verifier allows each
break position to differ by +/-1 from the golden, and the page count
to differ by +/-1 when the break counts disagree.

`tests/oracle/workbook_builder_test.cpp` covers the print builder
directly: `print::paginate` runs fully in C++, so those tests are real
local verification (wide table -> vertical break, tall table ->
horizontal break, manual break honored, fit-to-width collapse) even
before any golden exists.

## Status

The pivot and print scaffolding and verifiers are in place: schema
validator, generator entry point (Windows COM driver + best-effort
macOS driver), the declarative workbook builder, the C++ runner +
parameterized test, and CMake wiring. The Windows-host Excel automation
that produces the goldens runs separately. The checked-in workbook files
are historical and marked `reference-only`; until a product-verified
Microsoft 365 capture opts in through `PROVENANCE.json`,
`workbook_oracle_test.cpp` registers zero golden cases and the build stays
green. The local builder and formula-probe unit tests remain active.
