# tests/oracle/cases_cf

Conditional-formatting (CF) oracle scaffold. Pins Formulon's CF
evaluator output against committed JSON goldens, so behaviour changes
to `src/cf/cf_evaluator.cpp` surface as a CI failure rather than as a
silent rendering regression in downstream UIs.

This pipeline is **separate** from the function oracle
(`tests/oracle/cases/` + `tests/oracle/golden/`). The function oracle
fixates `=FORMULA` evaluation results: case YAML carries `formula` +
cell `setup`, golden JSON carries Excel's observed `Value`. CF cases
don't fit that shape — a CF case is a workbook with cell data + a list
of CF rules + a per-cell observed format payload (resolved fill colour,
data bar length, icon bucket). The data model is large enough to
warrant its own schema and its own test binary.

## Overview

```
 cases_cf/<suite>.yaml         human-authored source of truth
 cases_cf/<suite>.case.json    JSON mirror (validator + future Mac driver consume this)
 golden_cf/<suite>.golden.json hand-authored Formulon-self-baseline (replace with Excel-actuals once Mac driver exists)
 tests/oracle/cf_oracle_test.cpp
                               C++ smoke test that mirrors the YAML/JSON cases as
                               hardcoded gtest fixtures. Runs on Linux.
 tools/oracle/cf_case_schema.py
                               Python validator for the YAML and JSON files.
```

## Schema

### CF case JSON (`cases_cf/<suite>.case.json`)

```json
{
  "name": "cf_smoke",
  "description": "...",
  "cases": [
    {
      "id": "cellis_greater_than_50",
      "description": "...",
      "sheet": [
        {"row": 0, "col": 0, "kind": "number", "value": 10}
      ],
      "cf_blocks": [
        {
          "sqref": [{"first_row": 0, "first_col": 0, "last_row": 2, "last_col": 0}],
          "rules": [
            {
              "type": "CellIs",
              "priority": 1,
              "operator": "GreaterThan",
              "dxf_id": 7,
              "formula1": "50"
            }
          ]
        }
      ],
      "range": {"first_row": 0, "first_col": 0, "last_row": 2, "last_col": 0}
    }
  ]
}
```

Rule field reference:

- `type` — one of the 18 `RuleType` enum names: `Expression`,
  `CellIs`, `ColorScale`, `DataBar`, `IconSet`, `Top10`,
  `AboveAverage`, `ContainsText`, `NotContainsText`, `BeginsWith`,
  `EndsWith`, `ContainsBlanks`, `NotContainsBlanks`,
  `ContainsErrors`, `NotContainsErrors`, `TimePeriod`,
  `DuplicateValues`, `UniqueValues`.
- `priority` — workbook-global priority (smaller = evaluates first).
- `operator` — `cellIs` only: `LessThan` / `LessThanOrEqual` /
  `Equal` / `NotEqual` / `GreaterThan` / `GreaterThanOrEqual` /
  `Between` / `NotBetween`.
- `dxf_id` — index into `styles.dxfs` for `DifferentialFormat`-style
  rules; absent for visual rules (`ColorScale` / `DataBar` /
  `IconSet`).
- `formula1`, `formula2` — formula sources without leading `=`.
- `color_scale`, `data_bar`, `icon_set` — visual rule payloads (see
  `cf_smoke.case.json` for the `color_scale` shape).

### CF golden JSON (`golden_cf/<suite>.golden.json`)

```json
{
  "name": "cf_smoke",
  "cases": [
    {
      "id": "cellis_greater_than_50",
      "cells": [
        {
          "row": 1, "col": 0,
          "matches": [
            {"kind": "DifferentialFormat", "priority": 1, "dxf_id": 7}
          ]
        }
      ]
    }
  ]
}
```

Match shapes by `kind`:

- `DifferentialFormat` — `priority`, `dxf_id`.
- `ColorScale` — `priority`, `color: {r, g, b, a}`.
- `DataBar` — `priority`, `bar_length_pct`, `bar_axis_position_pct`,
  `is_negative`, `fill: {r, g, b, a}`.
- `IconSet` — `priority`, `icon_set_name` (e.g. `Three_Arrows`),
  `icon_index`.

### Comparison tolerance

Pinned in `cf_oracle_test.cpp` (mirrors the values documented here so
the C++ checks don't drift from author intent):

- `DifferentialFormat` — exact `dxf_id`, exact `priority`.
- `ColorScale` — `priority` exact; RGBA channels within +/- 2 each
  (interpolation rounds in different orders depending on stop count).
- `DataBar` — `priority` exact; `bar_length_pct` /
  `bar_axis_position_pct` within +/- 0.5; `is_negative` exact;
  `fill` channels within +/- 2.
- `IconSet` — exact `priority`, `icon_set_name`, `icon_index`.

## Adding a case

1. Author the case in `cases_cf/<suite>.yaml`.
2. Mirror the same case as JSON in `cases_cf/<suite>.case.json` (the
   YAML is for humans; the JSON is for the validator and the
   eventual Mac driver).
3. Author the golden in `golden_cf/<suite>.golden.json`. On macOS the
   preferred path is to let `tools/oracle/cf_oracle_gen.py` regenerate
   it from Excel; on other hosts hand-author a Formulon-self-baseline
   and document it as such in the suite description.
4. Mirror the new case in `tests/oracle/cf_oracle_test.cpp` as a
   hardcoded gtest fixture. If you change the YAML or the golden,
   mirror the change in the C++ test.
5. Validate the on-disk files:
   ```
   python3 tools/oracle/cf_case_schema.py \
       tests/oracle/cases_cf/<suite>.case.json \
       tests/oracle/golden_cf/<suite>.golden.json
   ```
6. Build and run the test binary:
   ```
   cmake --build build --target formulon_cf_oracle_tests --parallel
   cd build && ctest -R CfOracleSmoke --output-on-failure
   ```

## macOS Excel capture

`tools/oracle/cf_oracle_gen.py` drives Mac Excel 365 to capture
Excel-actual fills for each CF case. The generator builds one xlsx
per case via openpyxl, opens it under a hidden Excel.app instance,
recalculates, and reads `DisplayFormat.Interior` per cell in the case
range. The captured fills are reverse-mapped against rule descriptors
to emit the golden JSON.

Run it via:

```
make oracle-gen-cf                  # all suites
make oracle-gen-cf SUITE=cf_smoke   # one suite
```

macOS uses snake_case appscript attribute names, distinct from the
Windows COM `DisplayFormat.Interior.Color` path:

```
cell.api.display_format.interior_object.color    # -> [r, g, b]
cell.api.display_format.interior_object.pattern  # -> k.pattern_solid / k.pattern_none
```

Supported rule types:

- `CellIs` — captured as a `DifferentialFormat` match. The generator
  paints each rule's `dxf` with a deterministic fingerprint colour
  derived from the declared `dxf_id`, so the captured fill round-trips
  unambiguously back to the originating rule.
- `ColorScale` — captured directly; `interior_object.color` returns
  the interpolated RGB for the cell.

Not yet captured by the generator:

- `DataBar`, `IconSet` — Excel does not expose the rendered bar length
  or the icon bucket through `DisplayFormat`. Capturing these would
  require either parsing the rendered cell or reverse-engineering
  Excel's private CF state. Smoke cases exercising these rule types
  stay on hand-authored Formulon-self-baseline goldens until a richer
  capture path lands.
- `Top10`, `AboveAverage`, text rules (`ContainsText` /
  `NotContainsText` / `BeginsWith` / `EndsWith`), blank/error rules
  (`ContainsBlanks` / `NotContainsBlanks` / `ContainsErrors` /
  `NotContainsErrors`), `TimePeriod`, `DuplicateValues`,
  `UniqueValues`, `Expression` — all `DifferentialFormat`-bearing.
  They could reuse the same dxf-fingerprint trick as `CellIs`; they
  are deferred until the smoke suite gains cases that exercise them.

The generator raises `NotImplementedError` on unsupported rule types
so cases cannot silently drop coverage. Requires a macOS host with
Excel 365 (ja-JP) and Automation permission for the host terminal /
IDE (same prereqs as the function-oracle generator).

## Coverage status

1 suite (`cf_smoke`), 2 cases:

- `cellis_greater_than_50` — `CellIs > 50` over `A1:A3` (Excel-actual).
- `colorscale_three_stop_red_yellow_green` — 3-stop colour scale
  (Excel-actual).

Goldens under `golden_cf/` are Excel-actual captures verified against
Mac Excel 365 (ja-JP). Re-running `make oracle-gen-cf` with no diff
is the regression check.
