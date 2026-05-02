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
3. Author the golden in `golden_cf/<suite>.golden.json`. Until the
   Mac driver exists, this is a Formulon-self-baseline: it documents
   what Formulon currently emits, so changes to the evaluator are
   surfaced as a diff.
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

## macOS Excel capture (TODO)

The next step is `tools/oracle/cf_oracle_gen.py`: a xlwings-driven
script that opens each case as a real `.xlsx`, lets Excel evaluate the
CF rules, and reads the resolved cell formats back out. The xlwings
APIs needed:

- `wb.sheets[0].api.FormatConditions` — to install rules that mirror
  the JSON case (the script can also serialise the case's CF blocks
  directly into the xlsx package).
- `wb.sheets[0].range(addr).api.DisplayFormat.Interior.Color` —
  resolved cell-fill colour (BGR-packed `int`; needs splitting into
  RGBA).
- `wb.sheets[0].range(addr).api.DisplayFormat.Interior` — full
  interior payload, including pattern type for non-solid fills.
- `wb.sheets[0].api.FormatConditions(...).Type` /
  `.IconCriteria` /  `.BarColor` — needed to decode data-bar and
  icon-set match payloads, since `DisplayFormat` collapses them.

Status: not yet implemented. Until then, goldens in `golden_cf/` are
hand-authored Formulon-self-baselines, not Excel-actuals. The schema,
the validator, and the C++ smoke test are sized to absorb the future
real goldens without further migration. Requires a macOS host with
Excel 365 (ja-JP) and Automation permission (same prereqs as the
function-oracle generator).

## Coverage status

1 case (smoke). Goldens are Formulon-self-baselines; replacing them
with Excel-actuals is the macOS work.
