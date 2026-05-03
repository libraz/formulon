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

- `CellIs`, `Top10`, `AboveAverage`, `ContainsText`,
  `NotContainsText`, `BeginsWith`, `EndsWith`, `ContainsBlanks`,
  `NotContainsBlanks`, `ContainsErrors`, `NotContainsErrors`,
  `DuplicateValues`, `UniqueValues`, `Expression` — captured as
  `DifferentialFormat` matches. The generator paints each rule's
  `dxf` with a deterministic fingerprint colour derived from the
  declared `dxf_id`, so the captured fill round-trips unambiguously
  back to the originating rule.
- `ColorScale` — captured directly; `interior_object.color` returns
  the interpolated RGB for the cell.
- `DataBar`, `IconSet` — Excel does not surface bar length or icon
  bucket through `DisplayFormat`, so the generator opens the
  workbook in Excel only to validate that the rule is loadable and
  the workbook recalculates without error, then computes per-cell
  match payloads in Python from the documented Microsoft VBA
  semantics:
  [DataBar](https://learn.microsoft.com/en-us/office/vba/api/excel.databar)
  and
  [IconSetCondition](https://learn.microsoft.com/en-us/office/vba/api/excel.iconsetcondition).
  This is a "documented-semantics pin" — Formulon's evaluator is
  locked against an independent Python implementation of the spec
  rather than against Excel's render path. The dxf_id registry is
  not used for these rules (visual rules carry render payloads
  instead).
- `TimePeriod` — covered indirectly via Expression-rule rewrites
  (see "Coverage status" below). The native `timePeriod` rule is
  deferred because Excel reads live `TODAY()` at evaluation time;
  pinning a deterministic golden requires either a frozen-clock
  capture path (libfaketime / similar) or rewriting the rule to
  reference a known serial. The Expression rewrite is the latter.

Not yet captured by the generator:

- `DataBar` / `IconSet` 1-bit Excel-render verification — Excel does
  not expose the rendered bar length or icon bucket via the public
  AppleScript surface. Capturing these against Excel-actual would
  need parsing the rendered cell pixels or reverse-engineering
  Excel's private CF state.
- `TimePeriod` four week-boundary buckets (`Last7Days`, `ThisWeek`,
  `LastWeek`, `NextWeek`) — these cannot be rewritten as a static
  Expression because the bucket boundaries shift with the live
  weekday. Capturing them deterministically needs a frozen-clock
  harness (e.g. libfaketime) so the Excel-side `TODAY()` returns a
  known value during capture; deferred.
- Native `TimePeriod` rules (the day/month variants are exercised
  through Expression rewrites today) — capture is supported in
  principle via the dxf-fingerprint trick, but pinning the golden
  requires the same frozen-clock harness as the week-boundary
  buckets. Deferred.

The generator raises `NotImplementedError` on unsupported rule types
so cases cannot silently drop coverage. Requires a macOS host with
Excel 365 (ja-JP) and Automation permission for the host terminal /
IDE (same prereqs as the function-oracle generator).

## Coverage status

1 suite (`cf_smoke`), 23 cases. Of the 18 declared `RuleType`
variants, 15 are pinned end-to-end (rule loaded by Excel + per-cell
match captured), 2 are pinned via documented-semantics compute
(`DataBar`, `IconSet`), and 1 (`TimePeriod`) is exercised via
Expression-rule rewrites for the 6 day/month buckets that admit a
deterministic translation.

Cases by rule type:

- `CellIs`        — `cellis_greater_than_50`
- `ColorScale`    — `colorscale_three_stop_red_yellow_green`
- `Top10`         — `top10_top2_of_a1_to_a5`
- `AboveAverage`  — `above_average_a1_to_a5`
- `BeginsWith`    — `begins_with_foo`
- `EndsWith`      — `ends_with_bar`
- `NotContainsText` — `not_contains_qux`
- `ContainsBlanks` — `contains_blanks_a2`
- `NotContainsBlanks` — `not_contains_blanks_a1_a3`
- `ContainsErrors` — `contains_errors_a2`
- `NotContainsErrors` — `not_contains_errors_a1_a3`
- `DuplicateValues` — `duplicate_values_a1_a3`
- `UniqueValues`  — `unique_values_a2`
- `Expression`    — `expression_even_row`,
                    `time_period_today_via_expression`,
                    `time_period_yesterday_via_expression`,
                    `time_period_tomorrow_via_expression`,
                    `time_period_last_month_via_expression`,
                    `time_period_this_month_via_expression`,
                    `time_period_next_month_via_expression`
- `ContainsText`  — `contains_text_foo`
- `DataBar`       — `databar_min0_max100`
- `IconSet`       — `iconset_three_arrows`

`dxf_id` registry (one fingerprint per dxf-bearing case): 7, 11–23
inclusive (15 dxfs), plus 24–29 for the six TimePeriod-via-Expression
cases. New dxf_ids should continue from 30.

Goldens under `golden_cf/` are Excel-actual captures verified against
Mac Excel 365 (ja-JP) for the dxf-bearing and ColorScale cases, and
documented-semantics computes for the DataBar / IconSet cases.
Re-running `make oracle-gen-cf` with no diff is the regression check.
