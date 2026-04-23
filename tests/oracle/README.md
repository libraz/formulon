# tests/oracle

Oracle testing surface for Formulon. The oracle's **source of truth** is
**Mac Excel 365 (ja-JP locale)**, captured through `tools/oracle/` and
committed here as golden JSON.

## Directory layout

```
tests/oracle/
├── README.md                     (this file)
├── CMakeLists.txt                formulon_oracle_tests target
├── ENVIRONMENT.md                Excel version / locale snapshot (oracle-gen writes this)
├── json_reader.{h,cpp}           in-test JSON reader (no external dep)
├── oracle_runner.{h,cpp}         case loader + A1 parser
├── oracle_test.cpp               parameterized TEST_P comparing Formulon to golden
├── cases/                        human-authored YAML sources
│   ├── count.yaml
│   ├── countif.yaml
│   ├── lookup.yaml
│   ├── stats.yaml
│   └── datetime.yaml
└── golden/                       Excel-generated JSON (commit this)
    └── <suite>.golden.json
```

## Flow

```
 cases/*.yaml  --[oracle-gen; macOS only]-->  golden/*.golden.json
                                                  │
                                                  ▼
                           oracle-verify (ctest -L oracle; any platform)
```

- `make oracle-gen` — regenerates every `golden/*.golden.json` from YAML
  by driving Excel. Writes `ENVIRONMENT.md` with the Excel version and
  locale used.
- `make oracle-verify` — builds the `formulon_oracle_tests` binary and
  runs it under ctest. Never starts Excel. Safe to run in CI.

## Case YAML format

```yaml
suite: count                  # category name; must match filename stem
description: ...              # free-form
tolerance: { abs: 0, rel: 0 } # suite default; cases can override
locale: ja-JP
options:
  date1904: false
  iterative: false
cases:
  - id: unique_within_suite
    description: ...          # optional
    setup:                    # optional, map A1 -> value
      A1: 10                  # shorthand: number
      A2: "text"              # shorthand: text
      A3: true                # shorthand: bool
      A4: "=SUM(1,2)"         # shorthand: formula cell
      A5: { kind: error, code: "#DIV/0!" }  # explicit error
    formula: "=COUNT(A1:A4)"
    tolerance: { abs: 1e-10 } # optional, overrides suite default
    expect: { kind: number, value: 2 }  # OPTIONAL, author-side doc only
```

`expect` in YAML is not consulted by the verifier. It's there so the
YAML is self-documenting — what the author *believed* Excel does. Truth
is whatever the golden captures.

## Golden JSON format

```json
{
  "suite": "count",
  "description": "...",
  "environment": {
    "excel_version": "Microsoft Excel for Mac 16.xx.x (Build ...)",
    "excel_locale": "ja-JP",
    "date1904": false,
    "iterative": false,
    "generated_at": "2026-04-23T10:00:00Z"
  },
  "tolerance": {"abs": 0.0, "rel": 0.0},
  "cases": [
    {
      "id": "unique_within_suite",
      "formula": "=COUNT(A1:A4)",
      "setup": {"A1": {"kind": "number", "value": 10}, "...": "..."},
      "expect": {"kind": "number", "value": 3.0}
    }
  ]
}
```

The C++ verifier consumes this file end-to-end; it never touches YAML.

## Adding cases

1. Append to the relevant YAML under `cases/` (or create a new file).
2. On a Mac: `python3 tools/oracle/oracle_gen.py --suite <name>`.
3. `make oracle-verify` should now run the new cases and pass.
4. Commit YAML + golden together.

## Accepted divergences

Known-diverging cases are listed in `tests/divergence.yaml`. Entries can
either skip oracle generation (volatile functions) or widen the verifier
tolerance (±1ulp transcendentals). Every entry requires a `reason` and
the Excel build that exhibited it.
