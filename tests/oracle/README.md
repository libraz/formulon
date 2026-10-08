# tests/oracle

Oracle testing surface for Formulon. Outputs and metadata for every Excel
environment live under `targets/<target>/`. The primary formula oracle is
**Mac Excel 365 (ja-JP)**; the English formula corpus also has a dedicated
oracle gate. The workbook track uses Windows Excel 365 (ja-JP).

Pass counts and skip counts for each track are published in the root
`README.md`. Optional formula captures require verified provenance before
they enter variant CTest coverage. Historical Office 2019 data remains under
`reference/win-2019-ja_JP/` and is excluded from target discovery.

The remaining skips are explicit entries in `tests/divergence.yaml` or a
variant `divergence.yaml`: volatile/environment-bound results, host-service
dependencies, accepted Excel/Formulon divergences, or known driver limits.
They should not be treated as silent implementation stubs.

## Directory layout

```
tests/oracle/
├── README.md                     (this file)
├── CMakeLists.txt                formulon_oracle_tests + opt-in formulon_oracle_variant_tests
├── json_reader.{h,cpp}           in-test JSON reader (no external dep)
├── oracle_runner.{h,cpp}         case loader + A1 parser + variant dir resolution
├── oracle_test.cpp               parameterized TEST_P comparing Formulon to golden
├── cases/                        human-authored YAML sources (shared across all targets)
│   ├── count.yaml
│   ├── countif.yaml
│   ├── lookup.yaml
│   ├── stats.yaml
│   └── datetime.yaml
├── cases_cf/                      conditional-formatting case sources
│   └── <suite>.case.json/.yaml    declarative CF cases
├── external/ironcalc/
│   ├── fixtures/                 vendored upstream workbooks
│   └── golden/                   imported external-corpus goldens
└── targets/                      target-scoped outputs and metadata
    ├── mac-365-ja_JP/            primary formula and CF target
    │   ├── ENVIRONMENT.md
    │   ├── divergence.yaml        (optional) per-target overrides
    │   ├── golden/
    │   │   └── <suite>.golden.json
    │   └── golden_cf/
    │       ├── <suite>.golden.json
    │       └── PROVENANCE.json
    └── <target>/                 e.g. win-365-ja_JP
        ├── ENVIRONMENT.md
        ├── divergence.yaml        (optional) per-target overrides
        ├── golden/
        │   ├── <suite>.golden.json
        │   └── PROVENANCE.json     optional formula capture record
        └── golden_wb/             workbook track, when captured
            ├── <suite>.golden.json
            └── PROVENANCE.json     optional workbook capture record
```

## Flow

```
 cases/*.yaml  --[oracle-gen; host Excel]-->  targets/<target>/golden/*.golden.json
                                                        │
                                                        ▼
                           oracle-verify (ctest -L oracle; any platform)
```

- `make oracle-gen` — regenerates the selected target’s `golden/*.golden.json` from YAML
  by driving Excel on the configured target. Writes `ENVIRONMENT.md` with
  the Excel version and locale used.
- `make oracle-gen-cf` — regenerates the manifest-selected
  `tracks.cf.primary` capture (currently `mac-365-ja_JP`) and writes
  target's `golden_cf/PROVENANCE.json` with the Excel build, capture id, case counts,
  and SHA-256 for each golden. A legacy/reference marker is not active
  coverage; use `tools/oracle/.venv/bin/python tools/oracle/provenance.py
  cf-active` when an active verified capture is required.
- Workbook closure is pending-aware in normal `oracle` metadata tests. The
  strict external-capture check (`--require-active`) is registered only with
  `-DFORMULON_ORACLE_RELEASE_GATES=ON`; a label alone never adds it to
  default CTest.
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
    formula_cell: AA5         # OPTIONAL, bare A1; default = Z1
    capture: shape            # OPTIONAL, "cells" (default) or "shape"
    samples: [AA5, AB5]       # required with capture: shape
    tolerance: { abs: 1e-10 } # optional, overrides suite default
    expect: { kind: number, value: 2 }  # OPTIONAL, author-side doc only
```

`expect` in YAML is not consulted by the verifier. It's there so the
YAML is self-documenting — what the author *believed* Excel does. Truth
is whatever the golden captures.

`formula_cell` moves the formula under test off the default `Z1` for
cases whose result depends on where the formula sits — whether a
whole-axis spill still fits below the formula row, or whether the
formula's own cell falls inside the range it references. Both Excel
drivers and this verifier honour it; the field is emitted into the
golden only when a case declares it.

`capture: shape` is for a result too large to record cell by cell. The
default capture materialises every cell of a spill and stops at
`MAX_CAPTURE_CELLS` (4,096); a whole-axis spill is 16,384 cells at the
smallest. Shape capture instead records the dynamic-array shape plus the
cells listed in `samples` (absolute A1 addresses inside the spill), and
the verifier checks the shape first, then each sample at its offset from
the formula cell. Excel answers the case identically either way — only
the recording differs — so reach for it when the alternative would be
skipping the case.

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
      "formula_cell": "AA5",
      "expect": {"kind": "number", "value": 3.0}
    },
    {
      "id": "a_spill_too_large_to_materialise",
      "formula": "=A:A",
      "setup": {"A1": {"kind": "number", "value": 1}},
      "formula_cell": "Z1",
      "capture": "shape",
      "expect": {
        "kind": "array_shape",
        "shape": [1048576, 1],
        "samples": {"Z1": 1.0, "Z2": 0.0}
      }
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

Optional `applies_to: [target, ...]` field scopes an entry to specific
targets; absent means it applies to every target. Per-target overrides
live in `targets/<target>/divergence.yaml` and merge on top of the
primary file.

## Variant oracles (opt-in)

Additional captures under `targets/<target>/` support divergence research
across Excel environments. Their formula goldens enter the optional variant
binary only after the manifest and `golden/PROVENANCE.json` establish
eligibility. The dedicated ja-JP primary and en-US formula gates run
independently of this option. Directory presence alone does not add coverage.

Build the variant test binary by enabling the CMake option:

```bash
cmake -B build -DFORMULON_ORACLE_VARIANTS=ON
cmake --build build --target formulon_oracle_variant_tests
./build/tests/oracle/formulon_oracle_variant_tests \
  --gtest_filter='Oracle/OracleTest.Matches/*__win_365_ja_JP'
```

Variant TEST_P parameter names get a `__<target>` suffix
(`Oracle/OracleTest.Matches/<suite>_<case_id>__<target>`) so primary and
variant entries are always distinguishable.

A broad `ctest -L VARIANT` run covers eligible captures selected at
configure time. Use the variant binary’s `--gtest_filter` when you only need
one target. Historical data that is no longer a maintained target —
notably the mislabelled Office 2019 capture — lives under
`tests/oracle/reference/` and is deliberately excluded from the variant
harness.

Generation is host-dependent — see `tools/oracle/README.md` for the
multi-target setup walkthrough (`make oracle-setup`, target manifest,
WSL2 → Windows Excel bridge, `cli.py setup` preflight).
