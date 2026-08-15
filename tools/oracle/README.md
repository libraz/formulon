# tools/oracle

Generates Formulon's **oracle golden JSON** by driving Excel under controlled
options. The primary oracle is **Mac Excel 365 (ja-JP)**; additional Excel
environments (Windows, other locales) are supported as opt-in **variants**.

- **Inputs**: `tests/oracle/cases/*.yaml` (shared across all targets)
- **Primary outputs**: `tests/oracle/golden/*.golden.json` + `tests/oracle/ENVIRONMENT.md`
- **Variant outputs**: `tests/oracle/variants/<target>/golden/*.golden.json` + per-variant `ENVIRONMENT.md`

The golden JSON is what the C++ oracle test target consumes. CI never starts
Excel — generation happens on a developer machine and the resulting JSON is
committed.

## Targets

`targets.yaml` enumerates available environments. Inspect from any host:

```bash
tools/oracle/.venv/bin/python tools/oracle/cli.py list
```

Output (example, on macOS):

```
host: Darwin
primary: mac-365-ja_JP
targets:
  * mac-365-ja_JP  driver=macos_excel    runs_on=['Darwin']
    win-365-ja_JP  driver=windows_excel  runs_on=['Windows', 'Linux']
```

`*` marks the primary target. `runs_on` is the list of host OS values
(`platform.system()`) under which the target can be generated. `Linux` here
means **WSL2** — plain Linux is rejected by the bridge driver via
`/proc/version`.

Target status is part of the evidence contract, not a display label:

```bash
tools/oracle/.venv/bin/python tools/oracle/provenance.py check
tools/oracle/.venv/bin/python tools/oracle/provenance.py active-variants
tools/oracle/.venv/bin/python tools/oracle/provenance.py cf-active
```

Only a non-primary target with `status: scaffolded` and a verified
`PROVENANCE.json` (`classification: active`, `verified: true`,
`active_ctest: true`) can enter default variant CTest coverage. `status:
wanted` is a reserved slot. The old `win-365-ja_JP` directory is retained as
`reference-only`; its `16.0`/Click-to-Run values do not prove Microsoft 365
and must not be described as a verified Windows 365 golden. A fresh Windows
capture must pass the built-in `ARRAYTOTEXT` M365 sentinel before any golden
is committed.

The conditional-formatting track is declared separately in
`targets.yaml`:
`tracks.cf.primary: mac-365-ja_JP`. `cf_oracle_gen.py` resolves and validates
that target (driver, host OS, locale, and M365 sentinel) before writing
`golden_cf/*.golden.json`. A complete run also writes
`golden_cf/PROVENANCE.json`, including the Excel build, capture id, suite
case counts, and SHA-256 for every golden. `cf-active` is the explicit
check for a verified capture; a legacy/reference marker is intentionally not
active coverage.

## Setup

`make oracle-setup` auto-detects the host and routes to the right path:

| Host | Routes to | Installs |
|------|-----------|----------|
| macOS (Darwin) | `oracle-setup-mac` | rye → tools/oracle/.venv (xlwings) |
| WSL2 (Linux + Microsoft kernel) | `oracle-setup-wsl` | rye → tools/oracle/.venv + Windows-side instructions |

After setup, verify the host is fully ready:

```bash
tools/oracle/.venv/bin/python tools/oracle/cli.py setup
```

This runs preflight checks per target (xlwings import, Excel reachability,
WSL bridge wiring) and reports PASS / FAIL / SKIP with copy-pasteable hints.

### Mac primary (required for all PRs)

```bash
make oracle-setup
# Grant Automation permission when macOS prompts:
#   System Settings -> Privacy & Security -> Automation
#     -> (your terminal) -> Microsoft Excel
```

### WSL2 → Windows (external probe; currently wanted/reference-only)

Prerequisites:
- Windows host with Office 365 installed, signed in, and activated. The
  Office display language can be any locale we ship a target for — the
  driver detects it via `Application.International(xlCountryCode)` and
  records it in `ENVIRONMENT.md`.
- WSL2 (Ubuntu / Debian / etc.) with `python.exe` reachable from `$PATH`.

One-time on Windows (PowerShell):

```powershell
winget install Python.Python.3.12
py -m pip install xlwings pywin32 pyyaml
```

One-time on WSL2:

```bash
make oracle-setup            # auto-detects WSL2; runs rye sync

# Tell the bridge where Windows-side python.exe lives. Two options:
#   (a) preferred — export an env var so your per-machine path never
#       lands in a committed file:
export FORMULON_WIN_PYTHON="/mnt/c/Users/<you>/AppData/Local/Programs/Python/Python312/python.exe"
#   (b) for a private fork — add `win_python:` under the variant target
#       in tools/oracle/targets.yaml (note: this is .gitignored from
#       upstream; do not commit author-specific paths).

make oracle-setup            # re-run; all preflight checks should now PASS
```

## Generate goldens

```bash
# Primary (Mac), all suites
make oracle-gen

# Primary, single suite
make oracle-gen SUITE=count

# Conditional-formatting primary (Mac Excel 365 ja-JP)
make oracle-gen-cf
make oracle-gen-cf SUITE=cf_smoke

# External probe by target name (requires a product-verified Windows host)
# The wanted win-365-ja_JP target is intentionally reference-only until
# Microsoft 365 is verified by the driver's preflight sentinel.
make oracle-gen TARGET=win-365-ja_JP
make oracle-gen TARGET=win-365-ja_JP SUITE=count

# Workbook track (pivot tables + print areas). Target is auto-detected from
# the host OS; pass TARGET= to override and SUITE= to narrow.
make oracle-gen-workbook
make oracle-gen-workbook TARGET=win-365-ja_JP
make oracle-gen-workbook TARGET=win-365-ja_JP SUITE=getpivotdata_page_data

# A target whose status is still `wanted` has no established provenance,
# so its capture stages outside the repository (the generator prints the
# path) and is landed separately after review. GOLDEN_DIR= overrides where
# it stages.
make oracle-promote TRACK=workbook TARGET=win-365-ja_JP DRY_RUN=1
make oracle-promote TRACK=workbook TARGET=win-365-ja_JP

# Direct CLI access (more flexibility)
tools/oracle/.venv/bin/python tools/oracle/cli.py gen --target win-365-ja_JP
tools/oracle/.venv/bin/python tools/oracle/cli.py gen --all     # all targets compatible with this host
tools/oracle/.venv/bin/python tools/oracle/oracle_gen.py --suite count --visible   # debug, shows Excel UI
```

Successful runs for a maintained target:
- Primary: updates `tests/oracle/golden/*.golden.json` + `tests/oracle/ENVIRONMENT.md`
- Variant: updates `tests/oracle/variants/<target>/golden/*.golden.json` + per-variant `ENVIRONMENT.md`

## Promoting a staged capture

`status: wanted` means the project has not yet decided that this host's
Excel is one it wants to be measured against, so the generator stages the
capture at `~/.cache/formulon/oracle-capture/<track>/<target>/` instead of
writing into the tree. A complete run on an M365 host also drops a
`PROVENANCE.candidate.json` there.

`make oracle-promote` is the second half. It re-derives each golden's
SHA-256 from the file rather than trusting the candidate, checks that every
declared suite is present and that all of them came from one capture, and
refuses a capture whose Excel cannot name itself — the shape of the Office
2019 mix-up that made this a separate step. On success it copies the
goldens in, writes the directory's `PROVENANCE.json` with
`classification: active`, and flips the manifest `status:` to
`scaffolded`.

```bash
make oracle-promote TRACK=workbook TARGET=win-365-ja_JP DRY_RUN=1   # check only
make oracle-promote TRACK=workbook TARGET=win-365-ja_JP
make oracle-promote FROM=/path/to/other/capture                     # non-default source

# Then confirm the promotion is coverage-eligible and the goldens agree:
tools/oracle/.venv/bin/python tools/oracle/provenance.py check
tools/oracle/.venv/bin/python tools/oracle/provenance.py workbook-active
make oracle-verify
```

The refusals are covered by `tools/oracle/promote_capture_test.py`, which
runs in the default CTest surface as `OracleCapturePromotion` — they are
what stands between an external golden and repository coverage, so they are
tested on every run rather than only on a host that has Excel.

Diffs on these files are part of the PR review surface. Every oracle-gen PR
should call out which cells changed and why.

## Verify (any platform)

```bash
make oracle-verify           # primary only; runs ctest -L oracle
```

The workbook closure check without an external capture is pending-aware and
belongs to the normal `oracle` metadata surface. The strict
`--require-active` release gate is registered only when CMake is configured
with `-DFORMULON_ORACLE_RELEASE_GATES=ON`; a label alone never adds it to a
default CTest run.

Maintained variants are verified through a **separate** test binary, gated
behind a CMake option. `status: wanted` targets and their historical
reference-only directories are excluded even when the option is enabled:

```bash
cmake -B build -DFORMULON_ORACLE_VARIANTS=ON
cmake --build build --target formulon_oracle_variant_tests
cd build && ctest -L VARIANT --output-on-failure
```

Variant test names get a `__<target>` suffix so primary and variant
parameters never collide. Default builds (no `-DFORMULON_ORACLE_VARIANTS=ON`)
ignore `tests/oracle/variants/` entirely.

## Adding cases

1. Pick or create `tests/oracle/cases/<category>.yaml`.
2. Append a case record:
   ```yaml
   - id: sum_mixed_range
     formula: "=SUM(A1:A3)"
     setup: { A1: 1, A2: "text", A3: 3 }
   ```
3. Regenerate the primary golden:
   ```bash
   make oracle-gen SUITE=<category>
   ```
4. (Optional) Regenerate variant goldens on the appropriate host:
   ```bash
   make oracle-gen TARGET=win-365-ja_JP SUITE=<category>
   ```
5. Commit the YAML + every refreshed `*.golden.json`.

## Divergences

Cases that Excel evaluates differently from Formulon on purpose (locale,
1-ulp drift, NOW-style volatiles) are declared in `tests/divergence.yaml`.

Per-entry options:

```yaml
entries:
  - id: my_case_id
    mode: skip-oracle               # excludes from generation entirely
    cause: harness-cannot-capture   # REQUIRED for skip-oracle
    reason: "..."
    prefer: formulon                # which side we trust if they differ
    first_noted: 2026-04-23
    applies_to: [mac-365-ja_JP]     # OPTIONAL; default = applies to all targets
```

`cause` is one of `excel-no-value`, `harness-cannot-capture`,
`accepted-divergence`, `engine-gap`, `non-identifiable`;
`divergence_check.py` rejects a skip without one and prints a per-cause
tally of skipped **cases**. Only `excel-no-value` is a case that can never
pass, which is the release gate's stated grounds for leaving skips out of
the pass-rate denominator. Of the rest, `harness-cannot-capture` and
`engine-gap` are work owed and the tally is what keeps them visible;
`accepted-divergence` and `non-identifiable` are long-lived by design, and
differ in who chose — we did in the first, the mathematics did in the
second. The full definitions live in the header of
`tests/divergence.yaml`.

Per-variant overrides live in `tests/oracle/variants/<target>/divergence.yaml`
and get merged on top of the primary file (variant entries win on key
collision).

### Re-probing skipped cases

To check whether a newer Excel build can now capture a historical
`skip-oracle` case, generate a temporary divergence file that leaves every
other skip in place:

```bash
tools/oracle/.venv/bin/python tools/oracle/reprobe_divergence.py \
  --output /tmp/formulon-reprobe-divergence.yaml \
  --allow-case regexextract_mode_3_all_capture_groups

tools/oracle/.venv/bin/python tools/oracle/oracle_gen.py \
  --target mac-365-ja_JP \
  --suite text_regex \
  --golden-dir /tmp/formulon-reprobe-golden \
  --divergence /tmp/formulon-reprobe-divergence.yaml \
  --progress
```

Use `--allow-case` multiple times to probe a set. Keeping all unrelated
skips prevents accidental hangs from volatile / external-service /
catastrophic-regex cases while still letting the target case run.

## Setup-cell address forms

Each entry in a case's `setup` map is keyed by a cell address. Two forms
are accepted:

  * **Bare A1** — e.g. `A1`, `BC42`. Targets the default sheet
    (`Sheet1`). This is the historical and most common form.
  * **Sheet-qualified A1** — e.g. `Sheet2!A1`, `'My Sheet'!B5`. The
    driver creates the named sheet inside the case's workbook on first
    reference; the C++ verifier mirrors this with `wb.add_sheet(name)`.
    Suites that use any sheet-qualified key run in per-case-workbook
    mode (one `books.add()` per case) so cross-sheet state is isolated
    case-by-case.

Single-quoted sheet names are required when the name has spaces or
other characters Excel forbids in bare references; Excel's `''` escape
for an embedded apostrophe is honoured.

The formula under test is placed at `Sheet1!Z1`; cross-sheet references
in the formula resolve through whatever sheets the setup created.

## Formula placement (`formula_cell`)

A case may override where its formula is written:

```yaml
- id: bare_full_col_below_row_one
  formula: "=A:A"
  formula_cell: AA5        # OPTIONAL; default = Z1
  setup:
    A1: 1
```

The address is a bare, relative A1 cell on the default sheet — no sheet
qualifier, no `$`, no range. Both Excel drivers write the formula there
and read the result (or the spill) back from there, and the native
verifier evaluates with the same formula cell, so `ROW()` / `COLUMN()`,
implicit intersection and spill-footprint geometry all agree between the
two sides.

Use it when the answer depends on where the formula sits: whether a
whole-axis spill still fits below the formula row, or whether the
formula's own cell falls inside the range it references. Omit it
everywhere else — an absent field is not emitted into the golden, so
existing cases are byte-identical.

## Capturing a spill too large to materialise (`capture: shape`)

The default capture walks every cell of a spill and refuses a result
above `MAX_CAPTURE_CELLS` (4,096, `drivers/base.py`) rather than emitting
a partial golden. A bare whole-axis reference alone spills at least
16,384 cells, so such a case would have to be skipped for a limitation
that is ours rather than Excel's. Shape capture avoids that:

```yaml
- id: whole_col_spills_full_declared_rectangle
  formula: "=A:A"
  formula_cell: Z1
  capture: shape
  samples: [Z1, Z2, Z3, Z4, Z1048576]
```

The driver reads the dynamic-array shape from the same non-invasive
`ROWS()/COLUMNS()` probe and then reads only the listed cells, so the
capture cost is fixed no matter how large the spill is. The golden
records

```json
"expect": {"kind": "array_shape", "shape": [1048576, 1],
           "samples": {"Z1": 1.0, "Z2": 2.0, "...": "..."}}
```

and the native verifier checks the shape first, then resolves each sample
address to its offset from the formula cell. Samples must lie inside the
spill and are capped at `case_schema.MAX_SHAPE_SAMPLES`.

## Open follow-ups

  * **Structured references (`Table1[Col1]`, `Table1[@Col1]`, totals
    rows, …)** — the engine has full `StructuredRef` evaluation
    (`src/eval/tree_walker.cpp:1218`), but no oracle suite yet. Adding
    one requires (1) a `tables:` block in the YAML schema, (2)
    `ListObjects.Add` plumbing in `windows_excel.py` /
    `macos_excel.py`, and (3) `wb.set_tables([...])` wiring in the C++
    `oracle_test.cpp`. Deliberately staged into a separate change
    because the schema extension is broader than `cross_sheet_refs`.
  * **`#GETTING_DATA`** — Excel's transient async sentinel (RTD /
    Power Query refresh). Formulon has no async evaluator, so this is
    structurally unreachable; not tracked as a divergence and not
    added to `_ERR_DISPLAY_NAMES` until a Formulon-side need surfaces.

## Architecture

```
tools/oracle/
├── targets.yaml                 declarative target manifest
├── cli.py                       cross-platform dispatcher (list / setup / gen)
├── oracle_gen.py                core generator (loads cases, calls driver, writes JSON)
├── case_schema.py               YAML loader + validator
├── driver.py                    backward-compat re-export shim
└── drivers/
    ├── base.py                  OracleDriver ABC + CaseResult / EnvironmentInfo
    ├── macos_excel.py           Mac driver (xlwings via AppleEvents)
    ├── windows_excel.py         Windows driver (xlwings via COM) + wire entrypoint
    ├── wsl_bridge.py            WSL2 wrapper that subprocess-invokes windows_excel
    └── __init__.py              select_driver(target) factory
```

The driver factory chooses the concrete class based on `target['driver']`
and the current host:

| target.driver | Darwin | Windows | WSL2 | Plain Linux |
|---|---|---|---|---|
| `macos_excel` | ✅ direct | ❌ | ❌ | ❌ |
| `windows_excel` | ❌ | ✅ direct | ✅ via wsl_bridge | ❌ |
| `wsl_bridge` (explicit) | ❌ | ❌ | ✅ direct | ❌ |

The WSL2 bridge spawns Windows-side `python.exe -m
tools.oracle.drivers.windows_excel --serve` once per oracle-gen run and
keeps it alive for the whole batch, ferrying newline-delimited JSON
requests over stdin/stdout. Without persistence, 90+ suites paid one
Excel cold-start each (5–15 s); with it, Excel boots once and every
subsequent suite reuses the workbook. Wire format (one JSON object per
line, no framing):

```json
// stdin → server
{"version": 1, "command": "run_suite", "suite_name": "...", "cases": [...]}
{"version": 1, "command": "probe_environment"}
{"version": 1, "command": "shutdown"}                                  // sent on __exit__
// stdout ← server
{"version": 1, "type": "ready", "environment": {...}}                   // sent on startup
{"version": 1, "environment": {...}, "results": [{"id": ..., "kind": ..., "value": ...}]}
{"version": 1, "type": "error",  "error": "..."}                        // surfaced on failure
```

The driver runs the Windows Python with `-X utf8=1` so traceback text
and Excel error strings come back as UTF-8 regardless of the host
console code page (CP932 on ja-JP, CP1252 on de-DE, etc.) — WSLENV
does not forward `PYTHONUTF8`, so the command-line flag is the only
reliable lever.

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `oracle-setup-mac: rye not found` | rye not installed | `brew install rye` or `curl -sSf https://rye.astral.sh/get \| bash` |
| `target driver 'windows_excel' needs Windows or WSL2 host, got Darwin` | trying to gen variant from wrong host | Run on the right host, or use `cli.py list` to see what's compatible |
| `WSL bridge requires `win_python` in targets.yaml` | neither `$FORMULON_WIN_PYTHON` nor `targets.yaml` `win_python:` is set | `export FORMULON_WIN_PYTHON=/mnt/c/.../python.exe` (preferred), or set `win_python:` under the variant target for a private fork |
| `oracle-gen: warning: target 'X' declares locale='Y' but Excel reports 'Z'` | host Excel's display language differs from the target's `locale:` | Pick the matching `--target` (or change Excel's display language and restart) — goldens record the **detected** locale, so a mismatch tags them as the wrong target |
| `windows_excel subprocess failed (rc=...)` | Win Python missing xlwings/pywin32, or Excel not activated | Re-run `make oracle-setup-wsl`, then `cli.py setup --target win-365-ja_JP` |
| Variant tests not appearing in ctest | CMake variant flag off | `cmake -B build -DFORMULON_ORACLE_VARIANTS=ON` |
