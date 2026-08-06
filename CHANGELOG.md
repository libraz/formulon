# Changelog

All notable changes to Formulon are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.9.7] - 2026-08-06

### Added

- `formulon paginate`, a CLI subcommand that prints the resolved print areas,
  row and column page breaks, and page count for one sheet. It is backed by
  `fm_workbook_paginate` with an owned `fm_pagination_t`, and reaches WASM,
  Node and Python as `paginate`.
- Workbook memory-footprint estimate: `Workbook::approximate_memory_bytes()`
  and `fm_workbook_memory_usage` report an `O(cells)` pressure signal covering
  the cell store, the shared-string storage every `Text` value borrows from,
  the passthrough part payloads and the workbook metadata. The Node addon
  reports the delta to V8 as external memory on create, load, recalc and
  handle destruction, and exposes `memoryUsage()`.
- `fm_styles_add_batch` installs and deduplicates fonts, fills, borders, cell
  xfs and number formats in one call, ordering the tables so an xf can
  reference indices produced by the same batch.
- `fm_error_display_name` (`errorDisplayName` / `error_display_name`) for the
  Excel literal of a cell error code, and `fm_workbook_xlsb_read_diagnostics`
  for the undecoded formula and defined-name counters captured during an XLSB
  load.
- Phonetic text, iterative-calculation settings, extended font records
  carrying `vertAlign`, and `fm_workbook_save_xlsb_with_result`. All are
  additive, so the existing ABI is unchanged. The binding surface gains
  `getCommentResult`, conditional-format visual rules, differential formats,
  comment enumeration and pivot cache metadata alongside them.
- XLSB writer: a native `xl/styles.bin` emitter, round-tripped row/column
  layout, merged rectangles and the `date1904` flag, the mandatory worksheet
  prefix with view flags, zoom and frozen panes, and an `xl/metadata.bin`
  carrying the dynamic-array entry. Worksheet-tail records the model does not
  express — conditional formatting, data validation, hyperlinks, auto-filter,
  print setup, breaks and the drawing / table part references — are retained
  verbatim along with their sheet relationships.
- OOXML writer interns literal text cells into a generated
  `xl/sharedStrings.xml` and writes those cells as `t="s"`.
- Parser accepts a 3-D whole-column (`Sheet1:Sheet3!A:A`) or whole-row span,
  and treats a space before a quoted sheet name or a parenthesized range as
  the intersection operator.
- `GROUPBY` and `PIVOTBY` honour a `total_depth` of ±2, emitting one subtotal
  row per outer group with the outer key restated in the first key column.
  Pivot items sort by a data field's aggregate when `SortSpec::by_field` is
  set.
- A configurable process-wide structured-log sink and minimum severity level.

### Fixed

- Number-format colour names are read in the UI locale: the ja-JP profile
  accepts the localized names and the indexed `[色N]` form, while the English
  `[Red]` / `[ColorN]` spellings surface `#VALUE!` exactly as Excel does.
- Dynamic-array and lambda semantics: `BYROW` / `BYCOL` spill an errored slice
  into its own output cell instead of collapsing the call, `REDUCE` and `SCAN`
  hand the body an errored cell verbatim so an `IFERROR` guard can recover,
  and every `ArrayValue` allocation routes through a bounds-checked seam that
  validates each axis against the Excel grid.
- Numeric and statistical edge cases: `T.INV`, `F.INV` and `BETA.INV` invert
  by bisection over an expanding bracket, `YIELD` and `ODDFYIELD` clamp to the
  yield domain `PRICE` accepts, `DDB` / `DB` / `VDB` cap the schedule length,
  `LINEST` reports a finite F for a perfect fit, `FORECAST.ETS` detrends
  before detecting seasonality, and non-finite results surface as `#NUM!`.
- Number-format rounding and text functions: a shared rounding helper replaces
  ad hoc rounding in `FIXED`, `DOLLAR`, `BAHTTEXT` and the numeric renderer,
  the 32,767 UTF-16 unit cap is enforced in `SUBSTITUTE`, and the half-to-full
  katakana voicing tables are deduplicated.
- Date builtins reject serials outside Excel's `0`..`9999-12-31` range and
  broadcast array arguments cell by cell.
- `AREAS` evaluates the `INDIRECT` call it is asked to count, so a resolution
  counts as one area and a failure propagates as that error.
- Parser: the depth limit is validated against the completed AST so a flat
  left-associative chain is covered, a malformed UTF-8 lead byte is consumed
  instead of spinning the tokenizer forever, and `$`-bearing identifier runs
  such as `A$$1` are rejected.
- Structural edits propagate across every referencing model — hyperlink
  locations, data-validation formulas, pivot-cache sources, conditional-format
  and table ranges — and formulas are re-indexed when an edit rewrites a
  defined name.
- OOXML read path preserves unmodelled parts: chart, dialog and macro sheets
  come through as opaque sheets, package-level and per-sheet relationships of
  unrecognised types are re-emitted with zip-slip validation, and workbook and
  worksheet elements the model does not express survive the next save
  verbatim. XML-invalid C0 controls are escaped per context.
- XLSB Ptg codec emits reference-class Ptgs for cell and range arguments and
  value-class Ptgs for function results, decodes the `PtgMem*` markers, and
  encodes `BrtColor` with `fValidRGB` in bit 0.
- Pivot layout projection through the C API honours the workbook's Excel
  profile and the selected Compact, Tabular or Outline report layout, so the
  default ja-JP profile emits localized labels instead of the legacy English
  grid.
- Cyclic component members are ordered by address before iterating, so the
  Gauss-Seidel solver commits in a fixed order across standard-library
  implementations rather than following DFS pop order.
- A DataBar rule whose min and max thresholds are equal renders a full bar.
- CLI: `eval`, `dump` and `recalc` share one atomic file-I/O path, exit codes
  collapse to `{0, 1, 64}`, and `eval` runs through the read-only array C API
  so `=A1+1` sees an empty `A1`.
- Recalc workers launch through `launch_thread`, which returns
  `Expected<Thread, Error>` instead of terminating when the OS refuses a
  thread; the pool keeps whichever workers started and falls through to serial
  evaluation when none do.
- The evaluation and load-time arenas carry byte ceilings, so a hostile
  formula degrades to a per-cell error instead of growing until the process or
  the WASM host page aborts.

### Performance

- Whole-axis references are tracked as compact rectangle dependencies instead
  of promoting the formula to volatile, and Tarjan runs over the induced
  subgraph of dirty cells rather than the workbook-wide graph.
- Recalc layer workers are pooled behind a barrier-synchronized
  `LayerWorkerPool` for the whole parallel pass instead of being spawned and
  joined per layer.
- Row storage became `RowCells`, a contiguous run beginning at the row's first
  populated column, so memory scales with content rather than used width.
  `Sheet::read_range` appends a whole rectangle under one lock acquisition.
- The shared-string and pivot-record parts, the OOXML worksheet parse and the
  metadata shell parse run through an in-place XML loader, halving peak parse
  memory, and the workbook solely owns the passthrough payload instead of
  mirroring it.

### Changed

- The bytecode compiler, optimizer and VM compile only when
  `FORMULON_BUILD_VM` is on (defaulting to `FM_BUILD_TESTING`), so release
  CLI, WASM and binding binaries no longer carry the experimental pipeline.
- `ZipReader` drops the per-archive cumulative extraction cap; the zip-bomb
  guard narrows to the per-entry size, entry-count and compression-ratio caps.
- Per-file license and copyright headers are removed; the terms live in the
  top-level `LICENSE`.

### Testing

- Divergence and coverage governance: `divergence_check.py` validates every
  entry against a real case, suite, alias or documented non-oracle scope, and
  `golden_coverage_check.py` fails when a declared case has no golden. Each
  secondary-oracle golden file is registered as its own ctest entry so one
  allowlist exception can no longer mask every failure.
- Goldens recaptured against Excel 365 ja-JP 16.111.2, with the
  `cross_sheet_refs`, `intersect_operator`, `iterative_calc` and
  `spill_collision` suites added.
- The cross-language parity harness returns ctest's skip code when fewer than
  two channels are active, instead of reporting success on one.
- An XLSB libFuzzer target with a portable Ptg seed format, and source-seam
  guards that fail when a second `ArrayValue` allocation site or raw-XML
  retention implementation appears.

### Build / CI

- A `native-fast` job runs the fast ctest labels on develop pushes, so a red
  develop surfaces before the promotion to main.
- The WASM size report gates the Brotli wire size (768 KiB hard, 640 KiB soft)
  on equal footing with the uncompressed size; a host without `brotli` on
  `PATH` skips that half instead of failing it.
- CI fails on `expected_flakes.txt` entries that no longer name a registered
  test.

### Documentation

- The READMEs document the CLI commands, both WASM size ceilings, and the
  WASM worksheet parsing memory profile.

**Detailed Release Notes**: [GitHub Release](https://github.com/libraz/formulon/releases/tag/v0.9.7)

## [0.9.6] - 2026-07-19

### Added

- Full Excel 365 dynamic-array spill semantics. Bare ranges, arithmetic and
  comparison operators, and `IF` conditions now spill to their argument
  shape, matching Excel's implicit array evaluation; bare-range spills fill
  blanks with `0`. Scalar functions also spill each range argument element-
  wise instead of reducing to the top-left cell.

### Fixed

- Bounded every attacker-driven work path with a request-scoped budget across
  evaluation, conditional formatting, pivot, and I/O, so a hostile workbook
  can no longer force unbounded computation.
- Rejected out-of-grid coordinates at the public C API and core entry points,
  out-of-grid `Print_Titles` repeat spans, and out-of-range pivot-cache field
  indices.
- Hardened OOXML/XLSB part names, detected encrypted containers, and bounded
  `BrtExternSheet` reservation to the available payload; validated row /
  column / array bounds when decoding XLSB sheet records; cast the XLSB
  `ExternSheet` reserve count to `size_t` for the 32-bit WASM build.
- Hardened the parser against literal postfix-call, exponent overflow, and
  out-of-memory name interning.
- Aligned the VM's lexical scope with the tree-walker and made it fail closed
  on invalid opcodes.
- Recalc dependency correctness: rebuild the dependency graph on sheet
  permutation and gate evaluation on a strict parse; rewrite cell references
  on sheet rename and surface off-grid spills as `#SPILL!`; track direct
  lambda-call body dependencies and invalidate on name retarget or
  spill-phantom write; invalidate a memoized pivot layout when its source
  cache mutates.
- Reconciled the raw x14 conditional-format overlay when CF rules are removed.
- Measured the sheet-name length limit in code units and validated it on add.
- Canonicalized and localized every enumerated function through the C API.
- `formulon_cli recalc` now writes its output atomically, so a failure no
  longer destroys the original workbook; `eval --repeat` re-evaluates on each
  pass instead of being a no-op.

### Performance

- Extract zip entries into a single caller-owned buffer.

### Build / CI

- Added a repo-wide formatter / linter (biome + ruff) with a `make format`
  fan-out and a CI check-only counterpart.

**Detailed Release Notes**: [GitHub Release](https://github.com/libraz/formulon/releases/tag/v0.9.6)

## [0.9.5] - 2026-07-04

### Added

- Ad-hoc array evaluation: `evaluateFormulaArray` (C API, Node addon,
  WASM) and `evaluate_formula_array` (Python) evaluate a dynamic-array or
  spilled formula against a workbook without mutating it, returning the
  whole `Array` result instead of reducing to its top-left element. The C
  ABI exposes a two-step surface — `fm_workbook_evaluate_formula_array`
  stashes the result on the handle, `fm_workbook_evaluate_formula_array_cell`
  reads it back by row-major index.
- Function-metadata provider seam: hosts can inject localized function
  documentation over the engine's structural catalog. `fm_function_metadata`
  now also recognizes lazy-dispatch forms (`XLOOKUP`, `SUMIFS`, …) and
  parser special forms (`LET`, `LAMBDA`), and surfaces the unbounded-arity
  sentinel as `null` / `None`. Pure merge helpers `mergeFunctionMetadata`
  (Node) and `merge_function_metadata` (Python) resolve
  signature/description/localized name by locale-override → entry-default →
  engine-value precedence. Contract documented in
  `docs/function-metadata-schema.md`.

### Fixed

- Spill-phantom fidelity: `Sheet::spill_phantom_addresses` enumerates
  phantom cells across all spill regions; cell enumeration and `cell_count`
  now fold them in, `fm_workbook_cell_at` no longer treats a phantom
  coordinate as an internal error, and pagination computes the used range
  against a spilled region's full extent rather than just its anchor.
- Range-shaped defined names (e.g. `Sheet1!$A$1:$A$5`) now evaluate as a
  `Value::Array` instead of collapsing to a scalar through implicit
  intersection.

### Build / CI

- Add a `python-smoke` job to `prebuild.yml` to catch C ABI drift before
  release.

**Detailed Release Notes**: [GitHub Release](https://github.com/libraz/formulon/releases/tag/v0.9.5)

## [0.9.4] - 2026-07-03

### Added

- Read-only ad-hoc formula evaluation: `evaluateFormulaText` /
  `evaluateConditionalFormula` (C API, Node addon, WASM) evaluate formula
  text against a workbook without mutating it — resolving local/
  cross-sheet refs, defined names, and `ROW()`/`COLUMN()` anchoring, and
  reducing array/spill results to their top-left element. CF-rule
  evaluation shifts relative references from the rule's anchor and
  applies Excel's CF-predicate coercion.
- Comment enumeration: `getComments` (Node addon) / `fm_sheet_get_comment_count`
  and `fm_sheet_get_comment_at_index` (C API) list every comment on a
  sheet, including comments anchored on otherwise-empty cells.
- `show_dropdown` on data validation, round-tripped through OOXML with
  the `showDropDown` attribute's inverted semantics corrected.
- `addConditionalFormat` / `fm_sheet_cf_add_rule` now return the new
  rule's flattened index.

See the
[GitHub release page](https://github.com/libraz/formulon/releases/tag/v0.9.4)
for the full auto-generated change list.

## [0.9.3] - 2026-07-03

### Added

- Full binding-surface parity across C API, Node addon, WASM, and Python:
  pivot-cache worksheet-source/layout, sheet-view display/orientation
  flags, `save_ex` (XLSX/XLSB selector), a static cell-error setter,
  sheet-scoped defined names, CF `ColorScale`/`DataBar`/`IconSet`
  payloads, and dxf differential-format record reads. CLI `recalc`/`dump`
  pick up the same surface (extension-driven XLSB output, sheet-scoped
  name printing).
- Conditional formatting: whole-row/whole-column `sqref` support, x14
  data-bar overlay decoding (gradient, axis position, negative
  fill/border), and verbatim `extLst` passthrough.
- XLSB reader/writer closes binary-format protocol gaps: styles
  (`BrtFmt`/`BrtXF`), workbook-scope names including future functions and
  `LET`, cross-sheet 3-D references, and dynamic-array spill formulas
  (`BrtArrFmla`) — several of these previously produced `.xlsb` files
  that real Excel could not open.
- OOXML round-trip fidelity: `workbookPr`/`bookViews`/
  `workbookProtection`, `date1904`, Default-content-type passthrough,
  table style info, and per-cell style color specs (theme/indexed) all
  survive a load-modify-save cycle on real Excel-authored workbooks.
- Evaluator/parser: `date1904` threaded through the tree-walker and VM,
  defined-name resolution with circular-reference detection, whole-
  column/row range expansion against the sheet's used range, Excel's
  actual array-broadcast rule, and 3-D range tails
  (`Sheet1:Sheet3!A1:B2`).

### Fixed

- `PIVOTBY`/`GROUPBY` grand totals re-aggregate correctly for
  non-additive functions (Average/Max/Min/StdDev/Var); multi-value
  grand-total column blanking restored.
- Pivot `ShowValuesAs` (RunningTotal direction, Index, DifferenceFrom/
  PercentOf) and `GETPIVOTDATA` field-name resolution fixed.
- Conditional-formatting text rules now match numeric and blank cells
  via General-format coercion, matching Excel's SEARCH-based generated
  formula.
- Print pagination excludes hidden rows/columns from the pagination
  extent, makes column-break counting symmetric with row-break
  handling, and no longer mis-splits print-area/print-titles tokens on
  a quoted sheet name containing a comma.
- `AREAS` recurses into `CHOOSE`/`IF` reference branches instead of
  always returning 1; `INDEX`/`XLOOKUP`/`INDIRECT` range results route
  through the dynamic-array allocator instead of collapsing to a
  scalar.
- `IFNA` no longer promotes `Blank` to `0`; volatile-function detection
  is case-insensitive.

### Changed

- A shared value-kind rank centralizes `GROUPBY`/`PIVOTBY`/`SORT`
  ordering so the three comparators cannot diverge.
- WASM size ceiling raised from 1.9 MiB/600 KB soft to 3.0 MiB hard /
  2.5 MiB soft (current binary: 2.09 MiB uncompressed / 560 KiB
  Brotli).
- GitHub Actions workflows bumped to Node 24-compatible major versions
  (`actions/checkout@v7`, `setup-node@v6`, `setup-python@v6`,
  `cache@v6`, `upload-artifact@v7`, `download-artifact@v8`,
  `codecov-action@v6`, `setup-emsdk@v16`, `action-gh-release@v3`),
  ahead of GitHub's Node 20 removal.

See the
[GitHub release page](https://github.com/libraz/formulon/releases/tag/v0.9.3)
for the full auto-generated change list.

## [0.9.2] - 2026-05-18

### Added

- Workbook oracle track for pivot tables and print/pagination, driven
  through a WSL2->Windows Excel bridge with the `win-365-ja_JP` profile
  as primary. Mac and Windows tracks share a single comparator.
- Function-availability classification distinguishing win-365-only from
  cross-version functions; win-365 mode-switch semantics now match the
  oracle for the primary `win-365-ja_JP` profile.
- Print pagination captures margins and `PageBreakPreview` reads,
  retiring three divergence skips.

### Fixed

- Parser truncates numeric literals to Excel's 15-significant-digit
  representation.
- `ARRAYTOTEXT` propagates a scalar error argument through to the result.
- `PIVOTBY` layout, `MAP` / `MAKEARRAY` error spills, `FREQUENCY` bin
  ordering, `WRAPROWS` / `WRAPCOLS` shape, and `TRIMRANGE` blank
  handling align with the Mac Excel oracle.
- Print pagination suppresses auto-column breaks; the min-title-reserve
  floor avoids inverted-scale page-break drift at scale <= 50.
- OOXML round-trip preserves unknown workbook rels and shifts
  shared-formula refs correctly across cell moves (close residual cases
  from v0.9.1).
- `PERCENTILE.EXC` at the upper boundary (`pos == n`) now consistently
  returns `#NUM!`, matching Mac Excel. Previously one of two internal
  code paths returned the largest sample value.

### Changed

- Internal refactors split 12 large translation units into per-area
  files, extracted opcode metadata, introduced a binding-codegen
  pipeline for simple passthroughs, and added a shared `tests/util`
  library used by 60 test files. No user-visible API change.
- Further consolidation deduplicates aggregate kernels shared by
  `SUBTOTAL` and `AGGREGATE`, numeric-argument helpers, RGB-hex
  parsing, sheet-index validation, and OOXML/XLSB relative-path
  resolution; splits `ooxml_writer.cpp`, `number_format_tokenizer.cpp`,
  and `stats.cpp` into per-area files. No user-visible API change.

See the
[GitHub release page](https://github.com/libraz/formulon/releases/tag/v0.9.2)
for the full auto-generated change list.

## [0.9.1] - 2026-05-11

### Added

- OOXML round-trip of worksheet print settings (`pageSetup`, `pageMargins`,
  `headerFooter`, `printOptions`) and pass-through of opaque
  `printerSettings.bin` parts.

### Fixed

- OOXML round-trip preserves unknown workbook relationships and shifts
  shared-formula refs correctly across cell moves.

See the
[GitHub release page](https://github.com/libraz/formulon/releases/tag/v0.9.1)
for the full auto-generated change list.

## [0.9.0] - 2026-05-11

First public release.

### Distribution

- **npm** `@libraz/formulon`: Excel 365 calculation engine as a WebAssembly
  module. Browser, Node.js, and Web Worker compatible via embind.
- **PyPI** `formulon`: pure-Python `py3-none-any` wheel that drives a
  bundled `formulon_capi.wasm` through
  [`wasmtime`](https://pypi.org/project/wasmtime/). One wheel works on
  every platform `wasmtime` supports (Linux x86_64 / aarch64, macOS
  x86_64 / arm64, Windows x86_64).
- **CLI**: native `formulon_cli` binaries for macOS arm64, Linux x86_64,
  and Linux arm64, attached to the GitHub release as `.tar.gz`.

See the
[GitHub release page](https://github.com/libraz/formulon/releases/tag/v0.9.0)
for the full auto-generated change list.

[Unreleased]: https://github.com/libraz/formulon/compare/v0.9.7...HEAD
[0.9.7]: https://github.com/libraz/formulon/compare/v0.9.6...v0.9.7
[0.9.6]: https://github.com/libraz/formulon/compare/v0.9.5...v0.9.6
[0.9.5]: https://github.com/libraz/formulon/compare/v0.9.4...v0.9.5
[0.9.4]: https://github.com/libraz/formulon/compare/v0.9.3...v0.9.4
[0.9.3]: https://github.com/libraz/formulon/compare/v0.9.2...v0.9.3
[0.9.2]: https://github.com/libraz/formulon/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/libraz/formulon/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/libraz/formulon/releases/tag/v0.9.0
