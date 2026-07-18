# Changelog

All notable changes to Formulon are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/libraz/formulon/compare/v0.9.6...HEAD
[0.9.6]: https://github.com/libraz/formulon/compare/v0.9.5...v0.9.6
[0.9.5]: https://github.com/libraz/formulon/compare/v0.9.4...v0.9.5
[0.9.4]: https://github.com/libraz/formulon/compare/v0.9.3...v0.9.4
[0.9.3]: https://github.com/libraz/formulon/compare/v0.9.2...v0.9.3
[0.9.2]: https://github.com/libraz/formulon/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/libraz/formulon/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/libraz/formulon/releases/tag/v0.9.0
