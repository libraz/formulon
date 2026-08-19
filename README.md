# Formulon

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/@libraz/formulon)](https://www.npmjs.com/package/@libraz/formulon)
[![PyPI](https://img.shields.io/pypi/v/formulon)](https://pypi.org/project/formulon/)
[![codecov](https://codecov.io/gh/libraz/formulon/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![C++17](https://img.shields.io/badge/C%2B%2B-17-blue?logo=c%2B%2B)](https://en.cppreference.com/w/cpp/17)
[![Platform](https://img.shields.io/badge/platform-Linux%20%7C%20macOS%20%7C%20WebAssembly-lightgrey)](https://github.com/libraz/formulon)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net)

Formulon is a headless, Excel-compatible calculation engine — a C++17 core that defaults to the **Windows Excel 365 (ja-JP)** behavior profile, with every known divergence explicitly tracked against Excel oracle data. The same engine is packaged for the browser (WebAssembly), for Python, and for native command-line use, so a workbook recalculates to the same values wherever it runs.

No Excel installation, no Microsoft runtime, no COM automation required. The WASM build runs in browsers, Node, and Python through `wasmtime`; native CLI packages currently ship for `darwin-arm64`, `linux-x64`, and `linux-arm64`.

## Install

```bash
npm install @libraz/formulon   # JavaScript / TypeScript (WASM)
pip install formulon            # Python
```

CLI binaries are available from [GitHub Releases](https://github.com/libraz/formulon/releases).

## Why Formulon

- **Strict oracle, not aspirational compatibility.** The runtime default is `win-365-ja_JP`, and profile-specific oracle suites pin observed Excel behavior. Formula results are pinned against Mac Excel 365 (ja-JP); pivot tables and print layout are pinned against Windows Excel 365 (ja-JP), because reliable PivotTable automation is only available through Windows COM. Both come from a verified Microsoft 365 install. Outputs are checked for bit-level parity against golden data regenerated from the real product; every accepted divergence (transcendental ulp drift, volatile-function snapshots, Excel quirks where Formulon deliberately keeps a saner answer) is recorded case-by-case in [`tests/divergence.yaml`](tests/divergence.yaml) with a reason and the last verified Excel build.
- **One C++ core, identical results everywhere.** JS-only competitors re-run the logic in the browser and the logic on the server. Formulon ships one engine to every surface (WASM, Python, CLI) so there is no second implementation to drift.
- **WASM size budget.** CI enforces a **3.00 MiB** uncompressed and **768 KiB** Brotli hard ceiling, and reports **2.75 MiB** / **736 KiB** soft ceilings. Brotli is the wire size that actually binds, so it gates on equal footing. Run `make size-check` to measure the current artifact.
- **Small dependency set.** Engine deps: `miniz` (zip/deflate), `pugixml` (XML + XPath 1.0), `PCRE2` (Excel-compatible regex for `REGEX*`), `double-conversion` (Grisu3 shortest-roundtrip `dtoa`). Linear algebra, UTF-8 handling, and most number coercion are in-tree.
- **Readable, reviewable code.** `Expected<T, Error>` error handling, RAII, `-fno-exceptions -fno-rtti`, Google C++ style.

## What it is useful for

Anywhere a spreadsheet needs to be computed without booting Excel:

- running `.xlsx` workbooks headlessly in batch jobs or data pipelines,
- evaluating Excel-style formulas inside a web application, in the browser,
- embedding calculation into internal tools, bots, or notebooks,
- validating formulas and migrating legacy spreadsheets.

## Non-goals (by design)

Formulon deliberately does **not** cover:

| Area | Reason |
|------|--------|
| VBA execution | Security. `vbaProject.bin` is preserved byte-for-byte, never executed. |
| Legacy `.xls` (BIFF8, Excel 97–2003) | Out of scope for Excel 365 compatibility. |
| Chart / drawing rendering | Belongs to a rendering layer, not the engine. |
| PowerQuery (M) / DAX | Separate engine, separate problem domain. |
| Pivot cache regeneration | The stored `pivotCacheRecords` snapshot is preserved as-is, never rebuilt from the source range. PivotTable *results* are evaluated on demand through the API. |
| Spreadsheet UI | A thin UI integration layer is planned; rendering is yours. |

These are **permanent** non-goals, not "not yet." The scope is finite on purpose.

## Packaging

| Surface | Name | Notes |
|---------|------|-------|
| npm | `@libraz/formulon` | WASM ESM module, type definitions included. Node 22+, browsers, workers. |
| PyPI | `formulon` | Python 3.9+ `py3-none-any` wheel that bundles `formulon_capi.wasm` plus a pure-Python wrapper. `pip` resolves the platform-specific `wasmtime` runtime. |
| GitHub Releases | `formulon-cli-<platform-arch>` | Standalone CLI binaries (`eval`, `recalc`, `dump`, `paginate`) for `darwin-arm64`, `linux-x64`, `linux-arm64`. |

Every surface computes the same results from the same input. One
deliberate difference is worth knowing before you size a workload: the
WASM builds — which is to say the npm and PyPI packages — read worksheet
XML through the DOM parser only, whereas the native CLI switches to a
streaming parser for worksheets past 256 KiB. Streaming costs binary
size that the WASM budget does not have, so opening a worksheet in WASM
needs memory proportional to that worksheet's XML rather than a fixed
window. Peak use is per worksheet, not per workbook — sheets are read
one at a time — and the practical ceiling is the host's 32-bit WASM
address space.

## Command line

After placing a release binary on `PATH`, use `eval` for a scalar formula,
`recalc` to write a recalculated workbook, `dump` for a text snapshot, and
`paginate` to resolve the print area, page breaks, and page count.

```bash
formulon eval '=SUM(1,2,3)'
formulon recalc input.xlsx -o output.xlsx
formulon recalc --threads 4 input.xlsx -o output.xlsx
formulon dump output.xlsx --formulas
formulon paginate output.xlsx --sheet 0
```

All four commands accept `--` to end option parsing. Put command options
before it, then pass exactly one positional formula (`eval`) or input path
(`recalc`, `dump`, `paginate`); this also allows a relative path beginning
with `-`, for example `formulon dump --sheets -- -input.xlsx`.

`recalc` accepts `.xlsx` or `.xlsb` input/output. It writes its success status
to stderr as `formulon: recalc: ok, wrote M bytes to 'OUT'`; pass `--quiet` to
suppress that status line. XLSB data-loss warnings remain visible under
`--quiet`. By default it uses the serial recalc contract. `--threads N` opts
into the parallel SCC scheduler (`0` auto-detects, `1` stays on the caller
thread, and `2..8` sets a worker cap) and reports per-pass telemetry in the
status line.

## Status

**All 522 catalogued Excel functions are recognized**, but recognition is not the same as full Excel-compatible execution. The function catalog exposes availability explicitly; `make function-status` reports the current split.

The two top-level rows are exclusive and sum to the full 522; the
environment-bound row is a **subset of the 507 real implementations**, called
out separately because a fixed golden cannot fully describe it — it is not an
additional category (507 + 15 = 522, not 524).

| Availability | Count | Meaning | Examples |
|--------------|-------|---------|----------|
| Real implementation | 507 | Evaluates inside the normal calculation engine and is covered by unit and/or oracle tests. | Math, statistics, lookup, text, dynamic arrays |
| &nbsp;&nbsp;↳ of which environment-bound | 2 | A real implementation whose result depends on host or workbook state, so a fixed golden cannot fully describe it. Counted within the 507 above. | `INFO`, `CELL` |
| Unavailable stub | 15 | Requires external services, network I/O, COM providers, or OLAP connections that Formulon does not embed; returns a fixed unavailable error surface. | `PY`, `WEBSERVICE`, `STOCKHISTORY`, `IMAGE`, `RTD`, `TRANSLATE`, `DETECTLANGUAGE`, `COPILOT`, `CUBE*` |

**103 oracle categories** are defined. The formula and conditional-formatting tracks regenerate from Mac Excel 365 ja-JP; the workbook track regenerates from Windows Excel 365 ja-JP, and its goldens carry a capture identifier that pins every suite to a single verified Microsoft 365 session. Current local verification:

| Check | Result |
|-------|--------|
| `ctest -LE "SLOW\|BENCH\|TSAN"` — `make test`, the PR gate | all passing |
| `ctest -LE "BENCH\|TSAN"` — `make test-slow`, adds the `SLOW` tier | all passing |
| Primary formula oracle | `4423/4423` passing, `125` documented skips |
| Conditional-formatting oracle | `23/23` |
| Workbook oracle (pivot + print) | `66/66` passing, `10` documented skips |
| Imported third-party engine corpus (cross-check) | `12510/12510` passing, `168` documented divergences |

Three labels partition the CTest suite: `SLOW` (minutes-scale integration, fuzz smoke, and concurrency cases), `TSAN` (thread-sanitizer runs), and `BENCH` (microbenchmark regression checks, whose threshold is tunable, so they run on demand). Everything else is the unlabeled fast tier that CI gates on; there is no separate load-test tier.

Every skip is an explicit divergence, host-service dependency, volatile/environment-bound case, or driver limitation, not a silent stub. Of the 522 catalogued functions, `518` satisfy all six closure conditions (`behaviors_declared` / `cases_cover_behaviors` / `golden_present` / `divergence_documented` / `not_in_pilot` / `behavior_drift`); the remaining `4` (`ARRAYTOTEXT`, `FILTERXML`, `GETPIVOTDATA`, `PHONETIC`) fail only `behaviors_declared` — their behavior taxonomy is under-specified. `JIS` closes as a declared alias of `DBCS`: Excel rewrites that ja-JP formula-bar spelling before it stores or evaluates a formula, so no oracle case can name it, and the closure harness resolves the alias to the function it defers to rather than taking the declaration on trust.

Beyond formula results, **pivot tables and print areas / pagination** have a dedicated **workbook oracle track**, captured through a WSL2 → Windows COM bridge. Its ten remaining skips are all the same Excel quirk: at a print scale or zoom of 50% or less, Excel's page-break preview emits column auto-breaks that do not follow geometric pagination, so the observed break set stops shrinking with the scale and grows again at 25%. Each skipped case records the Microsoft 365 observation it was measured against.

New workbooks use the `win-365-ja_JP` formula profile by default; callers can switch with the profile-id API (`mac-365-ja_JP`, `win-365-ja_JP`). English-locale profiles are intentionally not exposed until matching EN oracle data and verified locale-specific behavior are available. A bytecode compiler and stack-machine VM exist alongside the tree-walker as an opt-in build mode (`-DFORMULON_VM_PARITY_CHECK=ON`) that cross-checks every evaluation against the tree-walker's result; every shipped artifact evaluates through the tree-walker alone. The OOXML reader/writer round-trips sheets, styles, conditional formatting, comments, hyperlinks, merges, data validations, defined names, tables, and pivot tables; an MS-XLSB reader/writer covers cell values, styles, cross-sheet 3-D references, and common tokenized formulas, with array-constant literals and post-2007 "future function" IDs still limited compared to the OOXML path. Workbook operations are available through the C ABI and language bindings; the CLI deliberately exposes only `eval`, `recalc`, `dump`, and `paginate`.

Feedback, issue reports, and oracle divergence reports are very welcome.

## Contributing

The fastest way to help right now is to **donate Excel oracle data from your locale**. If you have Excel 365 in any locale beyond Mac ja-JP, one command (`make oracle-contribute`) drives Excel, captures goldens, and walks you through the PR. See [CONTRIBUTING.md](CONTRIBUTING.md) for the full flow and the rationale for why this is community-driven.

## License

Apache License 2.0. See [LICENSE](LICENSE) and [NOTICE](NOTICE).
