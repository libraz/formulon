# Formulon

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/@libraz/formulon)](https://www.npmjs.com/package/@libraz/formulon)
[![PyPI](https://img.shields.io/pypi/v/formulon)](https://pypi.org/project/formulon/)
[![codecov](https://codecov.io/gh/libraz/formulon/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![C++17](https://img.shields.io/badge/C%2B%2B-17-blue?logo=c%2B%2B)](https://en.cppreference.com/w/cpp/17)
[![Platform](https://img.shields.io/badge/platform-Linux%20%7C%20macOS%20%7C%20WebAssembly-lightgrey)](https://github.com/libraz/formulon)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-blue)](https://formulon.libraz.net)

Formulon is a headless, Excel-compatible calculation engine — a C++17 core that defaults to the **Windows Excel 365 (ja-JP)** behavior profile, with every known divergence explicitly tracked against Excel oracle data. The same engine is packaged for the browser (WebAssembly), for Python, and for native command-line use, so a workbook recalculates to the same values wherever it runs.

No Excel installation, no Microsoft runtime, no COM automation required. The WASM build runs in browsers, Node, and Python through `wasmtime`; native CLI packages currently ship for `darwin-arm64`, `linux-x64`, and `linux-arm64`.

## Install

```bash
npm install @libraz/formulon   # JavaScript / TypeScript (WASM)
pip install formulon            # Python
```

CLI binaries are available from [GitHub Releases](https://github.com/libraz/formulon/releases).

## Why Formulon

- **Strict oracle, not aspirational compatibility.** The runtime default is `win-365-ja_JP`, and profile-specific oracle suites pin observed Excel behavior. The primary checked-in oracle remains Mac Excel 365 (ja-JP), while Windows Excel 365 (ja-JP) is verified through variant goldens. Outputs are checked for bit-level parity against golden data regenerated from the real product; every accepted divergence (transcendental ulp drift, volatile-function snapshots, Excel quirks where Formulon deliberately keeps a saner answer) is recorded case-by-case in [`tests/divergence.yaml`](tests/divergence.yaml) with a reason and the last verified Excel build.
- **One C++ core, identical results everywhere.** JS-only competitors re-run the logic in the browser and the logic on the server. Formulon ships one engine to every surface (WASM, Python, CLI) so there is no second implementation to drift.
- **WASM size budget.** CI enforces a **3.00 MiB** hard ceiling and reports a **2.50 MiB** soft ceiling. Run `make size-check` to measure the current artifact.
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
| Pivot cache recomputation | Structurally preserved; recomputation is out of scope. |
| Spreadsheet UI | A thin UI integration layer is planned; rendering is yours. |

These are **permanent** non-goals, not "not yet." The scope is finite on purpose.

## Packaging

| Surface | Name | Notes |
|---------|------|-------|
| npm | `@libraz/formulon` | WASM ESM module, type definitions included. Node 22+, browsers, workers. |
| PyPI | `formulon` | Python 3.9+ `py3-none-any` wheel that bundles `formulon_capi.wasm` plus a pure-Python wrapper. `pip` resolves the platform-specific `wasmtime` runtime. |
| GitHub Releases | `formulon-cli-<platform-arch>` | Standalone CLI binaries (`eval`, `recalc`, `dump`) for `darwin-arm64`, `linux-x64`, `linux-arm64`. |

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
`recalc` to write a recalculated workbook, and `dump` for a text snapshot.

```bash
formulon eval '=SUM(1,2,3)'
formulon recalc input.xlsx -o output.xlsx
formulon dump output.xlsx --formulas
```

`recalc` writes its success status to stderr as
`formulon: recalc: ok, wrote M bytes to 'OUT'`; pass `--quiet` to suppress it.

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

**92 oracle categories** are defined and regenerated from Mac Excel 365 ja-JP, with Windows Excel 365 ja-JP covered by the `win-365-ja_JP` variant goldens. Current local verification is `14342/14342` fast tests passing, `4026/4026` primary formula oracle cases passing with `166` documented skips. Every remaining skip is an explicit divergence, host-service dependency, volatile/environment-bound case, or driver limitation, not a silent stub. Of the 522 catalogued functions, `515` satisfy all six closure conditions (`behaviors_declared` / `cases_cover_behaviors` / `golden_present` / `divergence_documented` / `not_in_pilot` / `behavior_drift`); the remaining `7` (`FILTERXML`, `ARRAYTOTEXT`, `CONCAT`, `CHAR`, `TRUE`, `GETPIVOTDATA`, `PHONETIC`) are blocked on oracle metadata gaps — missing goldens or under-specified behavior taxonomies — rather than implementation gaps.

Beyond formula results, **pivot tables and print areas / pagination** have a dedicated **workbook oracle track** whose primary is `win-365-ja_JP` (reliable PivotTable automation needs Windows Excel COM). Goldens are captured end-to-end: the pivot suites close at `28/28`, and the `print_basic`, `print_pagination`, `print_fit`, and `print_matrix` suites pass `35/41` cases via the `formulon_workbook_oracle_tests` harness with `6` documented divergence-skip entries scoped to `win-365-ja_JP` for a known Excel PageBreakPreview COM quirk at `PageSetup.Zoom <= 50` (low-zoom column auto-breaks invert the intuitive shrink-to-fit rule; the engine emits no break, matching the geometric model rather than Excel's PBP overlay).

New workbooks use the `win-365-ja_JP` formula profile by default; callers can switch with the profile-id API (`mac-365-ja_JP`, `win-365-ja_JP`). English-locale profiles are intentionally not exposed until matching EN oracle data and verified locale-specific behavior are available. A bytecode compiler and stack-machine VM run in parallel with the tree-walker for parity verification. The OOXML reader/writer round-trips sheets, styles, conditional formatting, comments, hyperlinks, merges, data validations, defined names, tables, and pivot tables; an MS-XLSB reader/writer covers cell values and common tokenized formulas, with styles, cross-sheet 3-D references, array-constant literals, and post-2007 "future function" IDs still limited compared to the OOXML path. Workbook operations are available through the C ABI and language bindings; the CLI deliberately exposes only `eval`, `recalc`, and `dump`.

Feedback, issue reports, and oracle divergence reports are very welcome.

## Contributing

The fastest way to help right now is to **donate Excel oracle data from your locale**. If you have Excel 365 in any locale beyond Mac ja-JP, one command (`make oracle-contribute`) drives Excel, captures goldens, and walks you through the PR. See [CONTRIBUTING.md](CONTRIBUTING.md) for the full flow and the rationale for why this is community-driven.

## License

Apache License 2.0. See [LICENSE](LICENSE) and [NOTICE](NOTICE).
