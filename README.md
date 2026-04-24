# Formulon

> **Status: Under active development. Not yet ready for production use.**
> APIs, file layout, and packaging may change without notice until the first tagged release.

Formulon is a headless, Excel-compatible calculation engine — a C++17 core
that aims to be **bit-exact against Mac Excel 365 (ja-JP)**, with every known
divergence explicitly tracked. The same engine is packaged for the browser
(WebAssembly), for Python, and for native command-line use, so a workbook
recalculates to the same values wherever it runs.

No Excel installation, no Microsoft runtime, no COM automation required.
Runs on macOS, Linux, Windows, in the browser, and in Node.

## Why Formulon

- **Strict oracle, not aspirational compatibility.** Mac Excel 365 (ja-JP)
  is the behavioral oracle. Outputs are checked for bit-level parity against
  golden data regenerated from the real product; 17 intentional divergences
  (transcendental ulp drift, volatile-function snapshots, etc.) are recorded
  in [`tests/divergence.yaml`](tests/divergence.yaml) with a reason and the
  last verified Excel build.
- **One C++ core, identical results everywhere.** JS-only competitors re-run
  the logic in the browser and the logic on the server. Formulon ships one
  engine to every surface (WASM, Python, CLI) so there is no second
  implementation to drift.
- **Strict WASM size budget.** Target **1.65 MB uncompressed / 530 KB
  Brotli**, hard ceiling **1.8 MB / 600 KB Brotli**. The budget is enforced
  in CI, not aspirational; features ship within the budget or do not ship.
- **Small dependency set.** Engine deps: `miniz` (zip), `pugixml` (XML +
  XPath 1.0). That is the complete list for the core. Linear algebra,
  number formatting, and UTF-8 handling are in-tree.
- **Readable, reviewable code.** `Expected<T, Error>` error handling, RAII,
  `-fno-exceptions -fno-rtti`, Google C++ style.

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

These are **permanent** non-goals, not "not yet." The scope is finite on
purpose.

## Packaging

| Surface | Name | Notes |
|---------|------|-------|
| npm | `@libraz/formulon` | WASM ESM module, type definitions included. Node 22+, browsers, workers. |
| PyPI | _name pending_ | CPython 3.10–3.13 wheels for macOS / Linux / Windows. |
| GitHub Releases | `formulon-cli-<os>-<arch>` | Standalone CLI binaries. |

## Status

As of 2026-04: formula parser and tree-walking evaluator are in place, with
**458 / 520 Excel functions implemented (88.1%)** across Math & Trig, Stats,
Logical, Text, Date/Time, Lookup, Financial, Engineering, Info, and
Database families. **43 oracle categories** are defined, regenerated from
Mac Excel 365 ja-JP. The OOXML reader, WASM/Python/npm packaging, CLI, and
the bytecode VM are under active construction.

Feedback, issue reports, and oracle divergence reports are welcome, but
please do not rely on Formulon for production workloads yet.

## License

Apache License 2.0. See [LICENSE](LICENSE) and [NOTICE](NOTICE).
