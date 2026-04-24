# Formulon

> **Status: Under active development. Not yet ready for production use.**
> APIs, file layout, and packaging may change without notice until the first tagged release.

Formulon aims to be a **headless Excel** — a C++17 calculation engine that
evaluates Excel 365 formulas and reads/writes `.xlsx` workbooks without
requiring Excel to be installed. The same engine is packaged for the browser
(WebAssembly), for Python, and for native command-line use, so a workbook can
be recalculated identically wherever it runs.

## Goals

- **Faithful Excel 365 semantics.** Mac Excel 365 (ja-JP locale) is treated as
  the reference oracle. Known intentional differences are tracked explicitly
  rather than hidden.
- **Headless, embeddable.** No Excel installation, no GUI, no COM automation.
  Just a library you can link, import, or load.
- **One engine, many surfaces.** A single C++ core is shipped as:
  - a **WebAssembly** module on npm, for use in browsers and Node,
  - a **Python** wheel on PyPI,
  - native **CLI** binaries for Linux, macOS, and Windows.
- **Small WASM footprint.** The WebAssembly build targets a modest size
  budget so Formulon can live inside a web application without dominating
  the payload.
- **Readable, reviewable code.** Structured error handling, RAII, no
  exceptions, and a small, deliberate dependency set — so the engine stays
  auditable as it grows.

## What it is useful for

Anywhere a spreadsheet needs to be computed without booting Excel:

- running `.xlsx` workbooks headlessly in batch jobs or data pipelines,
- evaluating Excel-style formulas inside a web application, in the browser,
- embedding calculation into internal tools, bots, or notebooks,
- validating formulas and migrating legacy spreadsheets.

## Scope

Formulon focuses on the calculation engine and file I/O. It is not a
spreadsheet UI. A light integration layer (formula bar, grid bindings) is
planned to make it easy to build one on top.

## Status and roadmap

The project is pre-release. The function catalog, OOXML I/O, and packaging
are being built up milestone by milestone, and behavior is continuously
checked against the Excel oracle. Expect rough edges until the first tagged
release.

Feedback, issue reports, and oracle divergence reports are welcome, but
please do not rely on Formulon for production workloads yet.

## License

Apache License 2.0. See [LICENSE](LICENSE) and [NOTICE](NOTICE).
