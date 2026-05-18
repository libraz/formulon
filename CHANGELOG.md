# Changelog

All notable changes to Formulon are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/libraz/formulon/compare/v0.9.2...HEAD
[0.9.2]: https://github.com/libraz/formulon/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/libraz/formulon/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/libraz/formulon/releases/tag/v0.9.0
