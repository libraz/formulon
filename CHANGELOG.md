# Changelog

All notable changes to Formulon are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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

[Unreleased]: https://github.com/libraz/formulon/compare/v0.9.0...HEAD
[0.9.0]: https://github.com/libraz/formulon/releases/tag/v0.9.0
