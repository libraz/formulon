# Formulon WASM smoke tests

Node-based smoke tests that exercise the Emscripten / embind bindings
shipped in the `formulon_wasm` target.

## Prerequisites

- A working Emscripten install (`emcmake`, `emcc` on `PATH`). See
  https://emscripten.org/docs/getting_started/downloads.html.
- Node.js 18 or later (any version that supports ES modules and
  `node:assert/strict`).

## Running

From the repository root:

```bash
make wasm        # builds build-wasm/formulon.{js,wasm} (Release)
make test-wasm   # runs tests/wasm/run.mjs against the built artifact
```

Both targets fail fast with a descriptive message when their
prerequisites are missing (no Emscripten, no Node, no build artifact).

## What is covered

`tests/wasm/run.mjs` is intentionally framework-free; it uses
`node:assert/strict` so contributors do not need an `npm install`. It
exercises every embind export at least once:

- `versionString`, `statusString`, `lastErrorMessage`, `lastErrorContext`
- `evalFormula` for `SUM`, `IF`, `CONCAT`, `1/0`, and a malformed input
- `Workbook.createDefault`, `createEmpty`, `loadBytes`
- `setNumber`, `setBool`, `setText`, `setBlank`, `setFormula`, `recalc`,
  `getValue`, `setIterative`
- `addSheet`, `sheetCount`, `sheetName`
- `cellCount`, `cellAt`
- `save()` / `loadBytes()` round-trip on a small workbook

A failing test prints `FAIL <name>` plus a stack trace, and the runner
exits 1. A successful run exits 0 with a one-line `Smoke summary:` line.

## Why this is not in `ctest`

The native `ctest` harness uses Google Test against the static archive;
running gtest under WASM requires extra plumbing (an Emscripten test
runner and a separate fixture path) that is not justified for a
boundary-only smoke suite. The Node script gives us the same
end-to-end signal — JS calls embind, embind calls the C ABI, the C ABI
drives the engine — without that complexity.

A future bundle can promote this into a real test harness (Mocha /
Vitest) if the WASM surface grows beyond what `node:assert` covers
ergonomically.
