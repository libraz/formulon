# @libraz/formulon

[![npm](https://img.shields.io/npm/v/@libraz/formulon)](https://www.npmjs.com/package/@libraz/formulon)
[![PyPI](https://img.shields.io/pypi/v/formulon)](https://pypi.org/project/formulon/)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net)

Excel 365 calculation engine, compiled to WebAssembly for browsers and
Node. Evaluates formulas and loads, recalculates, and saves `.xlsx` and
`.xlsb` workbooks without an Excel installation. It defaults to the
`win-365-ja_JP` behavior profile; hosts can select the separately
supported `mac-365-ja_JP` profile when required. Results are checked
against real Excel, and known differences are listed in
[`tests/divergence.yaml`](https://github.com/libraz/formulon/blob/main/tests/divergence.yaml).

## Install

```sh
npm install @libraz/formulon
```

Requires Node.js 22 or newer (the package is shipped as ES modules).

## Quick start

```js
import createFormulon from '@libraz/formulon';

const Module = await createFormulon();

const r = Module.evalFormula('=SUM(1,2,3)');
console.log(r.value.number); // 6
```

`evalFormula` returns an envelope of the form `{ status, value }`. Excel
errors (e.g. `#DIV/0!`) are surfaced as a `value.kind === 4` (Error)
result rather than a failed status; only host-side problems
(out-of-memory, parser crashes) populate `status.ok === false`.

## Workbook example

```js
import createFormulon from '@libraz/formulon';
import { readFile, writeFile } from 'node:fs/promises';

const Module = await createFormulon();

// Load an existing workbook from disk.
const bytes = await readFile('input.xlsx');
const wb = Module.Workbook.loadBytes(bytes);
try {
  if (!wb.isValid()) {
    throw new Error(`load failed: ${Module.lastErrorMessage()}`);
  }

  // Mutate, recalc, save.
  wb.setNumber(0, 0, 0, 42);          // Sheet1!A1 = 42
  wb.setFormula(0, 1, 0, '=A1*2');    // Sheet1!A2 = =A1*2
  wb.recalc();

  const a2 = wb.getValue(0, 1, 0);
  console.log(a2.value.number); // 84

  const saved = wb.save();
  if (saved.status.ok) {
    await writeFile('output.xlsx', saved.bytes);
  }
} finally {
  // Always release the native handle.
  wb.delete();
}
```

## Choosing a build

The package ships two builds of the same API:

| Import | Threads | Hosting requirement |
| --- | --- | --- |
| `@libraz/formulon` | none | none; loads in any page, worker, Electron `file://` page, or Node |
| `@libraz/formulon/threads` | up to 8 Web Workers for `recalcParallel` | a browser page must be cross-origin isolated |

The default build uses ordinary (non-shared) memory and starts no workers.
`recalcParallel` is still callable there and runs the pass serially,
reporting `workerThreadsStarted: 0`.

The threads build allocates its memory as a `SharedArrayBuffer` and
pre-spawns its worker pool when the factory is called, so in a browser it
only loads when the page is served with `Cross-Origin-Opener-Policy:
same-origin` and `Cross-Origin-Embedder-Policy: require-corp`. Node needs
no extra setup.

```js
import createFormulon from '@libraz/formulon/threads';
```

## Bundler integration (Vite, webpack, esbuild)

Each build is an ES module factory plus a companion `.wasm`
(`formulon.wasm`, or `formulon_threads.wasm` for the threads build).
Consumer-side concerns worth knowing about:

**1. The threads build needs ES module workers.** Its pthread workers are
spawned by Emscripten with
`new Worker(new URL("formulon_threads_core.js", import.meta.url), {type: "module"})`.
Bundlers default to classic (IIFE) workers and must be told otherwise:

```ts
// vite.config.ts
export default defineConfig({
  worker: { format: 'es' },
});
```

webpack 5 picks up the `{type: "module"}` automatically when
`output.module: true`. esbuild requires `--format=esm` for the worker
chunk.

**2. Node bridging code is compiled in.** The factory contains a Node
branch (lazy-loaded `node:module`, plus `node:worker_threads` in the
threads build) that browser bundlers will warn about as
"externalised". The warnings are harmless — the branch is dead code at
runtime in browsers — but if you want to silence them, mark the imports
external:

```ts
// vite.config.ts
export default defineConfig({
  optimizeDeps: { exclude: ['@libraz/formulon'] },
  build: {
    rollupOptions: {
      external: [/^node:/],
    },
  },
});
```

If your bundler errors on the `await import("node:...")` form rather
than just warning, raise the build target to `es2022` so the
async-function-scoped `await` lexes cleanly:

```ts
// vite.config.ts
export default defineConfig({
  build: { target: 'es2022' },
});
```

### Parallel recalculation

`Workbook.recalc()` remains the serial, caller-thread recalculation API.
`Workbook.recalcParallel(threadCount)` starts workers only in the
`@libraz/formulon/threads` build; it is synchronous and returns a
`{ status, stats }` result after all workers have joined. A thread count of
`0` selects automatic detection capped at 8, `1` keeps all evaluation on the
caller thread and starts no workers, and `2..8` sets the maximum worker count.
Values above 8 return a failed status with `kInvalidArgument`.

The five 64-bit scheduler counters in `stats` are surfaced as JavaScript
`number` values (not `bigint`) and are exact through `Number.MAX_SAFE_INTEGER`
(`2^53 - 1`). `workerThreadsStarted` and `workerThreadsUsed` report actual
worker telemetry; the scheduler may use fewer workers than requested.

## API reference

The full TypeScript surface is shipped as `dist/formulon.d.ts` and is the
authoritative reference. Highlights:

- `createFormulon(opts?)` -- the default export. Returns a Promise
  resolving to a `FormulonModule`.
- `Module.evalFormula(formula)` -- one-shot evaluation against a fresh
  workbook.
- `Module.Workbook.createDefault() / createEmpty() / loadBytes(bytes)` --
  factory methods for building or loading a workbook.
- `Workbook.setNumber / setBool / setText / setBlank / setFormula` --
  cell mutators.
- `Workbook.recalc()` -- triggers a serial, full dependency-ordered
  recalculation.
- `Workbook.recalcParallel(threadCount)` -- synchronously recalculates with
  the parallel SCC scheduler and returns status plus telemetry (serial
  outside the threads build).
- `Workbook.getValue(sheet, row, col)` -- reads a cached cell value.
- `Workbook.save()` -- serialises back to an in-memory `.xlsx` byte buffer.

Workbook handles wrap a native pointer; always call `wb.delete()` (in a
`finally` block) when done.

Two C ABI capabilities are bound in the Python package but not exposed
here: `fm_workbook_set_iterative_enabled` (toggling iterative calculation
without re-supplying the existing iteration cap and residual threshold --
`setIterative(enabled, maxIterations, maxChange)` here always takes all
three) and `fm_styles_add_batch` (bulk style registration). This is a
genuine gap, not an intentional exclusion.

## Memory when loading a workbook

Both surfaces inflate each worksheet part into memory whole before
parsing it (the zip reader caps this at 100 MiB per entry, 256 MiB per
load); `loadBytes` here then always builds a DOM tree from that buffer.
The native CLI switches to a streaming parser for worksheets past 256
KiB instead, which skips building the DOM tree -- that implementation
costs binary size the WASM budget does not have -- but still holds the
same inflated XML buffer first. Sheets are read one at a time, so the
peak is per worksheet, and the practical ceiling here is the 32-bit WASM
address space. Results are identical either way.

## Project

Documentation — guides, compatibility notes, and the per-runtime API
reference — is at <https://formulon.libraz.net>, and the
[demos](https://formulon.libraz.net/demos) run this package in the
browser. Source, design notes,
and the oracle test suite live at <https://github.com/libraz/formulon>.
[formulon-cell](https://github.com/libraz/formulon-cell) is a browser
spreadsheet UI built on this package; it doubles as an integration test
and a worked example of embedding the engine.

## License

Apache License 2.0. See [LICENSE](./LICENSE) and [NOTICE](./NOTICE).
