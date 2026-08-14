# @libraz/formulon

Excel 365 calculation engine, compiled to WebAssembly. Evaluates formulas,
loads and saves `.xlsx` workbooks, and defaults to the `win-365-ja_JP`
behavior profile. Hosts can select the separately supported
`mac-365-ja_JP` profile when required.

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

## Bundler integration (Vite, webpack, esbuild)

The package ships a single ES module factory plus the companion
`formulon.wasm`. Two consumer-side concerns are worth knowing about:

**1. `recalcParallel` requires ES module workers.** The parallel recalc
scheduler runs on Web Workers spawned by Emscripten with
`new Worker(new URL("formulon.js", import.meta.url), {type: "module"})`.
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
branch (lazy-loaded `node:module` / `node:worker_threads` for Node
runtime support) that browser bundlers will warn about as
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

**3. Workers + SharedArrayBuffer require cross-origin isolation.** When
hosting in a browser, serve the page with `Cross-Origin-Opener-Policy:
same-origin` and `Cross-Origin-Embedder-Policy: require-corp` headers,
or pthread workers will refuse to start.

### Parallel recalculation

`Workbook.recalc()` remains the serial, caller-thread recalculation API.
`Workbook.recalcParallel(threadCount)` is synchronous and returns a
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
  the parallel SCC scheduler and returns status plus telemetry.
- `Workbook.getValue(sheet, row, col)` -- reads a cached cell value.
- `Workbook.save()` -- serialises back to an in-memory `.xlsx` byte buffer.

Workbook handles wrap a native pointer; always call `wb.delete()` (in a
`finally` block) when done.

## Memory when loading a workbook

`loadBytes` reads worksheet XML through the DOM parser only. The native
CLI switches to a streaming parser for worksheets past 256 KiB; that
implementation costs binary size the WASM budget does not have, so
loading here needs memory proportional to the largest single worksheet's
XML rather than a fixed window. Sheets are read one at a time, so the
peak is per worksheet, and the practical ceiling is the 32-bit WASM
address space. Results are identical either way.

## Project

Source, design notes, and the oracle test suite live at
<https://github.com/libraz/formulon>.

## License

Apache License 2.0. See [LICENSE](./LICENSE) and [NOTICE](./NOTICE).
