# @libraz/formulon-native

[![npm: @libraz/formulon](https://img.shields.io/npm/v/@libraz/formulon?label=%40libraz%2Fformulon)](https://www.npmjs.com/package/@libraz/formulon)
[![PyPI](https://img.shields.io/pypi/v/formulon)](https://pypi.org/project/formulon/)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net)

Native N-API binding for [Formulon](https://github.com/libraz/formulon),
the Excel 365 calculation engine. This package is not distributed on npm;
it is built from the repository, as described under
[Building from source](#building-from-source). The badges above track the
sibling WebAssembly and Python packages, which are published.

Formula evaluation uses Formulon's default `win-365-en_US` profile. Call
`setExcelProfileId()` to select any of the 28
`{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}` profiles. The
`mac-*` profiles and `win-365-ja_JP` are measured against Excel; the other
`win-*` profiles are estimated from the Mac measurements.
On a `mac-*` profile, text comparison, sorting, lookups and grouping follow
that locale's measured Mac Excel collation; `win-*` profiles add the Windows
word-sort rules measured on Windows Excel 365 ja-JP to that locale's collation.

## What this is

`@libraz/formulon-native` is a Node.js addon (`.node`) built directly
against the engine's static archive. It exposes the shared Workbook
surface of the WASM-backed `@libraz/formulon` package — see
[Surface parity](#surface-parity) for the methods that differ — but runs
as a native binary rather than via the V8 WASM runtime.

Why prefer the native build:

- **Multithreaded recalc**: pthreads are wired in by default through
  the engine's parallel scheduler.
- **No V8 ↔ WASM heap copies** on `load` / `save` (the most
  expensive operations on large workbooks).
- **Larger workbook capacity**: not bound by the WASM package's 2 GiB
  growth ceiling (`ALLOW_MEMORY_GROWTH` without an explicit
  `MAXIMUM_MEMORY` caps growth at Emscripten's 2 GiB default).

## Surface parity

This package exposes the shared `Workbook` surface of the WASM-backed
`@libraz/formulon` package, all marshalling to the identical C-ABI
functions. Its TypeScript declarations and its native class table
register 435 instance methods plus the four static factories. Of those
instance methods, 434 are shared with WASM; zero remain WASM-only, while
`getMemoryUsage()` is a native-only helper.
The shared `Workbook` methods take the same arguments, return the same values
and throw the same `FormulonError`; switching packages still requires updating
the module import and validating the target platform's native prebuild. The
declarations are generated from the WASM package's, so the two cannot differ
in any shared record, type or method.

One method exists only here, because a native workbook lives outside the JS
heap in a way the WASM build's does not: `getMemoryUsage()` reports the
workbook's estimated native footprint. `dispose()`, also reachable as
`[Symbol.dispose]()` for a `using` declaration, releases the handle on both
packages without waiting for finalization. `recalcParallel(threadCount)` is
available on both the native and WASM packages and returns the parallel SCC
telemetry described in the API declarations. See
[Memory accounting](#memory-accounting).

Two C ABI functions bound in the Python package have no JS method of their
own; what they do is reachable from both JS packages:

- `fm_workbook_set_iterative_enabled` toggles iterative calculation and keeps
  the stored cap and threshold. In JS, read them back and pass them on:
  `const it = wb.getIterative(); wb.setIterative(on, it.maxIterations, it.maxChange)`.
- `fm_workbook_add_batch` registers fonts, fills, borders and cell xfs in one
  all-or-nothing call. JS registers the same records one call each
  (`addFont`, `addFill`, `addBorder`, `addXf`), so a failure part-way leaves
  the records added before it in place.

The TypeScript declarations in `dist/index.d.ts` are the authoritative
method list. The methods are the WASM package's, documented one by one in
the [WASM API reference](https://formulon.libraz.net/api/wasm); the rows
where the packages differ are tabulated in
[Package surfaces](https://formulon.libraz.net/api/surfaces).

`createFormulon` is both the default export and a named export.

The `prebuild` CI workflow stages a build per supported OS/arch slot into
`dist/prebuilds/` as a workflow artifact for npm-publish staging; it is not
attached to the GitHub release page or pushed to the npm registry. From a
source checkout, build with the steps below.

## Building from source

A source checkout does not include a prebuilt `.node` binary; build it
via:

```bash
make node-native     # cmake -B build -DFM_BUILD_NODE_ADDON=ON, build
make node-package    # stage build/bin/formulon.node into dist/
make node-test       # node --test against the staged dist/
```

To use it directly out of a source checkout, point Node at the staged
package:

```bash
node -e "import('./packages/npm-native/dist/index.mjs').then(async m => console.log((await m.createFormulon()).evalFormula('=1+2')))"
```

## Usage

```js
import createFormulon from '@libraz/formulon-native';

const formulon = await createFormulon();
console.log(formulon.evalFormula('=SUM(1,2,3)'));
// { kind: 1, number: 6 }

const wb = formulon.Workbook.createDefault();
try {
  wb.setFormula(0, 0, 0, '=1+2');
  wb.recalc();
  const v = wb.getValue(0, 0, 0);
  if (v.kind === formulon.ValueKind.Number) {
    console.log(v.number);  // 3
  }
} finally {
  wb.dispose(); // wb.disposed is now true
}
```

`wb.recalc()` remains serial. For synchronous parallel recalculation, call
`wb.recalcParallel(threadCount)`: `threadCount` must be a finite integer;
`0` auto-detects up to 8 workers, `1` starts no workers, and `2..8` set the
worker upper bound. Missing, fractional, non-finite, negative, or above-8
values throw `RangeError` or `TypeError`. The result is a
`ParallelRecalcStats` that reports evaluated cells, processed SCCs, parallel and
serial steps, cycle recoveries, and started/used worker counts. The five
64-bit engine counters are exposed as JavaScript `number` values and are exact
through `Number.MAX_SAFE_INTEGER`.

Failures follow the WASM package's contract. `createFormulon()` rejects when
the native prebuild is not available. After that, an engine operation that
cannot complete throws `FormulonError` with `code`, `message` and `context`.
`instanceof FormulonError` holds for the class this package exports; code that
also loads `@libraz/formulon` should test `e.name === 'FormulonError'`.

## Memory accounting

A workbook is one pointer on the JS heap and can be hundreds of megabytes
in native memory, so garbage collection on its own does not feel the
pressure of holding many of them. The addon therefore reports each
workbook's estimated native footprint to V8 as external memory when it is
created, loaded, recalculated or disposed.

How much that report influences collection is up to the runtime, and it
is not visible through `process.getMemoryUsage().external` on current Node
versions. Treat it as a hint, not as a substitute for releasing
workbooks: `dispose()` is the one mechanism that frees the memory at a
point you choose.

`wb.getMemoryUsage()` returns the current estimate — the cell store, shared
strings, passthrough parts and workbook metadata — and refreshes the
report at the same time. That refresh matters for a long-running writer:
between the points listed above the reported figure goes stale, because
filling a million cells grows the workbook without V8 hearing about it.

## Project

Documentation — guides, compatibility notes, and the per-runtime API
reference — is at <https://formulon.libraz.net>. Source, design notes,
and the oracle test suite live at <https://github.com/libraz/formulon>.

## License

Apache-2.0 — see `LICENSE` and `NOTICE`.
