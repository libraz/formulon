# @libraz/formulon-native

[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net)

> **Not published.** This package is not on npm and has no release page. Build it from a source checkout ([Installation](#installation)); the packages `@libraz/formulon` (npm) and `formulon` (PyPI) are the published ones.

**Recalculate Excel workbooks and formulas from Node on a native addon, with results checked against real Excel 365.**
The addon is built directly against the engine's static archive and exposes the `Workbook` surface of the WebAssembly package `@libraz/formulon`, without the V8 WebAssembly runtime.

📖 **[Documentation](https://formulon.libraz.net)** &nbsp;·&nbsp; **[API](https://formulon.libraz.net/api/wasm)** &nbsp;·&nbsp; **[Package surfaces](https://formulon.libraz.net/api/surfaces)**

## What's inside

- **Multithreaded recalc** — pthreads are wired in by default through the engine's parallel scheduler.
- **No V8 / WASM heap copies** — `load` and `save`, the most expensive operations on large workbooks, marshal straight to the engine.
- **Larger workbooks** — not bound by the WASM package's 2 GiB growth ceiling.
- **Same surface** — the declarations and the native class table register 422 instance methods plus the four static factories. Of those instance methods, 422 are shared with WASM; zero remain WASM-only. [Package surfaces](https://formulon.libraz.net/api/surfaces)
- **Excel profiles** — a new workbook and the one-shot `evaluateFormula` start on the default `win-365-en_US` profile; `wb.setProfileId()` selects any of the 28 `{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}` profiles. [Locale profiles](https://formulon.libraz.net/compatibility/locale-profiles)

## Installation

Build the addon from a checkout and use the staged package:

```bash
git clone https://github.com/libraz/formulon.git && cd formulon
make node-native     # build the addon (cmake -DFM_BUILD_NODE_ADDON=ON)
make node-package    # stage build/bin/formulon.node into packages/npm-native/dist/
```

Then import `packages/npm-native/dist/index.mjs`, or link the `packages/npm-native` directory into your project with `npm install /path/to/formulon/packages/npm-native`. A checkout carries no prebuilt `.node` binary; the `prebuild` CI workflow stages one per OS and architecture as a workflow artifact only. `createFormulon()` rejects when the native prebuild is not available.

The package is ES modules only (no `require`); `createFormulon` is both the default export and a named export. Its TypeScript declarations need TypeScript 5.2 or later. Changes between releases are in the [CHANGELOG](https://github.com/libraz/formulon/blob/main/CHANGELOG.md).

## Quick start

```js run
import createFormulon from '@libraz/formulon-native';

const formulon = await createFormulon();
console.log(formulon.evaluateFormula('=SUM(1,2,3)')); // 6

const wb = formulon.Workbook.createDefault();
try {
  const ws = wb.sheet(0);
  ws.setFormula('A1', '=1+2');
  ws.setValues('B1', [[10, 20], [30, 40]]);
  console.log(wb.hasDirtyCells()); // true
  wb.recalc();
  console.log(ws.getValue('A1')); // 3
  console.log(ws.getValues('B1:C2')); // [ [ 10, 20 ], [ 30, 40 ] ]
} finally {
  wb.dispose(); // wb.disposed is now true
}
```

`getValue` and `getValues` return plain JavaScript values: a number, string, boolean, `null` for a blank cell, and an `ExcelError` for an Excel error. `getTypedValue` / `getTypedValues` return the tagged `Value` instead. A string that starts with `=` passed to `setValue` or `setValues` is entered as a formula; enter untrusted text with `ws.setText('A1', text)`, which always stores text.

A format is set with `updateCellFormat`; only the keys you give change, and a colour is `'#RRGGBB'`:

```js run
import createFormulon from '@libraz/formulon-native';

const formulon = await createFormulon();
const wb = formulon.Workbook.createDefault();
try {
  const ws = wb.sheet(0);
  ws.setNumber('A1', 1234.5);
  ws.updateCellFormat('A1:B2', {
    numberFormat: '#,##0.00',
    font: { bold: true, color: '#C00000' },
    alignment: { horizontal: formulon.HorizontalAlign.Right },
  });
  const format = ws.getCellFormat('A1');
  console.log(format.numberFormat); // #,##0.00
  console.log(format.font.bold); // true
  console.log(ws.getDisplayText('A1').text); // 1,234.50
} finally {
  wb.dispose();
}
```

Cell methods take `(row, col)` or an A1 string, and rectangle arguments take `'A1:B2'`. In the default automatic calculation mode a read calculates pending cells first; in manual mode it returns the value as of the last `recalc()`, and `wb.hasDirtyCells()` reports edits still waiting for one.

On Node 24+, `using wb = formulon.Workbook.createDefault()` releases the workbook at the end of the scope instead.

### Parallel recalculation

`wb.recalc()` remains serial. For synchronous parallel recalculation, call `wb.recalcParallel(threadCount)`; it returns a `ParallelRecalcStats` with evaluated cells, processed SCCs, parallel and serial steps, cycle recoveries, and started and used worker counts. The five 64-bit engine counters are `number` values and are exact through `Number.MAX_SAFE_INTEGER`.

`threadCount` is `0` to auto-detect up to 8 workers, `1` to start no workers, or `2..8` for the worker upper bound. It is optional (0 = automatic, capped at 8); a non-number is `TypeError`, a fractional or non-finite number `RangeError`, a negative number or one above 8 `InvalidArgument`.

### Lifetime and errors

`dispose()`, also reachable as `[Symbol.dispose]()` for a `using` declaration, releases a workbook when you call it. A workbook dropped without it is released when V8 finalizes the object, which is best effort, so dispose workbooks you are done with.

Failures follow the WASM package's contract. After `createFormulon()` succeeds, an engine operation that cannot complete throws `FormulonError` with `code`, `message` and `context`. `instanceof FormulonError` holds for the class this package exports; code that also loads `@libraz/formulon` should test `e.name === 'FormulonError'`.

### Memory accounting

A workbook is one pointer on the JS heap and can be hundreds of megabytes in native memory, so garbage collection on its own does not feel the pressure of holding many of them. The addon therefore reports each workbook's estimated native footprint to V8 as external memory when it is created, loaded, recalculated or disposed.

How much that report influences collection is up to the runtime, and it is not visible through `process.memoryUsage().external` on current Node versions. Treat it as a hint, not as a substitute for releasing workbooks.

`wb.getMemoryUsage()` returns the current estimate (the cell store, shared strings, passthrough parts and workbook metadata) and refreshes the report at the same time. Between the points listed above the reported figure goes stale, because filling a million cells grows the workbook without V8 hearing about it.

## Build from source

```bash
make node-native     # cmake -B build -DFM_BUILD_NODE_ADDON=ON, build
make node-package    # stage build/bin/formulon.node into dist/
make node-test       # node --test against the staged dist/
```

To use it directly out of a source checkout, point Node at the staged package:

```bash
node -e "import('./packages/npm-native/dist/index.mjs').then(async m => console.log((await m.createFormulon()).evaluateFormula('=1+2')))"
```

## Documentation

Guides, compatibility notes and the per-runtime API reference are at <https://formulon.libraz.net>; source and the oracle test suite are at <https://github.com/libraz/formulon>. The TypeScript declarations in `dist/index.d.ts` are the authoritative method list.

- **Learn** — [Documentation](https://formulon.libraz.net) · [Locale profiles](https://formulon.libraz.net/compatibility/locale-profiles)
- **Reference** — [WASM API reference](https://formulon.libraz.net/api/wasm) · [Package surfaces](https://formulon.libraz.net/api/surfaces)

## Non-goals

VBA execution, legacy `.xls`, chart rendering, Power Query / DAX, live external connections, and a spreadsheet UI are permanently out of scope. See [Non-goals](https://formulon.libraz.net/compatibility/non-goals).

## License

[Apache-2.0](https://github.com/libraz/formulon/blob/main/LICENSE)
