# @libraz/formulon-native

Native N-API binding for [Formulon](https://github.com/libraz/formulon),
the Excel 365 calculation engine.

## What this is

`@libraz/formulon-native` is a Node.js addon (`.node`) built directly
against the engine's static archive. It exposes the same Workbook
surface as the WASM-backed `@libraz/formulon` package, but runs as a
native binary rather than via the V8 WASM runtime.

Why prefer the native build:

- **Multithreaded recalc**: pthreads are wired in by default through
  the engine's parallel scheduler.
- **No V8 ↔ WASM heap copies** on `loadBytes` / `save` (the most
  expensive operations on large workbooks).
- **Larger workbook capacity**: not bound by the WASM 4 GiB ceiling.

## Status: MVP scaffold

This package currently exposes a deliberate subset of the engine's
JS surface — enough to build, load, edit, recalc, and save a workbook
end-to-end. The full ~50-method parity with `@libraz/formulon`,
prebuilt binaries for the supported OS/arch matrix, and the parity-test
channel are forthcoming bundles. Expect the API to grow, not to break.

Methods exposed today:

```
Workbook.createDefault()
Workbook.createEmpty()
Workbook.loadBytes(bytes)
  setNumber, setBool, setText, setBlank, setFormula
  getValue
  recalc, save
  addSheet, removeSheet, renameSheet
  sheetCount, sheetName
  setDefinedName

Top-level: evalFormula, version, lastErrorMessage,
           lastErrorContext, statusString
```

## Building from source

The published package ships a prebuilt `.node` binary; from a source
checkout you build it via:

```bash
make node-native     # cmake -B build -DFM_BUILD_NODE_ADDON=ON, build
make node-package    # stage build/bin/formulon.node into dist/
make node-test       # node --test against the staged dist/
```

To use it directly out of a source checkout, point Node at the staged
package:

```bash
node -e "import('./packages/npm-native/dist/index.mjs').then(m => console.log(m.evalFormula('=1+2')))"
```

## Usage

```js
import { Workbook, evalFormula, ValueKind } from '@libraz/formulon-native';

console.log(evalFormula('=SUM(1,2,3)'));
// { status: { ok: true, ... }, value: { kind: 1, number: 6, ... } }

const wb = Workbook.createDefault();
wb.setFormula(0, 0, 0, '=1+2');
wb.recalc();
const r = wb.getValue(0, 0, 0);
if (r.status.ok && r.value.kind === ValueKind.Number) {
  console.log(r.value.number);  // 3
}
```

The `{ status, value }` envelope mirrors the WASM package's contract.
There is no exception path: every fallible method returns
`{ ok: false, status, message, context }` on failure.

## License

Apache-2.0 — see `LICENSE` and `NOTICE`.
