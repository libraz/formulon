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

## Surface parity

This package exposes the full `Workbook` surface of the WASM-backed
`@libraz/formulon` package — the same 171 instance methods plus the
three static factories, all marshalling to the identical C-ABI
functions. JS callers can swap between the two packages without code
changes; the `{ status, value }` envelopes and result shapes are
byte-identical. The native-only differences are operational (native
threads, no V8 ↔ WASM heap copies, no 4 GiB ceiling), not API-shaped.

The TypeScript declarations in `dist/index.d.ts` are the authoritative
method list; the categories below summarise what is registered on the
`Workbook` class (see `DefineClass` in
`src/node_addon/parts/workbook_class.cc`):

```
Static factories
  Workbook.createDefault(), createEmpty(), loadBytes(bytes)

Cells & recalc
  setNumber, setBool, setText, setBlank, setFormula
  getValue, getLambdaText
  recalc, partialRecalc, setIterative, setIterativeProgress, save

Workbook policy / catalog
  calcMode, setCalcMode, excelProfileId, setExcelProfileId
  functionMetadata, functionNames,
  localizeFunctionName, canonicalizeFunctionName
  precedents, dependents, spillInfo, getExternalLinks

Sheets & structure
  addSheet, removeSheet, renameSheet, moveSheet, sheetCount, sheetName
  insertRows, deleteRows, insertCols, deleteCols
  cellCount, cellAt, definedNameCount, definedNameAt,
  tableCount, tableAt, passthroughCount, passthroughAt
  setDefinedName

Sheet view / layout / protection
  getSheetView, setSheetZoom, setSheetFreeze, setSheetTabHidden
  getSheetProtection, setSheetProtection
  getSheetColumns, setColumnWidth, setColumnHidden, setColumnOutline
  getSheetRowOverrides, setRowHeight, setRowHidden, setRowOutline

Styles
  getCellXfIndex, setCellXfIndex, getCellXf,
  getFont, getFill, getBorder, getNumFmt
  addFont, addFill, addBorder, addNumFmt, addXf
  fontCount, fillCount, borderCount, xfCount
  cellStyleCount, cellStyleXfCount, getCellStyle, getCellStyleXf

Merges / comments / hyperlinks / validations
  addMerge, removeMerge, removeMergeAt, clearMerges, getMerges
  getComment, setComment
  addHyperlink, getHyperlinks, removeHyperlink,
  removeHyperlinkAt, clearHyperlinks
  getValidations, addValidation, removeValidationAt, clearValidations

Conditional formatting
  getConditionalFormats, addConditionalFormat,
  removeConditionalFormatAt, clearConditionalFormats, evaluateCfRange

PivotTables & pivot caches
  pivotCount, pivotCreate, pivotRemove, pivotLayout, ... (full pivot
  cache + table mutation surface; see dist/index.d.ts)

Top-level: evalFormula, version, lastErrorMessage,
           lastErrorContext, statusString
```

Prebuilt binaries are shipped per supported OS/arch slot under
`dist/prebuilds/`; from a source checkout, build with the steps below.

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
