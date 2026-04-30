# @libraz/formulon

Excel 365 calculation engine, compiled to WebAssembly. Evaluates formulas,
loads and saves `.xlsx` workbooks, and aims for 1-bit compatibility with
Mac Excel 365 (ja-JP locale).

## Install

```sh
npm install @libraz/formulon
```

Requires Node.js 18 or newer (the package is shipped as ES modules).

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
- `Workbook.recalc()` -- triggers a full dependency-ordered recalculation.
- `Workbook.getValue(sheet, row, col)` -- reads a cached cell value.
- `Workbook.save()` -- serialises back to an in-memory `.xlsx` byte buffer.

Workbook handles wrap a native pointer; always call `wb.delete()` (in a
`finally` block) when done.

## Project

Source, design notes, and the oracle test suite live at
<https://github.com/libraz/formulon>.

## License

Apache License 2.0. See [LICENSE](./LICENSE) and [NOTICE](./NOTICE).
