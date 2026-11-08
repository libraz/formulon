# @libraz/formulon

[![npm](https://img.shields.io/npm/v/@libraz/formulon)](https://www.npmjs.com/package/@libraz/formulon)
[![PyPI](https://img.shields.io/pypi/v/formulon)](https://pypi.org/project/formulon/)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net)

**Recalculate Excel workbooks and formulas in the browser and Node, with results checked against real Excel 365.**
The engine is a C++17 core compiled to WebAssembly, so it needs no Excel install, no Windows and no server: a workbook a user opens in the page is calculated on their machine and never uploaded.

Known differences from Excel are listed case by case in [`tests/divergence.yaml`](https://github.com/libraz/formulon/blob/main/tests/divergence.yaml), each with a reason and the Excel build it was last verified on.

📖 **[Documentation](https://formulon.libraz.net)** &nbsp;·&nbsp; **[WASM guide](https://formulon.libraz.net/runtimes/wasm)** &nbsp;·&nbsp; **[API](https://formulon.libraz.net/api/wasm)** &nbsp;·&nbsp; **[Demos](https://formulon.libraz.net/demos)**

## What's inside

- **Formula engine** — 526 Excel function names, 511 implemented locally; dynamic arrays, `LET` / `LAMBDA`, and `REGEX*`. [Coverage](https://formulon.libraz.net/compatibility/formula-coverage)
- **Workbook I/O** — read, recalculate and write `.xlsx` and `.xlsb`, including styles, conditional formatting, tables, pivot tables, charts and print layout. [File formats](https://formulon.libraz.net/compatibility/file-format-support)
- **Locale profiles** — 28 Excel behavior profiles across fourteen locales on Mac and Windows hosts. [Locale profiles](https://formulon.libraz.net/compatibility/locale-profiles)
- **Excel oracle** — formula results are compared bit for bit against goldens captured from Mac Excel 365 in fourteen locales. [Oracle testing](https://formulon.libraz.net/compatibility/oracle-testing)
- **Two builds** — a single-threaded default that loads anywhere, and `@libraz/formulon/threads` for parallel recalculation. [Choosing a build](#choosing-a-build)

## Installation

```bash
npm install @libraz/formulon   # browsers and Node 22+
```

The package is ES modules only (no `require`) and ships its TypeScript declarations, which need TypeScript 5.2 or later. It runs on Node 22+ and the current releases of the major browsers; the `threads` build also needs a cross-origin isolated page (see [Choosing a build](#choosing-a-build)). Deno and Bun are not tested.

Changes between releases are listed in the [CHANGELOG](https://github.com/libraz/formulon/blob/main/CHANGELOG.md). Until 1.0 a minor release may break the API; a patch release does not.

## Quick start

```js run
import createFormulon from '@libraz/formulon';

const formulon = await createFormulon();
console.log(formulon.evaluateFormula('=SUM(1,2,3)')); // 6
console.log(formulon.evaluateFormula('=1/0') instanceof formulon.ExcelError); // true
```

`evaluateFormula` returns a plain JavaScript value: `null` for a blank, a number, a boolean or a string, an `ExcelError` for an Excel error such as `#DIV/0!`, and a 2-D array for an array result. `evaluateTypedFormula` returns the tagged `Value` instead. A host-side failure throws `FormulonError`, whose `code` is a `StatusCode`. See [Errors](https://formulon.libraz.net/compatibility/errors).

## Recalculate a workbook

```js run
import createFormulon from '@libraz/formulon';
import { readFile, writeFile } from 'node:fs/promises';

const formulon = await createFormulon();
const wb = formulon.Workbook.load(await readFile('input.xlsx'));
try {
  const ws = wb.sheet(0);          // or wb.sheet('Sheet1')
  ws.setValue('A1', 42);
  ws.setFormula('A2', '=A1*2');
  console.log(wb.hasDirtyCells()); // true
  wb.recalc();
  console.log(wb.hasDirtyCells()); // false
  console.log(ws.getValue('A2')); // 84

  await writeFile('output.xlsx', wb.save());
} finally {
  wb.dispose(); // workbooks wrap a native handle; `wb.disposed` turns true
}
```

On Node 24+ and in bundlers that support explicit resource management, `using wb = formulon.Workbook.load(...)` releases the workbook at the end of the scope instead. `dispose()` frees the workbook at the moment you call it. A workbook dropped without it is released when the garbage collector reclaims it, which is best effort and may never happen, so dispose workbooks you are done with.

A `Sheet` handle stays on the same sheet when sheets are added, moved or removed around it. `wb.sheet('Sheet1')` finds a sheet by name (it throws `FormulonError` with `StatusCode.NotFound` when there is none), `wb.listSheets()` lists them all, and `ws.name` and `ws.index` read its current name and position. Every method that takes a cell accepts either `(row, col)` or an A1 string (`ws.setValue('A1', 42)`, `ws.getValue('A2')`), and a rectangle argument accepts `'A1:B2'`. `ws.getValues('A1:B2')` reads a block as rows of values and `ws.setValues('A1', [[1, 'x'], [true, null]])` writes one.

`ws.setCellFormat(range, patch)` changes the number format, font, fill, border, alignment and protection of a cell or range, and leaves every attribute the patch does not name as it was. Whole-column and whole-row spans take the same strings as the size operations:

```js run
import createFormulon from '@libraz/formulon';

const formulon = await createFormulon();
const wb = formulon.Workbook.createDefault();
try {
  const ws = wb.sheet(0);
  ws.setValue('A1', 0.256);
  ws.setCellFormat('A1', { numberFormat: '0.0%', font: { bold: true, color: '#C00000' } });
  ws.setCellFormat('B:B', { alignment: { horizontal: formulon.HorizontalAlign.Center } });
  ws.setColumnWidth('B:D', 12);
  console.log(ws.getDisplayText('A1').text); // 25.6%
  console.log(ws.getCellFormat('A1').font.bold); // true
} finally {
  wb.dispose();
}
```

Reads return the value as of the last full `recalc()`; `wb.hasDirtyCells()` tells whether edits are still waiting for one. `ErrorDisplayName[ErrorCode.Div0]` gives `'#DIV/0!'` without a workbook.

## Tables, pivot tables and charts

`addTable`, `addPivot`, `addPivotCache`, `addChart`, `addShape`, `addSlicer` and `addPicture` return a handle that stays on its object while other tables, pivots or drawings come and go. The operations on the object are methods of the handle; the workbook and the sheet keep adding, listing, looking up and removing:

```js run
import createFormulon from '@libraz/formulon';

const formulon = await createFormulon();
const wb = formulon.Workbook.createDefault();
try {
  const ws = wb.sheet(0);
  ws.setValues('A1', [['Region', 'Amount'], ['East', 10], ['West', 30]]);

  const chart = ws.addChart({ kind: formulon.ChartInsertKind.Column, source: 'Sheet1!$A$1:$B$3' });
  chart.setTitle('Sales');
  console.log(chart.getProperties().title?.text); // Sales

  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ ref: 'A1:B3', sheet: 'Sheet1' });
  cache.addField('Region');
  cache.addField('Amount');
  for (const [region, amount] of [['East', 10], ['West', 30]]) {
    const record = cache.addRecord();
    cache.setRecordValue(record, 0, region);
    cache.setRecordValue(record, 1, amount);
  }
  const pivot = ws.addPivot({ name: 'ByRegion', cacheId: cache.cacheId, anchorRow: 0, anchorCol: 4 });
  pivot.addField({ sourceName: 'Region', axis: formulon.PivotAxis.Row });
  const amount = pivot.addField({ sourceName: 'Amount', axis: formulon.PivotAxis.Value });
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: formulon.PivotAggregation.Sum });
  pivot.rename('Totals');
  console.log(ws.pivot('Totals') === pivot); // true

  ws.removePivot(pivot);
  try {
    pivot.getProperties();
  } catch (e) {
    console.log(e instanceof formulon.FormulonError && e.code === formulon.StatusCode.NotFound); // true
  }
} finally {
  wb.dispose();
}
```

A handle of a removed object throws `FormulonError` with `StatusCode.NotFound`. Tables are reached with `wb.table(nameOrIndex)`, charts and the other drawings with `ws.chart(id)`, `ws.shape(id)`, `ws.slicer(id)`, `ws.picture(id)` and `ws.drawing(id)`.

## Switching Excel profiles

New workbooks use `win-365-en_US`, the default Excel profile. Formulas are always written with English function names and the stored separators; the Excel profile decides how text is parsed and how results are rendered. Switch it per workbook and recalculate:

```js run
import createFormulon from '@libraz/formulon';

const formulon = await createFormulon();
const wb = formulon.Workbook.createDefault();
try {
  const ws = wb.sheet(0);
  ws.setFormula('A1', '=LENB("日本")');
  ws.setFormula('A2', '=VALUE("1,5")');
  ws.setFormula('A3', '=ISEVEN(2)');

  for (const id of /** @type {const} */ (['mac-365-en_US', 'mac-365-ja_JP', 'mac-365-de_DE'])) {
    wb.setProfileId(id);
    wb.recalc(); // a profile change marks formulas dirty
    console.log(id, [1, 2, 3].map((row) => ws.getDisplayText(`A${row}`).text));
  }
} finally {
  wb.dispose();
}
// mac-365-en_US [ '2', '#VALUE!', 'TRUE' ]
// mac-365-ja_JP [ '4', '#VALUE!', 'TRUE' ]
// mac-365-de_DE [ '2', '1,5', 'WAHR' ]
```

The ids are `{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}`. Every `mac-*` profile and `win-365-ja_JP` is measured against Excel; the other `win-*` profiles are estimated from the Mac measurements. On a `mac-*` profile, text comparison, sorting, lookups and grouping follow that locale's measured Mac Excel collation; `win-*` profiles add the Windows word-sort rules measured on Windows Excel 365 ja-JP to that locale's collation. The profile is not saved into the file, so store the id your application targets and apply it again after loading.

For a new workbook that matches what Excel creates in a locale, use `formulon.Workbook.createForProfile(id)` instead of `createDefault`: it selects the profile and also sets the body and phonetic fonts, the localized Normal style and first sheet names, and the page margins Excel uses there. Sheets added later take the first sheet's row metrics, page margins and phonetic settings.

## Choosing a build

| Import | Threads | Hosting requirement |
| --- | --- | --- |
| `@libraz/formulon` | none; `recalcParallel` runs serially | none; loads in any page, worker, Electron `file://` page, or Node |
| `@libraz/formulon/threads` | up to 8 workers for `recalcParallel` | a browser page must be cross-origin isolated (COOP / COEP); bundlers need ES module workers |

Bundler settings for Vite, webpack and esbuild are in [Bundler requirements](https://formulon.libraz.net/runtimes/wasm#bundler-requirements).

## Non-goals

VBA execution, legacy `.xls`, chart rendering, Power Query / DAX, pivot cache refresh from source data, live external connections, and a spreadsheet UI are permanently out of scope. See [Non-goals](https://formulon.libraz.net/compatibility/non-goals).

## License

[Apache License 2.0](./LICENSE). See also [NOTICE](./NOTICE).
