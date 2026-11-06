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

The package is ES modules only (no `require`) and ships its TypeScript declarations.

## Quick start

```js run
import createFormulon from '@libraz/formulon';

const formulon = await createFormulon();
console.log(formulon.evalFormula('=SUM(1,2,3)').number); // 6
```

Excel errors such as `#DIV/0!` come back as values (`v.kind === ValueKind.Error`); a host-side failure throws `FormulonError`, whose `code` is a `StatusCode`. See [Errors](https://formulon.libraz.net/compatibility/errors).

## Recalculate a workbook

```js run
import createFormulon from '@libraz/formulon';
import { readFile, writeFile } from 'node:fs/promises';

const formulon = await createFormulon();
const wb = formulon.Workbook.load(await readFile('input.xlsx'));
try {
  const ws = wb.sheet(0);          // or wb.sheet('Sheet1')
  ws.setNumber(0, 0, 42);          // A1
  ws.setFormula(1, 0, '=A1*2');    // A2
  wb.recalc();
  console.log(ws.getValue(1, 0).number); // 84

  await writeFile('output.xlsx', wb.save());
} finally {
  wb.dispose(); // workbooks wrap a native handle; `wb.disposed` turns true
}
```

On Node 24+ and in bundlers that support explicit resource management, `using wb = formulon.Workbook.load(...)` releases the workbook at the end of the scope instead.

A `Sheet` handle stays on the same sheet when sheets are added, moved or removed around it. `wb.sheet('Sheet1')` finds a sheet by name (it throws `FormulonError` with `StatusCode.NotFound` when there is none), `wb.listSheets()` lists them all, and `ws.name` and `ws.index` read its current name and position. A1 notation works through `ws.set('A1', 42)` and `ws.get('A2')`.

## Switching locale profiles

New workbooks use `win-365-en_US`. Formulas are always written with English function names and the stored separators; the profile decides how text is parsed and how results are rendered. Switch it per workbook and recalculate:

```js run
import createFormulon from '@libraz/formulon';

const formulon = await createFormulon();
const wb = formulon.Workbook.createDefault();
try {
  const ws = wb.sheet(0);
  ws.setFormula(0, 0, '=LENB("日本")');  // A1
  ws.setFormula(1, 0, '=VALUE("1,5")');  // A2
  ws.setFormula(2, 0, '=ISEVEN(2)');     // A3

  for (const id of ['mac-365-en_US', 'mac-365-ja_JP', 'mac-365-de_DE']) {
    wb.setExcelProfileId(id);
    wb.recalc(); // a profile change marks formulas dirty
    console.log(id, [0, 1, 2].map((row) => ws.getDisplayText(row, 0).text));
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
