# Formulon

[![CI](https://img.shields.io/github/actions/workflow/status/libraz/formulon/ci.yml?branch=main&label=CI)](https://github.com/libraz/formulon/actions/workflows/ci.yml)
[![npm](https://img.shields.io/npm/v/@libraz/formulon)](https://www.npmjs.com/package/@libraz/formulon)
[![PyPI](https://img.shields.io/pypi/v/formulon)](https://pypi.org/project/formulon/)
[![codecov](https://codecov.io/gh/libraz/formulon/branch/main/graph/badge.svg)](https://codecov.io/gh/libraz/formulon)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![C++17](https://img.shields.io/badge/C%2B%2B-17-blue?logo=c%2B%2B)](https://en.cppreference.com/w/cpp/17)
[![Platform](https://img.shields.io/badge/platform-Linux%20%7C%20macOS%20%7C%20WebAssembly-lightgrey)](https://github.com/libraz/formulon)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net)

**Formulon recalculates Excel workbooks and formulas without Excel, with results checked against real Excel 365.**
No Excel install, no Windows, no COM automation: one C++17 core ships as WebAssembly for browsers and Node, as a Python package, and as native CLI binaries, so a workbook gives the same values wherever it runs.

**Use it when you need to:**

- **Recalculate workbooks on a server** — load `.xlsx` or `.xlsb` in a batch job, CI run or data pipeline, change inputs, and save it with fresh values.
- **Run spreadsheet logic in the browser** — evaluate formulas and whole workbooks client-side, so uploaded files never leave the user's machine.
- **Keep the model in Excel** — call the workbook your team already maintains from Node or Python instead of re-implementing its formulas.
- **Get the answer Excel gives in a given locale** — Japanese byte counting, German decimal commas and localized `TRUE` / `FALSE` follow the Excel profile you pick.
- **Give AI agents spreadsheet tools** — [formulon-mcp](https://github.com/libraz/formulon-mcp) exposes the engine over MCP.

Known differences from Excel are listed case by case in [`tests/divergence.yaml`](https://github.com/libraz/formulon/blob/main/tests/divergence.yaml), each with a reason and the Excel build it was last verified on.

📖 **[Documentation](https://formulon.libraz.net)** &nbsp;·&nbsp; **[Getting started](https://formulon.libraz.net/start/)** &nbsp;·&nbsp; **[Locale profiles](https://formulon.libraz.net/compatibility/locale-profiles)** &nbsp;·&nbsp; **[Formula coverage](https://formulon.libraz.net/compatibility/formula-coverage)** &nbsp;·&nbsp; **[API](https://formulon.libraz.net/api/)**

## What's inside

- **Formula engine** — 526 Excel function names, 511 implemented locally; dynamic arrays, `LET` / `LAMBDA`, and `REGEX*`. The rest call cloud or COM services and return a fixed error. [Coverage](https://formulon.libraz.net/compatibility/formula-coverage)
- **Workbook I/O** — read, recalculate and write `.xlsx` and `.xlsb`, including styles, conditional formatting, tables, pivot tables, charts and print layout. [File formats](https://formulon.libraz.net/compatibility/file-format-support)
- **Locale profiles** — 28 Excel behavior profiles across fourteen locales on Mac and Windows hosts. [Locale profiles](https://formulon.libraz.net/compatibility/locale-profiles)
- **Excel oracle** — formula results are compared bit for bit against goldens captured from Mac Excel 365 in fourteen locales; pivot and print layout against Windows Excel 365. [Oracle testing](https://formulon.libraz.net/compatibility/oracle-testing)
- **Size-budgeted WASM** — CI fails the build above 4.5 MiB uncompressed or 1344 KiB Brotli. [Size budgets](https://formulon.libraz.net/development/size-budgets)

## Installation

```bash
npm install @libraz/formulon   # browsers and Node 22+
pip install formulon            # Python 3.9+
```

CLI binaries for `darwin-arm64`, `linux-x64` and `linux-arm64` are on [GitHub Releases](https://github.com/libraz/formulon/releases). Changes between releases are in the [CHANGELOG](CHANGELOG.md); until 1.0 a minor release may break the API and a patch release does not.

## Quick start

```js run
import createFormulon from '@libraz/formulon';

const formulon = await createFormulon();
const value = formulon.evaluateFormula('=SUM(1,2,3)');
console.log(formulon.toNative(value)); // 6
```

```python run
import formulon

print(formulon.evaluate_formula("=SUM(1,2,3)"))  # 6.0
```

```js run
import createFormulon from '@libraz/formulon';

const formulon = await createFormulon();
const wb = formulon.Workbook.createDefault();
try {
  const ws = wb.sheet(0);
  ws.setValue('A1', 0.256);
  ws.setCellFormat('A1', { numberFormat: '0.0%', font: { bold: true } });
  console.log(ws.getDisplayText('A1').text); // 25.6%
} finally {
  wb.dispose();
}
```

```python run
import formulon

with formulon.Workbook.create_default() as wb:
    ws = wb.sheet(0)
    ws["A1"] = 0.256
    ws.set_cell_format("A1", number_format="0.0%", font={"bold": True})
    print(ws.get_display_text("A1").text)  # 25.6%
```

```bash
formulon eval '=SUM(1,2,3)'
formulon recalc input.xlsx -o output.xlsx
```

Failures throw `FormulonError` in JavaScript and `FormulonError` subclasses in Python. Loading, editing and saving workbooks are covered in [Recalculate a workbook](https://formulon.libraz.net/start/recalculate).

## Switching Excel profiles

New workbooks use `win-365-en_US`, the default Excel profile. Formulas are always written with English function names and the stored separators; the Excel profile decides how text is parsed and how results are rendered. `=LENB("日本")` gives `2` on `mac-365-en_US` and `4` on `mac-365-ja_JP`, and `=ISEVEN(2)` shows `WAHR` on `mac-365-de_DE`. Switch it per workbook and recalculate:

```js run
import createFormulon from '@libraz/formulon';

const formulon = await createFormulon();
const wb = formulon.Workbook.createDefault();
try {
  const ws = wb.sheet(0);
  ws.setFormula('A1', '=LENB("日本")');
  wb.setProfileId('mac-365-ja_JP'); // Python: set_profile_id
  wb.recalc(); // a profile change marks formulas dirty
  console.log(ws.getDisplayText('A1').text); // 4
} finally {
  wb.dispose();
}
```

The ids are `{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}`. Every `mac-*` profile and `win-365-ja_JP` is measured against Excel; the other `win-*` profiles are estimated from the Mac measurements. The profile is not saved into the file, so apply it again after loading. `formulon.Workbook.createForProfile(id)` also gives a new workbook the fonts, style names and sheet names Excel uses in that locale. See [Locale profiles](https://formulon.libraz.net/compatibility/locale-profiles) and [Text comparison and sorting](https://formulon.libraz.net/compatibility/text-collation).

## Non-goals

VBA execution, legacy `.xls`, chart rendering, Power Query / DAX, pivot cache refresh from source data, live external connections, and a spreadsheet UI are permanently out of scope. See [Non-goals](https://formulon.libraz.net/compatibility/non-goals).

## Contributing

The most useful contribution is Excel oracle data from your locale, especially from Windows Excel 365: `make oracle-contribute` drives Excel and captures goldens. See [CONTRIBUTING.md](CONTRIBUTING.md) and [Oracle contribution](https://formulon.libraz.net/development/oracle-contribution).

## License

[Apache License 2.0](LICENSE). See also [NOTICE](NOTICE).

## Related projects

- [formulon-cell](https://github.com/libraz/formulon-cell) — browser spreadsheet UI built on `@libraz/formulon`
- [formulon-mcp](https://github.com/libraz/formulon-mcp) — MCP server that gives AI agents workbook tools
