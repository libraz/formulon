# formulon

[![PyPI](https://img.shields.io/pypi/v/formulon)](https://pypi.org/project/formulon/)
[![npm](https://img.shields.io/npm/v/@libraz/formulon)](https://www.npmjs.com/package/@libraz/formulon)
[![License](https://img.shields.io/badge/license-Apache--2.0-blue)](https://github.com/libraz/formulon/blob/main/LICENSE)
[![Python](https://img.shields.io/pypi/pyversions/formulon)](https://pypi.org/project/formulon/)
[![Docs](https://img.shields.io/badge/docs-formulon.libraz.net-2563eb)](https://formulon.libraz.net)

**Recalculate Excel workbooks and formulas from Python, with results checked against real Excel 365.**
No Excel install, no Windows, no COM automation: the wheel is pure Python around the engine compiled to WebAssembly, so one `py3-none-any` wheel runs anywhere [`wasmtime`](https://pypi.org/project/wasmtime/) does.

Known differences from Excel are listed case by case in [`tests/divergence.yaml`](https://github.com/libraz/formulon/blob/main/tests/divergence.yaml), each with a reason and the Excel build it was last verified on.

📖 **[Documentation](https://formulon.libraz.net)** &nbsp;·&nbsp; **[Python guide](https://formulon.libraz.net/runtimes/python)** &nbsp;·&nbsp; **[API](https://formulon.libraz.net/api/python)** &nbsp;·&nbsp; **[Batch recalculation](https://formulon.libraz.net/scenarios/python-batch)**

## What's inside

- **Formula engine** — 526 Excel function names, 511 implemented locally; dynamic arrays, `LET` / `LAMBDA`, and `REGEX*`. [Coverage](https://formulon.libraz.net/compatibility/formula-coverage)
- **Workbook I/O** — read, recalculate and write `.xlsx` and `.xlsb`, including styles, conditional formatting, tables, pivot tables, charts and print layout. [File formats](https://formulon.libraz.net/compatibility/file-format-support)
- **Workbook editing** — cells, sheets, rows and columns, styles, merges, comments, validations, pivot tables, print settings and furigana, through typed `Workbook` and `Sheet` APIs. [API](https://formulon.libraz.net/api/python)
- **Locale profiles** — 28 Excel behavior profiles across fourteen locales on Mac and Windows hosts. [Locale profiles](https://formulon.libraz.net/compatibility/locale-profiles)
- **Excel oracle** — formula results are compared bit for bit against goldens captured from Mac Excel 365 in fourteen locales. [Oracle testing](https://formulon.libraz.net/compatibility/oracle-testing)

## Installation

```bash
pip install formulon   # Python 3.9+
```

`wasmtime` is the only runtime dependency; `pip` picks its platform build. Recalculation is serial in this package, with results identical to the threaded surfaces. Changes between releases are in the [CHANGELOG](https://github.com/libraz/formulon/blob/main/CHANGELOG.md); until 1.0 a minor release may break the API and a patch release does not.

## Quick start

```python run
import formulon

print(formulon.evaluate_formula("=SUM(1,2,3)").to_python())  # 6.0

v = formulon.evaluate_formula("=1/0")
print(v.kind.name, v.error_code.name)  # ERROR DIV0
```

Excel errors come back as values (`Value.to_python()` gives an `ExcelError`); a host-side failure raises `FormulonError`. Its subclasses are `IndexOutOfRangeError`, `NotFoundError`, `InvalidArgumentError` and `WorkbookClosedError`. See [Errors](https://formulon.libraz.net/compatibility/errors).

## Recalculate a workbook

```python run
from formulon import Workbook

with Workbook.load("input.xlsx") as wb:
    ws = wb.sheet(0)  # or wb["Sheet1"]
    ws["A1"] = 42
    ws.set_formula("A2", "=A1*2")
    print(wb.has_dirty_cells())  # True
    wb.recalc()
    print(wb.has_dirty_cells())  # False
    print(ws["A2"])  # 84.0

    wb.save("output.xlsx")

    print(ws.name, ws.index)  # Sheet1 0
    ws.set_values("A1", [[1, "x"], [True, None]])
    print(ws["A1"])  # 1.0
```

Every method that takes a cell accepts either `(row, col)` or an A1 string, and a rectangle argument accepts `"A1:B2"`. `ws["A1"]` and `ws[row, col]` read and write native values; `get_values` and `set_values` move a block of cells. Reads return the value as of the last full `recalc()`, and `wb.has_dirty_cells()` tells whether edits are still waiting for one. `ErrorCode.DIV0.display_name` is `"#DIV/0!"`.

A `Sheet` handle stays on the same sheet when sheets are added, moved or removed around it. `wb.sheet(key)` and `wb[key]` take an index or a name and raise `NotFoundError` or `IndexOutOfRangeError` when nothing matches; `wb.list_sheets()` lists every sheet.

`wb.close()` or leaving the `with` block releases the workbook at that moment; `wb.closed` reports it, and any later call raises `WorkbookClosedError`. A workbook dropped without closing is released when the garbage collector reclaims it, which is best effort, so close workbooks you are done with. `Workbook.load` also takes bytes or a binary stream, and `wb.save()` without a destination returns the bytes.

## Format cells

```python run
import formulon

with formulon.Workbook.create_default() as wb:
    ws = wb.sheet(0)
    ws.set_cell_format(
        "A1:B2",
        number_format="0.00%",
        font={"bold": True, "color": "#C00000"},
        fill={"fg": "#FFF2CC"},
        alignment={"horizontal": formulon.HorizontalAlign.CENTER},
    )
    ws.set_cell_format("D:D", border={"left": {"style": formulon.BorderStyle.THIN}})

    fmt = ws.get_cell_format("A1")
    print(fmt.number_format, fmt.font.bold, fmt.fill.pattern.name)  # 0.00% True SOLID
```

`set_cell_format` merges only the groups and keys it is given into each cell's format; a column (`"D:D"`), row (`"3:3"`) or whole-sheet range formats that axis, so cells entered there later take the format too. Equal formats share one style record.

## Switching Excel profiles

New workbooks use `win-365-en_US`, the default Excel profile. Formulas are always written with English function names and the stored separators; the Excel profile decides how text is parsed and how results are rendered. Switch it per workbook and recalculate:

```python run
import formulon

with formulon.Workbook.create_default() as wb:
    ws = wb.sheet(0)
    ws.set_formula("A1", '=LENB("日本")')
    ws.set_formula("A2", '=VALUE("1,5")')
    ws.set_formula("A3", "=ISEVEN(2)")

    for profile in ("mac-365-en_US", "mac-365-ja_JP", "mac-365-de_DE"):
        wb.set_excel_profile_id(profile)
        wb.recalc()
        print(profile, [ws.get_display_text(f"A{row}").text for row in (1, 2, 3)])
# mac-365-en_US ['2', '#VALUE!', 'TRUE']
# mac-365-ja_JP ['4', '#VALUE!', 'TRUE']
# mac-365-de_DE ['2', '1,5', 'WAHR']
```

The ids are `{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}`. Every `mac-*` profile and `win-365-ja_JP` is measured against Excel; the other `win-*` profiles are estimated from the Mac measurements. On a `mac-*` profile, text comparison, sorting, lookups and grouping follow that locale's measured Mac Excel collation; `win-*` profiles add the Windows word-sort rules measured on Windows Excel 365 ja-JP to that locale's collation. The profile is not saved into the file, so store the id your application targets and apply it again after loading.

For a new workbook that matches what Excel creates in a locale, use `Workbook.create_for_profile(id)` instead of `create_default`: it selects the profile and also sets the body and phonetic fonts, the localized Normal style and first sheet names, and the page margins Excel uses there. Sheets added later take the first sheet's row metrics, page margins and phonetic settings.

## Non-goals

VBA execution, legacy `.xls`, chart rendering, Power Query / DAX, pivot cache refresh from source data, live external connections, and a spreadsheet UI are permanently out of scope. See [Non-goals](https://formulon.libraz.net/compatibility/non-goals).

## License

[Apache License 2.0](https://github.com/libraz/formulon/blob/main/LICENSE). See also [NOTICE](https://github.com/libraz/formulon/blob/main/NOTICE).
