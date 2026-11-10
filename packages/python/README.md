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
pip install formulon   # Python 3.10+
```

`wasmtime` is the only runtime dependency; `pip` picks its platform build. Recalculation is serial in this package, with results identical to the threaded surfaces. Changes between releases are in the [CHANGELOG](https://github.com/libraz/formulon/blob/main/CHANGELOG.md); until 1.0 a minor release may break the API and a patch release does not.

## Quick start

```python run
import formulon

print(formulon.evaluate_formula("=SUM(1,2,3)"))  # 6.0

v = formulon.evaluate_formula("=1/0")
print(type(v).__name__, v.code.name)  # ExcelError DIV0
```

Numbers come back as `float`, because every Excel number is a double: `6.0`, not `6`.

Excel errors come back as `ExcelError` values; a host-side failure raises `FormulonError`. Its subclasses are `IndexOutOfRangeError`, `NotFoundError`, `InvalidArgumentError`, `FormulaSyntaxError`, `WorkbookClosedError` and one class per status range (`IoError`, `EvaluationError` and the others). See [Errors](https://formulon.libraz.net/compatibility/errors).

## Recalculate a workbook

```python run
from formulon import Workbook

# Create input.xlsx first; any existing workbook file works the same way.
with Workbook.create_default() as seed:
    seed.save("input.xlsx")

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

Every method that takes a cell accepts either `(row, col)` or an A1 string, and a rectangle argument accepts `"A1:B2"`. `ws["A1"]` and `ws[row, col]` read and write native values; `get_values` and `set_values` move a block of cells. A string that starts with `=` is entered as a formula; enter untrusted text with `ws.set_text("A1", text)`, which always stores text. In the default automatic calculation mode a read calculates pending cells first; in manual mode it returns the value as of the last `recalc()`, and `wb.has_dirty_cells()` tells whether edits are still waiting for one. `ErrorCode.DIV0.display_name` is `"#DIV/0!"`.

A `Sheet` handle stays on the same sheet when sheets are added, moved or removed around it. `wb.sheet(key)` and `wb[key]` take an index or a name and raise `NotFoundError` or `IndexOutOfRangeError` when nothing matches (`wb[-1]` is the last sheet, while a negative index passed to `wb.sheet()` is `IndexOutOfRangeError`); `wb.list_sheets()` lists every sheet.

`wb.close()` or leaving the `with` block releases the workbook at that moment; `wb.closed` reports it, and any later call raises `WorkbookClosedError`. A workbook dropped without closing is released when the garbage collector reclaims it, which is best effort, so close workbooks you are done with. `Workbook.load` also takes bytes or a binary stream, and `wb.save()` without a destination returns the bytes.

## Format cells

```python run
import formulon

with formulon.Workbook.create_default() as wb:
    ws = wb.sheet(0)
    ws.update_cell_format(
        "A1:B2",
        number_format="0.00%",
        font={"bold": True, "color": "#C00000"},
        fill={"foreground": "#FFF2CC"},
        alignment={"horizontal": formulon.HorizontalAlign.CENTER},
    )
    ws.update_cell_format("D:D", border={"left": {"style": formulon.BorderStyle.THIN}})

    fmt = ws.get_cell_format("A1")
    print(fmt.number_format, fmt.font.bold, fmt.fill.pattern.name)  # 0.00% True SOLID
```

`update_cell_format` merges only the groups and keys it is given into each cell's format; a column (`"D:D"`), row (`"3:3"`) or whole-sheet range formats that axis, so cells entered there later take the format too. Equal formats share one style record.

## Threads and memory

- One WebAssembly instance serves the whole process, and calls into it are serialised, so threads do not run workbook code in parallel.
- Memory the instance grew is not given back while the process lives.
- The instance has a 32-bit address space.
- A trap makes every open workbook unusable; start a new process.
- For parallel work, use `multiprocessing` and give each process its own workbooks.

New workbooks use `win-365-en_US`; `wb.set_profile_id(id)` takes any of `{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}`. Switching profiles is covered in the [root README](https://github.com/libraz/formulon#switching-excel-profiles).

## Non-goals

VBA execution, legacy `.xls`, chart rendering, Power Query / DAX, live external connections, and a spreadsheet UI are permanently out of scope. See [Non-goals](https://formulon.libraz.net/compatibility/non-goals).

## License

[Apache License 2.0](https://github.com/libraz/formulon/blob/main/LICENSE). See also [NOTICE](https://github.com/libraz/formulon/blob/main/NOTICE).
