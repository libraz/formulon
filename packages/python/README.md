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
- **Workbook I/O** — read, recalculate and write `.xlsx` and `.xlsb`, including styles, conditional formatting, tables, pivot tables and print layout. [File formats](https://formulon.libraz.net/compatibility/file-format-support)
- **Workbook editing** — cells, sheets, rows and columns, styles, merges, comments, validations, pivot tables, print settings and furigana, through a typed `Workbook` API. [API](https://formulon.libraz.net/api/python)
- **Locale profiles** — 14 Excel behavior profiles across seven locales on Mac and Windows hosts. [Locale profiles](https://formulon.libraz.net/compatibility/locale-profiles)
- **Excel oracle** — formula results are compared bit for bit against goldens captured from Mac Excel 365 in seven locales. [Oracle testing](https://formulon.libraz.net/compatibility/oracle-testing)

## Installation

```bash
pip install formulon   # Python 3.9+
```

`wasmtime` is the only runtime dependency; `pip` picks its platform build. Recalculation is serial in this package, with results identical to the threaded surfaces.

## Quick start

```python
import formulon

print(formulon.eval_formula("=SUM(1,2,3)").to_python())  # 6.0

v = formulon.eval_formula("=1/0")
print(v.kind.name, v.error_code)  # ERROR 1
```

Excel errors come back as values; `FormulonError` is raised only for host-side failures. See [Errors](https://formulon.libraz.net/compatibility/errors).

## Recalculate a workbook

```python
from formulon import Workbook

with open("input.xlsx", "rb") as f:
    data = f.read()

with Workbook.load(data) as wb:
    wb.set_number(0, 0, 0, 42.0)  # Sheet1!A1
    wb.set_formula(0, 1, 0, "=A1*2")  # Sheet1!A2
    wb.recalc()
    print(wb.get_value(0, 1, 0).to_python())  # 84.0

    with open("output.xlsx", "wb") as f:
        f.write(wb.save())
```

## Switching locale profiles

New workbooks use `win-365-en_US`. Formulas are always written with English function names and the stored separators; the profile decides how text is parsed and how results are rendered. Switch it per workbook and recalculate:

```python
with formulon.Workbook.create_default() as wb:
    wb.set_formula(0, 0, 0, '=LENB("日本")')
    wb.set_formula(0, 1, 0, '=VALUE("1,5")')
    wb.set_formula(0, 2, 0, "=ISEVEN(2)")

    for profile in ("mac-365-en_US", "mac-365-ja_JP", "mac-365-de_DE"):
        wb.set_excel_profile_id(profile)
        wb.recalc()
        print(profile, [wb.get_display_text(0, row, 0)[0] for row in range(3)])
# mac-365-en_US ['2', '#VALUE!', 'TRUE']
# mac-365-ja_JP ['4', '#VALUE!', 'TRUE']
# mac-365-de_DE ['2', '1,5', 'WAHR']
```

The ids are `{mac,win}-365-{ja_JP,en_US,de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,pt_BR,ru_RU,zh_TW,it_IT,nl_NL}`. Every `mac-*` profile and `win-365-ja_JP` is measured against Excel; the other `win-*` profiles are estimated from the Mac measurements. The profile is not saved into the file, so store the id your application targets and apply it again after loading.

## Non-goals

VBA execution, legacy `.xls`, chart rendering, Power Query / DAX, pivot cache refresh from source data, live external connections, and a spreadsheet UI are permanently out of scope. See [Non-goals](https://formulon.libraz.net/compatibility/non-goals).

## License

[Apache License 2.0](https://github.com/libraz/formulon/blob/main/LICENSE). See also [NOTICE](https://github.com/libraz/formulon/blob/main/NOTICE).
