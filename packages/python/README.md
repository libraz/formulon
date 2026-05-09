# formulon

Excel 365 calculation engine, exposed as a pure-`ctypes` Python binding.
Evaluates formulas, loads and saves `.xlsx` workbooks, and aims for 1-bit
compatibility with Mac Excel 365 (ja-JP locale).

## Install

```sh
pip install formulon
```

Requires Python 3.9 or newer. PyPI publishes platform wheels only. Each wheel
bundles a precompiled `libformulon.{so,dylib,dll}`; there is no Python
build-time dependency on NumPy, Cython, or pybind11.

Source distributions are intentionally not published for the 0.9 series.
Building the native engine from source requires CMake and a C++17 compiler, so
release artifacts are cut as verified wheels for supported platforms instead.

## Quick start

```python
import formulon

# One-shot formula evaluation against a fresh workbook.
v = formulon.eval_formula("=SUM(1,2,3)")
print(v.to_python())  # 6.0

# Cell-level Excel errors surface as Value(kind=ValueKind.ERROR), not
# Python exceptions.
v = formulon.eval_formula("=1/0")
print(v.kind)         # ValueKind.ERROR
print(v.error_code)   # 1 (ErrorCode::Div0)
```

## Workbook example

```python
from formulon import Workbook

with Workbook.create_default() as wb:
    wb.set_number(0, 0, 0, 21.0)        # Sheet1!A1 = 21
    wb.set_formula(0, 0, 1, "=A1*2")    # Sheet1!B1 = =A1*2
    wb.recalc()

    print(wb.get_value(0, 0, 1).to_python())  # 42.0

    # Serialise back to an in-memory .xlsx.
    blob = wb.save()
    with open("output.xlsx", "wb") as f:
        f.write(blob)

# Load back from disk.
with open("output.xlsx", "rb") as f:
    blob = f.read()

with Workbook.load(blob) as wb:
    wb.recalc()
    for cell in wb.iter_cells(0):
        print(cell)
```

## API reference

The public surface is documented via Python docstrings and the
hand-rolled type stubs in `formulon/__init__.pyi`. Highlights:

- `formulon.eval_formula(formula: str) -> Value` -- one-shot evaluation.
- `formulon.library_version() -> str` -- version of the underlying
  `libformulon` build.
- `formulon.Workbook.create_default() / create_empty() / load(bytes)` --
  factory methods; always use them as context managers (`with ... as wb:`).
- `Workbook.set_number / set_bool / set_text / set_blank / set_formula` --
  cell mutators.
- `Workbook.recalc()` -- triggers a full dependency-ordered recalculation.
- `Workbook.get_value(sheet, row, col) -> Value` -- read a cached value.
- `Workbook.save() -> bytes` -- serialise to `.xlsx`.
- `Workbook.iter_cells(sheet)`, `iter_defined_names()`, `iter_tables()`,
  `iter_passthrough()` -- iteration helpers.

`Value` exposes `kind`, `number`, `boolean`, `text`, `error_code`, plus
`to_python()` which converts to the natural Python type
(`None` / `float` / `bool` / `str`) or returns the `Value` itself for
errors and reserved kinds.

`FormulonError` is raised only for host-side problems (NULL handle,
parser crash inside `Workbook.load`, OOM). Excel cell errors travel
inside `Value(kind=ValueKind.ERROR)`.

## Building from source

The wheel is intentionally not built with `pip install` from source --
that would require CMake and a C++17 compiler on every host. Instead:

```sh
# From the repository root:
make python-package    # builds libformulon and stages it into _lib/
make python-test       # runs the smoke tests against the staged package
make python-wheel      # produces a platform-tagged build-py/dist/formulon-*.whl
```

`packages/python/scripts/stage.py` is the entry point; it shells out to
CMake with `-DFM_BUILD_C_API_SHARED=ON` and copies the resulting library
into `packages/python/formulon/_lib/`.

Do not upload a `py3-none-any` wheel for this package. The binding is pure
Python, but the wheel is platform-specific because it contains the native
shared library. Linux release wheels should be built in a manylinux container
and repaired with `auditwheel`; macOS release wheels should be checked with
`delocate`.

## Project

Source, design notes, and the oracle test suite live at
<https://github.com/libraz/formulon>.

## License

Apache License 2.0. See [LICENSE](./LICENSE) and [NOTICE](./NOTICE).
