# formulon

Excel 365 calculation engine, exposed as a pure-Python binding driven by
WebAssembly. Evaluates formulas, loads and saves `.xlsx` workbooks, and
aims for 1-bit compatibility with Mac Excel 365 (ja-JP locale).

## Install

```sh
pip install formulon
```

Requires Python 3.9 or newer. The wheel is `py3-none-any`: it contains a
single `formulon_capi.wasm` (a standalone reactor-style WebAssembly
module exporting the engine's C ABI) plus a thin Python wrapper that
drives it through [`wasmtime`](https://pypi.org/project/wasmtime/). One
wheel works on every platform `wasmtime` supports: Linux x86_64 /
aarch64, macOS x86_64 / arm64, Windows x86_64.

`wasmtime` is the only runtime dependency, declared in the wheel's
metadata; `pip` resolves the right platform-specific `wasmtime` build at
install time.

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
- `formulon.library_version() -> str` -- version of the engine compiled
  into the bundled `formulon_capi.wasm`.
- `formulon.Workbook.create_default() / create_empty() / load(bytes)` --
  factory methods; always use them as context managers (`with ... as wb:`).
- `Workbook.set_number / set_bool / set_text / set_blank / set_formula` --
  cell mutators.
- `Workbook.recalc()` -- triggers a full dependency-ordered recalculation.
  Always serial under WASM (the parallel scheduler requires a pthread
  runtime that wasmtime does not provide; the native CLI uses up to 8
  worker threads).
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

```sh
# From the repository root:
make wasm-capi         # builds build-wasm-capi/formulon_capi.wasm (Emscripten)
make python-package    # stages the wasm into packages/python/formulon/_wasm/
make python-test       # runs the smoke tests against the staged package
make python-wheel      # produces a py3-none-any build-py/dist/formulon-*.whl
```

`packages/python/scripts/stage.py` is the entry point; it just copies
the pre-built `formulon_capi.wasm` into the package data directory --
no compilation happens inside the Python build.

The wheel is intentionally not built with `pip install` from source:
that would require Emscripten on the user's machine. CI builds the wheel
once on Linux and publishes it to PyPI as `py3-none-any`.

## Project

Source, design notes, and the oracle test suite live at
<https://github.com/libraz/formulon>.

## License

Apache License 2.0. See [LICENSE](./LICENSE) and [NOTICE](./NOTICE).
