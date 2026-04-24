# IronCalc fixtures

This directory vendors a subset of the
[IronCalc](https://github.com/ironcalc/IronCalc) project's test xlsx
fixtures so the Formulon oracle pipeline can evaluate them as a
secondary golden source against Formulon's evaluator.

## Source

- Upstream: https://github.com/ironcalc/IronCalc
- Commit:   `5c9f145b4a8d2b23972b5c4bea3f27d4b0604652`
- Path in upstream: `xlsx/tests/{calc_tests,statistical,docs,templates,calc_test_no_export}/*.xlsx`

## License

IronCalc is dual-licensed under either:

- MIT License (see `LICENSE-MIT`)
- Apache License 2.0 (see `LICENSE-Apache-2.0`)

at the user's option. Copyright (c) 2023 EqualTo GmbH, 2023 Nicolas Hatcher.

The fixtures are retained unmodified under `fixtures/`. Per IronCalc's
dual-license terms, downstream users of Formulon may redistribute these
files under either license. Formulon itself is distributed under
Apache-2.0; see the repository root `LICENSE` and `NOTICE` files for the
combined notice.

## Layout

```
fixtures/
  calc_tests/           # 111 function-correctness xlsx cases
  statistical/          # 28 statistical functions
  docs/                 # 14 documentation-example xlsx
  templates/            # 2 larger workbook templates
  calc_test_no_export/  # 1 xlsx that is not round-tripped
```

Root-level xlsx files from upstream (`basic_text.xlsx`,
`dynamic_arrays.xlsx`, `freeze.xlsx`, `NoGrid.xlsx`,
`openpyxl_example.xlsx`, `split.xlsx`, `example.xlsx`) are intentionally
omitted because they exercise UI / view-layer features (freeze panes,
grid visibility, dynamic-array spill) that the Formulon oracle does not
currently target.

## Regeneration

These fixtures are a point-in-time snapshot. To update:

1. Clone IronCalc at the new commit.
2. Copy the four subdirectories listed above over the ones here,
   preserving relative paths.
3. Update the commit SHA above and re-run
   `make ironcalc-import && make ironcalc-verify` to refresh goldens.

## Running

```
make ironcalc-import      # regenerate golden JSON from xlsx (reads cached <v>)
make ironcalc-verify      # build formulon_ironcalc_oracle_tests and run ctest -L ironcalc
```

See `tools/oracle/ironcalc_import.py` for import details and
`tests/oracle/CMakeLists.txt` for the CTest wiring.
