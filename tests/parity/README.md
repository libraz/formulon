# Cross-channel parity gate

This harness verifies that Formulon's three distribution channels --
the native CLI (`build/bin/formulon_cli`), the npm WASM package
(`packages/npm/dist/`), and the Python wheel (`packages/python/`) --
produce **identical** results for the same formula. All three call into
the same C++ engine, so any divergence implicates a binding-layer bug
(number-formatting drift, UTF-8 handling, error-code translation, ...).

The fixture corpus is `fixtures.json`; each entry is one formula plus a
human-readable `expect` hint. The runner does not assert against
`expect` directly -- the comparison is *between* channels, not against a
golden -- but the hint shows up in divergence reports so failures are
self-explanatory.

## How to run

```bash
make parity-test
```

or directly:

```bash
python3 tests/parity/run_parity.py [--verbose]
```

The harness also runs as part of `ctest` under the `PARITY` label:

```bash
cd build && ctest -L PARITY --output-on-failure
```

## What gets compared

For each fixture and each available channel the runner produces a
normalized record:

  * Numbers   -- big-endian IEEE-754 hex, after canonicalizing through
                 `%.15g` (Excel General format / CLI `format_double`).
                 This rounding is *required* because the CLI's only
                 public output is a 15-significant-digit decimal; npm
                 and Python expose the engine's full-precision double,
                 which would otherwise disagree on values like `PI()`.
  * Booleans  -- raw `True` / `False`.
  * Text      -- the UTF-8 string verbatim.
  * Errors    -- the Excel display name (`#DIV/0!`, `#VALUE!`, ...),
                 mapped from the C ABI's integer ordinal.
  * Blanks    -- the bare kind tag.

Channels disagree iff their records differ.

## Skip semantics

The runner *never* requires every channel to be present:

  * Missing `build/bin/formulon_cli`     => CLI channel skipped.
  * Missing `node` on PATH or
    `packages/npm/dist/formulon.js`      => npm channel skipped.
  * `import formulon` fails (and there
    is no `packages/python/` source tree
    fallback)                            => Python channel skipped.

If fewer than 2 channels remain active, the harness prints a warning
and exits 0: a single channel cannot diverge from itself, so the gate
is a no-op rather than a false positive. The harness only fails when
two or more channels were exercised and they disagreed.

This makes the test safe to include in the default `ctest` suite even
on developer machines where only one of the three channels has been
built.

## What to do when a divergence is reported

The output names the baseline channel (the first one that produced a
record) and lists every other channel's normalized record next to it.
The divergent channel -- the one that does not match the others -- is
the bug.

  1. Read the offending channel's binding shim
     (`src/cli/render.cpp`, `src/wasm/bindings.cpp`,
     `packages/python/formulon/workbook.py`).
  2. Reproduce the divergence by hand against the same fixture.
  3. Fix the binding, re-run `make parity-test`, confirm zero
     divergences.

Do *not* "fix" a divergence by changing the fixture corpus. The
fixtures are deliberately small and load-bearing; if a fixture exposes
a real bug, fix the bug.

## Notes on what is *not* tested here

  * Workbook load/save parity (CLI's `recalc` vs npm's `loadBytes` vs
    Python's `Workbook.load`) is a separate harness; this one is
    `eval_formula`-only.
  * Locale-specific formatting (Japanese era dates, ja-JP separators)
    is owned by the oracle harness (`tests/oracle/`).
  * Status-envelope failures (parser crashes that surface as a non-OK
    status, not a `Value.Error`) are out of scope. This harness exits
    early on any host-side failure with a `channel-failure` line in the
    summary; investigating those is a separate workflow.
