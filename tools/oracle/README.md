# tools/oracle

Generates Formulon's **oracle golden JSON** by driving Mac Excel 365
through xlwings.

- **Platform**: macOS only (xlwings uses AppleEvents)
- **Excel**: Excel for Mac 365 installed, Automation permission granted
- **Inputs**: `tests/oracle/cases/*.yaml`
- **Outputs**: `tests/oracle/golden/*.golden.json` + `tests/oracle/ENVIRONMENT.md`

The golden JSON is what the C++ oracle test target (`make oracle-verify`)
consumes. CI never starts Excel — generation happens on a developer Mac and
the resulting JSON is committed.

## Setup (one-time)

Python tooling is managed by [rye](https://rye.astral.sh/). Install it first
(`curl -sSf https://rye.astral.sh/get | bash` or `brew install rye`), then:

```bash
make oracle-setup       # runs `rye sync` in tools/oracle/ (creates .venv)
```

This materializes `tools/oracle/.venv/` from `pyproject.toml` +
`requirements.lock`. The lockfile is committed so every developer Mac
resolves the same xlwings / appscript versions.

The first xlwings run will prompt for Automation permission
(Terminal / iTerm → System Settings → Privacy → Automation → Excel).

## Regenerate goldens

```bash
make oracle-gen                      # all suites (slow)
make oracle-gen SUITE=count          # one suite by name
python3 tools/oracle/oracle_gen.py --suite count --visible   # debug
```

Successful runs update `tests/oracle/golden/*.golden.json` and
`tests/oracle/ENVIRONMENT.md` (Excel version / locale snapshot). Diffs on
these files are part of the PR review surface — every oracle-gen PR should
call out which cells changed and why.

## Verify (any platform)

```bash
make oracle-verify      # runs ctest -L oracle
```

## Adding cases

1. Pick or create `tests/oracle/cases/<category>.yaml`.
2. Append a case record:
   ```yaml
   - id: sum_mixed_range
     formula: "=SUM(A1:A3)"
     setup: { A1: 1, A2: "text", A3: 3 }
   ```
3. Regenerate: `python3 tools/oracle/oracle_gen.py --suite <category>`
4. Commit both the YAML and the refreshed `*.golden.json`.

## Divergence

Cases that Excel evaluates differently from Formulon on purpose (locale,
1-ulp drift, NOW-style volatiles) are declared in `tests/divergence.yaml`.
Entries with `mode: skip-oracle` are excluded from generation; entries with
`tolerance: {abs, rel}` are generated but the verifier widens its compare
threshold accordingly.
