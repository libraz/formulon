# tools/oracle

Generates Formulon's **oracle golden JSON** by driving Excel under controlled
options. The primary oracle is **Mac Excel 365 (ja-JP)**; additional Excel
environments (Windows, other locales) are supported as opt-in **variants**.

- **Inputs**: `tests/oracle/cases/*.yaml` (shared across all targets)
- **Primary outputs**: `tests/oracle/golden/*.golden.json` + `tests/oracle/ENVIRONMENT.md`
- **Variant outputs**: `tests/oracle/variants/<target>/golden/*.golden.json` + per-variant `ENVIRONMENT.md`

The golden JSON is what the C++ oracle test target consumes. CI never starts
Excel — generation happens on a developer machine and the resulting JSON is
committed.

## Targets

`targets.yaml` enumerates available environments. Inspect from any host:

```bash
tools/oracle/.venv/bin/python tools/oracle/cli.py list
```

Output (example, on macOS):

```
host: Darwin
primary: mac-365-ja_JP
targets:
  * mac-365-ja_JP  driver=macos_excel    runs_on=['Darwin']
    win-365-ja_JP  driver=windows_excel  runs_on=['Windows', 'Linux']
```

`*` marks the primary target. `runs_on` is the list of host OS values
(`platform.system()`) under which the target can be generated. `Linux` here
means **WSL2** — plain Linux is rejected by the bridge driver via
`/proc/version`.

## Setup

`make oracle-setup` auto-detects the host and routes to the right path:

| Host | Routes to | Installs |
|------|-----------|----------|
| macOS (Darwin) | `oracle-setup-mac` | rye → tools/oracle/.venv (xlwings) |
| WSL2 (Linux + Microsoft kernel) | `oracle-setup-wsl` | rye → tools/oracle/.venv + Windows-side instructions |

After setup, verify the host is fully ready:

```bash
tools/oracle/.venv/bin/python tools/oracle/cli.py setup
```

This runs preflight checks per target (xlwings import, Excel reachability,
WSL bridge wiring) and reports PASS / FAIL / SKIP with copy-pasteable hints.

### Mac primary (required for all PRs)

```bash
make oracle-setup
# Grant Automation permission when macOS prompts:
#   System Settings -> Privacy & Security -> Automation
#     -> (your terminal) -> Microsoft Excel
```

### WSL2 → Windows variant (opt-in, for divergence research)

Prerequisites:
- Windows host with Office 365 (ja-JP locale) installed and signed in.
- WSL2 (Ubuntu / Debian / etc.) with `python.exe` reachable from `$PATH`.

One-time on Windows (PowerShell):

```powershell
winget install Python.Python.3.12
py -m pip install xlwings pywin32 pyyaml
```

One-time on WSL2:

```bash
make oracle-setup            # auto-detects WSL2; runs rye sync
# Edit tools/oracle/targets.yaml -> set win_python under win-365-ja_JP, e.g.
#   win_python: "/mnt/c/Users/<you>/AppData/Local/Programs/Python/Python312/python.exe"
make oracle-setup            # re-run; all preflight checks should now PASS
```

## Generate goldens

```bash
# Primary (Mac), all suites
make oracle-gen

# Primary, single suite
make oracle-gen SUITE=count

# Variant by target name (works on any compatible host)
make oracle-gen TARGET=win-365-ja_JP
make oracle-gen TARGET=win-365-ja_JP SUITE=count

# Direct CLI access (more flexibility)
tools/oracle/.venv/bin/python tools/oracle/cli.py gen --target win-365-ja_JP
tools/oracle/.venv/bin/python tools/oracle/cli.py gen --all     # all targets compatible with this host
tools/oracle/.venv/bin/python tools/oracle/oracle_gen.py --suite count --visible   # debug, shows Excel UI
```

Successful runs:
- Primary: updates `tests/oracle/golden/*.golden.json` + `tests/oracle/ENVIRONMENT.md`
- Variant: updates `tests/oracle/variants/<target>/golden/*.golden.json` + per-variant `ENVIRONMENT.md`

Diffs on these files are part of the PR review surface. Every oracle-gen PR
should call out which cells changed and why.

## Verify (any platform)

```bash
make oracle-verify           # primary only; runs ctest -L oracle
```

Variants are verified through a **separate** test binary, gated behind a
CMake option:

```bash
cmake -B build -DFORMULON_ORACLE_VARIANTS=ON
cmake --build build --target formulon_oracle_variant_tests
cd build && ctest -L VARIANT --output-on-failure
```

Variant test names get a `__<target>` suffix so primary and variant
parameters never collide. Default builds (no `-DFORMULON_ORACLE_VARIANTS=ON`)
ignore `tests/oracle/variants/` entirely.

## Adding cases

1. Pick or create `tests/oracle/cases/<category>.yaml`.
2. Append a case record:
   ```yaml
   - id: sum_mixed_range
     formula: "=SUM(A1:A3)"
     setup: { A1: 1, A2: "text", A3: 3 }
   ```
3. Regenerate the primary golden:
   ```bash
   make oracle-gen SUITE=<category>
   ```
4. (Optional) Regenerate variant goldens on the appropriate host:
   ```bash
   make oracle-gen TARGET=win-365-ja_JP SUITE=<category>
   ```
5. Commit the YAML + every refreshed `*.golden.json`.

## Divergences

Cases that Excel evaluates differently from Formulon on purpose (locale,
1-ulp drift, NOW-style volatiles) are declared in `tests/divergence.yaml`.

Per-entry options:

```yaml
entries:
  - id: my_case_id
    mode: skip-oracle               # excludes from generation entirely
    reason: "..."
    prefer: formulon                # which side we trust if they differ
    first_noted: 2026-04-23
    applies_to: [mac-365-ja_JP]     # OPTIONAL; default = applies to all targets
```

Per-variant overrides live in `tests/oracle/variants/<target>/divergence.yaml`
and get merged on top of the primary file (variant entries win on key
collision).

## Architecture

```
tools/oracle/
├── targets.yaml                 declarative target manifest
├── cli.py                       cross-platform dispatcher (list / setup / gen)
├── oracle_gen.py                core generator (loads cases, calls driver, writes JSON)
├── case_schema.py               YAML loader + validator
├── driver.py                    backward-compat re-export shim
└── drivers/
    ├── base.py                  OracleDriver ABC + CaseResult / EnvironmentInfo
    ├── macos_excel.py           Mac driver (xlwings via AppleEvents)
    ├── windows_excel.py         Windows driver (xlwings via COM) + wire entrypoint
    ├── wsl_bridge.py            WSL2 wrapper that subprocess-invokes windows_excel
    └── __init__.py              select_driver(target) factory
```

The driver factory chooses the concrete class based on `target['driver']`
and the current host:

| target.driver | Darwin | Windows | WSL2 | Plain Linux |
|---|---|---|---|---|
| `macos_excel` | ✅ direct | ❌ | ❌ | ❌ |
| `windows_excel` | ❌ | ✅ direct | ✅ via wsl_bridge | ❌ |
| `wsl_bridge` (explicit) | ❌ | ❌ | ✅ direct | ❌ |

The WSL2 bridge ferries one JSON command file and one JSON result file
between WSL Python and Windows Python via `wslpath -w`. Wire format:

```json
// input.json
{"version": 1, "command": "run_suite", "suite_name": "...", "cases": [...]}
// output.json
{"version": 1, "environment": {...}, "results": [{"id": ..., "kind": ..., "value": ...}]}
```

## Troubleshooting

| Symptom | Likely cause | Fix |
|---|---|---|
| `oracle-setup-mac: rye not found` | rye not installed | `brew install rye` or `curl -sSf https://rye.astral.sh/get \| bash` |
| `target driver 'windows_excel' needs Windows or WSL2 host, got Darwin` | trying to gen variant from wrong host | Run on the right host, or use `cli.py list` to see what's compatible |
| `WSL bridge requires `win_python` in targets.yaml` | targets.yaml not configured | Edit `tools/oracle/targets.yaml`, set `win_python` under the variant target |
| `windows_excel subprocess failed (rc=...)` | Win Python missing xlwings/pywin32, or Excel not activated | Re-run `make oracle-setup-wsl`, then `cli.py setup --target win-365-ja_JP` |
| Variant tests not appearing in ctest | CMake variant flag off | `cmake -B build -DFORMULON_ORACLE_VARIANTS=ON` |
