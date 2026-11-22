# Contributing to Formulon

Contributions are welcome, including small and unfinished ones. A maintainer reviews every pull request, finishes what is missing, and merges it, so you do not need to get everything in this document right, or even read all of it.

Donating Excel measurements is a contribution that needs no code: one command captures results from the Excel you already run. See [Donating oracle data](#donating-oracle-data).

## What a pull request needs

- A description of what the change does and why. If the intent is clear, the maintainer can verify the behaviour.
- If you can, a build and the tests for the part you changed.

That is all. A red check, a missing test, an untouched binding or a missing type declaration is fine: say what you know in the pull request, and the maintainer will finish it or tell you what is left. Target `develop` if you can; a pull request against `main` is retargeted.

For a larger change, such as a new feature family or a change to the public API, opening an issue first to describe the idea can save you work. It is not required.

Reporting a bug you do not intend to fix is a useful contribution too, and needs none of the setup below.

### What the maintainer can verify

The maintainer measures Excel on macOS in every locale the project ships, and on Windows in Japanese only. If a change or a bug depends on Windows in another language (a locale-specific result, a regional date format, a system font), describe what you see there, ideally with what Excel shows for the same input, because it cannot be reproduced on the maintainer's side.

## What you are contributing to

A C++17 calculation engine for Excel 365 with OOXML and MS-XLSB I/O, shipped as:

- WebAssembly on npm (`@libraz/formulon`, with a `/threads` build), in `src/wasm` and `packages/npm`.
- A Python wheel (`formulon`), in `packages/python`.
- A native Node addon (`@libraz/formulon-native`) and a CLI.

The goal is results identical to Excel 365 within a fixed WebAssembly size ceiling. Formula and workbook behaviour is judged against goldens captured from real Excel (the oracle) under `tests/oracle/`. Divergences from Excel are recorded one by one in `tests/divergence.yaml`, each with a reason.

A spreadsheet UI problem belongs in the [formulon-cell repository](https://github.com/libraz/formulon-cell/issues). A problem with an MCP tool belongs in the [formulon-mcp repository](https://github.com/libraz/formulon-mcp/issues).

The rest of this document has two parts. [Checklist](#checklist) describes what a finished change looks like. [Reference](#reference) explains how the checks and the size ceiling behave.

## Development setup

CMake, a C++17 compiler and Ninja or Make. `make` targets cover the routine work:

```bash
make build            # Debug build in build/
make test             # fast tier: everything except SLOW, BENCH and TSAN
make test-slow        # adds the SLOW tier
make format           # clang-format, ruff and biome; make format-check only checks
make lint             # ruff check; clang-tidy is advisory
```

The npm, Python and addon packages have their own targets. Run the ones for the binding you touched:

```bash
make npm-test         # builds the WASM package and runs its tests
make node-test        # native addon
make python-test      # Python wheel (needs uv)
python3 tools/dev/check_binding_drift.py all
```

The WebAssembly build needs Emscripten at the version pinned in `tools/wasm/emsdk-version.txt`:

```bash
make wasm             # single-threaded build in build-wasm/
make wasm-threads     # pthread build in build-wasm-threads/
make size-check       # compares both .wasm files with the size ceilings
```

## What CI requires

A red check on your pull request is not a reason to hold it back; the maintainer can look at it with you.

| Job | Workflow | Runs on |
| --- | --- | --- |
| format check (`make format-check`, `make lint`) and the README and binding drift checks | `develop-ci.yml` | pushes to `develop` and pull requests to it |
| build and fast test, Linux, Debug | `develop-ci.yml` | the same events |
| format check and drift checks | `ci.yml` | pushes to `main` and pull requests to it |
| WASM build and size check | `ci.yml` | the same events |
| build, test and coverage, Linux | `ci.yml` | pull requests to `main`, weekly, and on demand |
| prebuild of the native addon (darwin-arm64, linux-x64, linux-arm64) | `prebuild.yml` | pushes to `main`, pull requests to `main`, `v*` tags |
| Python binding smoke test | `prebuild.yml` | the same events |
| ThreadSanitizer | `sanitizers.yml` | started by hand |
| npm, PyPI and CLI publish | `release.yml` | `v*.*.*` tags |

Every workflow can also be started by hand from the Actions tab.

## Checklist

This is what a change looks like by the time it is merged. The maintainer checks it and fills in what is missing, so you are not expected to know all of it.

- Library code reports failure through `Expected<T, Error>`, with no exceptions and no RTTI; both are compiled out on native and WebAssembly builds.
- Library code stays WebAssembly-safe: no `<filesystem>`, no `<iostream>`, and no third-party header that relies on exceptions.
- A new Excel built-in goes in the family file under `src/eval/builtins/` and is registered at the bottom of that file's `register_<family>_builtins()`. A function that short-circuits or preserves range shape is lazy and takes an entry in `src/eval/tree_walker_lazy_table.cpp`.
- A new built-in has unit tests under `tests/unit/eval/` and, where Excel's behaviour is observable, cases under `tests/oracle/cases/`.
- A change to `src/c_api/formulon_c.h` passes `make node-test` and `make python-test`.
- A new or changed public name lands in every binding (C ABI, WASM/JS, native addon, Python, CLI), in `tools/api/names.yaml` and in `CHANGELOG.md` together, and `python3 tools/dev/check_binding_drift.py all` passes.
- A change that grows the `.wasm` by more than a few kilobytes states the size effect in the pull request description; the build stays under the ceilings in [Reference](#reference).
- A new divergence from Excel is an entry in `tests/divergence.yaml` with its reason and the Excel version it was last verified on.
- Where a change fixes a defect, include a test that fails without the fix.
- Write comments in English and document exported symbols.

### Commit messages

Conventional Commits, in English: `type(scope): summary`, lowercase after the colon, no trailing period. A body, when there is one, is a blank line followed by `- ` bullets. Commit messages can be rewritten at merge, so do not let them hold up a pull request.

## Reference

Nothing in this section asks anything of you; read it when a checklist item or a check puzzles you.

- Exceptions and RTTI are disabled on every target, so there is no exception path to fall back to; a fallible function returns `Expected<T, Error>`. Error codes are ranged by module in `src/utils/error.h`.
- The WebAssembly size gate fails above 6 MiB uncompressed or 1920 KiB Brotli. Brotli size is the binding constraint, and `make size-check` reports both.
- Eager built-ins are plain function pointers with no evaluation context. A function that needs workbook state, such as the date system, is lazy.
- `make build` configures a Debug tree, where `assert` runs; a Release build compiles asserts out, so a test that passes in one can abort in the other. CI on `develop` builds Debug.
- `make test` builds nothing; run `make build` first.
- The test labels are `SLOW`, `TSAN`, `BENCH` and `PARITY`; the unlabeled fast tier is what CI gates on. A new test binary registers with `gtest_discover_tests(... DISCOVERY_TIMEOUT 120)`, because listing the coverage-instrumented binary overruns the default 5 seconds.
- The C ABI is exercised by the Node and Python test suites, not by `ctest`, which is why a header change needs `make node-test` and `make python-test`.
- The release gate is an oracle pass rate of 99.5% or more of the executed cases with zero failures, and `python3 tools/oracle/divergence_check.py --strict` exiting 0.
- A pull request to `develop` starts `develop-ci.yml` only. The WASM size check runs from `ci.yml` on `main`, so the maintainer runs `make size-check` before merging a change that could grow the binary.

## Donating oracle data

The 1-bit compatibility claim rests on goldens captured from real Excel. If you run Excel 365 in a locale the project has not measured, one command captures a golden snapshot and prints the `git` commands to submit it. No code is involved.

### Requirements

Oracle data must come from Excel 365 (a Microsoft 365 subscription). Office 2019 and earlier return `#NAME?` for `ARRAYTOTEXT`, `LAMBDA` and the dynamic-array functions (`SORT`, `FILTER`, `UNIQUE`, `XLOOKUP`), so a golden captured there records `#NAME?` for each of them. Pull requests with goldens from such installs are not merged.

The three generators (`oracle-gen`, `oracle-gen-cf`, `oracle-gen-workbook`) evaluate `=ARRAYTOTEXT(1)` at startup and abort if Excel does not recognise it.

### Why this helps

Each track has one primary oracle: Mac Excel 365 ja-JP for formula and conditional-formatting results, and Windows Excel 365 ja-JP for pivot tables and print layout, which need Windows COM automation. The maintainer captures all fourteen supported Mac locales, and their formula goldens gate CI. Windows profiles other than ja-JP are estimated from the Mac capture of the same locale plus the known host differences.

Some results differ across locales and across the Mac and Windows ports:

- Locale-bound text functions: `BAHTTEXT`, `NUMBERSTRING`, `TEXT` with localized format strings, `DATEVALUE` on locale date strings, weekday and month names.
- System-locale formatting: decimal separator, date order, currency symbols, list separator.
- Mac and Windows divergences: date arithmetic near the 1900-02-29 phantom leap day, some `[$-...]` locale-tag fallbacks, a few text-encoding and rounding edge cases.

A single install cannot reproduce the Windows side. Windows Excel takes its locale from the OS language, the Office language pack and the registry, and the maintainer has one Windows installation (Japanese). A golden captured by someone who uses Excel in that locale is also easier to trust, because that person notices when results look wrong.

The runtime profile switch selects one of 28 profiles across fourteen locales, named `{mac,win}-365-<locale>`; the default is `win-365-en_US`. Only `win-365-ja_JP` is measured on Windows. Goldens from a real Windows install in another locale replace an estimate with a measurement.

### What is collected

A target name has the shape `<host>-<excel-major>-<locale>`, with `<host>` one of `mac` or `win` and `<locale>` in BCP 47 with `_` instead of `-`. [`tools/oracle/targets.yaml`](tools/oracle/targets.yaml) lists every target with a status:

| status | meaning |
|---|---|
| `primary` | the 1-bit reference, maintained by the maintainer |
| `scaffolded` | goldens exist. The thirteen non-primary Mac formula targets gate CI; `win-365-ja_JP` covers the workbook track |
| `wanted` | reserved slot with no goldens yet; these are the ones to donate |

`make oracle-contribute-list` prints the current set with the subset your host can run. The `wanted` slots are the thirteen Windows targets other than `win-365-ja_JP`. If your locale is not listed, see [Adding a new target](#adding-a-new-target).

### Prerequisites

| Host | Excel | Tooling |
|---|---|---|
| macOS 12+ | Excel 365, any locale | Python 3.10+; `make oracle-contribute` creates a virtual environment with `xlwings`. macOS asks for Automation permission on first use: System Settings, Privacy & Security, Automation, your terminal, Microsoft Excel. |
| Windows 10/11 | Excel 365, any locale | Python 3.10+ (`winget install Python.Python.3.12`) with `xlwings pywin32 pyyaml` installed in it. |
| WSL2 on Windows | Excel 365 on the Windows side | As on Windows. WSL2 Python only orchestrates and carries JSON. Set `FORMULON_WIN_PYTHON` to the Windows-side `python.exe`, for example `/mnt/c/Users/<you>/AppData/Local/Programs/Python/Python312/python.exe`, so the path stays out of `targets.yaml`. |

### Run it

From a fresh checkout:

```bash
make oracle-contribute
```

The command:

1. Creates the Python environment with `xlwings` on first run.
2. Reads the installed Excel version and locale.
3. Maps them to a target in `targets.yaml`. For an untracked target it prints how to reserve the slot.
4. Compares versions. If the committed `ENVIRONMENT.md` for that target came from an Excel build at or above yours, it thanks you and exits; no pull request is needed.
5. Otherwise it runs preflight, generates the goldens (15 to 25 minutes for the full suite), records `ENVIRONMENT.md`, and prints the `git` commands and the pull request URL.

If auto-detection picks the wrong target, for example with several Excel languages installed, name it:

```bash
make oracle-contribute TARGET=win-365-de_DE
```

The locale that counts is Excel's, not the OS's. On Mac Excel, check Excel, Preferences, General, Display Language; it must match System Settings, Language & Region. On Windows Excel, check File, Options, Language, Office display language. If they disagree, change one and restart Excel before running.

### Capturing another locale on a Mac

Mac Excel reads two settings from its own defaults domain. `AppleLanguages` moves the `CHAR` and `CODE` code page, `LENB`, `ASC`, era and day names and `NUMBERSTRING`. `AppleLocale` moves currency, date order and `TEXT` format keywords. Set both and relaunch Excel:

```bash
osascript -e 'tell application "Microsoft Excel" to quit'
defaults write com.microsoft.Excel AppleLanguages -array de-DE
defaults write com.microsoft.Excel AppleLocale de_DE
open -a "Microsoft Excel"
make oracle-gen TARGET=mac-365-de_DE
```

`tools/oracle/oracle_gen.py --target <target> --incremental` refreshes an existing target. The generator records the locale it reads back and warns when it differs from the target's. Remove both keys (`defaults delete com.microsoft.Excel AppleLanguages`, likewise `AppleLocale`) to return Excel to the system locale.

### Submit the result

When the command finishes it prints something like:

```
Done. Your oracle data is on disk -- thank you for the donation!

  Target:   win-365-de_DE
  Locale:   de-DE
  Host:     Windows
  Env file: tests/oracle/targets/win-365-de_DE/ENVIRONMENT.md

  Files git considers new or changed:
    tests/oracle/targets/win-365-de_DE/golden/...
    tests/oracle/targets/win-365-de_DE/ENVIRONMENT.md
    ...

  Next: push your branch and open a PR.

    git checkout -b oracle/win-365-de_DE-2026-04-28
    git add tests/oracle/targets/win-365-de_DE/
    git add tests/oracle/targets/win-365-de_DE/ENVIRONMENT.md
    git add tools/oracle/targets.yaml    # bump status: wanted -> scaffolded
    git commit -m 'test(oracle): contribute win-365-de_DE goldens'
    git push -u origin oracle/win-365-de_DE-2026-04-28

  Then open a PR (oracle template auto-loads):
    https://github.com/libraz/formulon/compare/develop...oracle/win-365-de_DE-2026-04-28?expand=1&template=oracle.md
```

Run those commands as printed; a pull request opened against `main` is retargeted. The oracle pull request template loads through that URL. The description should include:

- The Excel version, pasted verbatim from Excel, About.
- The OS version (`sw_vers -productVersion` on Mac, `winver` on Windows).
- Confirmation that the run used real Excel, not a modified driver or a debugger that could intercept results.
- Anything you noticed during generation: warnings, or cases that opened Excel dialogs.

### Two runs for a first-time target

When a target moves from `wanted` to its first goldens, a second contributor with the same target regenerates independently and posts a `git diff` against the first pull request. Identical output from two independent runs is what justifies `scaffolded`. To verify another contributor's pull request:

```bash
git fetch origin pull/<PR-NUMBER>/head:contrib-pr
git checkout contrib-pr
make oracle-contribute TARGET=<target-name>
git diff
```

A diff limited to `generated_at` timestamps and explicit divergences means the target is accepted.

### Refreshing an existing target

When your Excel is newer than the committed `ENVIRONMENT.md`, `make oracle-contribute` regenerates on its own. The pull request is small: `ENVIRONMENT.md` plus the goldens that moved. State which results changed so each can be classified as a bug, ULP drift or a new reading of the spec.

### Adding a new target

If your locale is not in `targets.yaml`, the command prints the derived target name and two ways forward:

```
[contribute] derived target name: mac-365-pt_BR
[contribute] this target isn't tracked in tools/oracle/targets.yaml.

Two ways forward:
  1) Open an issue so we can reserve the slot, then re-run:
     https://github.com/libraz/formulon/issues/new?title=Oracle+target%3A+mac-365-pt_BR
  2) Or pass --target=<existing-name> if you intend to map your
     environment onto an already-tracked target.
```

Open the issue with the derived target name, your Excel version and OS, and any locale-specific functions where you expect divergences (`pt_BR` exercises decimal-comma `TEXT` and Brazilian date formats). The maintainer adds a `wanted` entry to `targets.yaml`; then run `make oracle-contribute` again.

## Reporting issues

Search the existing issues first, then include what you did, what happened, what you expected, the binding (npm/WASM, Python, native addon or CLI), the version (`formulon --version` or the package version), and your OS and its display language. For a wrong result, give the formula, what Excel shows, and the Excel version and locale; a YAML case in the format under `tests/oracle/cases/` can be dropped straight into the suite. For a problem with a loaded workbook, the file matters more than anything else in the report.

Security vulnerabilities go through the process in [SECURITY.md](SECURITY.md), not the public issue tracker.

## License

By contributing, you agree that your contributions are licensed under the [Apache License 2.0](LICENSE).
