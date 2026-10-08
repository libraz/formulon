# Oracle Environment

This corpus contains 114 suites captured by the maintainers from Mac Excel 365 with the calculation locale set to th-TH.

- **Excel version**: `16.113.3`
- **Excel locale**: `th-TH` (`AppleLanguages` = `th-TH`, `AppleLocale` = `th_TH`)
- **Capture timestamps (UTC)**: `2026-10-08T16:42:02Z` through `2026-10-08T19:50:13Z`
- **date1904 / iterative**: recorded separately in each golden

`formulon_oracle_th_th_tests` evaluates these goldens with the explicit `mac-365-th_TH` profile. It is registered in CTest with the `oracle` label, independently of the optional variant binary. The ja-JP primary oracle and the `win-365-ja_JP` runtime default remain unchanged.

The target retains `status: scaffolded` in `tools/oracle/targets.yaml`. These maintainer captures use the dedicated oracle gate rather than the contributor provenance opt-in path. Shared skips come from `tests/divergence.yaml`; target-specific overrides belong in this directory's `divergence.yaml`.
