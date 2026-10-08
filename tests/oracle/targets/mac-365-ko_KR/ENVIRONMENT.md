# Oracle Environment

This corpus contains 114 suites captured by the maintainers from Mac Excel 365 with the calculation locale set to ko-KR.

- **Excel version**: `16.113.3`
- **Excel locale**: `ko-KR` (`AppleLanguages` = `ko-KR`, `AppleLocale` = `ko_KR`)
- **Capture timestamps (UTC)**: `2026-10-08T15:13:25Z` through `2026-10-08T19:52:05Z`
- **date1904 / iterative**: recorded separately in each golden

`formulon_oracle_ko_kr_tests` evaluates these goldens with the explicit `mac-365-ko_KR` profile. It is registered in CTest with the `oracle` label, independently of the optional variant binary. The ja-JP primary oracle and the `win-365-ja_JP` runtime default remain unchanged.

The target retains `status: scaffolded` in `tools/oracle/targets.yaml`. These maintainer captures use the dedicated oracle gate rather than the contributor provenance opt-in path. Shared skips come from `tests/divergence.yaml`; target-specific overrides belong in this directory's `divergence.yaml`.
