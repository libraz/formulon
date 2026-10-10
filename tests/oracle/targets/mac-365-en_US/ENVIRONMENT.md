# Oracle Environment

This corpus contains 123 suites captured by the maintainers from Mac Excel 365 with the calculation locale set to en-US.

- **Excel version**: `16.113.3`, `16.113.4`
- **Excel locale**: `en-US`
- **Capture timestamps (UTC)**: `2026-10-08T05:15:07Z` through `2026-10-10T06:29:32Z`
- **date1904 / iterative**: recorded separately in each golden

`formulon_oracle_en_us_tests` evaluates these goldens with the explicit `mac-365-en_US` profile. It is registered in CTest with the `oracle` label, independently of the optional variant binary. The ja-JP primary oracle and the `win-365-ja_JP` runtime default remain unchanged.

The target retains `status: scaffolded` in `tools/oracle/targets.yaml`. These maintainer captures use the dedicated oracle gate rather than the contributor provenance opt-in path. Shared skips come from `tests/divergence.yaml`; target-specific overrides belong in this directory's `divergence.yaml`.
