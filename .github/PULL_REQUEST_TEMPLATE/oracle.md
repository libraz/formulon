<!--
Thank you for donating Excel oracle data!

This template is loaded automatically by the URL `make oracle-contribute`
prints. If you got here some other way and your PR is *not* an oracle
contribution, switch to the default template by removing
`?template=oracle.md` from the compare URL.
-->

## Target

<!-- The target name as printed by `make oracle-contribute`, e.g.
     mac-365-en_US, win-365-de_DE. -->

`<target-name>`

## Environment

<!-- Paste verbatim from your machine. The Excel version is what
     ENVIRONMENT.md should also record; reviewers cross-check. -->

| Field | Value |
|---|---|
| Excel version | `<from Excel -> About>` |
| Excel locale (display language) | `<e.g. en-US>` |
| Host OS | `<sw_vers -productVersion / winver>` |
| Host arch | `<arm64 / x86_64>` |
| Generation date | `<YYYY-MM-DD>` |

## Generation

- [ ] I ran `make oracle-contribute` (auto-detect path), or
- [ ] I ran `make oracle-contribute TARGET=<name>` (explicit override).
- [ ] Goldens were generated against **real Excel** -- no mocks, no
      modified driver, no debugger intercepting results.
- [ ] I quit Excel before re-running anything; no other workbooks were
      open during generation.

Suites regenerated:

<!-- Either "all suites" or list specific ones. The contribute flow
     regenerates all by default. -->

- [ ] All suites under `tests/oracle/cases/`
- Or, specific suites: `<list>`

## What changed

<!-- Run `git diff --stat` against the variant directory and paste the
     summary. Call out which goldens moved, especially any that aren't
     just `generated_at` timestamp churn. -->

```
<paste git diff --stat>
```

Notable golden moves (if any):

-

## Divergences observed

<!-- Anything Excel did that surprised you, or that the verifier flags
     against the primary (Mac ja-JP) golden. New entries probably
     belong in tests/divergence.yaml or
     tests/oracle/variants/<target>/divergence.yaml -- include the
     proposed entry below if you can. -->

- None / `<list>`

## Two-person verify

This is a **first-time** target (status was `wanted`):

- [ ] Yes -- a second contributor with the same target should run
      `make oracle-contribute TARGET=<name>` independently and post the
      diff before merge. I have **not** acted as my own second verifier.
- [ ] No -- this is a refresh of an already-`scaffolded` target, so the
      single-contributor rule applies.

If this is a **second-verifier PR** (you are independently regenerating
to confirm a prior PR's goldens), link the original PR here:

> Verifying #<PR-NUMBER>
>
> `git diff` against the original PR's branch:
> ```
> <paste relevant lines, or "clean diff (only generated_at moved)">
> ```

## License

- [ ] I agree my contribution is licensed under
      [Apache 2.0](../LICENSE).
