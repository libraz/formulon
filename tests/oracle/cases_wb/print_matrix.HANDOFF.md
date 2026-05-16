# print_matrix oracle capture — Windows-side handoff

**Audience**: the engineer on the Windows host running
`tools/oracle/workbook_oracle_gen.py` (via WSL2 bridge or direct).

**One-line ask**: capture goldens for
`tests/oracle/cases_wb/print_matrix.case.json` (14 cases),
under the strict capture conditions described below, then commit
`tests/oracle/golden_wb/print_matrix.golden.json` back to `develop`.

This document is throw-away. Delete it (and this paragraph) after the
goldens land on `develop`.

---

## Why these 14 cases exist

The prior workbook-track closure attempt ended at **42 pass / 9 fail /
1 skip**, where the 9 residual failures clustered into three groups:

| Group | Count | Suspected cause |
|---|---|---|
| `*_v_breaks` with sparse data | 6 | `VPageBreaks.Count` ignores blank stretches **or** stale `FitToPagesWide=1` from workbook default |
| `print_titles_repeat_rows*` (~50pt off) | 2 | unknown subtraction rule for repeat-rows height |
| `print_fit_scale_25_extreme_shrink` | 1 | COM-side `FitToPages` flag override of Zoom |

Single-shot guarded re-capture would tell us "this configuration
matches now" but would not tell us **why**. The strategy these 14 cases
encode is the empirical alternative: vary one axis at a time so the
golden differences reveal the actual transformation rule. Once the
rules are known, we update `src/print/pagination.cpp` so the rule is
encoded, and any residual cases get `divergence.yaml` entries with
real reasons rather than guesses.

## The four axes

The suite is grouped into four blocks; **each block varies one axis
and pins everything else**. Do not let the Windows-side automation
re-order or merge cases — the diff between cases inside a block is
where the rule lives.

| Block | Axis | Pinned everything else | Variable | Cases |
|---|---|---|---|---|
| A | Zoom % | FitToPages OFF, dense A1:H30, A4 portrait | scale ∈ {100, 75, 50, 25} | `zoom_*_dense_fit_off` |
| B | FitToPages | Zoom=100, dense A1:H30, A4 portrait | (W,H) ∈ {(0,0), (1,0), (0,1), (2,2)} | `fit_*_zoom_100_dense` |
| C | Data density | Zoom=100, FitToPages OFF, A4 portrait, same widths/heights, same `print_area=A1:H30` | populated cells ∈ {dense, sparse_corners, col_sparse_only_a} | `density_*` |
| D | Print-title rows | Zoom=100, FitToPages OFF, A4 portrait, dense A1:D40 | repeat rows ∈ {"1:1", "1:3", "1:5"} | `print_titles_repeat_*` |

## Capture procedure

```bash
# On Windows host (or via WSL2 bridge):
git pull origin develop                                # picks up print_matrix.{yaml,case.json}
python tools/oracle/cli.py workbook --suite print_matrix
# -> writes tests/oracle/golden_wb/print_matrix.golden.json
```

The default generator path (`workbook_oracle_gen.py` → Windows COM
driver) is fine. **Do not** add an ad-hoc capture script — we want the
same driver path used by the other workbook suites so subsequent
verification under `make oracle-verify` shares its code.

### CRITICAL — what the driver must do per case

The driver must apply the `page_setup` block **exactly as authored**,
then read back what Excel actually set, then read the breaks.
Specifically:

1. Set `Worksheet.PageSetup.Zoom = <scale>` (or `False` when `scale`
   is unset).
2. Set `Worksheet.PageSetup.FitToPagesWide = <fit_to_width>` /
   `FitToPagesTall = <fit_to_height>`. A value of 0 in the case JSON
   means "OFF" — apply that as `False` (or `1` with the opposite axis
   `False`, whichever matches Excel's "no fit" state on your version).
3. **Round-trip read**: after applying, read back `.Zoom`,
   `.FitToPagesWide`, `.FitToPagesTall` and **include them in the
   golden's `expect.print` block** (new field
   `applied_page_setup`). This is the single most important addition
   — it converts "Excel disagreed" into "Excel disagreed *and here is
   what Excel said it was running with*". Without this, the prior
   `print_fit_scale_25_extreme_shrink` mystery (Excel returned
   `pages=6` for what should be 25% scaling) cannot be resolved.
4. Then read `HPageBreaks` / `VPageBreaks` and the page count as the
   other suites already do.

If the workbook driver does not currently round-trip
`PageSetup.Zoom/FitToPagesWide/FitToPagesTall`, add that field; it is
load-bearing for this matrix. The schema change is additive (extra
field under `expect.print`) so existing suites stay valid.

### Per-block pre-conditions

To keep the axes clean:

- **Block A (zoom)**: before applying `Zoom=<scale>`, set
  `FitToPagesWide=False, FitToPagesTall=False` explicitly. Excel's
  workbook default is `FitToPages=ON` in some templates — that would
  contaminate the Zoom column otherwise.
- **Block B (fit)**: before applying `FitToPagesWide/Tall`, set
  `Zoom=100`. Confirms whether Fit overrides Zoom or sums with it.
- **Block C (density)**: same `page_setup` across all three cases.
  Only `sheets` (the populated cells) changes. If Excel rounds the
  print area to a populated bounding box, the bounding box differs
  per case; if it honours the explicit `print_area=A1:H30`, breaks
  should be identical. **Either answer is interesting** — the goal is
  to learn which.
- **Block D (print titles)**: same `print_area=A1:D40`. Repeating
  rows should subtract their summed height from the first-page body,
  leaving subsequent pages with full body. Authoring fills enough
  header data so Excel does not collapse the repeat-rows band.

## Expected golden shape

Same as the other print suites, with one added field:

```json
{
  "id": "zoom_50_dense_fit_off",
  "spec":  { /* mirror of the .case.json entry */ },
  "expect": {
    "print": {
      "print_area": "A1:H30",
      "h_breaks":   [15],
      "v_breaks":   [4],
      "pages":      4,
      "applied_page_setup": {
        "zoom": 50,
        "fit_to_width": false,
        "fit_to_height": false
      }
    }
  }
}
```

`applied_page_setup` is the new round-trip-read field described above.

## After commit

Push the golden to `develop` (no PR per the routine-work rule). Then
on the Mac side `make oracle-verify` will pick it up and the matrix
becomes the empirical basis for the next round of `src/print/`
adjustments. Delete this HANDOFF.md in the same commit.
