# print_matrix follow-up — Windows-side handoff (round 2)

**Audience**: same Windows engineer who captured `print_matrix.golden.json`
in commit `2780199`.

**One-line ask**: extend the workbook driver to round-trip-read two
additional things (page margins + a more reliable Pages.Count), then
re-capture three suites whose current goldens are blocking further work.

This document is throw-away. Delete it (and this paragraph) once the
re-capture lands on `develop`.

---

## What the first matrix capture revealed

The 14-case `print_matrix` round 1 capture took the workbook oracle
from **20 fail / 46 pass / 1 skip** down to **8 fail / 57 pass / 1
skip** after `d5dd461 fix(print): Suppress auto-column breaks ...`.

Empirical rules derived from your goldens:

- **Block C (density)**: confirmed Excel never auto-paginates columns at
  scale=100. A wide print area renders on a single page-column,
  clipped at the right margin; `VPageBreaks` and `Pages.Count` both
  ignore the overflow. `src/print/pagination.cpp` now matches.
- **Block B (fit)**: confirmed `FitToPages` correctly forces
  `Zoom=False` and produces predictable page counts.

## The 8 residual failures break into 3 groups

### Group 1: driver-artifact (3 cases) — declared in divergence.yaml

| Case | Symptom |
|---|---|
| `print_fit.scale_25_extreme_shrink` | `pages=6` with `h_breaks=[], v_breaks=[]` (formula `(H+1)*(V+1)=1`) |
| `print_matrix.zoom_50_dense_fit_off` | `pages=4` with empty break arrays |
| `print_matrix.zoom_25_dense_fit_off` | `pages=6` with empty break arrays |

These are now in `tests/divergence.yaml` with `mode: skip-oracle`
scoped to `win-365-ja_JP`. **Re-capture goal**: change the driver to
read `PrintPreview`-mode `Pages.Count` rather than the display-zoom
property (whichever COM call the driver currently uses). The
`PrintPreview` count is what reconciles with `HPageBreaks /
VPageBreaks`. Once re-captured, remove those three entries from
`divergence.yaml`.

### Group 2: body-height calibration (4 cases) — needs driver extension

| Case | Symptom |
|---|---|
| `print_basic.print_titles_repeat_rows` | `h_breaks=[]` vs want `[39]` |
| `print_basic.print_titles_repeat_rows_and_cols` | `h_breaks=[]` vs want `[39]` |
| `print_matrix.print_titles_repeat_1_row` | `h_breaks=[]` vs want `[39]` |
| `print_matrix.print_titles_repeat_3_rows` | `h_breaks=[]` vs want `[39]` |
| (`print_matrix.print_titles_repeat_5_rows` passes — accidentally, because the prior body-subtraction logic compensates for one bug with another at depth=5) |

**Root cause hypothesis**: the body height our `compute_printable_area`
returns (`663.2pt` for A4 portrait Normal margins) is incompatible
with Excel's actual body height (`~595pt` inferred from the goldens).
A 40-row × 15pt print area (`A1:D40`) needs body `< 600pt` to fire a
break at row 39 -- and Excel does. The OOXML defaults we use (`top=0.75,
bottom=0.75, left=0.7, right=0.7, header=0.3, footer=0.3` in inches)
should yield body `= 691.2pt` minus our 28pt header/footer text strip
reservation `= 663.2pt`.

**Suspicion**: when the case YAML omits `page_margins`, Excel COM
applies a per-user / per-workbook-template preset that differs from
the OOXML spec defaults. **Wide** (1.0/1.0/0.5) would give body `~558pt`
which is close to the inferred `~595pt`; some other preset may give
exactly `~595pt`.

**Driver extension needed** -- extend `_apply_and_read_print` to
round-trip-read the actual margins and write them under a new
`applied_page_setup.margins` field:

```json
"applied_page_setup": {
  "zoom": 100,
  "fit_to_width": false,
  "fit_to_height": false,
  "margins": {
    "left": 0.7,
    "right": 0.7,
    "top": 0.75,
    "bottom": 0.75,
    "header": 0.3,
    "footer": 0.3
  }
}
```

Then re-capture all of `print_basic`, `print_matrix`, `print_pagination`,
`print_fit`. With actual margins recorded, the C++ side can either
(a) apply matching margins to its own `PageMargins` struct, or
(b) we add a small set of cases that explicitly set margins
(`{page_margins: {top: 1.0, ...}}` in the YAML) to exercise both
defaults and overrides. Either way, the absolute body number stops
being a moving target.

### Group 3: scale_50_shrinks_breaks mystery (1 case) — needs investigation

```
print_fit.scale_50_shrinks_breaks
  spec: A1:H1 (1 row, 8 cols * 30 chars), scale=50, no fit
  golden: pages=2, h_breaks=[], v_breaks=[]
  Formulon: pages=1, h_breaks=[], v_breaks=[]
```

At `scale=50`, content width = `1260 * 0.5 = 630pt`, body width = `494pt`.
Excel reports `pages=2` — geometrically consistent with an auto-column
break that splits the 630pt content into two 494pt-wide page-columns,
yet `v_breaks=[]` (consistent with the Block C rule that VPageBreaks
hides automatic column breaks).

But `print_matrix.zoom_75_dense_fit_off` has content width `1260 * 0.75
= 945pt`, body `494pt`, ratio `1.91` (would need 2 page-columns), and
reports `pages=1`. So the rule "count auto-col-pages at scale<100" is
not universal.

**Re-capture ask**: add 4 small cases to `print_matrix.yaml`:

```yaml
- id: scale_50_one_row_wide      # mirror of scale_50_shrinks_breaks
  sheets: {Sheet1: {A1: x, H1: x}}
  column_widths: {A:H: 30}
  print: {sheet: Sheet1, print_area: A1:H1,
          page_setup: {orientation: portrait, paper: 9, scale: 50}}

- id: scale_50_thirty_rows       # mirror of zoom_50 but only 1-row content
  sheets: {Sheet1: {A1: x, H30: x}}
  column_widths: {A:H: 30}
  print: {sheet: Sheet1, print_area: A1:H30,
          page_setup: {orientation: portrait, paper: 9, scale: 50}}

- id: scale_75_one_row_wide
  sheets: {Sheet1: {A1: x, H1: x}}
  column_widths: {A:H: 30}
  print: {sheet: Sheet1, print_area: A1:H1,
          page_setup: {orientation: portrait, paper: 9, scale: 75}}

- id: scale_75_thirty_rows       # mirror of zoom_75 but one shape change
  sheets: {Sheet1: {A1: x, H30: x}}
  column_widths: {A:H: 30}
  print: {sheet: Sheet1, print_area: A1:H30,
          page_setup: {orientation: portrait, paper: 9, scale: 75}}
```

The diff between `scale_50_one_row_wide` and `scale_50_thirty_rows` —
and between `scale_75` variants — should reveal whether Excel's
`Pages.Count` cares about row count when columns overflow at scale<100.

## Capture procedure (re-affirming)

```bash
# On Windows:
git pull origin develop                       # picks up driver + matrix updates
python tools/oracle/cli.py workbook --suite print_basic print_fit print_matrix print_pagination
git add tests/oracle/golden_wb/*.golden.json
# Confirm divergence-skipped cases were dropped from the regenerated goldens:
python tools/oracle/.venv/bin/python tools/oracle/workbook_case_schema.py \
    tests/oracle/cases_wb/print_matrix.case.json \
    tests/oracle/golden_wb/print_matrix.golden.json
git commit -m "feat(oracle): Capture margins and re-run pagination matrix"
git push origin develop
```

The Mac side picks it up via `git pull` and re-runs `make oracle-verify`.

## Acceptance criteria

After this round, the workbook track should be at **0 fail** for the
print suites (modulo the 3 divergence-skipped Block A entries). If
any case still fails, the failure mode is no longer "we don't know what
Excel is doing" -- it's a specific gap with documented evidence.
