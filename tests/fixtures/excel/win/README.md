# Windows Excel fixtures

Files saved by Windows Excel 365 ja-JP. Row heights and column widths depend on the display
scaling Excel ran under, and a `win-*` profile means 100% (96 DPI), so each set records it:

| Set | Display scaling | Layout figures usable for the profile |
|---|---|---|
| `column_width_dpi100/` | 100% (96 DPI, `results.json` `dpi`) | yes |
| `print_dpi100/` | 100% (96 DPI) | yes |
| `column_width/` | 175% (168 DPI, `results.json` `dpi`) | no; kept for the scaling comparison |
| `book_protection/` | not 100% | no; the files test protection records only |
| `pivot_days_end/`, `pivot_week_1900/`, `x14_structural/` | not recorded | not used for layout |
