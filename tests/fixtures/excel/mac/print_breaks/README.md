# Mac print breaks

`breaks.json` lists the vertical page breaks Mac Excel 16.115 (build 1108) reports for one-row print areas, used by `tests/unit/print/mac_print_breaks_test.cpp`.

Capture condition, shared by every case:

- No printer installed, so Excel paginates with its default page metrics.
- A4 portrait (`paperSize="9"`), scale 100, default margins (left/right 0.7 in, top/bottom 0.75 in, header/footer 0.3 in).
- One sheet, print area `$A$1:$T$1`, values in A1 and T1, and columns A:T all `<col width="W" customWidth="1">`.
- Font 0 (the Normal style's font) is the named face and size, with no `scheme` attribute and no theme part, so the name is read literally.

Each case gives the Normal font, the stored width `W`, the column's `Range.Width` in points (`width_pt`) and `Range.ColumnWidth` (`column_width`), and the 0-based columns each break precedes. `truncated` marks a case whose break list Excel cut short (it lists fewer breaks than the page count implies). Only the leading breaks it lists are compared for those cases.

The cases cover the Normal fonts whose Mac metrics `src/font_metrics.cpp` tabulates: Calibri 8, 11, 14 and 18, and 游ゴシック 11 and 14, at W = 2.00 to 41.00 in steps of 0.5.
