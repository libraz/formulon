# Upgrading from 0.18 to 0.19

0.19 makes the C ABI, the two JavaScript packages and the Python package one interface: the same names, the same records and the same error model. Almost every program written against 0.18 needs changes, and most of them are renames. A C caller must also recompile.

Read [Every binding](#every-binding) first, then the section for your binding. [Topics](#topics) explains the larger changes with before and after code, and the [reference tables](#reference-tables) list every former name with its replacement. The short list of what changed is in the [CHANGELOG](../../CHANGELOG.md).

## Overview

- Names follow one convention in C, JavaScript and Python. The former names are removed, with no aliases; [Name migration](#name-migration) maps each one.
- Sheet-scoped calls move from the workbook to a `Sheet` handle (`wb.sheet(key)`), and tables, pivots, pivot caches and drawings are handle classes with their own methods. See [Sheet handles](#sheet-handles).
- A failure throws in JavaScript and raises `FormulonError` or a subclass in Python. An argument error reads the same in every binding: the call, the argument path, then the constraint.
- One cell rectangle, `fm_cell_rect_t` (`CellRect`), replaces the per-feature range records. Records that live under a name or id are looked up by it, collections are read with one `list*` method, and a partial change is an `update<X>(patch)` call.
- An absent value reads as `NULL`, `null` or `None` on every read.
- The C ABI takes the sheet index explicitly, returns counts through an out-parameter and takes input records by `const` pointer. See [C ABI changes](#c-abi-changes).
- Several enums are renumbered; [Enums](#enums) lists the new values.
- Some calls keep their name but behave differently; see [Behaviour changes under an unchanged name](#behaviour-changes-under-an-unchanged-name).

## Steps by binding

### Every binding

- **Names:** one naming convention covers every public function, type and
  field in C, JavaScript and Python, and the former names are removed with no
  aliases. The table under [Name migration](#name-migration) lists each former
  name and what replaces it. Python static methods are module-level functions,
  and the `version` / `version_string` aliases are gone.
- **C ABI, JavaScript, Python:** one cell rectangle record, `fm_cell_rect_t`
  (`CellRect`), with inclusive 0-based corners, replaces `fm_merge_range`,
  `fm_print_range_t`, `fm_cf_cell_range_t`, `fm_viewport`, `MergeRange`,
  `DataValidationRange` and `ConditionalFormatRange`. Inverted or out-of-grid
  corners are rejected instead of normalised. Conditional-format and
  validation ranges, protected ranges and print areas take `CellRect` arrays
  instead of A1 text; table, chart and reference records carry `range` or a
  `SheetRect`.
- **C ABI, JavaScript, Python:** records that live under a name are looked up
  by that name (defined name, table style, cell style, hyperlink, person,
  threaded comment, shape, slicer, drawing object, passthrough part, theme
  script font). A missing key is reported as `out_present = 0` in C, as
  `null` in JavaScript and as `None` in Python. Adds that allocate an
  index (`fm_workbook_add_sheet`, `fm_sheet_add_validation`,
  `fm_workbook_add_cell_style_xf` and the others) take a trailing
  `size_t* out_index`. `fm_workbook_evaluate_formula_array_cell` is replaced by
  an owned `fm_value_array_t`, and slicers are addressed by the object id
  `fm_sheet_add_slicer` returned.
- **C ABI, JavaScript, Python:** an argument that cannot be right is refused
  by what it addresses. A position (a sheet index, a collection index) that is
  negative or past the end is `IndexOutOfRange` (`SheetIndexOutOfRange` for a
  sheet; Python `IndexOutOfRangeError`), and a coordinate (row, column) or a
  quantity (count, length) that is negative or outside its limit is
  `InvalidArgument` (Python `InvalidArgumentError`). A non-number is
  `TypeError`; in JavaScript a fractional or non-finite number is `RangeError`;
  in Python any `float` or `bool` where an integer is expected is `TypeError`
  and an integer wider than the parameter is `ValueError`. A negative value
  used to be `RangeError` or `ValueError`. The message names the method and the
  parameter and never a C type; match on the class and `code`, not the text.
  Python `wb[-1]` counts from the end and is the last sheet, while
  `wb.sheet(-1)` is `IndexOutOfRangeError`.
- **C ABI, JavaScript, Python:** the statuses `kSheetIndexOutOfRange`,
  `kInvalidSheetName` and `kSheetCountLimitExceeded` move from 5050, 5051 and
  5053 to 10, 11 and 12 (`SheetIndexOutOfRange`, `InvalidSheetName`,
  `SheetCountLimitExceeded`), so that general codes sit in the general range;
  5050 to 5053 are retired and never reused. A caller that compares numbers
  must update them.
- **C ABI, JavaScript, Python:** the status 9006, documented as
  `kThreadedCommentInvalid`, is retired and the number is never reused. A
  malformed timestamp, a mention outside the text, an empty display name,
  `done` on a reply, removing a person a comment still names and the other
  cases it covered are `InvalidArgument`.
- **C ABI, JavaScript, Python:** "absent" is one thing on every read: `NULL` in
  C, `null` in JavaScript and `None` in Python, never `""`, `0`, `-1` or a
  presence flag in the record. `getSpill` returns `Spill | null`,
  `getCellXfIndex` is `null` when no cell is stored at the address (a stored
  cell without a style of its own reads 0),
  `PivotCell.fieldName`, `FunctionMetadata.signature` and `description`,
  `AutoFilterColumn.dxfIndex` and the sparkline group colours read `null`, and
  `PageSetup.orientationStated` is gone (`orientation === null` says it).
  `PhoneticProperties.present` is `stated`; like every read-only key of a read
  record it is accepted and ignored on input, so `set(get())` round-trips. In
  Python the same holds for the password fields of a protected range,
  `PivotCell.field_name` and `number_format`, the totals-row label and formula
  of a table column, `ExternalLink.target`, the `formula1`, `formula2` and
  `text` of a conditional format, `ChartTitle.text` and `formula`,
  `PivotCacheOptions.missing_items_limit` (`None` is automatic, and `-1` is
  `InvalidArgumentError`), and `Pagination.print_titles`. `requested_id` of
  `add_pivot_cache` is `Optional[int] = None`, and an explicit `0` is
  `InvalidArgument` in every binding (C keeps `0` for "assign one").
- **JavaScript, Python, C ABI:** an argument error reads
  `<Class>.<method> <param>[.<field>] <constraint>` in WebAssembly, native and
  Python alike, for example `Sheet.updateView patch.zoomPct must be an integer`
  and `Sheet.remove_validation validation 99 is out of range (0 items)`. The
  class is the one the call was made on (`WorkbookOoxml`, `PivotCache`,
  `Picture`, `Chart`, `Workbook.load` for a static, `Formulon.<fn>` for a
  module function), a record field is named by its path (`patch.sparklines[2].location`),
  and the constraint comes from a fixed vocabulary: `must be a boolean` /
  `a number` / `a string` / `an object` / `an array`, `must be an integer`,
  `is required`, `has no field <key>`, `must not be null`, `is too large`,
  `must not be negative`, `must be in <lo>..<hi>`, `must not be empty`,
  `<v> is out of range (<n> items)` and `takes at most <n> arguments`. The
  Python `formulon: ` prefix, the `argument <n>` wording and C type names
  (`uint32`, `[0, 4294967295]`) no longer appear. Error classes follow the
  constraint: a wrong type is `TypeError`; a fractional, NaN or infinite number
  is `RangeError` (Python `TypeError`); a negative value for an unsigned field
  or a coordinate is `InvalidArgument` instead of `RangeError`; a value beyond
  32 bits is `RangeError` (Python `ValueError`); an index past the end is
  `IndexOutOfRange`.
- **Behaviour:** a conditional formatting rule whose id is already in use gets
  a fresh id instead of being refused. A JavaScript `Date` written to a
  General cell gets the date format `m/d/yyyy` (id 14) at local midnight and
  `m/d/yyyy h:mm` (id 22) otherwise; a Python `date` takes 14, a `datetime`
  22 and a `time` `h:mm:ss` (21).
- **C ABI, JavaScript, Python:** one unit per quantity, named by its suffix
  and converted at the boundary. Row heights are points (`heightPt`,
  `setRowHeight(rows, heightPt)`, `setRowAutoHeight(rows, heightPt)`,
  `defaultRowHeightPt`, `effectiveDefaultRowHeightPt`), column widths are
  characters (`widthChars`, `defaultColumnWidthChars`, `baseColumnWidthChars`,
  `effectiveDefaultColumnWidthChars`), page margins are inches (`leftIn`,
  `rightIn`, `topIn`, `bottomIn`, `headerIn`, `footerIn`), and the print scale
  is a percentage (`scalePct` on page setup and pagination). Other renames:
  `zoomScale` is `zoomPct`, `transparency` is `transparencyPct`, `rotation` is
  `rotationDeg`, `firstSliceAngle` is `firstSliceAngleDeg`, `gapWidth` is
  `gapWidthPct`, `normalFontSize` is `normalFontSizePt`, a font's `size` is
  `sizePt`, and an image's `pxWidth` / `pxHeight` are `widthPx` / `heightPx`.
  Fields with no unit (`indent`, `textRotation`, `fitToWidth`, `fitToHeight`
  and the like) keep their names.
- **C ABI, JavaScript, Python:** a method that changes part of a record is
  `update<X>(patch)` and merges the keys it names: `updateIterative`,
  `updateTableStyleDefaults`, `updatePageSetup`, `updatePageMargins`,
  `updatePrintOptions`, `updateHeaderFooter`, `updateView`,
  `updateFormatDefaults`, `updateCellFormat`, `Pivot.updateGrandTotals`,
  `Pivot.updateOptions`, `PivotCache.updateOptions`, `Chart.updateLegend`,
  `Chart.updateAxis`, `Chart.updateDisplay`, `Chart.updateSeries`,
  `Chart.updatePoint` and `Chart.updatePlot` (C `fm_sheet_update_*` and
  `fm_workbook_update_*`). `set<X>(record)` replaces the whole value. Python
  patches are keyword-only arguments that default to `UNCHANGED`.
- **C ABI, JavaScript, Python:** a record that exists once (sheet protection,
  an AutoFilter, a table AutoFilter, a chart title, a chart legend, a chart
  point override, the pinned time, the theme, a row or column break, a
  hyperlink) is removed with `set<X>(null)` (C: `NULL`), and its separate
  `remove*` / `clear*` method is gone. `remove*` of a keyed record that is
  missing fails with `NotFound`; `clear*` does nothing when there is nothing
  to clear. `reset<X>` restores a default and replaces the former `clear`
  names: `resetRowHeight`, `resetColumnWidth`, `resetCellXfIndex`,
  `resetColumnXfIndex`, `resetRowXfIndex` and `resetTheme`. The criteria of an
  AutoFilter are emptied with `clearAutoFilterCriteria()` on the sheet and on
  a table.
- **C ABI, JavaScript, Python:** table lifecycle lives on `Sheet`, like
  pivots: `ws.addTable(input)`, `ws.table(ref)`, `ws.listTables()` and
  `ws.removeTable(ref)`. A table index is local to its sheet, and every C table
  function is `fm_sheet_*table*`. `wb.table(name)` (`fm_workbook_find_table`)
  stays as the one workbook-wide lookup by name. A table's `name` is the name
  formulas use, and `table.rename('Rev')` changes it; the file's separate
  display name is no longer a public field.
- **C ABI, JavaScript, Python:** a table or PivotTable style is read and
  written with `getStyleOptions()` / `updateStyleOptions(patch)` and the
  records `TableStyleOptions` / `PivotStyleOptions` (`styleName` and the
  `show*` toggles); `TableProperties` no longer carries them. C has
  `fm_table_style_options_t`, `fm_pivot_style_options_t`,
  `fm_sheet_get_table_style_options`, `fm_sheet_update_table_style_options`
  and `fm_sheet_update_pivot_style_options`.
- **C ABI, JavaScript, Python:** sheet protection is
  `ws.setSheetProtection(input | null)` (replace, or unprotect) and
  `ws.updateSheetProtection(patch)`. There is no `enabled` field: a record
  means protected and `null` means unprotected. The `allow*` flags carry
  Excel's dialog defaults and the engine stores them without enforcing locks.
  Each `allow*` flag is the inverse of the 0.18.0 field it replaces, which
  mirrored the OOXML attribute: `formatCells: 1` prevented formatting, and
  `allowFormatCells: true` permits it (Python `format_cells` and
  `allow_format_cells`; C `format_cells` and `allow_format_cells`). Code that
  carried a value across must negate it.
- **C ABI, JavaScript, Python:** page breaks are `RowBreak { row, first, last,
  manual }` and `ColumnBreak { col, first, last, manual }` (`first` and `last`
  are the span the break covers). `setRowBreak(row, { first?, last?, manual? }
  | null)` and `setColumnBreak(col, ...)` set or remove one,
  `listRowBreaks()` / `listColumnBreaks()` read them, and
  `Pagination.rowBreaks` / `columnBreaks` hold the records. C has
  `fm_sheet_set_row_break(wb, sheet_index, row, const fm_row_break_t* or
  NULL)` and `fm_sheet_set_column_break`.
- **C ABI, JavaScript, Python:** a cell has at most one hyperlink, so
  `ws.setHyperlink(cell, link | null)` replaces `addHyperlink` and
  `addHyperlinkRange`; the range the link covers is `range` in the
  `HyperlinkInput` record (omitted or `null` is the key cell alone, and a range
  that does not start at the key cell is `InvalidArgument`).
- **C ABI, JavaScript, Python:** `Pivot.addFieldItem(field, name | {
  cacheIndex }, options)` takes either a label or a cache index (Python
  `add_field_item(field, name=None, *, cache_index=None, ...)` with exactly
  one of the two; C `fm_sheet_add_pivot_field_item_by_cache_index`).
  `addFieldSharedItem` exists only on `PivotCache`. `Pivot.setAnchor` takes
  one cell as `PivotInput.anchor` does, and the rename function is
  `fm_sheet_rename_pivot`.
- **C ABI, JavaScript, Python:** conditional-format rules refer to a
  differential format by `dxfIndex` (C `fm_cf_rule_t.dxf_index`); get one from
  `wb.ooxml.addDxf(format)`. A `Dxf` carries `numberFormat`, the format code
  (`string | null`, C `fm_dxf_t.number_format`), and the number-format id is
  allocated on save. A cell XF uses `styleXfIndex`, and `horizontal` /
  `vertical` for its alignment. The XF-index family is spelled `*Xf*`
  (`setColumnXfIndex`, `setRowXfIndex`, `getColumnXfIndex`, `getRowXfIndex`,
  `remapRangeXfs`, `snapshotRangeXfs`, `restoreRangeXfs`), and "style" means
  a named style only. Icon sets use `iconSet` with `iconIndex`, and
  `setCellPhoneticProperties` takes `fontIndex`.
- **C ABI, JavaScript, Python:** one table per concept replaces the per-feature
  copies, and a table's 0 is the file format's default. `ComparisonOperator`
  (C `fm_comparison_operator_t`) is the operator of a `cellIs` rule, a
  validation and an AutoFilter condition; it replaces `FilterOperator` and the
  bare operator numbers of rules and validations. `DatePeriod`
  (`fm_date_period_t`) is the `timePeriod` of a conditional-format rule, which
  was a bare number. `BlanksAs` (`fm_blanks_as_t`) covers a chart's `blanksAs` and
  a sparkline group's `emptyCellsAs` and replaces `ChartBlanksAs` and
  `SparklineEmptyCells`. The new numbers are listed under [Enums](#enums).
  A custom AutoFilter condition states its operator (an
  `AutoFilterCondition` in the `conditions` list), and the setter rejects
  `Between` and `NotBetween` with `InvalidArgument`.
- **C ABI, JavaScript, Python:** a string that named a closed set is an enum
  number: the totals-row function of a table column
  (`TableTotalsRowFunction`), a table style element type
  (`TableStyleElementType`), an axis tick-label position
  (`ChartTickLabelPosition`), the domain of chart data (`ChartDomain`), the
  platform of a width model (`WidthPlatform`) and the icon set of a custom
  icon level (`IconSetName`, which gains the four sets only a custom level
  takes). In C each field is an `int32_t` holding the matching typedef
  (`fm_table_totals_row_function_t`, `fm_table_style_element_type_t`,
  `fm_chart_tick_label_position_t`, `fm_width_platform_t`,
  `fm_icon_set_name_t`); an absent axis tick-label position reads
  `FM_CHART_TICK_LABEL_POSITION_ABSENT` (-1).
- **C ABI, JavaScript, Python:** a bare number with a fixed set of meanings is
  an enum: `PhoneticType` and `PhoneticAlignment` (`fm_phonetic_type_t`,
  `fm_phonetic_alignment_t`, values unchanged), `PageOrder`
  (`Pagination.pageOrder`, `fm_page_order_t`), `FunctionAvailability`
  (Python `MergedFunctionMetadata.availability`) and `ChartPointState`
  (Python `ChartDataSeries.states`). `PaperSize` names the eight codes the
  engine knows (Letter 1, Legal 5, A3 8, A4 9, A5 11, B4 12, B5 13, A6 70);
  `paperSize` stays a number, because other codes round-trip. C also has
  `fm_effective_style_source_t` and `fm_theme_source_t`. The mask fields of
  pivot date grouping are typed `PivotDatePeriodMask` in JavaScript (an OR of
  `PivotDatePeriod` bits) and are an `IntFlag` in Python.
- **C ABI, JavaScript, Python:** the theme script font takes `scheme`, a
  `FontScheme` (`Major` or `Minor`), in place of the `major` boolean:
  `wb.setThemeScriptFont({ scheme, script, typeface })`,
  `wb.getThemeScriptFont({ scheme, script })`, Python
  `set_theme_script_font(ThemeScriptFont(scheme=..., ...))`, and `int32_t scheme` in the C
  functions. `FontScheme.None` and any other value are `InvalidArgument`.
- **C ABI, JavaScript, Python:** one drawing surface. `DrawingProperties`
  carries `name`, `description`, `hidden` and one nested `anchor`
  (`DrawingAnchor`: `kind`, `editAs`, `fromRow`, `fromCol`, the four offsets,
  `toRow`, `toCol`, `widthEmu`, `heightEmu`) for every kind, and
  `ShapeProperties`, `SlicerProperties` and `ChartProperties` hold only
  their own fields (`ChartProperties` has no name or description; `format`
  is on `PictureProperties`). Input takes `anchor: DrawingAnchorInput` in
  place of the flat `anchorKind`, `from*`, `to*` and offset keys
  (`ShapeInput`, `SlicerInput`, `DrawingPlacementInput`, the chart's
  `placement`; Python `DrawingAnchorInput`, `DrawingPlacementInput`).
  `Drawing.update(patch)` (Python `update(name=, description=, hidden=,
  anchor=)`) changes the common fields of a picture, chart, shape or slicer;
  `rename` and `setAnchor` (`set_anchor`) are removed from them, and
  `Pivot.rename` and `Pivot.setAnchor` stay. `AnchorKind` and `AnchorEditAs`
  gain `Default` (0), so the other members move up by one ([Enums](#enums)),
  and a reader reports the stored `editAs`, never `Default`.
- **C ABI, JavaScript, Python:** the series of a chart is keyed by `seriesId`
  (`ChartSeries.seriesId`, `ChartDataSeries.seriesId`,
  `ChartUnresolved.seriesId`, and the `seriesId` parameter of `updateSeries`,
  `setPoint`, `updatePoint` and `removeSeries`; Python `series_id`). `addSeries`
  returns it. A point's `index` is still a position.
- **C ABI, JavaScript, Python:** `insertRows`, `deleteRows`, `insertColumns`
  and `deleteColumns` take `(first, last)` instead of `(start, count)`;
  a call written for the old form inserts a different number of rows.
  `ws.insertRows(3, 5)` inserts three rows, and `first > last` is
  `InvalidArgument`.
- **C ABI, JavaScript, Python:** `insertRows` and `insertColumns` fail with
  `PreconditionFailed` and change nothing when they would push a value,
  formula, note, fill, border or merge past the last row or column of the grid;
  they used to drop it. A font or number format, hyperlink, validation,
  conditional format or row or column format pushed off is still dropped.
- **C ABI, JavaScript, Python:** `Sheet.insertCells(range, shift)` and
  `deleteCells(range, shift)` fail with `PreconditionFailed` and change nothing
  when the shifted region meets a table, a multi-cell array formula or a data
  table that extends beyond the band, a PivotTable edge or inside, or when a
  value, formula, note, fill, border or merge would be pushed off the grid;
  `InvalidArgument` for a wrong shift direction or a whole-column range
  shifted down (whole-row shifted right).
- **C ABI, JavaScript, Python:** `PivotAxis.Col` is `PivotAxis.Column`
  (`FM_PIVOT_AXIS_COL` is `FM_PIVOT_AXIS_COLUMN`; Python `PivotAxis.COLUMN`),
  and `ChartBarDirection.Col` is `ChartBarDirection.Column`. The
  show-values-as members `PercentOfCol`, `RunningTotalInCol` and
  `PercentOfParentCol` are spelled out the same way (C
  `FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_COLUMN`, `_RUNNING_TOTAL_IN_COLUMN` and
  `_PERCENT_OF_PARENT_COLUMN`; Python `PERCENT_OF_COLUMN` and so on). `Col`
  remains only inside coordinate names such as `firstCol`.
- **C ABI, JavaScript, Python:** a chart title of `''` is `InvalidArgument` in
  `addChart`, as it already was in `setTitle`; `null` (`None`) is no title and
  an omitted title is Excel's automatic one. `fm_chart_spec_t` (was
  `fm_chart_insert`) replaces `const char* title` with `int32_t has_title` and `fm_chart_title_t title`
  (`has_title == 0` is automatic, a title of kind `NONE` is no title), so the
  insert takes the title shape `fm_sheet_set_chart_title` takes.
- **C ABI, JavaScript, Python:** every named collection (sheets, tables, table
  columns, defined names, cell styles, custom table styles, protected ranges,
  pivots) refuses to add, rename or copy a name that equals an existing one
  under Excel's text comparison, which folds letter case, full- and half-width
  forms and hiragana / katakana: `Data`, `ＤＡＴＡ`, `ﾃﾞｰﾀ` and `データ` collide.
  A lookup resolves an exact match, else the one entry equal under Unicode
  simple case folding, else the single entry equal under the full comparison;
  two or more such entries resolve nothing. A file with names that collide
  only under the full comparison loads as it is, and a new name colliding with
  either is refused. GUID ids and person ids compare by letter case only (a
  GUID is accepted in either case and stored uppercase), and `=ä` resolves a
  name `Ä`. Names used to fold ASCII letters only, or compare exactly.
- **C ABI, JavaScript, Python:** `addMerge` refuses a range that overlaps an
  existing merge, an identical one included, with `AlreadyExists`, and one
  that overlaps a table, a pivot, an array formula or a data table with
  `PreconditionFailed`. A workbook that already carries overlapping merges
  still loads.
- **C ABI, JavaScript, Python:** the page-axis field order of a pivot is read
  and written like the row and column orders: `Pivot.getPageFieldOrder()` and
  `setPageFieldOrder(indices)` (`get_page_field_order`, `set_page_field_order`),
  `fm_sheet_get_pivot_page_field_order` and `fm_sheet_set_pivot_page_field_order`.
  The Values level is `InvalidArgument` there, and a saved file lists the page
  fields in this order.
- **C ABI, JavaScript, Python:** a conditional-format rule is addressed by its
  id and a protected range by its name, as validations are. `addCfRule` returns
  the stored `CfRuleId` (a string; `CfRuleRef` is a position or an id) and
  `addProtectedRange` returns the name. The C functions
  `fm_sheet_update_cf_rule`, `fm_sheet_remove_cf_rule`,
  `fm_sheet_update_protected_range` and `fm_sheet_remove_protected_range` take
  the id or name in place of a position, `fm_sheet_get_cf_rule` is new, and
  `fm_sheet_get_protected_range` reads by name; the positional reads are
  `fm_sheet_get_cf_rule_at` and `fm_sheet_get_protected_range_at`. A rule's
  `id` is a brace-wrapped GUID; one that is omitted, empty or already used on
  the sheet is replaced by an upper-case one the engine generates. See
  [C ABI changes](#c-abi-changes).

### C ABI callers

- **C ABI:** a function that acts on a sheet takes `size_t sheet_index`
  explicitly. Count getters return `fm_status_t` and write the count through
  `size_t* out_count`; a NULL handle reports
  `FM_STATUS_BINDING_NULL_POINTER`. Getters take `const` handles, enum
  parameters and fields cross the ABI as `int32_t`, and
  `fm_workbook_paginate` is `fm_sheet_paginate`.
- **C ABI:** an out-of-range sheet index returns
  `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10) instead of
  `FM_STATUS_INVALID_ARGUMENT`. Every sheet-taking function checks the handle,
  then NULL pointers, then the sheet index, then the remaining arguments. A
  successful call leaves `fm_get_last_error_message` empty.
- **C ABI:** fifteen functions take their input record as `const T*` instead of
  by value (`fm_sheet_set_hyperlink`, `fm_sheet_add_merge`,
  `fm_sheet_add_validation`, `fm_sheet_add_cf_rule`, `fm_workbook_add_font`,
  `fm_workbook_add_cell_xf` and the other record-taking add and set
  functions); a NULL record returns `FM_STATUS_BINDING_NULL_POINTER`. The row
  setters (`fm_sheet_set_row_height`, `_set_row_auto_height`, `_set_row_hidden`,
  `_set_row_outline`, `_set_row_collapsed`, `fm_sheet_reset_row_height`) take an
  inclusive `(sheet_index, first, last, ...)` span like the column functions;
  `last < first` returns `FM_STATUS_INVALID_ARGUMENT`.
- **C ABI:** an index past the end of a collection other than sheets returns
  `FM_STATUS_INDEX_OUT_OF_RANGE`, and removing a record by a natural key that
  does not exist returns `FM_STATUS_NOT_FOUND`. Error message text no longer
  names C functions and is for display, not comparison.
  `fm_workbook_get_sheet_index` finds a sheet by name and
  `fm_get_version_number` returns the version as one number.
- **C ABI:** a few records and functions change shape; a caller compiled
  against an earlier header must be recompiled. See [C ABI changes](#c-abi-changes).
- **C ABI:** `fm_workbook_save` takes the `int32_t format` that
  `fm_workbook_save_as` took, and `fm_workbook_save_as` is removed:
  `fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len)`.
- **C ABI:** the message of an argument error leads with the C parameter or
  field (`max_iterations must be in 1..32767`), without a function name. A NULL
  top-level pointer returns `FM_STATUS_BINDING_NULL_POINTER`; a NULL pointer or
  string inside a record where a value is required returns
  `FM_STATUS_INVALID_ARGUMENT` (the AutoFilter arrays, threaded-comment text
  and mentions, validation, sparkline items and collation fields returned
  `FM_STATUS_BINDING_NULL_POINTER` before).
  `fm_iterative_settings_patch_t.max_iterations` must be 1..32767 when set;
  0 or a larger value is `FM_STATUS_INVALID_ARGUMENT` and is no longer stored
  as 32767 (`updateIterative`, `update_iterative`). A `dxf_index` or
  `style_xf_index` past the end of its table returns
  `FM_STATUS_INDEX_OUT_OF_RANGE`. The split fields of `fm_sheet_view_t`,
  `split_x` and `split_y`, are `split_x_twips` and `split_y_twips` in
  `fm_sheet_view_info_t` and `fm_sheet_view_patch_t` (with the `_op`
  companions).
- **C ABI:** `fm_sheet_set_pivot_field_date_group` (`Pivot.setFieldDateGroup`,
  `Pivot.set_field_date_group`) takes a mask of `FM_PIVOT_PERIOD_*` bits, a
  calendar, an interval in days and start and end serials (NaN for automatic)
  in place of a granularity and year arguments. `fm_pivot_date_grouping_t` and
  the `date_*` members of `fm_pivot_field_info_t` are gone; the struct now
  reports `group_base_index`, `insert_blank_row` and `fill_down_labels`, and
  a field's date grouping is read with `fm_sheet_get_pivot_field_date_group`
  (`Pivot.getFieldDateGroup`, `Pivot.get_field_date_group`), which returns
  the settings `fm_sheet_set_pivot_field_date_group` takes. The old "year and month" grouping
  is not the same thing as Excel's months period, so it has no replacement.
  `fm_pivot_item_info_t` gains `caption` and `show_detail` at its end. A caller
  compiled against an earlier header must be recompiled.
- **C ABI:** `fm_pivot_filter_t` gains `value_high_text` at its end, the
  upper bound of a label Between / Not Between filter, and
  `fm_pivot_filter_kind_t` covers every OOXML pivot filter type (label, value,
  top / bottom count, percent and sum, date comparison, relative period,
  quarter and month); the existing numbers are unchanged.
  `fm_pivot_field_info_t` gains `show_all` at its end. A caller compiled
  against an earlier header must be recompiled.
- **C ABI:** `fm_sheet_update_chart_display` takes a `fm_chart_display_patch_t`
  (keep / clear / set per field, like the legend and axis patches) in place of
  the `blanks_as` and `plot_vis_only` arguments of the former
  `fm_sheet_set_chart_display`. To port a call, clear a field
  whose old value was `ABSENT` and set the others; `ABSENT` as a set value is
  now `FM_STATUS_INVALID_ARGUMENT`. `fm_chart_info_t` gains `na_as_blank`,
  `fm_chart_preview_overrides_spec_t` gains `na_as_blank_op` / `na_as_blank`, and
  `fm_spill_info_t` gains `blocked`, each at its end. A caller compiled
  against an earlier header must be recompiled.
- **C ABI:** every input and output record starts with `uint32_t struct_size`.
  Set it with `FM_STRUCT_INIT(T)` (`{.struct_size = sizeof(T)}` in C,
  `fm_struct_init<T>()` in C++) before passing the record; a size below the
  layout of the record in this release returns `FM_STATUS_INVALID_ARGUMENT`,
  and a larger size is read as far as the library knows. From 1.0 a record
  gains fields only at its end. Value types (`fm_cell_rect_t`, `fm_value_t`,
  `fm_color_t`, the style records and the chart parts) are frozen.
  `FM_CHECK_VERSION()` compares the header version with the library's.
- **C ABI:** colours share one record. `fm_color_t` (`kind`, `argb`,
  `theme`, `tint`, `indexed`) replaces `fm_cf_color_t` and the `*_argb`
  fields of the chart fill, line and label font records. Before, a record held
  an `*_argb` integer beside a theme index; after, one `fm_color_t` field holds
  either. The shape fields `fill_rgb` and `line_rgb` are `fill` and `line`, an
  `fm_chart_fill_t` and an `fm_chart_line_t` as in the chart records.
  A caller compiled against an earlier header must be recompiled.
- **C ABI:** the text-comparison functions take the collation strength before
  the profile id. `fm_compare_text(a, b, profile_id, strength, out_order)`
  is `fm_compare_text(a, b, strength, profile_id, out_order)`, and
  `fm_get_collation_classes(values, count, profile_id, strength, ...)` is
  `fm_get_collation_classes(values, count, strength, profile_id, ...)`. Both
  take `const char*` and `int32_t`, so a stale call compiles and passes the
  wrong arguments.
- **C ABI:** the three CLI-only statuses `kCliInvalidArgs` (7005),
  `kCliFileNotFound` (7006) and `kCliOutputFailed` (7007) are removed from the
  engine's status range, and the numbers stay unused. They are internal to the
  command line now.
- **C ABI:** formula entry fails with a parser status (1000 to 1014) instead
  of a generic invalid-argument status, and `fm_get_last_parse_error` returns
  the code, the UTF-16 start and end, and the line and column. Reading a value
  recalculates first when the calculation mode is automatic, so a value read
  can return the `FM_STATUS_GRAPH_*` recalculation statuses. The values of the
  CF rule, operator, time-period and value-object fields, the validation type
  and operator and the icon set have typed enums; the fields keep their
  integer types. `fm_workbook_get_profile_id` and `fm_workbook_set_profile_id`
  replace the `*_excel_profile_id` functions.
- **C ABI:** every function that returns a pointer states how long it stays
  valid: static (the process), transient (until the next `fm_*` call on the
  same thread or the next call on the same handle) or owned (until the named
  destroy function). Code that kept a transient pointer across calls was
  already wrong; the header now says so per function.
- **C ABI:** every C enum constant starts with its type name
  (`FM_PIVOT_FILTER_KIND_*`, `FM_PIVOT_FILTER_OPERAND_KIND_*`,
  `FM_SLICER_CROSS_FILTER_*`), `plot_visible_only` is spelled out, and the
  pivot layout handle is `fm_pivot_report_t` with `fm_pivot_report_*`
  functions (`fm_pivot_report_layout_t` is the report-layout enum). `col` is used
  only for coordinates and `column` for every concept, byte counts are
  `out_byte_count`, and a picture is added with `fm_sheet_add_picture` and
  `fm_drawing_placement_spec_t`. An optional boolean in a `_spec_t` input is the
  new `fm_tristate_t` (`FM_TRISTATE_DEFAULT` 0 is Excel's default,
  `FM_TRISTATE_FALSE` 1, `FM_TRISTATE_TRUE` 2); read records and `_patch_t`
  booleans stay `int32_t` 0 or 1.
- **C ABI:** `fm_workbook_evaluate_formula` is removed; use
  `fm_sheet_evaluate_formula_array` and take the top-left cell of a result.
  The bindings' `evaluateFormula` does that.

### JavaScript and Python callers

- **JavaScript, Python:** every collection is read with one `list*` method
  (`listCfRules()`, `listValidations()`, `listTables()`, `listFonts()`
  and the others in the migration table) that returns a fresh array. A count
  method paired with a getter by index (`tableCount()` + `tableAt(i)`) is
  removed together with the getter; `listTables()[i]` replaces the pair. A
  phonetic run list is read with `getCellPhoneticRuns(row, col)` instead of
  `getCellPhoneticRunCount` and `getCellPhoneticRun`.
- **JavaScript, Python:** `spillInfo` (`spill_info`) is `getSpill`
  (`get_spill`) and returns `{ blocked, range }` or `null`. The constants
  `PIVOT_SHOW_AS_BASE_PREVIOUS` and `PIVOT_SHOW_AS_BASE_NEXT` are
  `PivotShowAsBase.Previous` and `PivotShowAsBase.Next` (Python an `IntEnum`
  with `PREVIOUS` and `NEXT`). `pivotFieldClearDateGroup`
  (`pivot_field_clear_date_group`) is gone: `setFieldDateGroup(field, null)`
  (`set_field_date_group(field, None)`) removes the grouping and does nothing
  for an ungrouped field.
- **JavaScript, Python:** `evaluateCfRange(range, { today, timePeriodRules })`
  takes `timePeriodRules: false` to match no `TimePeriod` rule (combined with
  `today` it is `InvalidArgument`), and Python
  `evaluate_cf_range(range, *, today=None, time_period_rules=True)` takes both
  as keywords; a positional `today` is `TypeError`.
- **JavaScript, Python:** a boolean argument or field takes `true` or `false`
  only. `setDate1904('yes')`, `setCalcOnSave(null)`, `parseReference(text,
  { r1c1: 1 })`, `updateView({ showGridLines: 1 })` and the Python counterparts
  (`set_date1904(1)`, `format_reference(range, absolute=1)`) throw `TypeError`;
  they used to be read as truthy or falsy. `Workbook.setIterative` errors are
  labelled `Workbook.updateIterative`, and `PivotCache` errors name
  `PivotCache.<method>`. A static factory (`Workbook.createDefault`,
  `Workbook.load`) and a module function refuse extra arguments with `takes at
  most <n> arguments`.
- **JavaScript, Python:** `enumName(StatusCode, code)` names a status.
  `SheetView.showRowColHeaders` is `showRowColumnHeaders`, and
  `SheetFormatDefaults` spells `defaultColWidth`, `baseColWidth` and
  `effectiveDefaultColWidth` as `defaultColumnWidthChars`,
  `baseColumnWidthChars` and `effectiveDefaultColumnWidthChars`, as do
  `SheetFormatDefaultsInput` and its patch (Python
  `default_column_width_chars` and so on, also as `update_format_defaults`
  keywords).
- **JavaScript, Python:** `formatValue(value, numberFormat)` takes a native
  cell value (a number, string, boolean, `Date`, `ExcelError` or tagged
  `Value`) and converts it as `setValue` does; `undefined` and an array are
  `TypeError`. Python `format_value` converts the same way. A `DateInput` is a
  `CivilTimeInput` (`hour`, `minute` and `second` default to 0) or a `Date`, in
  every slot including `convertCivilToSerial`. `ThreadedCommentInput.created`
  reads a `Date` through its UTC getters and a civil object as UTC; Python
  converts an aware `datetime` to UTC, treats a naive one and a `date` as UTC,
  and returns `created` as naive UTC. `setPinnedNow` keeps the local getters.
- `Chart.updateDisplay(patch)` (JavaScript) and `Chart.update_display(patch)`
  (Python) take one patch object in place of the positional arguments of
  `setChartDisplay(sheet, objectId, blanksAs, plotVisibleOnly)`: `blanksAs`, `plotVisibleOnly` and
  `naAsBlank` (`blanks_as`, `plot_visible_only`, `na_as_blank`). An omitted
  key keeps the element, `null` / `None` removes it, a value sets it.
- **Handles:** `Table`, `Pivot`, `PivotCache` and `Drawing` (with its
  subclasses `Picture`, `Chart`, `Shape` and `Slicer`) are handle classes that
  stay on their object while others come and go. An operation on the object is
  a method of the handle: `pivotDataFieldAdd(sheet, p, spec)` is
  `ws.pivot(p).addDataField(spec)`, and a chart's title is set with
  `ws.chart(id).setTitle(...)`; Python has the same classes in `formulon.Table`,
  `formulon.Pivot`, `formulon.PivotCache` and `formulon.Drawing`. The parent
  keeps `add*`, `remove*`, `list*` and the accessor (`ws.table(name)`,
  `ws.pivot(p)`, `ws.chart(id)`), and every `add*` returns the handle.
  `pivotSetName` is `Pivot.rename`, and a drawing is renamed with
  `Drawing.rename`. `wb.renameSheet(s, n)` is
  `wb.sheet(s).rename(n)` and `wb.moveSheet(s, i)` is
  `wb.sheet(s).move(insertBefore)` (see below).
- **Handles:** `listTables()`, `listPivots()`, `listCharts()`, `listShapes()`,
  `listSlicers()` and the other `list*` calls return handles. Code that passed
  the result to `JSON.stringify` maps each element through `getProperties()`
  first. The record types are named `<HandleClass>Properties`: `TableEntry` is
  `TableProperties`, `ChartInfo` is `ChartProperties`, `ImageResult` is
  `PictureProperties`, `ImageInsertOptions` is `DrawingPlacementInput`, `DrawingObject`
  is `DrawingProperties`, `Shape` is `ShapeProperties`, `Slicer` is
  `SlicerProperties`, and `PivotProperties` is new. In Python `Table`, `Chart`,
  `Shape` and `Slicer` are handle classes instead of records, and `__all__`
  exports both the classes and the `*Properties` records.
- **Pivot tables:** a pivot number format is a format code. The
  `String(wb.addNumFmt(code))` bridge is gone: pass the code itself.
- **Renames:** `evalFormula` is `evaluateFormula` and `eval_formula` is
  `evaluate_formula`. `setSheetFitToPage(sheet, b)` is `ws.updatePageSetup({ fitToPage: b })`.
  `setSparklineGroupSparklines` and `setThreadResolved` are folded into
  `updateSparklineGroup` and `updateThreadedComment`. `insertImage` is
  `addPicture`; `removeShape` and `removeSlicer` are `removeDrawing`;
  `listDrawingObjects` is `listDrawings`, and `listPictures` is new. Python
  uses the snake_case equivalents.
- **Cell formats:** `updateCellFormat(range, patch)` and `getCellFormat(cell)`
  (`update_cell_format`, `get_cell_format`) change and read a cell's number
  format, font, fill, border, alignment and protection. A patch key that is
  omitted keeps the attribute and `null` (`None`) clears it. A whole column
  (`"D:D"`), row (`"3:3"`) or sheet range formats that axis. In Python
  `add_border` takes a record and `add_cell_xf` a `CellXfInput`.
- **Values:** `formulon.toNative(value)` converts a `Value` to a plain
  JavaScript value (`null`, number, boolean or string; an error, array,
  reference or lambda stays a `Value`; Python's
  `formulon.to_native(value)` does the same).
- **Spans:** row and column operations (`setColumnWidth`, `insertRows`,
  `deleteRows` and the others that take a first and last index) also accept a
  span string such as `"B:D"` or `"2:5"`.
- **Enums:** ten enum tables are new (`BorderStyle`, `FillPattern`,
  `FontFamily`, `FontScheme`, `FontVerticalAlign`, `GradientType`,
  `HorizontalAlign`, `ReadingOrder`, `UnderlineStyle`, `VerticalAlign`; Python
  `IntEnum`s of the same names) and the style fields use them. The values are
  unchanged, so a caller passing the number keeps working.
- **JavaScript, Python:** `getCellFormat` reports `source`, where the
  effective format comes from; the XF index stays on the `ooxml` layer
  (`ws.ooxml.getCellXfIndex`). `numberFormat` is the name everywhere
  (`getNumberFormat(id)` / `addNumberFormat(code)` for the table of formats);
  `numFmtId` remains only for the XF-level id. The registries of table styles
  and cell styles (`list`, `get`, `set`, `remove`) are on `Workbook`.
  `activeSheet()` returns the `Sheet` instead of its index, and
  `hasPassthrough(path)` returns a
  boolean. `setDefaultFont` gains `getDefaultFont()`.
- **JavaScript, Python:** a pivot filter states its operand as `value` and
  `valueHigh` (number, string or `null`) instead of `valueKind`, `valueInt`,
  `valueDouble`, `valueText` and the `valueHigh*` keys; the enum
  `PivotFilterValueKind` is removed from both bindings. `valueHigh` is the
  upper bound of a `Between` filter. C keeps the operand union. A pivot cell's
  `value` is a native value (a number, string, boolean or `null`) instead of a
  `Value`; `cell.value.text` is `cell.value`. A pivot cell has no date flag,
  because the layout has none: a date label reads as its serial and the cell's
  `numberFormat` says it is a date, as a sheet cell does. A shared item reads
  as a `PivotCacheValue` with `isDate` (Python `is_date`), and the value of a
  date item is its serial in both bindings.
- **JavaScript, Python:** a date-time input takes the same union everywhere: a
  `CivilTime` or a `Date` (Python `date` or `datetime`). A `Date` is read
  through its local getters (`getFullYear()` to `getSeconds()`, milliseconds
  dropped), as Python reads a naive `datetime`; `setValue(Date)` and
  `setPinnedNow(Date)` used its UTC fields. An Invalid Date throws
  `RangeError`. The slots are `setValue`, `setPinnedNow`,
  `ThreadedCommentInput.created` (read through its UTC getters),
  `PivotCache.setRecordValue`,
  `PivotCache.addFieldSharedItem` (the `{ date: true }` option is gone: a
  `Date` is a date and a number is a number), the `start` and `end` of
  `PivotDateGroupInput`, and `evaluateCfRange(range, { today })` (Python
  `evaluate_cf_range(range, today=None)`; the positional `todaySerial` is
  gone, and an omitted `today` is the pinned clock, else the local date).
  Reads return one type: `ThreadedComment.created` and the group bounds are
  `CivilTime` (Python `datetime`), `null` when the stamp is absent or cannot
  be parsed, and `getDate` returns a `Date` whose local fields carry the
  serial, so `ws.setValue(r, c, ws.getDate(r, c))` round-trips. C keeps
  serials and `fm_civil_time_t`. Python refuses an aware `datetime` with
  `ValueError` in every one of these slots, as a cell write already did,
  except `ThreadedCommentInput.created`, which converts it to UTC.
- **JavaScript, Python:** `evaluateCfFormula` returns a `boolean`, as
  `evaluate_cf_formula` does in Python. `Value.to_python()` is removed;
  `formulon.to_native(value)` is the counterpart of `toNative`. `setError`
  accepts an `ExcelError` as well as an `ErrorCode`, and `validateValue`
  (`validate_value`) accepts the value union of `setValue`.
- **JavaScript, Python:** a setter that takes a record has an `*Input` type
  (`FontInput`, `FillInput`, `BorderInput`, `CellXfInput`, `DxfInput`,
  `CellStyleInput`, `TableStyleInput`, `CommentInput`, `PersonInput`,
  `ArrayFormulaInput`, `WorkbookProtectionInput`, `ProtectedRangeInput`,
  `SparklineInput`, `PrintTitlesInput`, `AutoFilterConditionInput`,
  `ChartFilterInput`, `PivotWorksheetSourceInput` and the others in
  [Added](../../CHANGELOG.md#added)), and also accepts the record its getter returns. An omitted
  key takes the default its type states, and the presence flags are gone:
  `CellXf.hasAlignment`, `hasHorizontalAlign`, `hasVerticalAlign`,
  `hasWrapText`, `hasJustifyLastLine` and `hasProtection`, and
  `PivotWorksheetSource.present`.
  `CfValueObjectInput` is `CfValueObject` and `CfIconInput` is `CfIcon`.
- **JavaScript, Python:** `DataTableInput` now names the data-table spec;
  the former cell-reference record is `DataTableInputCell`. The spec was
  `DataTableSpec` in JavaScript.
- **JavaScript, Python:** a key that holds a cell rectangle is `range`, and one
  that holds several is `ranges`. `rect` is `range` on `ArrayFormula`,
  `Spill`, `ParsedReference`, `DataTable` and
  `PivotWorksheetSource` (which also replaces `ref`), `ref` is `range` on
  `AutoFilterSortCondition` and `AutoFilterSortState`, and `source` is `range`
  on `ChartSourceInput`, `ChartInput` and `ChartPreviewOverridesInput`.
  A sparkline's `source` is a `SheetRect` (it was the text `formula`), and
  `SparklineGroup.dateAxisRef` is `dateAxisRange`, a `SheetRect` or `null`.
  Input takes a `SheetRect` or an A1 range where the type says so.
- **JavaScript, Python:** a cell position is a `CellAddress`:
  `SheetView.topLeft` and `paneTopLeft` replace `topLeftRow` / `topLeftCol` and
  `paneTopLeftRow` / `paneTopLeftCol`, `SheetSelection.active` replaces
  `activeRow` / `activeCol`, and `ThreadedComment.cell` replaces `row` / `col`
  (`ThreadedCommentInput.cell` takes a `CellAddress` or an A1 cell, A1 when
  omitted). `listThreadedComments(a1)` and `listThreadedComments(row, col)`
  list one cell's thread. `setPrintTitles` takes `PrintTitlesInput`: each axis
  is a `Span`, a span string (`'1:2'`, `'B'`) or `null`.
- **JavaScript, Python:** `ChartInsertKind` is removed. `addChart({ type, range,
  barDirection? })` takes a `ChartPlotType` (`Bar`, `Line`, `Area`, `Pie`,
  `Doughnut` or `Scatter`) and `barDirection` (`Column` when omitted) for `Bar`;
  `Chart.setType(plot, type, { barDirection, grouping })` likewise. An explicit
  `barDirection` on another type throws `TypeError`. The
  [ordinal table](#enums) maps the old numbers. `Chart.setFilter(filter)` takes
  `{ hiddenSeries, hiddenCategories }` (Python a `ChartFilter`), and
  `Chart.setSource({ range, direction, headerRow, headerColumn })` takes
  `range` as a `SheetRect` or an A1 range; `ChartProperties.source` is a
  `ChartSource` with `range`. A legend hides entries by `hiddenEntries`
  (`ChartLegend`, `Chart.updateLegend`, Python `update_legend(hidden_entries=)`);
  `hiddenSeries` stays on the chart filter.
- **JavaScript, Python:** an AutoFilter column lists its custom filter as
  `conditions`, one or two `{ operator, value }` records, in place of
  `customCount`, `operator1`, `value1`, `operator2` and `value2`. Any other
  count for a custom column is `InvalidArgument`.
- **JavaScript, Python:** `CfMatch` is a union told apart by `kind`, and each
  member has only its own fields: `m.bar` (`lengthPct`, `axisPositionPct`,
  `isNegative`, `fill`, `border`, `gradient`, `direction`) for a data bar in
  place of the flat `bar*` keys, `m.icon` (`iconSet`, `iconIndex`) for an icon
  set in place of the flat `iconSetName` and `iconIndex`, and `color` for a
  colour scale. `dxfIdEngaged` is gone, and `dxfId` is `dxfIndex`, `null` for a
  rule that applies no format. A custom table style element reads `dxfIndex: null` for a region
  without formatting.
- **JavaScript, Python:** `pivotFieldSetSort` (Python `pivot_field_set_sort`)
  is removed; `Pivot.updateField(i, { sortOrder, sortByField, showDetail })`
  replaces it.
  `PivotCache.getWorksheetSource()` returns `null` when the cache has none and
  otherwise `{ range, name }` with `range` a `SheetRect` or `null`;
  `setWorksheetSource({ range, sheet, name })` takes the same. A table column's
  totals use the keys of `TableColumn`: `totalsRowFunction`, `totalsRowLabel`
  and `totalsRowFormula` replace `function`, `label` and `formula` in the
  elements of `Table.update`'s `columns`; `setTableColumnTotals` and
  `TableColumnTotalsInput` are removed.
- **JavaScript, Python:** `ExternalLink.index` is `bookNumber` (`book_number`),
  the `n` of `[n]` in formula text. `pivotFieldAdd` (`pivot_field_add`) is
  removed: a pivot has one field per cache field, and a field is named by a
  `PivotFieldRef`, the pivot field index or the source name of the cache
  field. A pivot filter states its field as `fieldIndex` instead of `axis` and
  `fieldName`.
- **JavaScript, Python:** `pivotFieldAddSubtotalFn` and
  `pivotFieldClearSubtotalFns` are `Pivot.setFieldSubtotalFunctions(fieldIndex,
  list)` (`set_field_subtotal_functions`), which replaces the set; an empty
  list clears it and a repeated function is `InvalidArgument`.
- **JavaScript, Python:** the tab's visibility leaves `SheetView` and
  `SheetViewPatch`; read it with `getVisibility()` and write it with
  `setVisibility(v)` (`get_visibility`, `set_visibility`).
- **JavaScript, Python:** `addSheet(name?, { insertBefore? })` and
  `copySheet(source, { insertBefore?, newName? })` take an options object
  (Python keyword-only `insert_before=`, `new_name=`). `moveSheet(from, to)` (Python
  `move_sheet`), whose `to` counted positions after the removal, is
  `Sheet.move(insertBefore)`, a position in the list as it is now.
- **JavaScript, Python:** `addValidation` and `addSparklineGroup` return the
  new rule's id (`ValidationId`, `SparklineGroupId`), and
  `removeValidation`, `updateSparklineGroup` and `removeSparklineGroup` take a
  `ValidationRef` or `SparklineGroupRef`: an id, or a position as before.
  `DataValidation` and `SparklineGroup` carry `id`; `ValidationOutcome`
  names the covering rule as `ruleId` in place of `ruleIndex` (`null` when no rule covers the cell).
- **JavaScript, Python:** `ThreadedCommentInput.id` is optional (the engine
  generates one) and `personId` is required. Python `ChartCategories.format_code`
  is `number_format`.
- **JavaScript, Python:** `Table.update({ totalsRow: true })`
  (`update(totals_row=True)`) adds the totals row below the table: the range
  grows by one row, the cells beneath the table's columns shift down, and each
  column's totals cell is written in the new row. It used to overwrite the last
  data row. When something straddles the shifted band the row is taken from
  free space below instead, and the update fails with `PreconditionFailed` and
  changes nothing when no free row is left, or when the table ends on the last
  row of the grid. `update({ totalsRow: false })` removes the totals row and
  shifts the cells beneath up, or only empties it when something straddles the
  columns below; it fails with `PreconditionFailed` while a formula or defined
  name references the row, and the columns keep their saved totals settings. A toggle that changes nothing is a no-op.
  `fm_sheet_update_table` behaves the same.
- **JavaScript, Python:** turning the totals row on writes fresh defaults when
  the table has never shown its totals row or lost it to a row delete: the
  first column gets the label in the profile's language (`Total`, `集計`), the
  last column a sum when every data cell is a non-date number and a count
  otherwise, and a one-column table gets the function only. Otherwise the
  columns' saved settings are restored. A row or cell delete over the totals
  row turns it off and clears every column's totals settings.
- **JavaScript, Python:** `Table.update({ headerRow: false })`
  (`update(header_row=False)`) deletes the header cells, shrinks the range by
  its top row, drops the AutoFilter with its criteria, shows the rows they hid
  and keeps the column names; it used to flip the flag only. While it is off `Tbl[#Headers]` evaluates to `#REF!`.
  `headerRow: true` writes the names into the row above the data and grows the
  range up when that row is empty in the table's columns, and otherwise shifts
  the table's columns down one row as `insertCells` with `CellShift.Down`
  does. Turning it off on a table with no data row fails with
  `PreconditionFailed`.
- **JavaScript, Python:** a patch with `range` and a toggle applies the toggle
  first and checks `range` against the toggled shape; it used to take `range`
  as final and flip the flag only. A `range` whose first row disagrees is
  `InvalidArgument`, and a refused patch changes nothing. Migrate by passing
  the range of the toggled shape (the old first row plus one for header off,
  minus one for header on by growing up, unchanged when the columns shift down)
  or by omitting `range`.
- **JavaScript, Python:** `deleteRows` and `deleteCells` with `Up` over a
  table's header row, without the whole table, fail with `PreconditionFailed`;
  a delete covering every data row keeps the first data row, emptied. A
  one-row `addTable` range with a header row creates the header and one empty
  data row, shifting occupied cells of the table's columns down, and fails
  with `PreconditionFailed` on the last grid row; `headerRow` with `totalsRow`
  on one row is `InvalidArgument`.
- **JavaScript, Python:** `Table.update({ columns: [...] })`
  (`update(columns=[TableColumnPatch(...)])`) replaces the column totals
  setter. Element `i` patches column `i`; each element is a partial
  `TableColumn` with the same names (`totalsRowFunction`, `totalsRowLabel`,
  `totalsRowFormula`, `calculatedFormula`), an omitted key keeps its setting,
  `null` (`None`) clears it, and a read `TableColumn` is accepted as an
  element (its `name` is ignored). An array shorter than the column count
  patches the leading columns, `{}` leaves a column alone, and a longer array
  is `IndexOutOfRange` on every surface (C `FM_STATUS_INDEX_OUT_OF_RANGE`). A `custom` function
  needs a non-empty formula, any other function reads its formula as `null`,
  and an empty string is `InvalidArgument`. `calculatedFormula` is the
  formula as typed in the column's first data row (`=A3*10`, `=[@a]*10`); a
  formula that differs from the column's current one fills every data cell,
  constants included, while the same formula is no change and keeps the
  constants placed in the column, and `null` removes the mark and leaves the
  cells. The patch is checked as a whole before anything changes; an array,
  dynamic-array or off-grid formula is `InvalidArgument` and one that does not
  parse reports its parse error. A malformed element (an unknown key) is a
  `TypeError`. `TableColumn.calculatedFormula` (`calculated_formula`) reads
  the formula in the same spelling, or `null`. In C, `fm_table_patch_t.columns`
  holds `columns_count` `fm_table_column_patch_t` elements, each field with its
  own op, and `fm_table_column_info_t.calculated_formula` reads the formula.
- **JavaScript, Python:** `null` (`None`) for a patch key that the patch type
  does not declare clearable throws `TypeError` in the facade, the raw
  WebAssembly and native layers and Python. `Sheet.updateView({
  showRowColumnHeaders: null })` and `updateFormatDefaults({
  baseColumnWidthChars: null })` used to reach the library and fail with
  `InvalidArgument`, and Python `Chart.update_plot`, `update_axis` (`visible`,
  `major_gridlines`, `reversed`, `crosses`, `crosses_at`) and `update_series`
  (`values`, `y`, `size`, `smooth`) raised `InvalidArgumentError` for `None`.
  Python `Sheet.update_protected_range` takes `None` to clear `algorithm_name`,
  `hash_value`, `salt_value` and `security_descriptor`, as the other bindings
  do.
- **JavaScript, Python:** adds that returned nothing return what they created:
  `addMerge` the stored `CellRect`, `addDataTable` the top-left `CellAddress` of
  the result range, `addPerson` the stored id and `PivotCache.addFieldSharedItem`
  (`add_field_shared_item`) the new item's index.

### JavaScript and TypeScript callers

- **JavaScript:** failures throw. A call that cannot complete throws
  `FormulonError` with a `code` (`StatusCode`), `message` and `context`;
  there are no `{ status, ... }` envelopes and no `.ok` to check. Replace
  `if (!r.status.ok)` with `try` / `catch`. `r.value`, `r.index` and `r.bytes`
  are the return value itself, list results are plain arrays, and a lookup
  with no match returns `null`. An argument of the wrong type throws
  `TypeError`, and an integer outside its range throws `RangeError`.
  `instanceof FormulonError` holds within one package; test
  `e.name === 'FormulonError'` across `@libraz/formulon` and
  `@libraz/formulon-native`.
- **JavaScript:** `createFormulon()` resolves to the `Formulon` object, and
  both packages expose it the same way. `@libraz/formulon-native` no longer
  exports `Workbook` and `evalFormula` at the top level; use
  `const formulon = await createFormulon()` and `formulon.Workbook`. Name the
  object `formulon`, not `Module`. `saveAs(format)` is `save(format)`, and
  `lastErrorMessage()` is the `message` of the caught error.
- **JavaScript:** a value is a union discriminated by `kind`. Compare with
  `ValueKind` (`v.kind === ValueKind.Error`), not with a number. A value has
  only the fields of its kind, and `boolean` is a real `boolean`.
  `Workbook.disposed` reports a released handle, and every other call on it
  throws `FormulonError` with `StatusCode.BindingInvalidHandle`.
- **JavaScript:** `Workbook.sheet(key)` returns the `Sheet` handle of a sheet
  given by index, name (compared without regard to case) or another `Sheet`,
  and every sheet-scoped method moves onto it (see [Sheet handles](#sheet-handles)).
  `setValue` stores a number, string, boolean, `null` or a `Value`
  through the matching typed setter. Result types lose their `Result` suffix
  and update inputs take `Patch<T>`. Enums are const objects with matching
  union types. The embind `delete()`, `clone()` and `isDeleted()` are gone;
  use `dispose()` or `[Symbol.dispose]()`. `Workbook` is a plain class that
  each `createFormulon()` call builds anew, so `wb instanceof Module.Workbook`
  and `wb.$$` no longer work and two calls never share a class.
- **JavaScript:** `canonicalizeFormula` and `localizeFormula` check the grammar
  and throw with `parseError` offsets (UTF-16) into the text passed, as do
  `canonicalize_formula` and `localize_formula` in Python and the CLI `eval`
  command with `--profile`, which prints the position in the text as typed. Malformed text that
  `canonicalizeFormula` used to rewrite is now refused.
- **JavaScript:** `createFormulon()` of `@libraz/formulon-native` looks up the
  prebuild itself. Importing the package no longer throws when none matches;
  `createFormulon()` rejects with a plain `Error` naming the platform, which is
  a host failure and not a `FormulonError`. `FORMULON_NATIVE_ROOT` replaces the
  directory searched.
- **JavaScript:** `recalcParallel(threadCount)`: `threadCount` is optional
  (0 = automatic, capped at 8); a non-number is `TypeError`, a fractional or
  non-finite number `RangeError`, a negative number or one above 8
  `InvalidArgument`. A negative number used to be `RangeError`. A byte input
  that is not an `ArrayBuffer`, `Uint8Array` or view is `TypeError` naming the
  method and parameter, and an enum ordinal outside its table is
  `InvalidArgument` with a message that names the method and the parameter.
- **JavaScript:** `getCellsInRange(range, { cursor, limit })` and
  `getInvalidCells({ cursor, limit })` take a `PageOptions` object instead of
  positional `cursor` and `limit`; a third positional argument is `TypeError`.
  `getColumnWidthPt`, `getRowHeightPt` and `SheetOoxml.getColumnXfIndex` /
  `getRowXfIndex` accept an A1 column or row (`'B'`, `'3'`) as well as an
  index; a span such as `'B:D'` is `InvalidArgument` (Python `col` / `row`
  take an `int` or `str` alike). `setLogSink` receives a `LogEvent`
  (`{ level, event, ...fields }`, with `level` a `LogLevel`) that it owns,
  decoded from the engine's JSON line. `createFormulon(options)` validates its options in both packages and
  rejects with `TypeError` for a non-object, an unknown key, a `locateFile`,
  `print` or `printErr` that is not a function, or an empty or non-binary
  `wasm`; the native package then ignores the values.
- **TypeScript:** `evaluateCfRange(range, options?: CfRangeEvaluationOptions)`,
  `PivotFilterInput.value` and `valueHigh` accept `null` for absent, so a read
  filter passes back to `addFilter`, and a read `TableColumn` passes as an
  element of `Table.update`'s `columns` (its `name` is ignored). `Sheet.addProtectedRange(input)` names its
  parameter `input`.
- **Patches (JavaScript):** `TableUpdateInput` is `TablePatch`, and its `rect`
  is `range` (an A1 string or a `SheetRect`). `TableProperties.headerRowCount`
  and `totalsRowCount` are the booleans `headerRow` and `totalsRow`.
  `ChartDataSource.kind` is the number `ChartSourceKind`, not text: compare
  with `ChartSourceKind.Ref`.
- **Patches (JavaScript):** `updateSparklineGroup` and `updateProtectedRange`
  are partial updates. A field left out keeps its stored value instead of
  returning to its default; pass the field to change it.
- **Patches (JavaScript):** an unknown key in a patch throws `TypeError` at any
  depth, and `null` for a key that is not nullable throws `TypeError` (in
  `Shape.update` it used to mean "do not apply"; leave the key out).
- **Patches (JavaScript):** record keys such as `objectId` and `kind` are
  ignored by `update`, so the record `getProperties()` returned can be passed
  back. `Slicer.update({ name })` fails with `PreconditionFailed`: a slicer's
  name ties its frame to the slicer part.
- **Pivot tables (JavaScript):** `Pivot.addFieldSharedItem`'s third argument is
  an options object. `PivotWorksheetSourceInput` has no `present` key; clear
  the source with `setWorksheetSource(null)`. The parameter of
  `Pivot.setAnchor` is named `range`; its type is unchanged.
- **JavaScript:** `getValue`, `getValues`, `listCells`, `getCellsInRange`,
  `iterRows` and `evaluateFormula` return native values: `null`, a number, a
  boolean, a string, an `ExcelError` or a 2-D array. The tagged `Value` comes
  from `getTypedValue`, `getTypedValues`, `evaluateTypedFormula` and
  `evaluateTypedFormulaArray`. `toNative` is only needed on a `Value` you
  already hold.
- **JavaScript:** the part-level members (`addCellXf`, `listCellXfs`,
  `getAutoFilterXml` and the others in the migration table) live on
  `wb.ooxml` and `ws.ooxml`. A colour in or out is `{ kind, rgb, theme, tint,
  indexed }`, or a `'#RRGGBB'` string on input; the `*Argb`, `fgArgb` and
  `bgArgb` keys are gone, and a shape takes `fill` and `line` (a
  `ChartFillPatch` and a `ChartLinePatch`) in place of `fillRgb` and
  `lineRgb`. A formula that does not parse throws
  `FormulonError` with `parseError` (`code`, `start`, `end`, `line`,
  `column`). `excelProfileId` / `setExcelProfileId` are `getProfileId` /
  `setProfileId`. `parseReference` throws `ParserInvalidReference` instead of
  returning `null`. A call on a disposed workbook reports the closed handle
  before it validates its arguments.
- **JavaScript:** `setValue` and `setValues` read a string that starts with
  `=` as a formula, with the failure of `setFormula`; `'='` alone is a syntax
  error. Enter untrusted text with `setText`, which always stores text.
- **JavaScript:** record types with required keys are separate from their
  patches: `ShapeInput` and `ShapePatch`, and the same for slicers and the
  chart series add. `signature` is the one name for a signature field.
- **JavaScript:** `Patch<T, Clearable>` names the keys that take `null`.
  `null` on any other key is a compile-time error (and a `TypeError` at run
  time, see above); an empty patch changes nothing.
- **JavaScript:** `DefinedName.localSheetId` is `scope` (`null` for workbook
  scope, else the 0-based index of the sheet), and `DefinedNameInput.scope`
  takes a `SheetRef`. C keeps `local_sheet_id`.
- **JavaScript:** `getRangeXfBasesExcluding(range, excluded)` takes a list of
  cells (`CellAddress` or A1) in place of two arrays of rows and columns,
  `remapRangeXfs(range, mapping)` takes a `Map` from old to new index in place
  of two arrays, and `convertCivilToSerial(civil, { date1904 })` and
  `convertSerialToCivil(serial, { date1904 })` take an options object in place
  of a boolean. A `date1904` that is not a boolean throws `TypeError`.
- **JavaScript:** an object passed where a record is expected refuses a key it
  does not define with `TypeError` at every depth. That now covers the options
  of `Pivot.addFieldItem`, the `CivilTime` objects (`setPinnedNow`,
  `convertCivilToSerial`, the `today` option of `evaluateCfRange`, a threaded
  comment's `created`), the date-system options and `Color` objects
  (`setTabColor` and every colour key); a misspelt key used to be ignored.

### Python callers

- **Python:** the handle classes (`Sheet`, `Table`, `Pivot`, `PivotCache`,
  `Chart`, `Picture`, `Shape`, `Slicer`, `SheetOoxml`, `WorkbookOoxml`) cannot
  be constructed directly; they come only from the workbook, and a direct
  constructor call raises `TypeError`.
- **Python:** tuple results are frozen dataclasses (`PivotGrandTotals`,
  `PrintTitles`, `SheetSelection`, `TableStyleDefaults`,
  `IterativeSettings`, `StyleBatchIndices`), as are `Cell` and `Table`.
  `get_display_text` and `format_value` return a `DisplayText`,
  `evaluate_auto_filter` and `evaluate_table_auto_filter` an `AutoFilterEval`,
  and `get_cells_in_range` and `get_invalid_cells` a `CellsInRange`, where they
  returned tuples. `DefinedName` is a frozen dataclass. `get_comment_visible` returns
  `Optional[bool]`, and an optional keyword argument defaults to `None`, which
  leaves the field unchanged. `formulon/__init__.pyi` is removed; the package
  carries inline annotations and `py.typed`.
- **Python:** failures raise `FormulonError` or one of its subclasses:
  `IndexOutOfRangeError` for an index past the end (also an `IndexError`),
  `NotFoundError` for a missing key (also a `KeyError`),
  `InvalidArgumentError` and `AlreadyExistsError` (both also a `ValueError`), and `WorkbookClosedError` for a
  call after `close()`. Code that catches `FormulonError` keeps working.
  `Workbook.closed` replaces `is_valid`, `Workbook.save(format=...)` replaces
  `save_as`, and `formulon.to_native(value)` returns an `ExcelError` for an
  Excel error value. `Workbook.sheet(key)`, `wb[key]`, `len(wb)`, iteration and `in`
  give `Sheet` handles that carry every sheet-scoped method (see
  [Sheet handles](#sheet-handles)). The text of exception messages changed;
  match on the exception class, not the message.
- **Python:** `FormulonError(code, message="", context="", parse_error=None)`
  builds an error, and `str(e)` is `[CODE_NAME] message` with the context only
  in `e.context`. `FormulaSyntaxError` carries `parse_error` (a `ParseError`
  with `code`, `start`, `end`, `line`, `column`). A ragged `set_values` grid is
  `InvalidArgumentError`, and `parse_reference` with only one of `base_row` and
  `base_col` is `TypeError`, as in JavaScript.
- **Python:** `copy.copy`, `copy.deepcopy` and pickling of a `Workbook`, of any
  handle class and of the `*Ooxml` views raise
  `TypeError: Workbook cannot be copied; use Workbook.load(wb.save())`.
- **Python:** parameter names follow the declarations: `set_pinned_now(time)`
  (was `t`), `get_person(id)` / `remove_person(id)`, `add_dxf(input)`,
  `set_dxf(dxf_index, input)`, `set_auto_filter(filter)`,
  `set_cell_phonetic(row, col, phonetic)`, `get_threaded_comment(id)`,
  `update_threaded_comment(id, ...)`, `remove_threaded_comment(id)`,
  `set_row_break(row, input)`, `set_column_break(col, input)`,
  `add_picture(bytes, input=None)`, `restore_drawing(bytes, ...)`,
  `add_shape(spec)`, `add_slicer(table, column, spec)`,
  `add_sparkline_group(spec)`, `add_protected_range(input)`,
  `Chart.set_point(plot, series_id, point_index, point)` and
  `Chart.add_series(plot, input)`. The calculated-item methods of
  `PivotCache` take `calculated_item_index`. `Slicer.update` no longer takes
  `name`, because a slicer's name is fixed by its cache.
  `Sheet.update_cell_format` drops its `source` parameter; copy a format with
  `ws.update_cell_format("B1", **vars(ws.get_cell_format("A1")))`.
  `Sheet.set_row_break` and `set_column_break` take a `PageBreakInput`
  (`first`, `last`, `manual`) or a read break; `list_threaded_comments` takes
  `(row, col)` or one A1 string.
- **Python:** `Shape.update` and `Slicer.update` default every argument to the
  `UNCHANGED` singleton (an enum member of type `Unchanged`). `None` clears
  only in `Slicer.update(style=None)`; elsewhere it is a `TypeError`.
- **Python:** an anchor is one `DrawingAnchorInput` value instead of separate
  `anchor_kind=` / `row=` / ... keyword arguments; pass it as `anchor=` to
  `add_picture`, `add_shape`, `add_slicer`, a chart's `placement` and
  `Drawing.update`.
- **Python:** `Chart.get_preview` takes keyword-only arguments in place of a
  `ChartPreviewOverridesInput` record, which is removed.
- **Python:** the parameter that takes a cell rectangle is named `range`
  (`get_values`, `list_cells`, `add_merge`, `remove_merge`,
  `get_cell_rect_pt`, `get_merges_in_range`, `get_range_xf_bases`,
  `remap_range_xfs`, `snapshot_range_xfs`, `set_range_xf_index`;
  `viewport` in `recalc_viewport`; `Table.update` and `Pivot.set_anchor`
  likewise). Only a call that passes it by keyword breaks.
- **Python:** `TableProperties` carries `header_row` and `totals_row` booleans
  in place of `header_row_count` and `totals_row_count`. `ChartDataSource.kind`
  is the `ChartSourceKind` `IntEnum`, not `"none"` / `"ref"` / `"literal"`.
- **Python:** `update_sparkline_group` and `update_protected_range` take
  keyword arguments and change only those given, as in JavaScript.
  `PivotCache.set_worksheet_source(None)` clears the source and the keywords
  `ref`, `sheet` and `name` set it.
- **Python:** `formulon.LIB` is no longer re-exported from the package. The
  engine functions that used it (`get_version`, `set_log_min_level`) are
  module-level functions of `formulon`.
- **Python:** the part-level members move to `wb.ooxml` and `ws.ooxml`
  (`get_font`, `list_fonts`, `add_cell_xf`, `get_auto_filter_xml` and the
  others), and the `style_xf` field of the column and row layout records is
  read with `ws.ooxml.get_column_xf_index` and `get_row_xf_index`. `ws["A1"]`,
  `ws[row, col]` and `ws["A1:B2"]` read and write native values (the range
  form is a list of rows); `get_value` and `get_values` agree with them, and
  `get_typed_value` / `get_typed_values` return `Value`. `evaluate_formula`
  returns the whole array result. A string that starts with `=` is a formula;
  use `set_text` for untrusted text.
- **Python:** colours are `Color` or `"#RRGGBB"` / `"#AARRGGBB"` text; the
  `*_argb` fields, `CfColor`, and the shape fields `fill_rgb` / `line_rgb`
  (now `fill` / `line`, a `ChartFill` and a `ChartLine`) are removed. `Cell` is `RangeCell`,
  `ExcelProfileId` is `ProfileId`, and `excel_profile_id` /
  `set_excel_profile_id` are `get_profile_id` / `set_profile_id`. A formula
  that does not parse raises `FormulaSyntaxError` (a `FormulonError` with the
  parse position); `parse_reference` raises it instead of returning `None`.
  `save` picks `.xlsx` or `.xlsb` from the file extension and refuses any
  other extension unless `format=` is given. Python 3.9 is no longer
  supported; the package requires 3.10 or later.
- **Python:** eight signatures take a record in place of keywords:
  `set_cell_phonetic_properties(row, col, properties)` takes a
  `PhoneticProperties`, `add_chart(input)` a `ChartInput`,
  `Chart.set_source(source)` a `ChartSourceInput`,
  `PivotCache.set_worksheet_source(source)` a `PivotWorksheetSourceInput` or
  `None`, `Pivot.set_field_date_group(field_index, group)` a
  `PivotDateGroupInput`, `set_selection(selection)` a `SheetSelectionInput`,
  `set_theme_colors(colors)` a `ThemeColorsInput` and
  `set_theme_script_font(font)` a `ThemeScriptFont`. The old keyword forms
  raise `TypeError`.
- **Python:** four record types take the JavaScript names: `ErrorName` is
  `LocaleErrorName`, `MarginsPt` is `PaginationMargins`, `PageLayout` is
  `PaginationPage` and `PaperInfo` is `PaginationPaper`.
  `Workbook.list_passthroughs()` returns
  a list of part names, and `PassthroughPart` is removed. `resolve_color`
  has no default for its context, and `format_value(value, number_format)`
  names its second argument as JavaScript does and requires it.
  `PivotDataFieldInput(field_index, name=None, ...)` takes the field first.
  `Sheet.set_comment` takes a `CommentInput(text, author="")`, as JavaScript
  takes `CommentInput`; a `Comment` read back is still accepted.
- **Python:** `Mention` / `MentionInput` and `PhoneticRun` spans and the
  `ParseError` positions (`start`, `end`, `column`) are code-point offsets into
  the Python `str`, as `len` and slicing count, instead of UTF-16 code units; a
  character outside the BMP counts 1. JavaScript keeps UTF-16 offsets, and
  length limits stay in UTF-16 code units on every surface.
- **Python:** the status ranges have exception classes. `IoError`,
  `EvaluationError`, `FunctionError`, `RecalcError`, `SecurityError`,
  `BindingError`, `PivotError` and `PrintError` are `FormulonError` subclasses,
  one per range, and a status without a narrower class raises the class of its
  range. `InvalidFileError` is an `IoError` and a `ValueError`.
  `FormulaSyntaxError` is a `FormulonError` and a `ValueError` and is no longer
  an `InvalidArgumentError`: code that caught a failed parse with
  `except InvalidArgumentError` must catch `FormulaSyntaxError` (or
  `ValueError`) instead.

### Command-line users

- **Command line:** plain output escapes only newline, carriage return, tab and
  backslash, so a value pastes back as written. `paginate` prints print-area ranges in A1 notation. `eval` with no formula
  exits with status 64, and an engine failure exits with status 1. `dump` and
  `paginate` accept `--json`, and `profiles --json` lists the profile ids as
  JSON; the usage text no longer lists the profile ids. `paginate` prints
  `name`, `pageCount`, `printArea`, `rowBreaks`, `columnBreaks` and `pages`,
  one per line in plain output and as the same keys in `--json`
  (`{"format":1,"sheets":[{"name":...,"pageCount":N,"printArea":[...],
  "rowBreaks":[...],"columnBreaks":[...],"pages":[{"range":...}]}]}`); the
  former `sheet=`, `pages=<count>`, `print_area`, `horizontal_breaks` and
  `vertical_breaks` keys are gone, and a `--sheet` run prints the sheet name
  instead of an index. `dump --view metadata` names its sections and JSON keys
  `definedNames`, `tables` and `passthroughParts` (they were `defined_names` and
  `passthrough_parts`). `dump --json` prints cells as `{"format":1,"sheets":[
  {"name":...,"cells":[{"ref","kind","value","formula"?}]}]}`, and `"format"` is
  the schema version.
- **Command line:** a formula syntax error exits with status 65. `eval -`
  reads the formula from standard input, and `recalc -o` takes the file format
  from the output extension.
- **Command line:** `dump` takes `--view formulas|values|sheets|metadata`
  (default `formulas`; `--view=values` also works); `--formulas`, `--values`,
  `--sheets` and `--metadata` are removed and exit with status 64. `paginate
  --sheet` takes a sheet name or an index and fails for an unknown sheet as
  `dump --sheet` does. `recalc --json` prints
  `{"format":1,"status":"ok","output":PATH,"bytes":N,"warnings":[{"kind":...,"count":N}]}`
  on stdout and puts the loss warnings in `warnings` instead of stderr; it is
  refused with status 64 together with `-o -`.

## Topics

### Sheet handles

A sheet-scoped call moves from the workbook onto a `Sheet` handle, and the
sheet argument disappears:

```js
// before
wb.setNumber(sheet, row, col, 42);
// after
wb.sheet(sheet).setNumber(row, col, 42);
```

```python
# before
wb.set_number(sheet, row, col, 42.0)
# after
wb[sheet].set_number(row, col, 42.0)   # or wb.sheet(sheet)
```

Calls that address the whole workbook (`save`, `recalc`, `addFont`,
`setDefinedName`, `setProfileId` and the others without a sheet
argument) stay on `Workbook`. A handle follows its sheet through `addSheet`,
`copySheet`, `moveSheet`, `renameSheet` and `removeSheet`; `wb.sheet(0)` and
`wb.sheet('Data')` are the same object when they name the same sheet. Using
the handle of a removed sheet throws `FormulonError` with
`StatusCode.NotFound` (Python: `NotFoundError`).

| Before | After |
|---|---|
| `sheetCount()` / `sheet_count()` (a method) | the property `sheetCount` / `len(wb)`; `hasSheet(ref)` / `ref in wb` tells whether a sheet exists |
| `sheetName(i)` / `sheet_name(i)` | `sheet(i).name` |
| `addSheet(name)` | `addSheet(name?, { insertBefore? })` returns the new `Sheet`; `copySheet(source, { insertBefore?, newName? })` likewise (Python keyword-only) |
| `activeSheet()` returning an index | `activeSheet()` returns a `Sheet`; `setActiveSheet`, `moveSheet`, `removeSheet` and `renameSheet` take a `Sheet`, an index or a name |

`sheet(name).index` finds a sheet by name and throws when there is no such
sheet. `ws.pivot(pivot).showDetails(row, col)` returns the `Sheet` it adds.

Every sheet method that takes a cell accepts an A1 address in place of
`row, col` (see [A1 addresses](#a1-addresses-and-block-access)). Positional guard messages read
`Sheet.setNumber argument 1 must be ...` and `Workbook.<name> argument N ...`
in `@libraz/formulon` and `@libraz/formulon-native` alike.

### Method shape changes

Names are JavaScript; Python uses the snake_case equivalents, and a record
argument is a keyword-only argument list there (see
[Keyword-only arguments](#keyword-only-arguments-python)).

| Before | After |
|---|---|
| `setIterative(enabled, maxIterations, maxChange)` | `updateIterative(patch)`; an omitted field keeps the stored value |
| `setPinnedNow(y, m, d, h, mi, s)` and `clearPinnedNow()` | `setPinnedNow(time)` with a `CivilTime` or a `Date` read through its local getters; `null` clears the pin. Python takes and returns a naive `datetime.datetime` (`set_pinned_now(None)` clears it) |
| `setSheetPrintTitles(sheet, repeatRows, repeatCols)` / `getSheetPrintTitles(sheet)` | `ws.setPrintTitles(titles)` / `ws.getPrintTitles()` with `PrintTitles = { rows: Span \| null, columns: Span \| null }` and `Span = { first, last }` (0-based, inclusive); `null` clears. `Pagination.printTitles` has the same type |
| `pivotSetGrandTotals(sheet, pivot, rows, cols)` | `ws.pivot(pivot).updateGrandTotals({ rows, columns })` |
| `setComment(sheet, row, col, author, text)` (empty text removed the comment) | `ws.setComment(row, col, { text, author? })` writes the comment, empty text included (an omitted `author` means none); `ws.setComment(row, col, null)` removes it and does nothing when the cell has none |
| `setWorkbookProtection(p)` / `setSheetProtection(sheet, p)` (an all-zero record cleared them) | `wb.setWorkbookProtection(null)` and `ws.setSheetProtection(null)` remove them; an all-zero workbook-protection record still does. Sheet protection is described under [Behaviour changes](#behaviour-changes-under-an-unchanged-name) |
| `setCellPhoneticRuns(sheet, row, col, runs)` | `ws.setCellPhoneticRuns(row, col, runs)`; `null` and `[]` clear |
| `addHyperlink(sheet, row, col, target, display, tooltip, location)` | `ws.setHyperlink(row, col, { target, location?, display?, tooltip? })` adds a link at the cell, replacing one only when it covers exactly the same range, and `null` removes every link anchored there; a `null` target is an in-workbook link and requires `location`. The range the link covers is `range` in the record |
| `setThemeScriptFont(major, script, typeface)` | `setThemeScriptFont({ scheme, script, typeface })`, with `scheme` a `FontScheme` (`Major` or `Minor`) |
| `editThreadedComment(sheet, id, text, mentions)` | `ws.updateThreadedComment(id, { text?, mentions? })` |
| `addThreadedComment(sheet, partial)` | `ws.addThreadedComment(input)` with `ThreadedCommentInput` |
| `setDefinedName`, `setDefinedNameScoped`, `setDefinedNameAttrs` | `setDefinedName({ name, formula, scope?, comment?, hidden? })`; `scope` is a sheet index or name, and `null` or omitted is workbook scope. `removeDefinedName(name, scope?)` takes the same. Setting an existing name replaces every attribute |
| `getTheme(): Theme` (fonts and colors) | `getThemeFonts(): ThemeFonts` and `getThemeColors(): ThemeColors`; the `Theme` type is gone |
| `setSheetViewMode(sheet, mode: string)` | `ws.updateView({ viewMode })` with a `SheetViewMode` |
| `setSheetTabHidden(sheet, hidden)` | removed; `ws.setVisibility(visibility)` with a `SheetVisibility` covers every tab state, very hidden included, and `ws.getVisibility()` reads it. `SheetView` carries no tab state, and `SheetView.viewMode` is a `SheetViewMode` |
| `setTableStyleDefaults`, `setSheetFormatDefaults` taking an inline object | `updateTableStyleDefaults`, `ws.updateFormatDefaults` take a `...Patch` record: an omitted key keeps the value, `null` clears a nullable one |
| `pivotFieldSetAxis(..., axis \| -1)` | `pivot.setFieldAxis(field, axis: PivotAxis \| null)` |
| `pivotCacheCreate(requestedId = 0)` | `addPivotCache(requestedId?: number)` (omitted picks one) |
| `pivotFieldSetNumberFormat(..., utf8)` | `pivot.updateField(field, { numberFormat })` takes a format code |
| `set*Xml(sheet, '')`, `setCellPhonetic(..., '')`, empty captions, custom names, number formats and style names clearing a value | pass `null`; the matching getters return `null` instead of `''` |
| `setSheetPrintArea(sheet, '')` | `ws.setPrintArea(null)` and `ws.setPrintArea([])` both clear |
| `setSheetPageSetup`, `setSheetPageMargins`, `setSheetPrintOptions`, `setSheetHeaderFooter` | `ws.updatePageSetup`, `updatePageMargins`, `updatePrintOptions`, `updateHeaderFooter`; each merges the keys of the patch |
| `pivotGetStyleInfo`, `pivotSetStyleInfo` | `ws.pivot(p).getStyleOptions()`, `updateStyleOptions(patch)`; `Table` has the same pair |
| `setChartAxis`, `setChartDisplay`, `setChartSeries`, `setChartPlot` | `ws.chart(id).updateAxis`, `updateDisplay`, `updateSeries`, `updatePlot` |
| `setChartLegend(sheet, id, position, overlay, hiddenSeries)`, `setChartPoint(..., patch)` | `updateLegend(patch)` and `updatePoint(..., patch)` merge; `setLegend(legend \| null)` and `setPoint(plot, seriesId, point, point \| null)` replace the record or remove it. Python `ch.set_legend(None)` removes the legend |
| `clearRowHeight`, `clearColumnWidth`, `clearCellXfIndex`, `clearColumnStyle`, `clearRowStyle` | `resetRowHeight`, `resetColumnWidth`, `resetCellXfIndex`, `resetColumnXfIndex`, `resetRowXfIndex` |
| `setColumnStyle`, `setRowStyle` | `setColumnXfIndex`, `setRowXfIndex` |
| `clearPinnedNow()`, `removeAutoFilter(sheet)`, `removeSheetRowBreak(sheet, row)`, `removeSheetColBreak(sheet, col)` | `setPinnedNow(null)`, `setAutoFilter(null)`, `setRowBreak(row, null)`, `setColumnBreak(col, null)` |
| `clearAutoFilter(sheet)`, `clearTableAutoFilter(table)` | `ws.clearAutoFilterCriteria()`, `Table.clearAutoFilterCriteria()` |
| `createTable(input)`, `removeTable(index)` with a workbook-wide table index | `ws.addTable(input)`, `ws.table(ref)`, `ws.listTables()`, `ws.removeTable(ref)`; `wb.table(name)` finds a table by name on any sheet |
| `setSheetProtection(p)` with `enabled` | `setSheetProtection(input \| null)`; `updateSheetProtection(patch)` changes single flags |
| `addSheetRowBreak`, `addSheetColBreak` | `setRowBreak(row, { first?, last?, manual? })`, `setColumnBreak(col, ...)` |
| `pivotFieldAddItemAt(sheet, pivot, field, cacheIndex, visible)` | `Pivot.addFieldItem(field, { cacheIndex }, options)` |

A chart without a title or legend reads `title: null` and `legend: null`
instead of a record whose `kind` or `position` is `None`; a patch with
`title: null` or `legend: null` removes it. `PivotGrandTotals.cols` is
`columns`, `Pagination.horizontalBreaks` is `rowBreaks` (and the other
renames in the table below).

A defined name with an empty formula is rejected with `InvalidArgument`
instead of removing the name; `removeDefinedName` removes it.

### Enums

A member that stood for "absent" is removed from its table, and the field or
argument is `null` (Python `None`) instead. Reading the removed member gives
`undefined` in JavaScript, and setting it throws `TypeError`; Python raises
`AttributeError`.

| Enum | Removed member | Replaced by |
|---|---|---|
| `ChartAxisCrosses` | `Absent` | `null` |
| `BlanksAs` (was `ChartBlanksAs`) | `Absent` | `null` |
| `ChartLegendPosition` | `None` | `null` (no legend) |
| `ChartTitleKind` | `None` | `null` (no title) |

`ChartRole` gains `Title` (6) and `AxisTitle` (7), and `unresolved.source` is
a `ChartRole` (C `int32_t` `fm_chart_role_t`) instead of a string. `PivotFilterType`
is `PivotFilterKind`, and its relative-period members are spelled as
`DatePeriod` and `DynamicFilterType` spell them (`Tomorrow` to `LastMonth`,
`NextQuarter` to `YearToDate`, `Q1` to `Q4`, `M1` to `M12`; C
`FM_PIVOT_FILTER_KIND_TODAY`, Python `TODAY`). Only the comparison kinds
(`DateEqual`, `DateBetween`, `DateOlderThan` and the rest) keep the `Date`
prefix. `ChartPlotType.Extended` marks the chartex plots, and
the numeric enum `ChartExtendedLayout` (`BoxWhisker` 0, `ClusteredColumn` 1,
`Funnel` 2, `ParetoLine` 3, `RegionMap` 4, `Sunburst` 5, `Treemap` 6,
`Waterfall` 7) is read from `extendedLayout` (`ChartExtendedLayout | null`),
with `isExtended`; Python has the same in snake_case.

A closed set that has a default in the file format puts that default at 0 and
lists the rest in the order of the format's schema. Where one concept had a
table per feature, one table now serves them all. The numbers moved, so code
that writes the numbers instead of the names changes meaning:

| Enum | Before | After |
|---|---|---|
| `AnchorKind` | `OneCell` 0, `TwoCell` 1, `Absolute` 2 | `Default` 0, `OneCell` 1, `TwoCell` 2, `Absolute` 3 |
| `AnchorEditAs` | `TwoCell` 0, `OneCell` 1, `Absolute` 2 | `Default` 0, `OneCell` 1, `TwoCell` 2, `Absolute` 3 |
| `ChartAxisCrosses` | `AutoZero` 1, `Min` 2, `Max` 3, `At` 4 | `AutoZero` 0, `Min` 1, `Max` 2, `At` 3 |
| `ChartLegendPosition` | `Bottom` 1, `Top` 2, `Left` 3, `Right` 4, `TopRight` 5 | `Right` 0, `Bottom` 1, `TopRight` 2, `Left` 3, `Top` 4 |
| `BlanksAs` for a chart's `blanksAs` (was `ChartBlanksAs`) | `Gap` 1, `Zero` 2, `Span` 3 | `Zero` 0, `Gap` 1, `Span` 2 |
| `BlanksAs` for a sparkline group's `emptyCellsAs` (was `SparklineEmptyCells`) | `Gap` 0, `Zero` 1, `Span` 2 | `Zero` 0, `Gap` 1, `Span` 2 |
| `PivotSortOrder` | `Descending` 0, `Ascending` 1, `Manual` 2 | `Manual` 0, `Ascending` 1, `Descending` 2 |
| `ComparisonOperator` for a `cellIs` rule's `operator` (was a bare number) | `LessThan` 0, `LessThanOrEqual` 1, `Equal` 2, `NotEqual` 3, `GreaterThanOrEqual` 4, `GreaterThan` 5, `Between` 6, `NotBetween` 7 | `Between` 0, `NotBetween` 1, `Equal` 2, `NotEqual` 3, `LessThan` 4, `LessThanOrEqual` 5, `GreaterThan` 6, `GreaterThanOrEqual` 7 |
| `ComparisonOperator` for a validation's `operator` (was a bare number) | `Between` 0, `NotBetween` 1, `Equal` 2, `NotEqual` 3, `GreaterThan` 4, `LessThan` 5, `GreaterThanOrEqual` 6, `LessThanOrEqual` 7 | `Between` 0, `NotBetween` 1, `Equal` 2, `NotEqual` 3, `LessThan` 4, `LessThanOrEqual` 5, `GreaterThan` 6, `GreaterThanOrEqual` 7 |
| `ComparisonOperator` for an AutoFilter condition (was `FilterOperator`) | `Equal` 0, `LessThan` 1, `LessThanOrEqual` 2, `NotEqual` 3, `GreaterThanOrEqual` 4, `GreaterThan` 5 | `Equal` 2, `NotEqual` 3, `LessThan` 4, `LessThanOrEqual` 5, `GreaterThan` 6, `GreaterThanOrEqual` 7 |
| `DatePeriod` for a rule's `timePeriod` (was a bare number) | `Today` 0, `Yesterday` 1, `Tomorrow` 2, `Last7Days` 3, `ThisWeek` 4, `LastWeek` 5, `NextWeek` 6, `ThisMonth` 7, `LastMonth` 8, `NextMonth` 9 | `Today` 0, `Yesterday` 1, `Tomorrow` 2, `Last7Days` 3, `ThisMonth` 4, `LastMonth` 5, `NextMonth` 6, `ThisWeek` 7, `LastWeek` 8, `NextWeek` 9 |

`SlicerSortOrder` keeps `Ascending` 0 and `Descending` 1. A custom AutoFilter
condition always states its operator and takes only the six members that are
not ranges; `Between` and `NotBetween` there throw `FormulonError` with
`StatusCode.InvalidArgument`. `descending` stays a boolean on sort conditions.

Strings that named a closed set are numbers too. A string passed to one of
these fields throws `TypeError`:

| Enum | Field | Before | After |
|---|---|---|---|
| `TableTotalsRowFunction` | `TableColumn.totalsRowFunction`, `TableColumnPatch.totalsRowFunction` | `'none'`, `'sum'`, `'min'`, `'max'`, `'average'`, `'count'`, `'countNums'`, `'stdDev'`, `'var'`, `'custom'` | `None` 0, `Sum` 1, `Min` 2, `Max` 3, `Average` 4, `Count` 5, `CountNumbers` 6, `StdDev` 7, `Var` 8, `Custom` 9 |
| `TableStyleElementType` | `TableStyleElement.type` | the OOXML text (`'wholeTable'`, `'headerRow'`, `'firstRowStripe'` and the rest) | `WholeTable` 0, `HeaderRow` 1, `TotalRow` 2, `FirstColumn` 3, `LastColumn` 4, `FirstRowStripe` 5, `SecondRowStripe` 6 and so on in the order of the OOXML schema, to `PageFieldValues` 27 |
| `ChartTickLabelPosition` | `ChartAxis.tickLabelPosition` and its patch | `'nextTo'`, `'high'`, `'low'`, `'none'` | `NextTo` 0, `High` 1, `Low` 2, `None` 3 |
| `ChartDomain` | `ChartData.domain` | `'category'`, `'xy'` | `Category` 0, `Xy` 1 |
| `WidthPlatform` | `WidthModel.platform` | `'win'`, `'mac'` | `Win` 0, `Mac` 1 |
| `IconSetName` | `CfIcon.iconSet` | the OOXML text (`'3Arrows'`, `'3Stars'` and the rest) | the existing numbers 0 to 16, and `ThreeStars` 17, `ThreeTriangles` 18, `FiveBoxes` 19, `NoIcons` 20 |

Only a custom icon level takes the four sets above 16; a rule's own `iconSet`
still rejects them with `InvalidArgument`. An absent tick-label position is
still `null`, and `ChartTickLabelPosition.None` means no labels are drawn.
Python has the same tables as `IntEnum`s (`TableTotalsRowFunction` was a
`Literal` alias).

`ChartInsertKind` is removed with `addChart` and `Chart.setType` taking a
`ChartPlotType` and a `ChartBarDirection`. A stored number maps as follows:

| `ChartInsertKind` | Replaced by |
|---|---|
| `Column` 0 | `ChartPlotType.Bar` 0 with `barDirection: ChartBarDirection.Column` |
| `Bar` 1 | `ChartPlotType.Bar` 0 with `barDirection: ChartBarDirection.Bar` |
| `Line` 2 | `ChartPlotType.Line` 2 |
| `Area` 3 | `ChartPlotType.Area` 4 |
| `Pie` 4 | `ChartPlotType.Pie` 6 |
| `Scatter` 5 | `ChartPlotType.Scatter` 10 |
| `Doughnut` 6 | `ChartPlotType.Doughnut` 8 |

`ChartLegendPosition` now shares `Bottom`, `Top`, `Left` and `Right` with the
new `ChartAxisPosition`, which replaces the strings `'b' | 't' | 'l' | 'r'` of
`ChartAxis.position`.

Closed value sets that were plain numbers or strings are typed:
`PageOrientation` (`Portrait` 0, `Landscape` 1; the former `Orientation` had
`Default` 0, `Portrait` 1, `Landscape` 2, and `Default` is now an absent
`orientation` that prints portrait), `SheetViewMode` (replaces the
`view` attribute strings `''`, `pageBreakPreview` and `pageLayout`), `ExternalLinkKind`,
`ChartBarDirection`, `CfDataBarDirection` and `CfDataBarAxisPosition`. An
enum argument outside its set throws `FormulonError` with
`StatusCode.InvalidArgument`; `-1`, a non-integer and `NaN` throw
`RangeError`, and a string throws `TypeError` (Python: `InvalidArgumentError`
for an out-of-range number, `TypeError` for the others).

### Result fields

A value that is not there reads `null` (Python `None`) instead of a sentinel.
A comparison such as `field === -1` is silently false now.

| Before | After |
|---|---|
| `-1` | `null`: `PivotField.axis`, `PivotField.cacheFieldIndex`, `PivotField.groupBaseIndex`, `PivotItem.cacheIndex`, `PivotDataField.showAsBaseField`, `PivotDataField.showAsBaseItem`, `PivotCacheOptions.missingItemsLimit`, `DefinedName.scope`, `ValidationOutcome.ruleId` |
| `''` | `null`: `AutoFilterColumn.valueIso` / `maxValueIso`, `AutoFilterSortCondition.customList`, `ChartCategories.numberFormat`, `Comment.author`, `DefinedName.comment`, `Hyperlink.location`, `Hyperlink.display`, `Hyperlink.tooltip`, `PivotCell.numberFormat`, `PivotDataField.numberFormat`, `PivotField.customName`, `PivotField.numberFormat`, `PivotOptions` captions and alternative text, `PivotWorksheetSource.name`, the password, algorithm, hash and salt fields of `SheetProtection` and `WorkbookProtection`, `SlicerProperties.styleName`, `SparklineGroup.dateAxisRange`, `TableStyleDefaults.tableStyleName` / `pivotStyleName` |
| a number or boolean next to a `has...` flag | a nullable field, and the flag is gone: `AutoFilterColumn.filterValue` / `dynamicValue` / `dynamicMaxValue` / `iconIndex`, `AutoFilterSortCondition.dxfIndex` / `iconIndex`, `CellXf` alignment, wrap, locked and hidden fields, `ColumnLayout.widthChars`, `RowLayout.heightPt`, `SheetFormatDefaults.defaultColumnWidthChars` / `defaultRowHeightPt`, `PageSetup.orientation`, `ChartInput.legend` |

`Table.hasStyleInfo` and `CellXf.hasAlignment` are removed; an absent value is
the field being `null`.

`visible` is the one field name on both the read record and the patch.
A data validation spells the drop-down flag `showInCellDropdown` (Python
`show_in_cell_dropdown`, C `fm_tristate_t show_in_cell_dropdown`), the opposite
of the file's `showDropDown`. `PivotOptions.showEmptyCol` is `showEmptyColumn`,
`DataTable.colInput` is `columnInput`, and an AutoFilter top-10 pair is told
apart as `topValue` and `filterValue`. A `Pagination` carries `rowBreaks` and
`columnBreaks` as records, and its margins are `*In` and its scale `scalePct`.

### Keyword-only arguments (Python)

These parameters can no longer be passed by position.

| Method | Keyword-only parameters |
|---|---|
| `HyperlinkInput` | `target`, `location`, `display`, `tooltip`, `range` (a dataclass passed to `Sheet.set_hyperlink`) |
| `Sheet.add_picture` | `name`, `description`, `anchor` |
| `Sheet.add_shape`, `Shape.update` | `anchor`, `preset_geometry` and the other shape fields |
| `Sheet.add_slicer`, `Slicer.update` | `anchor`, `row_height_pt` and the other slicer fields |
| `Chart.update_legend` | `position`, `overlay`, `hidden_entries` |
| `Pivot.update_grand_totals` | `rows`, `columns` |
| `Pivot.add_field_item` | `cache_index`, `visible` |
| `Sheet.update_threaded_comment` | `text`, `mentions`, `done` |
| `TableInput` | `range`, `name`, `columns`, `style_name`, `header_row`, `totals_row` and the style toggles (a dataclass passed to `Sheet.add_table`) |
| `Table.update` | `range`, `header_row`, `totals_row`, `columns` |
| `Table.update_style_options` | `style_name` and the style toggles |
| `Workbook.update_table_style_defaults` | `table_style_name`, `pivot_style_name` |
| `Workbook.update_iterative` | `enabled`, `max_iterations`, `max_change` |
| `Workbook.get_theme_script_font` | `scheme`, `script` |

`Chart.set_source`, `Sheet.set_selection`,
`Sheet.add_chart` and the other signatures listed under BREAKING take one
record instead of keywords. `Sheet.add_table` takes the fields of the JavaScript `TableInput` (`range` is a
`SheetRect`, `columns` the column names); `Sheet.set_selection` takes a
`SheetSelection`; `get_border` returns and `add_border` takes a `Border`
dataclass. `PrintTitles` keeps its name but holds `rows` and `columns` as
`Span | None` instead of `has_rows`, `first_row`, `last_row`, `has_cols` and so on. `wb.sheet(-1)`
raises `IndexOutOfRangeError`. `formulon.UNCHANGED` is a public name, and
`formulon.__all__` lists public names only.

### Command line

- Exit status: `0` success, `1` the engine refused the request (an empty or
  corrupt workbook included), `64` usage error, `66` an input file that is
  missing, unreadable or a directory, `73` an output file that cannot be
  created, `74` an I/O error while reading or writing a stream (a closed pipe
  on stdout included). Failures to open a file reported `1` before.
- `eval --json` prints `{"format":1,"rows":[[{"kind":...,"value":...}]]}` for
  every result, a single cell included; it printed a bare value object for one
  cell and a bare array of arrays otherwise. Every `--json` output that writes
  to stdout (`eval`, `dump`, `paginate`, `profiles`) is an object with
  `"format": 1`, and `dump --json` may now combine with `--view sheets` and
  `--view metadata`.
- An input named `-` reads the workbook from stdin, and `-o -` writes the
  result to stdout; both are refused with status 64 on a terminal. A file that
  is really named `-` is written `./-`.
- Options take `--opt value` and `--opt=value`, `--opt=` is an empty value, a
  flag that takes no value rejects `=value` with status 64, and everything
  after `--` is a positional argument.

### A1 addresses and block access

- **JavaScript, Python:** every method with a cell argument takes either
  `(row, col)` or an A1 string (`ws.setValue('A1', 42)`, `ws.getValue('A2')`;
  Python `ws.set_value("A1", 42)`), and every method with a rectangle argument
  takes `'A1:B2'` as well as a `CellRect`. A string with a sheet prefix, a
  range where a single cell is expected, or text that is not an address throws
  `InvalidArgument`. `Sheet.get` and `Sheet.set` are removed; Python
  `ws["A1"]` and `ws[row, col]` read and write native values. Python
  coordinates are positional only (`row=` and `col=` are no longer accepted).
- **JavaScript, Python:** `getValues(range)` returns the rows of a block as a
  dense array with `Blank` for empty cells, and `setValues(a1, values)` writes
  one with the `setValue` rules (`=` text stays text). Input is validated
  before any cell is written; a write the engine refuses part-way leaves the
  earlier cells in place.
- **JavaScript, Python:** `hasDirtyCells()` / `has_dirty_cells()` reports edits
  made since the last full `recalc`. Reads return the value as of that
  recalculation.
- **JavaScript, Python:** every collection is listed by `list*`, and the
  `get...Count` / `get...At` pairs are removed (`listCells`, `listPivots`,
  `PivotCache.listFields`, `PivotCache.listFieldSharedItems`, `listFonts`,
  `listFills`, `listBorders`, `listCellXfs`, `listCellStyleXfs`, `listDxfs`,
  `Pivot.listFieldItems`). Only `PivotCache.getRecordCount` remains.
- **JavaScript, Python:** a partial update is one patch record whose omitted
  keys keep their value: `updateView`, `Pivot.updateField`, `Pivot.updateFieldItem`,
  `Pivot.updateStyleOptions`, `updatePrintOptions`, `updateHeaderFooter` and
  `Chart.updateLegend` replace the per-property setters. `null` clears a nullable
  key. Python takes keyword-only arguments defaulting to `UNCHANGED`.
- **JavaScript, Python:** tables, pivots, conditional-format rules,
  validations and protected ranges are addressed by position or by key
  (`removeTable('Sales')`, `removePivot('Pivot1')`): a table by its name
  (matched case-insensitively), a rule or validation by its id and a protected
  range by its name.
- **JavaScript:** the state-free functions (`getFunctionMetadata`,
  `listFunctionNames`, `localizeFunctionName`, `canonicalizeFunctionName`,
  `localizeFormula`, `canonicalizeFormula`, `getLocaleFacts`,
  `compareText`, `getCollationClasses`,
  `getImageInfo`) are members of `Formulon`, not of `Workbook`.
  `fm_workbook_probe_image` is `fm_get_image_info` in C, and Python has
  `get_image_info` at module level.
- **JavaScript, Python:** `errorDisplayName` is replaced by the
  `ErrorDisplayName` table (`ErrorDisplayName[ErrorCode.Div0]`) and Python
  `ErrorCode.display_name`.
- **JavaScript, Python:** `PivotCache.addFieldSharedItem(field, value)` and
  `PivotCache.setRecordValue(...)` replace the twelve typed
  writers and dispatch on the type of `value`.
- **Python:** an explicit `save` `format` that contradicts a `.xlsx` or `.xlsb`
  extension raises `InvalidArgumentError`. `save_with_diagnostics` takes `dest`. `add_batch`
  and `has_phonetic_properties` are removed (`PhoneticProperties.stated`
  replaces the last).
- **JavaScript, Python:** `removeMergeAt` is removed in favour of
  `removeMerge(range)`. `removeHyperlink(cell)` is removed in favour of
  `setHyperlink(cell, null)`, which clears every link anchored at the cell.
  `removeHyperlinkAt(index)` and `insertHyperlinkAt(index, link)` are public,
  where `index` is the position in `listHyperlinks()` order. `getDefinedName` and
  `removeDefinedName` take an optional `scope`, `getThemeScriptFont` and
  `setThemeScriptFont` take `{ scheme, script }`, `Chart.setType` takes
  `{ barDirection, grouping }` options, and `Pivot.addFieldItem` takes an
  `options` object.

## Behaviour changes under an unchanged name

- Formula text read back carries a leading `=` where Excel shows one:
  `getFormulaR1C1`, `PivotCacheField.formula`, table totals-row formulas,
  calculated-item formulas, conditional-format and validation `formula1` /
  `formula2`, formula thresholds of colour scales, data bars and icon sets,
  array-formula text and `getLambdaText`. C reads carry it too. Ranges (chart titles and series,
  print areas) stay without it. Input accepts text with or without the `=`,
  in C as well.
- A table column carries a calculated-column formula when the last data row
  holds a formula and at least all but one of the last five data rows hold the
  same formula moved to their row; a new table, a resize and a loaded file
  recognise it. It survives a one-cell write and goes when the last non-blank
  cell of the column is cleared. Writing a value, a formula or a blank to a
  visible totals-row cell updates the column's totals settings to match: a
  value becomes the label (a number is stored as its text), a `SUBTOTAL` over
  the column with one of the eight built-in codes becomes that function, any
  other formula, a constant `=5` included, is custom, and a blank clears the
  setting.
  A loaded file's totals-row cells are read the same way.
- A sparkline group reads an absent `displayEmptyCellsAs` as zero, the schema
  default, and leaves zero implicit on write. It used to read as gap and leave
  gap implicit, so a file without the attribute now draws blanks as zero, as
  Excel does, and a gap group is written with `displayEmptyCellsAs="gap"`.
- `removeDrawing`, `Chart.getData` and `Chart.getPreview` report
  `NotFound` for a missing id instead of `InvalidArgument`. So does every
  pivot-cache lookup by `cacheId`, which makes a missing id `NotFound` on
  every function.
- A cell coordinate, row index or column index outside the sheet grid is
  `InvalidArgument` on every function that takes one. The cell setters,
  the phonetic, lambda-text, spill, hyperlink, comment, data-table, style
  and trace getters, `validateValue` and `Pivot.showDetails` used to accept
  one, and either reported nothing found or failed further in. Removing a
  page break outside the grid was `NotFound`.
- `setAutoFilter` takes an `AutoFilterInput`. Only `range`, each column's
  `colId` and each sort condition's `ref` are required, and the record
  `getAutoFilter` returns is still accepted. An omitted `top` on a top-10
  column reads as `true` (the top items), as in Python and the file format.
  It used to read as the bottom items.
- `addCfRule` takes a brace-wrapped GUID as `id`; an id the sheet already uses
  is replaced by a generated one, and text that is not a GUID is
  `InvalidArgument`.
- `addSheet()` without a name (Python `add_sheet()`, C `NULL`) takes the
  profile's first sheet name without its trailing `1` and the smallest free
  number from 2, ignoring case: `Sheet2`, `Tabelle2`, `Feuil2`.
- A threaded comment without an `id` gets a generated GUID and one without
  `created` the pinned clock, else now; both are written to the file (the
  `id` and the `dT` attribute). `addThreadedComment` returns the id.
- A slicer's fallback shape in the drawing carries the frame's id and name,
  as Excel writes them.
- A pivot item's `name` reads its label: the label it was added with, else its
  shared item's value. An item bound only by cache index read `null` until the
  workbook was saved and loaded again; it now reads the label at once. `null`
  is left for the blank item and for an item bound to nothing.
- The anchor of a drawing object merges per key: `Drawing.update({ anchor })`
  keeps every key it leaves out, and an omitted `kind` keeps the object's
  kind. A kind left out of an add is the operation's default (one-cell for a
  picture, two-cell for a chart, shape and slicer). A read anchor can be passed
  back to `update`: a two-cell anchor takes its box and ignores the size, a
  one-cell anchor takes the size.
- A shape without a fill or outline of its own reads `fill: null` /
  `line: null` (it draws its style's paint); omitting the key on add means the
  same, and a patch of `null` removes the shape's own paint. `ChartFill` and
  `ChartLine` carry `none` (explicit no-fill / no-line) and `other` (a
  gradient, pattern or picture the record cannot state; read only).
- A straight connector may have a zero width or a zero height.
- `Slicer.update` and `Shape.update` accept the common keys (`name`,
  `description`, `hidden`, `anchor`) as well, and apply them before their own.
- A validation or sparkline group is addressed by its id, which stays valid
  across structural edits; ids are opaque strings and are assigned again
  when a workbook is loaded or a sheet is copied. A number is still accepted
  and means a position (`IndexOutOfRange` past the end); an unknown id is
  `NotFound`.
- `updateValidation` and `updateCfRule` accept the record a read returned. A
  change of the rule's type is `InvalidArgument`.
- `Sheet.move(insertBefore)` takes the position in the list before the move
  (`null` appends; a move before itself or before the next sheet changes
  nothing), as `addSheet` and `copySheet` do.
- `getPrintArea` returns `null` (Python `None`) when no print area is set,
  where JavaScript returned `[]`.
- `convertColumnCharsToPt`, `convertColumnPtToChars` and `getColumnWidthPt`
  take `mode` as the last argument. A call with the old order throws
  `RangeError`, except that a `mode` of 0 or 1 in the first position can be
  misread.
- Python pivot options read an unset string as `None` instead of `""`.
- An unknown error code has no display name: `ErrorDisplayName[code]` is
  `undefined` in JavaScript and `ErrorCode(code)` raises `ValueError` in
  Python, where the old function returned `#UNKNOWN!`.
- `updateHeaderFooter` clears a text field for `null` as well as `""`.
- Every JavaScript patch method (`updatePageSetup`, `updatePageMargins`,
  `updateIterative`, `updateView`, `Pivot.updateField` and the rest) takes the record
  its getter returns, ignoring the fields it cannot set, and throws
  `TypeError` for any other key, and so does every setter that takes an input
  record (`setComment`, `addFont`, `setSelection`, `addChart` and the rest).
  Several ignored every unknown key, and `Pivot.updateField` refused `axis`.
- CLI `dump --view metadata` prints defined-name formulas with a leading `=`.
- A drawing loaded from `.xlsb` (pictures and charts linked by one drawing
  record) is editable like one loaded from `.xlsx`: `addPicture`, `addChart`
  and the chart setters succeed instead of failing with `IoDrawingUnparseable`,
  `copySheet` copies it, and row and column inserts move its anchors. A
  drawing the reader cannot lift stays read-only. Its relationship id may
  change on save.
- Creating, resizing or renaming a table refuses what Excel refuses with
  `InvalidArgument`: a display name outside Excel's name syntax (a cell
  address such as `T1`, a name opening with an R1C1 reference such as
  `R2D2`, `TRUE`, a leading digit, a space), a name a defined name of any
  scope already uses, and a range overlapping another table on the sheet.
  Creating a table also refuses a range overlapping a PivotTable, a data
  table or a multi-cell array formula, or one that cuts through merged
  cells, and unmerges the merges the range covers, as Excel does. Where
  Excel widens the range to take in a merge it cuts through, the table is
  refused instead, because its columns are given by the caller. Resizing
  refuses the same overlaps, any merged cell in the new range, a move of
  the first row, and a range left with no data row under a header row.
  Setting a new defined name refuses a name outside that syntax or one a
  table uses. Such workbooks were written without complaint and Excel
  refused to open them.
- A dynamic-array result of more than one cell whose spill range touches a
  table, its anchor cell included, is `#SPILL!`, as in Excel. Adding,
  resizing or removing a table, and row or column edits that move one,
  recalculate the spills it blocks or releases.

- An empty patch changes nothing on every patch method. `Chart.updatePoint`
  with no key used to remove the point's formatting; it leaves the point as it
  is now, and a point whose override ends up with no fill, line or marker loses
  it (`Chart.setPoint(plot, series, point, null)` removes it outright). The same
  holds for `Pivot.updateField`, `Pivot.updateFieldItem` and
  `Pivot.updateStyleOptions`, where an all-`KEEP` C patch is a no-op.
- The block reads `getValues`, `getTypedValues`, `getCellsInRange`, `iterRows`
  and `listCells` (Python `get_values` and the others) read the used range when
  `range` is omitted or `null`, and a sheet with no used range yields no rows.
- `PivotField.showDetail` is `true` only when every item and the field's
  default marker are expanded (`true` for a field without items), and writing
  it expands or collapses all of them. `PivotDataField.aggregation` is optional on
  input and defaults to `PivotAggregation.Sum`.
- `Chart.setFilter` hides nothing for an omitted list. `addChart` with an
  omitted `legend` places it at the bottom, as Excel inserts a chart, and
  `null` means no legend. A literal series source is written (it was
  rejected), a literal series name as `<c:tx><c:v>`.
- A sparkline `source` and a `dateAxisRange` whose stored text does not parse
  read `null`, and the stored text is kept on write.
- `getFunctionMetadata` returns `localizedName` for every function; with no
  localized entry it is the function's own `name`.
- A pivot lookup by an unknown `cacheId` in `fm_sheet_add_pivot` is
  `FM_STATUS_NOT_FOUND`, and an out-of-range pivot, field or item index in
  `fm_sheet_update_pivot_field`, `fm_sheet_update_pivot_field_item` and
  `fm_sheet_update_pivot_style_options` is `FM_STATUS_INDEX_OUT_OF_RANGE`
  (13) instead of `FM_STATUS_INVALID_ARGUMENT`.
- The chart `style` (`c:style`) reads and writes one scale, 1 to 48; a read
  prefers `c:style`, and a file that carries only an `mc:AlternateContent`
  `c14:style` of 100 + n reads as n.

- Data validation booleans default to true, as in Excel's dialog:
  `allowBlank`, `showInputMessage`, `showErrorMessage` and
  `showInCellDropdown` (Python `allow_blank` and so on). In C each is an
  `fm_tristate_t` whose `FM_TRISTATE_DEFAULT` is Excel's default. A rule that
  omitted `allowBlank` used to reject an empty cell and now accepts it.
- `SheetProtection` `allow*` flags carry Excel's defaults: selecting locked and
  unlocked cells is allowed and every other action is not, so
  `ws.setSheetProtection({ allowFormatCells: true })` protects the sheet with
  formatting allowed and nothing else. `ws.setSheetProtection(null)` unprotects
  it. In C the two select flags are `fm_tristate_t` (`DEFAULT` = allowed).
- A table created without a style takes the workbook's `tableStyleName`
  default (`updateTableStyleDefaults({ tableStyleName: 'TableStyleLight9' })`),
  and `TableStyleMedium2` with row stripes when no default is set. `addPivot`
  takes `pivotStyleName` from the same defaults, so a new table or pivot is no
  longer always `TableStyleMedium2` / `PivotStyleLight16`.
- Table indices are sheet-local: `ws.table(0)` is the first table of that
  sheet, not of the workbook, and `ws.listTables()` lists that sheet's tables.
- Margins are read and written in inches and the print scale as a percentage
  (`scalePct`), on page setup and on pagination, in JavaScript, Python and C.
  Row heights are points and column widths characters on every surface.
- Singletons are removed with `set<X>(null)` (see BREAKING): a call that used
  `remove*` or `clear*` on one of them is removed rather than aliased.
  `remove*` on a missing keyed record fails with `NotFound`, and `clear*` on an
  empty collection succeeds.
- `update<X>(patch)` merges: a key the patch names changes, every other key
  keeps its value, and `null` clears a nullable key. `set<X>(record)` replaces
  the whole value, so `ws.setSheetProtection(input)`, `Chart.setLegend(legend)`
  and `Chart.setPoint(..., point)` drop whatever the record leaves out.

- A name or id that is already taken is `AlreadyExists` (C `FM_STATUS_ALREADY_EXISTS`
  (7), JavaScript `StatusCode.AlreadyExists`, Python `AlreadyExistsError`), and
  a refusal caused by the workbook's state is `PreconditionFailed` (C
  `FM_STATUS_PRECONDITION_FAILED` (9), Python `FormulonError` with
  `code == StatusCode.PRECONDITION_FAILED`). Both used to be `InvalidArgument`
  or a specific code. `AlreadyExists` covers a sheet, defined name, table,
  table style, cell style, person, threaded-comment id, slicer, protected
  range, pivot cache id, pivot field and pivot item name that collides.
  `PreconditionFailed` covers removing the last sheet or the last visible
  sheet (`kCannotRemoveLastSheet`, 5052, is gone), hiding the only
  visible sheet, making a hidden sheet the active one, a table or pivot range that overlaps another object, removing
  a pivot cache still in use, a field another calculated field or item names,
  and a write to a non-anchor cell of a legacy array formula.
- A lookup by an unknown object id (chart, drawing object, shape, slicer,
  picture, threaded comment, pivot cache) is `NotFound` on every function,
  where several reported `InvalidArgument`. `localizeFunctionName` with a
  function name that has no localized form is `NotFound`; an unknown profile
  id stays `InvalidArgument`.
- A position past the end of a collection is `IndexOutOfRange` on every
  function that takes an element index, and a `localSheetId` outside the
  workbook is `SheetIndexOutOfRange`.
- A foreign handle (a `Sheet`, `Table`, `Chart` or `Pivot` of another workbook)
  is `InvalidArgument` in JavaScript and Python alike.
- A PivotTable report lives in the sheet cells, and where it may be drawn
  follows Excel:
  - An edit of a pivot (a field, an item, an option, a filter, a grouping), a
    cache refresh and `Pivot.project()` never fail because the new report
    would grow over something. The change is applied, the old report's cells
    are cleared and the anchor holds `#SPILL!`; the report is blocked. A value,
    a `=""` formula, a merge (empty or not), an array formula, a data table's
    result cells, a table, a spill range and another pivot's report all block
    it, as does the sheet edge. A cell holding only a format does not, and
    loses that format when the report is drawn over it. Nothing is cleared to
    make room.
  - The next recalc after the range comes free draws the report, so in
    automatic calculation the first read after the clearing write shows it; in
    manual calculation it waits for an explicit recalc. A format-only change,
    such as removing a merge, starts no recalc.
  - `Pivot.getSpill()` (Python `get_spill`, C `fm_sheet_get_pivot_spill`)
    returns `{ blocked: true, range }` for a blocked report, and
    `Sheet.getSpill` returns the same at its anchor, so `Spill.blocked` no
    longer implies a formula at the anchor. A range past the sheet edge is not
    reported, as for a formula spill. `isProjectionPending()` is false for a
    blocked report, and `getLayout()` returns the layout it would draw.
    GETPIVOTDATA over a blocked report is `#REF!`.
  - Of two reports refreshed together, the one growing into the other's
    current report is blocked; when only their new ranges overlap, the one
    created first is drawn.
  - `addPivot` draws Excel's empty placeholder for a pivot with no field, a
    bordered frame 3 columns wide and 18 rows high with no values, so the cell
    count, the used range and `getCellsInRange` change as soon as a pivot is
    created. Creating a pivot whose range is occupied or runs past the sheet
    edge fails with `PreconditionFailed` and creates nothing.
  - `Pivot.setAnchor` moves a report as Excel's cut and paste does. In the
    destination, values, formulas, merges, spill cells and array formulas or
    data tables lying wholly inside are overwritten, and a table is deleted.
    The move fails with `PreconditionFailed`, changing nothing, when the
    destination covers another pivot, meets an array formula or data table
    crossing its edge, or runs past the sheet edge.
- Setters refuse what Excel refuses with `InvalidArgument`: a row height above
  409 pt, a column width above 255 characters, a font size outside 1 to 409, a
  cell text over 32,767 characters, a zoom outside 10 to 400, a print scale
  outside 10 to 400, a column outline level above 7 and a pane origin outside
  the grid. A file that carries such a value still loads and the value is
  clamped as Excel clamps it. A negative `maxChange` in an iterative-settings
  patch is `InvalidArgument`.
- `ws.updateView(patch)` (Python `ws.update_view`) is the single call that
  changes the view. A patch that sets both a freeze and a split is
  `InvalidArgument`. `ws.setVisibility` hiding the only visible sheet is
  `PreconditionFailed`.
- A JavaScript `undefined` or an omitted Python argument keeps a patch field,
  `null` / `None` clears it on a field that can be cleared, and a value sets
  it. `null` on a field that cannot be cleared is `TypeError` in both
  runtimes. The five password strings of a protected range and
  `PageSetupPatch.orientation` are clearable; clearing `orientation` removes
  the attribute.
- An optional string that is absent reads `null` in JavaScript and `None` in
  Python (`ParsedReference.sheet`, record `comment` and `name` fields, a table
  style name), where C returns `NULL`.
- A reference text that does not parse is `ParserInvalidReference`, and a font
  size of zero or less is `InvalidArgument` in every binding, with no
  `RangeError`.
- Clearing a cell-font attribute follows one rule. `sizePt` / `size_pt` and
  `name`, which a font always carries, take the workbook's Normal-style font
  value when set to `null` / `None` or left out of a new font, where JavaScript
  used to write 11 pt and Python refused the call. `color` can be cleared too
  and then means Excel's automatic colour. In C, a cell font with
  `has_size_pt == 0` or a `NULL` `name` takes the Normal font's value, and
  `size_pt` is range-checked only when stated.

## Reference tables

### Name migration

<!-- migration-table:begin generated by tools/api/gen_migration_table.py; do not edit -->
#### C ABI

| Before | After |
|---|---|
| `fm_auto_filter` | `fm_auto_filter_t` |
| `fm_border_record` | `fm_border_t` |
| `fm_border_side` | `fm_border_side_t` |
| `fm_buffer_free` | `fm_buffer_destroy` |
| `fm_cell_clear_xf_index` | `fm_sheet_reset_cell_xf_index` |
| `fm_cell_get_xf_index` | `fm_sheet_get_cell_xf_index` |
| `fm_cell_node_t.sheet` | `fm_cell_node_info_t.sheet_index` |
| `fm_cell_node_t` | `fm_cell_node_info_t` |
| `fm_cell_nodes_at` | `fm_cell_nodes_get_node_at` |
| `fm_cell_nodes_count` | `fm_cell_nodes_get_node_count` |
| `fm_cell_range_at` | `fm_cell_range_get_cell_at` |
| `fm_cell_range_count` | `fm_cell_range_get_cell_count` |
| `fm_cell_set_xf_index` | `fm_sheet_set_cell_xf_index` |
| `FM_CELL_STYLE_BUILTIN_ID_NONE` | removed |
| `fm_cell_style_record_t.i_level` | `fm_cell_style_t.outline_level` |
| `fm_cell_style_record_t` | `fm_cell_style_t` |
| `fm_cell_xf.has_horizontal_align` | `fm_cell_xf_t.has_horizontal` |
| `fm_cell_xf.has_vertical_align` | `fm_cell_xf_t.has_vertical` |
| `fm_cell_xf.horizontal_align` | `fm_cell_xf_t.horizontal` |
| `fm_cell_xf.vertical_align` | `fm_cell_xf_t.vertical` |
| `fm_cell_xf` | `fm_cell_xf_t` |
| `fm_cf_cell_range_t` | folded into `fm_cell_rect_t` |
| `FM_CF_COLOR_SCALE` | `FM_CF_MATCH_KIND_COLOR_SCALE` |
| `fm_cf_color_t` | folded into `fm_color_t` |
| `FM_CF_DATA_BAR` | `FM_CF_MATCH_KIND_DATA_BAR` |
| `FM_CF_DIFFERENTIAL_FORMAT` | `FM_CF_MATCH_KIND_DIFFERENTIAL_FORMAT` |
| `FM_CF_ICON_SET` | `FM_CF_MATCH_KIND_ICON_SET` |
| `fm_cf_match_t.dxf_id_engaged` | `fm_cf_match_info_t.has_dxf_index` |
| `fm_cf_match_t.icon_set_name` | `fm_cf_match_info_t.icon_set` |
| `fm_cf_match_t` | `fm_cf_match_info_t` |
| `fm_cf_results_cell_at` | `fm_cf_results_get_cell_at` |
| `fm_cf_results_cell_count` | `fm_cf_results_get_cell_count` |
| `fm_cf_results_match_at` | `fm_cf_results_get_match_at` |
| `fm_cf_rule_t.data_bar_axis_position_engaged` | `fm_cf_rule_t.has_data_bar_axis_position` |
| `fm_cf_rule_t.data_bar_engaged` | `fm_cf_rule_t.has_data_bar` |
| `fm_cf_rule_t.data_bar_gradient_engaged` | `fm_cf_rule_t.has_data_bar_gradient` |
| `fm_cf_rule_t.dxf_id_engaged` | `fm_cf_rule_t.has_dxf_index` |
| `fm_cf_rule_t.dxf_id` | `fm_cf_rule_t.dxf_index` |
| `fm_cf_rule_t.icon_set_engaged` | `fm_cf_rule_t.has_icon_set` |
| `fm_cf_rule_t.icon_set_floor_engaged` | `fm_cf_rule_t.has_icon_set_floor` |
| `fm_cf_rule_t.icon_set_name` | `fm_cf_rule_t.icon_set` |
| `fm_cf_rule_t.op_engaged` | `fm_cf_rule_t.has_op` |
| `fm_cf_rule_t.rank_engaged` | `fm_cf_rule_t.has_rank` |
| `fm_cf_rule_t.std_dev_engaged` | `fm_cf_rule_t.has_std_dev` |
| `fm_cf_rule_t.time_period_engaged` | `fm_cf_rule_t.has_time_period` |
| `fm_chart_axis_at` | `fm_chart_get_axis_at` |
| `FM_CHART_AXIS_CATEGORY` | `FM_CHART_AXIS_KIND_CATEGORY` |
| `FM_CHART_AXIS_DATE` | `FM_CHART_AXIS_KIND_DATE` |
| `fm_chart_axis_patch.tick_label_pos_op` | `fm_chart_axis_patch_t.tick_label_position_op` |
| `fm_chart_axis_patch.tick_label_pos` | `fm_chart_axis_patch_t.tick_label_position` |
| `fm_chart_axis_patch` | `fm_chart_axis_patch_t` |
| `FM_CHART_AXIS_POS_BOTTOM` | `FM_CHART_AXIS_POSITION_BOTTOM` |
| `FM_CHART_AXIS_POS_LEFT` | `FM_CHART_AXIS_POSITION_LEFT` |
| `FM_CHART_AXIS_POS_RIGHT` | `FM_CHART_AXIS_POSITION_RIGHT` |
| `FM_CHART_AXIS_POS_TOP` | `FM_CHART_AXIS_POSITION_TOP` |
| `FM_CHART_AXIS_SERIES` | `FM_CHART_AXIS_KIND_SERIES` |
| `fm_chart_axis_t.deleted` | `fm_chart_axis_info_t.visible` |
| `fm_chart_axis_t.tick_label_pos` | `fm_chart_axis_info_t.tick_label_position` |
| `fm_chart_axis_t` | `fm_chart_axis_info_t` |
| `fm_chart_axis_title_text_t` | `fm_chart_axis_title_text_info_t` |
| `FM_CHART_AXIS_VALUE` | `FM_CHART_AXIS_KIND_VALUE` |
| `FM_CHART_BAR_DIR_ABSENT` | `FM_CHART_BAR_DIRECTION_ABSENT` |
| `FM_CHART_BAR_DIR_BAR` | `FM_CHART_BAR_DIRECTION_BAR` |
| `FM_CHART_BAR_DIR_COL` | `FM_CHART_BAR_DIRECTION_COLUMN` |
| `FM_CHART_BLANKS_ABSENT` | `FM_BLANKS_AS_ABSENT` |
| `fm_chart_blanks_as_t` | `fm_blanks_as_t` |
| `FM_CHART_BLANKS_GAP` | `FM_BLANKS_AS_GAP` |
| `FM_CHART_BLANKS_SPAN` | `FM_BLANKS_AS_SPAN` |
| `FM_CHART_BLANKS_ZERO` | `FM_BLANKS_AS_ZERO` |
| `FM_CHART_CROSSES_ABSENT` | `FM_CHART_AXIS_CROSSES_ABSENT` |
| `FM_CHART_CROSSES_AT` | `FM_CHART_AXIS_CROSSES_AT` |
| `FM_CHART_CROSSES_AUTO_ZERO` | `FM_CHART_AXIS_CROSSES_AUTO_ZERO` |
| `FM_CHART_CROSSES_MAX` | `FM_CHART_AXIS_CROSSES_MAX` |
| `FM_CHART_CROSSES_MIN` | `FM_CHART_AXIS_CROSSES_MIN` |
| `FM_CHART_DASH_DASH_DOT` | `FM_CHART_LINE_DASH_DASH_DOT` |
| `FM_CHART_DASH_DASH` | `FM_CHART_LINE_DASH_DASH` |
| `FM_CHART_DASH_DOT` | `FM_CHART_LINE_DASH_DOT` |
| `FM_CHART_DASH_OTHER` | `FM_CHART_LINE_DASH_OTHER` |
| `FM_CHART_DASH_SOLID` | `FM_CHART_LINE_DASH_SOLID` |
| `fm_chart_data_axis_title_at` | `fm_chart_data_get_axis_title_at` |
| `fm_chart_data_category_at` | `fm_chart_data_get_category_at` |
| `fm_chart_data_category_level_at` | `fm_chart_data_get_category_level_at` |
| `fm_chart_data_category_level_size` | `fm_chart_data_get_category_level_count` |
| `fm_chart_data_info` | `fm_chart_data_get_info` |
| `fm_chart_data_points_t` | `fm_chart_data_points_info_t` |
| `fm_chart_data_range_at` | `fm_chart_data_get_range_at` |
| `fm_chart_data_series_at` | `fm_chart_data_get_series_at` |
| `fm_chart_data_series_points` | `fm_chart_data_get_series_points` |
| `fm_chart_data_series_t.format_code` | `fm_chart_data_series_info_t.number_format` |
| `fm_chart_data_series_t.idx` | `fm_chart_data_series_info_t.series_id` |
| `fm_chart_data_series_t.x_format_code` | `fm_chart_data_series_info_t.x_number_format` |
| `fm_chart_data_series_t` | `fm_chart_data_series_info_t` |
| `fm_chart_data_t.category_format_code` | `fm_chart_data_t.category_number_format` |
| `fm_chart_data_unresolved_at` | `fm_chart_data_get_unresolved_at` |
| `FM_CHART_DIRECTION_AUTO` | `FM_CHART_SERIES_DIRECTION_AUTO` |
| `FM_CHART_DIRECTION_COLUMNS` | `FM_CHART_SERIES_DIRECTION_COLUMNS` |
| `FM_CHART_DIRECTION_ROWS` | `FM_CHART_SERIES_DIRECTION_ROWS` |
| `fm_chart_fill_patch.transparency_op` | `fm_chart_fill_patch_t.transparency_pct_op` |
| `fm_chart_fill_patch` | `fm_chart_fill_patch_t` |
| `FM_CHART_HEADER_AUTO` | removed |
| `fm_chart_header_mode_t` | folded into `fm_tristate_t` |
| `FM_CHART_HEADER_NO` | removed |
| `FM_CHART_HEADER_YES` | removed |
| `fm_chart_hidden_category_at` | `fm_chart_get_hidden_category_at` |
| `fm_chart_info_t.plot_vis_only` | `fm_chart_info_t.plot_visible_only` |
| `fm_chart_info` | `fm_chart_get_info` |
| `FM_CHART_INSERT_AREA` | removed |
| `FM_CHART_INSERT_BAR` | removed |
| `FM_CHART_INSERT_COLUMN` | removed |
| `FM_CHART_INSERT_DOUGHNUT` | removed |
| `fm_chart_insert_kind_t` | folded into `fm_chart_plot_type_t` |
| `FM_CHART_INSERT_LINE` | removed |
| `FM_CHART_INSERT_PIE` | removed |
| `FM_CHART_INSERT_SCATTER` | removed |
| `fm_chart_insert` | `fm_chart_spec_t` |
| `FM_CHART_LABEL_ABOVE` | `FM_CHART_LABEL_POSITION_ABOVE` |
| `FM_CHART_LABEL_BELOW` | `FM_CHART_LABEL_POSITION_BELOW` |
| `FM_CHART_LABEL_BEST_FIT` | `FM_CHART_LABEL_POSITION_BEST_FIT` |
| `FM_CHART_LABEL_CENTER` | `FM_CHART_LABEL_POSITION_CENTER` |
| `FM_CHART_LABEL_INSIDE_BASE` | `FM_CHART_LABEL_POSITION_INSIDE_BASE` |
| `FM_CHART_LABEL_INSIDE_END` | `FM_CHART_LABEL_POSITION_INSIDE_END` |
| `FM_CHART_LABEL_LEFT` | `FM_CHART_LABEL_POSITION_LEFT` |
| `FM_CHART_LABEL_OUTSIDE_END` | `FM_CHART_LABEL_POSITION_OUTSIDE_END` |
| `FM_CHART_LABEL_RIGHT` | `FM_CHART_LABEL_POSITION_RIGHT` |
| `fm_chart_labels_patch.font_size_op` | `fm_chart_labels_patch_t.font_size_pt_op` |
| `fm_chart_labels_patch` | `fm_chart_labels_patch_t` |
| `FM_CHART_LEGEND_BOTTOM` | `FM_CHART_LEGEND_POSITION_BOTTOM` |
| `fm_chart_legend_hidden_at` | `fm_chart_get_legend_hidden_at` |
| `FM_CHART_LEGEND_LEFT` | `FM_CHART_LEGEND_POSITION_LEFT` |
| `FM_CHART_LEGEND_NONE` | `FM_CHART_LEGEND_POSITION_ABSENT` |
| `fm_chart_legend_patch.hidden_series_count` | `fm_chart_legend_patch_t.hidden_entries_count` |
| `fm_chart_legend_patch.hidden_series_op` | `fm_chart_legend_patch_t.hidden_entries_op` |
| `fm_chart_legend_patch.hidden_series` | `fm_chart_legend_patch_t.hidden_entries` |
| `fm_chart_legend_patch` | `fm_chart_legend_patch_t` |
| `FM_CHART_LEGEND_RIGHT` | `FM_CHART_LEGEND_POSITION_RIGHT` |
| `FM_CHART_LEGEND_TOP_RIGHT` | `FM_CHART_LEGEND_POSITION_TOP_RIGHT` |
| `FM_CHART_LEGEND_TOP` | `FM_CHART_LEGEND_POSITION_TOP` |
| `fm_chart_line_patch.transparency_op` | `fm_chart_line_patch_t.transparency_pct_op` |
| `fm_chart_line_patch.width_op` | `fm_chart_line_patch_t.width_pt_op` |
| `fm_chart_line_patch` | `fm_chart_line_patch_t` |
| `fm_chart_literal_t` | `fm_chart_literal_info_t` |
| `FM_CHART_MARKER_AUTO` | `FM_CHART_MARKER_SYMBOL_AUTO` |
| `FM_CHART_MARKER_CIRCLE` | `FM_CHART_MARKER_SYMBOL_CIRCLE` |
| `FM_CHART_MARKER_DASH` | `FM_CHART_MARKER_SYMBOL_DASH` |
| `FM_CHART_MARKER_DIAMOND` | `FM_CHART_MARKER_SYMBOL_DIAMOND` |
| `FM_CHART_MARKER_DOT` | `FM_CHART_MARKER_SYMBOL_DOT` |
| `FM_CHART_MARKER_NONE` | `FM_CHART_MARKER_SYMBOL_NONE` |
| `fm_chart_marker_patch.size_op` | `fm_chart_marker_patch_t.size_pt_op` |
| `fm_chart_marker_patch` | `fm_chart_marker_patch_t` |
| `FM_CHART_MARKER_PICTURE` | `FM_CHART_MARKER_SYMBOL_PICTURE` |
| `FM_CHART_MARKER_PLUS` | `FM_CHART_MARKER_SYMBOL_PLUS` |
| `FM_CHART_MARKER_SQUARE` | `FM_CHART_MARKER_SYMBOL_SQUARE` |
| `FM_CHART_MARKER_STAR` | `FM_CHART_MARKER_SYMBOL_STAR` |
| `fm_chart_marker_t.has_size` | `fm_chart_marker_t.has_size_pt` |
| `fm_chart_marker_t.size` | `fm_chart_marker_t.size_pt` |
| `FM_CHART_MARKER_TRIANGLE` | `FM_CHART_MARKER_SYMBOL_TRIANGLE` |
| `FM_CHART_MARKER_X` | `FM_CHART_MARKER_SYMBOL_X` |
| `FM_CHART_PATCH_CLEAR` | `FM_PATCH_OP_CLEAR` |
| `FM_CHART_PATCH_KEEP` | `FM_PATCH_OP_KEEP` |
| `fm_chart_patch_op_t` | `fm_patch_op_t` |
| `FM_CHART_PATCH_SET` | `FM_PATCH_OP_SET` |
| `FM_CHART_PLOT_AREA_3D` | `FM_CHART_PLOT_TYPE_AREA_3D` |
| `FM_CHART_PLOT_AREA` | `FM_CHART_PLOT_TYPE_AREA` |
| `fm_chart_plot_at` | `fm_chart_get_plot_at` |
| `fm_chart_plot_axis_id_at` | `fm_chart_get_plot_axis_id_at` |
| `FM_CHART_PLOT_BAR_3D` | `FM_CHART_PLOT_TYPE_BAR_3D` |
| `FM_CHART_PLOT_BAR` | `FM_CHART_PLOT_TYPE_BAR` |
| `FM_CHART_PLOT_BUBBLE` | `FM_CHART_PLOT_TYPE_BUBBLE` |
| `FM_CHART_PLOT_CHARTEX` | `FM_CHART_PLOT_TYPE_EXTENDED` |
| `FM_CHART_PLOT_DOUGHNUT` | `FM_CHART_PLOT_TYPE_DOUGHNUT` |
| `FM_CHART_PLOT_LINE_3D` | `FM_CHART_PLOT_TYPE_LINE_3D` |
| `FM_CHART_PLOT_LINE` | `FM_CHART_PLOT_TYPE_LINE` |
| `FM_CHART_PLOT_OF_PIE` | `FM_CHART_PLOT_TYPE_OF_PIE` |
| `fm_chart_plot_patch.first_slice_angle_op` | `fm_chart_plot_patch_t.first_slice_angle_deg_op` |
| `fm_chart_plot_patch.gap_width_op` | `fm_chart_plot_patch_t.gap_width_pct_op` |
| `fm_chart_plot_patch.hole_size_op` | `fm_chart_plot_patch_t.hole_size_pct_op` |
| `fm_chart_plot_patch.hole_size` | `fm_chart_plot_patch_t.hole_size_pct` |
| `fm_chart_plot_patch.overlap_op` | `fm_chart_plot_patch_t.overlap_pct_op` |
| `fm_chart_plot_patch.overlap` | `fm_chart_plot_patch_t.overlap_pct` |
| `fm_chart_plot_patch` | `fm_chart_plot_patch_t` |
| `FM_CHART_PLOT_PIE_3D` | `FM_CHART_PLOT_TYPE_PIE_3D` |
| `FM_CHART_PLOT_PIE` | `FM_CHART_PLOT_TYPE_PIE` |
| `FM_CHART_PLOT_RADAR` | `FM_CHART_PLOT_TYPE_RADAR` |
| `FM_CHART_PLOT_SCATTER` | `FM_CHART_PLOT_TYPE_SCATTER` |
| `FM_CHART_PLOT_STOCK` | `FM_CHART_PLOT_TYPE_STOCK` |
| `FM_CHART_PLOT_SURFACE_3D` | `FM_CHART_PLOT_TYPE_SURFACE_3D` |
| `FM_CHART_PLOT_SURFACE` | `FM_CHART_PLOT_TYPE_SURFACE` |
| `fm_chart_plot_t.bar_dir` | `fm_chart_plot_info_t.bar_direction` |
| `fm_chart_plot_t.chartex_layout` | `fm_chart_plot_info_t.extended_layout` |
| `fm_chart_plot_t.first_slice_angle` | `fm_chart_plot_info_t.first_slice_angle_deg` |
| `fm_chart_plot_t.gap_width` | `fm_chart_plot_info_t.gap_width_pct` |
| `fm_chart_plot_t.has_first_slice_angle` | `fm_chart_plot_info_t.has_first_slice_angle_deg` |
| `fm_chart_plot_t.has_gap_width` | `fm_chart_plot_info_t.has_gap_width_pct` |
| `fm_chart_plot_t.has_hole_size` | `fm_chart_plot_info_t.has_hole_size_pct` |
| `fm_chart_plot_t.has_overlap` | `fm_chart_plot_info_t.has_overlap_pct` |
| `fm_chart_plot_t.hole_size` | `fm_chart_plot_info_t.hole_size_pct` |
| `fm_chart_plot_t.overlap` | `fm_chart_plot_info_t.overlap_pct` |
| `fm_chart_plot_t` | `fm_chart_plot_info_t` |
| `FM_CHART_POINT_CONNECT` | `FM_CHART_POINT_STATE_CONNECT` |
| `FM_CHART_POINT_GAP` | `FM_CHART_POINT_STATE_GAP` |
| `fm_chart_point_patch` | `fm_chart_point_patch_t` |
| `fm_chart_point_t` | `fm_chart_point_info_t` |
| `FM_CHART_POINT_VALUE` | `FM_CHART_POINT_STATE_VALUE` |
| `fm_chart_preview_overrides.plot_vis_only_op` | `fm_chart_preview_overrides_spec_t.plot_visible_only_op` |
| `fm_chart_preview_overrides.plot_vis_only` | `fm_chart_preview_overrides_spec_t.plot_visible_only` |
| `fm_chart_preview_overrides` | `fm_chart_preview_overrides_spec_t` |
| `fm_chart_series_at` | `fm_chart_get_series_at` |
| `fm_chart_series_literal_at` | `fm_chart_get_series_literal_at` |
| `fm_chart_series_patch` | `fm_chart_series_patch_t` |
| `fm_chart_series_point_at` | `fm_chart_get_series_point_at` |
| `fm_chart_series_t.idx` | `fm_chart_series_info_t.series_id` |
| `fm_chart_series_t` | `fm_chart_series_info_t` |
| `fm_chart_sheet_range_t` | folded into `fm_sheet_rect_t` |
| `FM_CHART_SOURCE_LITERAL` | `FM_CHART_SOURCE_KIND_LITERAL` |
| `FM_CHART_SOURCE_NONE` | `FM_CHART_SOURCE_KIND_NONE` |
| `FM_CHART_SOURCE_REF` | `FM_CHART_SOURCE_KIND_REF` |
| `FM_CHART_TITLE_AUTO` | `FM_CHART_TITLE_KIND_AUTO` |
| `FM_CHART_TITLE_NONE` | `FM_CHART_TITLE_KIND_NONE` |
| `FM_CHART_TITLE_REF` | `FM_CHART_TITLE_KIND_REF` |
| `FM_CHART_TITLE_TEXT` | `FM_CHART_TITLE_KIND_TEXT` |
| `FM_CHART_TRI_ABSENT` | `FM_TRISTATE_DEFAULT` |
| `FM_CHART_TRI_FALSE` | `FM_TRISTATE_FALSE` |
| `fm_chart_tri_t` | `fm_tristate_t` |
| `FM_CHART_TRI_TRUE` | `FM_TRISTATE_TRUE` |
| `fm_chart_unresolved_t.series_idx` | `fm_chart_unresolved_info_t.series_id` |
| `fm_chart_unresolved_t.source` | `fm_chart_unresolved_info_t.role` |
| `fm_chart_unresolved_t` | `fm_chart_unresolved_info_t` |
| `fm_collation_classes` | `fm_get_collation_classes` |
| `FM_COLLATION_CLASSIC_EXACT` | `FM_COLLATION_STRENGTH_CLASSIC_EXACT` |
| `FM_COLLATION_COMPARE` | `FM_COLLATION_STRENGTH_COMPARE` |
| `FM_COLLATION_DYNAMIC` | `FM_COLLATION_STRENGTH_DYNAMIC` |
| `FM_COLOR_CONTEXT_FILL_BG` | `FM_COLOR_CONTEXT_FILL_BACKGROUND` |
| `FM_COLOR_CONTEXT_FILL_FG` | `FM_COLOR_CONTEXT_FILL_FOREGROUND` |
| `fm_color_spec` | `fm_color_t` |
| `fm_column_layout_t.has_style` | `fm_column_layout_info_t.has_xf_index` |
| `fm_column_layout_t.has_width` | `fm_column_layout_info_t.has_width_chars` |
| `fm_column_layout_t.style_xf` | `fm_column_layout_info_t.xf_index` |
| `fm_column_layout_t` | `fm_column_layout_info_t` |
| `fm_comment` | `fm_comment_t` |
| `fm_data_table.col_input_col` | `fm_data_table_t.column_input_col` |
| `fm_data_table.col_input_deleted` | `fm_data_table_t.column_input_deleted` |
| `fm_data_table.col_input_row` | `fm_data_table_t.column_input_row` |
| `fm_data_table.has_col_input` | `fm_data_table_t.has_column_input` |
| `fm_data_table` | `fm_data_table_t` |
| `fm_data_validation` | `fm_data_validation_t` |
| `fm_date_group_item` | `fm_date_group_item_t` |
| `FM_DISPLAY_INVALID_FORMAT` | `FM_DISPLAY_STATUS_INVALID_FORMAT` |
| `FM_DISPLAY_OK` | `FM_DISPLAY_STATUS_OK` |
| `FM_DISPLAY_OVERFLOW` | `FM_DISPLAY_STATUS_OVERFLOW` |
| `fm_drawing_anchor` | `fm_drawing_anchor_t` |
| `FM_DRAWING_OBJECT_CHART` | `FM_DRAWING_OBJECT_KIND_CHART` |
| `FM_DRAWING_OBJECT_CONNECTOR` | `FM_DRAWING_OBJECT_KIND_CONNECTOR` |
| `FM_DRAWING_OBJECT_GRAPHIC_FRAME` | `FM_DRAWING_OBJECT_KIND_GRAPHIC_FRAME` |
| `FM_DRAWING_OBJECT_GROUP` | `FM_DRAWING_OBJECT_KIND_GROUP` |
| `FM_DRAWING_OBJECT_OTHER` | `FM_DRAWING_OBJECT_KIND_OTHER` |
| `FM_DRAWING_OBJECT_PICTURE` | `FM_DRAWING_OBJECT_KIND_PICTURE` |
| `FM_DRAWING_OBJECT_SHAPE` | `FM_DRAWING_OBJECT_KIND_SHAPE` |
| `fm_drawing_object` | `fm_drawing_object_info_t` |
| `fm_dxf_record.border_engaged` | `fm_dxf_t.has_border` |
| `fm_dxf_record.fill_engaged` | `fm_dxf_t.has_fill` |
| `fm_dxf_record.font_engaged` | `fm_dxf_t.has_font` |
| `fm_dxf_record.num_fmt_code` | `fm_dxf_t.number_format` |
| `fm_dxf_record` | `fm_dxf_t` |
| `FM_DYNAMIC_FILTER_ABOVE_AVERAGE` | `FM_DYNAMIC_FILTER_TYPE_ABOVE_AVERAGE` |
| `FM_DYNAMIC_FILTER_BELOW_AVERAGE` | `FM_DYNAMIC_FILTER_TYPE_BELOW_AVERAGE` |
| `FM_DYNAMIC_FILTER_LAST_MONTH` | `FM_DYNAMIC_FILTER_TYPE_LAST_MONTH` |
| `FM_DYNAMIC_FILTER_LAST_QUARTER` | `FM_DYNAMIC_FILTER_TYPE_LAST_QUARTER` |
| `FM_DYNAMIC_FILTER_LAST_WEEK` | `FM_DYNAMIC_FILTER_TYPE_LAST_WEEK` |
| `FM_DYNAMIC_FILTER_LAST_YEAR` | `FM_DYNAMIC_FILTER_TYPE_LAST_YEAR` |
| `FM_DYNAMIC_FILTER_M10` | `FM_DYNAMIC_FILTER_TYPE_M10` |
| `FM_DYNAMIC_FILTER_M11` | `FM_DYNAMIC_FILTER_TYPE_M11` |
| `FM_DYNAMIC_FILTER_M12` | `FM_DYNAMIC_FILTER_TYPE_M12` |
| `FM_DYNAMIC_FILTER_M1` | `FM_DYNAMIC_FILTER_TYPE_M1` |
| `FM_DYNAMIC_FILTER_M2` | `FM_DYNAMIC_FILTER_TYPE_M2` |
| `FM_DYNAMIC_FILTER_M3` | `FM_DYNAMIC_FILTER_TYPE_M3` |
| `FM_DYNAMIC_FILTER_M4` | `FM_DYNAMIC_FILTER_TYPE_M4` |
| `FM_DYNAMIC_FILTER_M5` | `FM_DYNAMIC_FILTER_TYPE_M5` |
| `FM_DYNAMIC_FILTER_M6` | `FM_DYNAMIC_FILTER_TYPE_M6` |
| `FM_DYNAMIC_FILTER_M7` | `FM_DYNAMIC_FILTER_TYPE_M7` |
| `FM_DYNAMIC_FILTER_M8` | `FM_DYNAMIC_FILTER_TYPE_M8` |
| `FM_DYNAMIC_FILTER_M9` | `FM_DYNAMIC_FILTER_TYPE_M9` |
| `FM_DYNAMIC_FILTER_NEXT_MONTH` | `FM_DYNAMIC_FILTER_TYPE_NEXT_MONTH` |
| `FM_DYNAMIC_FILTER_NEXT_QUARTER` | `FM_DYNAMIC_FILTER_TYPE_NEXT_QUARTER` |
| `FM_DYNAMIC_FILTER_NEXT_WEEK` | `FM_DYNAMIC_FILTER_TYPE_NEXT_WEEK` |
| `FM_DYNAMIC_FILTER_NEXT_YEAR` | `FM_DYNAMIC_FILTER_TYPE_NEXT_YEAR` |
| `FM_DYNAMIC_FILTER_NULL` | `FM_DYNAMIC_FILTER_TYPE_NULL` |
| `FM_DYNAMIC_FILTER_Q1` | `FM_DYNAMIC_FILTER_TYPE_Q1` |
| `FM_DYNAMIC_FILTER_Q2` | `FM_DYNAMIC_FILTER_TYPE_Q2` |
| `FM_DYNAMIC_FILTER_Q3` | `FM_DYNAMIC_FILTER_TYPE_Q3` |
| `FM_DYNAMIC_FILTER_Q4` | `FM_DYNAMIC_FILTER_TYPE_Q4` |
| `FM_DYNAMIC_FILTER_THIS_MONTH` | `FM_DYNAMIC_FILTER_TYPE_THIS_MONTH` |
| `FM_DYNAMIC_FILTER_THIS_QUARTER` | `FM_DYNAMIC_FILTER_TYPE_THIS_QUARTER` |
| `FM_DYNAMIC_FILTER_THIS_WEEK` | `FM_DYNAMIC_FILTER_TYPE_THIS_WEEK` |
| `FM_DYNAMIC_FILTER_THIS_YEAR` | `FM_DYNAMIC_FILTER_TYPE_THIS_YEAR` |
| `FM_DYNAMIC_FILTER_TODAY` | `FM_DYNAMIC_FILTER_TYPE_TODAY` |
| `FM_DYNAMIC_FILTER_TOMORROW` | `FM_DYNAMIC_FILTER_TYPE_TOMORROW` |
| `FM_DYNAMIC_FILTER_YEAR_TO_DATE` | `FM_DYNAMIC_FILTER_TYPE_YEAR_TO_DATE` |
| `FM_DYNAMIC_FILTER_YESTERDAY` | `FM_DYNAMIC_FILTER_TYPE_YESTERDAY` |
| `fm_effective_style.fill_bg_argb` | `fm_effective_style_info_t.fill_background_argb` |
| `fm_effective_style.fill_bg_resolution` | `fm_effective_style_info_t.fill_background_resolution` |
| `fm_effective_style.fill_fg_argb` | `fm_effective_style_info_t.fill_foreground_argb` |
| `fm_effective_style.fill_fg_resolution` | `fm_effective_style_info_t.fill_foreground_resolution` |
| `fm_effective_style.num_fmt_code` | `fm_effective_style_info_t.number_format` |
| `fm_effective_style` | `fm_effective_style_info_t` |
| `fm_error_display_name` | `fm_get_error_display_name` |
| `fm_external_link_record_t.index` | `fm_external_link_info_t.book_number` |
| `fm_external_link_record_t` | `fm_external_link_info_t` |
| `fm_fill_record.gradient_degree` | `fm_fill_t.gradient_angle_deg` |
| `fm_fill_record` | `fm_fill_t` |
| `fm_filter_column.col_id` | `fm_filter_column_t.column_offset` |
| `fm_filter_column.dxf_id` | `fm_filter_column_t.dxf_index` |
| `fm_filter_column.has_icon_id` | `fm_filter_column_t.has_icon_index` |
| `fm_filter_column.icon_id` | `fm_filter_column_t.icon_index` |
| `fm_filter_column` | `fm_filter_column_t` |
| `FM_FILTER_OPERATOR_EQUAL` | `FM_COMPARISON_OPERATOR_EQUAL` |
| `FM_FILTER_OPERATOR_GREATER_THAN_OR_EQUAL` | `FM_COMPARISON_OPERATOR_GREATER_THAN_OR_EQUAL` |
| `FM_FILTER_OPERATOR_GREATER_THAN` | `FM_COMPARISON_OPERATOR_GREATER_THAN` |
| `FM_FILTER_OPERATOR_LESS_THAN_OR_EQUAL` | `FM_COMPARISON_OPERATOR_LESS_THAN_OR_EQUAL` |
| `FM_FILTER_OPERATOR_LESS_THAN` | `FM_COMPARISON_OPERATOR_LESS_THAN` |
| `FM_FILTER_OPERATOR_NOT_EQUAL` | `FM_COMPARISON_OPERATOR_NOT_EQUAL` |
| `fm_filter_operator_t` | `fm_comparison_operator_t` |
| `fm_font_record.has_size` | `fm_font_t.has_size_pt` |
| `fm_font_record` | `fm_font_t` |
| `fm_formula_canonicalize` | `fm_canonicalize_formula` |
| `fm_formula_localize` | `fm_localize_formula` |
| `fm_function_canonicalize` | `fm_canonicalize_function_name` |
| `fm_function_count` | `fm_get_function_count` |
| `FM_FUNCTION_ENVIRONMENT_BOUND` | `FM_FUNCTION_AVAILABILITY_ENVIRONMENT_BOUND` |
| `FM_FUNCTION_IMPLEMENTED_UNVERIFIED` | `FM_FUNCTION_AVAILABILITY_IMPLEMENTED_UNVERIFIED` |
| `FM_FUNCTION_IMPLEMENTED` | `FM_FUNCTION_AVAILABILITY_IMPLEMENTED` |
| `fm_function_localize` | `fm_localize_function_name` |
| `fm_function_metadata_t` | `fm_function_metadata_info_t` |
| `fm_function_metadata` | `fm_get_function_metadata` |
| `fm_function_name_at` | `fm_get_function_name_at` |
| `FM_FUNCTION_UNAVAILABLE_STUB` | `FM_FUNCTION_AVAILABILITY_UNAVAILABLE_STUB` |
| `FM_GEOMETRY_DISPLAY` | `FM_GEOMETRY_MODE_DISPLAY` |
| `FM_GEOMETRY_PRINT` | `FM_GEOMETRY_MODE_PRINT` |
| `FM_GRADIENT_LINEAR` | `FM_GRADIENT_TYPE_LINEAR` |
| `FM_GRADIENT_PATH` | `FM_GRADIENT_TYPE_PATH` |
| `fm_gradient_stop` | `fm_gradient_stop_t` |
| `fm_header_footer_t.align_with_margins_engaged` | `fm_header_footer_t.has_align_with_margins` |
| `fm_header_footer_t.different_first_engaged` | `fm_header_footer_t.has_different_first` |
| `fm_header_footer_t.different_odd_even_engaged` | `fm_header_footer_t.has_different_odd_even` |
| `fm_header_footer_t.scale_with_doc_engaged` | `fm_header_footer_t.has_scale_with_doc` |
| `fm_hyperlink` | `fm_hyperlink_t` |
| `fm_image_info` | `fm_image_info_t` |
| `fm_image_insert` | `fm_drawing_placement_spec_t` |
| `fm_last_error_context` | `fm_get_last_error_context` |
| `fm_last_error_message` | `fm_get_last_error_message` |
| `fm_locale_error_name_count` | `fm_get_locale_error_name_count` |
| `fm_locale_error_name` | `fm_get_locale_error_name_at` |
| `fm_locale_facts_t` | `fm_locale_facts_info_t` |
| `fm_locale_facts` | `fm_get_locale_facts` |
| `fm_margins_pt` | `fm_margins_pt_info_t` |
| `fm_mention.length` | `fm_mention_t.end` |
| `fm_mention` | `fm_mention_t` |
| `fm_merge_range` | folded into `fm_cell_rect_t` |
| `fm_page_break_t` | `fm_row_break_t` |
| `fm_page_layout` | `fm_page_layout_info_t` |
| `fm_page_margins_t.bottom_engaged` | `fm_page_margins_t.has_bottom_in` |
| `fm_page_margins_t.footer_engaged` | `fm_page_margins_t.has_footer_in` |
| `fm_page_margins_t.header_engaged` | `fm_page_margins_t.has_header_in` |
| `fm_page_margins_t.left_engaged` | `fm_page_margins_t.has_left_in` |
| `fm_page_margins_t.right_engaged` | `fm_page_margins_t.has_right_in` |
| `fm_page_margins_t.top_engaged` | `fm_page_margins_t.has_top_in` |
| `fm_page_setup_t.fit_to_height_engaged` | `fm_page_setup_t.has_fit_to_height` |
| `fm_page_setup_t.fit_to_page_engaged` | `fm_page_setup_t.has_fit_to_page` |
| `fm_page_setup_t.fit_to_width_engaged` | `fm_page_setup_t.has_fit_to_width` |
| `fm_page_setup_t.orientation_engaged` | `fm_page_setup_t.has_orientation` |
| `fm_page_setup_t.paper_size_engaged` | `fm_page_setup_t.has_paper_size` |
| `fm_page_setup_t.scale_engaged` | `fm_page_setup_t.has_scale_pct` |
| `fm_pagination_horizontal_break_at` | `fm_pagination_get_row_break_at` |
| `fm_pagination_horizontal_break_count` | `fm_pagination_get_row_break_count` |
| `fm_pagination_horizontal_break_is_manual` | removed |
| `fm_pagination_margins` | `fm_pagination_get_margins` |
| `fm_pagination_page_at` | `fm_pagination_get_page_at` |
| `fm_pagination_page_count` | `fm_pagination_get_page_count` |
| `fm_pagination_page_order` | `fm_pagination_get_page_order` |
| `fm_pagination_paper` | `fm_pagination_get_paper` |
| `fm_pagination_print_area_at` | `fm_pagination_get_print_area_at` |
| `fm_pagination_print_area_count` | `fm_pagination_get_print_area_count` |
| `fm_pagination_print_titles` | `fm_pagination_get_print_titles` |
| `fm_pagination_printable` | `fm_pagination_get_printable` |
| `fm_pagination_scale` | `fm_pagination_get_scale` |
| `fm_pagination_vertical_break_at` | `fm_pagination_get_column_break_at` |
| `fm_pagination_vertical_break_count` | `fm_pagination_get_column_break_count` |
| `fm_pagination_vertical_break_is_manual` | removed |
| `fm_paper_info` | `fm_paper_info_t` |
| `fm_parallel_recalc_stats` | `fm_parallel_recalc_stats_info_t` |
| `fm_person` | `fm_person_t` |
| `fm_phonetic_run_t.eb` | `fm_phonetic_run_t.end` |
| `fm_phonetic_run_t.sb` | `fm_phonetic_run_t.start` |
| `FM_PIVOT_AGG_AVERAGE` | `FM_PIVOT_AGGREGATION_AVERAGE` |
| `FM_PIVOT_AGG_COUNT_NUMBERS` | `FM_PIVOT_AGGREGATION_COUNT_NUMBERS` |
| `FM_PIVOT_AGG_COUNT` | `FM_PIVOT_AGGREGATION_COUNT` |
| `FM_PIVOT_AGG_MAX` | `FM_PIVOT_AGGREGATION_MAX` |
| `FM_PIVOT_AGG_MIN` | `FM_PIVOT_AGGREGATION_MIN` |
| `FM_PIVOT_AGG_PRODUCT` | `FM_PIVOT_AGGREGATION_PRODUCT` |
| `FM_PIVOT_AGG_STDDEV` | `FM_PIVOT_AGGREGATION_STDDEV` |
| `FM_PIVOT_AGG_STDDEVP` | `FM_PIVOT_AGGREGATION_STDDEVP` |
| `FM_PIVOT_AGG_SUM` | `FM_PIVOT_AGGREGATION_SUM` |
| `FM_PIVOT_AGG_VAR` | `FM_PIVOT_AGGREGATION_VAR` |
| `FM_PIVOT_AGG_VARP` | `FM_PIVOT_AGGREGATION_VARP` |
| `FM_PIVOT_AXIS_COL` | `FM_PIVOT_AXIS_COLUMN` |
| `FM_PIVOT_CELL_BLANK` | `FM_PIVOT_CELL_KIND_BLANK` |
| `FM_PIVOT_CELL_COL_LABEL` | `FM_PIVOT_CELL_KIND_COLUMN_LABEL` |
| `FM_PIVOT_CELL_COL_SUBTOTAL` | `FM_PIVOT_CELL_KIND_COLUMN_SUBTOTAL` |
| `FM_PIVOT_CELL_DATA` | `FM_PIVOT_CELL_KIND_DATA` |
| `FM_PIVOT_CELL_GRAND_TOTAL` | `FM_PIVOT_CELL_KIND_GRAND_TOTAL` |
| `FM_PIVOT_CELL_HEADER` | `FM_PIVOT_CELL_KIND_HEADER` |
| `FM_PIVOT_CELL_ROW_LABEL` | `FM_PIVOT_CELL_KIND_ROW_LABEL` |
| `FM_PIVOT_CELL_ROW_SUBTOTAL` | `FM_PIVOT_CELL_KIND_ROW_SUBTOTAL` |
| `fm_pivot_cell_t` | `fm_pivot_cell_info_t` |
| `fm_pivot_cells_at` | `fm_pivot_report_get_cell_at` |
| `fm_pivot_cells_bounds` | `fm_pivot_report_get_bounds` |
| `fm_pivot_cells_count` | `fm_pivot_report_get_cell_count` |
| `fm_pivot_cells_destroy` | `fm_pivot_report_destroy` |
| `fm_pivot_cells_t` | `fm_pivot_report_t` |
| `fm_pivot_data_field_spec_t` | `fm_pivot_data_field_t` |
| `fm_pivot_field_spec_t` | removed |
| `FM_PIVOT_FILTER_LABEL_BEGINS_WITH` | `FM_PIVOT_FILTER_KIND_LABEL_BEGINS_WITH` |
| `FM_PIVOT_FILTER_LABEL_CONTAINS` | `FM_PIVOT_FILTER_KIND_LABEL_CONTAINS` |
| `FM_PIVOT_FILTER_LABEL_DATE` | `FM_PIVOT_FILTER_KIND_DATE_BETWEEN` |
| `fm_pivot_filter_spec_t.type` | `fm_pivot_filter_t.kind` |
| `fm_pivot_filter_spec_t` | `fm_pivot_filter_t` |
| `fm_pivot_filter_type_t` | `fm_pivot_filter_kind_t` |
| `FM_PIVOT_FILTER_VALUE_BETWEEN` | `FM_PIVOT_FILTER_KIND_VALUE_BETWEEN` |
| `FM_PIVOT_FILTER_VALUE_DOUBLE` | `FM_PIVOT_FILTER_OPERAND_KIND_DOUBLE` |
| `FM_PIVOT_FILTER_VALUE_GREATER_THAN` | `FM_PIVOT_FILTER_KIND_VALUE_GREATER_THAN` |
| `FM_PIVOT_FILTER_VALUE_INT` | `FM_PIVOT_FILTER_OPERAND_KIND_INT` |
| `fm_pivot_filter_value_kind_t` | `fm_pivot_filter_operand_kind_t` |
| `FM_PIVOT_FILTER_VALUE_NONE` | `FM_PIVOT_FILTER_OPERAND_KIND_ABSENT` |
| `FM_PIVOT_FILTER_VALUE_TEXT` | `FM_PIVOT_FILTER_OPERAND_KIND_TEXT` |
| `FM_PIVOT_FILTER_VALUE_TOP_10` | `FM_PIVOT_FILTER_KIND_VALUE_TOP_COUNT` |
| `FM_PIVOT_LAYOUT_COMPACT` | `FM_PIVOT_REPORT_LAYOUT_COMPACT` |
| `FM_PIVOT_LAYOUT_OUTLINE` | `FM_PIVOT_REPORT_LAYOUT_OUTLINE` |
| `fm_pivot_layout_t` | `fm_pivot_report_layout_t` |
| `FM_PIVOT_LAYOUT_TABULAR` | `FM_PIVOT_REPORT_LAYOUT_TABULAR` |
| `FM_PIVOT_SHOW_AS_DIFFERENCE_FROM` | `FM_PIVOT_SHOW_VALUES_AS_DIFFERENCE_FROM` |
| `FM_PIVOT_SHOW_AS_INDEX` | `FM_PIVOT_SHOW_VALUES_AS_INDEX` |
| `FM_PIVOT_SHOW_AS_NORMAL` | `FM_PIVOT_SHOW_VALUES_AS_NORMAL` |
| `FM_PIVOT_SHOW_AS_PERCENT_DIFFERENCE_FROM` | `FM_PIVOT_SHOW_VALUES_AS_PERCENT_DIFFERENCE_FROM` |
| `FM_PIVOT_SHOW_AS_PERCENT_OF_COL` | `FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_COLUMN` |
| `FM_PIVOT_SHOW_AS_PERCENT_OF_PARENT_COL` | `FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_PARENT_COLUMN` |
| `FM_PIVOT_SHOW_AS_PERCENT_OF_PARENT_ROW` | `FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_PARENT_ROW` |
| `FM_PIVOT_SHOW_AS_PERCENT_OF_PARENT` | `FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_PARENT` |
| `FM_PIVOT_SHOW_AS_PERCENT_OF_ROW` | `FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_ROW` |
| `FM_PIVOT_SHOW_AS_PERCENT_OF_RUNNING_TOTAL` | `FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_RUNNING_TOTAL` |
| `FM_PIVOT_SHOW_AS_PERCENT_OF_TOTAL` | `FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_TOTAL` |
| `FM_PIVOT_SHOW_AS_PERCENT_OF` | `FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF` |
| `FM_PIVOT_SHOW_AS_RANK_ASCENDING` | `FM_PIVOT_SHOW_VALUES_AS_RANK_ASCENDING` |
| `FM_PIVOT_SHOW_AS_RANK_DESCENDING` | `FM_PIVOT_SHOW_VALUES_AS_RANK_DESCENDING` |
| `FM_PIVOT_SHOW_AS_RUNNING_TOTAL_IN_COL` | `FM_PIVOT_SHOW_VALUES_AS_RUNNING_TOTAL_IN_COLUMN` |
| `FM_PIVOT_SHOW_AS_RUNNING_TOTAL_IN_ROW` | `FM_PIVOT_SHOW_VALUES_AS_RUNNING_TOTAL_IN_ROW` |
| `fm_pivot_show_as_t` | `fm_pivot_show_values_as_t` |
| `FM_PIVOT_SORT_ASCENDING` | `FM_PIVOT_SORT_ORDER_ASCENDING` |
| `FM_PIVOT_SORT_DESCENDING` | `FM_PIVOT_SORT_ORDER_DESCENDING` |
| `FM_PIVOT_SORT_MANUAL` | `FM_PIVOT_SORT_ORDER_MANUAL` |
| `fm_pivot_style_info_t.show_col_headers` | `fm_pivot_style_options_t.show_column_headers` |
| `fm_pivot_style_info_t.show_col_stripes` | `fm_pivot_style_options_t.show_column_stripes` |
| `fm_pivot_style_info_t` | `fm_pivot_style_options_t` |
| `fm_print_options_t.grid_lines_engaged` | `fm_print_options_t.has_grid_lines` |
| `fm_print_options_t.headings_engaged` | `fm_print_options_t.has_headings` |
| `fm_print_options_t.horizontal_centered_engaged` | `fm_print_options_t.has_horizontal_centered` |
| `fm_print_options_t.vertical_centered_engaged` | `fm_print_options_t.has_vertical_centered` |
| `fm_print_range_t` | folded into `fm_cell_rect_t` |
| `fm_print_titles.has_cols` | `fm_print_titles_t.has_columns` |
| `fm_print_titles` | `fm_print_titles_t` |
| `fm_read_diagnostics_t` | `fm_read_diagnostics_info_t` |
| `fm_rect_pt` | `fm_rect_pt_info_t` |
| `fm_row_layout_t.has_height` | `fm_row_layout_info_t.has_height_pt` |
| `fm_row_layout_t.has_style` | `fm_row_layout_info_t.has_xf_index` |
| `fm_row_layout_t.style_xf` | `fm_row_layout_info_t.xf_index` |
| `fm_row_layout_t` | `fm_row_layout_info_t` |
| `fm_save_diagnostics_t` | `fm_save_diagnostics_info_t` |
| `FM_SHAPE_FIELD_ANCHOR` | removed |
| `FM_SHAPE_FIELD_FILL` | removed |
| `FM_SHAPE_FIELD_HIDDEN` | removed |
| `FM_SHAPE_FIELD_LINE_WIDTH` | removed |
| `FM_SHAPE_FIELD_LINE` | removed |
| `FM_SHAPE_FIELD_NAME` | removed |
| `FM_SHAPE_FIELD_PRST` | removed |
| `FM_SHAPE_FIELD_ROTATION` | removed |
| `FM_SHAPE_FIELD_TEXT_BOX` | removed |
| `FM_SHAPE_FIELD_TEXT` | removed |
| `FM_SHAPE_PAINT_NONE` | removed |
| `FM_SHAPE_PAINT_OTHER` | removed |
| `FM_SHAPE_PAINT_SOLID` | removed |
| `fm_shape_paint_t` | removed |
| `FM_SHAPE_PAINT_UNSET` | removed |
| `fm_shape` | `fm_shape_info_t` |
| `fm_sheet_add_col_break` | `fm_sheet_set_column_break` |
| `fm_sheet_add_hyperlink` | `fm_sheet_set_hyperlink` |
| `fm_sheet_add_row_break` | `fm_sheet_set_row_break` |
| `fm_sheet_cell_rect_pt` | `fm_sheet_get_cell_rect_pt` |
| `fm_sheet_cells_in_range` | `fm_sheet_get_cells_in_range` |
| `fm_sheet_cf_add_rule` | `fm_sheet_add_cf_rule` |
| `fm_sheet_cf_clear` | `fm_sheet_clear_cf_rules` |
| `fm_sheet_cf_count` | `fm_sheet_get_cf_rule_count` |
| `fm_sheet_cf_get_at` | `fm_sheet_get_cf_rule_at` |
| `fm_sheet_cf_remove_at` | `fm_sheet_remove_cf_rule` |
| `fm_sheet_chart_data_open` | `fm_sheet_get_chart_data` |
| `fm_sheet_chart_open` | `fm_sheet_get_chart` |
| `fm_sheet_chart_preview_open` | `fm_sheet_get_chart_preview` |
| `fm_sheet_clear_auto_filter` | `fm_sheet_clear_auto_filter_criteria` |
| `fm_sheet_clear_column_style` | `fm_sheet_reset_column_xf_index` |
| `fm_sheet_clear_column_width` | `fm_sheet_reset_column_width` |
| `fm_sheet_clear_row_height` | `fm_sheet_reset_row_height` |
| `fm_sheet_clear_row_style` | `fm_sheet_reset_row_xf_index` |
| `fm_sheet_col_break_at` | `fm_sheet_get_column_break_at` |
| `fm_sheet_col_break_count` | `fm_sheet_get_column_break_count` |
| `fm_sheet_column_chars_to_pt` | `fm_sheet_convert_column_chars_to_pt` |
| `fm_sheet_column_pt_to_chars` | `fm_sheet_convert_column_pt_to_chars` |
| `fm_sheet_column_width_pt` | `fm_sheet_get_column_width_pt` |
| `fm_sheet_data_table_at` | `fm_sheet_get_data_table` |
| `fm_sheet_drawing_object_at` | `fm_sheet_get_drawing_object_at` |
| `fm_sheet_drawing_object_count` | `fm_sheet_get_drawing_object_count` |
| `fm_sheet_edit_threaded_comment` | `fm_sheet_update_threaded_comment` |
| `fm_sheet_format_defaults.base_col_width` | `fm_sheet_format_defaults_t.base_column_width_chars` |
| `fm_sheet_format_defaults.default_col_width` | `fm_sheet_format_defaults_t.default_column_width_chars` |
| `fm_sheet_format_defaults.effective_default_col_width` | `fm_sheet_format_defaults_t.effective_default_column_width_chars` |
| `fm_sheet_format_defaults.has_default_col_width` | `fm_sheet_format_defaults_t.has_default_column_width_chars` |
| `fm_sheet_format_defaults.has_default_row_height` | `fm_sheet_format_defaults_t.has_default_row_height_pt` |
| `fm_sheet_format_defaults` | `fm_sheet_format_defaults_t` |
| `fm_sheet_get_column_count` | `fm_sheet_get_column_layout_count` |
| `fm_sheet_get_column` | `fm_sheet_get_column_layout_at` |
| `fm_sheet_get_comment_at_index` | `fm_sheet_get_comment_at` |
| `fm_sheet_get_comment_at` | `fm_sheet_get_comment` |
| `fm_sheet_get_image` | `fm_sheet_get_picture` |
| `fm_sheet_get_protection` | `fm_sheet_get_sheet_protection` |
| `fm_sheet_get_row_override_count` | `fm_sheet_get_row_layout_count` |
| `fm_sheet_get_row_override` | `fm_sheet_get_row_layout_at` |
| `FM_SHEET_HIDDEN` | `FM_SHEET_VISIBILITY_HIDDEN` |
| `fm_sheet_insert_chart` | `fm_sheet_add_chart` |
| `fm_sheet_insert_image` | `fm_sheet_add_picture` |
| `fm_sheet_insert_shape` | `fm_sheet_add_shape` |
| `fm_sheet_list_invalid_cells` | `fm_sheet_get_invalid_cells` |
| `fm_sheet_merges_in_range` | `fm_sheet_get_merges_in_range` |
| `fm_sheet_protected_range_count` | `fm_sheet_get_protected_range_count` |
| `fm_sheet_range_style_bases` | `fm_sheet_get_range_xf_bases` |
| `fm_sheet_remap_range_style` | `fm_sheet_remap_range_xfs` |
| `fm_sheet_remove_auto_filter` | removed; use `fm_sheet_set_auto_filter` |
| `fm_sheet_remove_col_break` | removed |
| `fm_sheet_remove_hyperlink` | removed; use `fm_sheet_set_hyperlink` |
| `fm_sheet_remove_row_break` | removed |
| `fm_sheet_remove_shape` | removed |
| `fm_sheet_remove_slicer` | removed |
| `fm_sheet_remove_validation_at` | `fm_sheet_remove_validation` |
| `fm_sheet_restore_range_style` | `fm_sheet_restore_range_xfs` |
| `fm_sheet_row_break_at` | `fm_sheet_get_row_break_at` |
| `fm_sheet_row_break_count` | `fm_sheet_get_row_break_count` |
| `fm_sheet_row_height_pt` | `fm_sheet_get_row_height_pt` |
| `fm_sheet_set_chart_axis` | `fm_sheet_update_chart_axis` |
| `fm_sheet_set_chart_display` | `fm_sheet_update_chart_display` |
| `fm_sheet_set_chart_plot` | `fm_sheet_update_chart_plot` |
| `fm_sheet_set_chart_point` | removed |
| `fm_sheet_set_chart_series` | `fm_sheet_update_chart_series` |
| `fm_sheet_set_column_style` | `fm_sheet_set_column_xf_index` |
| `fm_sheet_set_drawing_anchor` | removed; use `fm_sheet_update_drawing` |
| `fm_sheet_set_fit_to_page` | removed; use `fm_sheet_update_page_setup` |
| `fm_sheet_set_format_defaults` | `fm_sheet_update_format_defaults` |
| `fm_sheet_set_freeze` | removed |
| `fm_sheet_set_header_footer` | `fm_sheet_update_header_footer` |
| `fm_sheet_set_page_margins` | `fm_sheet_update_page_margins` |
| `fm_sheet_set_page_setup` | `fm_sheet_update_page_setup` |
| `fm_sheet_set_pane_top_left` | removed |
| `fm_sheet_set_print_options` | `fm_sheet_update_print_options` |
| `fm_sheet_set_protection` | `fm_sheet_set_sheet_protection` |
| `fm_sheet_set_right_to_left` | removed |
| `fm_sheet_set_row_style` | `fm_sheet_set_row_xf_index` |
| `fm_sheet_set_show_formulas` | removed |
| `fm_sheet_set_show_grid_lines` | removed |
| `fm_sheet_set_show_outline_symbols` | removed |
| `fm_sheet_set_show_row_col_headers` | removed |
| `fm_sheet_set_show_zeros` | removed |
| `fm_sheet_set_sparkline_group_sparklines` | removed; use `fm_sheet_update_sparkline_group` |
| `fm_sheet_set_split` | removed |
| `fm_sheet_set_tab_hidden` | removed |
| `fm_sheet_set_tab_selected` | removed |
| `fm_sheet_set_thread_resolved` | removed; use `fm_sheet_update_threaded_comment` |
| `fm_sheet_set_top_left` | removed |
| `fm_sheet_set_view_mode` | removed |
| `fm_sheet_set_zoom` | removed |
| `fm_sheet_shape_at` | removed |
| `fm_sheet_shape_count` | removed |
| `fm_sheet_slicer_at` | removed |
| `fm_sheet_slicer_count` | removed |
| `fm_sheet_snapshot_range_style` | `fm_sheet_snapshot_range_xfs` |
| `fm_sheet_sparkline_at` | `fm_sheet_get_sparkline_at` |
| `fm_sheet_sparkline_group_at` | `fm_sheet_get_sparkline_group_at` |
| `fm_sheet_sparkline_group_count` | `fm_sheet_get_sparkline_group_count` |
| `fm_sheet_threaded_comment_at` | `fm_sheet_get_threaded_comment_at` |
| `fm_sheet_threaded_comment_count` | `fm_sheet_get_threaded_comment_count` |
| `FM_SHEET_VERY_HIDDEN` | `FM_SHEET_VISIBILITY_VERY_HIDDEN` |
| `fm_sheet_view_t.show_row_col_headers` | `fm_sheet_view_info_t.show_row_column_headers` |
| `fm_sheet_view_t.split_x` | `fm_sheet_view_info_t.split_x_twips` |
| `fm_sheet_view_t.split_y` | `fm_sheet_view_info_t.split_y_twips` |
| `fm_sheet_view_t` | `fm_sheet_view_info_t` |
| `FM_SHEET_VISIBLE` | `FM_SHEET_VISIBILITY_VISIBLE` |
| `fm_sheet_width_model` | `fm_sheet_get_width_model` |
| `fm_slicer.style` | `fm_slicer_info_t.style_name` |
| `FM_SLICER_CROSS_NONE` | `FM_SLICER_CROSS_FILTER_NONE` |
| `FM_SLICER_CROSS_SHOW_ITEMS_WITH_DATA_AT_TOP` | `FM_SLICER_CROSS_FILTER_SHOW_ITEMS_WITH_DATA_AT_TOP` |
| `FM_SLICER_CROSS_SHOW_ITEMS_WITH_NO_DATA` | `FM_SLICER_CROSS_FILTER_SHOW_ITEMS_WITH_NO_DATA` |
| `FM_SLICER_FIELD_ANCHOR` | removed |
| `FM_SLICER_FIELD_CAPTION` | removed |
| `FM_SLICER_FIELD_COLUMN_COUNT` | removed |
| `FM_SLICER_FIELD_CROSS_FILTER` | removed |
| `FM_SLICER_FIELD_NAME` | removed |
| `FM_SLICER_FIELD_ROW_HEIGHT` | removed |
| `FM_SLICER_FIELD_SORT_ORDER` | removed |
| `FM_SLICER_FIELD_STYLE` | removed |
| `FM_SLICER_SORT_ASCENDING` | `FM_SLICER_SORT_ORDER_ASCENDING` |
| `FM_SLICER_SORT_DESCENDING` | `FM_SLICER_SORT_ORDER_DESCENDING` |
| `fm_slicer_sort_t` | `fm_slicer_sort_order_t` |
| `fm_slicer` | `fm_slicer_info_t` |
| `fm_sort_condition.dxf_id` | `fm_sort_condition_t.dxf_index` |
| `fm_sort_condition.has_dxf_id` | `fm_sort_condition_t.has_dxf_index` |
| `fm_sort_condition.has_icon_id` | `fm_sort_condition_t.has_icon_index` |
| `fm_sort_condition.icon_id` | `fm_sort_condition_t.icon_index` |
| `fm_sort_condition` | `fm_sort_condition_t` |
| `FM_SPARKLINE_AXIS_CUSTOM` | `FM_SPARKLINE_AXIS_TYPE_CUSTOM` |
| `FM_SPARKLINE_AXIS_GROUP` | `FM_SPARKLINE_AXIS_TYPE_GROUP` |
| `FM_SPARKLINE_AXIS_INDIVIDUAL` | `FM_SPARKLINE_AXIS_TYPE_INDIVIDUAL` |
| `FM_SPARKLINE_COLUMN` | `FM_SPARKLINE_TYPE_COLUMN` |
| `fm_sparkline_empty_cells_t` | folded into `fm_blanks_as_t` |
| `FM_SPARKLINE_EMPTY_GAP` | removed |
| `FM_SPARKLINE_EMPTY_SPAN` | removed |
| `FM_SPARKLINE_EMPTY_ZERO` | removed |
| `fm_sparkline_group.line_weight` | `fm_sparkline_group_t.line_width_pt` |
| `fm_sparkline_group` | `fm_sparkline_group_t` |
| `FM_SPARKLINE_LINE` | `FM_SPARKLINE_TYPE_LINE` |
| `FM_SPARKLINE_STACKED` | `FM_SPARKLINE_TYPE_STACKED` |
| `fm_sparkline` | `fm_sparkline_t` |
| `fm_status_string` | `fm_get_status_name` |
| `fm_styles_add_batch` | `fm_workbook_add_batch` |
| `fm_styles_add_border` | `fm_workbook_add_border` |
| `fm_styles_add_cell_style_xf` | `fm_workbook_add_cell_style_xf` |
| `fm_styles_add_cell_xf` | `fm_workbook_add_cell_xf` |
| `fm_styles_add_dxf` | `fm_workbook_add_dxf` |
| `fm_styles_add_fill` | `fm_workbook_add_fill` |
| `fm_styles_add_font` | `fm_workbook_add_font` |
| `fm_styles_add_num_fmt` | `fm_workbook_add_number_format` |
| `fm_styles_batch` | `fm_styles_batch_spec_t` |
| `fm_styles_get_border_count` | `fm_workbook_get_border_count` |
| `fm_styles_get_border` | `fm_workbook_get_border_at` |
| `fm_styles_get_cell_style_count` | `fm_workbook_get_cell_style_count` |
| `fm_styles_get_cell_style_xf_count` | `fm_workbook_get_cell_style_xf_count` |
| `fm_styles_get_cell_style_xf` | `fm_workbook_get_cell_style_xf_at` |
| `fm_styles_get_cell_style` | `fm_workbook_get_cell_style_at` |
| `fm_styles_get_cell_xf_count` | `fm_workbook_get_cell_xf_count` |
| `fm_styles_get_cell_xf` | `fm_workbook_get_cell_xf_at` |
| `fm_styles_get_dxf_count` | `fm_workbook_get_dxf_count` |
| `fm_styles_get_dxf` | `fm_workbook_get_dxf_at` |
| `fm_styles_get_fill_count` | `fm_workbook_get_fill_count` |
| `fm_styles_get_fill` | `fm_workbook_get_fill_at` |
| `fm_styles_get_font_count` | `fm_workbook_get_font_count` |
| `fm_styles_get_font` | `fm_workbook_get_font_at` |
| `fm_styles_get_num_fmt_string` | `fm_workbook_get_number_format` |
| `fm_styles_get_table_style_count` | `fm_workbook_get_table_style_count` |
| `fm_styles_get_table_style_defaults` | `fm_workbook_get_table_style_defaults` |
| `fm_styles_get_table_style` | `fm_workbook_get_table_style_at` |
| `fm_styles_remove_cell_style` | `fm_workbook_remove_cell_style` |
| `fm_styles_remove_table_style` | `fm_workbook_remove_table_style` |
| `fm_styles_set_cell_style` | `fm_workbook_set_cell_style` |
| `fm_styles_set_font` | `fm_workbook_set_font` |
| `fm_styles_set_table_style_defaults` | `fm_workbook_update_table_style_defaults` |
| `fm_styles_set_table_style` | `fm_workbook_set_table_style` |
| `fm_table_apply_auto_filter` | `fm_sheet_apply_table_auto_filter` |
| `fm_table_clear_auto_filter` | `fm_sheet_clear_table_auto_filter_criteria` |
| `fm_table_evaluate_auto_filter` | `fm_sheet_evaluate_table_auto_filter` |
| `fm_table_get_auto_filter` | `fm_sheet_get_table_auto_filter` |
| `fm_table_remove_auto_filter` | removed; use `fm_sheet_set_table_auto_filter` |
| `fm_table_set_auto_filter` | `fm_sheet_set_table_auto_filter` |
| `fm_table_style_element_t.has_dxf_id` | `fm_table_style_element_t.has_dxf_index` |
| `fm_text_compare` | `fm_compare_text` |
| `fm_theme_colors` | `fm_theme_colors_t` |
| `fm_theme_fonts` | `fm_theme_fonts_t` |
| `fm_threaded_comment.mention_count` | `fm_threaded_comment_t.mentions_count` |
| `fm_threaded_comment` | `fm_threaded_comment_t` |
| `FM_VAL_ARRAY` | `FM_VALUE_KIND_ARRAY` |
| `FM_VAL_BLANK` | `FM_VALUE_KIND_BLANK` |
| `FM_VAL_BOOL` | `FM_VALUE_KIND_BOOL` |
| `FM_VAL_ERROR` | `FM_VALUE_KIND_ERROR` |
| `FM_VAL_LAMBDA` | `FM_VALUE_KIND_LAMBDA` |
| `FM_VAL_NUMBER` | `FM_VALUE_KIND_NUMBER` |
| `FM_VAL_REF` | `FM_VALUE_KIND_REF` |
| `FM_VAL_TEXT` | `FM_VALUE_KIND_TEXT` |
| `fm_validation_outcome` | `fm_validation_outcome_info_t` |
| `fm_version_string` | `fm_get_version` |
| `fm_viewport` | folded into `fm_cell_rect_t` |
| `fm_width_model` | `fm_width_model_info_t` |
| `fm_workbook_calc_mode` | `fm_workbook_get_calc_mode` |
| `fm_workbook_calc_on_save` | `fm_workbook_get_calc_on_save` |
| `fm_workbook_cell_at` | `fm_sheet_get_cell_at` |
| `fm_workbook_cell_count` | `fm_sheet_get_cell_count` |
| `fm_workbook_cf_evaluate_range` | `fm_sheet_evaluate_cf_range` |
| `fm_workbook_clear_pinned_now` | removed; use `fm_workbook_set_pinned_now` |
| `fm_workbook_create` | `fm_workbook_create_default` |
| `fm_workbook_date1904` | `fm_workbook_get_date1904` |
| `fm_workbook_defined_name_at` | `fm_workbook_get_defined_name_at` |
| `fm_workbook_defined_name_attrs_at` | removed |
| `fm_workbook_defined_name_count` | `fm_workbook_get_defined_name_count` |
| `fm_workbook_delete_cols` | `fm_sheet_delete_columns` |
| `fm_workbook_delete_rows` | `fm_sheet_delete_rows` |
| `fm_workbook_dependents` | `fm_sheet_get_dependents` |
| `fm_workbook_evaluate_cf_formula` | `fm_sheet_evaluate_cf_formula` |
| `fm_workbook_evaluate_formula_array` | `fm_sheet_evaluate_formula_array` |
| `fm_workbook_evaluate_formula` | removed |
| `fm_workbook_excel_profile_id` | `fm_workbook_get_profile_id` |
| `fm_workbook_external_link_at` | `fm_workbook_get_external_link_at` |
| `fm_workbook_external_link_count` | `fm_workbook_get_external_link_count` |
| `FM_WORKBOOK_FORMAT_UNKNOWN` | removed |
| `fm_workbook_full_precision` | `fm_workbook_get_full_precision` |
| `fm_workbook_get_cell_phonetic_properties` | `fm_sheet_get_cell_phonetic_properties` |
| `fm_workbook_get_cell_phonetic_run_count` | `fm_sheet_get_cell_phonetic_run_count` |
| `fm_workbook_get_cell_phonetic_run` | `fm_sheet_get_cell_phonetic_run_at` |
| `fm_workbook_get_cell_phonetic` | `fm_sheet_get_cell_phonetic` |
| `fm_workbook_get_cell_show_phonetic` | `fm_sheet_get_cell_show_phonetic` |
| `fm_workbook_get_display_text` | `fm_sheet_get_display_text` |
| `fm_workbook_get_formula_r1c1` | `fm_sheet_get_formula_r1c1` |
| `fm_workbook_get_formula` | `fm_sheet_get_formula` |
| `fm_workbook_get_protection` | `fm_workbook_get_workbook_protection` |
| `fm_workbook_get_value` | `fm_sheet_get_value` |
| `fm_workbook_has_cell_phonetic_properties` | removed; use `fm_sheet_get_cell_phonetic_properties` |
| `fm_workbook_insert_cols` | `fm_sheet_insert_columns` |
| `fm_workbook_insert_rows` | `fm_sheet_insert_rows` |
| `fm_workbook_lambda_text_at` | `fm_sheet_get_lambda_text` |
| `fm_workbook_memory_usage` | `fm_workbook_get_memory_usage` |
| `fm_workbook_paginate` | `fm_sheet_paginate` |
| `fm_workbook_partial_recalc` | `fm_sheet_recalc_viewport` |
| `fm_workbook_passthrough_at` | `fm_workbook_get_passthrough_at` |
| `fm_workbook_passthrough_count` | `fm_workbook_get_passthrough_count` |
| `fm_workbook_person_at` | `fm_workbook_get_person_at` |
| `fm_workbook_person_count` | `fm_workbook_get_person_count` |
| `fm_workbook_pinned_now` | `fm_workbook_get_pinned_now` |
| `fm_workbook_pivot_cache_count` | `fm_workbook_get_pivot_cache_count` |
| `fm_workbook_pivot_cache_create` | `fm_workbook_add_pivot_cache` |
| `fm_workbook_pivot_cache_field_add_shared_item_blank` | `fm_workbook_add_pivot_cache_field_shared_item_blank` |
| `fm_workbook_pivot_cache_field_add_shared_item_bool` | `fm_workbook_add_pivot_cache_field_shared_item_bool` |
| `fm_workbook_pivot_cache_field_add_shared_item_error` | `fm_workbook_add_pivot_cache_field_shared_item_error` |
| `fm_workbook_pivot_cache_field_add_shared_item_number` | `fm_workbook_add_pivot_cache_field_shared_item_number` |
| `fm_workbook_pivot_cache_field_add_shared_item_text` | `fm_workbook_add_pivot_cache_field_shared_item_text` |
| `fm_workbook_pivot_cache_field_add` | `fm_workbook_add_pivot_cache_field` |
| `fm_workbook_pivot_cache_field_clear_shared_items` | `fm_workbook_clear_pivot_cache_field_shared_items` |
| `fm_workbook_pivot_cache_field_clear` | `fm_workbook_clear_pivot_cache_fields` |
| `fm_workbook_pivot_cache_field_count` | `fm_workbook_get_pivot_cache_field_count` |
| `fm_workbook_pivot_cache_field_name` | `fm_workbook_get_pivot_cache_field_name` |
| `fm_workbook_pivot_cache_field_shared_item_count` | `fm_workbook_get_pivot_cache_field_shared_item_count` |
| `fm_workbook_pivot_cache_get_worksheet_source` | `fm_workbook_get_pivot_cache_worksheet_source` |
| `fm_workbook_pivot_cache_id_at` | `fm_workbook_get_pivot_cache_id_at` |
| `fm_workbook_pivot_cache_record_add` | `fm_workbook_add_pivot_cache_record` |
| `fm_workbook_pivot_cache_record_clear` | `fm_workbook_clear_pivot_cache_records` |
| `fm_workbook_pivot_cache_record_count` | `fm_workbook_get_pivot_cache_record_count` |
| `fm_workbook_pivot_cache_record_set_blank` | `fm_workbook_set_pivot_cache_record_blank` |
| `fm_workbook_pivot_cache_record_set_bool` | `fm_workbook_set_pivot_cache_record_bool` |
| `fm_workbook_pivot_cache_record_set_error` | `fm_workbook_set_pivot_cache_record_error` |
| `fm_workbook_pivot_cache_record_set_number` | `fm_workbook_set_pivot_cache_record_number` |
| `fm_workbook_pivot_cache_record_set_text` | `fm_workbook_set_pivot_cache_record_text` |
| `fm_workbook_pivot_cache_remove` | `fm_workbook_remove_pivot_cache` |
| `fm_workbook_pivot_cache_set_worksheet_source` | `fm_workbook_set_pivot_cache_worksheet_source` |
| `fm_workbook_pivot_count` | `fm_sheet_get_pivot_count` |
| `fm_workbook_pivot_create` | `fm_sheet_add_pivot` |
| `fm_workbook_pivot_data_field_add` | `fm_sheet_add_pivot_data_field` |
| `fm_workbook_pivot_data_field_at` | `fm_sheet_get_pivot_data_field_at` |
| `fm_workbook_pivot_data_field_clear` | `fm_sheet_clear_pivot_data_fields` |
| `fm_workbook_pivot_data_field_count` | `fm_sheet_get_pivot_data_field_count` |
| `fm_workbook_pivot_data_field_set` | `fm_sheet_set_pivot_data_field` |
| `fm_workbook_pivot_field_add_item_at` | `fm_sheet_add_pivot_field_item_by_cache_index` |
| `fm_workbook_pivot_field_add_item` | `fm_sheet_add_pivot_field_item` |
| `fm_workbook_pivot_field_add_subtotal_fn` | removed; use `fm_sheet_set_pivot_field_subtotal_functions` |
| `fm_workbook_pivot_field_add` | removed |
| `fm_workbook_pivot_field_at` | `fm_sheet_get_pivot_field_at` |
| `fm_workbook_pivot_field_clear_date_group` | removed; use `fm_sheet_set_pivot_field_date_group` |
| `fm_workbook_pivot_field_clear_items` | `fm_sheet_clear_pivot_field_items` |
| `fm_workbook_pivot_field_clear_subtotal_fns` | removed; use `fm_sheet_set_pivot_field_subtotal_functions` |
| `fm_workbook_pivot_field_clear` | `fm_sheet_clear_pivot_fields` |
| `fm_workbook_pivot_field_count` | `fm_sheet_get_pivot_field_count` |
| `fm_workbook_pivot_field_item_at` | `fm_sheet_get_pivot_field_item_at` |
| `fm_workbook_pivot_field_item_count` | `fm_sheet_get_pivot_field_item_count` |
| `fm_workbook_pivot_field_set_axis` | `fm_sheet_set_pivot_field_axis` |
| `fm_workbook_pivot_field_set_date_group` | `fm_sheet_set_pivot_field_date_group` |
| `fm_workbook_pivot_field_set_item_visible` | removed |
| `fm_workbook_pivot_field_set_number_format` | removed |
| `fm_workbook_pivot_field_set_sort` | removed |
| `fm_workbook_pivot_field_set_subtotal_top` | removed |
| `fm_workbook_pivot_filter_add` | `fm_sheet_add_pivot_filter` |
| `fm_workbook_pivot_filter_at` | `fm_sheet_get_pivot_filter_at` |
| `fm_workbook_pivot_filter_clear` | `fm_sheet_clear_pivot_filters` |
| `fm_workbook_pivot_filter_count` | `fm_sheet_get_pivot_filter_count` |
| `fm_workbook_pivot_filter_remove_at` | `fm_sheet_remove_pivot_filter` |
| `fm_workbook_pivot_get_cache_id` | `fm_sheet_get_pivot_cache_id` |
| `fm_workbook_pivot_get_col_field_order` | `fm_sheet_get_pivot_column_field_order` |
| `fm_workbook_pivot_get_grand_totals` | `fm_sheet_get_pivot_grand_totals` |
| `fm_workbook_pivot_get_layout` | `fm_sheet_get_pivot_report_layout` |
| `fm_workbook_pivot_get_name` | `fm_sheet_get_pivot_name` |
| `fm_workbook_pivot_get_row_field_order` | `fm_sheet_get_pivot_row_field_order` |
| `fm_workbook_pivot_get_style_info` | `fm_sheet_get_pivot_style_options` |
| `fm_workbook_pivot_layout` | `fm_sheet_get_pivot_layout` |
| `fm_workbook_pivot_remove` | `fm_sheet_remove_pivot` |
| `fm_workbook_pivot_set_anchor` | `fm_sheet_set_pivot_anchor` |
| `fm_workbook_pivot_set_col_field_order` | `fm_sheet_set_pivot_column_field_order` |
| `fm_workbook_pivot_set_grand_totals` | `fm_sheet_update_pivot_grand_totals` |
| `fm_workbook_pivot_set_layout` | `fm_sheet_set_pivot_report_layout` |
| `fm_workbook_pivot_set_name` | `fm_sheet_rename_pivot` |
| `fm_workbook_pivot_set_row_field_order` | `fm_sheet_set_pivot_row_field_order` |
| `fm_workbook_pivot_set_style_info` | `fm_sheet_update_pivot_style_options` |
| `fm_workbook_precedents` | `fm_sheet_get_precedents` |
| `fm_workbook_probe_image` | `fm_get_image_info` |
| `fm_workbook_read_diagnostics` | `fm_workbook_get_diagnostics` |
| `fm_workbook_ref_mode` | `fm_workbook_get_ref_mode` |
| `fm_workbook_save_as` | removed; use `fm_workbook_save` |
| `fm_workbook_set_blank` | `fm_sheet_set_blank` |
| `fm_workbook_set_bool` | `fm_sheet_set_bool` |
| `fm_workbook_set_cell_phonetic_properties` | `fm_sheet_set_cell_phonetic_properties` |
| `fm_workbook_set_cell_phonetic_runs` | `fm_sheet_set_cell_phonetic_runs` |
| `fm_workbook_set_cell_phonetic` | `fm_sheet_set_cell_phonetic` |
| `fm_workbook_set_cell_show_phonetic` | `fm_sheet_set_cell_show_phonetic` |
| `fm_workbook_set_defined_name_attrs` | removed |
| `fm_workbook_set_defined_name_scoped` | removed |
| `fm_workbook_set_error` | `fm_sheet_set_error` |
| `fm_workbook_set_excel_profile_id` | `fm_workbook_set_profile_id` |
| `fm_workbook_set_formula` | `fm_sheet_set_formula` |
| `fm_workbook_set_iterative_enabled` | removed |
| `fm_workbook_set_iterative` | `fm_workbook_update_iterative` |
| `fm_workbook_set_number` | `fm_sheet_set_number` |
| `fm_workbook_set_protection` | `fm_workbook_set_workbook_protection` |
| `fm_workbook_set_text` | `fm_sheet_set_text` |
| `fm_workbook_sheet_count` | `fm_workbook_get_sheet_count` |
| `fm_workbook_sheet_name` | `fm_workbook_get_sheet_name` |
| `fm_workbook_spill_info` | `fm_sheet_get_spill` |
| `fm_workbook_table_at` | `fm_sheet_get_table_at` |
| `fm_workbook_table_count` | `fm_sheet_get_table_count` |
| `fm_workbook_table_create` | `fm_sheet_add_table` |
| `fm_workbook_table_get_column` | `fm_sheet_get_table_column_at` |
| `fm_workbook_table_get_info` | removed |
| `fm_workbook_table_remove` | `fm_sheet_remove_table` |
| `fm_workbook_table_rename` | `fm_sheet_rename_table` |
| `fm_workbook_table_set_column_totals` | removed; use `fm_sheet_update_table` |
| `fm_workbook_table_set_style_info` | `fm_sheet_update_table_style_options` |
| `fm_workbook_table_update` | `fm_sheet_update_table` |
| `fm_workbook_theme_script_font_at` | `fm_workbook_get_theme_script_font_at` |
| `fm_workbook_theme_script_font_count` | `fm_workbook_get_theme_script_font_count` |
| `kFmColorAuto` | `FM_COLOR_KIND_AUTO` |
| `kFmColorIndexed` | `FM_COLOR_KIND_INDEXED` |
| `kFmColorNone` | `FM_COLOR_KIND_NONE` |
| `kFmColorRgb` | `FM_COLOR_KIND_RGB` |
| `kFmColorTheme` | `FM_COLOR_KIND_THEME` |

#### JavaScript (`@libraz/formulon`, `@libraz/formulon-native`)

| Before | After |
|---|---|
| `addChartSeries()` | `ws.chart(...).addSeries()` |
| `AddChartSeriesResult` | removed |
| `addConditionalFormat()` | `addCfRule()` |
| `addHyperlink()` | `setHyperlink()` |
| `addHyperlinkRange()` | removed; use `setHyperlink()` |
| `addNumFmt()` | `addNumberFormat()` |
| `AddNumFmtResult` | removed |
| `addSheetColBreak()` | `setColumnBreak()` |
| `addSheetRowBreak()` | `setRowBreak()` |
| `AddStyleResult` | removed |
| `addXf()` | `addCellXf()` |
| `applyTableAutoFilter()` | `ws.table(...).applyAutoFilter()` |
| `AutoFilterColumn.colId` | `AutoFilterColumn.columnOffset` |
| `AutoFilterColumn.dxfId` | `AutoFilterColumn.dxfIndex` |
| `AutoFilterColumn.dynMaxVal` | `AutoFilterColumn.dynamicMaxValue` |
| `AutoFilterColumn.dynVal` | `AutoFilterColumn.dynamicValue` |
| `AutoFilterColumn.filterVal` | `AutoFilterColumn.filterValue` |
| `AutoFilterColumn.iconId` | `AutoFilterColumn.iconIndex` |
| `AutoFilterColumn.maxValIso` | `AutoFilterColumn.maxValueIso` |
| `AutoFilterColumn.topVal` | `AutoFilterColumn.topValue` |
| `AutoFilterColumn.valIso` | `AutoFilterColumn.valueIso` |
| `AutoFilterEvalResult` | `AutoFilterEval` |
| `AutoFilterResult` | removed |
| `AutoFilterSortCondition.dxfId` | `AutoFilterSortCondition.dxfIndex` |
| `AutoFilterSortCondition.iconId` | `AutoFilterSortCondition.iconIndex` |
| `AutoFilterSortCondition.ref` | `AutoFilterSortCondition.range` |
| `AutoFilterSortState.ref` | `AutoFilterSortState.range` |
| `borderCount()` | removed; use `listBorders()` |
| `BorderRecord` | `Border` |
| `BorderResult` | removed |
| `calcMode()` | `getCalcMode()` |
| `calcOnSave()` | `getCalcOnSave()` |
| `cellAt()` | removed |
| `cellCount()` | removed |
| `CellEntry` | folded into `RangeCell` |
| `CellRectResult` | removed |
| `CellResult` | removed |
| `CellShowPhoneticResult.show` | `value` |
| `CellShowPhoneticResult` | removed |
| `CellsInRangeResult` | `CellsInRange` |
| `cellStyleCount()` | removed; use `listCellStyles()` |
| `CellStyleRecord.iLevel` | `CellStyle.outlineLevel` |
| `CellStyleRecord.xfId` | `CellStyle.styleXfIndex` |
| `CellStyleRecord` | `CellStyle` |
| `cellStyleXfCount()` | removed; use `listCellStyleXfs()` |
| `CellXf.horizontalAlign` | `CellXf.horizontal` |
| `CellXf.verticalAlign` | `CellXf.vertical` |
| `CellXf.xfId` | `CellXf.styleXfIndex` |
| `CellXfIndexResult.xfIndex` | `value` |
| `CellXfIndexResult` | removed |
| `CellXfResult` | folded into `CellXf` |
| `CfCellResult` | `CfCell` |
| `CfColor` | folded into `Color` |
| `CfIconInput.iconId` | `CfIconInput.iconIndex` |
| `CfMatch.dxfId` | `CfMatch.dxfIndex` |
| `CfRangeResult` | removed |
| `ChartAxisInfo.deleted` | `ChartAxis.visible` |
| `ChartAxisInfo` | `ChartAxis` |
| `ChartBlanksAs` | `BlanksAs` |
| `ChartCategories.formatCode` | `ChartCategories.numberFormat` |
| `ChartDataResult` | removed |
| `ChartDataSeries.idx` | `ChartDataSeries.seriesId` |
| `ChartFillInfo.transparency` | `ChartFill.transparencyPct` |
| `ChartFillInfo` | `ChartFill` |
| `ChartFillPatch.transparency` | `ChartFillPatch.transparencyPct` |
| `ChartInfo.isChartex` | `ChartProperties.isExtended` |
| `ChartInfo` | `ChartProperties` |
| `ChartInsertKind` | folded into `ChartPlotType` |
| `ChartInsertOptions.source` | `ChartInput.range` |
| `ChartInsertOptions` | `ChartInput` |
| `ChartLabelsInfo` | `ChartLabels` |
| `ChartLineInfo.transparency` | `ChartLine.transparencyPct` |
| `ChartLineInfo` | `ChartLine` |
| `ChartLinePatch.transparency` | `ChartLinePatch.transparencyPct` |
| `ChartMarkerInfo.size` | `ChartMarker.sizePt` |
| `ChartMarkerInfo` | `ChartMarker` |
| `ChartMarkerPatch.size` | `ChartMarkerPatch.sizePt` |
| `ChartPlotInfo.chartexLayout` | `ChartPlot.extendedLayout` |
| `ChartPlotInfo.firstSliceAngle` | `ChartPlot.firstSliceAngleDeg` |
| `ChartPlotInfo.gapWidth` | `ChartPlot.gapWidthPct` |
| `ChartPlotInfo.holeSize` | `ChartPlot.holeSizePct` |
| `ChartPlotInfo.overlap` | `ChartPlot.overlapPct` |
| `ChartPlotInfo` | `ChartPlot` |
| `ChartPlotPatch.firstSliceAngle` | `ChartPlotPatch.firstSliceAngleDeg` |
| `ChartPlotPatch.gapWidth` | `ChartPlotPatch.gapWidthPct` |
| `ChartPlotPatch.holeSize` | `ChartPlotPatch.holeSizePct` |
| `ChartPlotPatch.overlap` | `ChartPlotPatch.overlapPct` |
| `ChartPlotType.Chartex` | `ChartPlotType.Extended` |
| `ChartPointInfo` | `ChartPoint` |
| `ChartPreviewOverrides.source` | `ChartPreviewOverridesInput.range` |
| `ChartPreviewOverrides` | `ChartPreviewOverridesInput` |
| `ChartResult.chart` | `value` |
| `ChartResult` | removed |
| `ChartSeriesInfo.idx` | `ChartSeries.seriesId` |
| `ChartSeriesInfo` | `ChartSeries` |
| `ChartSheetRange` | folded into `SheetRect` |
| `ChartTitleInfo` | `ChartTitle` |
| `ChartUnresolved.seriesIdx` | `ChartUnresolved.seriesId` |
| `clearAutoFilter()` | `clearAutoFilterCriteria()` |
| `clearCellXfIndex()` | `resetCellXfIndex()` |
| `clearColumnStyle()` | `resetColumnXfIndex()` |
| `clearColumnWidth()` | `resetColumnWidth()` |
| `clearConditionalFormats()` | `clearCfRules()` |
| `clearPinnedNow()` | removed; use `setPinnedNow()` |
| `clearRowHeight()` | `resetRowHeight()` |
| `clearRowStyle()` | `resetRowXfIndex()` |
| `clearSheetBreaks()` | `clearBreaks()` |
| `clearTableAutoFilter()` | `ws.table(...).clearAutoFilterCriteria()` |
| `collationClasses()` | `getCollationClasses()` |
| `CollationClassesResult` | `CollationClasses` |
| `ColorSpec.rgb` | `Color.argb` |
| `ColorSpec` | `Color` |
| `columnCharsToPt()` | `convertColumnCharsToPt()` |
| `ColumnLayout.width` | `ColumnLayout.widthChars` |
| `columnPtToChars()` | `convertColumnPtToChars()` |
| `ColumnsResult` | removed |
| `CommentEntry` | `Comment` |
| `CommentResult.comment` | `value` |
| `CommentResult` | removed |
| `CompareTextResult.order` | `value` |
| `CompareTextResult` | removed |
| `ConditionalFormatEntry.dxfId` | `ConditionalFormat.dxfIndex` |
| `ConditionalFormatEntry.op` | `ConditionalFormat.operator` |
| `ConditionalFormatEntry` | `ConditionalFormat` |
| `ConditionalFormatInput.dxfId` | `ConditionalFormatInput.dxfIndex` |
| `ConditionalFormatInput.op` | `ConditionalFormatInput.operator` |
| `ConditionalFormatRange` | folded into `CellRect` |
| `createTable()` | `addTable()` |
| `DataTable.colInput` | `DataTable.columnInput` |
| `DataTable.rect` | `DataTable.range` |
| `DataTableResult.dataTable` | `value` |
| `DataTableResult` | removed |
| `DataTableSpec.rect` | `DataTableInput.range` |
| `DataTableSpec` | `DataTableInput` |
| `DataValidationEntry.op` | `DataValidation.operator` |
| `DataValidationEntry.showDropDown` | `DataValidation.showInCellDropdown` |
| `DataValidationEntry` | `DataValidation` |
| `DataValidationInput.op` | `DataValidationInput.operator` |
| `DataValidationInput.showDropDown` | `DataValidationInput.showInCellDropdown` |
| `DataValidationRange` | folded into `CellRect` |
| `date1904()` | `getDate1904()` |
| `DefinedName.localSheetId` | `DefinedName.scope` |
| `definedNameAt()` | removed; use `listDefinedNames()` |
| `definedNameCount()` | removed; use `listDefinedNames()` |
| `delete()` | removed |
| `deleteCols()` | `deleteColumns()` |
| `dependents()` | `getDependents()` |
| `DisplayTextResult` | `DisplayText` |
| `DrawingObject` | `DrawingProperties` |
| `dxfCount()` | removed; use `listDxfs()` |
| `DxfRecord.alignmentXml` | `Dxf.alignment` |
| `DxfRecord.numFmt` | `Dxf.numberFormat` |
| `DxfRecord.protectionXml` | `Dxf.protection` |
| `DxfRecord` | `Dxf` |
| `DxfResult` | removed |
| `editThreadedComment()` | `updateThreadedComment()` |
| `EffectiveStyleResult.font` | `EffectiveStyle.fontColor` |
| `EffectiveStyleResult.numFmtCode` | `EffectiveStyle.numberFormat` |
| `EffectiveStyleResult` | `EffectiveStyle` |
| `errorDisplayName()` | removed |
| `EvalArrayResult` | removed; evaluateFormulaArray returns the rows directly |
| `evalFormula()` | `evaluateFormula()` |
| `EvalResult` | removed |
| `evaluateConditionalFormula()` | `evaluateCfFormula()` |
| `evaluateFormulaText()` | `evaluateFormula()` |
| `evaluateTableAutoFilter()` | `ws.table(...).evaluateAutoFilter()` |
| `excelProfileId()` | `getProfileId()` |
| `ExcelProfileId` | `ProfileId` |
| `ExternalLinkRecord.index` | `ExternalLink.bookNumber` |
| `ExternalLinkRecord` | `ExternalLink` |
| `fillCount()` | removed; use `listFills()` |
| `FillRecord.bg` | `Fill.background` |
| `FillRecord.fg` | `Fill.foreground` |
| `FillRecord` | `Fill` |
| `FillResult` | removed |
| `FilterOperator` | `ComparisonOperator` |
| `fontCount()` | removed; use `listFonts()` |
| `FontRecord.size` | `Font.sizePt` |
| `FontRecord` | `Font` |
| `FontResult` | removed |
| `FormulaResult.formula` | `value` |
| `FormulaResult` | removed |
| `FormulonModule` | `Formulon` |
| `FormulonModuleOptions` | `FormulonOptions` |
| `fullPrecision()` | `getFullPrecision()` |
| `functionMetadata()` | `getFunctionMetadata()` |
| `FunctionMetadata.signatureTemplate` | `FunctionMetadata.signature` |
| `FunctionMetadataResult` | removed |
| `functionNames()` | `listFunctionNames()` |
| `getChart()` | `ws.chart(...).getProperties()` |
| `getChartData()` | `ws.chart(...).getData()` |
| `getCommentResult()` | removed; use `getComment()` |
| `getComments()` | `listComments()` |
| `getConditionalFormats()` | removed |
| `getExternalLinks()` | removed |
| `getHyperlinks()` | `listHyperlinks()` |
| `getImage()` | `ws.picture(...).getProperties()` |
| `getMerges()` | `listMerges()` |
| `getNumFmt()` | `getNumberFormat()` |
| `getPersons()` | `listPeople()` |
| `getProtectedRanges()` | removed |
| `getSheetAutoFilterXml()` | `getAutoFilterXml()` |
| `getSheetColBreaks()` | `listColumnBreaks()` |
| `getSheetColumns()` | `listColumnOverrides()` |
| `getSheetFormatDefaults()` | `getFormatDefaults()` |
| `getSheetHeaderFooterXml()` | `getHeaderFooterXml()` |
| `getSheetPageMargins()` | `getPageMargins()` |
| `getSheetPageMarginsXml()` | `getPageMarginsXml()` |
| `getSheetPageSetup()` | `getPageSetup()` |
| `getSheetPageSetupXml()` | `getPageSetupXml()` |
| `getSheetPrintArea()` | `getPrintArea()` |
| `getSheetPrintOptionsXml()` | `getPrintOptionsXml()` |
| `getSheetPrintTitles()` | `getPrintTitles()` |
| `getSheetRowBreaks()` | `listRowBreaks()` |
| `getSheetRowOverrides()` | `listRowOverrides()` |
| `getSheetSelection()` | `getSelection()` |
| `getSheetSheetPrXml()` | `getSheetPrXml()` |
| `getSheetView()` | `getView()` |
| `getTableAutoFilter()` | `ws.table(...).getAutoFilter()` |
| `getTheme()` | `getThemeFonts()` |
| `getThreadedComments()` | `listThreadedComments()` |
| `getValidations()` | removed |
| `GradientFill.degree` | `GradientFill.angleDeg` |
| `HyperlinkEntry.rect` | `Hyperlink.range` |
| `HyperlinkEntry` | `Hyperlink` |
| `ImageInfo.pxHeight` | `ImageInfo.heightPx` |
| `ImageInfo.pxWidth` | `ImageInfo.widthPx` |
| `ImageInfoResult` | folded into `ImageInfo` |
| `ImageInsertOptions` | `DrawingPlacementInput` |
| `ImagePlacement` | folded into `DrawingAnchor` |
| `ImageResult` | folded into `PictureProperties` |
| `ImageSnapshotResult` | removed |
| `insertChart()` | `addChart()` |
| `InsertChartResult` | removed |
| `insertCols()` | `insertColumns()` |
| `insertImage()` | `addPicture()` |
| `InsertImageResult` | removed |
| `insertShape()` | `addShape()` |
| `isValid()` | `disposed()` |
| `IterativeSettingsResult` | `IterativeSettings` |
| `LambdaTextResult.text` | `value` |
| `LambdaTextResult` | removed |
| `lastErrorContext()` | removed |
| `lastErrorMessage()` | removed |
| `listDrawingObjects()` | `listDrawings()` |
| `listInvalidCells()` | `getInvalidCells()` |
| `ListResult` | removed |
| `localeFacts()` | `getLocaleFacts()` |
| `LocaleFactsResult` | removed |
| `memoryUsage()` | `getMemoryUsage()` |
| `Mention.length` | `Mention.end` |
| `MergedFunctionMetadataResult` | removed |
| `MergeRange` | folded into `CellRect` |
| `moveSheet()` | `wb.sheet(...).move()` |
| `NumberResult` | removed |
| `NumFmtResult` | removed |
| `Orientation` | folded into `PageOrientation` |
| `PageBreak.id` | `RowBreak.row` |
| `PageBreak.max` | `RowBreak.last` |
| `PageBreak.min` | `RowBreak.first` |
| `PageBreak` | `RowBreak` |
| `PageMargins.bottom` | `PageMargins.bottomIn` |
| `PageMargins.bottomStated` | `PageMargins.bottomInStated` |
| `PageMargins.footer` | `PageMargins.footerIn` |
| `PageMargins.footerStated` | `PageMargins.footerInStated` |
| `PageMargins.header` | `PageMargins.headerIn` |
| `PageMargins.headerStated` | `PageMargins.headerInStated` |
| `PageMargins.left` | `PageMargins.leftIn` |
| `PageMargins.leftStated` | `PageMargins.leftInStated` |
| `PageMargins.right` | `PageMargins.rightIn` |
| `PageMargins.rightStated` | `PageMargins.rightInStated` |
| `PageMargins.top` | `PageMargins.topIn` |
| `PageMargins.topStated` | `PageMargins.topInStated` |
| `PageMarginsInput.bottom` | `PageMarginsInput.bottomIn` |
| `PageMarginsInput.footer` | `PageMarginsInput.footerIn` |
| `PageMarginsInput.header` | `PageMarginsInput.headerIn` |
| `PageMarginsInput.left` | `PageMarginsInput.leftIn` |
| `PageMarginsInput.right` | `PageMarginsInput.rightIn` |
| `PageMarginsInput.top` | `PageMarginsInput.topIn` |
| `PageSetup.scale` | `PageSetup.scalePct` |
| `PageSetup.scaleStated` | `PageSetup.scalePctStated` |
| `PageSetupInput.scale` | `PageSetupInput.scalePct` |
| `paginate()` | `getPagination()` |
| `PaginationMargins.bottom` | `PaginationMargins.bottomIn` |
| `PaginationMargins.footer` | `PaginationMargins.footerIn` |
| `PaginationMargins.header` | `PaginationMargins.headerIn` |
| `PaginationMargins.left` | `PaginationMargins.leftIn` |
| `PaginationMargins.right` | `PaginationMargins.rightIn` |
| `PaginationMargins.top` | `PaginationMargins.topIn` |
| `PaginationPage.rect` | `PaginationPage.range` |
| `PaginationPrintTitles` | folded into `PrintTitles` |
| `PaginationResult.horizontalBreaks` | `Pagination.rowBreaks` |
| `PaginationResult.scale` | `Pagination.scalePct` |
| `PaginationResult.verticalBreaks` | `Pagination.columnBreaks` |
| `PaginationResult` | `Pagination` |
| `ParallelRecalcResult` | removed |
| `partialRecalc()` | `recalcViewport()` |
| `PartialRecalcResult.recomputed` | `value` |
| `PartialRecalcResult` | removed |
| `passthroughAt()` | removed; use `listPassthroughs()` |
| `passthroughCount()` | removed; use `listPassthroughs()` |
| `PersonsResult` | removed |
| `PhoneticProperties.fontId` | `PhoneticProperties.fontIndex` |
| `PhoneticPropertiesResult` | folded into `PhoneticProperties` |
| `PhoneticRun.eb` | `PhoneticRun.end` |
| `PhoneticRun.sb` | `PhoneticRun.start` |
| `pinnedNow()` | `getPinnedNow()` |
| `PinnedNowResult` | removed |
| `PivotAxis.Col` | `PivotAxis.Column` |
| `pivotCacheCount()` | removed |
| `pivotCacheCreate()` | `addPivotCache()` |
| `pivotCacheFieldAdd()` | `wb.pivotCache(...).addField()` |
| `pivotCacheFieldAddSharedItemBlank()` | removed |
| `pivotCacheFieldAddSharedItemBool()` | removed |
| `pivotCacheFieldAddSharedItemError()` | removed |
| `pivotCacheFieldAddSharedItemNumber()` | removed |
| `pivotCacheFieldAddSharedItemText()` | removed |
| `pivotCacheFieldClear()` | `wb.pivotCache(...).clearFields()` |
| `pivotCacheFieldClearSharedItems()` | `wb.pivotCache(...).clearFieldSharedItems()` |
| `pivotCacheFieldCount()` | removed |
| `pivotCacheFieldName()` | removed |
| `pivotCacheFieldSharedItemCount()` | removed |
| `pivotCacheGetWorksheetSource()` | `wb.pivotCache(...).getWorksheetSource()` |
| `pivotCacheId()` | `ws.pivot(...).cache()` |
| `pivotCacheIdAt()` | removed |
| `pivotCacheRecordAdd()` | `wb.pivotCache(...).addRecord()` |
| `pivotCacheRecordClear()` | `wb.pivotCache(...).clearRecords()` |
| `pivotCacheRecordCount()` | `wb.pivotCache(...).getRecordCount()` |
| `pivotCacheRecordSetBlank()` | removed |
| `pivotCacheRecordSetBool()` | removed |
| `pivotCacheRecordSetError()` | removed |
| `pivotCacheRecordSetNumber()` | removed |
| `pivotCacheRecordSetText()` | removed |
| `pivotCacheRemove()` | `removePivotCache()` |
| `pivotCacheSetWorksheetSource()` | `wb.pivotCache(...).setWorksheetSource()` |
| `PivotCellKind.ColLabel` | `PivotCellKind.ColumnLabel` |
| `PivotCellKind.ColSubtotal` | `PivotCellKind.ColumnSubtotal` |
| `pivotColFieldOrder()` | `ws.pivot(...).getColumnFieldOrder()` |
| `pivotCount()` | removed |
| `pivotCreate()` | `addPivot()` |
| `pivotDataFieldAdd()` | `ws.pivot(...).addDataField()` |
| `pivotDataFieldAt()` | `ws.pivot(...).getDataField()` |
| `pivotDataFieldClear()` | `ws.pivot(...).clearDataFields()` |
| `pivotDataFieldCount()` | removed; use `ws.pivot(...).listDataFields()` |
| `PivotDataFieldResult` | `PivotDataField` |
| `pivotDataFieldSet()` | `ws.pivot(...).setDataField()` |
| `PivotDataFieldSpec` | `PivotDataFieldInput` |
| `pivotFieldAdd()` | removed |
| `pivotFieldAddItem()` | `ws.pivot(...).addFieldItem()` |
| `pivotFieldAddItemAt()` | removed |
| `pivotFieldAddSubtotalFn()` | removed; use `ws.pivot(...).setFieldSubtotalFunctions()` |
| `pivotFieldAt()` | `ws.pivot(...).getField()` |
| `pivotFieldClear()` | `ws.pivot(...).clearFields()` |
| `pivotFieldClearDateGroup()` | removed; use `ws.pivot(...).setFieldDateGroup()` |
| `pivotFieldClearItems()` | `ws.pivot(...).clearFieldItems()` |
| `pivotFieldClearSubtotalFns()` | removed; use `ws.pivot(...).setFieldSubtotalFunctions()` |
| `pivotFieldCount()` | removed; use `ws.pivot(...).listFields()` |
| `pivotFieldItemAt()` | `ws.pivot(...).getFieldItem()` |
| `pivotFieldItemCount()` | removed; use `ws.pivot(...).listFieldItems()` |
| `PivotFieldOrderResult` | removed |
| `PivotFieldResult` | folded into `PivotField` |
| `pivotFieldSetAxis()` | `ws.pivot(...).setFieldAxis()` |
| `pivotFieldSetDateGroup()` | `ws.pivot(...).setFieldDateGroup()` |
| `pivotFieldSetItemVisible()` | removed |
| `pivotFieldSetNumberFormat()` | removed |
| `pivotFieldSetSort()` | removed |
| `pivotFieldSetSubtotalTop()` | removed |
| `pivotFilterAdd()` | `ws.pivot(...).addFilter()` |
| `pivotFilterAt()` | `ws.pivot(...).getFilter()` |
| `pivotFilterClear()` | `ws.pivot(...).clearFilters()` |
| `pivotFilterCount()` | removed; use `ws.pivot(...).listFilters()` |
| `pivotFilterRemoveAt()` | `ws.pivot(...).removeFilter()` |
| `PivotFilterResult` | `PivotFilter` |
| `PivotFilterSpec.type` | `PivotFilterInput.kind` |
| `PivotFilterSpec` | `PivotFilterInput` |
| `PivotFilterType.LabelDate` | `PivotFilterKind.DateBetween` |
| `PivotFilterType.ValueTop10` | `PivotFilterKind.ValueTopCount` |
| `PivotFilterType` | `PivotFilterKind` |
| `PivotFilterValueKind` | removed |
| `pivotGetGrandTotals()` | `ws.pivot(...).getGrandTotals()` |
| `pivotGetLayout()` | `ws.pivot(...).getReportLayout()` |
| `pivotGetStyleInfo()` | `ws.pivot(...).getStyleOptions()` |
| `PivotGrandTotalsResult.cols` | `PivotGrandTotals.columns` |
| `PivotGrandTotalsResult` | `PivotGrandTotals` |
| `PivotItemResult` | folded into `PivotItem` |
| `pivotLayout()` | `ws.pivot(...).getLayout()` |
| `PivotLayoutResult.rect` | `PivotLayout.range` |
| `PivotLayoutResult` | `PivotLayout` |
| `pivotName()` | `ws.pivot(...).name` |
| `pivotRemove()` | `removePivot()` |
| `PivotReportLayoutResult.layout` | `value` |
| `PivotReportLayoutResult` | removed |
| `pivotRowFieldOrder()` | `ws.pivot(...).getRowFieldOrder()` |
| `pivotSetAnchor()` | `ws.pivot(...).setAnchor()` |
| `pivotSetColFieldOrder()` | `ws.pivot(...).setColumnFieldOrder()` |
| `pivotSetGrandTotals()` | `ws.pivot(...).updateGrandTotals()` |
| `pivotSetLayout()` | `ws.pivot(...).setReportLayout()` |
| `pivotSetName()` | `ws.pivot(...).rename()` |
| `pivotSetRowFieldOrder()` | `ws.pivot(...).setRowFieldOrder()` |
| `pivotSetStyleInfo()` | `ws.pivot(...).updateStyleOptions()` |
| `PivotShowValuesAs.PercentOfCol` | `PivotShowValuesAs.PercentOfColumn` |
| `PivotShowValuesAs.PercentOfParentCol` | `PivotShowValuesAs.PercentOfParentColumn` |
| `PivotShowValuesAs.RunningTotalInCol` | `PivotShowValuesAs.RunningTotalInColumn` |
| `PivotStyleInfo.name` | `PivotStyleOptions.styleName` |
| `PivotStyleInfo` | `PivotStyleOptions` |
| `PivotStyleInfoResult` | folded into `PivotStyleOptions` |
| `PivotWorksheetSource.ref` | `PivotWorksheetSource.range` |
| `PivotWorksheetSourceResult` | folded into `PivotWorksheetSource` |
| `precedents()` | `getPrecedents()` |
| `previewChartData()` | `ws.chart(...).getPreview()` |
| `probeImage()` | `getImageInfo()` |
| `ProtectedRange.password` | `ProtectedRange.legacyPasswordHash` |
| `rangeStyleBases()` | `getRangeXfBases()` |
| `RangeStyleBasesResult` | removed |
| `RangeStyleSnapshotResult` | removed |
| `readDiagnostics()` | `getDiagnostics()` |
| `ReadDiagnosticsResult` | `ReadDiagnostics` |
| `RecalcViewport` | folded into `CellRect` |
| `RectPt.height` | `RectPt.heightPt` |
| `RectPt.width` | `RectPt.widthPt` |
| `refMode()` | `getRefMode()` |
| `remapRangeStyle()` | `remapRangeXfs()` |
| `removeAutoFilter()` | removed; use `setAutoFilter()` |
| `removeChartSeries()` | `ws.chart(...).removeSeries()` |
| `removeConditionalFormatAt()` | `removeCfRule()` |
| `removeHyperlink()` | removed; use `setHyperlink()` |
| `removeMergeAt()` | removed |
| `removeShape()` | `removeDrawing()` |
| `removeSheetColBreak()` | removed |
| `removeSheetRowBreak()` | removed |
| `removeSlicer()` | `removeDrawing()` |
| `removeTableAutoFilter()` | removed; use `ws.table(...).setAutoFilter()` |
| `removeValidationAt()` | `removeValidation()` |
| `renameSheet()` | `wb.sheet(...).rename()` |
| `ResolvedColorResult` | removed |
| `RestoreImageOptions.newId` | `RestoreDrawingOptions.assignNewId` |
| `RestoreImageOptions` | `RestoreDrawingOptions` |
| `restoreRangeStyle()` | `restoreRangeXfs()` |
| `RowLayout.height` | `RowLayout.heightPt` |
| `RowsResult` | removed |
| `saveAs()` | removed; use `save()` |
| `SaveDiagnosticsResult` | `SaveDiagnostics` |
| `SaveResult` | removed |
| `setChartAxis()` | `ws.chart(...).updateAxis()` |
| `setChartDisplay()` | `ws.chart(...).updateDisplay()` |
| `setChartFilter()` | `ws.chart(...).setFilter()` |
| `setChartLegend()` | `ws.chart(...).setLegend()` |
| `setChartPlot()` | `ws.chart(...).updatePlot()` |
| `setChartPoint()` | `ws.chart(...).setPoint()` |
| `setChartSeries()` | `ws.chart(...).updateSeries()` |
| `setChartSource()` | `ws.chart(...).setSource()` |
| `setChartTitle()` | `ws.chart(...).setTitle()` |
| `setChartType()` | `ws.chart(...).setType()` |
| `setColumnStyle()` | `setColumnXfIndex()` |
| `setDefinedNameAttrs()` | removed |
| `setDefinedNameScoped()` | removed |
| `setDrawingAnchor()` | removed; use `ws.drawing(...).update()` |
| `setDrawingZOrder()` | `ws.drawing(...).setZOrder()` |
| `setExcelProfileId()` | `setProfileId()` |
| `setIterative()` | `updateIterative()` |
| `setRowStyle()` | `setRowXfIndex()` |
| `setSheetAutoFilterXml()` | `setAutoFilterXml()` |
| `setSheetFitToPage()` | removed; use `updatePageSetup()` |
| `setSheetFormatDefaults()` | `updateFormatDefaults()` |
| `setSheetFreeze()` | removed |
| `setSheetHeaderFooter()` | `updateHeaderFooter()` |
| `setSheetHeaderFooterXml()` | `setHeaderFooterXml()` |
| `setSheetPageMargins()` | `updatePageMargins()` |
| `setSheetPageMarginsXml()` | `setPageMarginsXml()` |
| `setSheetPageSetup()` | `updatePageSetup()` |
| `setSheetPageSetupXml()` | `setPageSetupXml()` |
| `setSheetPaneTopLeft()` | removed |
| `setSheetPrintArea()` | `setPrintArea()` |
| `setSheetPrintOptions()` | `updatePrintOptions()` |
| `setSheetPrintOptionsXml()` | `setPrintOptionsXml()` |
| `setSheetPrintTitles()` | `setPrintTitles()` |
| `setSheetRightToLeft()` | removed |
| `setSheetSelection()` | `setSelection()` |
| `setSheetSheetPrXml()` | `setSheetPrXml()` |
| `setSheetShowFormulas()` | removed |
| `setSheetShowGridLines()` | removed |
| `setSheetShowOutlineSymbols()` | removed |
| `setSheetShowRowColHeaders()` | removed |
| `setSheetShowZeros()` | removed |
| `setSheetSplit()` | removed |
| `setSheetTabHidden()` | removed |
| `setSheetTabSelected()` | removed |
| `setSheetTopLeft()` | removed |
| `setSheetViewMode()` | removed |
| `setSheetVisibility()` | `setVisibility()` |
| `setSheetZoom()` | removed |
| `setSparklineGroupSparklines()` | removed; use `updateSparklineGroup()` |
| `setTableAutoFilter()` | `ws.table(...).setAutoFilter()` |
| `setTableColumnTotals()` | removed; use `ws.table(...).update()` |
| `setTableStyleDefaults()` | `updateTableStyleDefaults()` |
| `setThreadResolved()` | removed; use `updateThreadedComment()` |
| `Shape.rotation` | `ShapeProperties.rotationDeg` |
| `Shape` | `ShapeProperties` |
| `ShapeInput.rotation` | `ShapeInput.rotationDeg` |
| `ShapePaint` | removed |
| `SheetAutoFilterXmlResult.xml` | `value` |
| `SheetAutoFilterXmlResult` | removed |
| `SheetCommentEntry` | `SheetComment` |
| `SheetFormatDefaults.baseColWidth` | `SheetFormatDefaults.baseColumnWidthChars` |
| `SheetFormatDefaults.defaultColWidth` | `SheetFormatDefaults.defaultColumnWidthChars` |
| `SheetFormatDefaults.defaultRowHeight` | `SheetFormatDefaults.defaultRowHeightPt` |
| `SheetFormatDefaults.effectiveDefaultColWidth` | `SheetFormatDefaults.effectiveDefaultColumnWidthChars` |
| `SheetFormatDefaults.effectiveDefaultRowHeight` | `SheetFormatDefaults.effectiveDefaultRowHeightPt` |
| `SheetFormatDefaultsResult` | folded into `SheetFormatDefaults` |
| `sheetName()` | removed |
| `SheetPageBreaksResult` | removed |
| `SheetPageMarginsResult` | folded into `PageMargins` |
| `SheetPageSetupResult` | folded into `PageSetup` |
| `SheetPrintAreaResult` | removed |
| `SheetPrintTitlesResult` | removed |
| `SheetPrintXmlResult.xml` | `value` |
| `SheetPrintXmlResult` | removed |
| `SheetProtection.autoFilter` | `SheetProtection.allowAutoFilter`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.deleteColumns` | `SheetProtection.allowDeleteColumns`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.deleteRows` | `SheetProtection.allowDeleteRows`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.formatCells` | `SheetProtection.allowFormatCells`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.formatColumns` | `SheetProtection.allowFormatColumns`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.formatRows` | `SheetProtection.allowFormatRows`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.insertColumns` | `SheetProtection.allowInsertColumns`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.insertHyperlinks` | `SheetProtection.allowInsertHyperlinks`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.insertRows` | `SheetProtection.allowInsertRows`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.legacyPassword` | `SheetProtection.legacyPasswordHash` |
| `SheetProtection.objects` | `SheetProtection.allowEditObjects`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.pivotTables` | `SheetProtection.allowPivotTables`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.scenarios` | `SheetProtection.allowEditScenarios`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.selectLockedCells` | `SheetProtection.allowSelectLockedCells`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.selectUnlockedCells` | `SheetProtection.allowSelectUnlockedCells`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.sort` | `SheetProtection.allowSort`; inverted: `true` allows what the former flag prevented |
| `SheetProtectionResult` | removed |
| `SheetSelection.activeRow` | `SheetSelection.active` |
| `SheetSelectionResult` | folded into `SheetSelection` |
| `SheetView.paneTopLeftRow` | `SheetView.paneTopLeft` |
| `SheetView.showRowColHeaders` | `SheetView.showRowColumnHeaders` |
| `SheetView.splitX` | `SheetView.splitXTwips` |
| `SheetView.splitY` | `SheetView.splitYTwips` |
| `SheetView.topLeftRow` | `SheetView.topLeft` |
| `SheetView.zoomScale` | `SheetView.zoomPct` |
| `SheetViewResult` | removed |
| `Slicer` | `SlicerProperties` |
| `snapshotDrawing()` | `ws.drawing(...).snapshot()` |
| `snapshotRangeStyle()` | `snapshotRangeXfs()` |
| `Sparkline.formula` | `Sparkline.source` |
| `SparklineEmptyCells` | folded into `BlanksAs` |
| `SparklineGroup.dateAxisRef` | `SparklineGroup.dateAxisRange` |
| `SparklineGroup.lineWeight` | `SparklineGroup.lineWidthPt` |
| `SparklineGroupInput.dateAxisRef` | `SparklineGroupInput.dateAxisRange` |
| `SparklineGroupInput.lineWeight` | `SparklineGroupInput.lineWidthPt` |
| `spillInfo()` | `getSpill()` |
| `SpillInfo.rect` | `Spill.range` |
| `SpillInfo` | `Spill` |
| `Status.status` | `Status.code` |
| `Status` | removed |
| `statusString()` | removed |
| `StringResult` | removed |
| `tableAt()` | `ws.table(...).getProperties()` |
| `TableColumnEntry` | `TableColumn` |
| `TableColumnTotalsInput` | folded into `TableColumnPatch` |
| `tableCount()` | removed |
| `TableEntry.headerRowCount` | `TableProperties.headerRow` |
| `TableEntry.totalsRowCount` | `TableProperties.totalsRow` |
| `TableEntry` | `TableProperties` |
| `tableStyleCount()` | removed; use `listTableStyles()` |
| `TableStyleDefaultsResult.pivotStyle` | `TableStyleDefaults.pivotStyleName` |
| `TableStyleDefaultsResult.tableStyle` | `TableStyleDefaults.tableStyleName` |
| `TableStyleDefaultsResult` | `TableStyleDefaults` |
| `TableStyleElement.dxfId` | `TableStyleElement.dxfIndex` |
| `TableUpdateInput` | folded into `TablePatch` |
| `ThemeResult` | removed |
| `ThreadedComment.row` | `ThreadedComment.cell` |
| `ThreadedCommentsResult` | removed |
| `updateShape()` | `ws.shape(...).update()` |
| `updateSlicer()` | `ws.slicer(...).update()` |
| `updateTable()` | `ws.table(...).update()` |
| `ValidationEntry` | folded into `DataValidation` |
| `ValidationOutcomeResult` | `ValidationOutcome` |
| `version()` | removed; use `getVersion()` |
| `versionString()` | `getVersion()` |
| `WidthModelResult.normalFontSize` | `WidthModel.normalFontSizePt` |
| `WidthModelResult` | `WidthModel` |
| `Workbook.loadBytes()` | `Workbook.load()` |
| `WorkbookFormat.Unknown` | removed |
| `WorkbookProtection.revisionsPassword` | `WorkbookProtection.revisionsLegacyPasswordHash` |
| `WorkbookProtection.workbookPassword` | `WorkbookProtection.workbookLegacyPasswordHash` |
| `WorkbookProtectionResult` | removed |
| `xfCount()` | removed; use `listCellXfs()` |

#### Python (`formulon`)

| Before | After |
|---|---|
| `add_batch()` | removed |
| `add_chart_series()` | `ws.chart(...).add_series()` |
| `add_col_break()` | `set_column_break()` |
| `add_conditional_format()` | `add_cf_rule()` |
| `add_hyperlink()` | `set_hyperlink()` |
| `add_hyperlink_range()` | removed; use `set_hyperlink()` |
| `add_num_fmt()` | `add_number_format()` |
| `add_row_break()` | `set_row_break()` |
| `apply_table_auto_filter()` | `ws.table(...).apply_auto_filter()` |
| `AutoFilterColumn.col_id` | `AutoFilterColumn.column_offset` |
| `AutoFilterColumn.dxf_id` | `AutoFilterColumn.dxf_index` |
| `AutoFilterColumn.dyn_max_val` | `AutoFilterColumn.dynamic_max_value` |
| `AutoFilterColumn.dyn_val` | `AutoFilterColumn.dynamic_value` |
| `AutoFilterColumn.filter_val` | `AutoFilterColumn.filter_value` |
| `AutoFilterColumn.icon_id` | `AutoFilterColumn.icon_index` |
| `AutoFilterColumn.max_val_iso` | `AutoFilterColumn.max_value_iso` |
| `AutoFilterColumn.top_val` | `AutoFilterColumn.top_value` |
| `AutoFilterColumn.val_iso` | `AutoFilterColumn.value_iso` |
| `AutoFilterSortCondition.dxf_id` | `AutoFilterSortCondition.dxf_index` |
| `AutoFilterSortCondition.icon_id` | `AutoFilterSortCondition.icon_index` |
| `AutoFilterSortCondition.ref` | `AutoFilterSortCondition.range` |
| `AutoFilterSortState.ref` | `AutoFilterSortState.range` |
| `border_count()` | removed; use `list_borders()` |
| `calc_mode()` | `get_calc_mode()` |
| `calc_on_save()` | `get_calc_on_save()` |
| `cell_style_count()` | removed; use `list_cell_styles()` |
| `cell_style_xf_count()` | removed; use `list_cell_style_xfs()` |
| `cell_xf_count()` | removed; use `list_cell_xfs()` |
| `Cell` | folded into `RangeCell` |
| `CellStyleRecord.i_level` | `CellStyle.outline_level` |
| `CellStyleRecord.xf_id` | `CellStyle.style_xf_index` |
| `CellStyleRecord` | `CellStyle` |
| `CellXf.horizontal_align` | `CellXf.horizontal` |
| `CellXf.vertical_align` | `CellXf.vertical` |
| `CellXf.xf_id` | `CellXf.style_xf_index` |
| `cf_count()` | removed; use `list_cf_rules()` |
| `CfCellResult` | `CfCell` |
| `CfColor` | folded into `Color` |
| `CfIcon.icon_id` | `CfIcon.icon_index` |
| `CfMatch.dxf_id` | `CfMatch.dxf_index` |
| `ChartAxisInfo` | `ChartAxis` |
| `ChartBlanksAs` | `BlanksAs` |
| `ChartCategories.format_code` | `ChartCategories.number_format` |
| `ChartDataSeries.format_code` | `ChartDataSeries.number_format` |
| `ChartDataSeries.idx` | `ChartDataSeries.series_id` |
| `ChartDataSeries.x_format_code` | `ChartDataSeries.x_number_format` |
| `ChartFillInfo.transparency` | `ChartFill.transparency_pct` |
| `ChartFillInfo` | `ChartFill` |
| `ChartFillPatch.transparency` | `ChartFillPatch.transparency_pct` |
| `ChartInfo.is_chartex` | `ChartProperties.is_extended` |
| `ChartInfo` | `ChartProperties` |
| `ChartInsertKind` | folded into `ChartPlotType` |
| `ChartLabelsInfo` | `ChartLabels` |
| `ChartLegendInfo.hidden_series` | `ChartLegend.hidden_entries` |
| `ChartLegendInfo` | `ChartLegend` |
| `ChartLineInfo.transparency` | `ChartLine.transparency_pct` |
| `ChartLineInfo` | `ChartLine` |
| `ChartLinePatch.transparency` | `ChartLinePatch.transparency_pct` |
| `ChartMarkerInfo.size` | `ChartMarker.size_pt` |
| `ChartMarkerInfo` | `ChartMarker` |
| `ChartMarkerPatch.size` | `ChartMarkerPatch.size_pt` |
| `ChartPlotInfo.chartex_layout` | `ChartPlot.extended_layout` |
| `ChartPlotInfo.first_slice_angle` | `ChartPlot.first_slice_angle_deg` |
| `ChartPlotInfo.gap_width` | `ChartPlot.gap_width_pct` |
| `ChartPlotInfo.hole_size` | `ChartPlot.hole_size_pct` |
| `ChartPlotInfo.overlap` | `ChartPlot.overlap_pct` |
| `ChartPlotInfo` | `ChartPlot` |
| `ChartPlotPatch.first_slice_angle` | `ChartPlotPatch.first_slice_angle_deg` |
| `ChartPlotPatch.gap_width` | `ChartPlotPatch.gap_width_pct` |
| `ChartPlotPatch.hole_size` | `ChartPlotPatch.hole_size_pct` |
| `ChartPlotPatch.overlap` | `ChartPlotPatch.overlap_pct` |
| `ChartPlotType.CHARTEX` | `ChartPlotType.EXTENDED` |
| `ChartPointInfo` | `ChartPoint` |
| `ChartPreviewOverrides.source` | `range` |
| `ChartPreviewOverrides` | removed |
| `ChartSeriesInfo.idx` | `ChartSeries.series_id` |
| `ChartSeriesInfo` | `ChartSeries` |
| `ChartSheetRange` | folded into `SheetRect` |
| `ChartSourceInfo` | `ChartSource` |
| `ChartTitleInfo` | `ChartTitle` |
| `ChartUnresolved.series_idx` | `ChartUnresolved.series_id` |
| `ChartUnresolved.source` | `ChartUnresolved.role` |
| `CivilTime` | removed |
| `clear_auto_filter()` | `clear_auto_filter_criteria()` |
| `clear_cell_xf_index()` | `reset_cell_xf_index()` |
| `clear_column_style()` | `reset_column_xf_index()` |
| `clear_column_width()` | `reset_column_width()` |
| `clear_conditional_formats()` | `clear_cf_rules()` |
| `clear_pinned_now()` | removed; use `set_pinned_now()` |
| `clear_row_height()` | `reset_row_height()` |
| `clear_row_style()` | `reset_row_xf_index()` |
| `clear_table_auto_filter()` | `ws.table(...).clear_auto_filter_criteria()` |
| `ColorScale` | removed; use CfColorScaleInput to write and CfColorScale to read |
| `ColorSpec` | `Color` |
| `column_chars_to_pt()` | `convert_column_chars_to_pt()` |
| `column_pt_to_chars()` | `convert_column_pt_to_chars()` |
| `ColumnLayout.width` | `ColumnLayout.width_chars` |
| `comment_count()` | removed; use `list_comments()` |
| `CommentEntry` | `SheetComment` |
| `ConditionalFormat.dxf_id` | `ConditionalFormat.dxf_index` |
| `ConditionalFormatInput.dxf_id` | `ConditionalFormatInput.dxf_index` |
| `DataBar` | removed; use CfDataBarInput to write and CfDataBar to read |
| `DataTable.col_input` | `DataTable.column_input` |
| `DataTable.rect` | `DataTable.range` |
| `DataTableInput.rect` | `DataTableInput.range` |
| `DataValidation.show_dropdown` | `DataValidation.show_in_cell_dropdown` |
| `DataValidationInput.show_dropdown` | `DataValidationInput.show_in_cell_dropdown` |
| `date1904()` | `get_date1904()` |
| `DefinedName.local_sheet_id` | `DefinedName.scope` |
| `delete_cols()` | `delete_columns()` |
| `dependents()` | `get_dependents()` |
| `DifferentialFormat.alignment_xml` | `DifferentialFormat.alignment` |
| `DifferentialFormat.num_fmt_code` | `DifferentialFormat.number_format` |
| `DifferentialFormat.protection_xml` | `DifferentialFormat.protection` |
| `DifferentialFormat` | folded into `Dxf` |
| `DrawingObject` | `DrawingProperties` |
| `dxf_count()` | removed; use `list_dxfs()` |
| `edit_threaded_comment()` | `update_threaded_comment()` |
| `EffectiveStyle.font` | `EffectiveStyle.font_color` |
| `EffectiveStyle.num_fmt_code` | `EffectiveStyle.number_format` |
| `error_display_name()` | removed |
| `ErrorName` | `LocaleErrorName` |
| `eval_formula()` | `evaluate_formula()` |
| `evaluate_table_auto_filter()` | `ws.table(...).evaluate_auto_filter()` |
| `excel_profile_id()` | `get_profile_id()` |
| `ExcelProfileId` | `ProfileId` |
| `external_link_count()` | removed; use `list_external_links()` |
| `ExternalLinkRecord.index` | `ExternalLink.book_number` |
| `ExternalLinkRecord` | `ExternalLink` |
| `fill_count()` | removed; use `list_fills()` |
| `FillRecord` | `Fill` |
| `FilterOperator` | `ComparisonOperator` |
| `font_count()` | removed; use `list_fonts()` |
| `FontRecord.size` | `Font.size_pt` |
| `FontRecord` | `Font` |
| `full_precision()` | `get_full_precision()` |
| `get_chart()` | `ws.chart(...).get_properties()` |
| `get_chart_data()` | `ws.chart(...).get_data()` |
| `get_col_breaks()` | `list_column_breaks()` |
| `get_comments()` | `list_comments()` |
| `get_conditional_format_at()` | removed |
| `get_conditional_formats()` | removed |
| `get_external_link_at()` | removed; use `list_external_links()` |
| `get_external_links()` | removed |
| `get_hyperlinks()` | `list_hyperlinks()` |
| `get_image()` | `ws.picture(...).get_properties()` |
| `get_merges()` | `list_merges()` |
| `get_num_fmt()` | `get_number_format()` |
| `get_persons()` | `list_people()` |
| `get_phonetic()` | `get_cell_phonetic()` |
| `get_phonetic_properties()` | `get_cell_phonetic_properties()` |
| `get_phonetic_runs()` | `get_cell_phonetic_runs()` |
| `get_pivot_cache_worksheet_source()` | `wb.pivot_cache(...).get_worksheet_source()` |
| `get_pivot_report_layout()` | `ws.pivot(...).get_report_layout()` |
| `get_protected_ranges()` | removed |
| `get_row_breaks()` | `list_row_breaks()` |
| `get_sheet_columns()` | `list_column_overrides()` |
| `get_sheet_format_defaults()` | `get_format_defaults()` |
| `get_sheet_row_overrides()` | `list_row_overrides()` |
| `get_sheet_selection()` | `get_selection()` |
| `get_sheet_view()` | `get_view()` |
| `get_show_phonetic()` | `get_cell_show_phonetic()` |
| `get_table_auto_filter()` | `ws.table(...).get_auto_filter()` |
| `get_theme()` | `get_theme_fonts()` |
| `get_threaded_comments()` | `list_threaded_comments()` |
| `get_validation_at()` | removed |
| `get_validations()` | removed |
| `GradientFill.degree` | `GradientFill.angle_deg` |
| `has_phonetic_properties()` | removed; use `get_cell_phonetic_properties()` |
| `Hyperlink.rect` | `Hyperlink.range` |
| `hyperlink_count()` | removed; use `list_hyperlinks()` |
| `IconSet` | removed; use CfIconSetInput to write and CfIconSet to read |
| `ImageInfo.px_height` | `ImageInfo.height_px` |
| `ImageInfo.px_width` | `ImageInfo.width_px` |
| `insert_chart()` | `add_chart()` |
| `insert_cols()` | `insert_columns()` |
| `insert_image()` | `add_picture()` |
| `insert_shape()` | `add_shape()` |
| `is_valid()` | `closed()` |
| `iter_cells()` | removed |
| `iter_defined_names()` | `list_defined_names()` |
| `iter_passthrough()` | `list_passthroughs()` |
| `iter_tables()` | removed |
| `lambda_text_at()` | `get_lambda_text()` |
| `library_version()` | `get_version()` |
| `list_drawing_objects()` | `list_drawings()` |
| `list_invalid_cells()` | `get_invalid_cells()` |
| `MarginsPt.bottom` | `PaginationMargins.bottom_in` |
| `MarginsPt.footer` | `PaginationMargins.footer_in` |
| `MarginsPt.header` | `PaginationMargins.header_in` |
| `MarginsPt.left` | `PaginationMargins.left_in` |
| `MarginsPt.right` | `PaginationMargins.right_in` |
| `MarginsPt.top` | `PaginationMargins.top_in` |
| `MarginsPt` | `PaginationMargins` |
| `memory_usage()` | `get_memory_usage()` |
| `Mention.length` | `Mention.end` |
| `merge_count()` | removed; use `list_merges()` |
| `MergeRange` | folded into `CellRect` |
| `move_sheet()` | `wb.sheet(...).move()` |
| `PageBreak` | `RowBreak` |
| `PageLayout.rect` | `PaginationPage.range` |
| `PageLayout` | `PaginationPage` |
| `PageMargins.bottom_stated` | `PageMargins.bottom_in_stated` |
| `PageMargins.bottom` | `PageMargins.bottom_in` |
| `PageMargins.footer_stated` | `PageMargins.footer_in_stated` |
| `PageMargins.footer` | `PageMargins.footer_in` |
| `PageMargins.header_stated` | `PageMargins.header_in_stated` |
| `PageMargins.header` | `PageMargins.header_in` |
| `PageMargins.left_stated` | `PageMargins.left_in_stated` |
| `PageMargins.left` | `PageMargins.left_in` |
| `PageMargins.right_stated` | `PageMargins.right_in_stated` |
| `PageMargins.right` | `PageMargins.right_in` |
| `PageMargins.top_stated` | `PageMargins.top_in_stated` |
| `PageMargins.top` | `PageMargins.top_in` |
| `PageSetup.scale_stated` | `PageSetup.scale_pct_stated` |
| `PageSetup.scale` | `PageSetup.scale_pct` |
| `paginate()` | `get_pagination()` |
| `PaginationResult.horizontal_breaks` | `Pagination.row_breaks` |
| `PaginationResult.scale` | `Pagination.scale_pct` |
| `PaginationResult.vertical_breaks` | `Pagination.column_breaks` |
| `PaginationResult` | `Pagination` |
| `PaperInfo` | `PaginationPaper` |
| `partial_recalc()` | `recalc_viewport()` |
| `PassthroughPart` | removed |
| `PhoneticProperties.font_id` | `PhoneticProperties.font_index` |
| `pinned_now()` | `get_pinned_now()` |
| `pivot_cache_count()` | removed |
| `pivot_cache_create()` | `add_pivot_cache()` |
| `pivot_cache_field_add()` | `wb.pivot_cache(...).add_field()` |
| `pivot_cache_field_add_shared_item_blank()` | removed |
| `pivot_cache_field_add_shared_item_bool()` | removed |
| `pivot_cache_field_add_shared_item_error()` | removed |
| `pivot_cache_field_add_shared_item_number()` | removed |
| `pivot_cache_field_add_shared_item_text()` | removed |
| `pivot_cache_field_clear()` | `wb.pivot_cache(...).clear_fields()` |
| `pivot_cache_field_clear_shared_items()` | `wb.pivot_cache(...).clear_field_shared_items()` |
| `pivot_cache_field_count()` | removed |
| `pivot_cache_field_name()` | removed |
| `pivot_cache_field_shared_item_count()` | removed |
| `pivot_cache_id()` | `ws.pivot(...).cache()` |
| `pivot_cache_id_at()` | removed |
| `pivot_cache_record_add()` | `wb.pivot_cache(...).add_record()` |
| `pivot_cache_record_clear()` | `wb.pivot_cache(...).clear_records()` |
| `pivot_cache_record_count()` | `wb.pivot_cache(...).get_record_count()` |
| `pivot_cache_record_set_blank()` | removed |
| `pivot_cache_record_set_bool()` | removed |
| `pivot_cache_record_set_error()` | removed |
| `pivot_cache_record_set_number()` | removed |
| `pivot_cache_record_set_text()` | removed |
| `pivot_cache_remove()` | `remove_pivot_cache()` |
| `pivot_col_field_order()` | `ws.pivot(...).get_column_field_order()` |
| `pivot_count()` | removed |
| `pivot_create()` | `add_pivot()` |
| `pivot_data_field_add()` | `ws.pivot(...).add_data_field()` |
| `pivot_data_field_at()` | `ws.pivot(...).get_data_field()` |
| `pivot_data_field_clear()` | `ws.pivot(...).clear_data_fields()` |
| `pivot_data_field_count()` | removed; use `ws.pivot(...).list_data_fields()` |
| `pivot_data_field_set()` | `ws.pivot(...).set_data_field()` |
| `pivot_field_add()` | removed |
| `pivot_field_add_item()` | `ws.pivot(...).add_field_item()` |
| `pivot_field_add_item_at()` | removed |
| `pivot_field_add_subtotal_fn()` | removed; use `ws.pivot(...).set_field_subtotal_functions()` |
| `pivot_field_at()` | `ws.pivot(...).get_field()` |
| `pivot_field_clear()` | `ws.pivot(...).clear_fields()` |
| `pivot_field_clear_date_group()` | removed; use `ws.pivot(...).set_field_date_group()` |
| `pivot_field_clear_items()` | `ws.pivot(...).clear_field_items()` |
| `pivot_field_clear_subtotal_fns()` | removed; use `ws.pivot(...).set_field_subtotal_functions()` |
| `pivot_field_count()` | removed; use `ws.pivot(...).list_fields()` |
| `pivot_field_item_at()` | `ws.pivot(...).get_field_item()` |
| `pivot_field_item_count()` | removed; use `ws.pivot(...).list_field_items()` |
| `pivot_field_set_axis()` | `ws.pivot(...).set_field_axis()` |
| `pivot_field_set_date_group()` | `ws.pivot(...).set_field_date_group()` |
| `pivot_field_set_item_visible()` | removed |
| `pivot_field_set_number_format()` | removed |
| `pivot_field_set_sort()` | removed |
| `pivot_field_set_subtotal_top()` | removed |
| `pivot_filter_add()` | `ws.pivot(...).add_filter()` |
| `pivot_filter_at()` | `ws.pivot(...).get_filter()` |
| `pivot_filter_clear()` | `ws.pivot(...).clear_filters()` |
| `pivot_filter_count()` | removed; use `ws.pivot(...).list_filters()` |
| `pivot_filter_remove_at()` | `ws.pivot(...).remove_filter()` |
| `pivot_grand_totals()` | `ws.pivot(...).get_grand_totals()` |
| `pivot_layout()` | `ws.pivot(...).get_layout()` |
| `pivot_name()` | `ws.pivot(...).name` |
| `pivot_remove()` | `remove_pivot()` |
| `pivot_row_field_order()` | `ws.pivot(...).get_row_field_order()` |
| `pivot_set_anchor()` | `ws.pivot(...).set_anchor()` |
| `pivot_set_col_field_order()` | `ws.pivot(...).set_column_field_order()` |
| `pivot_set_grand_totals()` | `ws.pivot(...).update_grand_totals()` |
| `pivot_set_name()` | `ws.pivot(...).rename()` |
| `pivot_set_row_field_order()` | `ws.pivot(...).set_row_field_order()` |
| `pivot_set_style_info()` | `ws.pivot(...).update_style_options()` |
| `pivot_style_info()` | `ws.pivot(...).get_style_options()` |
| `PivotAxis.COL` | `PivotAxis.COLUMN` |
| `PivotCellKind.COL_LABEL` | `PivotCellKind.COLUMN_LABEL` |
| `PivotCellKind.COL_SUBTOTAL` | `PivotCellKind.COLUMN_SUBTOTAL` |
| `PivotDataFieldSpec` | `PivotDataFieldInput` |
| `PivotFieldInfo.subtotal_fns` | `PivotField.subtotal_functions` |
| `PivotFieldInfo` | `PivotField` |
| `PivotFilterSpec.type` | `PivotFilterInput.kind` |
| `PivotFilterSpec` | `PivotFilterInput` |
| `PivotFilterType.LABEL_DATE` | `PivotFilterKind.DATE_BETWEEN` |
| `PivotFilterType.VALUE_TOP_10` | `PivotFilterKind.VALUE_TOP_COUNT` |
| `PivotFilterType` | `PivotFilterKind` |
| `PivotFilterValueKind` | removed |
| `PivotItemInfo` | `PivotItem` |
| `PivotLayout.rect` | `PivotLayout.range` |
| `PivotShowValuesAs.PERCENT_OF_COL` | `PivotShowValuesAs.PERCENT_OF_COLUMN` |
| `PivotShowValuesAs.PERCENT_OF_PARENT_COL` | `PivotShowValuesAs.PERCENT_OF_PARENT_COLUMN` |
| `PivotShowValuesAs.RUNNING_TOTAL_IN_COL` | `PivotShowValuesAs.RUNNING_TOTAL_IN_COLUMN` |
| `PivotStyleInfo.name` | `PivotStyleOptions.style_name` |
| `PivotStyleInfo` | `PivotStyleOptions` |
| `PivotWorksheetSource.ref` | `PivotWorksheetSource.range` |
| `precedents()` | `get_precedents()` |
| `preview_chart_data()` | `ws.chart(...).get_preview()` |
| `probe_image()` | `get_image_info()` |
| `protected_range_count()` | removed; use `list_protected_ranges()` |
| `ProtectedRange.password` | `ProtectedRange.legacy_password_hash` |
| `range_style_bases()` | `get_range_xf_bases()` |
| `read_diagnostics()` | `get_diagnostics()` |
| `RectPt.height` | `RectPt.height_pt` |
| `RectPt.width` | `RectPt.width_pt` |
| `ref_mode()` | `get_ref_mode()` |
| `remap_range_style()` | `remap_range_xfs()` |
| `remove_auto_filter()` | removed; use `set_auto_filter()` |
| `remove_chart_series()` | `ws.chart(...).remove_series()` |
| `remove_col_break()` | removed |
| `remove_conditional_format_at()` | `remove_cf_rule()` |
| `remove_hyperlink()` | removed; use `set_hyperlink()` |
| `remove_merge_at()` | removed |
| `remove_row_break()` | removed |
| `remove_shape()` | `remove_drawing()` |
| `remove_slicer()` | `remove_drawing()` |
| `remove_table_auto_filter()` | removed; use `ws.table(...).set_auto_filter()` |
| `remove_validation_at()` | `remove_validation()` |
| `rename_sheet()` | `wb.sheet(...).rename()` |
| `restore_range_style()` | `restore_range_xfs()` |
| `RowLayout.height` | `RowLayout.height_pt` |
| `save_as()` | removed; use `save()` |
| `set_chart_axis()` | `ws.chart(...).update_axis()` |
| `set_chart_display()` | `ws.chart(...).update_display()` |
| `set_chart_filter()` | `ws.chart(...).set_filter()` |
| `set_chart_legend()` | `ws.chart(...).set_legend()` |
| `set_chart_plot()` | `ws.chart(...).update_plot()` |
| `set_chart_point()` | `ws.chart(...).set_point()` |
| `set_chart_series()` | `ws.chart(...).update_series()` |
| `set_chart_source()` | `ws.chart(...).set_source()` |
| `set_chart_title()` | `ws.chart(...).set_title()` |
| `set_chart_type()` | `ws.chart(...).set_type()` |
| `set_column_style()` | `set_column_xf_index()` |
| `set_defined_name_attrs()` | removed |
| `set_defined_name_scoped()` | removed |
| `set_drawing_anchor()` | removed; use `ws.drawing(...).update()` |
| `set_drawing_z_order()` | `ws.drawing(...).set_z_order()` |
| `set_excel_profile_id()` | `set_profile_id()` |
| `set_fit_to_page()` | removed; use `update_page_setup()` |
| `set_header_footer()` | `update_header_footer()` |
| `set_iterative()` | `update_iterative()` |
| `set_iterative_enabled()` | removed |
| `set_page_margins()` | `update_page_margins()` |
| `set_page_setup()` | `update_page_setup()` |
| `set_phonetic()` | `set_cell_phonetic()` |
| `set_phonetic_properties()` | `set_cell_phonetic_properties()` |
| `set_phonetic_runs()` | `set_cell_phonetic_runs()` |
| `set_pivot_cache_worksheet_source()` | `wb.pivot_cache(...).set_worksheet_source()` |
| `set_pivot_report_layout()` | `ws.pivot(...).set_report_layout()` |
| `set_print_options()` | `update_print_options()` |
| `set_row_style()` | `set_row_xf_index()` |
| `set_sheet_format_defaults()` | `update_format_defaults()` |
| `set_sheet_freeze()` | removed |
| `set_sheet_pane_top_left()` | removed |
| `set_sheet_right_to_left()` | removed |
| `set_sheet_selection()` | `set_selection()` |
| `set_sheet_show_formulas()` | removed |
| `set_sheet_show_grid_lines()` | removed |
| `set_sheet_show_outline_symbols()` | removed |
| `set_sheet_show_row_col_headers()` | removed |
| `set_sheet_show_zeros()` | removed |
| `set_sheet_split()` | removed |
| `set_sheet_tab_hidden()` | removed |
| `set_sheet_tab_selected()` | removed |
| `set_sheet_top_left()` | removed |
| `set_sheet_view_mode()` | removed |
| `set_sheet_visibility()` | `set_visibility()` |
| `set_sheet_zoom()` | removed |
| `set_show_phonetic()` | `set_cell_show_phonetic()` |
| `set_sparkline_group_sparklines()` | removed; use `update_sparkline_group()` |
| `set_table_auto_filter()` | `ws.table(...).set_auto_filter()` |
| `set_table_style_defaults()` | `update_table_style_defaults()` |
| `set_thread_resolved()` | removed; use `update_threaded_comment()` |
| `Shape.rotation` | `ShapeProperties.rotation_deg` |
| `Shape` | `ShapeProperties` |
| `ShapePaint` | removed |
| `sheet_count()` | removed |
| `sheet_name()` | removed |
| `SheetFormatDefaults.base_col_width` | `SheetFormatDefaults.base_column_width_chars` |
| `SheetFormatDefaults.default_col_width` | `SheetFormatDefaults.default_column_width_chars` |
| `SheetFormatDefaults.default_row_height` | `SheetFormatDefaults.default_row_height_pt` |
| `SheetFormatDefaults.effective_default_col_width` | `SheetFormatDefaults.effective_default_column_width_chars` |
| `SheetFormatDefaults.effective_default_row_height` | `SheetFormatDefaults.effective_default_row_height_pt` |
| `SheetProtection.auto_filter` | `SheetProtection.allow_auto_filter`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.delete_columns` | `SheetProtection.allow_delete_columns`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.delete_rows` | `SheetProtection.allow_delete_rows`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.format_cells` | `SheetProtection.allow_format_cells`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.format_columns` | `SheetProtection.allow_format_columns`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.format_rows` | `SheetProtection.allow_format_rows`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.insert_columns` | `SheetProtection.allow_insert_columns`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.insert_hyperlinks` | `SheetProtection.allow_insert_hyperlinks`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.insert_rows` | `SheetProtection.allow_insert_rows`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.legacy_password` | `SheetProtection.legacy_password_hash` |
| `SheetProtection.objects` | `SheetProtection.allow_edit_objects`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.pivot_tables` | `SheetProtection.allow_pivot_tables`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.scenarios` | `SheetProtection.allow_edit_scenarios`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.select_locked_cells` | `SheetProtection.allow_select_locked_cells`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.select_unlocked_cells` | `SheetProtection.allow_select_unlocked_cells`; inverted: `true` allows what the former flag prevented |
| `SheetProtection.sort` | `SheetProtection.allow_sort`; inverted: `true` allows what the former flag prevented |
| `SheetSelection.active_row` | `SheetSelection.active` |
| `SheetView.pane_top_left_row` | `SheetView.pane_top_left` |
| `SheetView.show_row_col_headers` | `SheetView.show_row_column_headers` |
| `SheetView.split_x` | `SheetView.split_x_twips` |
| `SheetView.split_y` | `SheetView.split_y_twips` |
| `SheetView.top_left_row` | `SheetView.top_left` |
| `SheetView.zoom_scale` | `SheetView.zoom_pct` |
| `Slicer` | `SlicerProperties` |
| `snapshot_drawing()` | `ws.drawing(...).snapshot()` |
| `snapshot_range_style()` | `snapshot_range_xfs()` |
| `Sparkline.formula` | `Sparkline.source` |
| `SparklineEmptyCells` | folded into `BlanksAs` |
| `SparklineGroup.date_axis_ref` | `SparklineGroup.date_axis_range` |
| `SparklineGroup.line_weight` | `SparklineGroup.line_width_pt` |
| `spill_info()` | `get_spill()` |
| `SpillInfo.rect` | `Spill.range` |
| `SpillInfo` | `Spill` |
| `StyleBatchIndices` | removed |
| `Table.header_row_count` | `TableProperties.header_row` |
| `Table.totals_row_count` | `TableProperties.totals_row` |
| `table_create()` | `add_table()` |
| `table_remove()` | `remove_table()` |
| `table_rename()` | `ws.table(...).rename()` |
| `table_set_column_totals()` | removed; use `ws.table(...).update()` |
| `table_set_style_info()` | `ws.table(...).update_style_options()` |
| `table_style_count()` | removed; use `list_table_styles()` |
| `table_style_defaults()` | `get_table_style_defaults()` |
| `table_update()` | `ws.table(...).update()` |
| `Table` | `TableProperties` |
| `TableColumnTotalsInput` | folded into `TableColumnPatch` |
| `TableStyleElement.dxf_id` | `TableStyleElement.dxf_index` |
| `Theme` | removed |
| `ThreadedComment.row` | `ThreadedComment.cell` |
| `update_shape()` | `ws.shape(...).update()` |
| `update_slicer()` | `ws.slicer(...).update()` |
| `validation_count()` | removed; use `list_validations()` |
| `version_string()` | removed; use `get_version()` |
| `WidthModel.normal_font_size` | `WidthModel.normal_font_size_pt` |
| `Workbook.canonicalize_formula()` | `canonicalize_formula()` |
| `Workbook.canonicalize_function_name()` | `canonicalize_function_name()` |
| `Workbook.collation_classes()` | `get_collation_classes()` |
| `Workbook.compare_text()` | `compare_text()` |
| `Workbook.function_count()` | removed; use `list_function_names()` |
| `Workbook.function_metadata()` | `get_function_metadata()` |
| `Workbook.function_name_at()` | removed; use `list_function_names()` |
| `Workbook.locale_facts()` | `get_locale_facts()` |
| `Workbook.localize_formula()` | `localize_formula()` |
| `Workbook.localize_function_name()` | `localize_function_name()` |
| `WorkbookFormat.UNKNOWN` | removed |
| `WorkbookProtection.revisions_password` | `WorkbookProtection.revisions_legacy_password_hash` |
| `WorkbookProtection.workbook_password` | `WorkbookProtection.workbook_legacy_password_hash` |
<!-- migration-table:end -->

### C ABI changes

A caller compiled against an earlier header must be recompiled.

**Functions.**

| Before | After |
|---|---|
| `fm_workbook_set_defined_name(wb, name, formula)`, `fm_workbook_set_defined_name_scoped`, `fm_workbook_set_defined_name_attrs` | `fm_workbook_set_defined_name(wb, const fm_defined_name_t*)` (name, formula, `local_sheet_id`, comment, hidden); an empty formula is `FM_STATUS_INVALID_ARGUMENT` instead of a removal |
| `fm_workbook_defined_name_at` (two strings and a scope), `fm_workbook_defined_name_attrs_at` | `fm_workbook_get_defined_name_at(wb, index, fm_defined_name_t*)` |
| `fm_workbook_set_iterative(wb, enabled, max_iterations, max_change)`, `fm_workbook_set_iterative_enabled` | `fm_workbook_update_iterative(wb, const fm_iterative_settings_patch_t*)`; a negative field keeps the stored value |
| `fm_sheet_set_print_titles(wb, sheet_index, repeat_rows, repeat_cols)`, `fm_sheet_get_print_titles` (A1 text) | take and fill an `fm_print_titles_t` (per-axis `has_*` flag and 0-based span); `NULL` clears |
| `fm_sheet_set_view_mode(wb, sheet_index, const char* view)` | removed; `fm_sheet_update_view` sets `view_mode`, an `int32_t` `fm_sheet_view_mode_t` |
| `fm_sheet_set_tab_hidden` | removed; `fm_sheet_set_visibility` states every tab state |
| `fm_sheet_set_fit_to_page` | removed; `fm_sheet_update_page_setup` carries the fit-to-page flag |
| `fm_sheet_set_sparkline_group_sparklines(wb, sheet_index, idx, items, count)`, `fm_sheet_update_sparkline_group(wb, sheet_index, idx, const fm_sparkline_group*)` | `fm_sheet_set_sparkline_group_sparklines` is removed; `fm_sheet_update_sparkline_group(wb, sheet_index, group_id, const fm_sparkline_group_patch_t*)` merges the group's fields and, when `items_op` is `SET`, replaces its sparklines; an empty `SET` `items` is `FM_STATUS_INVALID_ARGUMENT`. `fm_sheet_add_sparkline_group` returns the group's id, `fm_sheet_get_sparkline_group` reads one by id and `fm_sparkline_group_t.id` carries it |
| `fm_sheet_set_thread_resolved` and `fm_sheet_edit_threaded_comment(wb, sheet_index, id, text, mentions, mention_count)` | both are removed; `fm_sheet_update_threaded_comment(wb, sheet_index, id, const fm_threaded_comment_patch_t*)` with `text`, `mentions` (applied when `mentions_engaged` is non-zero) and `done` (`-1` keeps the resolved state). `done` on a reply id returns `FM_STATUS_INVALID_ARGUMENT` |
| `fm_workbook_set_cell_phonetic_runs`, `fm_sheet_add_threaded_comment` | `fm_sheet_set_cell_phonetic_runs`, `fm_sheet_add_threaded_comment` and `fm_sheet_update_threaded_comment` take an `int32_t offset_unit` (`fm_text_offset_unit_t`: `FM_TEXT_OFFSET_UNIT_UTF16` 0, `FM_TEXT_OFFSET_UNIT_CODE_POINT` 1) before the record, which picks the offset pair each `fm_phonetic_run_t` or `fm_mention_t` is read from. Both records carry `start` / `end` in UTF-16 units and `start_cp` / `end_cp` in code points, and reads fill both. `fm_parse_error_info_t` adds `start_cp`, `end_cp` and `column_cp` |
| (none) | `fm_sheet_get_print_options` and `fm_sheet_get_header_footer` read the print options and header/footer back |
| (none) | `fm_sheet_get_table_runtime_id` / `fm_workbook_resolve_table_runtime_id` and `fm_sheet_get_pivot_runtime_id` / `fm_sheet_resolve_pivot_runtime_id` map a table or pivot index to an id that survives other tables or pivots being added or removed, and back |
| (none) | `fm_sheet_get_pivot_spill` reads the range a blocked pivot report needs into an `fm_spill_info_t`, and `fm_sheet_get_spill` reports the same at the report's anchor |
| a pivot `number_format` (`fm_pivot_field_spec_t`, the data-field spec and the pivot field setters) given as the text of a number-format id | `number_format` is a format code (`0.00%`). A digits-only string is read as a format code, so a call written for the old form silently selects a different format |
| `fm_sheet_set_page_setup`, `_set_page_margins`, `_set_print_options`, `_set_header_footer`, `_set_format_defaults`, `fm_styles_set_table_style_defaults` | the same names with `update` (`fm_sheet_update_page_setup` and so on) and `fm_workbook_update_table_style_defaults`; a field of the patch record that is unset keeps the stored value |
| `fm_workbook_pivot_set_grand_totals`, `fm_workbook_pivot_set_style_info`, `fm_sheet_set_chart_axis`, `_set_chart_display`, `_set_chart_series`, `_set_chart_plot` | `fm_sheet_update_pivot_grand_totals`, `fm_sheet_update_pivot_style_options`, `fm_sheet_update_chart_axis`, `_update_chart_display`, `_update_chart_series`, `_update_chart_plot` |
| (none) | `fm_sheet_update_view`, `fm_sheet_update_chart_legend` and `fm_sheet_update_chart_point` merge a patch; `fm_sheet_set_chart_legend(..., const fm_chart_legend_t* or NULL)` replaces or removes the legend |
| `fm_sheet_remove_auto_filter`, `fm_sheet_remove_row_break`, `fm_sheet_remove_col_break`, `fm_workbook_clear_pinned_now` | removed; `fm_sheet_set_auto_filter(wb, sheet_index, NULL)`, `fm_sheet_set_row_break(..., NULL)`, `fm_sheet_set_column_break(..., NULL)` and `fm_workbook_set_pinned_now(wb, NULL)` remove, and removing what is not there is a no-op |
| `fm_sheet_add_hyperlink` | `fm_sheet_set_hyperlink`; it replaces the link of the cell and a `NULL` link removes it |
| `fm_sheet_clear_row_height`, `fm_sheet_clear_column_width`, `fm_sheet_clear_column_style`, `fm_sheet_clear_row_style` | `fm_sheet_reset_row_height`, `fm_sheet_reset_column_width`, `fm_sheet_reset_column_xf_index`, `fm_sheet_reset_row_xf_index` |
| `fm_sheet_clear_auto_filter`, `fm_table_clear_auto_filter` | `fm_sheet_clear_auto_filter_criteria`, `fm_sheet_clear_table_auto_filter_criteria` |
| the `fm_workbook_*table*` functions | `fm_sheet_*table*` with a sheet-local `table_index`; `fm_workbook_find_table` is unchanged |
| `fm_workbook_pivot_set_anchor` taking a rectangle | `fm_sheet_set_pivot_anchor` takes one cell, as `fm_pivot_spec_t.anchor` does |
| `fm_workbook_evaluate_formula` (top-left cell only) | removed; `fm_sheet_evaluate_formula_array` returns the array and a caller takes element (0, 0) |
| `fm_pagination_horizontal_break_at(p, i, uint32_t* out_row)`, `fm_pagination_vertical_break_at(p, i, uint32_t* out_col)`, `fm_pagination_horizontal_break_is_manual`, `fm_pagination_vertical_break_is_manual` | `fm_pagination_get_row_break_at` / `fm_pagination_get_column_break_at` fill an `fm_row_break_t` / `fm_column_break_t` (span and `manual` included); the two `*_is_manual` functions are removed |
| `fm_workbook_table_create(wb, sheet_index, ref, name, display_name, column_names, column_count, style_name, header_row, totals_row, out_index)` | `fm_sheet_add_table(wb, sheet_index, const fm_table_spec_t*, out_table_index)`; the spec also carries the four style toggles, and the call is atomic |
| `fm_workbook_pivot_create(wb, sheet_index, utf8_name, cache_id, anchor_row, anchor_col, out_pivot_index)` | `fm_sheet_add_pivot(wb, sheet_index, const fm_pivot_spec_t*, out_pivot_index)`; an unknown `cache_id` is `FM_STATUS_NOT_FOUND` |
| `fm_workbook_pivot_set_style_info(wb, sheet_index, pivot_index, utf8_style_name, show_row_headers, show_col_headers, show_row_stripes, show_col_stripes, show_last_column)` | `fm_sheet_update_pivot_style_options(wb, sheet_index, pivot_index, const fm_pivot_style_options_patch_t*)`; an all-`KEEP` patch changes nothing |
| `fm_workbook_pivot_field_set_date_group(..., granularity, calendar, start_year_or_neg1, end_year_or_neg1, interval_days, start_serial_or_neg1, end_serial_or_neg1)` | `fm_sheet_set_pivot_field_date_group(..., const fm_pivot_date_group_spec_t*)`; `has_start` / `has_end` replace the `-1` sentinels |
| `fm_workbook_pivot_field_set_sort`, `_set_subtotal_top`, `_set_number_format`, `fm_workbook_pivot_field_set_item_visible` | removed; `fm_sheet_update_pivot_field(wb, sheet_index, pivot_index, field_index, const fm_pivot_field_patch_t*)` and `fm_sheet_update_pivot_field_item(..., item_index, const fm_pivot_item_patch_t*)` carry every one of them. Each field has an `_op` companion (`KEEP`, `SET`, `CLEAR`), the whole patch is checked before anything changes, and an all-`KEEP` patch is a no-op |
| `fm_workbook_pivot_cache_get_worksheet_source(wb, cache_id, out_present, out_ref, out_sheet, out_name)`, `fm_workbook_pivot_cache_set_worksheet_source(wb, cache_id, present, ref, sheet, name)` | `fm_workbook_get_pivot_cache_worksheet_source(wb, cache_id, out_present, fm_pivot_worksheet_source_t* out)` and `fm_workbook_set_pivot_cache_worksheet_source(wb, cache_id, const fm_pivot_worksheet_source_t* or NULL)`; `NULL` clears the source |
| `fm_sheet_set_chart_type(wb, sheet_index, object_id, plot, kind, grouping)` | `fm_sheet_set_chart_type(..., plot, type, bar_direction, grouping)` with an `fm_chart_plot_type_t` and an `fm_chart_bar_direction_t`; `fm_chart_insert_kind_t` and `FM_CHART_INSERT_KIND_*` are removed. A zero `bar_direction` is `COLUMN` and is accepted for any type, `BAR` with a type other than bar is `FM_STATUS_INVALID_ARGUMENT` |
| `fm_workbook_set_default_font(wb, const fm_font_record*)` | takes `const fm_font_t*` |
| `fm_sheet_get_comment_count(wb, sheet_index, out)` | `size_t sheet_index`, the count through `size_t* out_count`, a `const` handle |
| `fm_sheet_set_comment(wb, sheet, row, col, author, text)` (empty text removed the note) | `size_t sheet_index` and a `const fm_comment_t*`; `NULL` removes the note and is a no-op when there is none |
| `fm_sheet_set_range_xf_index` | `size_t sheet_index` and `xf_index`, the rectangle as `const fm_cell_rect_t*` |
| `fm_sheet_set_row_hidden`, `fm_sheet_set_row_outline` | take the inclusive row span `(first, last)`, like `fm_sheet_set_column_hidden` and `fm_sheet_set_column_outline` |
| `fm_sheet_set_column_style`, `fm_sheet_set_row_style` | `fm_sheet_set_column_xf_index`, `fm_sheet_set_row_xf_index`, with `xf_index` as `size_t` |
| `fm_sheet_remap_range_style`, `fm_sheet_snapshot_range_style`, `fm_sheet_restore_range_style` | `fm_sheet_remap_range_xfs`, `fm_sheet_snapshot_range_xfs`, `fm_sheet_restore_range_xfs`; the rectangle is a `const fm_cell_rect_t*` |
| `fm_pivot_cells_destroy` | `fm_pivot_report_destroy`, on the `fm_pivot_report_t` handle |
| `fm_sheet_set_drawing_anchor` | removed; `fm_sheet_update_drawing(wb, sheet_index, object_id, const fm_drawing_patch_t*)` changes the name, description, hidden flag and anchor of any drawing object. The anchor is an embedded `fm_drawing_anchor_patch_t` that merges per field (an omitted `kind` keeps the object's kind); a slicer's name stays unchangeable |
| `fm_sheet_remove_shape`, `fm_sheet_remove_slicer`, `fm_sheet_shape_count` / `_at`, `fm_sheet_slicer_count` / `_at` | removed; `fm_sheet_remove_drawing` removes a picture, chart, shape, connector, group or slicer, and `fm_sheet_get_drawing_object_count` / `_at` enumerate every kind |
| `fm_sheet_get_validation_at(wb, uint32_t sheet, uint32_t index, fm_data_validation*)` | `fm_sheet_get_validation_at(const wb, size_t sheet_index, size_t validation_index, fm_data_validation_t*)` |
| `fm_workbook_save_with_diagnostics` filling an `fm_save_diagnostics_t` | fills an `fm_save_diagnostics_info_t` |
| `fm_sheet_get_view` filling an `fm_sheet_view_t` | fills an `fm_sheet_view_info_t` |
| the flat anchor fields of the shape, slicer and drawing records | one `fm_drawing_anchor_t anchor` (`kind`, `edit_as`, `from_row`, `from_col`, the four offsets, `to_row`, `to_col`, `width_emu`, `height_emu`); `fm_drawing_object_info_t` also carries `hidden` and `description` for every kind; `fm_chart_info_t` has no anchor, name or description |
| `fm_workbook_insert_rows(wb, sheet, row, count)`, `fm_workbook_delete_rows`, `fm_workbook_insert_cols`, `fm_workbook_delete_cols` | `fm_sheet_insert_rows`, `fm_sheet_delete_rows`, `fm_sheet_insert_columns` and `fm_sheet_delete_columns` take `(wb, sheet_index, first, last)`, inclusive; `last < first` or a `last` past the grid is `FM_STATUS_INVALID_ARGUMENT` |
| `fm_workbook_add_sheet(wb, utf8_name)`, `fm_workbook_move_sheet(wb, from_index, to_index)` | `fm_workbook_add_sheet(wb, name, insert_before, out_sheet_index)` and `fm_workbook_move_sheet(wb, from_index, insert_before)`, both with a position in the list before the call, as `fm_workbook_copy_sheet` has; `insert_before == sheet_count` appends. A `NULL` name takes the next default name |
| `fm_sheet_add_validation(wb, sheet, fm_data_validation v)`, `fm_sheet_remove_validation_at(wb, sheet, index)` | `fm_sheet_add_validation(..., out_validation_id)` returns an opaque `const char*` id (also `fm_data_validation_t.id`); `fm_sheet_get_validation`, `fm_sheet_update_validation` and `fm_sheet_remove_validation` take the id; an unknown id is `FM_STATUS_NOT_FOUND`; `_count` and `_at` stay for enumeration. `fm_validation_outcome_info_t.rule_index` is `rule_id` |
| `fm_sheet_remove_sparkline_group(wb, sheet_index, idx)` | takes a `const char* group_id` |
| `fm_sheet_add_threaded_comment(...)` with a required `fm_threaded_comment.id` | `id` may be `NULL` and the engine generates a GUID; the id comes back through `const char** out_id`. `created` left empty is the pinned clock, else now. An unknown parent id is `FM_STATUS_NOT_FOUND` |
| `fm_chart_series_t.idx`, `fm_chart_data_series_t.idx`, `fm_chart_unresolved_t.series_idx` and the series parameters of the chart functions | `series_id` (point `idx` stays a position) |
| `fm_external_link_record_t.index` | `book_number`, still 1-based: the `n` of `[n]` in formula text |
| `fm_pivot_filter_spec_t` `axis` and `field_name` | `field_index`; the axis is the field's |
| `fm_workbook_pivot_field_add_subtotal_fn`, `fm_workbook_pivot_field_clear_subtotal_fns` | `fm_sheet_set_pivot_field_subtotal_functions(wb, sheet_index, pivot_index, field_index, aggregations, count)` replaces the set; an empty list clears it and a repeated function is `FM_STATUS_INVALID_ARGUMENT` |
| (none) | `fm_workbook_get_pivot_cache_record_value` reads one cache cell as an `fm_value_t` |
| `fm_sheet_view_t.tab_hidden` | removed; `fm_sheet_get_visibility` and `fm_sheet_set_visibility` |
| (none) | `fm_sheet_update_cf_rule` merges an `fm_cf_rule_patch_t`; `fm_sheet_update_validation` merges an `fm_data_validation_patch_t`. A type change is `FM_STATUS_INVALID_ARGUMENT` |
| (none) | `fm_workbook_set_fill`, `_set_border`, `_set_cell_xf` and `_set_dxf` replace the record at an index, as `fm_workbook_set_font` does |
| (none) | `fm_get_profile_id_count` and `fm_get_profile_id_at` list the 28 Excel profile ids |
| (none) | `fm_sheet_set_value(wb, sheet_index, row, col, const fm_value_t*)` writes any non-array value that `fm_sheet_get_value` reads |
| `fm_sheet_add_row_break` | `fm_sheet_set_row_break(wb, sheet_index, row, const fm_row_break_t*)`, which upserts the break with its span and removes it with `NULL` |
| `fm_sheet_get_row_override_count` | `fm_sheet_get_row_layout_count` |
| `fm_sheet_set_header_footer`, `fm_sheet_set_page_margins`, `fm_sheet_set_print_options` | `fm_sheet_update_header_footer`, `fm_sheet_update_page_margins` (fields in inches with an `_in` suffix), `fm_sheet_update_print_options`; each merges |
| `fm_sheet_set_show_row_col_headers` | removed; `fm_sheet_update_view` sets `show_row_column_headers` |
| `fm_workbook_set_cell_phonetic_properties` and `fm_workbook_get_cell_phonetic_properties` taking `uint32_t` type and alignment | `fm_sheet_set_cell_phonetic_properties` and `fm_sheet_get_cell_phonetic_properties` with `int32_t` `fm_phonetic_type_t` and `fm_phonetic_alignment_t`; a value outside the set is `FM_STATUS_INVALID_ARGUMENT` |
| `fm_workbook_theme_script_font_count`, `fm_workbook_theme_script_font_at`, `fm_workbook_set_theme_script_font` taking `int32_t major` | `fm_workbook_get_theme_script_font_count`, `_at`, `fm_workbook_get_theme_script_font` and `fm_workbook_set_theme_script_font` take `int32_t scheme`, an `fm_font_scheme_t`; `FM_FONT_SCHEME_NONE` and any other value are `FM_STATUS_INVALID_ARGUMENT` |

Parameter names changed without a change of type: the out-parameter that
returns a sheet's index is `out_sheet_index` (it was `out_index`), and an
index into a collection is named for the collection (`table_index`,
`validation_index`, `column_index`, `plot_index`). The sheet
name argument of `fm_workbook_add_sheet` is `name`, and the string out-parameter
of `fm_workbook_get_sheet_name` is `out_name`.

**Enums.** Values of these typedefs moved. A member for "absent" is `-1`
where one exists, and the other members start at 0.

| Typedef | Before | After |
|---|---|---|
| `fm_chart_axis_crosses_t` | `ABSENT` 0, `AUTO_ZERO` 1, `MIN` 2, `MAX` 3, `AT` 4 | `ABSENT` -1, `AUTO_ZERO` 0, `MIN` 1, `MAX` 2, `AT` 3 |
| `fm_chart_bar_direction_t` | `ABSENT` 0, `COL` 1, `BAR` 2 | `ABSENT` -1, `COLUMN` 0, `BAR` 1 |
| `fm_blanks_as_t` (was `fm_chart_blanks_as_t`; also replaces `fm_sparkline_empty_cells_t`) | chart: `ABSENT` 0, `GAP` 1, `ZERO` 2, `SPAN` 3; sparkline: `GAP` 0, `ZERO` 1, `SPAN` 2 | `ABSENT` -1, `ZERO` 0, `GAP` 1, `SPAN` 2 |
| `fm_chart_legend_position_t` | `NONE` 0, `BOTTOM` 1, `TOP` 2, `LEFT` 3, `RIGHT` 4, `TOP_RIGHT` 5 | `NONE` -1, `RIGHT` 0, `BOTTOM` 1, `TOP_RIGHT` 2, `LEFT` 3, `TOP` 4 |
| `fm_pivot_sort_order_t` | `DESCENDING` 0, `ASCENDING` 1, `MANUAL` 2 | `MANUAL` 0, `ASCENDING` 1, `DESCENDING` 2 |
| `fm_comparison_operator_t` (replaces `fm_filter_operator_t` and the bare operator numbers of `fm_cf_rule_t` and `fm_data_validation`) | one table per feature (see the JavaScript table under [Enums](#enums)) | `BETWEEN` 0, `NOT_BETWEEN` 1, `EQUAL` 2, `NOT_EQUAL` 3, `LESS_THAN` 4, `LESS_THAN_OR_EQUAL` 5, `GREATER_THAN` 6, `GREATER_THAN_OR_EQUAL` 7 |
| `fm_date_period_t` (`fm_cf_rule_t.time_period` was a bare number) | today 0, yesterday 1, tomorrow 2, last 7 days 3, this week 4, last week 5, next week 6, this month 7, last month 8, next month 9 | `TODAY` 0, `YESTERDAY` 1, `TOMORROW` 2, `LAST_7_DAYS` 3, `THIS_MONTH` 4, `LAST_MONTH` 5, `NEXT_MONTH` 6, `THIS_WEEK` 7, `LAST_WEEK` 8, `NEXT_WEEK` 9 |
| `fm_anchor_kind_t`, `fm_anchor_edit_as_t` | `ONE_CELL` 0, `TWO_CELL` 1, `ABSOLUTE` 2 (`edit_as`: `TWO_CELL` 0, `ONE_CELL` 1, `ABSOLUTE` 2) | `DEFAULT` 0 (the operation's default kind on input, the anchor's own kind for `edit_as`), `ONE_CELL` 1, `TWO_CELL` 2, `ABSOLUTE` 3 |
| `fm_shape_paint_t` | `UNSET` 0, `SOLID` 1, `NONE` 2, `OTHER` 3 | `DEFAULT` 0 (no fill or outline of its own: the shape style's), `SOLID` 1, `NONE` 2, `OTHER` 3 |

New enum typedefs replace `#define` sets and plain integers:
`fm_page_orientation_t` (`PORTRAIT` 0, `LANDSCAPE` 1, replacing
`FM_ORIENTATION_DEFAULT` 0, `_PORTRAIT` 1 and `_LANDSCAPE` 2; an absent
orientation prints portrait), `fm_external_link_kind_t` (the values of
`FM_EXTERNAL_LINK_KIND_*` are unchanged), `fm_sheet_view_mode_t` (`NORMAL` 0,
`PAGE_BREAK_PREVIEW` 1, `PAGE_LAYOUT` 2), `fm_cf_data_bar_direction_t`
(`CONTEXT` 0, `LEFT_TO_RIGHT` 1, `RIGHT_TO_LEFT` 2) and
`fm_cf_data_bar_axis_position_t` (`AUTOMATIC` 0, `MIDDLE` 1, `NONE` 2).
The closed sets that were text or `#define` constants are enums as well:
`fm_table_totals_row_function_t` (`NONE` 0, `SUM` 1, `MIN` 2, `MAX` 3,
`AVERAGE` 4, `COUNT` 5, `COUNT_NUMBERS` 6, `STD_DEV` 7, `VAR` 8, `CUSTOM` 9),
`fm_table_style_element_type_t` (`WHOLE_TABLE` 0 to `PAGE_FIELD_VALUES` 27),
`fm_chart_tick_label_position_t` (`ABSENT` -1, `NEXT_TO` 0, `HIGH` 1, `LOW` 2,
`NONE` 3), `fm_width_platform_t` (`WIN` 0, `MAC` 1), `fm_page_order_t`
(`DOWN_THEN_OVER` 0, `OVER_THEN_DOWN` 1), `fm_phonetic_type_t` and
`fm_phonetic_alignment_t` (the values of the `FM_PHONETIC_*` constants they
replace), `fm_effective_style_source_t` (`CELL` 0, `ROW` 1, `COLUMN` 2,
`DEFAULT` 3), `fm_theme_source_t` (`PART` 0, `DEFAULT` 1, `UNPARSEABLE` 2) and
`fm_icon_set_name_t`, which gains `3_STARS` 17, `3_TRIANGLES` 18, `5_BOXES` 19
and `NO_ICONS` 20. The `fm_pivot_filter_kind_t` members for relative periods
drop the `DATE_` infix (`FM_PIVOT_FILTER_KIND_TODAY`); their values are
unchanged.

**Zero-initialised records.** A record that was zeroed and then filled in
field by field reads differently when a field is left at 0:

| Field | 0 meant | 0 means |
|---|---|---|
| an anchor's `kind` and `edit_as` | `FM_ANCHOR_KIND_ONE_CELL`; `FM_ANCHOR_EDIT_AS_TWO_CELL` | `FM_ANCHOR_KIND_DEFAULT` (picture one-cell, chart, shape and slicer two-cell); `FM_ANCHOR_EDIT_AS_DEFAULT` (the anchor's own kind) |
| a legend position (`fm_chart_legend_position_t`) | no legend | `FM_CHART_LEGEND_POSITION_RIGHT` |
| chart blanks, crossing and bar direction | absent | zero, automatic zero and column |
| a sparkline group's `empty_cells_as` | `FM_SPARKLINE_EMPTY_GAP` | `FM_BLANKS_AS_ZERO` |
| a pivot field sort order (`fm_pivot_sort_order_t`) | descending | manual |
| a conditional-format rule's `op` | less than | between |
| an AutoFilter condition's operator | equal | between, which the setter rejects |

Set the `ABSENT` or `NONE` member (`-1`) to say "not there".
`fm_chart_display_patch_t` rejects `-1` as a value to set: clear the field
through its operation instead.

**Structs.** The first column gives the former name of a renamed record; the
current spelling is in the [Name migration](#name-migration) table.

| Struct (former name) | Change |
|---|---|
| (none) | `fm_defined_name_t` and `fm_iterative_settings_patch_t` are new (see the functions above); `fm_defined_name_t` is both the input of `fm_workbook_set_defined_name` and the record the getters fill |
| (none) | `fm_table_spec_t`, `fm_pivot_spec_t`, `fm_parse_reference_spec_t`, `fm_pivot_date_group_spec_t`, `fm_pivot_worksheet_source_t`, `fm_pivot_style_options_patch_t`, `fm_pivot_field_patch_t` and `fm_pivot_item_patch_t` are new (see the functions above); each begins with `struct_size` |
| (none) | `fm_filter_condition_t` and `fm_chart_data_source_spec_t` are new element records (no `struct_size`): one custom AutoFilter condition (`op`, `value`) and one chart series input (`kind`, `formula`, `literal`, `literal_count`) |
| `fm_filter_column` | `custom_count`, `op1`, `val1`, `op2`, `val2` are `const fm_filter_condition_t* conditions` and `uint32_t condition_count`; another count for a custom column is `FM_STATUS_INVALID_ARGUMENT`, and a non-custom column reads 0 |
| `fm_pivot_field_info_t` | `show_detail` is appended: 1 only when every item and the default expand marker are expanded, and a field with no items reads 1 |
| `fm_chart_insert` | `fm_chart_spec_t`; `kind` is `type` (an `fm_chart_plot_type_t`) and `bar_direction` is new; a zero `legend` is `FM_CHART_LEGEND_POSITION_RIGHT`, the file default, where the bindings pass `BOTTOM` for an omitted option |
| `fm_chart_series_patch` | the six `const char*` references (`name`, `categories`, `values`, `x`, `y`, `size`) are each an `_op` and an `fm_chart_data_source_spec_t`: `REF` writes `formula`, `LITERAL` writes the points, `NONE` removes the input of `name`, `categories` and `x` and is `FM_STATUS_INVALID_ARGUMENT` for the others |
| `fm_chart_legend_patch` | `hidden_series` and `hidden_series_count` are `hidden_entries` and `hidden_entries_count` (the op is `hidden_entries_op`), as in the new `fm_chart_legend_t`; `fm_sheet_set_chart_filter` keeps `hidden_series` |
| `fm_chart_axis_patch`, `fm_chart_labels_patch` | gain `minor_unit`, `minor_gridlines`, `position` and `source_linked`; and `source_linked` and `deleted`, each with its `_op` |
| `fm_chart_point_patch` | an all-`KEEP` patch changes nothing (it removed the point's formatting); a point whose override ends empty loses its `c:dPt`, and `CLEAR` on a line's `none` resets it to 0 |
| `fm_image_insert` | `fm_drawing_placement_spec_t`, which also carries `hidden`; `descr` is `description`; the anchor fields (`anchor_kind`, `edit_as`, `row`, `col`, `row_off_emu`, `col_off_emu`, `width_emu`, `height_emu`) are the `anchor` member, an `fm_drawing_anchor_t` |
| `fm_drawing_object` | `fm_drawing_object_info_t`; `descr` is `description`; in the anchor, `from_row_off`, `from_col_off`, `to_row_off` and `to_col_off` are `from_row_offset_emu` and so on, and `cx` and `cy` are `width_emu` and `height_emu` |
| `fm_shape` | `fm_shape_info_t`; the same offset and size renames; `prst` is `preset_geometry`; the anchor and the name are read from `fm_drawing_object_info_t` |
| `fm_slicer` | `fm_slicer_info_t`; `row_height` (EMU) is `row_height_pt` (points); the anchor and the name are read from `fm_drawing_object_info_t` |
| `fm_sheet_view_t` | `fm_sheet_view_info_t`; `tab_hidden` is removed; `view_mode` is an `int32_t` `fm_sheet_view_mode_t` instead of text |
| `fm_page_setup_t` | `orientation` is an `int32_t` `fm_page_orientation_t` (was `uint32_t`) |
| `fm_external_link_record_t` | `kind` is an `int32_t` `fm_external_link_kind_t` (was `uint32_t`) |
| `fm_chart_data_info_t` | `point_count`, `series_count`, `category_label_count`, `category_value_count`, `category_level_count`, `axis_title_count`, `unresolved_count` are `uint32_t` |
| `fm_chart_data_points_t`, `fm_chart_data_source_t` | `count` and `literal_count` are `uint32_t` |
| `fm_chart_data_series_t` | `plot_index`, `point_count` and the `range_count` array are `uint32_t` |
| `fm_chart_info_t` | `legend_hidden_count`, `plot_count`, `axis_count`, `hidden_category_count` are `uint32_t` |
| `fm_chart_legend_patch`, `fm_chart_plot_t`, `fm_chart_series_t` | `hidden_series_count` (now `hidden_entries_count`), `axis_id_count`, `series_count`, `point_count` are `uint32_t` |
| `fm_fill_record` | `gradient_stop_count` is `uint32_t` |
| `fm_sparkline_group` | `sparkline_count` is `uint32_t` |
| `fm_cell_xf` | `reading_order` is `int32_t` (an `fm_reading_order_t`) |
| `fm_table_column_info_t` | `totals_row_function` is an `int32_t` `fm_table_totals_row_function_t` instead of text; `fm_table_column_patch_t.totals_row_function` is the same enum |
| `fm_table_style_element_t` | `type` is an `int32_t` `fm_table_style_element_type_t` instead of text |
| `fm_chart_axis_t`, `fm_chart_axis_patch` | `tick_label_pos` is `tick_label_position`, an `int32_t` `fm_chart_tick_label_position_t` (`FM_CHART_TICK_LABEL_POSITION_ABSENT` -1 when the axis has no element) instead of text |
| `fm_width_model` | `platform` is an `int32_t` `fm_width_platform_t` instead of text |
| `fm_cf_icon_t` | `icon_set` is an `int32_t` `fm_icon_set_name_t` instead of text |
| `fm_cf_rule_t`, `fm_data_validation`, the AutoFilter column | the operator fields are an `fm_comparison_operator_t`; the sparkline group's `empty_cells_as` is an `fm_blanks_as_t` and `FM_BLANKS_AS_ABSENT` is `FM_STATUS_INVALID_ARGUMENT` there |
| `fm_styles_batch` | `font_count`, `fill_count`, `border_count`, `cell_xf_count`, `num_fmt_count` are `uint32_t` |

**Records, patches and getters.**

| Before | After |
|---|---|
| every `_t` record under a role | renamed by role: `_spec_t` input only, `_info_t` output only, plain `_t` read and written, `_patch_t` partial update; the table under [Name migration](#name-migration) lists each one. `fm_sheet_rect_t`, `fm_drawing_anchor_t` and the chart fill, line, marker and label value types keep their names |
| per-record patch conventions (`engaged` flags, `-1` sentinels, `FM_CHART_PATCH_*`) | `fm_patch_op_t` (`FM_PATCH_OP_KEEP` 0, `_CLEAR` 1, `_SET` 2) and one `<field>_op` before each field of the shape, slicer, chart, table, iterative-settings, view and threaded-comment patches. `FM_CHART_PATCH_*` and `fm_chart_patch_op_t` are gone |
| `fm_sheet_insert_shape`, `fm_sheet_update_shape` and the slicer pair taking an `FM_SHAPE_FIELD_*` / `FM_SLICER_FIELD_*` mask | `fm_shape_spec_t` and `fm_slicer_spec_t` for the adds and `fm_shape_patch_t` and `fm_slicer_patch_t` (kind-specific fields only) for the updates; the masks are removed. The read records are `fm_shape_info_t` and `fm_slicer_info_t` |
| the fourteen single-field view setters (`fm_sheet_set_zoom` and the rest; `fm_sheet_set_selection` stays) | `fm_sheet_update_view(wb, sheet_index, const fm_sheet_view_patch_t*)` |
| `fm_workbook_table_at` and `fm_workbook_table_get_info` | `fm_sheet_get_table_at(wb, sheet_index, table_index, fm_table_info_t*)`, which carries the name and range |
| a positional getter without `_at` | `_at` on every positional getter, for example `fm_workbook_get_external_link_at` |
| a keyed or optional read that failed with a not-found status for "none", or that wrote a sentinel | `(…, int32_t* out_present, T* out)`; `fm_sheet_get_chart_data`, `fm_sheet_get_auto_filter`, `fm_workbook_get_iterative`, `fm_workbook_get_defined_name` and `fm_sheet_get_comment_visible` use it, and a miss is `*out_present == 0` with status OK |
| optional strings as `""` or a flag-only record | `NULL`, with `out_present` or a `has_*` field |
| `fm_chart_data_category_at(…, double* out_value)` and NaN markers | `int32_t* out_has_value` before `out_value`; `today_serial` of `fm_sheet_evaluate_cf_range` takes `int32_t has_today_serial` before it, with no NaN switch |
| `FM_CELL_STYLE_BUILTIN_ID_NONE` | `has_builtin_id` on `fm_cell_style_t` |
| a list getter with a per-list buffer shape | `(T* out, size_t capacity, size_t* out_count)` |
| series `c:idx`, xf indices and counts | `uint32_t`; positional indices are `size_t`, and every enum field is `int32_t` |
| 52 status codes | removed from `formulon_status.h` because no function returned them: 1, 5, 1002, 1008, 1011, 1012, 2000 to 2008, 3000 to 3003, 3005 to 3008, 4000, 4001, 4002, 4004, 4006, 5001, 5008, 5009, 5013, 5015, 5017, 5052, 6000 to 6004, 6006, 6007, 7004, 8000 to 8004, 9000, 9001, 9003 and 9004. The matching `StatusCode` members and Python names go with them; the general codes 7 and 9 and the specific codes still returned stay |
| `fm_workbook_spill_info(wb, sheet, row, col, fm_spill_info_t*)` and `engaged` in `fm_spill_info_t` | `fm_sheet_get_spill(wb, sheet_index, row, col, int32_t* out_present, fm_spill_info_t* out)` and `fm_sheet_get_array_formula(..., int32_t* out_present, fm_array_formula_info_t* out)`; the records lose `engaged`, and a blocked anchor is present with `blocked` set |
| `fm_cell_get_xf_index(wb, sheet, row, col, uint32_t* out_xf_index)`, which wrote 0 for a cell with no XF | `fm_sheet_get_cell_xf_index(…, int32_t* out_present, uint32_t* out_xf_index)` |
| `fm_pivot_style_info_t.present` | `int32_t* out_present` before the out record of `fm_sheet_get_pivot_style_options`, and of `fm_sheet_get_table_style_options`; a NULL `out_present` is `FM_STATUS_INVALID_ARGUMENT` |
| `fm_workbook_pivot_field_clear_date_group` | removed; `fm_sheet_set_pivot_field_date_group(..., NULL)` clears the grouping |
| `kSheetIndexOutOfRange` (5050), `kInvalidSheetName` (5051), `kSheetCountLimitExceeded` (5053) | `FM_STATUS_SHEET_INDEX_OUT_OF_RANGE` (10), `FM_STATUS_INVALID_SHEET_NAME` (11) and `FM_STATUS_SHEET_COUNT_LIMIT_EXCEEDED` (12). 5050 to 5053 are retired and never reused |
| `<field>_engaged` presence fields (`fm_page_setup_t`, `fm_page_margins_t`, `fm_print_options_t`, `fm_header_footer_t`, `fm_cf_rule_t`, `fm_dxf_record` records and the others) | `has_<field>` |
| a `-1`, `UINT32_MAX` or `0xFFFFFFFF` meaning "none" in a record field: `fm_pivot_field_info_t.cache_field_index`, `fm_pivot_item_info_t.cache_index`, `fm_pivot_data_field_spec_t` `show_as_base_field` and `show_as_base_item`, `fm_function_metadata_t.max_arity`, `fm_filter_column.dxf_id` | `has_<field>` before a `uint32_t` field. `fm_filter_column_t` has no `struct_size` and gains `has_dxf_index`, so its array stride changes |
| `-1` as a pivot `axis` (`fm_pivot_field_spec_t`, `fm_pivot_field_info_t`, `fm_workbook_pivot_field_set_axis`) | `FM_PIVOT_AXIS_ABSENT` |
| `fm_cf_match_t.bar_border_engaged` and the colour presence flags of `fm_cf_rule_t` (`data_bar_*_engaged`) | removed; an absent colour is `FM_COLOR_KIND_NONE`, as in every other colour field. A sparkline group colour is cleared with `CLEAR` |
| `-1` as the workbook scope of a defined name (the `local_sheet_id` arguments), the numbers 1048828 and 1048829 as show-as base items | `FM_DEFINED_NAME_SCOPE_WORKBOOK`, `FM_PIVOT_SHOW_AS_BASE_PREVIOUS` and `FM_PIVOT_SHOW_AS_BASE_NEXT` |
| `fm_chart_title_t` `text` and `formula` read as `""` when the title kind does not use them | `NULL` for the unused field; as input `NULL` and `""` both read as empty |
| a C parameter that names a column or a plot group as `index` | `column_index` and `plot_index` (names only, no ABI change) |
| `fm_sheet_cf_add_rule(wb, sheet_index, fm_cf_rule_t rule, size_t* out_index)` | `fm_sheet_add_cf_rule(..., const fm_cf_rule_t*, const char** out_cf_rule_id)` returns the stored id (a view valid until the next scratch-backed read) |
| `fm_sheet_cf_remove_at` (a position) | `fm_sheet_remove_cf_rule` and `fm_sheet_update_cf_rule` take the rule's id (`const char* id`); a missing id is `FM_STATUS_NOT_FOUND` |
| (none) | `fm_sheet_get_cf_rule(wb, sheet_index, id, int32_t* out_present, fm_cf_rule_t* out)` reads a rule by id; `fm_sheet_get_cf_rule_at` stays the positional read |
| `fm_sheet_update_protected_range`, `fm_sheet_remove_protected_range` (a position) | take the range's name |
| `fm_sheet_get_protected_range(wb, sheet, index, out)` (the positional read) | `fm_sheet_get_protected_range_at`; `fm_sheet_get_protected_range(wb, sheet_index, name, int32_t* out_present, out)` now reads by name |
| `fm_workbook_pivot_cache_field_add_shared_item_number`, `_text`, `_bool`, `_blank`, `_error` | `fm_workbook_add_pivot_cache_field_shared_item_number` and so on, with `_date` added; each takes a trailing `size_t* out_shared_item_index` |
| `fm_workbook_add_person(wb, person)` | `fm_workbook_add_person(wb, person, const char** out_person_id)` returns the stored id |
| `fm_chart_insert.title` (`const char*`; `NULL` automatic, `""` none) | `int32_t has_title` and `fm_chart_title_t title`; the struct layout changes |
| (none) | `fm_sheet_get_pivot_page_field_order` and `fm_sheet_set_pivot_page_field_order` |
| `FM_PIVOT_AXIS_COL`, `FM_PIVOT_SHOW_AS_PERCENT_OF_COL`, `_RUNNING_TOTAL_IN_COL`, `_PERCENT_OF_PARENT_COL` | `FM_PIVOT_AXIS_COLUMN`, `FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_COLUMN`, `_RUNNING_TOTAL_IN_COLUMN`, `_PERCENT_OF_PARENT_COLUMN` (same numbers) |

**Additive changes (no ABI break).** `fm_convert_civil_to_serial` and
`fm_convert_serial_to_civil` convert between a civil time and a serial for
either date system; see [Added](../../CHANGELOG.md#added). `fm_sheet_update_protected_range`
keeps the stored ranges when `ranges` is `NULL` and `range_count` is 0, which
lets a caller change the other fields of a range whose address the engine
cannot parse. The drawing functions for name, z-order, snapshot and restore
accept shapes, and the z-order functions also accept slicers.
`fm_sheet_add_slicer` and `fm_sheet_update_slicer` accept an ascending sort
order.
