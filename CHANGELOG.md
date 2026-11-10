# Changelog

All notable changes to Formulon are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.19.0] - 2026-11-10

### BREAKING

- **Names:** one naming convention covers every public function, type and
  field in C, JavaScript and Python, and the former names are removed with no
  aliases. The table under [Name migration](#name-migration) lists each former
  name and what replaces it. Python static methods are module-level functions,
  and the `version` / `version_string` aliases are gone.
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
- **JavaScript, Python:** every collection is read with one `list*` method
  (`listCfRules()`, `listValidations()`, `listTables()`, `listFonts()`
  and the others in the migration table) that returns a fresh array. A count
  method paired with a getter by index (`tableCount()` + `tableAt(i)`) is
  removed together with the getter; `listTables()[i]` replaces the pair. A
  phonetic run list is read with `getCellPhoneticRuns(row, col)` instead of
  `getCellPhoneticRunCount` and `getCellPhoneticRun`.
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
- **C ABI:** an index past the end of a collection other than sheets returns
  `FM_STATUS_INDEX_OUT_OF_RANGE`, and removing a record by a natural key that
  does not exist returns `FM_STATUS_NOT_FOUND`. Error message text no longer
  names C functions and is for display, not comparison.
  `fm_workbook_get_sheet_index` finds a sheet by name and
  `fm_get_version_number` returns the version as one number.
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
  `getCellXfIndex` is `null` for a cell with no XF of its own,
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
- **Python:** `FormulonError(code, message="", context="", parse_error=None)`
  builds an error, and `str(e)` is `[CODE_NAME] message` with the context only
  in `e.context`. `FormulaSyntaxError` carries `parse_error` (a `ParseError`
  with `code`, `start`, `end`, `line`, `column`). A ragged `set_values` grid is
  `InvalidArgumentError`, and `parse_reference` with only one of `base_row` and
  `base_col` is `TypeError`, as in JavaScript.
- **Python:** `copy.copy`, `copy.deepcopy` and pickling of a `Workbook`, of any
  handle class and of the `*Ooxml` views raise
  `TypeError: Workbook cannot be copied; use Workbook.load(wb.save())`.
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
- **C ABI:** a few records and functions change shape; a caller compiled
  against an earlier header must be recompiled. See [C ABI changes](#c-abi-changes).
- **C ABI:** `fm_workbook_save` takes the `int32_t format` that
  `fm_workbook_save_as` took, and `fm_workbook_save_as` is removed:
  `fm_workbook_save(wb, FM_WORKBOOK_FORMAT_XLSX, &bytes, &len)`.
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
- **JavaScript, Python:** a boolean argument or field takes `true` or `false`
  only. `setDate1904('yes')`, `setCalcOnSave(null)`, `parseReference(text,
  { r1c1: 1 })`, `updateView({ showGridLines: 1 })` and the Python counterparts
  (`set_date1904(1)`, `format_reference(range, absolute=1)`) throw `TypeError`;
  they used to be read as truthy or falsy. `Workbook.setIterative` errors are
  labelled `Workbook.updateIterative`, and `PivotCache` errors name
  `PivotCache.<method>`. A static factory (`Workbook.createDefault`,
  `Workbook.load`) and a module function refuse extra arguments with `takes at
  most <n> arguments`.
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
  filter passes back to `addFilter`, and `TableColumnTotalsInput` carries the
  totals fields only, so a read `TableColumn` passes to `setColumnTotals`
  (its `name` is ignored). `Sheet.addProtectedRange(input)` names its
  parameter `input`.
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
- **C ABI:** `fm_sheet_set_pivot_field_date_group` (`Pivot.setFieldDateGroup`,
  `Pivot.set_field_date_group`) takes a mask of `FM_PIVOT_PERIOD_*` bits, a
  calendar, an interval in days and start and end serials (NaN for automatic)
  in place of a granularity and year arguments. `fm_pivot_date_grouping_t` and
  the `date_*` members of `fm_pivot_field_info_t` are gone; the struct now
  reports `date_periods`, `date_calendar`, `group_base_index`,
  `insert_blank_row` and `fill_down_labels`. The old "year and month" grouping
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
- **Command line:** a formula syntax error exits with status 65. `eval -`
  reads the formula from standard input, and `recalc -o` takes the file format
  from the output extension.
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
- **JavaScript, Python:** `getCellFormat` reports `source`, where the
  effective format comes from; the XF index stays on the `ooxml` layer
  (`ws.ooxml.getCellXfIndex`). `numberFormat` is the name everywhere
  (`getNumberFormat(id)` / `addNumberFormat(code)` for the table of formats);
  `numFmtId` remains only for the XF-level id. The registries of table styles
  and cell styles (`list`, `get`, `set`, `remove`) are on `Workbook`.
  `activeSheet()` returns the `Sheet` instead of its index, and
  `hasPassthrough(path)` returns a
  boolean. `setDefaultFont` gains `getDefaultFont()`.
- **JavaScript:** record types with required keys are separate from their
  patches: `ShapeInput` and `ShapePatch`, and the same for slicers and the
  chart series add. `signature` is the one name for a signature field.
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
- **C ABI, JavaScript, Python:** the theme script font takes `scheme`, a
  `FontScheme` (`Major` or `Minor`), in place of the `major` boolean:
  `wb.setThemeScriptFont({ scheme, script, typeface })`,
  `wb.getThemeScriptFont({ scheme, script })`, Python
  `set_theme_script_font(ThemeScriptFont(scheme=..., ...))`, and `int32_t scheme` in the C
  functions. `FontScheme.None` and any other value are `InvalidArgument`.
- **C ABI:** `fm_workbook_evaluate_formula` is removed; use
  `fm_sheet_evaluate_formula_array` and take the top-left cell of a result.
  The bindings' `evaluateFormula` does that.

- **JavaScript, Python:** a setter that takes a record has an `*Input` type
  (`FontInput`, `FillInput`, `BorderInput`, `CellXfInput`, `DxfInput`,
  `CellStyleInput`, `TableStyleInput`, `CommentInput`, `PersonInput`,
  `ArrayFormulaInput`, `WorkbookProtectionInput`, `ProtectedRangeInput`,
  `SparklineInput`, `PrintTitlesInput`, `AutoFilterConditionInput`,
  `ChartFilterInput`, `PivotWorksheetSourceInput` and the others in
  [Added](#added)), and also accepts the record its getter returns. An omitted
  key takes the default its type states, and the presence flags are gone:
  `CellXf.hasAlignment`, `hasHorizontalAlign`, `hasVerticalAlign`,
  `hasWrapText`, `hasJustifyLastLine` and `hasProtection`, and
  `PivotWorksheetSource.present`.
  `CfValueObjectInput` is `CfValueObject` and `CfIconInput` is `CfIcon`.
- **JavaScript, Python:** `DataTableInput` now names the data-table spec;
  the former cell-reference record is `DataTableInputCell`. The spec was
  `DataTableSpec` in JavaScript.
- **JavaScript:** `Patch<T, Clearable>` names the keys that take `null`.
  `null` on any other key is a compile-time error (and a `TypeError` at run
  time, see above); an empty patch changes nothing.
- **JavaScript, Python:** a key that holds a cell rectangle is `range`, and one
  that holds several is `ranges`. `rect` is `range` on `ArrayFormula`,
  `Spill`, `ParsedReference`, `DataTable` and
  `PivotWorksheetSource` (which also replaces `ref`), `ref` is `range` on
  `AutoFilterSortCondition` and `AutoFilterSortState`, and `source` is `range`
  on `ChartSourceInput`, `ChartInput` and `ChartPreviewOverridesInput`.
  A sparkline's `source` is a `SheetRect` (it was the text `formula`), and
  `SparklineGroup.dateAxisRef` is `dateAxisRange`, a `SheetRect` or `null`.
  Input takes a `SheetRect` or an A1 range where the type says so.
- **JavaScript:** `DefinedName.localSheetId` is `scope` (`null` for workbook
  scope, else the 0-based index of the sheet), and `DefinedNameInput.scope`
  takes a `SheetRef`. C keeps `local_sheet_id`.
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
  and `totalsRowFormula` replace `function`, `label` and `formula` in
  `Table.setColumnTotals`.
- **JavaScript:** `getRangeXfBasesExcluding(range, excluded)` takes a list of
  cells (`CellAddress` or A1) in place of two arrays of rows and columns,
  `remapRangeXfs(range, mapping)` takes a `Map` from old to new index in place
  of two arrays, and `convertCivilToSerial(civil, { date1904 })` and
  `convertSerialToCivil(serial, { date1904 })` take an options object in place
  of a boolean. A `date1904` that is not a boolean throws `TypeError`.
- **Python:** nine signatures take a record in place of keywords:
  `set_cell_phonetic_properties(row, col, properties)` takes a
  `PhoneticProperties`, `add_chart(input)` a `ChartInput`,
  `Chart.set_source(source)` a `ChartSourceInput`,
  `PivotCache.set_worksheet_source(source)` a `PivotWorksheetSourceInput` or
  `None`, `Pivot.set_field_date_group(field_index, group)` a
  `PivotDateGroupInput`, `set_selection(selection)` a `SheetSelectionInput`,
  `Table.set_column_totals(column, totals)` a `TableColumnTotalsInput`,
  `set_theme_colors(colors)` a `ThemeColorsInput` and
  `set_theme_script_font(font)` a `ThemeScriptFont`. The old keyword forms
  raise `TypeError`.

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
- **C ABI, JavaScript, Python:** `insertRows`, `deleteRows`, `insertColumns`
  and `deleteColumns` take `(first, last)` instead of `(start, count)`;
  a call written for the old form inserts a different number of rows.
  `ws.insertRows(3, 5)` inserts three rows, and `first > last` is
  `InvalidArgument`.
- **C ABI, JavaScript, Python:** `insertRows` and `insertColumns` fail with
  `PreconditionFailed` and change nothing when they would push a cell that
  holds a value or formula past the last row or column of the grid; they used
  to drop it.
- **C ABI, JavaScript, Python:** `Sheet.insertCells(range, shift)` and
  `deleteCells(range, shift)` fail with `PreconditionFailed` and change nothing
  when the shifted region meets a table, a multi-cell array formula or a data
  table that extends beyond the band, a PivotTable edge or inside, or when a
  value or formula would be pushed off the grid; `InvalidArgument` for a wrong
  shift direction or a whole-column range shifted down (whole-row shifted
  right).
- **JavaScript, Python:** `addValidation` and `addSparklineGroup` return the
  new rule's id (`ValidationId`, `SparklineGroupId`), and
  `removeValidation`, `updateSparklineGroup` and `removeSparklineGroup` take a
  `ValidationRef` or `SparklineGroupRef`: an id, or a position as before.
  `DataValidation` and `SparklineGroup` carry `id`; `ValidationOutcome`
  names the covering rule as `ruleId` in place of `ruleIndex` (`null` when no rule covers the cell).
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
- **Python:** `Mention` / `MentionInput` and `PhoneticRun` spans are code-point
  offsets into the Python `str`, as `len` and slicing count, instead of UTF-16
  code units; a character outside the BMP counts 1. JavaScript keeps UTF-16
  offsets, and length limits stay in UTF-16 code units on every surface.
- **JavaScript, Python:** `ThreadedCommentInput.id` is optional (the engine
  generates one) and `personId` is required. Python `ChartCategories.format_code`
  is `number_format`.
- **C ABI, JavaScript, Python:** `PivotAxis.Col` is `PivotAxis.Column`
  (`FM_PIVOT_AXIS_COL` is `FM_PIVOT_AXIS_COLUMN`; Python `PivotAxis.COLUMN`),
  and `ChartBarDirection.Col` is `ChartBarDirection.Column`. The
  show-values-as members `PercentOfCol`, `RunningTotalInCol` and
  `PercentOfParentCol` are spelled out the same way (C
  `FM_PIVOT_SHOW_VALUES_AS_PERCENT_OF_COLUMN`, `_RUNNING_TOTAL_IN_COLUMN` and
  `_PERCENT_OF_PARENT_COLUMN`; Python `PERCENT_OF_COLUMN` and so on). `Col`
  remains only inside coordinate names such as `firstCol`.
- **Python:** the status ranges have exception classes. `IoError`,
  `EvaluationError`, `FunctionError`, `RecalcError`, `SecurityError`,
  `BindingError`, `PivotError` and `PrintError` are `FormulonError` subclasses,
  one per range, and a status without a narrower class raises the class of its
  range. `InvalidFileError` is an `IoError` and a `ValueError`.
  `FormulaSyntaxError` is a `FormulonError` and a `ValueError` and is no longer
  an `InvalidArgumentError`: code that caught a failed parse with
  `except InvalidArgumentError` must catch `FormulaSyntaxError` (or
  `ValueError`) instead.
- **JavaScript:** an object passed where a record is expected refuses a key it
  does not define with `TypeError` at every depth. That now covers the options
  of `Pivot.addFieldItem`, the `CivilTime` objects (`setPinnedNow`,
  `convertCivilToSerial`, the `today` option of `evaluateCfRange`, a threaded
  comment's `created`), the date-system options and `Color` objects
  (`setTabColor` and every colour key); a misspelt key used to be ignored.
- **C ABI, JavaScript, Python:** a chart title of `''` is `InvalidArgument` in
  `addChart`, as it already was in `setTitle`; `null` (`None`) is no title and
  an omitted title is Excel's automatic one. `fm_chart_spec_t` (was
  `fm_chart_insert`) replaces `const char* title` with `int32_t has_title` and `fm_chart_title_t title`
  (`has_title == 0` is automatic, a title of kind `NONE` is no title), so the
  insert takes the title shape `fm_sheet_set_chart_title` takes.
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
  name references the row, and the columns keep their saved totals labels and
  functions. A toggle that changes nothing is a no-op, and with `range` in the
  same patch the range is final and `totalsRow` only sets the flag.
  `fm_sheet_update_table` behaves the same.
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
- **JavaScript, Python:** adds that returned nothing return what they created:
  `addMerge` the stored `CellRect`, `addDataTable` the top-left `CellAddress` of
  the result range, `addPerson` the stored id and `PivotCache.addFieldSharedItem`
  (`add_field_shared_item`) the new item's index.
- **Command line:** `dump` takes `--view formulas|values|sheets|metadata`
  (default `formulas`; `--view=values` also works); `--formulas`, `--values`,
  `--sheets` and `--metadata` are removed and exit with status 64. `paginate
  --sheet` takes a sheet name or an index and fails for an unknown sheet as
  `dump --sheet` does. `recalc --json` prints
  `{"format":1,"status":"ok","output":PATH,"bytes":N,"warnings":[{"kind":...,"count":N}]}`
  on stdout and puts the loss warnings in `warnings` instead of stderr; it is
  refused with status 64 together with `-o -`.

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
| `setComment(sheet, row, col, author, text)` (empty text removed the comment) | `ws.setComment(row, col, { text, author? })` writes the comment, empty text included (an omitted `author` is `""`); `ws.setComment(row, col, null)` removes it and does nothing when the cell has none |
| `setWorkbookProtection(p)` / `setSheetProtection(sheet, p)` (an all-zero record cleared them) | `wb.setWorkbookProtection(null)` and `ws.setSheetProtection(null)` remove them; an all-zero workbook-protection record still does. Sheet protection is described under [Behaviour changes](#behaviour-changes-under-an-unchanged-name) |
| `setCellPhoneticRuns(sheet, row, col, runs)` | `ws.setCellPhoneticRuns(row, col, runs)`; `null` and `[]` clear |
| `addHyperlink(sheet, row, col, target, display, tooltip, location)` | `ws.setHyperlink(row, col, { target, location?, display?, tooltip? })` sets or replaces the cell's link and `null` removes it; a `null` target is an in-workbook link and requires `location`. the range the link covers is `range` in the record |
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
| `TableTotalsRowFunction` | `TableColumn.totalsRowFunction`, `TableColumnTotalsInput.totalsRowFunction` | `'none'`, `'sum'`, `'min'`, `'max'`, `'average'`, `'count'`, `'countNums'`, `'stdDev'`, `'var'`, `'custom'` | `None` 0, `Sum` 1, `Min` 2, `Max` 3, `Average` 4, `Count` 5, `CountNumbers` 6, `StdDev` 7, `Var` 8, `Custom` 9 |
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
| `''` | `null`: `DefinedName.comment`, `Hyperlink.location`, `Hyperlink.display`, `Hyperlink.tooltip`, `PivotCell.numberFormat`, `PivotDataField.numberFormat`, `PivotField.customName`, `PivotField.numberFormat`, `PivotOptions` captions and alternative text, `PivotWorksheetSource.name`, the password, algorithm, hash and salt fields of `SheetProtection` and `WorkbookProtection`, `SlicerProperties.styleName`, `SparklineGroup.dateAxisRange`, `TableStyleDefaults.tableStyleName` / `pivotStyleName` |
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
| `Table.update` | `range`, `name`, `header_row`, `totals_row` |
| `Table.update_style_options` | `style_name` and the style toggles |
| `Workbook.update_table_style_defaults` | `table_style_name`, `pivot_style_name` |
| `Workbook.update_iterative` | `enabled`, `max_iterations`, `max_change` |
| `Workbook.get_theme_script_font` | `scheme`, `script` |

`Chart.set_source`, `Table.set_column_totals`, `Sheet.set_selection`,
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
- **JavaScript, Python:** `removeMergeAt` and `removeHyperlinkAt` are removed
  in favour of `removeMerge` and `removeHyperlink`. `getDefinedName` and
  `removeDefinedName` take an optional `scope`, `getThemeScriptFont` and
  `setThemeScriptFont` take `{ scheme, script }`, `Chart.setType` takes
  `{ barDirection, grouping }` options, and `Pivot.addFieldItem` takes an
  `options` object.

### Behaviour changes under an unchanged name

- Formula text read back carries a leading `=` where Excel shows one:
  `getFormulaR1C1`, `PivotCacheField.formula`, table totals-row formulas,
  calculated-item formulas, conditional-format and validation `formula1` /
  `formula2`, formula thresholds of colour scales, data bars and icon sets,
  array-formula text and `getLambdaText`. C reads carry it too. Ranges (chart titles and series,
  print areas) stay without it. Input accepts text with or without the `=`,
  in C as well.
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
| `fm_workbook_table_set_column_totals` | `fm_sheet_set_table_column_totals` |
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
| `removeHyperlinkAt()` | removed |
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
| `setTableColumnTotals()` | `ws.table(...).setColumnTotals()` |
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
| `TableColumnTotalsInput.formula` | `TableColumnTotalsInput.totalsRowFormula` |
| `TableColumnTotalsInput.function` | `TableColumnTotalsInput.totalsRowFunction` |
| `TableColumnTotalsInput.label` | `TableColumnTotalsInput.totalsRowLabel` |
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
| `cols` | `columns` |
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
| `remove_hyperlink_at()` | removed |
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
| `table_set_column_totals()` | `ws.table(...).set_column_totals()` |
| `table_set_style_info()` | `ws.table(...).update_style_options()` |
| `table_style_count()` | removed; use `list_table_styles()` |
| `table_style_defaults()` | `get_table_style_defaults()` |
| `table_update()` | `ws.table(...).update()` |
| `Table` | `TableProperties` |
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
| (none) | `fm_sheet_get_print_options` and `fm_sheet_get_header_footer` read the print options and header/footer back |
| (none) | `fm_sheet_get_table_runtime_id` / `fm_workbook_resolve_table_runtime_id` and `fm_sheet_get_pivot_runtime_id` / `fm_sheet_resolve_pivot_runtime_id` map a table or pivot index to an id that survives other tables or pivots being added or removed, and back |
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
| `fm_table_column_info_t` | `totals_row_function` is an `int32_t` `fm_table_totals_row_function_t` instead of text; the totals setter's function argument is the same enum |
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
either date system; see [Added](#added). `fm_sheet_update_protected_range`
keeps the stored ranges when `ranges` is `NULL` and `range_count` is 0, which
lets a caller change the other fields of a range whose address the engine
cannot parse. The drawing functions for name, z-order, snapshot and restore
accept shapes, and the z-order functions also accept slicers.
`fm_sheet_add_slicer` and `fm_sheet_update_slicer` accept an ascending sort
order.

### Added

- `Sheet.insertCells(range, shift)` and `deleteCells(range, shift)` (Python
  `insert_cells`, `delete_cells`; C `fm_sheet_insert_cells`,
  `fm_sheet_delete_cells`) shift the cells of a band down or right (insert) or
  up or left (delete), with `CellShift` naming the direction. A reference moves
  only when its extent across the band lies wholly inside it; a range crossing
  the band and a whole-row or whole-column reference stay, and a range wholly
  deleted becomes `#REF!`. Conditional-format and validation ranges are split
  across the band, and a merge that only partly meets it is unmerged.
- Lookups and lists: `Workbook.sheetCount` and `hasSheet(ref)` (Python
  `len(wb)` and `ref in wb`), `formulon.listProfileIds()` (Python
  `list_profile_ids()`, 28 ids), `Workbook.listThemeScriptFonts(scheme)` with
  the `ThemeScriptFontKey` type for `getThemeScriptFont`, and
  `PivotCache.getRecordValue(recordIndex, fieldIndex)` (`get_record_value`).
  `Pivot.getField(i)`, `getDataField(i)` and `getFilter(i)` read one element
  of the list at that position (`IndexOutOfRange` past the end), and
  `Pivot.getProperties()` returns `{ name, cacheId, reportLayout, grandTotals,
  styleOptions }`, the definition of the pivot.
- `Sheet.updateValidation(ref, patch)` and `updateCfRule(ref, patch)` (Python
  keyword-only `update_validation`, `update_cf_rule`) change part of a rule.
  `ooxml.setFill`, `setBorder`, `setCellXf` and `setDxf` replace the record at an
  index, as `setFont` does.
- Reference types: `ValidationRef`, `SparklineGroupRef`, `PivotCacheRef` (a
  number or a `PivotCache`: `pivotCache`, `removePivotCache` and
  `PivotInput.cacheId` take it), `PivotFieldRef`, `ChartAxisRef` (an axis id, or
  `{ kind, plot? }`: `updateAxis(axis, patch)` finds the axis of that kind).
  Python exports the same aliases.
- `Sheet.getVisibility()` and `setVisibility(visibility)`, and the C
  functions of the same names.
- `DrawingAnchorInput`, `DrawingAnchorPatch` and `DrawingPlacementInput`
  (JavaScript and Python); `DrawingProperties.hidden` and `description` for every
  kind, written as the `hidden` and `descr` attributes of the drawing's
  `cNvPr`. `Drawing.name` reads the object's current name.
  `addSlicer(table, column, spec)` takes the column by name as well as by
  number.
- Chart parts are named and read-only: `ChartCategories`, `ChartAxisTitleText`,
  `ChartUnresolved` and `ChartSeriesRanges`.
- Conditional-format payloads: `CfColorScale`, `CfDataBar` and `CfIconSet` are
  the records a read returns, with `CfColorScaleInput`, `CfDataBarInput` and
  `CfIconSetInput` for input (they replace Python's `ColorScale`, `DataBar` and
  `IconSet`), and `CfExpressionRuleInput`,
  `CfCellIsRuleInput`, `CfColorScaleRuleInput`, `CfDataBarRuleInput`,
  `CfIconSetRuleInput`, `CfTop10RuleInput`, `CfAboveAverageRuleInput`,
  `CfTextRuleInput`, `CfTimePeriodRuleInput` and `CfPlainRuleInput` are the
  inputs of `addCfRule` (Python dataclasses of the same names, with
  `ConditionalFormatInput` the union).
- `EffectiveBorders` (`left`, `right`, `top`, `bottom`, `diagonal`) is the
  type of `EffectiveStyle.borders` in both bindings. `PrintOptionsInput`,
  `HeaderFooterInput` and their patches are exported, and
  `toNative` and `mergeFunctionMetadata` are named exports of every entry
  point. `getImageInfo`, `addPicture` and `restoreDrawing` accept the byte
  inputs `load` accepts. Python `Pivot` has a `repr`, `Workbook.ooxml` is typed,
  and iterating a `Sheet` or testing `in` raises `TypeError` that names
  `iter_rows()`.
- C: `fm_sheet_set_value`, `fm_get_profile_id_count`, `fm_get_profile_id_at`,
  `fm_sheet_get_sparkline_group` and `fm_sheet_get_validation`; see [C ABI
  changes](#c-abi-changes).
- Input types: `FontInput`, `FillInput` (`GradientFillInput`,
  `GradientStopInput`), `BorderInput` (`BorderSideInput`), `CellXfInput`,
  `DxfInput`, `CellStyleInput`, `TableStyleInput` (`TableStyleElementInput`),
  `CommentInput`, `PersonInput`, `ArrayFormulaInput`, `WorkbookProtectionInput`,
  `ProtectedRangeInput`, `SparklineInput`, `PrintTitlesInput`,
  `AutoFilterConditionInput`, `ChartFilterInput` and `DateSystemOptions`; read
  records `AutoFilterCondition`, `ChartSource`, `ChartFilter`, `CfValueObject`,
  `CfIcon`, `CfMatchBar` and `CfMatchIcon`. Python has the dataclass
  counterparts (`ChartInput`, `ChartSourceInput`, `ChartFilter`,
  `PivotWorksheetSourceInput`, `PivotDateGroupInput`, `SheetSelectionInput`,
  `TableColumnTotalsInput`, `ThemeColorsInput`, `WorkbookProtectionInput`,
  `ProtectedRangeInput`, `PrintTitlesInput`, `SparklineInput`,
  `SparklineGroupInput`, `AutoFilterInput`, `AutoFilterCondition`), each also
  accepting the record its getter returns.
- `ChartAxisPatch` gains `minorUnit`, `minorGridlines`, `position` and
  `sourceLinked`, `ChartLabelsPatch` gains `sourceLinked` and `deleted`, and
  `ChartDisplayPatch` gains `style` (1 to 48); `PivotField.showDetail` and
  `PivotFieldPatch.sortOrder`, `sortByField` and `showDetail` are new, as are
  the matching C fields.
- C: `fm_sheet_update_pivot_field` and `fm_sheet_update_pivot_field_item` merge
  an `fm_pivot_field_patch_t` / `fm_pivot_item_patch_t`;
  `fm_pivot_field_info_t` gains `show_detail`.

- A trailing argument with an obvious default can be left out:
  `getPrecedents(cell)` / `getDependents(cell)` follow one level,
  `recalcParallel()` picks the thread count itself (at most 8), `copySheet(index)` appends
  the copy, and `evaluateCfRange(range)` evaluates time-period rules against
  today (the pinned date when the workbook has one). Python takes the same
  defaults. An argument beyond the ones a method declares throws `TypeError`
  naming the method and how many arguments it takes.
- `fm_get_last_parse_error` (`FormulonError.parseError`,
  `FormulaSyntaxError.parse_error`) reports where a formula failed to parse:
  a `ParseErrorCode` (eight codes added for formula entry: argument count,
  reference required, qualified built-in, refused whitespace,
  parameter-shaped name, leading `==`, byte order mark and invalid structured
  reference), UTF-16 start and end offsets, line and column. The same error
  carries the parser statuses 1000 to 1014.
- `fm_sheet_get_tab_color` / `fm_sheet_set_tab_color` (`getTabColor` /
  `setTabColor`, `get_tab_color` / `set_tab_color`); `null` / `None` clears
  the colour.
- `getTypedValue`, `getTypedValues`, `getDate`, `getUsedRange`, `iterRows`,
  `evaluateTypedFormula`, `evaluateTypedFormulaArray`, `enumName` and the
  `ExcelError` class in JavaScript; `get_typed_value`, `get_typed_values`,
  `get_date`, `get_used_range`, `iter_rows`, `list_cells`, `set_values`,
  `evaluate_typed_formula` and `enum_name` in Python. A `Workbook` iterates
  its sheets, `Workbook.load` accepts an `ArrayBuffer` or typed array, and
  `getMemoryUsage` is on the WebAssembly package as on the native one.
- Enum tables for `CfRuleType`, `ComparisonOperator`, `DatePeriod`,
  `CfValueObjectType`, `IconSetName`, `ValidationType` and `ParseErrorCode`.
- Record fields that hold a cell or a range accept A1 text (`'B2'`,
  `'A1:C3'`) in JavaScript and Python, including `PivotInput.anchor`,
  validation and conditional-format ranges.
- `fm_get_image_info` reads picture properties from bytes without a workbook.
- `GROUPBY` and `PIVOTBY` accept `total_depth` beyond plus or minus 2 (nested
  subtotals for every key prefix), and `relative_to` follows the parent key
  prefix at any depth.

- Pictures and charts are written to `.xlsb` and read back as model drawings,
  so `listDrawings` and `ws.chart(id)` work after an XLSB round trip. A
  drawing holding other graphic frames (slicers, SmartArt) is not written and
  counts in `deferredFeatureCount`.
- `FM_VERSION_MAJOR`, `FM_VERSION_MINOR` and `FM_VERSION_PATCH` in
  `formulon_c.h` give the library version at compile time. The header opens
  with a quick start, the five pointer lifetime classes, the last-error
  contract and the rule that public structs gain no fields from 1.0 on.
- `fm_workbook_get_defined_name` and `fm_workbook_remove_defined_name` look up
  and remove a defined name by name and scope.

- A workbook file name: `fm_workbook_set_file_name` / `fm_workbook_get_file_name`
  (`setFileName` / `getFileName`, `set_file_name` / `get_file_name`). When it is
  set, every function that returns formula text, and `FORMULATEXT`, shows a
  reference to this workbook's own names as the formula bar does
  (`=Book.xlsx!Rate`, `='My Book.xlsx'!Rate`) instead of the stored
  `[0]!Rate`. The same spelling, and `[Book.xlsx]Sheet1!A1`, can be entered
  and refer to this workbook; no external link is created. The name must end
  in a workbook extension. The command-line tools set it from the input
  file.

- Date grouping works the way Excel's Group command does. Seconds, minutes,
  hours, days (with an interval), months, quarters and years each become a
  group field in the pivot cache, so a grouped workbook loads, evaluates and
  saves in Excel's own form. Labels and field names follow the 14 UI locales,
  years can follow the Japanese calendar, and ungrouping removes the trailing
  group fields.
- Pivot field and item labels that hold dates or numbers are drawn in the
  field's number format.
- Pivot report options: per field `defaultSubtotal`, `insertBlankRow`,
  `fillDownLabels` and `showDetail` (`Pivot.updateField` keys and
  `fm_pivot_field_patch_t` members), and the grand total caption and autofit flag in `fm_pivot_options_t`. The layout
  draws the grand total caption, blank rows, repeated labels and collapsed
  items as Excel does, and loaded files keep them.
- The `caption` member of `fm_pivot_item_patch_t` renames an item without
  changing the value it stands for, and `show_detail` expands or collapses one
  item (`Pivot.updateFieldItem` keys `caption` and `showDetail`). `fm_pivot_item_info_t` reports both. A
  caption is at most 255 UTF-16 units and may not repeat another item's label
  in the field.
- The Values level ("Σ Values") appears in the row and column field orders as
  `FM_PIVOT_VALUES_FIELD` (-2 in JavaScript and Python), at its position
  among the fields. Listing it in `fm_sheet_set_pivot_row_field_order` or
  `_set_column_field_order` moves it to that axis. It needs two or more data
  fields, and the axis that holds it must keep listing it.
- `fm_sheet_show_pivot_details` (`Pivot.showDetails`, `Pivot.show_details`)
  runs Show Details on a data, subtotal or grand total cell: it adds a sheet
  with the matching source rows in a table, before the pivot's sheet, and
  makes it the active sheet. Sheets from the pivot's sheet onward move up by
  one.
- `Drawing.update({ name })` renames a picture, chart or shape. Default
  object ids now count per sheet from 2 and are shared by pictures, charts,
  shapes and slicers, and default chart, picture and table names use the
  workbook locale's word.

- `fm_workbook_get_pivot_cache_field_shared_item_value`
  (`PivotCache.listFieldSharedItems`, `PivotCache.list_field_shared_items`)
  reads the shared items of a pivot cache field, so a refresh can keep the
  existing item order.
- Pivot cache dates: `fm_workbook_add_pivot_cache_field_shared_item_date`
  and `fm_workbook_set_pivot_cache_record_date` store a serial as a date
  (`<d>`) rather than a number, and
  `fm_workbook_is_pivot_cache_field_shared_item_date` reads the mark back. In
  JavaScript and Python a `Date` (`date` or `datetime`) passed to
  `PivotCache.addFieldSharedItem` or `setRecordValue` is a date item, and
  `listFieldSharedItems` reports the mark as `isDate` (`is_date`). A date outside the workbook's
  date system is refused. Date filters, date grouping and the saved field
  hints treat these values as they treat dates loaded from a file.
- `fm_sheet_set_row_auto_height` (`setRowAutoHeight`, `set_row_auto_height`)
  stores the automatic height of the rows `first` to `last` of a sheet
  without marking them custom, so a refreshed height is saved as `ht` alone
  and the rows keep following their content and font.
- `fm_sheet_get_range_xf_bases_excluding` (`getRangeXfBasesExcluding`,
  `get_range_xf_bases_excluding`) lists the XFs a selection resolves to while
  leaving out given cells, for hosts that resolve those cells' style
  themselves.
- `fm_sheet_get_array_formula` / `fm_sheet_set_array_formula`
  (`getArrayFormula` / `setArrayFormula`, `get_array_formula` /
  `set_array_formula`) read and enter legacy (Ctrl+Shift+Enter) array
  formulas. The getter reports the block's range and formula from any of its
  cells. The setter refuses a rectangle that partly overlaps another block or
  intersects a merge, a data table or a PivotTable.

- Calculated fields: `fm_workbook_add_pivot_cache_calculated_field`,
  `_calculated_field_set_formula`, `_calculated_field_remove` and
  `fm_workbook_get_pivot_cache_field_formula` (`PivotCache.addCalculatedField`,
  `PivotCache.add_calculated_field`, and so on). A calculated field shows its
  formula over the Sums of the fields it names, totals included, and is saved
  in Excel's form.
- Calculated items: `fm_workbook_add_pivot_cache_calculated_item`, `_count`,
  `_at` (`fm_pivot_calculated_item_info_t`), `_set_formula`, `_move` and `_remove`
  (`PivotCache.addCalculatedItem`, `PivotCache.add_calculated_item`, and so
  on). Items are evaluated in the context of each cell, the field's totals
  include them, and where items meet the one later in solve order wins. A data
  field cannot use Average, StdDev, StdDevP, Var or VarP beside them.
- PivotTable options in one struct: `fm_sheet_get_pivot_options` /
  `fm_sheet_update_pivot_options` with `fm_pivot_options_t`
  (`Pivot.getOptions` / `Pivot.updateOptions`, `Pivot.get_options` /
  `Pivot.update_options`): error and empty-cell captions, field headers, classic
  grid layout, empty rows and columns, report filter wrapping, merged labels,
  compact indent, print titles, alternative text and the rest of the options
  dialog. The layout draws those that change the report.
- Pivot cache options: `fm_workbook_get_pivot_cache_options` /
  `fm_workbook_update_pivot_cache_options` with `fm_pivot_cache_options_t` (`PivotCache.getOptions`,
  `PivotCache.get_options`): saving the source records, the missing-items
  limit and refresh on open. A cache saved without records writes no records
  part.
- The `customName` key of `Pivot.updateField` (`fm_pivot_field_patch_t`)
  renames a pivot field in the report, and the `showAll` key lists its items
  without data.
- `fm_sheet_get_pivot_data_field_default_name` (`Pivot.getDataFieldDefaultName`,
  `Pivot.get_data_field_default_name`) reads the name a new data field would get,
  without adding one.
- `fm_sheet_project_pivot` (`Pivot.project`) writes a
  pivot's report into its cells now, and `fm_sheet_is_pivot_projection_pending`
  (`Pivot.isProjectionPending`, `Pivot.is_projection_pending`) tells whether a cache
  change left them stale.
- Charts read and write "Show #N/A as an empty cell" (`c16r3:dispNaAsBlank`):
  `na_as_blank` in `fm_chart_info_t` (`naAsBlank`, `na_as_blank`), the
  `na_as_blank` field of the display patch, and a preview override. Turning it
  off removes the extension, as Excel writes it.
- `fm_sheet_get_spill` (`getSpill`, `get_spill`) reports a spill that is
  blocked with `#SPILL!`: at the anchor, `blocked` is 1 and the anchor, rows
  and columns give the range the spill tried to fill, so a host can outline
  it as Excel does.
- `fm_workbook_parse_reference` (`parseReference`, `parse_reference`) reads a
  reference string with the grammar `INDIRECT` uses under the workbook's
  profile, A1 or R1C1, into a sheet name and a 0-based rectangle;
  `fm_workbook_format_reference` (`formatReference`, `format_reference`)
  writes a rectangle back as A1 text, quoting the sheet name as a formula
  does.

- `fm_sheet_rename_table_column` (`Table.renameColumn`,
  `Table.rename_column`) renames a table column as typing over its header
  does, and returns the final name. An empty name takes the locale's word
  with the lowest unused number (`Column1`, `列1`, `Spalte1`), and a name
  another column already has gets the lowest unused number from 2 appended
  (`Amount2` becomes `Amount22`). Structured references to the column follow
  in cell formulas on every sheet, totals-row formulas, defined names and
  chart series. A column added inside a table is named with the same locale
  word.
- `Shape.setZOrder` and `Shape.snapshot`, and `Slicer.setZOrder` (Python
  `set_z_order`, `snapshot` and `Slicer.set_z_order`). A snapshot restores the shape, picture fill and
  hyperlink included. A slicer inside an
  `mc:AlternateContent` block can be moved in z-order; a slicer still cannot
  be removed, anchored or renamed through them.
- `Formulon.convertCivilToSerial` and `Formulon.convertSerialToCivil`
  (Python `convert_civil_to_serial`, `convert_serial_to_civil`; C
  `fm_convert_civil_to_serial`, `fm_convert_serial_to_civil`) convert between
  a calendar date and time and a serial number for the 1900 or 1904 date
  system. Civil input is checked, not normalised: month 13 is
  `InvalidArgument`. In the 1900 system `1900-02-29` is accepted and is
  serial 60, as in Excel; the 1904 system refuses it and refuses years before
  1904. A serial in `[60, 61)` converts back to `1900-02-29` in C and
  JavaScript, and Python raises `ValueError` because `datetime` cannot hold
  it. A fractional second rounds to the nearest second, and a serial outside
  `[1, max)` (`[0, max)` for 1904) is `InvalidArgument`. Python takes a
  `date` or a naive `datetime`, and a `datetime` with a time zone raises
  `ValueError`.
- `FormulonOptions.wasm` takes the module as bytes (`BufferSource`) or a
  compiled `WebAssembly.Module`, for hosts that cannot fetch it by URL. The
  native package ignores it.
- `Workbook.dispose()` deletes the raw handle at once. A WASM workbook that is
  dropped without `dispose()` is released when it is garbage collected, on a
  best-effort basis: do not rely on it for timing, and call `dispose()` (or
  `using`) when memory matters.
- `make npm-typecheck` type-checks the README examples and the public
  declarations with TypeScript.

- `Workbook.getDefaultFont()` beside `setDefaultFont`, `ws.updateSheetProtection(patch)`,
  `getStyleOptions()` / `updateStyleOptions(patch)` on `Table` and `Pivot`,
  `Table.rename(name)`, `listRowBreaks()` / `listColumnBreaks()` with
  `RowBreak` and `ColumnBreak` records, `Pivot.updateFieldItem`,
  `activeSheet()`, and `ChartExtendedLayout` with `extendedLayout` /
  `isExtended` on chart plots. C adds `fm_tristate_t`, `fm_row_break_t`,
  `fm_column_break_t`, `fm_table_style_options_t`,
  `fm_pivot_style_options_t` and `fm_sheet_add_pivot_field_item_by_cache_index`.
- `getCellFormat` reports `source`, and a conditional-format input accepts
  `format: CellFormatPatch`, which is resolved to a differential format;
  `dxfIndex` stays as the low-level alternative.

### Changed

- Chart fill, line and data-label font colours and shape fill and outline
  colours accept and return theme references (`ColorKind.Theme`, `theme` 0..11
  numbered as in a cell colour, `tint` in -1..1). A scheme colour whose only
  adjustments are luminance and alpha reads as Theme; the tint is stored as
  `lumMod` / `lumOff` and quantised to 1e-5. Indexed and automatic colours are
  rejected with `FM_STATUS_INVALID_ARGUMENT`. A colour with any other
  transform reads as before: its rendered RGB on a chart, `Other` on a shape.
- Changing a chart between bar, column and area types keeps each series' fill
  as it is, theme colours with their transforms, gradients and patterns
  included, instead of only an RGB colour. A change to or from a line or
  scatter type re-derives the series colour, an RGB one included, as Excel
  does; previously an RGB colour was carried to the scatter markers.
- A chart without chart style parts, as `addChart` creates and as many
  non-Excel writers save, keeps its series and point formatting through a
  type change, pie and doughnut included, as Excel does for such a chart: a
  line plot drops every line when the series line is hidden and shows no
  markers, and a scatter plot hides every line (2 pt wide when the series
  line, and for a point its own line too, was visible with a width, 3 pt
  otherwise) and keeps only shown markers. Previously such a chart was
  restyled as a chart with style parts is.
- A chart type change between category and value axes (scatter to or from
  column, bar, line or area) converts the existing axes in place, as Excel
  does, instead of replacing them with new default axes: each axis keeps its
  id, direction, gridlines, title, ticks and formatting, and a kept value axis
  its bounds and units. A category axis's label and mark intervals become the
  major and minor units of the value axis; an axis turned into a category
  axis drops its bounds and units (its log base too on a bar plot) and is not
  multi-level. A converted axis links its number format to the category data
  (General for text). A date axis stays as it is over a column, bar, line or
  area plot and becomes a value axis, keeping its units and dropping its time
  units and bounds, over a scatter plot; a scatter plot whose X values are
  dates becomes a plot over a date axis. Every type change drops a category axis's label and
  mark intervals, as Excel does.
- A chart type change to or from pie or doughnut follows Excel's: pie points
  become points in the series colour that keep their outline; a series line
  reaches the slices less what the old type's style draws by default; new
  axes have outside tick marks and the default crossing (a bar chart lists
  its value axis first, and a scatter chart's axes have no outline and its
  horizontal axis no gridlines or number format); and a new bar plot takes a
  gap width of 100.
- A chart added over categories that are all dates (blanks aside) gets a
  date axis, as Excel makes one; a single plain number or text among them
  keeps a category axis. A date axis's base unit, whether made by adding a
  chart or by a type change, follows the smallest gap between the sorted
  dates: days under 28 days, months under 365, years otherwise. Each axis of
  a new chart links to the number format of its data's first cell (the
  category axis is General when a category is text, and a percent-stacked
  value axis stays 0%) instead of always General.
- Changing a chart's data range re-links its axes as Excel does: an automatic
  category axis becomes a date axis over date categories and turns back over
  text, a date axis already in place keeps its base unit, a manual axis keeps
  its kind, and linked axis number formats follow the new first cells.
- A shape fill or outline colour is no longer dropped silently: given without
  its kind it fails with a `TypeError` in JavaScript and Python, and given with
  a kind other than solid it fails with `FM_STATUS_INVALID_ARGUMENT`.
- Number-format codes (`fm_workbook_add_number_format`, `fm_workbook_get_number_format`
  and their JavaScript and Python forms) use the vocabulary of the workbook's
  Excel profile, as the pivot `number_format` field does. Built-in id 14 is
  `m/d/yy` under en-US and `yyyy/m/d` under ja-JP, so adding that code
  returns 14, and the ECMA-376 spelling `mm-dd-yy` is stored as a custom
  format. `get(add(code))` returns `code`.
- `evaluate_cf_range` with `today` omitted evaluates time-period rules against
  the pinned clock, else the local date, instead of disabling them, and the
  JavaScript `evaluateCfRange(range, { today })` takes the date as an option
  that may be left out.
- `fm_get_function_metadata` (`getFunctionMetadata`, `function_metadata`) reports the
  argument counts formula entry accepts for every built-in, including the
  functions it used to report as `0` with no upper bound (`IF` 2 to 3,
  `XLOOKUP` 3 to 6, `SUMIFS` 3 or more, `LET` 3 to 253, `LAMBDA` 1 to 254).
- A conditional formatting rule added without an id gets one in the form
  Excel uses, `{00000000-000E-0000-<sheet>-0000<priority>}`, in place of the
  `{FC000000-…}` form.
- `fm_sheet_get_cell_at`, `fm_workbook_get_defined_name_at` and the validation,
  table and pivot field formula views are documented as valid until the next
  read that uses the handle's scratch buffer, like the other text getters. A
  view now points into scratch storage when a file name is set and the text
  needs the workbook name spelled out.

- A pivot table's report is kept in the sheet cells, as Excel keeps it. Every
  pivot table edit writes the report into its range with Excel's cell formats,
  and formulas that read the range see it; deleting a pivot table clears its
  range. An edit or a new pivot table whose report would overlap another
  pivot table, a table, a data table, an array formula or a merge now fails
  with `FM_STATUS_PRECONDITION_FAILED` and changes nothing. A cache edit leaves the
  cells stale until the pivot's next edit, `fm_sheet_project_pivot` or a
  save, which writes them first.
- A data field named by the engine takes a number after the name, as Excel
  does, when the name is already taken by another data field or a field
  ("Sum of Amt2").
- A pivot table created through the API starts with the options Excel gives a
  new pivot: item print titles on, compact indent 0 and one filter per field.
- Loaded pivot filters of every type are listed by `fm_sheet_get_pivot_filter_at`
  and removed by `fm_sheet_clear_pivot_filters`.

- Functions that return formula text spell structured references as the
  formula bar does: `[#This Row]` is written `@`, a reference from inside the
  table to one of its own columns drops the table name (`=[@Qty]`,
  `=SUM([Qty])`; `=tblSales[@Qty]` from outside), and `tblSales[]` reads
  `tblSales`. `FORMULATEXT` shows the same form at the target cell, with
  `#All`, `#Data`, `#Headers` and `#Totals` and the item separator inside the
  brackets in the workbook's locale (`[[#Kopfzeilen];[Qty]]`). Formula
  setters accept this form and store Excel's long form; a bare table name is
  stored as `tblSales[]`.
- Writing a value into a table's header cell (`fm_sheet_set_text`,
  `_set_number`, `_set_bool`, `_set_error`, `_set_blank`) renames the column
  as `fm_sheet_rename_table_column` does. Clearing it gives the default
  name instead of an empty header.
- `Table.update` takes a `TablePatch` and moves the table with `range`, given
  as A1 text or a `SheetRect` on the table's sheet. `getCellRectPt` and
  `getWidthModel` (Python `get_cell_rect_pt`, `get_width_model`) use the
  Display geometry mode when none is given.
- `fm_sheet_update_protected_range` leaves the stored ranges alone when
  `ranges` is `NULL` and `range_count` is 0, so a read-modify-write does not
  lose a range the engine cannot parse.

- The record-taking C functions accept the new record layouts and
  `size_t` indexes; each line below is a break for a stale caller.
  `fm_sheet_add_hyperlink`, `fm_sheet_add_validation`,
  `fm_sheet_add_merge`, `fm_sheet_remove_merge` and `fm_workbook_add_person`
  take their records as `const` pointers and under renamed types
  (`fm_hyperlink_t`, `fm_data_validation_t`, `fm_cell_rect_t`,
  `fm_person_t`); `fm_sheet_add_threaded_comment` takes `fm_threaded_comment_t`.
  `fm_sheet_get_auto_filter` and `fm_sheet_set_auto_filter` use
  `fm_auto_filter_t`, `fm_sheet_get_effective_style` uses
  `fm_effective_style_info_t`, `fm_sheet_get_format_defaults` uses
  `fm_sheet_format_defaults_t`,
  `fm_sheet_validate_value` uses `fm_validation_outcome_info_t`,
  `fm_workbook_recalc_parallel` uses `fm_parallel_recalc_stats_info_t`,
  `fm_workbook_get_theme_colors` / `fm_workbook_set_theme_colors` use
  `fm_theme_colors_t`, and `fm_workbook_get_theme_fonts` /
  `fm_workbook_set_theme_fonts` use `fm_theme_fonts_t`.
  `fm_workbook_resolve_color` takes a `const fm_color_t*` in place of an
  `fm_color_spec` by value.
- The sheet-scoped list and edit functions take `size_t sheet_index` and
  `size_t` positions: `fm_sheet_clear_hyperlinks`, `fm_sheet_clear_merges`,
  `fm_sheet_clear_validations`, `fm_sheet_get_hyperlink_at`,
  `fm_sheet_get_hyperlink_count`, `fm_sheet_get_merge_at`,
  `fm_sheet_get_merge_count`, `fm_sheet_get_validation_count`,
  `fm_sheet_remove_hyperlink`, `fm_sheet_remove_hyperlink_at`,
  `fm_sheet_remove_merge_at`, and the sheet operations
  `fm_workbook_move_sheet`, `fm_workbook_remove_sheet` and
  `fm_workbook_rename_sheet`. The count getters write through
  `size_t* out_count`.
- `fm_sheet_get_picture` takes `int32_t* out_present` before its outputs. A
  cell with no picture is `FM_STATUS_OK` with `*out_present = 0`; it was
  `FM_STATUS_INVALID_ARGUMENT`.
- `fm_sheet_get_print_area` and `fm_sheet_set_print_area` read and take
  `fm_cell_rect_t` ranges with a count instead of A1 text.
- Python: the package runs one WebAssembly instance per process, so calls are
  serialised and memory it grew is not returned. The Python README lists the
  limits.
- `GROUPBY` and `PIVOTBY` with `total_depth` of 2 or more over a single key
  column return `#VALUE!`. `TEXTBEFORE` and `TEXTAFTER` return `#N/A` on a
  miss when `if_not_found` is omitted (they returned `0`). A union that holds
  a 3-D reference is `#NAME?`, and a union of areas on different sheets is
  `#VALUE!`. A list validation whose source is a range compares the proposed
  number exactly, so `0.1+0.2` matches its own list entry.

### Removed

- `fm_cf_color_t` and the `*_argb` fields of the chart and shape records
  (see `fm_color_t` under BREAKING), `CfColor`, the `*Argb` keys of
  JavaScript colour patches, `fillRgb` / `lineRgb`, and `EvalArrayResult`.
- `kCliInvalidArgs`, `kCliFileNotFound` and `kCliOutputFailed` (7005 to 7007).
- `fm_workbook_evaluate_formula_array_cell`: `fm_sheet_evaluate_formula_array`
  returns an owned `fm_value_array_t` read with `fm_value_array_get_value_at`.
- `fm_sheet_set_tab_hidden`, `fm_workbook_set_defined_name_scoped`,
  `fm_workbook_set_iterative_enabled`, `fm_sheet_set_fit_to_page` and
  `fm_sheet_set_thread_resolved`; the replacements are in the table under
  [C ABI changes](#c-abi-changes).
- `fm_workbook_probe_image`: use `fm_get_image_info`, which takes no workbook.
- The former names of the renamed C functions (the `removed` rows of the
  migration table), with no aliases.
- The singleton `remove*` / `clear*` methods (`removeAutoFilter`,
  `clearPinnedNow`, `removeSheetRowBreak`, `removeSheetColBreak`, and the
  table AutoFilter equivalent), replaced by `set<X>(null)`; `addHyperlink`,
  `addHyperlinkRange`, `addSheetRowBreak` and `addSheetColBreak`, replaced by
  `setHyperlink`, `setRowBreak` and `setColumnBreak`;
  `pivotFieldAddItemAt`, replaced by `addFieldItem(field, { cacheIndex })`;
  the `enabled` field of sheet protection; `DrawingObject.mediaPath`, whose
  role `format` covers; and the `fm_sheet_remove_auto_filter`,
  `fm_sheet_remove_row_break`, `fm_sheet_remove_col_break` and
  `fm_workbook_clear_pinned_now` C functions.
- `fm_workbook_evaluate_formula`, which returned the top-left cell of a
  result: use `fm_sheet_evaluate_formula_array`. `FilterOperator`,
  `ChartBlanksAs`, `SparklineEmptyCells` and `PivotFilterValueKind` in the
  bindings, with the C typedefs they mirrored where the tables merged (see
  [Enums](#enums)); `Value.to_python()`; the `valueKind`, `valueInt`,
  `valueDouble`, `valueText` and `valueHigh*` keys of a pivot filter; the `todaySerial` argument of `evaluateCfRange`.

### Performance

- The threads WebAssembly build skips Binaryen's memory packing, which trades
  raw module size for a smaller Brotli size.
- Both WebAssembly builds are smaller, from shared argument readers in the
  JavaScript binding layer and shared helpers in the engine. Behaviour is
  unchanged.

### Fixed

- Saving a new workbook with an inserted image or chart as `.xlsb` no longer
  fails with `IoContentTypeInvalid`.
- `PivotField.subtotalTop` reads `false` after it is set to `false` in
  JavaScript. It used to read `true` whatever was set.
- A reference deleted by a structural edit is spelled as Excel spells it. A
  row or column deletion keeps the qualifier the reference was written with
  (`=Sheet3!#REF!*2`, `='Data 1'!#REF!`, a defined name `Sheet3!#REF!`), and
  that qualifier follows a later sheet rename. A sheet deletion keeps the
  address in the formula bar (`=#REF!A1`, `=SUM(#REF!A:A)`, `#REF!#REF!` for
  an already-deleted reference) while the saved formula stays bare `#REF!`;
  copying a sheet that one end of a 3-D reference names spells it the same
  way. Removing an end sheet of a 3-D span that leaves one sheet gives the
  single-sheet reference (`=SUM(Sheet2!A1:B2)`). The parser accepts all of
  these spellings, including `#REF!A:A` and `#REF!1:3`.
- Inserting or deleting rows or columns recalculates formulas that read a
  whole column, a whole row or a large range across the edited band, such as
  `=SUM(Sheet3!A:A)` after a row of Sheet3 is deleted; they kept their old
  value before.
- A spill reference whose anchor a structural edit deletes keeps its `#`, as
  in Excel: `=SUM(Sheet3!#REF!#)` and `=SUM(#REF!#)` after a row deletion,
  `=SUM(#REF!D5#)` after a sheet deletion, saved as
  `_xlfn.ANCHORARRAY(#REF!)`. It evaluates to `#REF!` (it gave `#VALUE!`),
  and `#REF!#` can be typed and read back.
- Deleting a sheet keeps its sheet-scoped names that other sheets still use:
  a name referenced as `Sheet2!Rate` becomes workbook-scoped and the
  reference becomes `[0]!Rate`, or binds to an existing workbook-scoped name
  of the same name. Names nothing else references are dropped as before.
  A reference from another surviving name, from a chart series or from a
  LAMBDA call (`=Sheet2!Fn(3)` becomes `=[0]!Fn(3)`, not `#REF!`) keeps the
  name too, and a kept name keeps the names its body reaches through
  `Sheet2!Name`. An unqualified name in its body does not keep that name.
- Deleting a sheet sends structured references to the tables on it to
  `#REF!` (`=SUM(#REF!)`); they kept naming the deleted table before.
- Conditional formatting formulas read from a file drop a qualifier naming
  their own sheet (`Sheet3!$A$1>0` on Sheet3 is `$A$1>0`), as Excel does on
  load, so a later row deletion gives `#REF!>0`. The same applies to color
  scale, data bar and icon set thresholds, to data validation formulas, and
  to a 3-D reference whose both ends are the own sheet
  (`Sheet3:Sheet3!$A$1`). A sheet-scoped name keeps its qualifier.
- Conditional formatting rules and data validations that reference another
  sheet are read from the x14 extension of an Excel file (`.xlsx`) and from
  its 2010 records (`.xlsb`). Before, they were missing from the model: they
  were not evaluated, not listed, and not updated by structural edits. They
  are written back the same way, so a rule that references another sheet no
  longer goes into the base element, and it stays in the extension after
  the sheet it references is deleted.
- In `.xlsb` conditional formatting and data validation formulas, a relative
  reference to another sheet is stored as an offset from the rule's first
  cell, as Excel stores it. Excel read the old form as a different column
  (`Sheet1!A:A` on G1:G3 came back as `Sheet1!G:G`).
- XLSB stores a qualified or deleted-sheet `#REF!` the way Excel does (a
  sheet-bound reference error, or the address through a deleted-sheet
  entry), and reads those forms back with their sheet or address: an Excel
  `.xlsb` shows `=Sheet3!#REF!*2` and `=#REF!A1` again. A reference to a
  deleted sheet in an Excel `.xlsb` no longer reads back as a plain `A1` on
  the formula's own sheet.
- PivotTable cache dates are read and written as Excel spells them: a serial
  counts days from 1899-12-30 in the 1900 date system (so serials 1 to 60
  name the day before their cell display) and from 1904-01-01 in the 1904
  system. Dates before 1900-03-01 and every date in a 1904 workbook were off
  before.
- A PivotTable cache built in memory saves the `<sharedItems>` content hints
  Excel writes for its values, so Excel opens it without a repair prompt:
  date fields get `containsDate` and `minDate` / `maxDate` (omitted when dates
  mix with numbers), number fields mixed with text get `minValue` /
  `maxValue`, and numeric fields off every axis get their full numeric hints
  in place of a bare `containsNumber`.
- A PivotTable label filter on numeric items orders numbers before a text
  bound, as Excel does: `less than "abc"` keeps every number, and `not
  between "2" and "abc"` keeps only the items below 2.
- A PivotTable with column fields but no row or data fields starts at its
  anchor column and is two rows tall. Outline and tabular forms caption the
  column area with the field name.
- With the classic drag-and-drop layout (`gridDropZones`) and no row fields,
  a PivotTable keeps its row-label column. It labels the value row with the
  locale's total word, puts a lone data field's name in the top-left corner,
  and writes the matching `firstDataCol` / `firstHeaderRow` offsets.
- A PivotTable read from a file and then edited (a filter, a field, a cache
  change) saves the range of its new layout. Before, it kept the range the file
  described whenever the report got smaller.
- `INDIRECT` reads an unquoted sheet name with non-ASCII characters or a dot
  (`INDIRECT("売上!A1")`, `INDIRECT("Sheet.1!A1")`) instead of returning
  `#REF!`, following the rule formulas use for sheet names.
- Writing a number, text or blank into a pivot cache record, or clearing a
  field's shared items, no longer leaves the cell saved as a date or as a
  shared-item index.
- A data bar without an x14 extension (the Excel 2007 form) is drawn as Excel
  draws it: no axis, and every bar, negative ones included, grows from the
  left edge in proportion to the minimum-to-maximum span. A data bar with an
  automatic axis, the default for one added through the API, now writes its
  x14 extension, so Excel keeps the axis after a save instead of showing the
  2007 form.
- A structured reference to one cell (`[@Qty]`, or `[#Headers]` or
  `[#Totals]` with one column) in a legacy formula loaded without a
  dynamic-array mark no longer gains an implicit-intersection `@`.
- Formulas with structured references in a loaded workbook recalculate when
  the table changes; totals, `SUBTOTAL`, `[#Totals]` and `[#Headers]` results
  kept their loaded values before. A `[#This Row]` reference now depends on
  its own row.
- A column name containing `,` or `:` is read as one name in a structured
  reference (`t[a,b]`, `t[a:b]`), as Excel stores it, and `'` escapes are
  read the same way in formulas, files and table edits.
- Formula entry refuses an argument Excel refuses for a parameter that takes
  only a reference: an empty or constant argument, an operator, or a function
  that never returns a reference, as in `=SUBTOTAL(9,)`, `=SUMIF(,1)` or
  `=COUNTIF(1,1)`. The parameters are those of `SUMIF`, `COUNTIF`, the
  `*IFS` and database functions, `SUBTOTAL`, `OFFSET`, `CELL`, `RANK`, `ROW`,
  `COLUMN`, `AREAS`, `COUNTBLANK`, `FORMULATEXT`, `ISFORMULA` and
  `PHONETIC`. Empty value arguments such as `=SUM(,)` are still accepted.
- Formula entry refuses an argument count Excel refuses: too few or too many
  arguments (`=ROWS()`, `=ABS(1,2)`, `=SUM()`, `=IF(1)`), an incomplete pair
  for the functions that take arguments in pairs (`=COUNTIFS(A1,1,A1)`,
  `=SUMIFS(A1,A1,1,A1)`, `IFS`, `SORTBY`, `GETPIVOTDATA`), and more than 255
  arguments to any function. An omitted argument counts, so `=ABS(1,)` is
  refused. A workbook that already holds such a formula still loads and
  calculates it.
- `LARGE`, `SMALL`, `PERCENTILE`, `QUARTILE` (and their `.INC` / `.EXC`
  forms), `TRIMMEAN` and `FVSCHEDULE` take exactly two arguments, and
  `TRANSLATE` takes one to three, as in Excel.
- `GROUPBY` and `PIVOTBY` call a function passed by name with the group values
  and the whole value column when its minimum is two arguments, so `XIRR` and
  `FORECAST.ETS.SEASONALITY` give Excel's results.
- Adding or updating a slicer with an ascending sort order (`fm_sheet_add_slicer`,
  `fm_sheet_update_slicer`, `addSlicer`, `Slicer.update`) failed with
  `InvalidArgument`; it is accepted and stored.
- A `GROUPBY` / `PIVOTBY` function that returns an array, of any shape and
  one cell included, gives `#VALUE!` in that value cell as Excel does (it was
  `#CALC!`); one that returns a `LAMBDA` stays `#CALC!`. `IF` passed by name is
  `#VALUE!` in every value cell.
- A scalar function lifted over an array argument returns an array even when
  that array is one cell (`=TYPE(ABS({1}))`, `=TYPE(ROUND(1.5,{0}))` and
  `=TYPE(ISNUMBER({1}))` are 64, through `LET` and `LAMBDA` too), so
  `=GROUPBY({"a"},{1},ABS,,0)` is `#VALUE!`. A one-cell reference reads as that
  cell: a one-cell spill (`=TYPE(C1#)` is 1), a name over one cell, and `TAKE`
  / `DROP` over a range. `ISREF` is `TRUE` for a spill reference and for `TAKE`
  / `DROP` over a range.
- `formulon paginate` exits with the library status when reading page breaks
  fails, as its other failures do.
- Text comparison, sorting and lookups under the Mac ja-JP profile treat a
  prolonged sound mark as Excel does: after a kana it keeps that kana's width,
  and after a half-width voiced mark it counts as half-width. A voiced or
  semi-voiced combining mark composes with its base kana on the Mac profiles.
- A threaded comment's `created` stamp is saved in UTC. A trailing `Z` in a
  file is accepted, and a stamp with a UTC offset or an unparseable stamp reads
  as absent (`null`) on load; the attribute text used to be kept as it was
  read.
- Engine error messages a call reports lead with the parameter or field they
  are about (`name is not a valid defined name`, `shift must be down or
  right`), or are a sentence when no argument is at fault (`Cannot change part
  of an array`). They used to lead with a prose word or an operation prefix
  (`insert_cells: ...`), which the bindings turned into a path that does not
  exist (`Workbook.setDefinedName definedName.invalid name`) or showed in its
  C spelling. Every engine error now carries the call's label
  (`Sheet.setValue Cannot change part of an array`), whatever its status, and
  a leading C path becomes the public path only when it names one of the call's
  arguments. File diagnostics keep their `<part>: <problem>` form.
- An out-of-range index reads alike in every binding:
  `Workbook.sheet key 5 is out of range (1 sheet)`, with the noun singular for
  a count of one. Python labels its sheet lookups as JavaScript does.
- Python: a number parameter given a string, `None` or a `bool` raises
  `TypeError` (`Sheet.set_number value must be a number`) instead of
  converting the string or failing with an unlabelled error.
- Python: page breaks, validation flags, the protection select flags, a
  colour-scale threshold's `gte`, table-style `pivot` / `table` and a pivot
  field's `subtotal_top` read `False` when the stored value is false; they used
  to read `True`.
- Every cell of a legacy array formula reads the block's formula from
  `getFormula`, `listCells` and `getCellsInRange`, as Excel's `Range.Formula`
  does; only the top-left cell did. `getSpill` returns `null` on every cell of
  a legacy array block, which is not a spill. A spill child still reads no
  formula.

## [0.18.0] - 2026-10-26

### BREAKING

- **C ABI:** `fm_sheet_view_t` grows by `split_x` / `split_y`,
  `show_formulas` and `show_outline_symbols` (64 to 88 bytes on 64-bit
  targets), `fm_font_record`
  by the `outline` / `shadow` / `condense` / `extend` toggles and their `has_*`
  flags (88 to 120 bytes), and `fm_fill_record` by the gradient members (64 to
  128 bytes); `fm_dxf_record` grows with them. `fm_cf_rule_t` gains
  `icon_set_icons` / `icon_set_icon_count` at its end, and
  `fm_pivot_field_info_t` gains `default_subtotal` / `subtotal_fns` at its end.
  `fm_sheet_format_defaults` takes `custom_height` / `zero_height` ahead of the
  `effective_*` members, which move (48 to 56 bytes). `fm_column_layout_t` and
  `fm_row_layout_t` keep their size and take `collapsed` from padding. A caller
  compiled against an earlier header must be recompiled.
- **C ABI:** text arguments that are not well-formed UTF-8 (surrogate code
  points, overlong forms, truncated sequences) are rejected with
  `kBindingUtf8EncodingError` (7002), and index, count and outline-level
  arguments are range-checked before anything changes (an outline level
  above 7, an `xf_index` outside the cell-XF table, hiding or removing the
  only visible sheet). Every binding reports the same statuses.
- JS packages: both packages check every positional argument the same way
  before any engine call, as the documented `Status` contract states. A
  missing required argument, a value of the wrong type, or a primitive or
  `null` where an object is declared throws `TypeError`; the native package
  no longer substitutes `0`, `""`, `false` or an empty object (for example
  `setNumber(0, 0, 0)`, `setIterative(true)` and `addFont(null)`). Malformed
  arguments a method reports as a status return 2 (`kInvalidArgument`) on
  both packages; the native package returned 7001.
- Python: `Value` validates its payload on construction. A payload the kind
  does not take, or a missing one, raises `ValueError`; a payload of the
  wrong type raises `TypeError`. Packed integer arguments are range-checked
  before the C call.
- Sheet names: a name that begins or ends with an apostrophe is refused on
  add and rename (`kInvalidSheetName`), and an `.xlsx` or `.xlsb` file
  carrying one fails to load (`kIoSheetCorrupt`), as Excel refuses the name
  and fails to open the file. `History`, in any case, is refused under the
  `win-*` profiles, the default included; Mac profiles accept it, as Mac
  Excel does.
- Row and column inserts and deletes that cut through a data table, a
  legacy array block or a PivotTable, or that would push one of them or a
  table past the grid edge, are refused with `kPreconditionFailed`, as Excel
  refuses them.
- What-If data tables in a loaded file are modelled (see Added) rather than
  frozen into values, so a write into a result cell is refused with
  `kPreconditionFailed`.
- A formula nesting more than 65 function calls is refused with
  `NestedFormulaTooDeep`, as Excel refuses it; up to 127 nested calls parsed
  before.

### Added

- Outline: `collapsed` on `<row>` and `<col>` is read, kept, written and
  reported by the row and column override getters, with
  `fm_sheet_set_row_collapsed` / `fm_sheet_set_column_collapsed` (and
  `setRowCollapsed` / `setColumnCollapsed` / `set_row_collapsed` /
  `set_column_collapsed`). `<sheetFormatPr outlineLevelRow outlineLevelCol>`
  is derived from the deepest row and column levels on every save, XLSX and
  XLSB, so the outline gutter shows its symbols.
- Sheet view: `showFormulas`, `showOutlineSymbols`, `windowProtection`,
  `showRuler`, `showWhiteSpace`, `defaultGridColor`, `colorId`,
  `zoomScaleNormal`, `zoomScaleSheetLayoutView`, `zoomScalePageLayoutView`,
  `<selection activeCellId>`, `<pivotSelection>` and the view's `<extLst>` are
  kept across a resave. `showFormulas` and `showOutlineSymbols` are reported by
  the view getter and have setters (`fm_sheet_set_show_formulas`,
  `fm_sheet_set_show_outline_symbols`). `<sheetView>` elements after the first
  (additional workbook windows) are kept verbatim and written after the
  modelled view.
- Split panes: `fm_sheet_set_split` (`setSheetSplit` / `set_sheet_split`)
  splits the view at a twips offset on either axis without freezing it. A split
  replaces a freeze and the other way round; a change resets the pane origin
  and active pane and keeps the active selection.
- Gradient fills (`<gradientFill>`, linear and path, with stops) are read,
  written and exposed by the fill and dxf getters and `add_fill` / `add_dxf`.
- Font `<outline>`, `<shadow>`, `<condense>` and `<extend>` are read, written
  and exposed by the font getters and setters.
- Custom icon sets: the per-level `<x14:cfIcon>` icons of an icon-set rule are
  read from and written to the linked x14 rule and exposed by the conditional
  format getter and `add_rule` (`icon_set_icons`; `icons` in the JS and Python
  bindings). A rule Excel stores only in the x14 extension, and the legacy
  copy Excel adds beside the x14 rule in `.xlsb`, each load as one rule.
- `<calcPr>`: `calcId`, `fullPrecision`, `calcOnSave`, `refMode`,
  `concurrentCalc`, `concurrentManualCount`, `forceFullCalc`, `fullCalcOnLoad`
  and `calcCompleted` are kept across a resave. `fullPrecision`, `calcOnSave`,
  `refMode` and the date system (`<workbookPr date1904>`) have typed getters
  and setters in the C ABI and every binding (`fm_workbook_full_precision` /
  `fm_workbook_set_full_precision`, `fm_workbook_calc_on_save` /
  `fm_workbook_set_calc_on_save`, `fm_workbook_ref_mode` /
  `fm_workbook_set_ref_mode`, `fm_workbook_date1904` /
  `fm_workbook_set_date1904`).
- Defined names: `comment` and `hidden` are read with
  `fm_workbook_defined_name_attrs_at` and set with
  `fm_workbook_set_defined_name_attrs` / `setDefinedNameAttrs` /
  `set_defined_name_attrs`; replacing a name's formula keeps both.
- Tables: `<tableStyleInfo>` (style name and the first-column, last-column,
  row-stripe and column-stripe toggles), the header and totals row counts and
  each column's totals-row function, label and formula are read with
  `fm_workbook_table_get_info` / `fm_workbook_table_get_column` (`tableAt` in
  JavaScript) and written with `fm_workbook_table_set_style_info` /
  `fm_workbook_table_set_column_totals` (`setTableColumnTotals` /
  `table_set_style_info` / `table_set_column_totals`). `fm_workbook_table_rename`
  (`table_rename`) renames a table and its display name and rewrites structured
  references to it in cells, defined names, conditional formats, data
  validations, calculated-column and totals-row formulas and chart series; a
  pivot cache sourced from the table follows. Turning the totals row on writes
  each column's totals cell.
- Custom table and PivotTable styles (`<tableStyles>`): listed, added or
  replaced, and removed with `fm_styles_get_table_style_count` /
  `fm_styles_get_table_style` / `fm_styles_set_table_style` /
  `fm_styles_remove_table_style` (`tableStyleCount` / `getTableStyle` /
  `setTableStyle` / `removeTableStyle`; `table_style_count` /
  `get_table_style` / `set_table_style` / `remove_table_style`), and the
  default table and PivotTable style names with
  `fm_styles_get_table_style_defaults` / `fm_styles_set_table_style_defaults`
  (`getTableStyleDefaults` / `setTableStyleDefaults`; `table_style_defaults` /
  `set_table_style_defaults`).
- Pivot table style: `<pivotTableStyleInfo>` is read and written with
  `fm_workbook_pivot_get_style_info` / `fm_workbook_pivot_set_style_info`
  (`pivotGetStyleInfo` / `pivotSetStyleInfo`; `pivot_style_info` /
  `pivot_set_style_info`).
- Workbook protection: `<workbookProtection>` (`lockStructure`, `lockWindows`,
  `lockRevision`, the legacy password hashes and the `workbook*` / `revisions*`
  algorithm, hash, salt and spin-count attributes) has a typed getter and setter
  in the C ABI and every binding (`fm_workbook_get_protection` /
  `fm_workbook_set_protection`, `fm_workbook_protection_t`;
  `getWorkbookProtection` / `setWorkbookProtection`;
  `get_workbook_protection` / `set_workbook_protection`). Hashes are stored as
  the host computes them; attributes the record does not model are kept.
- Protected ranges: a sheet's `<protectedRange>` entries (name, sqref, password
  hash, algorithm, hash, salt, spin count, security descriptor) can be listed,
  added, updated and removed (`fm_sheet_protected_range_count`,
  `fm_sheet_get_protected_range`, `fm_sheet_add_protected_range`,
  `fm_sheet_update_protected_range`, `fm_sheet_remove_protected_range`,
  `fm_protected_range_t`; `getProtectedRanges` / `addProtectedRange` /
  `updateProtectedRange` / `removeProtectedRange`; the `*_protected_range`
  methods in Python). The ranges move with row and column edits.
- Legacy notes: a note's shown/hidden state is read and written with
  `fm_sheet_get_comment_visible` / `fm_sheet_set_comment_visible`
  (`getCommentVisible` / `setCommentVisible`, `get_comment_visible` /
  `set_comment_visible`). A note's box (`x:Anchor`), its visibility and the
  authors table of its comments part, including authors no remaining note
  names, are kept across a resave; adding, removing or showing a note
  regenerates the drawing without resetting the other notes' boxes or
  visibility, and a box follows its note through row and column edits. Each
  remaining note re-emits its own original `<v:shape>` (fill, stroke, shadow,
  text box style, font, size) with only its anchor, visibility and cell
  patched; only a new note gets the default shape.
- Drawing shapes: the `xdr:sp` shapes and text boxes, `xdr:cxnSp` lines and
  connectors and `xdr:grpSp` groups of a sheet's drawing can be listed (id,
  name, kind, preset geometry, anchor, text, solid fill, outline colour and
  width, rotation, hidden), inserted, updated and removed
  (`fm_sheet_shape_count`, `fm_sheet_shape_at`, `fm_sheet_insert_shape`,
  `fm_sheet_update_shape`, `fm_sheet_remove_shape`, `fm_shape`;
  `listShapes` / `insertShape` / `updateShape` / `removeShape`; the `*_shape`
  methods in Python). An update rewrites only the fields it names and keeps
  the rest of the shape's XML; shapes the record cannot express are listed
  with `ShapePaint.Other` or empty fields and pass through a save unchanged.
  Shapes follow row and column edits like pictures do.
- Sparkline groups: the `x14:sparklineGroups` of a sheet's worksheet extension
  list can be listed with every attribute and colour the group carries, added,
  updated, given new sparklines and removed
  (`fm_sheet_sparkline_group_count`, `fm_sheet_sparkline_group_at`,
  `fm_sheet_sparkline_at`, `fm_sheet_add_sparkline_group`,
  `fm_sheet_update_sparkline_group`, `fm_sheet_set_sparkline_group_sparklines`,
  `fm_sheet_remove_sparkline_group`, `fm_sparkline_group`, `fm_sparkline`;
  `listSparklineGroups` / `addSparklineGroup` / `updateSparklineGroup` /
  `setSparklineGroupSparklines` / `removeSparklineGroup`; the
  `*_sparkline_group*` methods in Python). Attributes and children outside the
  record (`rightToLeft`, `uid`, unknown extensions) are kept by an update.
  Sparklines of a workbook loaded from `.xlsb` are listed and editable the same
  way, and groups from any source are written to `.xlsb`.
- Table slicers: a loaded table slicer (slicer part, slicer cache, the workbook
  and worksheet extension entries, the `Slicer_*` defined name and the drawing
  frame with its fallback) is kept across a resave and its frame moves with row
  and column edits. A new C ABI (`fm_sheet_slicer_count`, `fm_sheet_slicer_at`,
  `fm_sheet_add_slicer`, `fm_sheet_update_slicer`, `fm_sheet_remove_slicer`,
  `fm_slicer`; `listSlicers` / `addSlicer` / `updateSlicer` / `removeSlicer`;
  `list_slicers` / `add_slicer` / `update_slicer` / `remove_slicer` in Python)
  lists the slicers of a sheet with their caption, cache, table and column,
  style, item columns, row height, sort order, cross filter and anchor, adds a
  slicer over a table column with all its parts, relationships and content
  types, updates caption, style, item columns, row height, sort order, cross
  filter and anchor, and removes a slicer together with its cache once no slicer
  uses it, and drops the sheet's drawing when the slicer was its last object.
  Removing a table or a sheet, or deleting the rows or columns under a
  slicer's frame, removes the slicers bound to it the same way. Slicers
  over a pivot cache pass through untouched and are not listed; the selection of
  a table slicer lives in the table's AutoFilter.
- What-If data tables (`<f t="dataTable">`) are modelled instead of frozen into
  values: the sheet keeps each table's result range, orientation and input
  cells with the `dt2D`, `dtr`, `r1`, `r2`, `del1`, `del2`, `ca` and `aca`
  attributes, and a save writes the formula back on the anchor cell. A recalc
  fills the result cells of one-variable (column or row input) and
  two-variable tables by substituting each value into the input cell(s),
  recalculating the cells that read them and restoring the inputs; under
  `calcMode="autoNoTable"` it leaves them alone until
  `fm_workbook_recalc_data_tables` (`recalcDataTables` /
  `recalc_data_tables`), and `partial_recalc` never computes them. The result
  range refuses cell writes ("cannot change part of a data table"). Row and
  column inserts and deletes shift the range and the input cells; deleting an
  input cell sets `del1` / `del2` and the table evaluates to `#REF!`, deleting
  the range removes the table. `fm_sheet_add_data_table` /
  `fm_sheet_data_table_at` / `fm_sheet_remove_data_table` (`addDataTable` /
  `getDataTable` / `removeDataTable` in JavaScript, `add_data_table` /
  `get_data_table` / `remove_data_table` in Python, `fm_data_table`) create a
  table, read the one covering a cell, and remove it with its result cells.
  XLSB reads and writes the table as `BrtTable` (428) ahead of the anchor's
  plain value record, with the always-calculate flag. Inserting or deleting
  rows or columns that cut into a data table (its result range, or the header
  row and column holding the formula and input values) is refused with
  `kPreconditionFailed`; deleting the whole block removes the table. Result
  cells report `=TABLE(row_input,col_input)` through the formula getters,
  `FORMULATEXT` returns `{=TABLE(...)}` like an array formula, and `ISFORMULA`
  is TRUE.
- XLSB parity with the features above: gradient fills (`BrtFill`, `BrtDXF`
  gradient props); dxf font outline, shadow, condense and extend, and explicit
  "off" toggles in dxf fonts; calculation settings (`BrtCalcProp`); workbook
  properties (`BrtWbProp`); workbook protection including lockWindows,
  lockRevision and the revisions password (`BrtBookProtection`,
  `BrtBookProtectionIso`); every workbook window (`BrtBookView`) and the sheet
  views of extra windows; sheet view mode and `activeCellId`; protected ranges
  (`BrtRangeProtection`, `BrtRangeProtectionIso`); custom icon sets as x14 rules
  (`BrtBeginCFRule14`, `BrtBeginIconSet14`, `BrtCFVO14`, `BrtCFIcon`); sparkline
  groups; tables (`BrtBeginList` and friends, with style, flags, totals,
  calculated columns, AutoFilter and sort state); structured references
  (`PtgList`); pivot filters and style info on read; and the workbook's slicer
  cache lists, so table slicers survive an `.xlsb` resave. Table slicers from an
  `.xlsx` source, security descriptors of protected ranges,
  `workbookPr dateCompatibility` / `allowRefreshQuery`, custom table and
  PivotTable styles, sparkline groups whose references cannot be encoded, and
  pivot tables from an `.xlsx` source are not written to `.xlsb` and are
  counted in `deferred_feature_count`.

### Changed

- Formula nesting follows Excel's limits: 256 parentheses and 1024 unary
  signs parse and evaluate, where an expression depth of 128 was the limit
  before. The WASM stack (1.5 MiB) and the worker-thread stack are sized for
  them.
- The evaluated-array element bound is raised from 53,687,091 to
  107,374,182, as measured on Mac Excel 365: `A:CX+0` evaluates, while
  `A:CY+0` and `SEQUENCE(1048576,103)` give the number 0. The bound applies
  under every profile, though Windows Excel abandons smaller arrays. `SUM`,
  `AVERAGE`, `MIN`, `MAX`, `PRODUCT` and `COUNTA` fold whole-axis operation
  results without materialising them.
- Recalc visits cells in a deterministic order, so spill winners and legacy
  array blocks resolve the same way on every run; a legacy array block whose
  formula fails to parse keeps its rectangle.

### Performance

- The `.wasm` builds compile the cold evaluator families (special
  functions, financial, `FORECAST.ETS`, format parsing, `GROUPBY` /
  `PIVOTBY`, regex, database, regression) and PCRE2 for size, about 26 KB
  smaller over Brotli with bit-identical results.
- Chart point gathering and sheet-name comparison allocate and decode less;
  table-name comparison allocates less, and pivot field-order updates are
  linear.

### Fixed

- `ADDRESS` quotes a sheet name by the active locale's rules (its own R1C1
  letters and boolean names, U+3000 in double-byte locales), and
  `CELL("address")` on another sheet quotes it by the cross-workbook rule
  instead of appending it bare. A sheet name opening with `.` is quoted in
  every sheet qualifier, like a leading digit.
- `TEXT` and number formats: a subnormal scientific exponent follows the
  integer placeholder count; a fraction format's integer part is grouped,
  with the locale's group separator and with literals between its
  placeholders kept in the group; in fr-FR / ru-RU a blank run after a digit
  placeholder is one grouping run; the 127-decimal cap is checked on the
  stored format.
- `VALUE` and text-to-number coercion reject a grouped number only when its
  first group is all zeros (`046.565.118` parses in de-DE; `000.509` and
  `0,123` are `#VALUE!`).
- The `^` operator flushes a subnormal square to zero, as `POWER` does.
- `ROUNDUP` rounds on the 15-digit view, so binary noise no longer adds a
  unit (`ROUNDUP(0.1+0.2,1)` is 0.3).
- `IMSUM` and `IMPRODUCT` fold ranges and arrays, a blank cell as 0.
- `FORECAST.ETS` targets more than 2^31-1 steps past the timeline are
  `#NUM!`, and calendar targets are placed by month difference; `FORECAST`
  and the regression family lift array-valued targets.
- `GROUPBY` / `PIVOTBY` number sort orders key columns first, then value
  columns, and call a two-argument aggregator (`PERCENTOF`, a two-parameter
  `LAMBDA`) with the group slice and the whole filtered column.
- `LOOKUP`, `MATCH` and `INDEX` read whole-axis references by their declared
  shape. In `LOOKUP` an empty whole-axis vector is `#N/A`, and a result past
  the populated extent is blank.
- `REGEXREPLACE` bounds an unterminated `\U` / `\L` run by the text cap, and
  `XMATCH` / `XLOOKUP` in regex mode return `#VALUE!` when the match limit is
  exhausted, as `REGEXTEST` does.
- Scalar errors broadcast against arrays.
- Conditional formatting: Cell Value rules compare with worksheet
  semantics (blank is 0 against a number and `""` against text; numbers <
  text < booleans). Error cells match "does not contain" text rules and no
  longer match No Blanks.
- PivotTables: Percent of Grand Total and Index divide by the whole-table
  aggregate when the grand totals are hidden.
- Structural edits: `:` chain operands shift on their qualified sheet,
  whole-axis references stay anchored, structured references to a removed
  table or column become `#REF!`, and internal hyperlink locations are left
  as written.
- Formula text that shows `SINGLE` / `ANCHORARRAY` as `@` / `#`, or places
  an implied `@`, is no longer garbled when the formula holds non-ASCII text
  (a Japanese sheet name, a string literal), and `ANCHORARRAY` of a cell
  that is not a spill anchor keeps its call spelling instead of an
  unparsable `x#`.
- Pivot field-order updates accept getter-owned storage and keep the values
  marker in place when fields move between axes.
- `.xlsx`: unmodelled attributes and children of comments, conditional
  formats, styles, pivot caches and tables are kept across a resave, and
  sizes read from a file are bounded on load, recalc and save (a pivot cache
  above 2^25 record slots fails the load with `kIoFileTooLarge`).
- `.xlsx` and `.xlsb` keep the source sheet ids instead of renumbering them
  by position.
- Python: the type stub declares the runtime dataclass constructors and enum
  members.
- Workbook-level extension entries that name a relationship by `r:id` (slicer
  caches, timeline caches) follow the relationship ids the writer assigns
  instead of pointing at a stale id after a resave.
- `.xlsb` tables kept verbatim (a part the model cannot hold) keep their
  name and id: creating a table onto either is refused.
- JS packages: any `Workbook` method called from that workbook's iterative
  progress callback throws `TypeError` naming `kGraphRecalcReentrant` (4005)
  and changes nothing, instead of blocking forever (native and threads
  builds) or changing the workbook mid-sweep (single-threaded build).
- `@libraz/formulon-native`: `dispose()` from a spec getter or callback while
  a method of the same workbook is running throws `TypeError` (`cannot
  dispose Workbook during an active method call`), as `delete()` does on the
  WASM package, instead of freeing the handle the running method still uses.
- `@libraz/formulon-native`: a workbook with an iterative progress callback
  is garbage-collected once unreachable, also when the callback refers to
  the workbook.
- `@libraz/formulon`: `setIterativeProgress` and `setLogSink` reject a value
  that is not a function or `null` with status 2 and keep the installed
  callback, as the native package does, instead of accepting it and failing
  at the next recalc.
- `@libraz/formulon`: a string field of a spec object that holds a lone UTF-16
  surrogate is stored with U+FFFD in its place, as on the native package,
  instead of reaching the engine as ill-formed UTF-8.
- `pivotFieldSetDateGroup` takes 7 to 10 arguments on both JS packages, with
  `intervalDays` defaulting to 1 and `startSerial` / `endSerial` to -1.
- A failed `getCellXf` / `getCellStyleXf` reports `locked: true` and a failed
  `validateValue` reports `valid: false` on both JS packages.
- The `@libraz/formulon-native` TypeScript declarations type pivot fields and
  parameters with the same enum types as `@libraz/formulon` (`PivotAxis`,
  `PivotAggregation`, `PivotSortOrder`, ...), so code written against one
  package's types compiles against the other's.

## [0.17.0] - 2026-10-21

### BREAKING

- **C ABI:** `fm_sheet_view_t` grows by four `uint32_t` fields
  (`top_left_row`, `top_left_col`, `pane_top_left_row`, `pane_top_left_col`),
  from 48 to 64 bytes on 64-bit targets. A caller compiled against an earlier
  header must be recompiled before calling `fm_sheet_get_view`.
- Column widths are Excel's stored width, padding included, as `<col width>`
  holds it: Calibri 11 stores the 8.43-character column as 9.140625. This
  applies to the column-width setters and getters, the sheet default column
  width and `fm_sheet_column_chars_to_pt` / `columnCharsToPt` /
  `column_chars_to_pt` and their inverses in every binding. On screen a width
  snaps to whole pixels of the Normal font's maximum digit width. The default
  column width is derived from `baseColWidth` and the Normal font instead of
  a fixed 8.43 characters, and print layout with an empty font table uses the
  Calibri 11 model.
- Hiding or removing the active sheet activates the next visible sheet to the
  right, else the nearest visible one on the left, and moves `tabSelected`
  with it.
- Saved XLSX and XLSB files no longer carry the Show Phonetic flag on formula
  cells, and loading drops it there; Excel keeps it on constants and blanks
  only.
- A sheet added to a workbook created from an Excel profile takes the first
  sheet's row metrics, page margins and `phoneticPr`, fixed at creation, so a
  later `set_excel_profile` does not change what added sheets get. Plain and
  loaded workbooks still add bare sheets.

### Added

- The active sheet: `fm_workbook_get_active_sheet` /
  `fm_workbook_set_active_sheet` (`activeSheet` / `setActiveSheet`,
  `active_sheet` / `set_active_sheet`), backed by `activeTab` and kept across
  sheet inserts, removals and moves. Setting it moves `tabSelected` and
  rejects hidden sheets.
- The per-cell Show Phonetic flag (`fm_workbook_get_cell_show_phonetic` /
  `fm_workbook_set_cell_show_phonetic`, `getCellShowPhonetic` /
  `setCellShowPhonetic`, `get_show_phonetic` / `set_show_phonetic`) and
  `fm_workbook_has_cell_phonetic_properties` for an explicit `phoneticPr`.
  Row and column phonetic flags are preserved.
- Pivot definition read-back: cache id, name, grand totals, fields and their
  items, data fields, and row and column field order
  (`fm_workbook_pivot_get_*` / `fm_workbook_pivot_field_*` /
  `fm_workbook_pivot_data_field_at`, `pivot*` in Node, `pivot_*` in Python).
  A data field added or replaced with no name is named from the profile's
  pivot wording ("Sum of Sales", "合計 / Sales").
- Sheet view scroll origin, pane origin and selection:
  `fm_sheet_set_top_left`, `fm_sheet_set_pane_top_left`,
  `fm_sheet_get_selection` and `fm_sheet_set_selection`
  (`setSheetTopLeft` / `setSheetPaneTopLeft` / `getSheetSelection` /
  `setSheetSelection`, `set_sheet_top_left` / `set_sheet_pane_top_left` /
  `get_sheet_selection` / `set_sheet_selection`). Unfrozen split panes and
  `frozenSplit` round-trip through XLSX and XLSB, and selections are saved one
  per pane in Excel's order.

### Fixed

- XLSB files with frozen panes open in Excel: panes are written as `BrtPane`
  (record 151) with the column split first, and Excel's own frozen panes are
  no longer rejected on load. Per-pane selections (`BrtSel`) are read and
  written, and pane and selection records of other windows or custom views
  are no longer folded into the main view.
- A selection in a pane the view lacks is dropped on save, since Excel refuses
  such a file, and an active cell outside every selected range moves to the
  first range's top-left.

### Performance

- Aggregates compact their values in place, one parsed sheet relationships
  document is reused across load phases, `VLOOKUP` / `HLOOKUP` / `MATCH` share
  one exact-text scan, and table references skip redundant validation.
- The WASM binding glue is compiled for size while the engine stays at `-O3`,
  which takes about 75 KB raw and 6 KB Brotli off each `.wasm`.

## [0.16.1] - 2026-10-21

### Fixed

- `fm_styles_get_num_fmt_string` / `getNumFmt` / `get_num_fmt` now resolve
  built-in ids that have no ECMA-376 code (East Asian date ids such as 31 and
  55-58, Thai-digit ids 59-81, the currency ids 5-8) to the code the
  workbook's Excel profile renders them with, instead of failing. A cell
  saved by a Japanese, Chinese or Korean desktop spreadsheet therefore reports
  its date format. Ids with an ECMA-376 code still return that code.

## [0.16.0] - 2026-10-19

### BREAKING

- **C ABI and every binding:** the five picture operations are renamed to
  drawing-object operations and now act on charts as well as pictures. A
  snapshot of a chart carries its chart part and every part the chart's
  relationships reach; snapshots taken by earlier versions still restore.
  Migration, by rename only:

  | Before | After |
  |---|---|
  | `fm_sheet_remove_image` / `removeImage` / `remove_image` | `fm_sheet_remove_drawing` / `removeDrawing` / `remove_drawing` |
  | `fm_sheet_set_image_anchor` / `setImageAnchor` / `set_image_anchor` | `fm_sheet_set_drawing_anchor` / `setDrawingAnchor` / `set_drawing_anchor` |
  | `fm_sheet_set_image_z_order` / `setImageZOrder` / `set_image_z_order` | `fm_sheet_set_drawing_z_order` / `setDrawingZOrder` / `set_drawing_z_order` |
  | `fm_sheet_snapshot_image` / `snapshotImage` / `snapshot_image` | `fm_sheet_snapshot_drawing` / `snapshotDrawing` / `snapshot_drawing` |
  | `fm_sheet_restore_image` / `restoreImage` / `restore_image` | `fm_sheet_restore_drawing` / `restoreDrawing` / `restore_drawing` |
  | `fm_image_anchor` | `fm_drawing_anchor` (same layout) |
  | `FM_IMAGE_RESTORE_NEW_ID` | `FM_DRAWING_RESTORE_NEW_ID` (same value) |

  `probeImage`, `getImage` and `insertImage` keep their names. The old names
  are gone rather than aliased, so a caller linked against an earlier
  shared library must be recompiled.
- Formula text is stored the way Excel stores entered text, so a formula
  reads back respelled: number literals are respelled from their value
  (`1.50` reads back as `1.5`), a unary sign touching a literal is absorbed
  (`+5` as `5`), and whitespace Excel drops is removed (trailing, before `,`,
  around `:`, inside array constants). Legacy (non-dynamic-array) formulas
  show Excel's implied `@`. This applies to cells, defined names,
  conditional-format and data-validation formulas.
- A formula number literal whose leading digit lies outside Excel's range
  (`1E308` and above, `1E-310` and below, other than zero) is rejected as an
  invalid number literal instead of evaluating to `#NUM!` or 0. Literals of
  64 bytes or more are accepted and keep 15 significant digits.
- Formula entry (`fm_workbook_set_formula`, defined names, conditional-format
  and data-validation formulas) rejects a tab, a carriage return, and
  whitespace between a built-in function name and its `(`, as Excel does,
  with `kParserUnexpectedToken`.
- Writing a value, text, formula or blank into a cell of a legacy array
  formula (CSE) block other than its anchor fails with `kPreconditionFailed`
  and leaves the block unchanged; writing to the anchor replaces the whole
  block. Every cell of a block reports the block's formula (`FORMULATEXT`
  shows `{=...}`, `ISFORMULA` is TRUE).
- Font sizes must be finite and non-negative, `fm_workbook_set_iterative`
  takes a finite non-negative `max_change`, and the pivot-cache number
  setters take finite values; anything else fails with `kInvalidArgument`.
  `fm_sheet_cf_add_rule` fails with `kInvalidArgument` instead of overflowing
  when the auto-assigned priority would pass `INT32_MAX`.
- Python integer parameters reject a fractional or non-finite number with
  `ValueError` and a non-number (including a numeric string) with `TypeError`
  instead of truncating or converting it.
- Ad-hoc formula evaluation (`fm_workbook_evaluate_formula`,
  `fm_workbook_evaluate_formula_array` and `fm_workbook_evaluate_cf_formula`,
  in every binding) rejects an out-of-grid target or anchor with
  `kInvalidArgument`.
- Both JavaScript bindings reject an iterative-calculation cap that is not an
  int32 (above 2^31-1, fractional or non-finite) with `RangeError` instead of
  wrapping it.

### Added

- Charts. A chart is addressed by sheet and drawing object id, the id
  `listDrawingObjects` reports for a chart object. `getChart` reads the
  chart's typed view (plots, series and their references, formatting, data
  labels, axes, title, legend, filters, and the source range the series were
  split from); `getChartData` resolves the series against the current cell
  values by Excel's rules for hidden cells, blanks and errors, and
  `previewChartData` does the same for uncommitted source, direction,
  heading and blank-handling changes without touching the workbook.
  `insertChart` creates a column, bar, line, area, pie, doughnut or scatter
  chart from one range. `setChartType`, `setChartTitle`, `setChartLegend`,
  `setChartAxis`, `setChartDisplay`, `setChartSeries`, `setChartPoint`,
  `setChartPlot`, `setChartFilter`, `addChartSeries`, `removeChartSeries` and
  `setChartSource` edit the chart part in place, keeping everything they do
  not touch. In the C ABI these are `fm_sheet_chart_open` /
  `fm_sheet_chart_data_open` / `fm_sheet_chart_preview_open` (owned handles
  read through `fm_chart_*` / `fm_chart_data_*` accessors),
  `fm_sheet_insert_chart` and one `fm_sheet_set_chart_*` /
  `fm_sheet_add_chart_series` / `fm_sheet_remove_chart_series` entry per
  edit, with the status codes `kIoChartUnparseable` (5025),
  `kIoChartInvalidReference` (5026) and `kIoChartUnsupportedKind` (5027).
  `setChartTitle` also takes a title object (`{ kind, text?, formula? }`;
  `ChartTitleInfo` in Python), so `ChartTitleKind.Auto` returns the title to
  Excel's automatic one. The `title` of a `setChartAxis` patch takes the same
  forms (`fm_chart_axis_patch.title` is an `fm_chart_title_t`). Each
  `getChartData` series carries `xFormatCode`
  (`x_format_code`, `fm_chart_data_series_t.x_format_code`), the number format
  of its x values.
- Chart caches are refreshed on save the way Excel writes them, and chart
  references follow row and column inserts and deletes and sheet renames and
  deletes (`#REF!` as Excel writes it). Changing a chart's type carries series
  colours, lines and points the way Excel's own type change does. A chart
  with no title and one unfiltered series reads as an automatic title, and
  marker symbols cover `x`, `star`, `plus`, `dot`, `dash` and (read-only)
  `picture`.
- Pivot data fields can show values as `PercentOf`, `PercentOfRunningTotal`,
  `RankAscending` and `RankDescending` (`FM_PIVOT_SHOW_AS_PERCENT_OF` = 12 to
  `FM_PIVOT_SHOW_AS_RANK_DESCENDING` = 15 in the C ABI, `ShowValuesAs` in the
  bindings). The x14 `pivotShowAs` extension and `subtotalHiddenItems` are
  read and written, and every value-filter comparison type is read.
- Built-in number format ids (5-8 and 41-44 included) resolve to the
  profile's locale code in cell display, effective styles and data
  validation, and the `fm_styles_*` setters accept those ids.

### Changed

- The WASM size report's ceilings moved to 4.25 MiB soft / 4.5 MiB hard
  uncompressed and 1216 KiB soft / 1280 KiB hard Brotli, to take in the
  chart API.
- Numbers Formulon writes into `.xlsx` parts (cell values, chart values,
  pivot cache records, external-link caches) are spelled the way Excel writes
  them: fixed notation with 15 significant digits when that round-trips
  (else 17) from 0.1 up to but not including 1E+16, 17-digit exponential
  form with a signed exponent (`1.23E-2`) outside it, and `0` for zero and
  denormals.
- Inserted rows take the row style, custom height and per-cell formats of the
  row above; inserted columns take the width, column style and per-cell
  formats of the column to the left. Sparklines hosted on the source row or
  column are cloned onto each inserted one. Hidden state and outline level
  are never inherited.
- On `mac-*` profiles `FILTERXML` and `WEBSERVICE` are unknown names
  (`#NAME?`, also as a function value), as Mac Excel provides neither;
  Windows profiles are unchanged. `CUBE*` functions return `#N/A` without a
  cube connection instead of `#NAME?`.
- A flat chain of same-precedence operators is limited by the 8,192-character
  formula length rather than the nesting cap. `ROWS` and `COLUMNS` size their
  operands from shape without reading values. An evaluated array of more than
  53,687,091 elements abandons the formula, which then holds 0. With
  iterative calculation off, a cell that is the only member of its cycle
  (`=A1+1` in `A1`) holds 0; cycles through several cells keep `#REF!`.
- `N` and `T` read the top-left cell of a reference argument instead of
  spilling over it, and `TYPE` of an array argument is 64.
- Results built on the platform's `sin`, `cos`, `tan`, `atan`, `exp`, `log`
  and `pow` take their last bits from the host math library. Mac Excel shares
  the macOS library's last bits, so on macOS they match exactly; on Linux
  and in WebAssembly a few arguments (for example `TAN(1)`, `ACOT(0.5)`)
  differ from Mac Excel by 1-2 ulp and give the correctly rounded value,
  which is what Windows Excel returns. The oracle compares those cases
  exactly on macOS and within 4 ulp elsewhere.

### Performance

- Pivot Sum aggregation over sparse records takes one pass. Order statistics
  (`AGGREGATE`, `LARGE`, `SMALL`, `PERCENTILE`) select with `nth_element`,
  sparse range precedent tracing walks stored rows only, and parallel
  recalculation drops per-pass copies.
- Releasing a range dependency is constant time and dependency edges are
  cleared in bulk; the iterative solver indexes cycle state by position.
- Conditional-format evaluation computes `AboveAverage` statistics once per
  population and visits only populated cells of oversized ranges; AutoFilter
  custom conditions are prepared once per filter, and clearing or remapping
  styles no longer walks every row.
- Effective style resolution parses the theme once per call, and print-title
  and manual page-break lookups use indexed tracks.

### Fixed

- Number formats follow Excel for fractions (`#` placeholders pad nothing, a
  numerator that rounds to 0 drops the fraction group, fixed denominators
  round half up, approximations take the last convergent that fits), `?` and
  `0` placeholders on a zero integer part, thousands separators, scientific
  mantissas taken from the 15-digit view, `General` and `@`, and subnormals.
  Ties in the 15-digit view round toward zero in `&`, `TEXT`, `FIXED` and
  `ROUND`. `TEXT`, `FIXED`, `DOLLAR` and `USDOLLAR` return `#VALUE!` when the
  output exceeds 255 characters.
- Text that converts to a number (`VALUE`, arithmetic on text, criteria)
  keeps 15 significant digits by truncation and follows Excel's range,
  percent and 255-character rules; `NUMBERVALUE` is a pre-pass in front of
  the same recogniser. Date and time text uses one recogniser for all
  fourteen locales, including Korean, Chinese meridiem, era and kanji forms,
  per-locale currency spellings and month-name matching.
- Subnormal results of `*`, `/`, `^`, unary `%`, `PRODUCT`, `SUMPRODUCT` and
  `SUMSQ` flush to 0 as Excel does, and `SIGN`, truthiness and unary minus
  use Excel's one-sided zero test. The cancellation snap scales with the left
  operand, and `SUM`, `AVERAGE`, `SUBTOTAL` and `AGGREGATE` snap only the
  final addition. `ROUND`, `ROUNDUP`, `ROUNDDOWN`, `TRUNC`, `CEILING` and
  `FLOOR` work on the 15-digit view, and `MOD` is `#NUM!` from a quotient of
  1,125,900,000,000. Native builds are compiled without fused multiply-add,
  so arm64 arithmetic matches x86-64 and WASM.
- Trigonometric functions are `#NUM!` from |x| = 2^27, and the hyperbolic,
  inverse trigonometric, `RADIANS`, `DEGREES`, `FACT`, `ERF` / `ERFC`, `LOG`
  and `TIME` results follow Excel's formulas. `BITLSHIFT` / `BITRSHIFT`
  decide on the net shift; `COMBIN`, `COMBINA`, `PERMUT` and `PERMUTATIONA`
  are `#NUM!` at 2^31-1; `BASE` and `DECIMAL` follow Excel's length and
  prefix rules. `BESSELJ`, `BESSELY`, `BESSELI` and `BESSELK` follow Excel's
  algorithms and order limit, `CONVERT` uses Excel's unit table and
  arithmetic, and the `IM*` functions use Excel's parsing, arithmetic and
  text output.
- `YEARFRAC` basis 1 counts the 1900-02-29 serial, `DATEDIF` `"YD"` follows
  Excel, `WORKDAY` truncates a fractional start, `EOMONTH` before 1900 is
  `#NUM!`, `DATE(1900,3,0)` is serial 60, and `HOUR` / `MINUTE` / `SECOND`
  honour the 1904 bound.
- The gamma, beta, normal, binomial and related distributions and their
  inverses are evaluated with Excel's kernels, so large and extreme
  parameters give Excel's values instead of `#NUM!` or a last-digit drift.
  `T.TEST`, `F.TEST`, `CHISQ.TEST`, `Z.TEST`, `CONFIDENCE.*`,
  `BINOM.DIST.RANGE`, `CORREL`, `RSQ`, `SLOPE`, `FORECAST` and `COVAR`
  compute in Excel's order. The variance family takes a second pass where
  the one-pass sum cancels; `PERCENTILE.*`, `QUARTILE.*` and `MEDIAN`
  interpolate as Excel does; the `*A` aggregates skip logical and text
  elements of array arguments. `LINEST`, `LOGEST`, `TREND` and `GROWTH` fit
  with Excel's pivoted QR and dropped-column rule, and `MINVERSE` solves by
  partial-pivot LU.
- `AGGREGATE` validates `k` as `LARGE` / `SMALL` do, never counts error cells
  for function 2, and counts them for function 3 under options that keep
  errors.
- `FORECAST.ETS` and its `.CONFINT`, `.SEASONALITY` and `.STAT` forms follow
  Excel's timeline handling, seasonality detection, smoothing-parameter
  candidates, holdout scoring and trend reset; out-of-range options are
  `#NUM!`.
- The payment, present/future value, depreciation and bond functions compute
  in Excel's operation order with its `log1p` / `expm1` / power kernels, so
  results match to the last digits and `PMT` no longer loses precision at
  small rates. `DDB` accepts fractional periods. `RATE`, `IRR`, `XIRR` and
  `YIELD` run Excel's solvers (a `RATE` / `IRR` guess at or below -1 is
  `#VALUE!`, a single `XIRR` flow is `#N/A`), and `IRR`, `NPV` and `PV`
  return `#DIV/0!` or `#NUM!` as Excel does when the discount factor
  underflows or overflows.
- `MATCH` type 1, `LOOKUP`, approximate `VLOOKUP` / `HLOOKUP`, `XLOOKUP` and
  `XMATCH` share one search that reproduces Excel's result on unsorted and
  mixed-type data and orders errors above booleans. Wildcard criteria match
  per UTF-16 unit, consistent with `SEARCH`.
- On Mac profiles `LEN`, `LEFT`, `RIGHT`, `MID`, `REPLACE`, `FIND` and
  `SEARCH` count a supplementary-plane character as one position; other
  profiles keep UTF-16 units. The `LENB` family uses per-locale byte widths.
  `UPPER`, `LOWER`, `PROPER` and the case escapes of `REGEXREPLACE` map case
  across Unicode.
- `REGEXTEST`, `REGEXEXTRACT` and `REGEXREPLACE` accept the empty pattern,
  expand replacements with Excel's grammar, and return `#VALUE!` when a match
  exceeds the step limit; `REGEXEXTRACT` rejects a `return_mode` above 2.
- Sheet names are quoted by Excel's rule in formulas and `ADDRESS`, and names
  qualified by a sheet or book are not localized as built-ins.
- Iterative calculation stops on a workbook-wide change test, sweeping
  independent circular components together, as Excel does.
- `ROW`, `COLUMN`, `OFFSET` and range endpoints on a missing sheet return
  `#REF!`, `Sheet!Name` with an unknown qualifier is `#NAME?`, and `OFFSET`
  resolves a chained `:` base to its bounding box. An arena exhausted during
  parallel recalculation reports `kOutOfMemory` instead of committing
  `#REF!`.
- Huge and non-finite numeric arguments to the `D*` functions, `XMATCH`,
  `XLOOKUP`, `TOCOL`, `WRAPROWS`, `WRAPCOLS`, `OFFSET`, `INDEX`, `RANK` and
  the date and bond functions saturate or give `#VALUE!` instead of an
  undefined conversion.
- Pivot values follow Excel cell for cell: Sum totals are built from the
  rendered cells, an all-blank cell or total renders blank instead of 0,
  custom subtotals switch Sum totals to the accumulator model, hidden, page
  and label filters compact records as Excel does, and show-values-as
  results follow Excel's blank and error precedence. A field without
  `sortType` keeps its `<items>` document order.
- `PIVOTBY` without column subtotals no longer reads past its subtotal table,
  and a spill anchor's cached text stays valid after the spill region is
  cleared or recommitted.
- Data validation evaluates its formulas as if the proposed value were in the
  cell, so a rule that refers to the cell it validates (for example
  `=COUNTIF($A:$A,A1)=1`) checks the new value rather than the stored one.
- On a sheet whose rows are hidden by default (`zeroHeight`), `SUBTOTAL`,
  `AGGREGATE`, print geometry and AutoFilter treat rows without their own
  visibility as hidden, row style, height and outline edits keep a row's
  visibility, an XLSB save keeps those rows hidden, and
  `effective_default_row_height` is 0.
- `SUBTOTAL` and `AGGREGATE` reached through a defined name recalculate when
  row visibility changes.
- Colour scales, data bars and percent thresholds stay finite for extreme
  endpoints, and an out-of-range percentile threshold is clamped.
- An external reference matches a cached sheet name case-insensitively
  beyond ASCII (`[Book.xlsx]ärea!A1` finds `Ärea`), in evaluation and in an
  XLSB save.
- `_xHHHH_` escapes are decoded when reading formulas, cached strings,
  comment authors and external-link text, so text no longer grows across
  saves. Relationship parts are matched by namespace under any prefix, zip
  entry names with an embedded NUL are rejected, and the streaming reader
  reads inline strings like the DOM reader.
- Non-finite numbers are stored as `#NUM!` in external-link caches and
  `.xlsb` cells, `.xlsb` formula string literals keep surrogate pairs, and
  the XLSB shared-string header counts total references separately from
  unique strings.
- Format detection reads the parsed `[Content_Types].xml` declarations, so
  incidental text in that part no longer routes a package to the XLSB reader.
- A drawing edit fails with `kPreconditionFailed` when no drawing object id
  is left, instead of writing a wrapped id.
- Malformed UTF-8 is rejected the same way by every code-point consumer and
  by sheet-name validation.
- Node `recalc()` reports its failure message and code instead of losing
  them, and Python workbook calls hold the handle lock for the whole
  operation, so closing a workbook on another thread cannot free it mid-call.

## [0.15.2] - 2026-10-14

### BREAKING

- Python `ConditionalFormatInput` drops `dxf_id_engaged`, `op_engaged`,
  `rank_engaged`, `std_dev_engaged` and `time_period_engaged`: `dxf_id`, `op`,
  `rank`, `std_dev` and `time_period` are `Optional` and a value that is not
  `None` is engaged, as in the WASM binding. `ConditionalFormat.op`, `rank`,
  `std_dev` and `time_period` are `None` when the rule does not state them, so
  a rule read back can be passed to `add_conditional_format` unchanged.
- `fm_workbook_pivot_field_set_sort` / `pivotFieldSetSort` /
  `pivot_field_set_sort` take a sort order, `fm_pivot_sort_order_t` /
  `PivotSortOrder` (descending, ascending, manual), in place of the
  `ascending` flag. Manual keeps a field's items in the order they were added
  with `pivot_field_add_item` and saves as `sortType="manual"`; it takes no
  by-field key. An order outside the three values is rejected. A field item
  added by label now binds to the cache's shared item of that label, so its
  saved `<item x>` matches. Pass `1` / `true` for the former ascending and
  `0` / `false` for descending.
- `fm_sheet_cf_add_rule`, `fm_sheet_add_validation`,
  `fm_workbook_set_defined_name[_scoped]` and formula-typed conditional-format
  thresholds store formula text without its leading `=`, the form files carry
  and readers produce: getters return `$B1>100` rather than `=$B1>100`, and
  saved files no longer contain `<formula>=…`. A formula starting with `==`
  and an unparseable formula-typed threshold are rejected with
  `kParserUnexpectedToken`.
- `fm_sheet_cf_add_rule` rejects an icon-set rule unless it carries exactly
  N-1 thresholds for its N-icon set (`kInvalidArgument`).
- Pivot fields added through the Python, Node and WASM bindings default
  `subtotal_top` to true, matching Excel and OOXML; set it to false explicitly
  for subtotals at the bottom.
- `fm_workbook_pivot_field_add` and `fm_workbook_pivot_field_set_axis` keep
  the row and column field orders in step with the field's axis, so a field
  added on the row or column axis shows without a separate order call.
  `fm_workbook_pivot_set_row_field_order` / `_col_field_order` set the axis's
  whole membership (unlisted former members leave the axis) and reject
  duplicate indices. `fm_workbook_pivot_field_clear` also empties both orders.
- `fm_font_record` gains `has_size` (`hasSize` / `has_size` in the bindings),
  so a differential-format font can omit `<sz>` and leave the size unchanged;
  cell fonts always keep their size. In JavaScript an omitted `size` implies
  `hasSize: false`; in Python pass `has_size=False`.

### Fixed

- On Windows profiles, the arithmetic operators (`+ - * / ^ %`), `EXP` and
  `LN` round as Windows Excel does (x87 extended precision, then double), so
  results that land on a rounding tie match it to the last bit. Aggregate
  functions and other builtins still compute in double.
- XLSB loading now decodes differential formats (`BrtDXF`), so
  conditional-format rules keep a resolvable `dxf_id`, `.xlsb` to `.xlsx`
  conversion keeps their formatting, and `add_dxf` after loading no longer
  collides with Excel's first dxf.
- Saving a loaded `.xlsb` as XLSB after editing its styles (for example
  adding a differential format) regenerates `styles.bin` from the model
  instead of failing with `kIoXlsbRetainedPartStale`; an unedited styles
  table is still written byte for byte.
- Text, blanks, errors and timePeriod conditional-format rules added without a
  formula now carry the formula Excel stores (e.g.
  `NOT(ISERROR(SEARCH("ab",B1)))`, `FLOOR(B2,1)=TODAY()-1`), anchored at the
  top-left cell of the first range, so both writers save it and `formula1`
  reads back populated.
- An icon-set rule with surplus thresholds no longer renders an out-of-range
  icon; evaluation ignores thresholds past the set's N-1.
- Compact and outline pivots with a subtotal at the bottom
  (`subtotalTop="0"`) or no subtotal (`defaultSubtotal="0"`) show the outer
  item on its own heading row, and a bottom subtotal row is labelled
  "<item> Total" in the locale's wording.
- Compact-layout row cells report the row's hierarchy depth and its own field
  name instead of depth 0 and the outermost field.
- The Python type stub declares `Workbook.set_column_style` and
  `clear_column_style`, which existed at runtime only.
- Conditional-format relative formulas, and the formulas derived for text,
  blanks, errors and time-period rules, anchor at the top-left of the
  bounding box of all sqref ranges, as Excel does, instead of at the first
  listed range.
- Pivot custom subtotal rows are labelled with their function name in the
  UI locale's wording and position ("East 最大", "Somme East") and follow
  Excel's fixed function order; two or more
  functions place the subtotals below the group whatever `subtotal_top`
  says. The pivot XML writer no longer emits the default subtotal item next
  to custom ones, and the reader treats custom functions as replacing the
  default subtotal.
- Excel opens pivot tables authored with Formulon: a table with two or more
  data fields writes the Values field on an axis (Excel refused the file
  without it), a pivot field is flagged `dataField` only when a data field
  refers to it, and authored tables carry Excel's `showAll="0"`,
  `updatedVersion`, location offsets and values-row extension, plus
  `refreshOnLoad` on the cache so Excel renders them on open.
- Windows profiles (`win-*`, including the default `win-365-en_US`) collate
  text with the locale-aware collator measured on Windows Excel 365 ja-JP
  instead of an ASCII case-insensitive byte comparison. Comparisons, sorting,
  lookups and criteria follow Windows word sort: hyphen and apostrophe count
  only after every other difference, `æ` equals `ae`, half-width kana sort
  before full-width and katakana before hiragana, Thai prefix vowels keep
  their position, and Hangul sorts before Han.
- `TEXT` on Windows ja-JP accepts an array of formats again (the `\` to `¥`
  rewrite applies per element instead of failing the call), and `[$-LCID]`
  tags render Windows' own month and day names (`[$-411]mmm` is `1`,
  `[$-404]mmmm` is `一月`). The system long-date tag drops the weekday on
  Windows ja-JP.
- `CODE` and `LENB` on Windows ja-JP treat Hangul as the single CP932 byte `?`
  (`CODE("한")` is 63, `LENB("한글")` is 2).
- Text collation composes a base letter with following combining marks
  (U+0300–U+036F) before comparing, so `"a"&UNICHAR(769)` equals `"á"` in
  comparisons, sorting and lookups, as on Mac and Windows Excel.
- `^` and `POWER` reproduce Excel's last bits: a fractional exponent is
  computed as `exp(log(x)*y)` (so `8^(1/3)` is 1.9999999999999998, as in
  Excel, not 2), 0.5 is a square root, a negative exponent is the reciprocal
  of the positive power, and an integer exponent is square-and-multiply
  (`1.0001^1000` is 1.105165392603197).
- `FILTERXML` returns `#VALUE!` for XML with text outside the root element or
  more than one root element, instead of reading past it.
- `BETA.INV` above the median solves on the upper tail, where `1 - p` is
  exact, so it lands on the true inverse instead of the edge of a flat CDF
  region (`BETA.INV(BETA.DIST(0.9,0.1,10,TRUE),0.1,10)` now agrees with Excel
  to within 1e-6, where it was off by 1.15e-6).

### Performance

- Conditional-format and AutoFilter colour/icon evaluation prepare the rule
  plan once per range walk instead of rescanning the range per cell.
- Pivot aggregation works over record-index views instead of copying values
  per group.
- Loading and saving `.xlsx` / `.xlsb` parse `[Content_Types].xml` and each
  rels part once, look up passthrough parts by path, and prune orphan parts
  without rescanning the part graph.
- `TEXTSPLIT`, `TEXTBEFORE` / `TEXTAFTER` and wildcard `XLOOKUP` / `XMATCH`
  prepare their delimiters and patterns once per call, and function-name
  lookup no longer allocates.

## [0.15.1] - 2026-10-12

### Changed

- `fm_workbook_cell_count` / `cellCount` and `fm_workbook_cell_at` / `cellAt`
  / `iter_cells` count and enumerate only the coordinates that were written,
  plus spill phantoms. A row holding cells far apart, such as A and XFD, no
  longer stores or counts the empty columns between them, so 50,000 such rows
  report 100,000 cells instead of aborting in WASM. `insertCols` no longer
  adds blank cells to the count, and `cellAt` no longer returns a blank cell
  for a gap. Indices stay dense in `[0, count)` and ordered by `(row, col)`.

### Fixed

- Saving is deterministic: every zip entry of a saved `.xlsx` / `.xlsb` carries
  the fixed timestamp 1980-01-01 00:00 instead of the save time, so an
  unchanged workbook saves to identical bytes.
- `fm_sheet_range_style_bases` / `rangeStyleBases` / `range_style_bases`
  return only the XFs some cell of the selection resolves to. A selection
  wholly covered by cells with their own style no longer reports XF 0, and a
  column style shadowed by row styles everywhere is no longer listed.
- A `.xlsx` sheet whose rows hold styled cells far apart no longer fails to
  load as too large. The per-sheet load budget charges each stored cell, which
  admits about 2.98 million cells per sheet.

### Performance

- Saving to `.xlsx` / `.xlsb`, recalculation indexing, reference rewriting and
  row or column edits walk only a row's stored cells, so a row with cells far
  apart no longer costs its full column width.

## [0.15.0] - 2026-10-12

### BREAKING

- C ABI: `fm_locale_facts_t` gains a trailing `collation_measured` field, so
  the struct grows by 4 bytes. Callers built against the old layout must
  rebuild.

### Added

- Row and column styles can be set and cleared (`fm_sheet_set_column_style`,
  `fm_sheet_clear_column_style`, `fm_sheet_set_row_style`,
  `fm_sheet_clear_row_style`; `setColumnStyle` / `clearColumnStyle` /
  `setRowStyle` / `clearRowStyle` on WASM and Node, `set_column_style` and
  siblings in Python). Both take an inclusive span of columns or rows.
  Width, height, hidden and outline state are kept.
- Attribute-only formatting of a range, row, column or whole-sheet selection,
  as Excel applies it: `fm_sheet_range_style_bases` / `rangeStyleBases` /
  `range_style_bases` list the XFs the selection resolves to, and
  `fm_sheet_remap_range_style` / `remapRangeStyle` / `remap_range_style` move
  each to its patched XF. Styles of covered rows and columns are rewritten,
  existing cells are remapped, and a cell is created only where a styled row
  or column crossing the selection would otherwise win; blank cells of a
  covered row or column are not enumerated.
  `fm_sheet_snapshot_range_style` / `fm_sheet_restore_range_style` (and their
  binding forms) capture and restore a selection's style state for undo.
- A cell's direct style can be removed (`fm_cell_clear_xf_index` /
  `clearCellXfIndex` / `clear_cell_xf_index`), so a blank cell inherits its
  row or column style again, unlike assigning xf 0, which pins Normal. Value,
  formula and comments are kept; the effective style's `source` tells the two
  states apart.
- Text comparison and value classification under a profile's collation:
  `fm_text_compare` / `compareText` / `compare_text` return the sign of
  comparing two texts, and `fm_collation_classes` / `collationClasses` /
  `collation_classes` give each value a class id, dense from 0 in ascending
  `SORT` order (numbers, text, booleans, errors by code, blanks). The strength
  (`Compare`, `Dynamic`, `ClassicExact`) selects the text fold of the
  matching formula family; a number, its text form and a boolean never share
  a class. Ids are valid within one call.
- `localeFacts` / `locale_facts` report `collationMeasured`: true for `mac-*`
  profiles, false for the estimated `win-*` collation (including
  `win-365-ja_JP`, whose other facts are measured).
- New workbooks shaped for a profile: `fm_workbook_create_for_profile` /
  `Workbook.createForProfile` / `Workbook.create_for_profile` take a profile
  id and return the workbook Excel creates in that locale: the body font
  (name, charset, the theme's minor scheme, 12 pt on Mac and 11 pt on
  Windows), the phonetic font and vertically centred Normal style of the
  Japanese, Korean and Chinese locales, the localized Normal style and first
  sheet names, page margins, and on Mac the default row height and base
  column width. Mac values are measured in all fourteen UI locales; `win-*`
  profiles are estimated from them. `createDefault` is unchanged.
- The theme's per-script fonts: `fm_workbook_theme_script_font_count` /
  `fm_workbook_theme_script_font_at` and `getTheme().majorScripts` /
  `minorScripts` (`major_scripts` / `minor_scripts` in Python) list the
  `<a:font script>` faces of the major and minor fonts, the Office theme's 47
  when the workbook has no theme part, and `fm_workbook_set_theme_script_font`
  / `setThemeScriptFont` / `set_theme_script_font` set one.

### Changed

- The WASM size report's Brotli soft ceiling moved to 1040 KiB; the hard
  ceiling stays at 1056 KiB.

### Performance

- Row entries are kept in row order, so effective-style lookups and `.xlsb`
  saves stay fast on sheets with many styled rows (100,000 rows: `.xlsb` save
  from 2.4 s to 83 ms).
- `UNIQUE` hashes each row's (or column's) identity instead of comparing
  every pair, so a 100,000-row text column takes milliseconds.

### Fixed

- A number, boolean, error, text or formula written into a cell without its
  own style takes its row's style, else its column's, as Excel stores it.
  A cell with an explicit style, including style 0, keeps it.
- `SEARCH` and `SEARCHB` treat two characters as equal when the profile's
  comparison collation does: letter case everywhere, and width and kana in
  every locale except zh-CN and zh-TW.
- Conditional-format text rules follow the formulas Excel stores for them:
  "contains" behaves as `SEARCH` (wildcards included) and "begins with" /
  "ends with" compare under the profile's collation. A "between" rule
  accepts its bounds in either order.
- `TEXTJOIN` over an operator applied to a whole column (`A:A&""`) counts the
  empty cells below the data, so a delimiter-joined result past 32,767
  characters is `#CALC!` as in Excel.
- `formatValue` and `validateValue` on WASM and Node reject a value without
  `kind` with an invalid-argument status instead of reading it as blank.
- The declarations no longer type unrelated `kind` fields as `FilterKind`:
  `ExternalLinkRecord.kind` is `ExternalLinkKind` and `ColorSpec.kind` a
  number in both packages, and the Node package's `Value.kind`,
  `PivotCell.kind` and `CfMatch.kind` use their own kind enums.
- An AutoFilter checklist entry ignores leading and trailing spaces on both
  the entry and the cell's shown text, so `abc` also selects `abc ` and
  ` abc`. Tabs, non-breaking and ideographic spaces still count.
- PivotTable caption filters "begins with", "ends with" and "contains" (and
  their negations) match like `SEARCH`: letter case, width and kana fold one
  character at a time under the profile's comparison collation, `*` and `?`
  are wildcards and `~` escapes them. On load the stored wildcard pattern is
  what is matched, as Excel does on refresh. Label filters set through the
  API use the same rule.
- An in-cell data-validation list compares what the cell shows: leading and
  trailing spaces are ignored, a spelling must match exactly (case and width
  included), and a value that reads as a number matches an item equal to it
  as a number (`01`, `１`, `1e0` against `1`; a percentage against `50%`). The
  space after each comma in the list is dropped. Lists that refer to a range
  are unchanged.
- Print column widths find the Normal style by its built-in id, so a
  workbook whose Normal style has a localized name sizes columns from that
  style's font.

## [0.14.0] - 2026-10-11

### BREAKING

- A new workbook evaluates with the `win-365-en_US` profile instead of
  `win-365-ja_JP`, in the C ABI, Node, WASM, Python and the CLI. Results
  that depend on the locale change with it: number, date and boolean text,
  `TEXT` and `DOLLAR`, the `CHAR` / `CODE` code page, the byte-counting
  `*B` functions, `ASC` / `DBCS`, kana and width folding, and PivotTable
  labels. Select `win-365-ja_JP` by profile id to keep the previous
  behaviour.
- `fm_locale_t` and the `FM_LOCALE_*` codes are removed; the function-name
  calls take a profile id instead of a numeric locale. `fm_function_metadata`,
  `functionMetadata` and `function_metadata` lose their locale argument, as
  the metadata is locale-invariant.
- C ABI: `fm_sheet_format_defaults` grows from 32 to 48 bytes
  (`effective_default_col_width`, `effective_default_row_height`). Callers
  built against the old layout must rebuild.
- Formula text that does not parse is rejected at entry:
  `fm_workbook_set_formula`, `fm_workbook_set_defined_name(_scoped)`,
  `fm_sheet_cf_add_rule` and `fm_sheet_add_validation` return
  `kParserUnexpectedToken` and leave the workbook unchanged, and the Node,
  WASM and Python setters over them do the same. A file that holds such a
  formula still loads it verbatim; it evaluates to `#NAME?`.

### Added

- `--profile <id>` on every CLI subcommand selects the Excel profile, from
  the same 28 ids the bindings accept.
- `fm_sheet_clear_column_width` (`clearColumnWidth`, `clear_column_width`)
  drops only the width override of a column range, keeping hidden, outline
  and style state, so the columns follow later default-width changes.
  `fm_sheet_get_format_defaults` reports the effective default column width
  and row height the sheet lays out with (`effectiveDefaultColWidth`,
  `effectiveDefaultRowHeight`).

- `de-DE`, `fr-FR`, `zh-CN`, `ko-KR`, `th-TH`, `es-ES`, `es-MX`, `pt-BR`,
  `ru-RU`, `zh-TW`, `it-IT` and `nl-NL` Excel profiles, each on a `mac` and a
  `win` host (`{mac,win}-365-{de_DE,fr_FR,zh_CN,ko_KR,th_TH,es_ES,es_MX,pt_BR,
  ru_RU,zh_TW,it_IT,nl_NL}`), selectable through the profile-id API in the C
  ABI, Node, WASM and Python. zh-TW counts and converts double-byte text in
  Big5, and ru-RU reads `CHAR` / `CODE` in Mac Cyrillic.
  The `mac-*` profiles follow Mac Excel 365 in that locale for number, date
  and boolean text, `TEXT` format codes, R1C1 letters, `CELL` codes,
  `FORMULATEXT`, double-byte text functions and `CHAR` / `CODE`; the `win-*`
  profiles are estimated from them plus the Windows host rules.
  Formulas are still entered with English names and stored separators, and
  CLI output stays locale-invariant.
- `--now <date[Thh:mm:ss]>` on every CLI subcommand pins the clock `NOW`,
  `TODAY` and relative pivot date filters read, as the pinned-clock API
  already does in the C ABI, Node, WASM and Python.
- Locale-dependent catalog calls are keyed by profile id in the C ABI, Node,
  WASM and Python: function-name translation (`fm_function_localize`,
  `fm_function_canonicalize`; `localizeFunctionName`,
  `canonicalizeFunctionName`; `localize_function_name`,
  `canonicalize_function_name`) for `de-DE`, `fr-FR`, `es-ES`, `es-MX`,
  `pt-BR`, `ru-RU`, `it-IT` and `nl-NL`; whole-formula text
  conversion between the stored English form and a profile's spelling
  (`fm_formula_localize`, `fm_formula_canonicalize`, `localizeFormula`,
  `canonicalizeFormula`, `localize_formula`, `canonicalize_formula`); and
  per-profile locale facts (`fm_locale_facts`, `fm_locale_error_name`,
  `localeFacts`, `locale_facts`): separators, boolean names, date order,
  currency layout, localized error names, and whether the profile is
  Excel-measured or estimated.
- A built-in function name written bare is a function value, as a `LAMBDA`
  is: `=TYPE(SUM)` is 128, `=SUM` shows `#CALC!`, `=ABS+1` is `#VALUE!`, and
  `=LET(f,ABS,f(-2))`, `=CHOOSE(1,SUM,ABS)(5)` and `=IF(TRUE,ABS,SUM)(-3)`
  call it. A defined name or LET binding of the same spelling still wins.
  XLSX and XLSB store it as `_xleta.SUM`, read and written.
- The trim-reference operators `A1:.A10` (trailing blanks), `A1.:A10`
  (leading) and `A1.:.A10` (both), over bounded ranges, whole columns and
  rows, 2-D and sheet-qualified ranges. XLSX and XLSB store them as
  `_xlfn._TRO_TRAILING(A1:A10)` and its siblings, and they read back in
  operator form.
- References to another workbook can be entered, displayed and evaluated:
  `=[Book.xlsx]Sheet1!A1`, `='/Users/x/[Book.xlsx]Sheet1'!A1`,
  `=SUM([Book.xlsx]Sheet1:Sheet2!A1)` and `=Book.xlsx!Name`, in cells, defined
  names, conditional formats and data validations. Entering one for a workbook
  that has no link yet creates the link, and XLSX and XLSB saves write it with
  its link part, so the formula survives a save and reload.
- Sheet names that look like a cell address (`S2`) or start with a digit
  (`2024`) are accepted unquoted before `!`, as Excel does.
- `Src2!Name`, where no sheet is called `Src2` and the linked book `Src2`
  has no file extension, binds to that book's name. It is written as
  `[N]!Name` in XLSX and XLSB saves.
- An external name that the link part lists without a body (no `refersTo` in
  XLSX, a zero-length formula in XLSB) reads as `#NAME?` and is written back
  the same way.
- XLSB DDE and OLE links are recognised by their link type instead of being
  read as supporting books.
- An external whole column or row has its declared size: `ROWS`, `COLUMNS`,
  `INDEX` and the lookup functions measure `[Book]Sheet!A:A` as the whole
  column, and one entered as the formula spills at that size, with blank
  cells past the cached extent. Aggregates still read the cached extent.
- Pictures can be moved and resized in place (`setImageAnchor`,
  `set_image_anchor`, `fm_sheet_set_image_anchor`) from the values
  `listDrawingObjects` reports. The object id, name, crop, rotation, flips and
  every other element are kept. A picture rotated a quarter turn gets an
  anchor box with width and height swapped about the same centre, and setting
  the values already in place writes nothing.
- The stacking order of pictures is settable: `setImageZOrder` /
  `set_image_z_order` / `fm_sheet_set_image_z_order` makes a picture item
  `index` of `listDrawingObjects`, which lists objects back to front.
- A picture can be captured with `snapshotImage` / `snapshot_image` /
  `fm_sheet_snapshot_image` and put back with `restoreImage(sheet, bytes,
  {newId})` / `restore_image(sheet, data, new_id=False)` /
  `fm_sheet_restore_image`. The capture holds the picture's element, its
  media and its position in the stacking order. Restoring replaces a picture
  with the same id or returns a deleted one at its place; `newId` adds a copy
  with a fresh id on top.
- `resetTheme` / `reset_theme` / `fm_workbook_reset_theme` returns a workbook
  to having no theme part, so `getTheme` reports the default theme again
  after `setThemeColors` or `setThemeFonts`. Excel adds a theme part again
  when it saves the file.

### Changed

- Text comparison on `mac-*` profiles follows Mac Excel's collation for the
  profile's locale: comparison operators, `SWITCH`, `COUNTIF`-family and
  D-function criteria, `MATCH` / `VLOOKUP` / `HLOOKUP` / `LOOKUP`,
  `XLOOKUP` / `XMATCH`, `SORT` / `SORTBY` / `UNIQUE`, `GROUPBY` / `PIVOTBY`,
  PivotTable item grouping, ordering and label filters, conditional formatting
  and AutoFilter. Case, width and kana type are ignored where Excel ignores
  them, accents order after base letters, `ß` equals `ss`, Hangul compares by
  jamo, and ideographs follow the locale's order (JIS in ja-JP, readings in
  ko-KR, pinyin in zh-CN, strokes in zh-TW). Lookups and criteria compare at
  the strength Excel uses for each function, so `XLOOKUP` matches `Ａ` to `A`
  in zh-CN where `MATCH(…, 1)` does not. `win-*` profiles keep the
  case-insensitive byte comparison, with unsigned bytes on every path, and
  `GROUPBY`, `PIVOTBY` and PivotTables on them no longer fold kana or width.
- Text functions follow the profile's code page. `CHAR` / `CODE` use Mac
  Roman or Windows-1252 on single-byte profiles, CP932 / JIS X 0208 on
  ja-JP, and the EUC forms of GB 2312 and KS X 1001 on zh-CN and ko-KR.
  Outside the double-byte profiles the `*B` functions count UTF-16 units,
  `ASC` / `DBCS` return their text unchanged and `TRIM` keeps U+3000.
  Half-width katakana is one byte only in ja-JP. Windows-1252's undefined
  slots decode to the C1 control of the same number, and on ja-JP a lone
  CP932 lead byte is a space.
- Number, date and boolean text is read and written by the profile's locale:
  `VALUE`, `NUMBERVALUE`, implicit coercion and criteria take the locale
  separators, currency and boolean names; numeric dates follow the locale
  date order, and kanji, wareki, full-width digits and yen markers are
  accepted only on Japanese profiles. `DOLLAR`, `FIXED` and `USDOLLAR` use
  the locale currency shape, number formats take weekday and month names
  from the locale, and the `!` escape, `DBNum` and the Japanese General
  spelling apply only on Japanese profiles.
- Exact `MATCH` / `VLOOKUP` / `HLOOKUP` / `XLOOKUP` / `XMATCH` on `mac-*`
  profiles fold full-width digits, and a D-function field argument folds
  half-width katakana, as Mac Excel does. `PHONETIC` returns `#N/A` in
  en-US.
- PivotTable labels, data-field names and subtotal rows follow Mac Excel 365
  in each profile's locale. An en-US pivot renders Excel's compact layout
  (`Row Labels` / `Column Labels` headers, `North Total` subtotal rows) instead
  of field-name headers, and names a Count Numbers, StdDevp or Varp data field
  `Count of`, `StdDevp of` or `Varp of`; a fr-FR subtotal row reads
  `Total North`.
- The WASM size report's Brotli ceilings moved to 1024 KiB soft and 1056 KiB
  hard; the uncompressed ceilings are unchanged.
- A workbook without a theme part now resolves theme colors and fonts with the
  current Excel Office theme (Aptos Display / Aptos Narrow, accent1 `156082`),
  as Excel does, instead of the Office 2013-2022 theme. A generated theme part
  carries the same values.
- Formulas loaded with cross-workbook references read back with the book name
  (`[Book.xlsx]Sheet1!A1`, or the quoted path form when the link has an
  absolute path) instead of `[N]`, and are written back with the original
  index.
- `[Book]Name` without a `!` is read as a name, so it evaluates to `#NAME?`
  when the book has no such name.
- A 3-D reference (`Sheet1:Sheet3!A1`) read as a value is `#REF!`, local or
  external. A selected 3-D arm of `IF` or `CHOOSE` gives `#VALUE!`, and
  `ROWS`-style shape functions and `INDEX` reject one with `#VALUE!`.
- The size of a picture in `listDrawingObjects` is read from its unrotated
  `a:xfrm` extent first, then from the anchor's `xdr:ext`; the two differ for
  a rotated one-cell picture or a part whose extents disagree.
- A typed `[1]` names a book called `1`; it no longer addresses the first link.
- `Sheet !A1`, with a space before `!`, is rejected.
- Sheet names shaped like `R`, `C` or `R1C1` are quoted in formula text.
- The XLSB writer encodes cross-workbook references as formulas instead of
  keeping only their cached values.
- A whole column or row inside an operator, a scalar function, `LET` or a
  defined name keeps its declared size (1,048,576 rows or 16,384 columns) with
  blank cells past the data, as Excel does, instead of being reduced to the
  cell in the formula's row or column: `SUM((A:A="x")*B:B)` and
  `SUMPRODUCT((A:A>1)*B:B)` sum the matching rows, `ROWS(B:B+0)` is 1048576,
  and `=B:B+0` spills the whole column. Only the populated part is stored.
- `COUNTBLANK` counts the empty cells of the whole declared range, and
  `COUNTIF` / `COUNTIFS` with a criterion that matches blank count the cells
  past the used range.
- `COUNTBLANK` and the `SUMIF`, `COUNTIF`, `AVERAGEIF`, `MAXIFS` and `MINIFS`
  families return `#VALUE!` for a range in another workbook, as Excel does
  while that workbook is closed. A cell missing from a linked workbook's
  cached values reads as blank instead of 0.
- A conditional-format or data-validation formula whose result is an array
  uses the array's first element for every cell it applies to. The ad-hoc
  rule probe (`fm_workbook_evaluate_cf_formula`) decides the same way, also
  where the array could not spill from the target cell.
- Text past the 32,767-character limit follows Excel function by function:
  `CONCAT` and `TEXTJOIN` return `#CALC!` (`TEXTJOIN` used to return
  `#VALUE!`), and `&` and `CONCATENATE` cut the result to the limit instead
  of returning longer text. `TEXTJOIN` with empty cells kept counts every
  blank row of a whole column, so `TEXTJOIN(",",FALSE,A:A)` is `#CALC!`.
- `ISREF` is TRUE for a cell, range or whole column in another workbook,
  written directly or bound by `LET`.
- Row and column inserts and deletes move the cells named by worksheet
  content kept for a lossless save, as Excel moves them: protected ranges,
  scenario input cells, a sheet sort state, cell watches, ignored-error
  ranges and web-publish ranges, and in XLSB the AutoFilter and manual page
  breaks. A watch on a deleted cell keeps its address, an insert inside an
  ignored-error range leaves the new cells out of it, and a sort state goes
  with the last of its sort conditions, the AutoFilter's included.
- An integer argument within 2^-22 below a whole number counts as that
  number in the functions where Excel does so, among them `DATE`, `TIME`,
  `INDEX`, `VLOOKUP` / `HLOOKUP`, the `D*` field, `LEFT` / `RIGHT`, `FIND`,
  `ROUND`, `FACT`, `LARGE`, `SEQUENCE`, `ADDRESS` and `WEEKDAY`; `MID`,
  `REPT`, `CHOOSE`, `SMALL` and the other plain truncations are unchanged.
- The former Analysis ToolPak functions (engineering, `IM*`, bond and
  coupon, `EDATE`, `EOMONTH`, `WORKDAY`, `NETWORKDAYS`, `YEARFRAC`,
  `WEEKNUM`, `QUOTIENT`, `GCD`, `MROUND` and the rest of the set) return
  `#VALUE!` for a boolean argument and `#N/A` for an omitted required one.
- An omitted argument counts as 0 in `COUNT`, `COUNTA`, `MEDIAN`, `STDEV`,
  `VAR`, `KURT`, `SKEW`, `AVEDEV`, `DEVSQ`, `GEOMEAN` and `HARMEAN`;
  `PRODUCT` skips it and `MODE` is `#VALUE!`. Direct booleans and numeric
  text count in `KURT`, `SKEW`, `GEOMEAN`, `HARMEAN`, `MEDIAN`, `AVEDEV`
  and `DEVSQ`, as in `STDEV`.
- Dates: serial 0 is 1900-01-00 in `EOMONTH` and `EDATE`; `DAY`, `MONTH`,
  `YEAR` and `WEEKDAY` round to the nearest second first; `DATE` bounds the
  month and saturates the day as Excel does; the `.INTL` weekend argument
  takes a boolean, blank or text as Excel does; and time text accepts the
  hour-only `6 PM` form and a trailing `.` after AM/PM.
- Statistics: `T.DIST` density at 0 degrees of freedom is `#DIV/0!`;
  `COVARIANCE.P`, `COVAR` and `COVARIANCE.S` take scalars; `GAMMA.DIST`
  density at 0 follows alpha; `CHISQ.TEST` drops non-numeric pairs and
  allows negative values; `BINOM.INV` rejects a probability of 0 or 1; and
  `GAMMA.INV`, `CHISQ.DIST.RT`, `CHISQ.INV` and `T.INV.2T` are accurate at
  extreme arguments.
- `CHOOSECOLS` and `CHOOSEROWS` take an array of indexes; `COUNT` counts
  across a 3-D reference and `COUNTIF` on one is `#VALUE!`; `CELL` with
  `@range` reads the intersected cell; `ARABIC` follows Excel's subtractive
  rule (`IIM` is 998); and `TOCOL` / `TOROW` fillers are empty text.
- `GROUPBY` and `PIVOTBY` default to a grand total, detected headers and an
  ascending sort, and their empty label cells are empty text.
- `<sheetFormatPr>` keeps `customHeight` and its other attributes through
  an XLSX or XLSB save.
- A result that cancels to within 2^-50 of its larger operand reads as 0
  where Excel snaps it: the unparenthesised root `+` / `-` of a formula
  (`=0.5-0.4-0.1` is 0, cell by cell for an array result) and each addition
  in `SUM`, `AVERAGE`, `AVERAGEA`, `SUBTOTAL` and `AGGREGATE`. A
  parenthesised or wrapped expression, `SUMPRODUCT`, the `*IF(S)` family
  and the `D*` functions keep the residue.
- A negative base raised to the reciprocal of an odd integer is an odd root
  (`(-8)^(1/3)` is -1.9999999999999998) in `^`, `POWER` and `SERIESSUM`,
  and `SERIESSUM` follows the operator's `0^0` and `0^-n` errors. `GCD`,
  `LCM` and `MULTINOMIAL` reject a boolean with `#VALUE!`.
- The single-value argument of `LARGE`, `SMALL`, `PERCENTILE`, `QUARTILE`
  (all variants) and `TRIMMEAN`, and `PERCENTRANK`'s `x` and
  `significance`, coerce a cell reference as a direct argument (a blank is
  0, `TRUE` is 1, numeric text is its number) and evaluate per element when
  given an array or range. `PERCENTILE.EXC` and `QUARTILE.EXC` accept the
  last position (`PERCENTILE.EXC({1,2,3},0.75)` is 3). `PERCENTRANK`
  between duplicates, `PROB` with blank limits, `FISHERINV` and `T.INV` at
  extreme arguments follow Excel.
- `AGGREGATE` functions 1 to 13 take references only, so a literal or an
  array argument is `#VALUE!`; a union, a LET-bound range and a
  reference-returning call are references, and an array `k` in functions
  14 to 19 evaluates per element.
- `XMATCH` and `XLOOKUP` `match_mode` 3 is an unanchored, case-sensitive
  regex search. `INDIRECT` wraps a relative R1C1 offset around the sheet.
  Numeric text above 9.99999999999999E+307 is `#VALUE!`. In the numeric and
  A-family aggregates a non-numeric text literal outranks a later error, and
  an `IF` that picks a plain value is a direct argument.
- `DAYS360` rounds its dates to the second and reads serial 0 as
  1900-01-00; `WORKDAY` and `WORKDAY.INTL` floor a fractional day count;
  `NETWORKDAYS.INTL` counts 0 for the all-weekend mask.
- `ACCRINT` counts periods on the quasi-coupon grid anchored on
  `first_interest`, and `ODDFPRICE`, `ODDFYIELD`, `ODDLPRICE` and
  `ODDLYIELD` measure each quasi-coupon period at its own length.
- A pivot with several data fields lays out the Values level where its
  file places it: on the rows, first on its axis, or between column fields.
  Each data field gets its own grand-total row or column in the locale's
  wording, a group above the Values level is subtotalled once per data
  field, and the tabular and outline headers show the stored Values
  caption. The report form is read from the row fields, where Excel
  records it, and written back in the attribute combinations Excel saves.
- `TEXT` format codes follow more of Mac Excel 365's rules in every
  profile. An unescaped `n` is `#VALUE!`, and so is a `/` that is neither a
  date separator nor a fraction bar: one with nothing on a side, or a letter
  that is no date code on its left (`x/x`, or `yyyy/m/d` where `y` and `d`
  are text). `r` is the era year in ja-JP (`rr` adds the era name) and the
  year in zh-TW, and `t` writes the digits in Thai in th-TH.
- `[DBNum1]`-`[DBNum4]` write numbers the way Mac Excel 365 does in ja-JP,
  zh-CN, ko-KR and zh-TW. General, elapsed time and the date fields other
  than the full year use place units in each locale's style: ja-JP drops
  the 1 before 十百千 (`百一`), zh-CN and zh-TW write it and a zero for
  skipped places (`一百○一`), ko-KR writes it without the zero, and
  `[DBNum4]` is Hangul in ko-KR. Digit placeholders and full years stay
  digit by digit, with each locale's digits (`[DBNum2]` 0 is `〇` in ja-JP,
  `零` elsewhere). `[DBNum4]` is accepted, and changes nothing outside ko-KR.
- `[$-…]` locale tags in number formats now take effect the way Mac Excel
  365 applies them, for every section of the format:
  - the language of the first section's tag (`[$-411]`, `[$-de-DE]`, `[$-ko]`)
    switches month and day names, AM/PM and `[DBNumN]` styles to that
    language's own spellings for the fourteen profile languages and the
    regional LCIDs that share them; any other language keeps the profile's
    names, and the Japanese era is available only under a Japanese tag;
  - calendars `03` (Japanese era years), `05` (Dangi), `07` (Thai Buddhist),
    `04` (Taiwan), `02` (English names), `0D` and `15` change years and
    names, and `ja-JP-x-gannen` writes the first era year as 元;
  - numeral systems `02`-`13` substitute the Unicode digits of Arabic-Indic,
    Devanagari, Thai, Lao, Khmer and the other scripts, and `1B`-`27` write
    the ja-JP, zh-CN, zh-TW and ko-KR DBNum styles;
  - `[$-F800]` / `[$-x-sysdate]` and `[$-F400]` / `[$-x-systime]` show the
    system long date and time; any tag makes `b` the Gregorian year, a
    second tag in a section or a malformed one is `#VALUE!`, and a text
    section writes a tag's currency symbol.

  Hijri, Hebrew, Persian, Saka and lunisolar calendars, and languages
  outside the profiles (such as Arabic or en-GB), are not modelled.
- `TEXT` accepts every colour name and the indexed-colour prefix in each
  profile's spelling (`[Blau]`, `[Farbe3]`, `[แดง]`), and `ARRAYTOTEXT`
  writes errors in each profile's spelling, including `#SPILL!` and `#CALC!`
  (`#スピル!` in ja-JP, `#ÜBERLAUF!` in de-DE, `#BEREKENEN!` in nl-NL).

### Performance

- The WASM module no longer links the 128-bit floating-point routines, and
  several single-use template instantiations are folded into shared code.

### Fixed

- `IFERROR` and `IFNA` replace the caught elements of an array or range
  value one by one, pairing an array fallback by position, instead of
  passing the array through unchanged.
- `NUMBERVALUE` drops spaces anywhere in its text (`"1 234.5"` is 1234.5).
- Date and time text accepts spaces around `-`, `/` and `:`
  (`DATEVALUE("2026 - 01 - 01")`), and an hour of 24 or more rolls the date
  forward.
- A `"TRUE"` / `"FALSE"` `COUNTIF`-family criterion matches only boolean
  cells, and `GCD` / `LCM` skip a scalar reference to a blank cell.
- `SUM(INDIRECT("A2"))` and the other range-aware functions skip text in a
  single-cell `INDIRECT` as they do for `SUM(A2)`.
- A PivotTable with axis fields and no data field renders labels only,
  without empty value columns or a grand-total row.
- Fraction formats (`TEXT(0.015, "?/100")`) give the same result on WASM and
  x86-64 Linux as on macOS.
- `TEXT` returns `#VALUE!` for a text section holding an unquoted digit
  (`"@ 9"`), as Excel does.
- `TEXT` returns `#VALUE!` for a number format with an unquoted colon before
  a digit placeholder (`"0:0"`), as Excel does.

### Known limitations

- Collation on `mac-*` profiles does not compose a letter with a following
  combining mark (`a` + U+0301 is not `á`), places scripts outside Latin,
  Greek, Cyrillic, Thai, Hangul, kana and CJK after them in code-point order
  (the zh profiles ignore them), and orders zh-CN ideographs outside GB 2312
  level 1 by radical and stroke instead of pinyin. Wildcard criteria and
  wildcard lookups keep the previous kana and width folding, and Windows text
  order is unmeasured.

## [0.13.0] - 2026-10-06

### BREAKING

- C ABI: `fm_cell_xf` grows from 88 to 128 bytes (apply flags, `quotePrefix`
  and protection) and `fm_row_layout_t` from 32 to 40 bytes (`has_height`,
  `custom_height`). Callers built against the old layouts must rebuild.
- C ABI: `fm_styles_set_cell_style` takes one `fm_cell_style_record_t` (name,
  `xf_id`, `builtin_id`, `i_level`, `hidden`, `custom_builtin`). The JS
  `setCellStyle` and Python `set_cell_style` take the matching record.
- The native Node addon and WASM validate arguments strictly before entering
  native code: a wrongly typed or out-of-range positional or nested argument
  throws `TypeError` / `RangeError` instead of being coerced. String fields
  take a string (or, on Node, a one-byte buffer), and 64-bit cursors are
  bounded by `Number.MAX_SAFE_INTEGER` on both surfaces.

### Added

Everything below is available through the C ABI, WASM, native Node and Python.

- Sheet geometry: default column width and row height can be read and set,
  and a row height cleared back to auto. Cell rectangles, column widths and
  row heights are reported in points in display or print mode, with the
  width model and chars/points conversion.
- Pagination results carry paper, margins, printable area, scale, page order,
  print titles, per-page layout and manual-break flags.
- Cells and formulas in a range are enumerated through a handle with a
  resumable cursor, and merges in a range are listed.
- A cell's formula in A1 and R1C1 notation, and its display text (or any
  value under a format code) with an `ok` / overflow / invalid status.
- The workbook theme: the 12 colours and the major and minor fonts can be
  read and set, including on a workbook with no theme part. Any colour spec
  (theme and tint, indexed, auto, RGB) resolves to ARGB for a font, fill or
  border context, and a cell's effective style reports its resolved fonts,
  fills and borders.
- Named cell styles can be removed, and built-in style ids 48..53 are
  accepted. `setCellStyle` is now available on the native Node addon.
- Typed AutoFilters for sheets and tables: filter columns (values, date
  groups, custom, top 10, dynamic, colour, icon) and sort state can be read,
  set, removed, cleared, evaluated to visible rows and applied. Applying a
  filter hides and unhides rows and recalculates `SUBTOTAL` and `AGGREGATE`.
  A typed set keeps `xr:uid`, `calendarType`, `extLst` and unknown
  attributes.
- Data-validation evaluation: whether a proposed value satisfies the rule
  that applies to a cell, which rule applied and its error style, plus a
  resumable listing of invalid cells.
- Threaded comments with replies, mentions, resolved state and the
  workbook's person list can be read, added, edited and removed. Each thread
  keeps Excel's legacy note stub in sync, and row and column edits move or
  remove threads with their anchors.
- Images: PNG, JPEG, GIF and BMP bytes are probed for format and pixel size,
  a sheet's drawing objects (pictures, charts, shapes, groups) are listed
  with anchors, and pictures can be read, inserted with a one-cell or
  two-cell anchor (an absolute anchor is rejected), and removed. Row and
  column edits move anchors the way Excel does.
- Enum constants mirroring the C enums on every surface: `GeometryMode`,
  `DisplayStatus`, `ColorContext`, `ColorResolution`, `ThemeSource`,
  `EffectiveStyleSource`, `FilterKind`, `FilterOperator`,
  `DynamicFilterType`, `SortBy`, `SortMethod`, `DateTimeGrouping`,
  `ValidationErrorStyle`, `ImageFormat`, `DrawingObjectKind`, `AnchorKind`
  and `AnchorEditAs`.

### Changed

- `SUBTOTAL` and `AGGREGATE` (options 0..3 and the array form 14..19) skip a
  referenced cell whose own formula calls `SUBTOTAL` or `AGGREGATE`,
  including inside `LET` / `LAMBDA` bodies and unevaluated `IF` branches,
  and every cell of a spilled or CSE array whose anchor qualifies. A name,
  `INDIRECT` or a plain reference to such a cell does not count. A
  `SUBTOTAL` whose cells are all excluded returns its empty-set value
  instead of `#VALUE!`.
- `SUBTOTAL` 1..11 treat hidden rows as filtered only when a sheet or table
  AutoFilter on that sheet has a criterion, and then skip every hidden row
  on the sheet. 101..111 keep skipping every hidden row.
- Rows with an auto height are no longer written with `customHeight`; only
  explicit heights are.
- `TEXT` renderings match Excel for fraction padding on whole numbers,
  `[DBNum1]` with General, serial 0 as 1900/1/0, negative 1904-system dates
  and times, and the 1904 upper serial bound.
- The WASM size report's ceilings moved to 3.50 MiB / 896 KiB Brotli soft
  and 3.75 MiB / 928 KiB Brotli hard, keeping the 0.25 MiB and 32 KiB gap
  between each soft and hard ceiling.
- Number formats are validated the way Excel validates them: more than four
  sections, a condition in the third section or later, `@` mixed with
  numeric or date tokens, and an unterminated quote, bracket, escape, `_`
  or `*` are rejected. Conditional sections take Excel's implicit sign
  arms, and a value no arm accepts is `#VALUE!` in `TEXT` and `#####` in
  display text. A text-only section shows numbers as General.
- Fractions are rendered by a continued-fraction search with fixed and
  zero-prefixed denominators and `%` applied first; scientific formats
  support engineering grouping, and commas before the decimal point scale.
  Dates render from one rounded second with midnight carry.
- A cell with an explicit style 0 is distinct from an unstyled cell and
  overrides its row or column style; the distinction round-trips through
  `.xlsx` and `.xlsb`. `CELL()` and display text read the effective
  cell, row or column format, and spilled cells are saved with their
  effective style.
- Formulas that call a sheet- or workbook-qualified built-in are rejected at
  every entry point (cells, conditional formats, defined names, validation),
  not only `set_cell_formula`.
- Pivot fields named ambiguously (a name shared by distinct fields) are
  rejected by filters and the C API without mutation, and `GETPIVOTDATA`
  returns `#REF!` for them. Value filters on an unknown field are rejected
  at authoring time.
- A row or column insert removes a pivot table pushed wholly off the grid;
  a `GETPIVOTDATA` reading it becomes `#REF!`.
- `TIME` caps each component at 32767 as Excel does, and the date and time
  functions return `#NUM!` for non-finite or out-of-range inputs.

### Performance

- The WASM module no longer links libc `printf`: numbers are formatted by a
  byte-identical in-tree formatter built on double-conversion.
- Loading a sheet with many dynamic-array anchors is no longer quadratic:
  committed spills are indexed by row band.
- Adding formulas under whole-column watchers no longer costs quadratic
  time during dirty marking.
- Duplicated function preludes and out-of-line paths are shared, shrinking
  the WASM module (`.wasm` 3,292,036 bytes raw, 901,702 bytes Brotli).

### Fixed

- Stored number formats with English colour names (`[Red]`, `[ColorN]`) no
  longer render as `########` in display text and `CELL`.
- `FIXED` and `DOLLAR` accept the full Excel decimals range.
- `RANDBETWEEN` and `RANDARRAY` stay finite and in range when the bounds are
  huge or span beyond the 64-bit integer range.
- Dates past month 12 were computed two days off in the civil-date
  conversion.
- ISO `t="d"` date cells are rebased onto the 1904 date system when the
  workbook uses it, and date-only ISO strings with a `Z` or `+hh:mm` suffix
  are accepted.
- Toggling the 1904 date system keeps the saved `<workbookPr>` in step with
  the model; package-root parts such as `customXml` get correct relationship
  targets in `.xlsx` and `.xlsb`; unknown relationship types are
  attribute-escaped.
- Formulas referencing a missing sheet recalculate when that sheet is added
  or renamed into existence. Changing iteration, the profile, the 1904 date
  system or the pinned clock dirties every formula (and clears pivot
  results where they depend on it). A partial recalc marks dependents
  outside its closure dirty, so a later full recalc refreshes them.
- Row and column edits rewrite defined names and unowned table and pivot
  cache ranges, dirty committed spill anchors, drop data validations and
  ranges pushed wholly off the grid, and no longer overflow drawing anchors
  near the grid edge. Removing or moving a sheet remaps the active and
  first visible tab.
- Pivot authoring edits dirty dependent formulas so `GETPIVOTDATA`
  refreshes; custom subtotals re-aggregate with their own function after a
  value filter; invalid date serials stay out of date groups.
- `INDIRECT` and `OFFSET` fill a clamped expansion shorter than its
  rectangle with blanks instead of uninitialised cells.
- Conditional-format theme, indexed and auto colours no longer render black:
  they resolve against the workbook theme and palette in the engine and in
  the C API payloads, and survive an XLSB save and load.
- Table AutoFilters follow row and column inserts and deletes, and
  AutoFilter `colId`s shift when columns are inserted or deleted inside the
  filter.

### Known limitations

- Saving as XLSB does not write threaded comments, inserted images or typed
  AutoFilter edits; each is reported as a diagnostic. AutoFilters loaded
  from XLSB are not visible to the typed API.
- Custom data-validation formulas and rule bounds evaluate against the
  stored workbook, not the proposed value.
- AutoFilter evaluation approximates Excel's text ordering for custom `<`
  and `>` text comparisons, and computes top N and average filters over all
  body rows.
- Display text does not depend on column width: narrow columns never show
  `####` and repeat fills contribute nothing.
- The legacy note text of a threaded comment and the regenerated comment VML
  geometry approximate Excel's.

## [0.12.0] - 2026-09-30

### Added

- `USDOLLAR`, which formats a number as dollar currency text with two
  decimals by default and negatives in parentheses (`($1,234.50)`). It
  shares its formatting with `DOLLAR`, which keeps the yen sign, zero
  default decimals and a leading minus.

- `.xlsb` workbooks keep their conditional formatting, data validation
  and protection. The XLSB reader decodes conditional formats (every rule
  type and operator, time periods, thresholds, RGB colours, icon sets and
  data bars including their `x14` extension), data validations, and sheet
  and workbook protection into the same model the `.xlsx` reader fills,
  and the writer emits them from that model, so a rule edited after load
  is saved rather than dropped. The differential formats (dxf) those
  rules use are written into `.xlsb`'s styles part, font name included,
  and an edited or new `x14` data bar is written from the model on both
  formats. Rule formulas are encoded against the rule's range the way
  Excel encodes them, so Excel opens the file with the rules intact.

- Conditional-format rules carry an icon set's floor (its first
  threshold, below which a cell gets no icon) and a data bar's direction.
  The C ABI appends `icon_set_floor_engaged`, `icon_set_floor` and
  `data_bar_direction` to `fm_cf_rule_t`, so earlier fields keep their
  offsets while the struct grows, and `fm_cf_match_t` reports
  `bar_direction` in a former padding byte. They surface as
  `iconSet.floor`, `dataBar.direction` and `barDirection` in npm and
  native Node, and as `IconSet.floor`, `DataBar.direction` and
  `CfMatch.bar_direction` in Python. Both formats read and write the
  floor and the direction instead of synthesising a percent-0 floor.

- Defined names work as range endpoints. `A1:MyName` composes with the
  reference the name's body yields, including sheet-qualified names and
  bodies that are ranges, other names or `INDEX` / `OFFSET` calls; a
  constant, expression or text body gives `#VALUE!` and an undefined name
  `#NAME?`. `INDEX`, `XLOOKUP`, `IFS` and `SWITCH` work as endpoints too
  (`A1:INDEX(...)`, `XLOOKUP(...):B3`), and `ROW` / `COLUMN` accept such a
  dynamic range. Endpoints on different sheets give `#VALUE!`.

- The self-book name form `[0]!Name` resolves, including `[0]!Fn(args)`
  calls, `A1:[0]!Rng` endpoints and intersections such as
  `[0]!Rng A1:A5`. It reads the workbook-scoped definition, falling back
  to the lowest-indexed sheet's local one, and round-trips through
  `.xlsb`.

### Changed

- **Breaking (npm):** `@libraz/formulon` now resolves to a
  single-threaded build that loads without cross-origin isolation, and
  the pthread build moves to `@libraz/formulon/threads`, with its binary
  at `@libraz/formulon/formulon_threads.wasm`. The previous default
  allocated its memory as a `SharedArrayBuffer` and spawned eight worker
  threads as soon as the factory ran, so a page or worker without
  COOP/COEP headers (an Electron `file://` page, most static hosts) never
  got a ready module, even when only the serial `recalc` was used. Both
  entries share one API and one `.d.ts`. In the default build
  `recalcParallel` still succeeds but evaluates serially and reports
  `workerThreadsStarted: 0`. A host that relies on parallel recalc
  changes its import from `@libraz/formulon` to
  `@libraz/formulon/threads` and keeps serving the COOP/COEP headers it
  already needed.

- **Breaking (npm and native Node):** the `Workbook` accessors that used
  to answer a rejected argument or a released handle with a plain zero,
  empty string, `null` or default value now carry the call's `Status`, so
  a failure no longer reads as a legitimate result. Every `*Count`
  accessor and `calcMode()` return `NumberResult` (`{ status, value }`);
  `excelProfileId()`, `localizeFunctionName()` and
  `canonicalizeFunctionName()` return `StringResult`; `pinnedNow()`
  returns `{ status, now }`; `precedents()`, `dependents()` and
  `functionNames()` return `ListResult`; and `SpillInfo` gains `status`.
  Read `.value` (or `.now`) where the bare value was used before. The
  Python binding already raised `FormulonError` for these cases and is
  unchanged.

- **Breaking (C ABI and every binding):** the pivot date grouping
  `FM_PIVOT_DATE_WEEK` / `PivotDateGrouping.Week` is replaced by
  `FM_PIVOT_DATE_DAYS` / `PivotDateGrouping.Days` (same ordinal, 4),
  Excel's "Days" grouping, in which a week is a 7-day interval.
  `fm_workbook_pivot_field_set_date_group` gains three trailing
  parameters: `uint32_t interval_days`, `double start_serial_or_neg1` and
  `double end_serial_or_neg1`. A C caller appends them, passing `7`,
  `-1.0`, `-1.0` for the old weekly grouping and `1`, `-1.0`, `-1.0` for
  any other granularity, where they are ignored. `pivotFieldSetDateGroup`
  (npm, native Node) and `pivot_field_set_date_group` (Python) take them
  as optional trailing arguments (`intervalDays` / `interval_days`
  defaulting to 1), so replacing `Week` with `Days` and passing an
  interval of 7 is the whole migration there. Buckets start at the
  field's data minimum unless Start is given, are labelled
  `yyyy/m/d - yyyy/m/d`, and records below Start or above End collapse
  into `<start` / `>end` buckets. The previous Sunday-aligned
  `YYYY-MM-DD` weeks matched no Excel grouping. A zero interval, or a
  Start after End, is rejected with `kInvalidArgument`.

- **Breaking (C ABI and every binding):** a pivot field's and a pivot
  data field's `number_format` is a `numFmtId` as a decimal string, either
  a built-in id or one returned by `addNumFmt` / `add_num_fmt`, and any
  other non-empty string is rejected with `kInvalidArgument`. Register a
  format code with `addNumFmt` first and pass the returned id. A pivot
  field's id now reaches the saved file as `<pivotField numFmtId>` and is
  read back from it; it was previously accepted and silently dropped.

- Weekday functions and weekday format tokens follow Excel's serial
  weekday for 1900-system serials 0 to 60, where serial 1 (1900-01-01) is
  a Sunday and the fictitious 1900-02-29 a Wednesday, instead of the
  proleptic Gregorian calendar. `WEEKDAY`, `NETWORKDAYS(.INTL)`,
  `WORKDAY(.INTL)`, `TEXT` `ddd`/`aaa` and conditional-format weekday
  checks now match Excel there, `WEEKNUM` and `ISOWEEKNUM` match across
  1900, and `WEEKDAY` of a blank cell is 7. Later serials and the 1904
  date system are unchanged.

- Text that spells a hexadecimal float (`"0x10"`), an infinity or a NaN
  no longer coerces to a number; `="0x10"+1`, `="inf"+1` and `="nan"+1`
  give `#VALUE!`, as in Excel, where they gave 17 and `#NUM!`. Decimal
  text parses exactly as before.

- `FORMULATEXT`, `SHEET` and `SHEETS` are volatile, as in Excel: their
  cells recalculate on every pass, `.xlsb` stores them with the volatile
  marker, and `.xlsx` marks their cells `ca="1"`.

- A sheet-qualified call to a built-in function (`=Sheet1!SUM(1)`) is
  rejected as a formula that fails to parse, as Excel refuses it at
  entry, and evaluates to `#NAME?`. A qualified defined name
  (`=Sheet1!MyFn(1)`) is still accepted. The formula-length limit is
  lowered to Excel's 8,192 characters.

- `getValue` / `get_value` (`fm_workbook_get_value`) rejects a row or
  column outside the sheet grid with `kInvalidArgument` instead of
  reading it as a blank cell, and `fm_workbook_set_number` rejects NaN
  and infinity instead of storing a value `ISNUMBER` accepts and a later
  save turns into `#NUM!`.

- A pivot item added by cache index, which carries no name of its own,
  derives its display label from the bound cache value at evaluation
  time. Hiding such an item now hides the value it displays, a sole
  visible page item is labelled by that value, and a show-as base item is
  matched against it.

- Printed page breaks size columns from the workbook's Normal-style font
  and size, using widths measured against Windows Excel for Calibri,
  游ゴシック, ＭＳ Ｐゴシック and Meiryo UI at 8 to 18 pt, instead of
  assuming Calibri 11 for every workbook.

- The WASM size report's ceilings moved to 3.00 MiB / 832 KiB Brotli soft
  and 3.25 MiB / 864 KiB Brotli hard. Shipped feature code had filled the
  previous margin, so the soft warning fired on every run; the gap
  between each soft and hard ceiling stays 0.25 MiB and 32 KiB.

- The WASM builds are produced with Emscripten 6.0.10, and the Python
  package requires wasmtime 49. The vendored C++ dependencies move to
  PCRE2 10.49, pugixml 1.16 and double-conversion 3.4.0.

### Removed

- **Breaking (C ABI and every binding):**
  `fm_workbook_pivot_field_add_aggregation` and
  `fm_workbook_pivot_field_clear_aggregations` are removed, with
  `pivotFieldAddAggregation` / `pivotFieldClearAggregations` (npm and
  native Node) and `pivot_field_add_aggregation` /
  `pivot_field_clear_aggregations` (Python). They wrote a field-level
  list that neither evaluation nor the saved file ever read, and Excel
  has no such concept. Drop the calls: a value's aggregation is set on
  its data field, and a field's subtotal functions with
  `pivotFieldAddSubtotalFn` / `pivot_field_add_subtotal_fn`.

- **Breaking (build):** the experimental bytecode compiler, optimizer and
  VM, together with the `FORMULON_BUILD_VM` and
  `FORMULON_VM_PARITY_CHECK` CMake options and the `kVm*` error codes
  (2050-2064). They never shipped in a release artifact; the tree-walker
  is the engine's only evaluator. A build script that sets either option
  drops it.

### Performance

- Loading a workbook with many formulas that watch whole columns is
  linear rather than quadratic: formula registration looks up the
  formula cells inside a watched rectangle through an ordered index
  instead of walking every populated row, and `.xlsx` literal cells are
  written straight to the sheet. 40,000 rows of `=ROWS(A:A)` load in
  about half a second.

- A range watched by many formulas is one node in the dependency graph,
  so N formulas watching a range that holds M formulas cost N+M edges
  instead of N×M. Changing a defined name or a table re-indexes only the
  formulas that reference it instead of rebuilding every formula's edges
  and clearing every spill in the workbook.

- `COUNTIF`-family criteria and lookup keys are normalized for comparison
  once per call instead of once per scanned cell, and `VLOOKUP`,
  `HLOOKUP` and `INDEX` over whole columns read only the cells they need.

- The WASM module parses decimal text through double-conversion instead
  of the C library's `strtod`, which drops the libc scanner and its
  128-bit float routines from the binary (about 10 KB uncompressed,
  5 KB Brotli).

### Fixed

- **Formula parsing and formula text:** a formula keeps the parenthesis
  pairs it was written with, in its formula text and in `.xlsb` bytes,
  instead of losing them on a round trip. A cell reference, a
  parenthesised range, union or intersection is accepted as the callee of
  a call (`=A1(1)`, `=Sheet1!A1(1)`, `=(A1,B1)(1)`), while
  `Sheet1!LOG10(100)` still calls the built-in. A parenthesised
  intersection or union is accepted as a range endpoint (`(A1 B1):C3`), a
  trailing space before `)` or `]` is no longer read as an intersection
  operator, a sheet named `TRUE` or `FALSE` is quoted so `Sheet!A1` does
  not tokenize as a boolean, a leading backslash is accepted in a defined
  name, a subnormal numeric literal flushes to zero, and `#REF!A1` and
  `Sheet1!#REF!` parse as a bare `#REF!`. Inside a structured-reference
  bracket an apostrophe escapes the next character (`Table1['#Items]`),
  and a doubled single quote in a column name is unescaped before it is
  matched against the header.

- **Defined names:** `Sheet!Name` resolves in that sheet's scope,
  sheet-local first then workbook, giving `#REF!` for an unknown sheet
  and `#NAME?` for an undefined name; a sheet rename rewrites the
  qualifier and a sheet removal collapses it to `#REF!`. `Sheet!Fn(args)`
  calls a sheet-scoped `LAMBDA` name. An unqualified name inside a
  sheet-local name's body, including a `LAMBDA` defined there, resolves in
  the owning sheet's scope rather than the caller's. Defining, renaming
  or removing a name re-indexes and recalculates formulas that call it as
  a function (`=Fn(3)`), not only ones that reference it.

- **References and implicit intersection:** `@` and `SINGLE` over a
  reference project it onto the formula's cell instead of taking its
  top-left value: over a spill reference (`@G1#`) they project the anchor
  plus its spill region, and over a defined name, intersection, `LAMBDA`
  call or reference-returning call (`INDEX`, `OFFSET`, `CHOOSE`,
  `INDIRECT`, `IF`, `IFS`, `SWITCH`, `XLOOKUP`) they project the
  reference it yields, giving `#VALUE!` outside it. Calling any reference
  shape as a function, including one held by a name or `LET` binding
  (`=(A1:A2)(1)`, `=LET(f,(A1,B1),f(1))`), gives `#REF!`. `LET` names and
  `LAMBDA` parameters bind as references whenever their source resolves
  to one, so `ISREF`, `ROW`, `CELL`, `OFFSET` and `AREAS` see the
  reference and `SUM` over a one-cell `INDEX` binding skips text;
  `MAP`, `BYROW`, `BYCOL`, `REDUCE` and `SCAN` pass each element of a
  reference source as a cell, row or column reference, and accept a bare
  built-in name as their function the way `GROUPBY` / `PIVOTBY` do.

- **`INDEX` and `OFFSET`:** `INDEX` on a two-dimensional reference with
  only a row argument returns `#REF!` for every row number, including 0;
  an array constant still spills its row, or for 0 the whole array.
  `INDEX`'s area-number form selects from a parenthesised union in source
  order, giving `#REF!` for an out-of-range area. `OFFSET` treats an empty
  height or width like an omitted one, keeping the base reference's size,
  and accepts a `LET`-bound or whole-row/whole-column reference as its
  base.

- **Dependencies and circular references:** a cycle is reported only
  when evaluation actually reads a cell of the cycle back into itself, so
  an untaken branch no longer makes a formula circular
  (`=IF(FALSE,A1,5)` in A1 gives 5; `=IF(TRUE,A1,5)` stays circular).
  Reading a reference only for its position creates no dependency, so
  `=ROW(A1)` in A1, `=ROWS(A:A)` down column A, `=LET(r,A1,ROW(r))` and
  `=ROW(INDEX(A1,1,1))` are no longer circular, and `CELL` depends on its
  reference's value only for the `contents` and `type` info types. A
  range ending in a reference-returning call depends only on the cells it
  reads. `OFFSET` and `INDIRECT` record the rectangles they resolve to,
  so a dependent recalculates when the target moves; a cycle closing
  through them is circular only while it points at the cell, and with
  iterative calculation off it restores every member to its pre-recalc
  value, so repeated recalcs no longer grow the cycle's values.
  A formula reading one cell of another formula's spill now depends on
  the spill's anchor (`=B2*10` next to `B1=SEQUENCE(2)` gave 0 and stayed
  stale). A `LAMBDA` passed to `MAP` and the other helpers contributes its
  body's references, so `MAP(A1,LAMBDA(x,x*B1))` recalculates when B1
  changes. A self-reference in ad-hoc evaluation (`evaluateFormulaText`)
  reads the cell's cached value instead of re-entering the iterative
  driver.

- **Recalculation after model edits:** formulas that read something the
  dependency graph cannot see are marked dirty when it changes: a
  phonetic-run, reading or cell-property write; a row's hidden flag (read
  by `SUBTOTAL` / `AGGREGATE`); an Excel-profile change; and a defined
  name, print area, print titles or table change.

- **Whole rows, whole columns and array shapes:** whole-row and
  whole-column references used together in one call share one walked
  length and pad the shorter ones with blanks, so uneven whole columns
  agree on shape in `SUMIFS` / `COUNTIFS`, `XLOOKUP` / `XMATCH` /
  `MATCH` / `LOOKUP` and `FILTER`, and a wide whole-column lookup no
  longer hits `#CALC!`. `XLOOKUP` over an empty column returns its
  `if_not_found` value. `WRAPCOLS` / `WRAPROWS` and the `GROUPBY` /
  `PIVOTBY` filter array judge vector-versus-2D from the reference's
  declared shape, so a multi-column reference populated in one row is not
  misread as a vector.

- **Conditional aggregates:** `COUNTIF`, `SUMIF`, `AVERAGEIF`,
  `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `MAXIFS` and `MINIFS` lift a range
  or array criterion and return one result per element
  (`COUNTIF(A1:A3,A1:A2)` is `{1;2}`), broadcasting several criteria the
  way operators broadcast, instead of returning 0. A position a shorter
  criterion array cannot supply uses `#N/A` as that criterion, as Excel
  does.

- **Text functions:** `TEXTBEFORE` / `TEXTAFTER` accept an empty
  delimiter, matching at the start of the text for a positive
  `instance_num` and at the end for a negative one. `TEXTJOIN` no longer
  expands a range or array delimiter into its argument list. `TEXT`,
  `FIXED` and `DOLLAR` print an integer beyond 15 significant digits
  without binary-representation noise, `DOLLAR` keeps the minus sign of a
  negative value that rounds to zero (`DOLLAR(-0.001,2)` is `¥-0.00`),
  and the localized ja-JP `General` format code, including its DBNum
  variants, is recognised as `General`. `REGEXEXTRACT`, `REGEXTEST` and
  `REGEXREPLACE` advance a find-all match by code point instead of byte,
  and `REGEXREPLACE` accepts a negative occurrence counted from the end.

- **Date and financial functions:** `DATEVALUE` and `VALUE` read a
  year-less date text against the current year. `YEARFRAC` basis 1
  divides a multi-year span by the average length of every calendar year
  in it, and the financial functions share that rule. The
  date-sensitive financial functions (`COUP*`, `ACCRINT*`, `DISC`,
  `INTRATE`, `RECEIVED`, `TBILL*`, `PRICE*`, `YIELD*`, `DURATION`,
  `MDURATION`, `ODDF*`, `ODDL*`, `AMOR*`) honour the workbook's 1904 date
  system. `YIELD` and `ODDFYIELD` return `#NUM!` instead of a clamped
  result, `AMORDEGRC` / `AMORLINC` cap their period, `DAYS` / `DAYS360`
  reject an invalid serial, `DAYS360`'s US February rule matches the
  30/360 year fraction, and `TIME` no longer picks up rounding noise.
  `WEEKNUM` with an unsupported return type returns `#NUM!` instead of
  treating it as type 1.

- **Math and statistics:** `ROUNDUP`, `ROUNDDOWN` and `TRUNC` snap
  binary-representation noise as `ROUND` does, so a value that scales to
  `28.999999999999996` rounds to 29. `VAR` / `STDEV` coerce a directly
  passed boolean or numeric text as documented. `COMBIN`, `COMBINA` and
  `MULTINOMIAL` stay exact within 2^53, `T.DIST` / `T.INV` stay accurate
  at very large degrees of freedom, and `MAKEARRAY` rejects a `rows` or
  `cols` below 1 with `#VALUE!`. The gamma-based distributions no longer
  race on shared state under parallel recalc.

- **Conditional formatting:** an icon set draws no icon for a value
  below its floor, at the floor itself when `gte` is off, instead of the
  lowest icon. Time-period rules (this/last/next week and month) read
  dates on the workbook's 1904 epoch when it uses one.
  `ContainsBlanks` / `NotContainsBlanks` treat a formula-produced empty
  string or a whitespace-only cell as blank, as Excel does. From
  `.xlsx`, a bare `<iconSet>` reads as `3TrafficLights1` and a missing
  `allowBlank` on a validation as false, the defaults Excel writes.

- **Pivot tables:** a Top-N, Bottom-N or percentage value filter on an
  inner field ranks and keeps whole groups within each parent group
  instead of across the flattened axis, and Bottom-N is supported.
  Subtotals and grand totals for average, min, max and distinct-count
  style aggregations are re-derived from the underlying records,
  including after a value filter or a Show Values As margin, instead of
  summed from child aggregates. Date grouping buckets against the
  workbook's date system; a table placing a date- or number-grouped field
  on an axis is rejected instead of collapsing it to one blank item; a
  non-finite aggregate evaluates to `#NUM!`; and a field's sort order and
  the Values pseudo-field's position round-trip. Tabular and outline
  layouts render pivots with column fields, one header column per field,
  with the grand-total header on the first column-header row and a
  repeated outer column item shown once. An `.xlsb` pivot location
  outside the sheet grid, or a pivot cache field holding a record the
  reader does not decode, fails the load with a named error instead of
  being trusted or skipped.

- **Print layout:** page width no longer counts hidden columns, and the
  JIS B4 and B5 paper sizes resolve instead of falling back to A4.

- **Formula spelling on load and save (`.xlsx` and `.xlsb`):**
  conditional-format, data-validation and defined-name formulas read back
  in formula-bar spelling from both formats, so a stored
  `_xlfn.SINGLE(x)` / `_xlfn.ANCHORARRAY(x)` reads as `@x` / `x#` and a
  known function loses its `_xlfn.` prefix, while an unknown function
  keeps it and round-trips unchanged. A loaded formula without the
  dynamic-array mark shows `@` exactly where Excel's formula bar does,
  and a typed `@` saves as `_xlfn.SINGLE` in both formats instead of
  being rejected. `_xHHHH_` escapes decode symmetrically with how the
  writer encodes them, a defined name's storage prefixes are re-applied
  on save, and a shared formula is re-parsed without its storage
  prefixes. A package-absolute relationship target (a leading `/`)
  resolves from the package root, so files written by other spreadsheet
  libraries load.

- **Dynamic arrays and cached values on save:** saved files mark a
  formula as dynamic-array (`cm` in `.xlsx`, the anchor form in `.xlsb`)
  exactly where Excel does, derived from the formula's static shape; the
  mark used to be set on formulas Excel leaves unmarked, such as
  `INDEX(A1:A2,1)`, and missing on `IFS` / `SWITCH`. A legacy CSE block
  keeps its fixed range. Spilled cells are saved with their values, so a
  reader that does not recalculate no longer sees blanks, and a loaded
  spill no longer turns into constants that block its own recalculation.
  Volatile and always-calculated formulas carry `ca="1"` / `aca="1"` in
  `.xlsx` and the matching flags in `.xlsb`. A cached newer error is
  stored as the legacy code Excel writes (`#SPILL!`, `#CALC!` and the
  other newer errors as `#VALUE!`, `#GETTING_DATA` as `#N/A`); the `.xlsb`
  writer stored the newer code, which made Excel refuse the whole file.

- **Structural edits:** inserting or deleting rows and columns moves the
  ranges and formulas of `x14` conditional-format rules and sparkline
  groups kept from a loaded file, as Excel does, where they kept their
  old coordinates. A delete that leaves two conditional-format ranges
  edge to edge joins them into one.

- **XLSB formula encoding:** formulas are written with the token types
  and classes Excel writes, so Excel reopens an engine-written file with
  the same formula text and values. An operator's argument
  (`=SUMPRODUCT(A1:A2*2)`, `=SUM(A1:A2*2)`) no longer gains an implicit
  intersection when Excel reopens it; `=(1+2)*3` keeps its parentheses; a
  formula that is only a reference (`=A10:A11`) spills instead of showing
  `#VALUE!`; `IF`, `CHOOSE` and `IFERROR` carry the jump data Excel
  expects; formulas calling `NOW`, `RAND`, `INDIRECT` or another volatile
  function recalculate on F9; `LAMBDA` definitions and named calls,
  sheet-qualified names, whole-row and whole-column areas, negative
  literals, `#REF!`, array constants and reference-returning calls encode
  and decode as Excel saves them; and an undefined function name is
  stored as a call instead of falling back to the cached value. Array
  constants holding text, booleans or errors load. An undefined
  `Sheet!Name` gets an empty record scoped to that sheet, so a same-named
  workbook definition no longer leaks into it.

- **XLSB workbook structure:** the tab-selected sheet is recorded
  correctly, so a multi-sheet book no longer opens with every tab
  grouped. Retained `x14` blocks keep pointing at the right sheets on
  save, and a retained pivot table, pivot cache, styles part or `x14`
  block that an edit left stale fails the save with a named error instead
  of writing a file with the edit silently missing. Workbook protection
  carrying only `lockWindows`, `lockRevision` or a revisions password is
  no longer written as an empty protection record. A load is bounded by a
  cumulative cell-record budget, so a wide sparse row cannot exhaust
  memory, and `.xlsx` data tables and `.xlsb` save-time losses are
  reported in the package diagnostics.

- **Bindings:** the pthread npm build no longer hangs when bundled. Its
  workers were spawned from the entry shim, which a bundler honouring
  `"sideEffects": false` reduced to an empty chunk, so `createFormulon()`
  never resolved; workers now boot from the Emscripten module, which
  `"sideEffects"` names. `delete()` / `dispose()` called from a workbook's
  own iterative-progress callback throws instead of freeing a handle the
  running recalc still uses. Spec objects are validated field by field,
  so a missing pivot or data-field name reaches the C ABI as absent
  rather than the string `"undefined"`, and an out-of-range
  `recalcParallel` thread count or a non-number `setError` code is
  rejected before any C call. The log sink receives a copied buffer,
  which the `/threads` build needs because most browsers' `TextDecoder`
  refuses a shared one, and a throwing sink no longer leaves a pending
  exception in the native Node addon.

- **Python:** a WASM trap during a call raises `FormulonError` instead of
  a raw `wasmtime.Trap`, and marks the shared module instance poisoned so
  every later call fails the same way instead of touching state the trap
  may have corrupted. The cell, defined-name, table and passthrough
  iterators raise `FormulonError` when `close()` runs between two
  `next()` calls instead of reading a freed handle, and
  `set_phonetic_runs` accepts an empty run's text, as Excel writes for
  out-of-range kana runs.

## [0.11.1] - 2026-08-22

### Added

- A workbook's default font can be declared. Font 0 is the record every
  unstyled cell resolves to, and a new workbook seeds it with Excel's
  Calibri 11; since the seeded table already owns index 0, `add_font`
  could only ever append beside it, leaving a ja-JP host no way to say
  what a never-styled cell should be saved as. `fm_workbook_set_default_font`
  fills that gap, and `fm_styles_set_font` overwrites any existing slot
  in place for the general case of restyling every `<xf>` that names it.
  Both reach WASM and the native Node addon as `setDefaultFont` /
  `setFont` and Python as `set_default_font` / `set_font`; the current
  default reads back through the existing `getFont(0)`.

- Phonetic guides can be authored span by span, not only as one reading
  for the whole cell. The core has kept one run per `<rPh>` block since
  0.11.0, but the bindings carried a single string in each direction, so
  reading a partially annotated cell and writing it back collapsed every
  span into one whole-cell annotation.
  `fm_workbook_set_cell_phonetic_runs` takes the runs as an ordered
  partition and `fm_workbook_get_cell_phonetic_run_count` /
  `fm_workbook_get_cell_phonetic_run` read them back with their spans.
  They reach WASM and the native Node addon as `setCellPhoneticRuns` /
  `getCellPhoneticRuns` and Python as `set_phonetic_runs` /
  `get_phonetic_runs`. The flattening `getCellPhonetic` is unchanged and
  still returns the readings concatenated.

- `getCellPhonetic` and `setCellPhonetic` reach the native Node addon,
  which had neither. They were the last cell-level pair that existed on
  WASM and Python only.

- How a phonetic guide renders can be authored, not only round-tripped.
  `fm_workbook_set_cell_phonetic_properties` /
  `fm_workbook_get_cell_phonetic_properties` carry the font that draws the
  ruby, the kana form Excel generates readings in and how the kana is
  distributed over the characters it covers. They reach WASM and the
  native Node addon as `setCellPhoneticProperties` /
  `getCellPhoneticProperties` and Python as `set_phonetic_properties` /
  `get_phonetic_properties`, with `FM_PHONETIC_TYPE_*` and
  `FM_PHONETIC_ALIGNMENT_*` naming the ordinals. Deliberately separate
  from the run entry points in both directions: editing the readings does
  not reset the rendering, and setting the rendering does not touch the
  readings. A value write still clears both.

- Phonetic guides survive the MS-XLSB container. `BrtSSTItem`'s phonetic
  tail is now decoded and emitted, so furigana no longer disappears when
  a workbook is saved as `.xlsb` or read back from one. The binary form
  stores the kana once and gives each run a start offset into that
  concatenation, and elides the run array entirely for a whole-string
  reading; both shapes are handled. The shared-string interner keys on
  the guide as well as the text, so two cells reading the same kanji
  differently no longer collapse onto one entry.

### Fixed

- A font's `<scheme>` theme link survives a load and save. A ja-JP
  workbook's Normal font carries `scheme="minor"`, which is what makes
  Excel show it as the body font and re-resolve it when the theme
  changes; the element was dropped on read, so re-saving rewrote the font
  as a literal name. It now round-trips through both containers — the
  binary form is `BrtFont`'s `bFontScheme`, whose ordinals the field
  shares — and reaches the record projections as `scheme` on WASM, the
  native Node addon and Python, so reading a font, editing one field and
  writing it back no longer unlinks it.

- A phonetic guide's `<phoneticPr>` block survives a load and save. Only
  the `<rPh>` runs were carried, so a guide set to hiragana or to a
  distributed layout came back as Excel's default half-width katakana on
  the next save. Which font renders the ruby, which kana form it uses and
  how it is distributed now round-trip through both containers, and are
  written beside every annotated string item the way Excel writes them.
  An absent element and a bare `<phoneticPr/>` resolve differently —
  half-width katakana / no control against full-width katakana / left —
  and are now read apart.

## [0.11.0] - 2026-08-22

### Added

- `pivotFieldAddItemAt` / `pivot_field_add_item_at` reaches WASM and the
  native Node addon, where only Python and the C ABI had it. It is the
  only way to name a pivot item by its cache index, and therefore the
  only way to express the blank member of a pivot axis — an empty label
  passed to `pivotFieldAddItem` cannot name it, and filtering on it
  silently did nothing. The C header had prescribed this entry point as
  the required alternative all along.

- `getIterative()` reads back the iterative-calculation settings on WASM
  and the native Node addon, which previously carried only the setter.
  It returns `enabled`, `maxIterations` and `maxChange` beside the usual
  status, so a host can render the settings dialog it is about to write
  to. Python already had it.

- A sheet tab's `veryHidden` visibility can now be set, not only
  preserved. Excel leaves a very-hidden sheet out of its "Unhide"
  dialog, which is how a workbook keeps a settings or lookup sheet out
  of a user's reach; until now that state round-tripped through a file
  but no surface could newly state it, and the boolean setter could only
  clear it. `fm_sheet_set_visibility` takes the three-state
  `fm_sheet_visibility_t` and reaches WASM and the native Node addon as
  `setSheetVisibility` and Python as `set_sheet_visibility`.
  `fm_sheet_view_t` gains a `visibility` field carrying the resolved
  state, surfaced as `SheetView.visibility` on all three bindings.

- Worksheet print settings can now be authored, not only read back. Page
  setup, margins, print options, print area, print titles, header/footer
  and manual page breaks each reach the C ABI as a typed setter beside the
  getter that already existed, and all of them are exposed through WASM,
  the native Node addon and Python. A header/footer is set from decoded
  section strings — the caller spells a literal ampersand the way Excel's
  header syntax does, as `&&`, and the engine escapes it for the file — so
  no caller has to assemble XML to change one section. Every setter has a
  raw-XML counterpart for the parts this engine does not model, and a
  fragment handed to one of those is validated as well-formed and bounded
  before it is stored, so a malformed fragment is rejected at the call
  rather than on save.

- A newly created workbook carries the minimum style table Excel writes:
  one font, the reserved `none` and `gray125` fills, one border, one cell
  style. The writer used to synthesise these on save only while the table
  was empty, so a caller that appended a single fill took the slot Excel
  reserves for `none` and shifted every `fillId` in the file by one — a
  break visible only after the file reached Excel. Seeding also makes
  index 0 a usable template, so "copy the default and change the size"
  works without a separate partial-input type.

- `fm_sheet_set_range_xf_index` / `setRangeXfIndex` / `set_range_xf_index`
  stores one style index across every cell in a rectangle in a single
  call, materialising cells that do not exist yet as styled blanks, so
  ruling a report's border no longer costs one ABI crossing per cell.

- Cross-workbook references resolve. The index-spelled forms Excel
  stores — `[1]Sheet1!A1`, `[1]Sheet1!A1:B2`, `[2]!Name` and the quoted
  `'[1]My Sheet'!A1` — parse and evaluate against the cached values in
  the external link part, which is the same cache Excel itself reads
  once the supporting workbook is closed. The path-spelled
  `[Book1.xlsx]Sheet1!A1` remains unsupported: it carries no link-table
  index to bind the file name against. On XLSB the supporting-book table
  is decoded too, so an external sheet index resolves to the book it
  names rather than being assumed to be this workbook — index 0 is not
  always the file itself, and assuming it rebound a cross-workbook
  reference onto a same-numbered local sheet. Such a reference is still
  not encoded on an XLSB save.

- The XLSB pivot parts — `pivotCacheDefinition`, `pivotCacheRecords` and
  the pivot table itself — decode into the same model the OOXML pivot
  reader builds, so a pivot in an `.xlsb` file is evaluated rather than
  dropped. A pivot whose record encoding has not been measured against a
  file Excel wrote is skipped rather than guessed at. `GETPIVOTDATA`
  resolves
  a data field by its display name and then by the source column it
  aggregates, which is the spelling Excel writes into a
  click-generated formula, and a lookup that names one axis in full
  while leaving the other open answers from the leaf totals instead of
  `#REF!`.

- Phonetic (furigana) annotations keep the spans they annotate. A cell's
  kana is stored as one run per `<rPh>` block carrying that block's
  UTF-16 span, rather than flattened to a single whole-string
  annotation, so a partially annotated string survives a load and save.
  `PHONETIC` substitutes only the annotated spans and passes the rest of
  the surface text through, surfacing kana and original characters side
  by side. Two cells share a shared-string entry only when their spans
  agree as well as their text.

- The spill operator `#` accepts an anchor that is computed rather than
  written out, so `OFFSET`, `INDIRECT`, a `CHOOSE` or `IF` branch, a
  parenthesised reference and a `LET`-bound name each anchor a spill. An
  anchor naming more than one cell is `#REF!` rather than being narrowed
  to its top-left corner. A formula of this shape evaluates through the
  tree walker, since the bytecode compiler's spill-reference opcode
  indexes a pool a computed anchor has no entry in.

- Two more authored pivot filter families are evaluated. The week
  windows (`thisWeek` / `lastWeek` / `nextWeek`) resolve against the
  calendar week running Sunday through Saturday rather than a rolling
  seven days, so a `thisWeek` read on a Friday still starts on the
  preceding Sunday and the three windows tile with no gap or overlap.
  The recurring `M1`..`M12` / `Q1`..`Q4` selectors name a calendar
  position rather than a contiguous range — `M1` keeps every January of
  every year — so they resolve without reading the clock. No authored
  `<filters>` family is left unevaluated.

- The range operator composes to any depth. A `:` chain folds into the
  rectangle bounding every leaf, and mixing axes is a composition rather
  than an error: a whole-column leaf spans every row and a whole-row
  leaf every column, so `A:C:1:3` bounds the whole grid and
  `SUM(A:C:E:G)` sums the bounded columns instead of answering
  `#VALUE!`. Every existing two-endpoint spelling keeps its anchors and
  sheet qualifier verbatim. Excel stores such a chain as written while
  this engine records the rectangle it names, so a re-serialised
  `H:H:E:E` is written back as `H:E` — the rectangle and every evaluated
  result are identical.

### Changed

- The style table a newly created workbook now carries (see above)
  reserves index 0 for its font, border and cell style, and indices 0
  and 1 for its `none` and `gray125` fills. The first `addFont` or
  `addBorder` a caller appends to a freshly created workbook therefore
  returns index 1, not 0, and the first `addFill` returns index 2, not
  0; callers that assumed a fresh workbook's style tables start empty
  should read the index back from the setter's return value rather than
  assuming it.

- `fm_workbook_set_iterative` / `setIterative` / `set_iterative` clamp
  `max_iterations` to 32767, and the getter reports the clamped value
  rather than the requested one, so a caller that set a larger cap and
  reads it back now sees 32767. A loaded file whose `iterateCount`
  exceeds that saves back as 32767 rather than its original figure —
  such a file is schema-valid, but Excel's own dialog cannot produce
  one, and the alternative was leaving an unbounded recalculation
  reachable from a file.

- `addValidation`'s `allowBlank` now defaults to `false` on WASM and the
  native Node addon, where it defaulted to `true`. Python already
  defaulted to `false`, and both TypeScript declarations state the
  general rule that an omitted boolean field defaults to `false` with
  `showDropDown` as the only carve-out. A validation rule that omits the
  field therefore changes meaning: it now rejects empty cells. Spell
  `allowBlank: true` to keep the old behaviour.

- Argument-validation failures on the native Node addon report
  `kBindingNullPointer` (7001) instead of `kBindingInvalidHandle`
  (7000). 7000 means "this workbook handle was already destroyed", so a
  caller handed that code for a malformed argument would reach for the
  one recovery that cannot help — recreating the workbook. The affected
  entry points are `pivotFieldAdd`, `pivotDataFieldAdd`,
  `pivotDataFieldSet`, `pivotFilterAdd`, `setIterativeProgress` and the
  module-level `setLogSink`. A caller branching on 7000 for those will
  see the new code.

- `formulon eval` no longer treats a malformed formula as a structural
  failure. Malformed syntax now resolves to Excel's ordinary `#NAME?`,
  which prints to stdout with an exit status of 0, the same result every
  other surface over the C ABI already produced. A caller that used the
  exit status to detect a typo should test the printed value instead.

- A number written into the pivot cache now carries its shortest
  round-tripping spelling, so a cached `0.1` is stored as `0.1` rather
  than `0.10000000000000001`. Cell values already used this form; the
  cache did not, so the same number appeared two ways in one file. The
  value each spelling parses back to is identical, but the saved bytes
  differ from those earlier releases produced.

- Row heights, column widths, the default sheet format metrics, font
  sizes and colour tints are written the same shortest round-tripping
  way, for the same reason. A default column width is now stored as
  `8.43` rather than `8.4299999999999997`, and a colour tint as `0.35`
  rather than `0.34999999999999998`. The values are unchanged and only
  the spelling differs, but a byte comparison against a file saved by
  an earlier release will show it.

- The WASM size report's soft ceilings moved to 2.75 MiB uncompressed and
  736 KiB Brotli. The previous pair sat below the shipped binary, so the
  warning fired on every run and said nothing about the change being
  measured. The new pair sits just above it, which is what a tripwire has
  to do to be read: one feature's worth of growth trips it, and the hard
  ceilings — unchanged at 3.00 MiB and 768 KiB — stay a further 0.25 MiB
  and 32 KiB away.

- `fm_sheet_view_t` gains a `visibility` field between `tab_hidden` and
  `show_grid_lines`. The struct is 48 bytes on a 64-bit host as before —
  the new field fills what was tail padding — but grows from 40 to 44
  bytes on wasm32, so code compiled against the previous layout must be
  rebuilt. `tab_hidden` is unchanged in meaning and is non-zero for both
  hidden states, so a caller reading only that field sees a very-hidden
  sheet as hidden rather than as visible. `fm_sheet_set_tab_hidden`
  still refuses to demote a very-hidden sheet to plain hidden;
  `fm_sheet_set_visibility` is the way to do that.

- The `fuzz-parser`, `fuzz-xlsx` and `fuzz-eval` make targets are
  replaced by `fuzz`, which builds every harness under
  `-DFM_BUILD_FUZZ=ON` and runs the smoke tier, and `fuzz-long`, which
  runs them to a wall-clock budget. The three old names printed that
  they were unimplemented and exited 0; they covered three of the five
  harnesses in any case. `bench` likewise printed that it was
  unimplemented and now runs the regression check. Scripts calling the
  four old names will now fail with "No rule to make target" rather than
  silently succeeding.

- The fuzz harnesses need a Clang that ships libFuzzer, which the Apple
  toolchain does not; `fuzz` looks for a Homebrew LLVM and falls back to
  `clang` on the `PATH`, overridable with `FUZZ_CC` / `FUZZ_CXX`. When
  neither can link `-fsanitize=fuzzer`, configuring now refuses in
  seconds and names the cause, rather than failing at link after a full
  sanitized rebuild of the core.
  AddressSanitizer stays off by default because it deadlocks in its own
  shadow-memory setup against recent macOS dynamic linkers; set
  `FUZZ_SANITIZERS` to opt back in.

- The xlsx fuzz target now seeds from the bundled workbook fixtures.
  Starting from an empty corpus, it only ever explored the ZIP header
  rejection path.

### Performance

- Parallel recalculation no longer routes a cell to the calling thread
  merely because it is volatile. The isolation exists so that a formula
  resolving a reference at evaluation time — `INDIRECT`, `OFFSET` —
  cannot read a cell a worker is writing, which says nothing about a
  formula whose only volatility is producing a fresh value. `RAND`,
  `NOW`, `TODAY` and their relatives were isolated all the same, so a
  sheet of them created no worker pool at all and a four-thread
  configuration ran exactly like a single-threaded one. Reference-
  resolving formulas stay isolated. The gain follows the work each
  formula does rather than the cell count: a heavy volatile formula
  scales with the pool, while a sheet of bare `=RAND()` is bound by the
  per-cell commit rather than by evaluation and does not.

- Every sort in the engine runs through one shared index-sort body — the
  C API parts list, conditional-format span merging, `MODE`'s frequency
  table, the dynamic-array `SORT` / `SORTBY` lane permutations, spill
  release, the OOXML and XLSB readers and writers, the pivot filter and
  hierarchy builders, and the sheet's own row operations — instead of
  each site instantiating its own copy of the introsort. This is a WASM
  code-size measure: the shared order is stable, which is at least as
  strong as what the replaced comparators guaranteed, so no result
  moves. `ROMAN` additionally walks a fixed value-descending pair table
  instead of building and sorting one on every call, with unchanged
  output.

### Fixed

- `TRIM` no longer rewrites an ideographic space (U+3000) as an ASCII
  space. A run of trimmable spaces still collapses to one character, but
  the character it keeps is the one the run started with, so `"a　　b"`
  trims to `"a　b"` and a lone interior full-width space is left alone.
  Rewriting it to U+0020 silently reflowed Japanese text that had been
  spaced deliberately, and shifted every later `FIND` / `MID` position by
  the difference in byte width.

- `ISOMITTED` now reports `TRUE` for an argument the caller omitted with
  an empty slot — `f(1, , 3)` as well as a missing trailing one. It had
  answered `FALSE` for the empty-slot spelling, which is the spelling
  Excel actually uses for an omitted argument, so the optional-parameter
  idiom (`IF(ISOMITTED(x), default, x)`) selected the wrong branch in a
  `LAMBDA` helper. Omission is a property of the call's syntax rather
  than of the argument's position, and applies to leading, middle and
  trailing slots alike.

- A zero-length string is no longer treated as a blank cell.
  `CELL("type", ...)` returns `"l"` for one, whether it was entered as a
  constant or produced by `=""`, and `"b"` is reserved for a genuinely
  blank cell. Wildcard criteria stop excluding it, so
  `COUNTIF(range, "*")` counts every text cell and leaves out only
  blanks. `COUNTIF(range, "=")` is now separated from
  `COUNTIF(range, "")`: both parse to an equality against an empty
  right-hand side, but the bare comparator is a blank-cell probe that a
  zero-length string fails. All three had been built around an empty
  string that reached the sheet as a blank cell, which is a state Excel
  does not produce for `=""`.

- A whole-axis pair is spliced in the stored form of a formula as well
  as in its canonical one: `RangeOp(A:A, C:C)` is written as `A:C`
  rather than `A:A:C:C`. Only the canonical formatter had the splice, so
  the text written into a file drifted from what Excel writes while
  every evaluated result stayed correct. `RangeOp(A:A, A:A)` stays
  unspliced, since `A:C`'s compaction would read back as a single
  whole-column reference rather than as the pair that was written.

- A row or column insert or delete now moves the worksheet auto-filter's
  `ref` rectangle with the cells it filters, and drops the auto-filter
  when the edit consumes its whole range. Previously the rectangle stayed
  where it was while the data moved out from under it. The criteria
  offsets on `<filterColumn>` are not remapped, so a column edit inside
  the filtered range still leaves the criteria attached to the wrong
  columns.

- Bulk range reads and spilled-range (`A1#`) reads handed back text whose
  bytes belonged to the sheet, so a value read from a range could be
  invalidated by a later write to the cell it came from or by a clear of
  the spill region behind it. The scalar cell read already returned owned
  bytes; the bulk paths now match it.

- A workbook could specify its own iteration budget with nothing
  bounding it. `<calcPr iterate="1" iterateCount="4294967295"
  iterateDelta="0"/>` over a two-cell cycle made the first
  recalculation run billions of sweeps: the delta of zero makes the
  convergence test unsatisfiable, there is no time limit anywhere in the
  engine, and cancellation only exists if the host registered a progress
  callback. On WASM that is an unrecoverable hang of the instance. The
  count is now capped at 32767 — Excel's own limit for the setting, so
  the cap costs no fidelity — on the file path, the API setter and the
  solver itself.

- `spinCount` values above 2^32-1 were truncated on read, and a value
  that truncated to exactly zero then lost the attribute entirely on
  save, leaving a protected sheet's hash and salt with no iteration
  count. Such values now saturate at 4294967295.

- Converting an `.xlsb` to an `.xlsx` silently dropped every cell's
  alignment, protection and `apply*` flags. The XLSB styles reader left
  seventeen `<xf>` fields unmodelled on the grounds that they round-trip
  through the retained `xl/styles.bin` — which holds for `.xlsb` to
  `.xlsb`, but not for a conversion, because the `.xlsx` writer emits
  from the model rather than from the retained bytes. `vertical="center"`
  is the ja-JP Excel default, so a typical Japanese-authored workbook
  came out of the conversion with every cell bottom-aligned. The reverse
  direction lost `textRotation`, `indent`, `justifyLastLine`,
  `shrinkToFit` and `readingOrder`. Both are now carried. One field
  still cannot survive an `.xlsb` save: the binary record has no slot for
  `relativeIndent`, which is now stated in the reader rather than left
  to be discovered.

- A deeply nested formula could corrupt memory or kill the WASM module.
  The parser's nesting cap is sized against a measured worst-case stack
  cost, but the WASM link inherited the toolchain's default stack of
  64 KiB while the deepest legal input needs about 97 KiB, and release
  builds disable the overflow check — so the shadow stack ran into the
  data segment underneath it without any diagnostic. A formula of about
  420 characters, well inside both the nesting cap and Excel's
  8,192-character limit, already wrote past the stack; around 600
  characters it faulted outright and took the module instance with it.
  Any host-supplied formula string reached this, as did a crafted
  workbook, since the same parser serves the file readers. The stack is
  now pinned explicitly at 320 KiB — a little over three times the
  measured worst case — for both the embind and the C-ABI variants, and
  a formula past the nesting cap reports the parser's error instead of
  faulting.

- The six sheet-layout setters accepted a width or height of NaN, ±Inf
  or a negative number, and a column or row index past the grid, then
  wrote the result to file. NaN and infinity serialise to an empty
  string, which is not a lexical `xsd:double` at all, so the emitted
  `<col width=""/>` and `<row ht=""/>` made Excel prompt to repair the
  workbook rather than merely rendering something odd; `min="16384"` and
  `r="1048577"` name tracks outside the grid. `fm_sheet_set_column_width`,
  `fm_sheet_set_column_hidden`, `fm_sheet_set_column_outline`,
  `fm_sheet_set_row_height`, `fm_sheet_set_row_hidden` and
  `fm_sheet_set_row_outline` now return `kInvalidArgument` and leave the
  model untouched. A width or height of zero remains legal — a
  zero-width column is a real thing. Callers that were getting a success
  status for these arguments will now see a failure, but the file that
  path produced could not be opened.

- A failed call on WASM and the native Node addon could report the
  message left behind by an unrelated earlier call. Failures raised
  inside the binding layer never set the thread-local diagnostics, so
  the envelope carried whatever residue was there — or nothing at all
  when the previous call had succeeded. A destroyed handle now says so,
  and a rejected argument names the parameter it rejected.

- `getCellStyleXf` omitted `xfId` from the record it returns, although
  both TypeScript declarations state it is always 0 there and the C
  layer sets it. The returned object now matches `getCellXf`'s shape, so
  the documented read-modify-write round trip no longer depends on the
  field's absence happening to coincide with its default.

- Evaluating a pivot with a manual filter walked every item of every
  field for every record, so a filter's cost grew with the length of the
  item list Excel writes for any field placed on an axis — even when
  nothing was hidden. The hidden-label set is a property of the table,
  not of a record, so it is now built once per evaluation and each
  record costs one lookup per field that hides something. On fifty
  thousand records the effect ranges from roughly forty-fold to two
  orders of magnitude depending on item-list length, and the cost no
  longer follows that length at all.

- `=A:C B:B` and its relatives returned `#VALUE!` instead of the
  intersection Excel computes. Deriving a rectangle from full-column or
  full-row endpoints had implementations left that predated the shared
  one; the intersection operator, the scalar shape seam and the
  aggregator expander now use the same derivation as every other
  consumer.

- `ROW` and `COLUMN` given a whole-column or whole-row reference
  collapsed to `1` instead of projecting every index of that axis, so
  `=SUM(ROW(A:A))` returned 1 rather than 549,756,338,176. `ROWS` and
  `COLUMNS` already reported the full axis for the same reference, so
  the singular and plural forms disagreed with each other. `=ROW(A:C)`
  and `=COLUMN(1:3)` were wrong the same way. The off-axis spellings —
  `=ROW(1:1)`, `=COLUMN(A:A)` — were already right and still are.

- `dump` and `paginate` discarded the diagnostics a load produces when
  part of a workbook could not be represented, so a snapshot or a page
  geometry could be taken from a silently incomplete workbook. All three
  subcommands now report them, prefixed with the subcommand that read
  the file.

- A cell containing a tab or a newline could forge column and row
  boundaries in `eval`'s plain grid output. Both are now escaped, the way
  `dump` already escaped them.

- `--version` and `--help` exited 0 even when the write to stdout failed,
  so a caller reading them through a closed pipe saw success and no
  output. They now fail.

- A manual page break written to an xlsx landed one row or column late
  when the file was opened in Excel, and a break read from an
  Excel-authored file paginated one track early. OOXML's `<brk id>` is
  already the zero-based index the break precedes, but the reader
  subtracted one from it and the writer added one back. The two errors
  cancelled inside a read/write cycle, so only pagination results and
  Excel disagreed with the model.

- `SUM(IF(condition_range, a, b))` and its relatives returned `#VALUE!`
  instead of aggregating the masked range. The array condition was
  coerced to a single boolean at the aggregator's argument seam rather
  than picked per cell, so the array idiom — which needs no
  Ctrl+Shift+Enter in Excel 365 — failed under every eager aggregator.
  The criteria and lookup families reached a second copy of that seam
  and failed the same way: `SUMPRODUCT` and `INDEX` returned `#VALUE!`
  and `#REF!`, and `COUNT` returned a plausible wrong number rather than
  an error, because a function that inspects error arguments drops a
  failed expansion as "not a number". A bare `=IF(A1:A5<=3,1,0)` spilled
  correctly throughout; only the consuming forms were affected.

- A range chaining `:` over whole-column or whole-row endpoints on one
  axis lost an endpoint on save. `=SUM(A:A:A:A)` was written as
  `=SUM(A:A)`, which reads back as a single whole-column reference
  rather than the pair. Where the two endpoints were anchored
  differently — `=SUM($A:$A:A:A)` — the surviving form also answered
  differently once filled, because one endpoint had stopped being
  relative. The splice that compacts `A:A:C:C` to `A:C` now applies only
  across axes, where there is something to compact. A side effect worth
  knowing: this engine does not yet evaluate a `:` chain whose endpoints
  are themselves whole-axis ranges, and the old compaction was
  incidentally hiding that on save, so such a formula now keeps its
  `#VALUE!` across a save and reload instead of appearing to repair
  itself.

- `IFS` with an array condition returned `#VALUE!` instead of deciding
  per cell. `=IFS(A1:A5<=3,A1:A5,TRUE,0)` now spills `{1;2;3;0;0}`, and
  several array conditions in sequence each win the cells they hold for.
  A condition that wins as a scalar still settles the call before any
  later condition is evaluated, so a later argument that would error is
  not reached.

- `SWITCH` with an array as its first argument returned a single value
  instead of selecting per cell. Nothing matched, so the walk fell
  through to the default: `=SWITCH(A1:A5,1,10,0)` answered `0` where
  Excel spills `{10;0;0;0;0}`, with no error to show something had gone
  wrong. Where there is no default, the rule that an unmatched subject
  yields `#N/A` now applies per cell, so the cells that match keep their
  value.

- A cell reference followed by a space and a parenthesis parsed as a
  function call rather than an intersection, so `=A1 (B1:C5)` returned
  `#NAME?` where Excel evaluates the intersection. Only a range tail was
  carved out of the call rewrite; a parenthesised operand of any other
  shape still became a call. A parenthesised comma list is also a legal
  intersection operand, and was separately rejected as an invalid range.
  `LOG10` is the only Excel function whose name is shaped like a cell
  reference, so it stays a call.

- Negative elements of an array constant were written back in a form
  neither Excel nor Formulon can read: `={-1,2}` was saved as
  `{(-1),2}`, and an array constant admits constants only. Formulas such
  as `=IRR({-100,40,40,40})` were affected on every save.

- A sheet name beginning with a digit was written without quotes in a
  3-D reference span, producing `3S1:Daa!A1`, which does not parse — the
  leading run is read as a number. Renaming a sheet to a name of that
  shape, such as `2026Q1`, corrupted every reference to it.

- A range whose endpoint is a defined name or `LAMBDA` parameter spelled
  like a column letter was written in a form that read back as a
  whole-column range: `RO:r` became `RO:R`. Such an endpoint is now
  parenthesised. Ordinary ranges, whole-column ranges and ranges ending
  in a function call are written as before.

### Documentation

- The row and column edits now state what they do not move. They remap
  every structure the engine models, but worksheet content kept
  byte-verbatim because the engine does not model it — the worksheet
  `<extLst>` behind DataBar extended settings, sparkline groups and
  slicer anchors, plus any unmodelled `<worksheet>` child — keeps its
  pre-edit rectangles, and nothing reports it. Excel resolves the
  mismatch differently per extension: it drops an `x14`
  conditional-formatting entry whose range no longer matches the rule it
  extends, so an extended DataBar reverts to its legacy rendering, while
  a sparkline keeps drawing from a source range that has moved out from
  under it. The limitation now reaches the C header, both TypeScript
  declaration files and the Python docstrings, where previously only an
  internal header mentioned it.

## [0.10.0] - 2026-08-18

### Added

- A pivot with a report-filter (page) field now renders the header block
  Excel draws above it: one row per page field carrying the field name and
  the item it is showing, then a blank separator row, all inside the pivot's
  own extent. The selection is resolved during evaluation, where the bound
  cache is available, so `PivotResult` carries it and the projection only
  draws it. Excel records a selection two ways and both are read — a single
  chosen item as `<pageField item>`, a wider one as the field's own hidden
  items — and the placeholder text follows the locale (`(すべて)` under the
  ja-JP profile). `<pageFields>` is decoded for the report order and the
  selection but still re-emitted from the passthrough bin, so a round trip
  is byte-identical; the writer synthesises the element only for a table
  that never carried one, which a page-axis field previously saved without.

- `INDIRECT(ref_text, FALSE)` reads R1C1 text instead of returning `#REF!`
  for every call. An axis is written absolutely (`R5C2`) or relative to the
  cell holding the formula (`R[-1]C`, or a bare `R` meaning the same row),
  and an endpoint naming one axis is unbounded along the other, so `R5` is
  the whole of row 5 exactly as `5:5` is. The `a1` flag selects a grammar
  rather than adding a fallback: A1 text under `FALSE` is `#REF!`, as R1C1
  text under `TRUE` already was. A relative axis evaluated with no formula
  cell — the ad-hoc "evaluate this text" entry points bind none — is `#REF!`
  rather than being measured from an assumed origin.

- A workbook-level clock seam, so results that depend on when they are
  computed can be pinned to one instant. `NOW`, `TODAY` and the pivot
  relative-period filters otherwise each read the host clock independently,
  which makes a recalc internally inconsistent across a midnight boundary
  and makes any such result untestable. It reaches the C ABI as
  `fm_workbook_pinned_now` / `fm_workbook_set_pinned_now` /
  `fm_workbook_clear_pinned_now` over a new 24-byte `fm_civil_time_t`, WASM
  and the native Node addon as `pinnedNow()` / `setPinnedNow(...)` /
  `clearPinnedNow()`, and Python as `pinned_now()` / `set_pinned_now(...)` /
  `clear_pinned_now()`. The reading is carried as local civil fields rather
  than a timestamp, so a pin has no residual timezone interpretation and
  reproduces identically on any host. It is model state, not file state: a
  save does not record it and a reloaded workbook comes back unpinned, and
  an unpinned workbook follows the host clock exactly as before. A pin is a
  calendar instant rather than a normalising constructor — a month of 13 is
  rejected instead of rolled into the next year. Purely additive.
- Authored pivot relative-period and top-N percent / sum filters are now
  evaluated instead of skipped. Thirteen relative-period families ("this
  month", "year to date", ...) prune records against the workbook clock
  before aggregation, and the two running-total flavours of the top-10
  dialog rank an axis leaf by its aggregate and accumulate in descending
  order until the running total first reaches the target, keeping the leaf
  that crosses it — `percent` reading the target as a share of the axis
  total and `sum` as an absolute amount. The three top-N flavours share a
  byte-identical `<top10>` element and are told apart only by the filter's
  `type`, so reading one as another silently produced a different table.
  Week-relative families and the recurring `M1`..`M12` / `Q1`..`Q4` families
  stay unevaluated and remain registered as a divergence.
- Data-bar `x14` settings on every binding. `gradient`, `axisPosition`,
  `negativeFill`, `border`, `negativeBorder` and `axisColor` reach WASM and
  the native Node addon on the `dataBar` object, and Python as the matching
  `DataBar` fields. They live in the `x14` extension rather than the legacy
  `<dataBar>` element, and now survive a save and load rather than
  collapsing to the defaults. Omitting one keeps the model default —
  gradient fill on, automatic axis, negative fill equal to the positive
  fill, no border, black axis — so an object read back from a rule can be
  handed straight to the adder and reproduces it. Purely additive; the C
  ABI already carried the fields.
- `fm_workbook_pivot_field_add_item_at(wb, sheet, pivot, field, cache_index,
  visible)` appends a manual-filter item addressed by its position in the
  bound cache field's shared items — the same index space as OOXML
  `<item x="N">` — and reaches Python as `pivot_field_add_item_at`. This is
  the only way to construct the blank item: it has no label of its own, so
  the filter engine matches it by what it binds to, and the name-addressed
  `fm_workbook_pivot_field_add_item` leaves that binding at index 0 whatever
  the caller passes. The file-load path was already index-addressed; only
  hand-built pivots were affected. Purely additive.
- Container-agnostic save/load loss counters across the C API, WASM, native
  Node, Python, and CLI surfaces. `fm_workbook_save_with_diagnostics` fills
  an `fm_save_diagnostics_t` and `fm_workbook_read_diagnostics` fills an
  `fm_read_diagnostics_t`; both structs are 20 bytes with identical layout
  on native and wasm32. A field means the same thing whichever container
  was written or read, so a caller never has to know which writer ran. On
  top of the previous XLSB-only counters this adds the OOXML reader's
  skipped presentation overlays and unrecognised workbook content type, the
  OOXML writer's renumbered table parts, and — for both writers — dropped
  passthrough parts and dropped relationships, including the XLSB
  sheet-scope relationships that have no OOXML counterpart. They reach the
  bindings as `saveWithDiagnostics` / `readDiagnostics` and
  `save_with_diagnostics()` / `read_diagnostics()`. The CLI reports the
  OOXML load counters on their own `warning: OOXML read diagnostics` line
  and labels the save line by the container actually written. Coverage is
  deliberately partial: these count part-, relationship- and feature-level
  loss in the package readers and writers, so an all-zero result means none
  of the documented losses occurred rather than that nothing was logged.
- Table authoring: `fm_workbook_table_create` / `_update` / `_remove` and
  their `createTable` / `updateTable` / `removeTable` counterparts, which
  reject a column list whose count disagrees with the width of `ref` because
  such a table is a file Excel refuses to open without repair. An update is
  partial — a `NULL` style name keeps the stored style payload, a negative
  header-row or totals-row flag keeps the current one, and an existing
  AutoFilter is retargeted by rewriting only its `ref`, so criteria and
  extensions survive a read-modify-write.
- Worksheet AutoFilter access as an opaque fragment:
  `fm_sheet_get_auto_filter_xml` / `fm_sheet_set_auto_filter_xml`
  (`getSheetAutoFilterXml` / `setSheetAutoFilterXml`) hand over the whole
  `<autoFilter>` element verbatim, so filter criteria, sort state and filter
  extensions can be read, modified and written back without loss.
- Alignment-complete cell formats. `fm_cell_xf` now carries
  `justifyLastLine`, `textRotation`, `indent`, `relativeIndent`,
  `shrinkToFit`, `readingOrder` and a presence flag for each alignment
  attribute, so an explicit zero / false is distinguishable from an omitted
  one. Named-style authoring is reachable through
  `fm_styles_add_cell_style_xf` and `fm_styles_set_cell_style`. The complete
  record reaches WASM, the native Node addon and Python. See **Changed** for
  the struct layout break this implies.
- Multi-cell hyperlinks: `fm_hyperlink` carries the inclusive rectangle end
  as `last_row` / `last_col`, `fm_sheet_add_hyperlink` accepts that
  rectangle, and the bindings expose it as `addHyperlinkRange` /
  `add_hyperlink_range` alongside `lastRow` / `lastCol` on a read-back
  hyperlink. OOXML and XLSB carry the full anchor span in both directions.
- A pivot value filter can name the measure it ranks.
  `PivotFilter::data_field_index` — reachable as `data_field_index` on the
  `fm_pivot_filter_spec_t` that `fm_workbook_pivot_filter_add` takes, and as
  `dataFieldIndex` / `data_field_index` in the bindings — indexes the table's
  data fields; a table with several measures previously always scored the
  first one. Label and date filters ignore the selector.
- `formulon recalc` accepts `.xlsx` or `.xlsb` on both sides, the output
  container following the extension of `-o`, and warns on stderr when a load
  left undecoded formulas, undecoded defined names or dropped package parts,
  or when the writer downgraded formula cells or omitted modelled features.
  Those warnings are data loss rather than status, so `--quiet` suppresses
  only the success line and leaves them visible.
- An XLSB load reports the package parts it could not carry as
  `XlsbReadResult::dropped_part_count`, together with an
  `xlsb.package.parts_dropped` structured warning naming the first such entry
  and why it was skipped. A part typed through an extension default — media,
  embedded OLE, printer settings, the relationships of an unmodelled part —
  was previously absent from a saved workbook with no signal.

### Changed

- `GROUPBY`'s `sort_order` and `PIVOTBY`'s `row_sort_order` /
  `col_sort_order` reject a supplied `0` with `#VALUE!`. The argument is a
  signed column index and that domain has no zero member, which is what
  Excel answers. Omitting the argument still selects the documented default
  of first-occurrence order, and an empty slot between commas counts as an
  omission — so the way to ask for the default is to leave it out rather
  than to spell it. A fraction truncating to zero lands on the same
  rejection. Excel pins the row half of the PIVOTBY pair; the column half is
  held to the same rule rather than accepting on one side what the other
  rejects.
- Saving a workbook whose PivotTable cache declares no worksheet source now
  fails instead of writing the package. Excel offers to repair any file
  carrying a bare `<cacheSource type="worksheet"/>`, and there is no form of
  it Excel accepts without a `<worksheetSource>` — so the writer was
  describing a state that cannot be saved. A declared range is enough even
  when the sheet holds no data, because the cache records carry the values
  themselves. Call `fm_workbook_pivot_cache_set_worksheet_source` (Python
  `set_pivot_cache_worksheet_source`, Node `pivotCacheSetWorksheetSource`)
  after creating a cache. This is a breaking change for a host that builds
  pivots through the API and never set one, but those saves were already
  producing a file Excel would not open cleanly, with nothing in the API
  reporting it; the save path is the only place the host can learn of it.
  Caches read from a file are unaffected — Excel always writes a source.
- The C ABI carries one entry point per operation instead of a base name plus
  an `_ex` / `_ex2` successor. The surviving rung is the one that represents
  the whole model; the narrower one is gone. This is a coordinated,
  source-and-binary breaking change against v0.9.7 — see **Removed** below
  for the per-symbol mapping and the two struct layout changes.
  - `fm_workbook_save_ex` is now `fm_workbook_save_as`, and reaches the
    bindings as `saveAs(format)` (WASM and native Node, replacing `saveEx`)
    and `save_as(fmt)` (Python, replacing `save_ex`). `fm_workbook_save`
    stays as the `.xlsx` default every binding's `save()` calls: it is the
    common case with no argument to get wrong, not a compatibility shim.
  - `fm_cell_xf` is now 88 bytes and carries every optional alignment
    attribute with its presence flag, replacing the 20-byte projection.
    `fm_styles_add_cell_xf` takes it **by value**, so this changes the
    calling convention rather than a buffer size: a caller built against the
    v0.9.7 header reads its arguments from the wrong registers and stack
    slots with nothing to diagnose it. Recompile.
  - `fm_sheet_view_t` is now 48 bytes native / 40 bytes wasm32 and carries
    the display and orientation flags. It is written through a
    caller-supplied pointer, so a stale caller is overwritten 32 (native) or
    24 (wasm32) bytes past the end of its own storage. Recompile.
  - `fm_workbook_defined_name_at` gained the `int32_t* out_local_sheet_id`
    fifth parameter. It is optional: pass `NULL` for the previous behaviour.
  - **Behavioural, and silent: `fm_styles_add_cell_xf` and
    `fm_styles_add_batch` write a different `<alignment>` than before for
    the same caller code.** They now read an alignment attribute only when
    its `has_*` flag is set, instead of inferring presence from the value
    differing from the model default. So a caller that set
    `record.horizontal_align = 3` and relied on that being enough
    recompiles cleanly against the widened record, still gets `kOk`, still
    gets an xf index — and now produces an xf with *no* `<alignment>` child,
    because the value is ignored without its flag. There is no compile error
    and no status to check; the difference is only visible in the emitted
    file. Set the matching `has_*` flag for every alignment attribute you
    mean to write. The same rule makes a value outside its Excel range
    ignored rather than rejected when its flag is clear.
    `fm_styles_add_batch` also now rejects a record whose `xf_id` names a
    `<cellStyleXfs>` entry that does not exist, which it previously could
    not express at all.
- `fm_hyperlink` gained the `last_row` / `last_col` rectangle end, so its
  layout differs from the previous release; recompile against the current
  header before linking.
- miniz is tracked at 3.1.2.

### Removed

Every entry below is a binary ABI break against v0.9.7: the symbols are
gone rather than deprecated, so a caller linked against the old shared
library must be recompiled.

Which consumers a removal reaches depends on which of three distribution
surfaces carried the symbol in v0.9.7, so each entry names them:

- **native** — declared in the header, so reachable by a third party linking
  the library directly.
- **wasm** — listed in `tools/wasm/capi_exports.txt`. The npm WASM package
  and the Python wheel load the same `.wasm`, so they are one surface, not
  two: a symbol absent here was never callable from either.
- **npm-native** — wrapped by the Node addon. The addon binds a subset by
  hand, so a symbol can be native-and-wasm reachable and still absent here.

Entries reaching **native only** are the ones easiest to under-report: they
have no binding to notice their absence and no test in this repo calls them.

- `fm_styles_get_cell_xf`, `fm_styles_add_cell_xf` and
  `fm_styles_get_cell_style_xf` keep their names but take the widened
  `fm_cell_xf` described above; the 20-byte forms are gone. All three:
  native + wasm + npm-native. No binding called them — every one already
  used the `_ex2` forms — so the JS and Python method surfaces are
  unchanged, but a native caller must recompile and a wasm caller passing a
  hand-built 20-byte record will now read past it.
- `fm_styles_get_cell_xf_ex2`, `fm_styles_add_cell_xf_ex2`,
  `fm_styles_get_cell_style_xf_ex2`, `fm_styles_add_cell_style_xf_ex2` —
  renamed to the base names above. Unshipped: on none of the three surfaces
  in v0.9.7, so no consumer is affected.
- `fm_cell_xf_ex2` — merged into `fm_cell_xf`. The `base` member is gone and
  its seven fields are now the first seven members of the flat record, at
  the same offsets. Unshipped.
- `fm_sheet_get_view_ex` / `fm_sheet_view_ex_t` — renamed to
  `fm_sheet_get_view` / `fm_sheet_view_t`, replacing the 16-byte form. The
  `_ex` entry point was native + wasm + npm-native; the base it replaces was
  native + wasm only, so an npm-native consumer sees no change here.
- `fm_workbook_defined_name_at_ex` — folded into
  `fm_workbook_defined_name_at`, which now takes the scope out-param
  directly. Same split as the sheet-view pair: the `_ex` form was on all
  three surfaces, the base was native + wasm only.
- `fm_workbook_save_ex` — renamed to `fm_workbook_save_as`, same signature.
  Native + wasm + npm-native.
- `fm_styles_get_font_ex` and `fm_styles_add_font_ex` — removed earlier in
  this cycle when the font record absorbed the fields they carried; use
  `fm_styles_get_font` / `fm_styles_add_font`, which now return and accept
  the complete `fm_font_record`. **Native only**: both were declared in the
  v0.9.7 header but never exported to wasm and never wrapped by the Node
  addon, so the only consumer affected is a third party linking the header.
- `fm_workbook_save_xlsb_with_result` — replaced by
  `fm_workbook_save_with_diagnostics(wb, FM_WORKBOOK_FORMAT_XLSB, &bytes,
  &len, &diagnostics)`, reading the downgrade count from
  `diagnostics.downgraded_formula_count`. The replacement also reports the
  four other save counters the old entry point discarded. **Native only**:
  it was declared in the v0.9.7 header but never exported to wasm and never
  wrapped by the Node addon, so neither the npm packages nor the Python
  wheel could ever call it.
- `fm_workbook_xlsb_read_diagnostics` — replaced by
  `fm_workbook_read_diagnostics(wb, &diagnostics)`, reading
  `diagnostics.undecoded_formula_count` and
  `diagnostics.undecoded_defined_name_count`. The old dropped-part
  projection is now `diagnostics.undecoded_part_count`, renamed so it can
  no longer be confused with the save-side `dropped_part_count`, which
  counts a different event. Native + wasm; the Node addon never wrapped it.

### Fixed

- A saved PivotTable closes each field's `<items>` with the subtotal entries
  the field displays — `<item t="default"/>` for the implicit subtotal, or one
  token per explicitly selected function. Excel treats a field whose item list
  lacks them as damaged and offers to repair the workbook on open, so every
  package carrying a pivot was affected: one built through the API, and one
  loaded from Excel and saved again, which also dropped the entries Excel had
  written. Nothing short of opening the file reported it — the markers are
  optional in the schema, and the reader skips them deliberately because the
  selection is modelled as `defaultSubtotal` / the `*Subtotal` family rather
  than as items, so a read-write round trip compared equal while shedding
  them.
- A namespace-qualified attribute retained from a consumed part is re-emitted
  together with the binding for its prefix, searched from the element up
  through its ancestors because the declaration usually sits on the part root.
  Excel writes `mc:Ignorable` and `xr:uid` on the pivot parts; re-emitting one
  with nothing binding its prefix produced XML that is not well-formed, so
  loading an Excel-authored workbook with a pivot and saving it yielded a
  package no parser would accept.
- `DATEDIF` matches its unit argument case-insensitively for every documented
  token (`Y`, `M`, `D`, `YM`, `YD`, `MD`). A lowercase or mixed-case spelling
  such as `DATEDIF(a, b, "y")` or `"yM"` returned `#NUM!` where Excel returns
  the interval.
- A spill is blocked when its footprint meets a merged range or another
  formula's spill rectangle, so `=SEQUENCE(2)` entered at a merged `A1:B1`
  reports `#SPILL!` as Excel does instead of writing over the merge. Only a
  region anchored at the requested anchor is ignored — that is the producer
  re-evaluating itself. A blocked spill leaves the sheet untouched apart from
  the anchor's cached `#SPILL!`, and a zero-sized, out-of-grid or
  end-overflowing footprint counts as a collision.
- Whole-axis and spill-derived dependencies are normalized on one path, so a
  partial recalculation reaches the same fixed point as a full one: extents
  are shared rather than recomputed per caller, range bounds are inclusive
  throughout, a spill footprint contributes its own dependency edges, and a
  semantic reindex invalidates the spills it invalidates.
- Structural edits keep formulas addressable. A row or column insert or
  delete reindexes defined names after the physical move rather than before,
  so a moved owner — including a 3-D span owner — is no longer registered at
  its pre-edit coordinates. Sheet rename, removal and reordering rewrite
  workbook-local 3-D spans through a single visitor that walks every formula
  holder.
- Pivot hierarchy and subtotal rows are emitted in one display order, and
  subtotals are counted per owner so two branches sharing a display label
  stay separate.
- The iterative solver no longer retains arena-backed values between passes:
  scalars are copied by value, text is copied byte-exactly into bounded
  solver-owned storage including embedded NUL, and arrays, lambdas and
  references are treated as incomparable so the residual stays infinite.
- OOXML and XLSB readers cap cumulative decompression at 256 MiB per open
  session and report `kIoFileTooLarge` before allocating beyond that budget.
  A reopen and a failed inflate are charged against the same budget.
- `fm_sheet_add_hyperlink`, `fm_sheet_add_merge`, `fm_sheet_set_comment` and
  `fm_sheet_add_validation` reject a rectangle or coordinate outside the
  Excel grid with `kInvalidArgument`. A validation rule's ranges are all
  checked before the rule is stored, so a rejected call leaves the sheet
  unchanged.
- `fm_styles_add_batch` stages the complete table and commits it in one step,
  so a failure leaves both the workbook and the caller's output arrays
  untouched, and `fm_styles_add_num_fmt` reports `kPreconditionFailed` when
  the 16-bit custom number-format id space is exhausted instead of wrapping.
- Every optional `xf` alignment attribute round-trips. `textRotation`,
  `indent`, `relativeIndent`, `shrinkToFit` and `readingOrder` each sit
  behind a presence flag, so an omitted attribute stays distinct from an
  explicit zero or false, and an explicitly empty `<alignment/>` or an
  explicit schema default (`horizontal="general"`, `wrapText="0"`) survives
  instead of collapsing away. A malformed attribute value returns
  `kIoSheetCorrupt` naming the table, xf index and attribute rather than
  reading as a default.
- `justifyLastLine` is read from `<alignment>` and written back. The
  attribute was parsed past and never emitted, so it was dropped from every
  imported workbook on save.
- A table part emits `<autoFilter>` and `<sortState>` before
  `<tableColumns>`, the order `CT_Table` declares; the previous order
  produced a part Excel flagged as needing repair.
- Every `pivotTable` part gets the `pivotCacheDefinition` relationship part
  it requires, so a consumer navigating the package by relationship alone can
  reach the cache a table draws from.
- `fm_sheet_set_auto_filter_xml` validates the fragment with a real XML parse
  — exactly one top-level element, named `autoFilter` — instead of a
  prefix/suffix shape check, so a malformed or truncated element is rejected
  rather than written into a package Excel refuses to open. Criteria and
  extension payloads are still preserved verbatim.
- The CLI resolves a symbolic link before its temp-then-rename write, so
  saving through a link updates the workbook the link names instead of
  replacing the link with a regular file and leaving the real workbook stale.
  A plain path, a dangling link or a failed resolution is replaced as-is.
- `eval`, `dump` and `paginate` report a failed write of their primary result
  as a nonzero exit status, so exit 0 means the complete result reached the
  output stream.
- The formula length cap bounds a single oversized token. A string literal,
  identifier or quoted sheet name longer than the cap was consumed whole
  because the cap was only re-checked between tokens; the input is now
  trimmed on a codepoint boundary before any scanner runs, an input of
  exactly the cap length is still accepted, and a truncation that lands
  mid-token reports `ExcessiveLength` rather than passing silently.
- Range-sourced arguments are filtered identically on both evaluation paths,
  and a scalar argument's error is detected at its own slot before later
  range arguments are flattened, so `SUM(1/0, A1)` and `SUM(A1, 1/0)` pick
  the same error whichever path evaluates them. The same filter applies to
  the per-group slice `GROUPBY` and `PIVOTBY` hand to a bare aggregate.

## [0.9.7] - 2026-08-06

### Added

- `formulon paginate`, a CLI subcommand that prints the resolved print areas,
  row and column page breaks, and page count for one sheet. It is backed by
  `fm_workbook_paginate` with an owned `fm_pagination_t`, and reaches WASM,
  Node and Python as `paginate`.
- Workbook memory-footprint estimate: `Workbook::approximate_memory_bytes()`
  and `fm_workbook_memory_usage` report an `O(cells)` pressure signal covering
  the cell store, the shared-string storage every `Text` value borrows from,
  the passthrough part payloads and the workbook metadata. The Node addon
  reports the delta to V8 as external memory on create, load, recalc and
  handle destruction, and exposes `memoryUsage()`.
- `fm_styles_add_batch` installs and deduplicates fonts, fills, borders, cell
  xfs and number formats in one call, ordering the tables so an xf can
  reference indices produced by the same batch.
- `fm_error_display_name` (`errorDisplayName` / `error_display_name`) for the
  Excel literal of a cell error code, and `fm_workbook_xlsb_read_diagnostics`
  for the undecoded formula and defined-name counters captured during an XLSB
  load.
- Phonetic text, iterative-calculation settings, extended font records
  carrying `vertAlign`, and `fm_workbook_save_xlsb_with_result`. All are
  additive, so the existing ABI is unchanged. The binding surface gains
  `getCommentResult`, conditional-format visual rules, differential formats,
  comment enumeration and pivot cache metadata alongside them.
- XLSB writer: a native `xl/styles.bin` emitter, round-tripped row/column
  layout, merged rectangles and the `date1904` flag, the mandatory worksheet
  prefix with view flags, zoom and frozen panes, and an `xl/metadata.bin`
  carrying the dynamic-array entry. Worksheet-tail records the model does not
  express — conditional formatting, data validation, hyperlinks, auto-filter,
  print setup, breaks and the drawing / table part references — are retained
  verbatim along with their sheet relationships.
- OOXML writer interns literal text cells into a generated
  `xl/sharedStrings.xml` and writes those cells as `t="s"`.
- Parser accepts a 3-D whole-column (`Sheet1:Sheet3!A:A`) or whole-row span,
  and treats a space before a quoted sheet name or a parenthesized range as
  the intersection operator.
- `GROUPBY` and `PIVOTBY` honour a `total_depth` of ±2, emitting one subtotal
  row per outer group with the outer key restated in the first key column.
  Pivot items sort by a data field's aggregate when `SortSpec::by_field` is
  set.
- A configurable process-wide structured-log sink and minimum severity level.

### Fixed

- Number-format colour names are read in the UI locale: the ja-JP profile
  accepts the localized names and the indexed `[色N]` form, while the English
  `[Red]` / `[ColorN]` spellings surface `#VALUE!` exactly as Excel does.
- Dynamic-array and lambda semantics: `BYROW` / `BYCOL` spill an errored slice
  into its own output cell instead of collapsing the call, `REDUCE` and `SCAN`
  hand the body an errored cell verbatim so an `IFERROR` guard can recover,
  and every `ArrayValue` allocation routes through a bounds-checked seam that
  validates each axis against the Excel grid.
- Numeric and statistical edge cases: `T.INV`, `F.INV` and `BETA.INV` invert
  by bisection over an expanding bracket, `YIELD` and `ODDFYIELD` clamp to the
  yield domain `PRICE` accepts, `DDB` / `DB` / `VDB` cap the schedule length,
  `LINEST` reports a finite F for a perfect fit, `FORECAST.ETS` detrends
  before detecting seasonality, and non-finite results surface as `#NUM!`.
- Number-format rounding and text functions: a shared rounding helper replaces
  ad hoc rounding in `FIXED`, `DOLLAR`, `BAHTTEXT` and the numeric renderer,
  the 32,767 UTF-16 unit cap is enforced in `SUBSTITUTE`, and the half-to-full
  katakana voicing tables are deduplicated.
- Date builtins reject serials outside Excel's `0`..`9999-12-31` range and
  broadcast array arguments cell by cell.
- `AREAS` evaluates the `INDIRECT` call it is asked to count, so a resolution
  counts as one area and a failure propagates as that error.
- Parser: the depth limit is validated against the completed AST so a flat
  left-associative chain is covered, a malformed UTF-8 lead byte is consumed
  instead of spinning the tokenizer forever, and `$`-bearing identifier runs
  such as `A$$1` are rejected.
- Structural edits propagate across every referencing model — hyperlink
  locations, data-validation formulas, pivot-cache sources, conditional-format
  and table ranges — and formulas are re-indexed when an edit rewrites a
  defined name.
- OOXML read path preserves unmodelled parts: chart, dialog and macro sheets
  come through as opaque sheets, package-level and per-sheet relationships of
  unrecognised types are re-emitted with zip-slip validation, and workbook and
  worksheet elements the model does not express survive the next save
  verbatim. XML-invalid C0 controls are escaped per context.
- XLSB Ptg codec emits reference-class Ptgs for cell and range arguments and
  value-class Ptgs for function results, decodes the `PtgMem*` markers, and
  encodes `BrtColor` with `fValidRGB` in bit 0.
- Pivot layout projection through the C API honours the workbook's Excel
  profile and the selected Compact, Tabular or Outline report layout, so the
  default ja-JP profile emits localized labels instead of the legacy English
  grid.
- Cyclic component members are ordered by address before iterating, so the
  Gauss-Seidel solver commits in a fixed order across standard-library
  implementations rather than following DFS pop order.
- A DataBar rule whose min and max thresholds are equal renders a full bar.
- CLI: `eval`, `dump` and `recalc` share one atomic file-I/O path, exit codes
  collapse to `{0, 1, 64}`, and `eval` runs through the read-only array C API
  so `=A1+1` sees an empty `A1`.
- Recalc workers launch through `launch_thread`, which returns
  `Expected<Thread, Error>` instead of terminating when the OS refuses a
  thread; the pool keeps whichever workers started and falls through to serial
  evaluation when none do.
- The evaluation and load-time arenas carry byte ceilings, so a hostile
  formula degrades to a per-cell error instead of growing until the process or
  the WASM host page aborts.

### Performance

- Whole-axis references are tracked as compact rectangle dependencies instead
  of promoting the formula to volatile, and Tarjan runs over the induced
  subgraph of dirty cells rather than the workbook-wide graph.
- Recalc layer workers are pooled behind a barrier-synchronized
  `LayerWorkerPool` for the whole parallel pass instead of being spawned and
  joined per layer.
- Row storage became `RowCells`, a contiguous run beginning at the row's first
  populated column, so memory scales with content rather than used width.
  `Sheet::read_range` appends a whole rectangle under one lock acquisition.
- The shared-string and pivot-record parts, the OOXML worksheet parse and the
  metadata shell parse run through an in-place XML loader, halving peak parse
  memory, and the workbook solely owns the passthrough payload instead of
  mirroring it.

### Changed

- The bytecode compiler, optimizer and VM compile only when
  `FORMULON_BUILD_VM` is on (defaulting to `FM_BUILD_TESTING`), so release
  CLI, WASM and binding binaries no longer carry the experimental pipeline.
- `ZipReader` drops the per-archive cumulative extraction cap; the zip-bomb
  guard narrows to the per-entry size, entry-count and compression-ratio caps.
- Per-file license and copyright headers are removed; the terms live in the
  top-level `LICENSE`.

### Testing

- Divergence and coverage governance: `divergence_check.py` validates every
  entry against a real case, suite, alias or documented non-oracle scope, and
  `golden_coverage_check.py` fails when a declared case has no golden. Each
  secondary-oracle golden file is registered as its own ctest entry so one
  allowlist exception can no longer mask every failure.
- Goldens recaptured against Excel 365 ja-JP 16.111.2, with the
  `cross_sheet_refs`, `intersect_operator`, `iterative_calc` and
  `spill_collision` suites added.
- The cross-language parity harness returns ctest's skip code when fewer than
  two channels are active, instead of reporting success on one.
- An XLSB libFuzzer target with a portable Ptg seed format, and source-seam
  guards that fail when a second `ArrayValue` allocation site or raw-XML
  retention implementation appears.

### Build / CI

- A `native-fast` job runs the fast ctest labels on develop pushes, so a red
  develop surfaces before the promotion to main.
- The WASM size report gates the Brotli wire size (768 KiB hard, 640 KiB soft)
  on equal footing with the uncompressed size; a host without `brotli` on
  `PATH` skips that half instead of failing it.
- CI fails on `expected_flakes.txt` entries that no longer name a registered
  test.

### Documentation

- The READMEs document the CLI commands, both WASM size ceilings, and the
  WASM worksheet parsing memory profile.

**Detailed Release Notes**: [GitHub Release](https://github.com/libraz/formulon/releases/tag/v0.9.7)

## [0.9.6] - 2026-07-19

### Added

- Full Excel 365 dynamic-array spill semantics. Bare ranges, arithmetic and
  comparison operators, and `IF` conditions now spill to their argument
  shape, matching Excel's implicit array evaluation; bare-range spills fill
  blanks with `0`. Scalar functions also spill each range argument element-
  wise instead of reducing to the top-left cell.

### Fixed

- Bounded every attacker-driven work path with a request-scoped budget across
  evaluation, conditional formatting, pivot, and I/O, so a hostile workbook
  can no longer force unbounded computation.
- Rejected out-of-grid coordinates at the public C API and core entry points,
  out-of-grid `Print_Titles` repeat spans, and out-of-range pivot-cache field
  indices.
- Hardened OOXML/XLSB part names, detected encrypted containers, and bounded
  `BrtExternSheet` reservation to the available payload; validated row /
  column / array bounds when decoding XLSB sheet records; cast the XLSB
  `ExternSheet` reserve count to `size_t` for the 32-bit WASM build.
- Hardened the parser against literal postfix-call, exponent overflow, and
  out-of-memory name interning.
- Aligned the VM's lexical scope with the tree-walker and made it fail closed
  on invalid opcodes.
- Recalc dependency correctness: rebuild the dependency graph on sheet
  permutation and gate evaluation on a strict parse; rewrite cell references
  on sheet rename and surface off-grid spills as `#SPILL!`; track direct
  lambda-call body dependencies and invalidate on name retarget or
  spill-phantom write; invalidate a memoized pivot layout when its source
  cache mutates.
- Reconciled the raw x14 conditional-format overlay when CF rules are removed.
- Measured the sheet-name length limit in code units and validated it on add.
- Canonicalized and localized every enumerated function through the C API.
- `formulon_cli recalc` now writes its output atomically, so a failure no
  longer destroys the original workbook; `eval --repeat` re-evaluates on each
  pass instead of being a no-op.

### Performance

- Extract zip entries into a single caller-owned buffer.

### Build / CI

- Added a repo-wide formatter / linter (biome + ruff) with a `make format`
  fan-out and a CI check-only counterpart.

**Detailed Release Notes**: [GitHub Release](https://github.com/libraz/formulon/releases/tag/v0.9.6)

## [0.9.5] - 2026-07-04

### Added

- Ad-hoc array evaluation: `evaluateFormulaArray` (C API, Node addon,
  WASM) and `evaluate_formula_array` (Python) evaluate a dynamic-array or
  spilled formula against a workbook without mutating it, returning the
  whole `Array` result instead of reducing to its top-left element. The C
  ABI exposes a two-step surface — `fm_workbook_evaluate_formula_array`
  stashes the result on the handle, `fm_workbook_evaluate_formula_array_cell`
  reads it back by row-major index.
- Function-metadata provider seam: hosts can inject localized function
  documentation over the engine's structural catalog. `fm_function_metadata`
  now also recognizes lazy-dispatch forms (`XLOOKUP`, `SUMIFS`, …) and
  parser special forms (`LET`, `LAMBDA`), and surfaces the unbounded-arity
  sentinel as `null` / `None`. Pure merge helpers `mergeFunctionMetadata`
  (Node) and `merge_function_metadata` (Python) resolve
  signature/description/localized name by locale-override → entry-default →
  engine-value precedence. Contract documented in
  `docs/function-metadata-schema.md`.

### Fixed

- Spill-phantom fidelity: `Sheet::spill_phantom_addresses` enumerates
  phantom cells across all spill regions; cell enumeration and `cell_count`
  now fold them in, `fm_workbook_cell_at` no longer treats a phantom
  coordinate as an internal error, and pagination computes the used range
  against a spilled region's full extent rather than just its anchor.
- Range-shaped defined names (e.g. `Sheet1!$A$1:$A$5`) now evaluate as a
  `Value::Array` instead of collapsing to a scalar through implicit
  intersection.

### Build / CI

- Add a `python-smoke` job to `prebuild.yml` to catch C ABI drift before
  release.

**Detailed Release Notes**: [GitHub Release](https://github.com/libraz/formulon/releases/tag/v0.9.5)

## [0.9.4] - 2026-07-03

### Added

- Read-only ad-hoc formula evaluation: `evaluateFormulaText` /
  `evaluateConditionalFormula` (C API, Node addon, WASM) evaluate formula
  text against a workbook without mutating it — resolving local/
  cross-sheet refs, defined names, and `ROW()`/`COLUMN()` anchoring, and
  reducing array/spill results to their top-left element. CF-rule
  evaluation shifts relative references from the rule's anchor and
  applies Excel's CF-predicate coercion.
- Comment enumeration: `getComments` (Node addon) / `fm_sheet_get_comment_count`
  and `fm_sheet_get_comment_at_index` (C API) list every comment on a
  sheet, including comments anchored on otherwise-empty cells.
- `show_dropdown` on data validation, round-tripped through OOXML with
  the `showDropDown` attribute's inverted semantics corrected.
- `addConditionalFormat` / `fm_sheet_cf_add_rule` now return the new
  rule's flattened index.

See the
[GitHub release page](https://github.com/libraz/formulon/releases/tag/v0.9.4)
for the full auto-generated change list.

## [0.9.3] - 2026-07-03

### Added

- Binding-surface parity across C API, Node addon, WASM, and Python for
  pivot-cache worksheet-source/layout, sheet-view display/orientation
  flags, `save_ex` (XLSX/XLSB selector), a static cell-error setter,
  sheet-scoped defined names, CF `ColorScale`/`DataBar`/`IconSet`
  payloads, and dxf differential-format record reads. CLI `recalc`/`dump`
  pick up the same surface (extension-driven XLSB output, sheet-scoped
  name printing).
- Conditional formatting: whole-row/whole-column `sqref` support, x14
  data-bar overlay decoding (gradient, axis position, negative
  fill/border), and verbatim `extLst` passthrough.
- XLSB reader/writer closes binary-format protocol gaps: styles
  (`BrtFmt`/`BrtXF`), workbook-scope names including future functions and
  `LET`, cross-sheet 3-D references, and dynamic-array spill formulas
  (`BrtArrFmla`) — several of these previously produced `.xlsb` files
  that real Excel could not open.
- OOXML round-trip fidelity: `workbookPr`/`bookViews`/
  `workbookProtection`, `date1904`, Default-content-type passthrough,
  table style info, and per-cell style color specs (theme/indexed) all
  survive a load-modify-save cycle on real Excel-authored workbooks.
- Evaluator/parser: `date1904` threaded through the tree-walker and VM,
  defined-name resolution with circular-reference detection, whole-
  column/row range expansion against the sheet's used range, Excel's
  actual array-broadcast rule, and 3-D range tails
  (`Sheet1:Sheet3!A1:B2`).

### Fixed

- `PIVOTBY`/`GROUPBY` grand totals re-aggregate correctly for
  non-additive functions (Average/Max/Min/StdDev/Var); multi-value
  grand-total column blanking restored.
- Pivot `ShowValuesAs` (RunningTotal direction, Index, DifferenceFrom/
  PercentOf) and `GETPIVOTDATA` field-name resolution fixed.
- Conditional-formatting text rules now match numeric and blank cells
  via General-format coercion, matching Excel's SEARCH-based generated
  formula.
- Print pagination excludes hidden rows/columns from the pagination
  extent, makes column-break counting symmetric with row-break
  handling, and no longer mis-splits print-area/print-titles tokens on
  a quoted sheet name containing a comma.
- `AREAS` recurses into `CHOOSE`/`IF` reference branches instead of
  always returning 1; `INDEX`/`XLOOKUP`/`INDIRECT` range results route
  through the dynamic-array allocator instead of collapsing to a
  scalar.
- `IFNA` no longer promotes `Blank` to `0`; volatile-function detection
  is case-insensitive.

### Changed

- A shared value-kind rank centralizes `GROUPBY`/`PIVOTBY`/`SORT`
  ordering so the three comparators cannot diverge.
- WASM size ceiling raised from 1.9 MiB/600 KB soft to 3.0 MiB hard /
  2.5 MiB soft (current binary: 2.09 MiB uncompressed / 560 KiB
  Brotli).
- GitHub Actions workflows bumped to Node 24-compatible major versions
  (`actions/checkout@v7`, `setup-node@v6`, `setup-python@v6`,
  `cache@v6`, `upload-artifact@v7`, `download-artifact@v8`,
  `codecov-action@v6`, `setup-emsdk@v16`, `action-gh-release@v3`),
  ahead of GitHub's Node 20 removal.

See the
[GitHub release page](https://github.com/libraz/formulon/releases/tag/v0.9.3)
for the full auto-generated change list.

## [0.9.2] - 2026-05-18

### Added

- Workbook oracle track for pivot tables and print/pagination, driven
  through a WSL2->Windows Excel bridge with the `win-365-ja_JP` profile
  as primary. Mac and Windows tracks share a single comparator.
- Function-availability classification distinguishing win-365-only from
  cross-version functions; win-365 mode-switch semantics now match the
  oracle for the primary `win-365-ja_JP` profile.
- Print pagination captures margins and `PageBreakPreview` reads,
  retiring three divergence skips.

### Fixed

- Parser truncates numeric literals to Excel's 15-significant-digit
  representation.
- `ARRAYTOTEXT` propagates a scalar error argument through to the result.
- `PIVOTBY` layout, `MAP` / `MAKEARRAY` error spills, `FREQUENCY` bin
  ordering, `WRAPROWS` / `WRAPCOLS` shape, and `TRIMRANGE` blank
  handling align with the Mac Excel oracle.
- Print pagination suppresses auto-column breaks; the min-title-reserve
  floor avoids inverted-scale page-break drift at scale <= 50.
- OOXML round-trip preserves unknown workbook rels and shifts
  shared-formula refs correctly across cell moves (close residual cases
  from v0.9.1).
- `PERCENTILE.EXC` at the upper boundary (`pos == n`) now consistently
  returns `#NUM!`, matching Mac Excel. Previously one of two internal
  code paths returned the largest sample value.

### Changed

- Internal refactors split 12 large translation units into per-area
  files, extracted opcode metadata, introduced a binding-codegen
  pipeline for simple passthroughs, and added a shared `tests/util`
  library used by 60 test files. No user-visible API change.
- Further consolidation deduplicates aggregate kernels shared by
  `SUBTOTAL` and `AGGREGATE`, numeric-argument helpers, RGB-hex
  parsing, sheet-index validation, and OOXML/XLSB relative-path
  resolution; splits `ooxml_writer.cpp`, `number_format_tokenizer.cpp`,
  and `stats.cpp` into per-area files. No user-visible API change.

See the
[GitHub release page](https://github.com/libraz/formulon/releases/tag/v0.9.2)
for the full auto-generated change list.

## [0.9.1] - 2026-05-11

### Added

- OOXML round-trip of worksheet print settings (`pageSetup`, `pageMargins`,
  `headerFooter`, `printOptions`) and pass-through of opaque
  `printerSettings.bin` parts.

### Fixed

- OOXML round-trip preserves unknown workbook relationships and shifts
  shared-formula refs correctly across cell moves.

See the
[GitHub release page](https://github.com/libraz/formulon/releases/tag/v0.9.1)
for the full auto-generated change list.

## [0.9.0] - 2026-05-11

First public release.

### Distribution

- **npm** `@libraz/formulon`: Excel 365 calculation engine as a WebAssembly
  module. Browser, Node.js, and Web Worker compatible via embind.
- **PyPI** `formulon`: pure-Python `py3-none-any` wheel that drives a
  bundled `formulon_capi.wasm` through
  [`wasmtime`](https://pypi.org/project/wasmtime/). One wheel works on
  every platform `wasmtime` supports (Linux x86_64 / aarch64, macOS
  x86_64 / arm64, Windows x86_64).
- **CLI**: native `formulon_cli` binaries for macOS arm64, Linux x86_64,
  and Linux arm64, attached to the GitHub release as `.tar.gz`.

See the
[GitHub release page](https://github.com/libraz/formulon/releases/tag/v0.9.0)
for the full auto-generated change list.

[Unreleased]: https://github.com/libraz/formulon/compare/v0.19.0...HEAD
[0.19.0]: https://github.com/libraz/formulon/compare/v0.18.0...v0.19.0
[0.18.0]: https://github.com/libraz/formulon/compare/v0.17.0...v0.18.0
[0.17.0]: https://github.com/libraz/formulon/compare/v0.16.1...v0.17.0
[0.16.1]: https://github.com/libraz/formulon/compare/v0.16.0...v0.16.1
[0.16.0]: https://github.com/libraz/formulon/compare/v0.15.2...v0.16.0
[0.15.2]: https://github.com/libraz/formulon/compare/v0.15.1...v0.15.2
[0.15.1]: https://github.com/libraz/formulon/compare/v0.15.0...v0.15.1
[0.15.0]: https://github.com/libraz/formulon/compare/v0.14.0...v0.15.0
[0.14.0]: https://github.com/libraz/formulon/compare/v0.13.0...v0.14.0
[0.13.0]: https://github.com/libraz/formulon/compare/v0.12.0...v0.13.0
[0.12.0]: https://github.com/libraz/formulon/compare/v0.11.1...v0.12.0
[0.11.1]: https://github.com/libraz/formulon/compare/v0.11.0...v0.11.1
[0.11.0]: https://github.com/libraz/formulon/compare/v0.10.0...v0.11.0
[0.10.0]: https://github.com/libraz/formulon/compare/v0.9.7...v0.10.0
[0.9.7]: https://github.com/libraz/formulon/compare/v0.9.6...v0.9.7
[0.9.6]: https://github.com/libraz/formulon/compare/v0.9.5...v0.9.6
[0.9.5]: https://github.com/libraz/formulon/compare/v0.9.4...v0.9.5
[0.9.4]: https://github.com/libraz/formulon/compare/v0.9.3...v0.9.4
[0.9.3]: https://github.com/libraz/formulon/compare/v0.9.2...v0.9.3
[0.9.2]: https://github.com/libraz/formulon/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/libraz/formulon/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/libraz/formulon/releases/tag/v0.9.0
