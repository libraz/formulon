# Changelog

All notable changes to Formulon are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

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
  column, and one entered as the formula spills at that size. Aggregates
  still read the cached extent.
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

- Full binding-surface parity across C API, Node addon, WASM, and Python:
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

[Unreleased]: https://github.com/libraz/formulon/compare/v0.13.0...HEAD
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
