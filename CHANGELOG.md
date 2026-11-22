# Changelog

All notable changes to Formulon are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.19.0] - 2026-11-22

0.19 makes the C ABI, the two JavaScript packages and the Python package one
interface with the same names, records and error model, and brings file output,
calculation and formula entry closer to Excel. Almost every 0.18 program needs
changes, most of them renames, and a C caller must recompile.

Upgrading from 0.18: see [Upgrading from 0.18 to 0.19](docs/releases/upgrading-0.19.md).

### BREAKING

- **Bindings.** Names follow one convention in C, JavaScript and Python, and
  the former names are removed with no aliases
  ([name migration](docs/releases/upgrading-0.19.md#name-migration)).
- Sheet-scoped calls move to a `Sheet` handle (`wb.sheet(key)`), and `Table`,
  `Pivot`, `PivotCache` and `Drawing` are handle classes that `list*` calls
  return ([sheet handles](docs/releases/upgrading-0.19.md#sheet-handles)).
- Failures throw `FormulonError` in JavaScript and raise `FormulonError` or a
  subclass in Python, sorted by what the argument addresses; there are no
  status envelopes ([every binding](docs/releases/upgrading-0.19.md#every-binding)).
- Status codes are renumbered, retired or removed and named `FM_STATUS_*` /
  `StatusCode`; several failures change status, and engine ceilings, `.xlsb`
  loads included, report `SecurityResourceLimit` with a `part=` context
  ([every binding](docs/releases/upgrading-0.19.md#every-binding)).
- `CellRect` replaces the per-feature range records, absent reads `null`, units
  are named by suffix, collections use one `list*` method and `update<X>(patch)`,
  and several enums are merged or renumbered
  ([method shape changes](docs/releases/upgrading-0.19.md#method-shape-changes), [enums](docs/releases/upgrading-0.19.md#enums)).
- Strings: an empty string is no longer an absent value, so it is refused where
  Excel cannot store it, and a lone UTF-16 surrogate fails with
  `BindingUtf8EncodingError` instead of being stored as U+FFFD
  ([empty strings](docs/releases/upgrading-0.19.md#empty-strings), [every binding](docs/releases/upgrading-0.19.md#every-binding)).
- C ABI: sheet functions take `size_t sheet_index`, counts come through
  out-parameters, input records are `const` with `struct_size`, and
  `fm_workbook_load` and the chart and shape records change. Recompile
  ([C ABI changes](docs/releases/upgrading-0.19.md#c-abi-changes)).
- JavaScript: `createFormulon()` resolves to `Formulon`, a value is a union on
  `kind`, `setValue` reads a leading `=` as a formula, malformed arguments throw
  `TypeError`, and the declarations need TypeScript 5.7
  ([JavaScript and TypeScript callers](docs/releases/upgrading-0.19.md#javascript-and-typescript-callers)).
- Python: tuple results are frozen dataclasses, some signatures take a record
  in place of keywords, positions are code points, `FormulonError` carries
  `.code`, and Python 3.10 is required
  ([Python callers](docs/releases/upgrading-0.19.md#python-callers)).
- A cell holding an in-cell picture reads as a new value kind,
  `ValueKind.Picture` (C `FM_VALUE_KIND_PICTURE`), where it read `#VALUE!`; a
  `switch` over every kind needs a new case
  ([picture values](docs/releases/upgrading-0.19.md#picture-values)).
- Setters refuse what Excel refuses to store or open: overlapping or colliding
  merges, tables and names, names outside Excel's syntax, format defaults,
  protected ranges, sparkline sources and chartsheet page settings
  ([behaviour changes](docs/releases/upgrading-0.19.md#behaviour-changes-under-an-unchanged-name)).
- Setters also refuse excluded style kinds, a changed cell-style id, repeated
  conditional-format ids, colour-scale automatic thresholds, colours without a
  kind, repeated pivot names, bad `sortByField` targets and numeric date groups
  ([behaviour changes](docs/releases/upgrading-0.19.md#behaviour-changes-under-an-unchanged-name)).
- Row and column edits take `(first, last)`; an insert that would push content
  off the grid fails, and conditional formats and validations grow over cells
  inserted next to them
  ([method shape changes](docs/releases/upgrading-0.19.md#method-shape-changes)).
- Cell storage: a blank cell exists only while it has a format of its own,
  `setRangeXfIndex` over whole rows or columns writes row or column formats,
  and `snapshotRangeXfs` writes capture version 2
  ([behaviour changes](docs/releases/upgrading-0.19.md#behaviour-changes-under-an-unchanged-name)).
- Hyperlinks are keyed by their range: `setHyperlink(link)` takes a required
  `range`, overlapping links coexist, setting one applies the Hyperlink cell
  style and the target is stored as Excel stores it
  ([behaviour changes](docs/releases/upgrading-0.19.md#behaviour-changes-under-an-unchanged-name)).
- Reads and defaults follow Excel: fonts report their theme face, family and
  charset, format codes use the profile's vocabulary, validation booleans
  default to true, and `activeSheet()` follows Excel's open rule
  ([behaviour changes](docs/releases/upgrading-0.19.md#behaviour-changes-under-an-unchanged-name)).
- **Formulas.** Formula entry refuses what Excel refuses, among them argument
  counts, the newer error literals, non-reference `:` endpoints and a typed
  `ANCHORARRAY` call; stored formulas still load
  ([formula entry](docs/releases/upgrading-0.19.md#formula-entry)).
- Formula text takes Excel's spelling on entry and `.xlsx` load (`=sum(a1)`
  reads `=SUM(A1)`), with a leading `=` and `@` structured references, and a
  sheet qualifier naming no sheet adds an external link
  ([formula entry](docs/releases/upgrading-0.19.md#formula-entry)).
- A loaded workbook keeps its cached values when Excel would: a manual book, or
  an automatic one saved by the current Excel, recomputes only what its file
  marks stale. `recalcFull()` recomputes everything
  ([calculation on load](docs/releases/upgrading-0.19.md#calculation-on-load)).
- `LEN`, `MID`, `FIND`, `SEARCH` and `REPLACE` count characters by the
  workbook's Compatibility Version, so under the default Windows profile a new
  workbook counts a character outside the BMP once
  ([behaviour changes](docs/releases/upgrading-0.19.md#behaviour-changes-under-an-unchanged-name)).
- **Files.** An `.xlsb` save of a feature the binary records cannot hold fails
  with `IoFeatureNotWritable` and writes nothing, where it left the feature
  out, and so does `setCellPicture` on a book keeping its rich values as stored
  ([`.xlsb` saves that fail](docs/releases/upgrading-0.19.md#xlsb-saves-that-fail)).
- **PivotTables and tables.** A pivot report growing over occupied cells is
  left blocked with `#SPILL!`, `setAnchor` acts as cut and paste, a
  `refreshOnLoad` cache refreshes on open, and a deleted source keeps its cache
  ([behaviour changes](docs/releases/upgrading-0.19.md#behaviour-changes-under-an-unchanged-name)).
- `addTable` numbers blank and repeated column names instead of refusing them,
  typing into a header renames the column, and the header and totals toggles
  write and delete their cells
  ([behaviour changes](docs/releases/upgrading-0.19.md#behaviour-changes-under-an-unchanged-name)).
- **Command line.** `dump --view` replaces the per-view flags, plain output
  escapes only newline, carriage return, tab and backslash, `paginate` prints
  new keys, and a usage error exits 64 and a formula syntax error 65
  ([command-line users](docs/releases/upgrading-0.19.md#command-line-users)).

### Added

- **API.** Typed and native reads (`getTypedValue`, `getTypedValues`,
  `getDate`, `getUsedRange`, `iterRows`, `evaluateTypedFormula`), `*Input`
  records that accept what the getters return, and `Workbook.dispose()`.
- `parseReference` / `formatReference` convert reference text,
  `localizeFormula` / `canonicalizeFormula` handle the profile's R1C1 spelling,
  and `FormulonError.parseError` reports where a formula fails to parse.
- `localizeDefinedName` and `resolveDefinedNameKey` give the spelling and
  scope a defined name is shown and stored under.
- `convertCivilToSerial` / `convertSerialToCivil` convert dates in the 1900
  and 1904 systems, and `listProfileIds()` lists the behaviour profiles.
- `Workbook.load(bytes, { profileId })` loads under a profile, and a workbook
  location (`getLocation` / `setLocation`) gives `CELL("filename")`, link
  targets and formula text the folder and file name Excel would use.
- `getCompatibilityVersion` / `setCompatibilityVersion` read and set Excel's
  Compatibility Version, which `.xlsx` and `.xlsb` read and write.
- `recalcFull()` / `recalcFullParallel()` recompute every formula,
  `Sheet.isCellStale` reports a formula not computed since its inputs changed,
  and `listCircularReferences()` lists the cells in an unresolved cycle.
- `applyTheme(bytes)` / `saveTheme()` apply and save a `.thmx` theme.
- `Workbook.copySheet` duplicates a worksheet as Excel's Move or Copy does,
  with its names, tables, PivotTables, pictures and charts.
- `Sheet.insertCells` / `deleteCells` shift a band of cells in a `CellShift`
  direction, and rows inserted inside a table take its calculated columns.
- Chartsheets load, save and convert between `.xlsx` and `.xlsb`; `Sheet.kind`
  tells them apart, `Sheet.getChart()` reads the chart, and their protection,
  view, page setup, margins and header and footer are readable and settable.
- In-cell pictures (Excel's Place in Cell): `getCellPicture` / `setCellPicture`
  read and set them, a formula carries the picture as a value (`TYPE` is 128),
  and `ExcelPicture` stores one through `setValue`.
- `insertHyperlinkAt` / `removeHyperlinkAt` edit the hyperlink list by
  position, and `removeHyperlink(range)` removes every link it intersects.
- `listCellXfIndices` and `listShowPhoneticCells` list formatted blanks, and
  `setRowAutoHeight` and `getArrayFormula` / `setArrayFormula` are added.
- `getSpill` reports a spill blocked with `#SPILL!`, with `Spill.cause` and
  `Spill.anchor`, and `Pivot.getSpill()` reports the range a pivot report needs.
- **Formulas.** `ECMA.CEILING`; `GROUPBY` takes its field relationship argument,
  and `GROUPBY` / `PIVOTBY` accept a total depth beyond 2 and a relative-to
  argument.
- **Files.** Saved `.xlsx` and `.xlsb` files carry Excel's rich error values
  (`#SPILL!`, `#CALC!`, a `#VALUE!` with a reason), and loading reads them back.
- `.xlsb` gains what only `.xlsx` had: pictures and charts, PivotTables read
  into the model, notes and threaded comments, AutoFilter and sort state, print
  and phonetic settings, named cell styles and custom table styles.
- Custom views survive a save between `.xlsx` and `.xlsb`, AutoFilter included.
- **PivotTables and tables.** `PivotCache.refresh()` rebuilds a cache from its
  source and re-projects its pivots, moving cell formats with their items, and
  `getWorksheetSourceRange()` reports what the source resolves to.
- Pivot grouping as Excel's Group command does it: `setFieldNumericGroup`,
  date groups with labels in every UI locale and the Japanese calendar, read
  back through `getFieldDateGroup` / `getFieldNumericGroup`.
- Pivot report options (subtotals, blank rows, repeated labels, `showDetail`,
  sorting, captions, `showAll`), calculated fields and items, loaded filters of
  every type, `Pivot.showDetails` and `PivotCell.expansion` for the buttons.
- `Table.renameColumn`, `Table.insertRows` / `insertColumns`, and
  `Table.update({ columns })` for totals and calculated columns.
- **Conditional formatting and styles.** `setCfRulePriority` places a rule at a
  priority, rule calls take a `CfRuleRef` (an id or a list position),
  `getCellFormat` reports a format's source, and a rule takes a `format` patch.
- **Drawings and charts.** Chart paint reads a `DrawingColor` with its base
  colour and transforms; chart and plot area paint, line ends and the system and
  long dash presets are readable and settable.
- Shapes read and write connector glue, preset adjust values, text direction
  and flips; pictures read and write rotation and flips.
- Every drawing has `hidden`, `description` and `name`; shapes and slicers take
  a z-order, chart patches gain minor units and gridlines, label position, style
  and "Show #N/A as an empty cell", and sparkline groups carry `rightToLeft`.

### Changed

- **API.** Error messages name the call and the parameter path first, and
  Python runs one WebAssembly instance per process, so calls are serialised.
- **Formulas.** `SIN`, `COS`, `TAN`, `EXP`, `LN`, `LOG`, `POWER` and their
  relatives no longer use the host math library: a result can differ from
  0.18.0 in the last bit and is the same on macOS, Linux and WebAssembly.
- A 3-D span is recognised in Excel's shapes and quoted as one unit, and moving
  a sheet past the other endpoint shrinks the span as Excel does.
- A built-in defined name (`Print_Area`, `_FilterDatabase`) is reached and
  stored as Excel does when typed without its prefix or in the profile's
  language, and `getFunctionMetadata` reports the argument counts entry accepts.
- **Files.** Saves write styles as Excel does: unused and duplicate cell formats
  and dxfs are dropped and renumbered, built-in styles take the profile
  language's names, and rule blocks follow Excel's order.
- Opening a file merges conditional-format rules and rewrites their ranges as
  Excel does on open, so a sheet can list fewer rules than the file holds.
- A save writes only the external links something references, numbered by
  first use, and a cached sheet with no data reads `#REF!`.
- A save under a `ja_JP`, `zh_CN`, `zh_TW` or `ko_KR` profile gives each sheet
  without phonetic settings the locale's phonetic font, as Excel does.
- Threaded comments are saved with their legacy note in the profile's UI
  language, CRLF line breaks and replies by creation time, and `created` in UTC.
- A PivotTable cache saves whether its source was edited since the last refresh
  (`dataSourceClean`), and edited or new pivots are saved with their row and
  column items instead of asking Excel to rebuild them on open.
- `listPassthroughs()` and the dropped-part counts report only what a save
  loses, and `deferredFeatureCount` / `skippedFeatureCount` count every
  feature a save or load leaves out.
- Opening an `.xlsb` whose workbook part points outside the package fails with
  the zip-path status, and a protected range's security descriptor reads back
  in Excel's spelling.
- **Tables and styles.** A column inserted inside a table is named from its
  left neighbour, writing a totals-row cell updates the column's totals
  setting, and `removeTableStyle` re-points its users to their default style.
- `getCellStyle` and its siblings find a built-in style by its name in the
  profile language, and `updateCellFormat` merges a patch the same way in every
  binding.
- Under a Mac profile, column geometry and pagination use Mac Excel's widths for
  the Normal font; `win-*` profiles are unchanged.
- **Drawings and charts.** Chart type changes and new charts over dates follow
  Excel, chart and shape colours take theme references, a glued connector is
  re-routed when its shapes move, and default drawing ids count per sheet from 2.
- **Command line.** The tools apply `--profile` at load and set the workbook
  location from the input path.
- The WASM size ceiling is 6 MiB uncompressed and 1920 KiB Brotli, up from
  4.5 MiB and 1280 KiB.

### Removed

- **API.** `fm_cf_color_t`, the `*_argb` fields, `CfColor`, `fillRgb` /
  `lineRgb`, `EvalArrayResult`, `fm_workbook_evaluate_formula` and the other
  replaced C functions
  ([C ABI changes](docs/releases/upgrading-0.19.md#c-abi-changes)).
- The singleton `remove*` / `clear*` methods, `addHyperlink`,
  `setTableColumnTotals`, sheet protection's `enabled`, `DrawingObject.mediaPath`
  and the raw cell-count and cell-at exports of the JavaScript packages.
- `FilterOperator`, `ChartBlanksAs`, `SparklineEmptyCells`, Python
  `Value.to_python()` and `FormulonError.status`, and the `todaySerial`
  argument of `evaluateCfRange`.
- The CLI-only statuses 7005 to 7007, and the `--formulas`, `--values`,
  `--sheets` and `--metadata` flags.
- **Files.** Smart tags and a sheet view's `<pivotSelection>` are no longer kept
  through a save, as Excel's own save drops them.

### Performance

- Both WebAssembly builds are smaller: the link and the bundled third-party
  libraries are optimised for size, and helpers are shared.
- Text criteria (`COUNTIF`, `SUMIFS`), non-ASCII lookups, `TEXT` over many
  cells and cross-sheet reads of formula cells are faster.
- A range format costs memory in proportion to its rows, not its cells.

### Fixed

- **API.** In JavaScript an in-workbook hyperlink reads `target: null`, several
  flags no longer read true when false was stored, and Python reads `NOW()` and
  the other unpinned clocks in local time instead of UTC.
- **Formulas.** `OFFSET`, `INDIRECT`, `INDEX`, `XLOOKUP`, `TAKE`, `DROP`, `IF`,
  `CHOOSE`, `LET` and lambda calls return the reference they select, so
  `SUBTOTAL`, `COUNTIF`, `ISREF` and the like read it as the range itself.
- 3-D references give the results Excel gives under `:`, intersection, union,
  parentheses, `INDEX` and the arms of `IF`, `CHOOSE`, `IFS` and `SWITCH`.
- Lookup functions (`VLOOKUP`, `HLOOKUP`, `MATCH`, `LOOKUP`, `XLOOKUP`,
  `XMATCH`, `INDEX`, `CHOOSE`, `INDIRECT`) return what Excel returns for blank
  lookup values, wildcards, array indices, whole columns and sheet names.
- Conditional aggregates and database functions match error criteria, read
  their sum ranges and report field errors as Excel does.
- Financial functions (`AMORDEGRC`, the coupon and price family, `IRR`,
  `MIRR`, `XIRR`, `XNPV`, `NPV`, `YEARFRAC`, `SERIESSUM`) follow Excel's date
  rules, argument coercion and error order.
- Statistical and matrix functions (`PERCENTILE`, `QUARTILE`, `PERCENTRANK`,
  `FREQUENCY`, `LINEST` and its family, `MMULT`, `MODE`, `ERROR.TYPE`) coerce
  and check their arguments as Excel does.
- When several arguments fail, a call reports the error Excel reports, and
  date, text and lookup functions evaluate per element over arrays.
- `TEXT`, `TEXTBEFORE`, `TEXTAFTER`, `TEXTSPLIT`, omitted arguments of
  `SEQUENCE`, `TEXTJOIN` and the like, and Mac ja-JP text comparison of
  prolonged sound and voiced marks give Excel's results.
- Spills, implicit intersection and one-cell arrays read as Excel reads them,
  and `GROUPBY` / `PIVOTBY` give `#VALUE!` for an array-returning function.
- Structured references and references into other workbooks evaluate as the
  ranges they denote, and `LET` and `LAMBDA` names follow Excel's rules.
- Data-validation formulas read a range under an operator as its top-left
  cell, structural edits spell deleted references as Excel does, and deleting a
  sheet keeps the sheet-scoped names other sheets use.
- **Files.** Excel refused to open an `.xlsb` with formulas over a table or
  re-saved with PivotTables or custom views, a table column with a header
  format, and an `.xlsx` with a second window's pivot selection.
- Saved files opened in Excel with columns hidden, custom row heights scaled
  or the view scrolled down, and a theme edit was lost on Mac; columns, sheet
  views, `.xlsb` row records and the theme version are now written as Excel does.
- In-cell pictures and other rich values survive a save; they came back as
  `#VALUE!`.
- An `.xlsb` save stored a CSE array formula or a validation formula with a
  range under an operator so that Excel computed a different result, and stored
  the later criteria ranges of `MAXIFS` / `MINIFS` with the wrong class.
- `.xlsb` loads keep Fit to page, outline settings, page setup, margins and
  validation input modes, and saves to `.xlsx` keep the
  header, footer and page breaks, cell watches, ignored errors and scenarios.
- An `.xlsb` save writes a data bar's formula minimum or maximum, keeps a
  drawing holding form controls and survives renaming or deleting a sheet with
  extension rules; an `.xlsx` save of an `.xlsb` chartsheet is a valid package.
- Saves keep the stored codes of the built-in currency and accounting number
  formats, a colour-only dxf fill, and a font with no colour as automatic.
- Conditional formats and validations referring to another sheet load from
  `.xlsx` and `.xlsb`, and `.xlsb` keeps a qualified `#REF!`.
- Loads read an `.xlsx` date with fractional seconds as the exact serial, show
  a listed row on a sheet whose rows default to hidden, and let formulas read a
  macro sheet's values, whose shared strings an `.xlsx` save now keeps.
- Row and column edits move form controls and OLE objects with their legacy
  shapes, and a delete covering a move-and-size object removes it.
- **PivotTables.** Row and column grand totals were swapped when a PivotTable
  was read from or saved to `.xlsx`.
- A field with a numeric range grouping on itself shows Excel's ranges, a
  field sorted by a value field keeps that sort, and a page field showing one
  of several items reads "(Multiple Items)".
- Data fields keep their base field and item through a save, a cache read from
  `.xlsb` loads its dates as dates, and API-built caches carry the shared-item
  hints Excel needs, so it no longer asks to repair the file.
- Pivot layout for numeric label filters, column-only pivots and the classic
  layout without row fields matches Excel.
- **Conditional formatting and validation.** Text rules (`containsText`,
  `beginsWith` and the like) format the cells their stored formula selects, and
  a data bar without an extension is drawn as Excel 2007 draws it.
- AutoFilter custom criteria with wildcards, a literal-comparison flag or empty
  text filter as Excel does, an icon filter with no icon opens, and an
  ascending slicer sort order no longer fails.
- A new workbook under a Windows profile states its Normal font's default row
  height, and a threaded-comment mention after a line break lands where Excel
  puts it.
- **Drawings and charts.** A bare `autoTitleDeleted` reads as true, so editing a
  pie adds no title, and a one-space chart title is kept.
- **Command line.** `formulon paginate` exits with the library status when
  reading page breaks fails.

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

## Earlier releases

Releases up to 0.17.0 are in the [changelog archive](docs/releases/changelog-archive.md).

[Unreleased]: https://github.com/libraz/formulon/compare/v0.19.0...HEAD
[0.19.0]: https://github.com/libraz/formulon/compare/v0.18.0...v0.19.0
[0.18.0]: https://github.com/libraz/formulon/compare/v0.17.0...v0.18.0
