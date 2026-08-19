# Changelog

All notable changes to Formulon are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Added

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
  rather than on save. A style table is now seeded when a workbook is
  created, which keeps a caller-appended font, fill or border off the
  index slots Excel reserves.

### Fixed

- A manual page break written to an xlsx landed one row or column late
  when the file was opened in Excel, and a break read from an
  Excel-authored file paginated one track early. OOXML's `<brk id>` is
  already the zero-based index the break precedes, but the reader
  subtracted one from it and the writer added one back. The two errors
  cancelled inside a read/write cycle, so only pagination results and
  Excel disagreed with the model.

### Changed

- The WASM size report's soft ceilings moved to 2.75 MiB uncompressed and
  736 KiB Brotli. The previous pair sat below the shipped binary, so the
  warning fired on every run and said nothing about the change being
  measured. The new pair sits just above it, which is what a tripwire has
  to do to be read: one feature's worth of growth trips it, and the hard
  ceilings — unchanged at 3.00 MiB and 768 KiB — stay a further 0.25 MiB
  and 32 KiB away.

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

[Unreleased]: https://github.com/libraz/formulon/compare/v0.10.0...HEAD
[0.10.0]: https://github.com/libraz/formulon/compare/v0.9.7...v0.10.0
[0.9.7]: https://github.com/libraz/formulon/compare/v0.9.6...v0.9.7
[0.9.6]: https://github.com/libraz/formulon/compare/v0.9.5...v0.9.6
[0.9.5]: https://github.com/libraz/formulon/compare/v0.9.4...v0.9.5
[0.9.4]: https://github.com/libraz/formulon/compare/v0.9.3...v0.9.4
[0.9.3]: https://github.com/libraz/formulon/compare/v0.9.2...v0.9.3
[0.9.2]: https://github.com/libraz/formulon/compare/v0.9.1...v0.9.2
[0.9.1]: https://github.com/libraz/formulon/compare/v0.9.0...v0.9.1
[0.9.0]: https://github.com/libraz/formulon/releases/tag/v0.9.0
