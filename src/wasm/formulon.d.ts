//
// Hand-written TypeScript declarations for the Formulon WASM bindings.
//
// The single entry point is the default export from `formulon.js`,
// which is the Emscripten module factory produced under
// MODULARIZE=1 / EXPORT_NAME=createFormulon / EXPORT_ES6=1. It returns
// a Promise resolving to the Module surface declared below.
//
// Mirror of `EMSCRIPTEN_BINDINGS(formulon)` in `src/wasm/embind.cpp`.
// Keep this file in sync when adding or removing bindings.

/** `fm_value_kind_t` ordinals (mirror of `fm_value_kind_t`). */
export enum ValueKind {
  Blank = 0,
  Number = 1,
  Bool = 2,
  Text = 3,
  Error = 4,
  Array = 5,
  Ref = 6,
  Lambda = 7,
}

/** Result envelope returned by every fallible binding call. */
export interface Status {
  /** True when the underlying C ABI returned `kOk`. */
  ok: boolean;
  /** Numeric `fm_status_t`. 0 on success. */
  status: number;
  /** Describes the failure this call hit; empty on success. Normally the
   *  thread-local last-error message the C ABI recorded, but a rejection
   *  the binding raised itself -- a destroyed handle, a missing or
   *  wrongly-shaped argument -- carries its own message instead, never the
   *  residue of an earlier call. */
  message: string;
  /** Optional thread-local context string. Empty on success, and empty for
   *  a failure the binding raised before reaching the C ABI. */
  context: string;
}

/** Per-call telemetry returned by `Workbook.recalcParallel`. The five
 * scheduler counters are `uint64_t` in the C ABI but are exposed as
 * JavaScript `number` values by the WASM binding. They are exact through
 * `Number.MAX_SAFE_INTEGER` (2^53 - 1). */
export interface ParallelRecalcStats {
  cellsEvaluated: number;
  sccsProcessed: number;
  parallelSteps: number;
  serialFallbackSteps: number;
  cycleRecoveries: number;
  workerThreadsStarted: number;
  workerThreadsUsed: number;
}

/** `{ status, stats }` returned by `Workbook.recalcParallel`. */
export interface ParallelRecalcResult {
  status: Status;
  stats: ParallelRecalcStats;
}

/** A JS array with the status of the list enumeration that produced it. */
export type ListResult<T> = ReadonlyArray<T> & { readonly status: Status };

/** Flattened mirror of `fm_value_t`. Only the field selected by `kind`
 *  is meaningful; the others carry default-zero values. */
export interface Value {
  kind: ValueKind;
  /** Active when `kind === ValueKind.Number`. */
  number: number;
  /** Active when `kind === ValueKind.Bool` (0 or 1). */
  boolean: number;
  /** Active when `kind === ValueKind.Text`. */
  text: string;
  /** Active when `kind === ValueKind.Error`; a `formulon::ErrorCode` ordinal. */
  errorCode: number;
}

/** `{ status, value }` pair returned by cell-read entry points. */
export interface CellResult {
  status: Status;
  value: Value;
}

/** Return type of `evalFormula(...)`. */
export interface EvalResult {
  status: Status;
  value: Value;
}

/**
 * Return type of `Workbook.evaluateFormulaArray(...)`: the whole multi-cell
 * result of an ad-hoc formula. `cells` is a `rows` x `cols` nested array in
 * row-major order (`cells[r][c]`); a scalar result is reported as a 1x1
 * array (`rows === cols === 1`).
 */
export interface EvalArrayResult {
  status: Status;
  rows: number;
  cols: number;
  cells: Value[][];
}

/** Return type of `Workbook.save()` / `Workbook.saveAs(format)`. */
export interface SaveResult {
  status: Status;
  /** Freshly-allocated `Uint8Array` on success; `null` on failure. */
  bytes: Uint8Array | null;
}

/**
 * Return type of `Workbook.saveWithDiagnostics(format)`.
 *
 * A field means the same thing whichever container was written; where the
 * underlying event exists in only one container, the other simply never
 * raises it. Coverage is partial by design — an all-zero result means
 * none of the losses below occurred, not that nothing was logged.
 *
 * Summing these does NOT give "how many things did I lose":
 * `droppedPartCount` and `droppedRelationshipCount` both rise for one
 * dropped part.
 */
export interface SaveDiagnosticsResult extends SaveResult {
  /** Formula cells emitted as cached literals. Always zero for XLSX. */
  downgradedFormulaCount: number;
  /** Sheet features not lowered to records. Always zero for XLSX. */
  deferredFeatureCount: number;
  /** Passthrough parts dropped for colliding with a generated path, or
   *  for two entries claiming one path. Both containers. */
  droppedPartCount: number;
  /** Relationships dropped because their target part is gone. Both
   *  containers, at package / workbook / (XLSB) sheet scope. Rises with
   *  `droppedPartCount` for one dropped part — do not add them. */
  droppedRelationshipCount: number;
  /** Tables emitted under a writer-assigned id. Always zero for XLSB. */
  renumberedPartCount: number;
}

/**
 * Return type of `Workbook.readDiagnostics()`. Same contract as
 * `SaveDiagnosticsResult`: one meaning per field across both containers,
 * partial coverage by design.
 */
export interface ReadDiagnosticsResult {
  status: Status;
  /** Formula cells whose stored formula could not be decoded. XLSB only. */
  undecodedFormulaCount: number;
  /** Defined names skipped for the same reason. XLSB only. */
  undecodedDefinedNameCount: number;
  /** Package parts whose content type could not be resolved, so they were
   *  not decoded. XLSB only — the XLSX reader keeps unmodelled parts as
   *  passthrough instead of dropping them. */
  undecodedPartCount: number;
  /** Presentation-overlay entries dropped for an unusable reference: one
   *  per merge / hyperlink / data validation, one per conditional-format
   *  block (a block can carry several rules). XLSX only. */
  skippedFeatureCount: number;
  /** Workbook parts whose content type was unrecognised. Max 1. XLSX only. */
  unknownContentTypeCount: number;
}

/** `fm_workbook_format_t` ordinals: container format for `saveAs`. */
export enum WorkbookFormat {
  Unknown = 0,
  Xlsx = 1,
  Xlsb = 2,
}

/** Return type of `Workbook.sheetName(idx)`. */
export interface StringResult {
  status: Status;
  value: string;
}

/** Return type of `Workbook.cellAt(sheet, idx)`. */
export interface CellEntry {
  status: Status;
  row?: number;
  col?: number;
  /** Raw formula text, or `null` for pure literals. */
  formula?: string | null;
  value?: Value;
}

/** Return type of `Workbook.definedNameAt(idx)`. */
export interface DefinedNameEntry {
  status: Status;
  name?: string;
  formula?: string;
  /** -1 for workbook scope; otherwise a 0-based sheet index. */
  localSheetId?: number;
}

/** Return type of `Workbook.tableAt(idx)`. */
export interface TableEntry {
  status: Status;
  name?: string;
  displayName?: string;
  ref?: string;
  sheetIndex?: number;
}

/** Authoring input for a worksheet table. `columns` must have exactly one
 * name per column in `ref`; `styleName` follows Excel's built-in TableStyle
 * names. With `headerRow` (the default), the caller still writes the header
 * cells: Excel expects the first row of `ref` to hold the column names. */
export interface TableInput {
  sheetIndex: number;
  ref: string;
  name: string;
  displayName?: string;
  columns: string[];
  styleName?: string;
  headerRow?: boolean;
  totalsRow?: boolean;
}

/** Partial table metadata update. Omitted fields preserve their existing
 * values; an empty `styleName` explicitly removes the table style. When
 * omitted, `ref` also preserves the current range. `headerRow` and
 * `totalsRow` update their flags only when supplied. */
export interface TableUpdateInput {
  ref?: string;
  styleName?: string;
  headerRow?: boolean;
  totalsRow?: boolean;
}

/** Return type of `Workbook.passthroughAt(idx)`. */
export interface PassthroughEntry {
  status: Status;
  path?: string;
}

/** PivotTable layout cell kind. Mirrors `fm_pivot_cell_kind_t`. */
export enum PivotCellKind {
  Header = 0,
  RowLabel = 1,
  ColLabel = 2,
  Data = 3,
  RowSubtotal = 4,
  ColSubtotal = 5,
  GrandTotal = 6,
  Blank = 7,
}

/** One concrete cell in a projected PivotTable layout. */
export interface PivotCell {
  /** Absolute 0-based sheet row. */
  row: number;
  /** Absolute 0-based sheet column. */
  col: number;
  value: Value;
  kind: PivotCellKind;
  /** Header nesting depth; 0 for data cells. */
  depth: number;
  /** Source field name when known. */
  fieldName: string;
  /** Excel number-format code when known. */
  numberFormat: string;
}

/** Return type of `Workbook.pivotLayout(sheet, pivotIndex)`. */
export interface PivotLayoutResult {
  status: Status;
  /** Absolute 0-based top row of the rectangular pivot layout. */
  top: number;
  /** Absolute 0-based left column of the rectangular pivot layout. */
  left: number;
  /** Rectangular layout row span. */
  rows: number;
  /** Rectangular layout column span. */
  cols: number;
  /** Sparse projected cells in row-major order. */
  cells: ReadonlyArray<PivotCell>;
}

/** PivotTable axis enumeration. Mirrors `fm_pivot_axis_t`. */
export enum PivotAxis {
  Row = 0,
  Col = 1,
  Value = 2,
  Page = 3,
}

/** Aggregation function for a value-axis field. Mirrors
 *  `fm_pivot_aggregation_t`. */
export enum PivotAggregation {
  Sum = 0,
  Count = 1,
  Average = 2,
  Max = 3,
  Min = 4,
  Product = 5,
  CountNumbers = 6,
  StdDev = 7,
  StdDevP = 8,
  Var = 9,
  VarP = 10,
}

/** Show-values-as derivation applied to a data-field aggregate.
 *  Mirrors `fm_pivot_show_as_t`. */
export enum PivotShowValuesAs {
  Normal = 0,
  PercentOfRow = 1,
  PercentOfCol = 2,
  PercentOfTotal = 3,
  RunningTotalInRow = 4,
  RunningTotalInCol = 5,
  Index = 6,
  DifferenceFrom = 7,
  PercentDifferenceFrom = 8,
  PercentOfParentRow = 9,
  PercentOfParentCol = 10,
  PercentOfParent = 11,
}

/** Sentinel value for `PivotDataFieldSpec.showAsBaseItem` meaning
 *  "(previous)". Mirrors the magic value used by Excel's OOXML
 *  `<dataField>` element. */
export const PIVOT_SHOW_AS_BASE_PREVIOUS = 1048828;
/** Sentinel value for `PivotDataFieldSpec.showAsBaseItem` meaning
 *  "(next)". */
export const PIVOT_SHOW_AS_BASE_NEXT = 1048829;

/** Filter type for an active (slicer-applied) filter. Mirrors
 *  `fm_pivot_filter_type_t`. */
export enum PivotFilterType {
  ValueTop10 = 0,
  ValueGreaterThan = 1,
  ValueBetween = 2,
  LabelContains = 3,
  LabelBeginsWith = 4,
  LabelDate = 5,
}

/** Date-grouping granularity. Mirrors `fm_pivot_date_grouping_t`. */
export enum PivotDateGrouping {
  Day = 0,
  Month = 1,
  Quarter = 2,
  Year = 3,
  Week = 4,
  Hour = 5,
  Minute = 6,
  Second = 7,
}

/** Calendar system used by date grouping. Mirrors
 *  `fm_pivot_calendar_t`. */
export enum PivotCalendar {
  Gregorian = 0,
  Japanese = 1,
}

/** Pivot report layout form. Mirrors `fm_pivot_layout_t`. */
export enum PivotReportLayout {
  Compact = 0,
  Tabular = 1,
  Outline = 2,
}

/** Worksheet source metadata for a pivot cache. */
export interface PivotWorksheetSource {
  present: boolean;
  ref?: string;
  sheet?: string;
  name?: string;
}

export interface PivotWorksheetSourceResult extends PivotWorksheetSource {
  status: Status;
  ref: string;
  sheet: string;
  name: string;
}

export interface PivotReportLayoutResult {
  status: Status;
  layout: PivotReportLayout;
}

/** Discriminator for the variant payload carried by a pivot filter
 *  spec. `None` (= -1) means the slot is unset (only meaningful for
 *  the optional upper-bound payload on range filters). Mirrors
 *  `fm_pivot_filter_value_kind_t`. */
export enum PivotFilterValueKind {
  None = -1,
  Int = 0,
  Double = 1,
  Text = 2,
}

/** Plain-data spec for `Workbook.pivotFieldAdd`. Mirrors
 *  `fm_pivot_field_spec_t`. */
export interface PivotFieldSpec {
  sourceName: string;
  customName?: string;
  axis: PivotAxis;
  subtotalTop?: boolean;
  numberFormat?: string;
}

/** Plain-data spec for `Workbook.pivotDataFieldAdd` /
 *  `pivotDataFieldSet`. Mirrors `fm_pivot_data_field_spec_t`. Pass
 *  `-1` for `showAsBaseField` / `showAsBaseItem` to mean "unset". */
export interface PivotDataFieldSpec {
  name: string;
  fieldIndex: number;
  aggregation: PivotAggregation;
  numberFormat?: string;
  showAs?: PivotShowValuesAs;
  showAsBaseField?: number;
  showAsBaseItem?: number;
}

/** Plain-data spec for `Workbook.pivotFilterAdd`. Mirrors
 *  `fm_pivot_filter_spec_t`. For non-range filters, `valueKind` must be
 *  one of `Int` / `Double` / `Text`; the high-bound payload is honoured
 *  only for range filter types (`ValueBetween`, `LabelDate`). */
export interface PivotFilterSpec {
  axis: PivotAxis;
  fieldName: string;
  type: PivotFilterType;
  /** Data-field aggregate scored by value filters. Defaults to 0. */
  dataFieldIndex?: number;
  valueKind?: PivotFilterValueKind;
  valueInt?: number;
  valueDouble?: number;
  valueText?: string;
  valueHighKind?: PivotFilterValueKind;
  valueHighInt?: number;
  valueHighDouble?: number;
}

/** Return type of `Workbook.pivotFilterAt(...)`: the stored filter plus a
 *  status envelope. Field names and semantics match `PivotFilterSpec`.
 *
 *  The active-filter list is **session state**. An entry added through
 *  `pivotFilterAdd` affects evaluation only while the workbook handle
 *  lives and is not written by `save`. Conversely a `<filters>` block
 *  Excel wrote is preserved verbatim on save but does not appear here, so
 *  `pivotFilterCount` reports only what this session added. The filter
 *  surface that does persist is pivot field item visibility. */
export interface PivotFilterResult {
  status: Status;
  axis: number;
  fieldName: string;
  type: number;
  dataFieldIndex: number;
  valueKind: number;
  valueInt: number;
  valueDouble: number;
  valueText: string;
  valueHighKind: number;
  valueHighInt: number;
  valueHighDouble: number;
}

/** Conditional-format match kind. Mirrors `formulon::cf::CFMatchKind`. */
export enum CfMatchKind {
  DifferentialFormat = 0,
  ColorScale = 1,
  DataBar = 2,
  IconSet = 3,
}

/** Excel cell-error ordinals. Mirrors `formulon::ErrorCode`; these are the
 *  values carried by `Value.errorCode` and accepted by `setError`. */
export enum ErrorCode {
  Null = 0,
  Div0 = 1,
  Value = 2,
  Ref = 3,
  Name = 4,
  Num = 5,
  NA = 6,
  GettingData = 7,
  Spill = 8,
  Calc = 9,
  Field = 10,
  Blocked = 11,
  Connect = 12,
  External = 13,
  Busy = 14,
  Python = 15,
  Unknown = 16,
}

/**
 * Workbook-level calc mode (`<calcPr calcMode>`).
 *
 * - `Auto` — recalc on every input change (Excel default).
 * - `Manual` — only recalc when explicitly requested.
 * - `AutoNoTable` — recalc on every change EXCEPT data-table cells.
 *
 * Round-trip metadata only — the engine recalcs all dirty cells
 * regardless of which mode is set.
 */
export enum CalcMode {
  Auto = 0,
  Manual = 1,
  AutoNoTable = 2,
}

export type ExcelProfileId = 'mac-365-ja_JP' | 'win-365-ja_JP';

/**
 * A wall-clock reading in local civil fields, as read back from
 * {@link Workbook.pinnedNow}. `month` is 1-12 and `day` is 1-31; the other
 * fields follow the usual 24-hour ranges.
 */
export interface CivilTime {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
}

/** RGBA colour. Channels are 0-255 (sRGB). */
export interface CfColor {
  r: number;
  g: number;
  b: number;
  a: number;
}

/** Conditional-format value object.
 *  `type`: 0 number, 1 percent, 2 percentile, 3 min, 4 max,
 *  5 formula, 6 autoMin, 7 autoMax. */
export interface CfValueObjectInput {
  type: number;
  value?: string;
  gte?: boolean;
}

/** Resolved CF match. Active fields depend on `kind`; the others carry
 *  default-zero values. */
export interface CfMatch {
  kind: CfMatchKind;
  priority: number;
  /** `1` when `dxfId` is meaningful; `0` otherwise. */
  dxfIdEngaged: number;
  dxfId: number;
  /** Active when `kind === ColorScale`. */
  color: CfColor;
  /** Active when `kind === DataBar`. */
  barLengthPct: number;
  barAxisPositionPct: number;
  barIsNegative: number;
  barFill: CfColor;
  barBorderEngaged: number;
  barBorder: CfColor;
  barGradient: number;
  /** Active when `kind === IconSet`; ordinal of `formulon::cf::IconSetName`. */
  iconSetName: number;
  iconIndex: number;
}

/** One cell's CF result inside a viewport-range evaluation. */
export interface CfCellResult {
  row: number;
  col: number;
  matches: ReadonlyArray<CfMatch>;
}

/** Return type of `Workbook.evaluateCfRange(...)`. `cells` is sparse:
 *  only cells that produced at least one match appear. */
export interface CfRangeResult {
  status: Status;
  cells: ReadonlyArray<CfCellResult>;
}

/** Resolved print geometry for one worksheet. All coordinates are 0-based. */
export interface PaginationResult {
  status: Status;
  /** The sheet's declared `_xlnm.Print_Area`. Empty when the sheet declares
   *  none -- it is not backfilled with the used range. Pagination itself
   *  still falls back to the used range in that case, so `pageCount` can be
   *  non-zero while this is empty. */
  printArea: Array<{ firstRow: number; firstCol: number; lastRow: number; lastCol: number }>;
  horizontalBreaks: number[];
  verticalBreaks: number[];
  pageCount: number;
}

/** Sheet tab visibility, mirroring OOXML `<sheet state>`.
 *
 *  `VeryHidden` differs from `Hidden` in that Excel leaves such a sheet
 *  out of its "Unhide" dialog, which is how a workbook keeps a settings
 *  or lookup sheet out of a user's reach. */
export enum SheetVisibility {
  Visible = 0,
  Hidden = 1,
  VeryHidden = 2,
}

/** Per-sheet view: zoom (10..400, default 100), frozen-pane row/col
 *  counts, tab visibility, and the display / orientation flags mirrored
 *  from OOXML `<sheetView>` (gridlines, row/col headers, zero display,
 *  right-to-left, tab-selected, view mode). Booleans are encoded as
 *  `0`/`1` to match the embind binding's wire shape. */
export interface SheetView {
  zoomScale: number;
  freezeRows: number;
  freezeCols: number;
  /** Boolean stored as 0/1 to match the embind binding's wire shape.
   *  The two-state view of `visibility`: `1` for both hidden states, so
   *  reading only this field sees a very-hidden sheet as hidden rather
   *  than visible. Read `visibility` to tell the two apart. */
  tabHidden: number;
  /** The authoritative tab state, as a `SheetVisibility` code. Always
   *  agrees with `tabHidden`. */
  visibility: number;
  /** `showGridLines`; default 1. */
  showGridLines: number;
  /** `showRowColHeaders`; default 1. */
  showRowColHeaders: number;
  /** `showZeros`; default 1. */
  showZeros: number;
  /** `rightToLeft`; default 0. */
  rightToLeft: number;
  /** `tabSelected`; default 0. */
  tabSelected: number;
  /** `view` mode: `""` (normal), `"pageBreakPreview"`, or `"pageLayout"`. */
  viewMode: string;
}

/** Return type of `Workbook.getSheetView(sheet)`. */
export interface SheetViewResult {
  status: Status;
  view: SheetView;
}

/**
 * Mirror of OOXML `<sheetProtection>` (ECMA-376 §18.3.1.85).
 *
 * Round-trip metadata only — the engine does not enforce locks at
 * evaluation time. The host UI inspects these flags to mirror Excel's
 * "Protect Sheet" dialog state. Booleans are encoded as `0`/`1` to
 * match the embind wire shape.
 *
 * `enabled` controls whether the `<sheetProtection>` element is
 * emitted at all; setting it to `0` clears the protection block on
 * save.
 */
export interface SheetProtection {
  enabled: number;
  algorithmName: string;
  hashValue: string;
  saltValue: string;
  spinCount: number;
  legacyPassword: string;
  sheet: number;
  objects: number;
  scenarios: number;
  formatCells: number;
  formatColumns: number;
  formatRows: number;
  insertColumns: number;
  insertRows: number;
  insertHyperlinks: number;
  deleteColumns: number;
  deleteRows: number;
  selectLockedCells: number;
  selectUnlockedCells: number;
  sort: number;
  autoFilter: number;
  pivotTables: number;
}

/** Return type of `Workbook.getSheetProtection(sheet)`. */
export interface SheetProtectionResult {
  status: Status;
  protection: SheetProtection;
}

/** Per-column-range layout override. Inclusive `[first, last]` columns
 *  carry the same width / hidden / outline level. */
export interface ColumnLayout {
  first: number;
  last: number;
  width: number;
  /** Boolean stored as 0/1. */
  hidden: number;
  outlineLevel: number;
  /** Whether the width is logically explicit (raw presence or non-zero legacy width). */
  hasWidth: number;
  /** Whether the source / model carries an explicit style attribute. */
  hasStyle: number;
  styleXf: number;
}

/** Return type of `Workbook.getSheetColumns(sheet)`. */
export interface ColumnsResult {
  status: Status;
  columns: ReadonlyArray<ColumnLayout>;
}

/** Per-row layout override. */
export interface RowLayout {
  row: number;
  height: number;
  /** Boolean stored as 0/1. */
  hidden: number;
  outlineLevel: number;
  /** Whether OOXML customFormat=1 makes the row style effective. */
  hasStyle: number;
  styleXf: number;
}

/** Return type of `Workbook.getSheetRowOverrides(sheet)`. */
export interface RowsResult {
  status: Status;
  rows: ReadonlyArray<RowLayout>;
}

/** Inclusive cell rectangle used by `addMerge` / `getMerges`. */
export interface MergeRange {
  firstRow: number;
  lastRow: number;
  firstCol: number;
  lastCol: number;
}

/** Sheet hyperlink entry as returned by `getHyperlinks(sheet)`. */
export interface HyperlinkEntry {
  row: number;
  col: number;
  /** Inclusive bottom-right endpoint of the hyperlink rectangle. */
  lastRow: number;
  lastCol: number;
  /** Absolute or relative target (URL, email, internal ref, …). */
  target: string;
  /** In-workbook destination (empty for an external target). */
  location: string;
  /** Display text override (empty when default). */
  display: string;
  /** Tooltip text (empty when none). */
  tooltip: string;
}

/** Cell comment entry returned by `getComment(sheet, row, col)`. */
export interface CommentEntry {
  author: string;
  text: string;
}

/** Status-bearing comment lookup returned by `getCommentResult`. */
export interface CommentResult {
  status: Status;
  comment: CommentEntry | null;
}

/** Sheet-wide comment entry returned by `getComments(sheet)`. Extends
 *  `CommentEntry` with the anchor cell so comments on otherwise-empty
 *  cells can be discovered without already knowing their `(row, col)`. */
export interface SheetCommentEntry extends CommentEntry {
  row: number;
  col: number;
}

/** One cell-range entry inside a `DataValidationEntry.ranges`. Identical
 *  shape to `MergeRange`; declared separately so the data-validation
 *  surface can evolve independently of the merge surface. */
export interface DataValidationRange {
  readonly firstRow: number;
  readonly firstCol: number;
  readonly lastRow: number;
  readonly lastCol: number;
}

/** One sheet `<dataValidation>` block as returned by
 *  `getValidations(sheet)`.
 *
 *  Field semantics (matches OOXML `dataValidations.xsd`):
 *    * `type`        — 0 none, 1 whole, 2 decimal, 3 list, 4 date,
 *                       5 time, 6 textLength, 7 custom.
 *    * `op`          — 0 between, 1 notBetween, 2 equal, 3 notEqual,
 *                       4 greaterThan, 5 lessThan,
 *                       6 greaterThanOrEqual, 7 lessThanOrEqual.
 *    * `errorStyle`  — 0 stop, 1 warning, 2 information.
 */
export interface DataValidationEntry {
  readonly ranges: ReadonlyArray<DataValidationRange>;
  readonly type: number;
  readonly op: number;
  readonly errorStyle: number;
  readonly allowBlank: boolean;
  readonly showInputMessage: boolean;
  readonly showErrorMessage: boolean;
  /** Whether the in-cell dropdown arrow is shown (`list` type only).
   *  Defaults to `true`, matching Excel's default. */
  readonly showDropDown: boolean;
  readonly formula1: string;
  readonly formula2: string;
  readonly errorTitle: string;
  readonly errorMessage: string;
  readonly promptTitle: string;
  readonly promptMessage: string;
}

/** Argument shape accepted by `addValidation(sheet, validation)`. Every
 *  field except `type` is optional; missing fields default to `0` for
 *  the small enum-shaped integers, `false` for booleans, and `""` for
 *  strings. `ranges` is required by the model but accepted as missing
 *  to allow zero-range rules. */
export interface DataValidationInput {
  ranges?: ReadonlyArray<DataValidationRange>;
  type: number;
  op?: number;
  errorStyle?: number;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  /** Whether the in-cell dropdown arrow is shown (`list` type only).
   *  Defaults to `true`, matching Excel's default. */
  showDropDown?: boolean;
  formula1?: string;
  formula2?: string;
  errorTitle?: string;
  errorMessage?: string;
  promptTitle?: string;
  promptMessage?: string;
}

/** @deprecated Use {@link DataValidationEntry} instead. */
export type ValidationEntry = DataValidationEntry;

/** Workbook-wide cell coordinate used by trace results. `sheet` is the
 *  0-based sheet index. */
export interface CellNode {
  readonly sheet: number;
  readonly row: number;
  readonly col: number;
}

/** Result envelope for `functionMetadata(name, locale)`.
 *
 *  `ok` is `false` when no function matches `name`; the remaining
 *  fields are absent. When `ok` is `true`, `name` / `minArity` /
 *  `maxArity` are always populated; `signatureTemplate` and
 *  `description` are always `undefined` — the engine ships no
 *  human-readable function text and a host merges its own provider
 *  document over this result (`docs/function-metadata-schema.md`). */
export interface FunctionMetadataResult {
  readonly ok: boolean;
  readonly name?: string;
  readonly minArity?: number;
  /** `null` denotes an unbounded variadic or a lazy / special form whose
   *  upper arity is unknown. */
  readonly maxArity?: number | null;
  /** Availability class: 0 implemented, 1 implemented_unverified,
   *  2 environment_bound, 3 unavailable_stub. */
  readonly availability?: number;
  readonly signatureTemplate?: string;
  readonly description?: string;
}

/** Per-locale display overrides inside a {@link FunctionMetadataEntry}. */
export interface FunctionMetadataLocalized {
  signature?: string;
  description?: string;
}

/** One host-injected metadata entry, keyed by canonical UPPERCASE function
 *  name inside a {@link FunctionMetadataProvider}. Display-only metadata
 *  (see `docs/function-metadata-schema.md`); it never affects formula
 *  parsing or evaluation. */
export interface FunctionMetadataEntry {
  /** Default (locale-agnostic) signature template. */
  signature?: string;
  /** Default (locale-agnostic) description. */
  description?: string;
  /** Map of BCP-47 locale tag -> localized display name. */
  aliases?: Record<string, string>;
  /** Map of BCP-47 locale tag -> per-locale signature/description overrides. */
  localized?: Record<string, FunctionMetadataLocalized>;
}

/** A whole host metadata document's `functions` map: canonical UPPERCASE
 *  function name -> {@link FunctionMetadataEntry}. */
export type FunctionMetadataProvider = Record<string, FunctionMetadataEntry>;

/** Result of a {@link MergeFunctionMetadata} call: a
 *  {@link FunctionMetadataResult} with the resolved localized display name
 *  attached. */
export interface MergedFunctionMetadataResult extends FunctionMetadataResult {
  /** `entry.aliases[locale]` when present, else the canonical `name`. */
  readonly localizedName?: string;
}

/** Merges a host-supplied {@link FunctionMetadataEntry} over the engine's
 *  structural `functionMetadata()` result.
 *
 *  Field precedence (first non-nullish wins):
 *    - `signatureTemplate`: `entry.localized[locale].signature` ->
 *      `entry.signature` -> `base.signatureTemplate`
 *    - `description`: `entry.localized[locale].description` ->
 *      `entry.description` -> `base.description`
 *    - `localizedName`: `entry.aliases[locale]` -> `base.name`
 *
 *  When `entry` is `undefined`, `base` is returned verbatim (signature /
 *  description stay `undefined`). `locale` is a BCP-47 display tag matching
 *  the keys in `aliases` / `localized`, independent of the numeric locale
 *  passed to `functionMetadata()`. See `docs/function-metadata-schema.md`.
 *
 *  A pure host-side helper: it touches no WASM state and is exported from
 *  this package's hand-written entry point rather than the generated module
 *  glue. `@libraz/formulon-native` exports the same function. */
export function mergeFunctionMetadata(
  base: FunctionMetadataResult,
  entry: FunctionMetadataEntry | undefined,
  locale: string,
): MergedFunctionMetadataResult;

/** Signature of {@link mergeFunctionMetadata}. */
export type MergeFunctionMetadata = typeof mergeFunctionMetadata;

/** Spill region info returned by `spillInfo(sheet, row, col)`. */
export interface SpillInfo {
  readonly engaged: boolean;
  readonly anchorRow: number;
  readonly anchorCol: number;
  readonly rows: number;
  readonly cols: number;
}

/** One inclusive cell-range entry inside a CF rule's `sqref` union. */
export interface ConditionalFormatRange {
  readonly firstRow: number;
  readonly firstCol: number;
  readonly lastRow: number;
  readonly lastCol: number;
}

/** One CF rule as returned by `getConditionalFormats(sheet)`.
 *
 *  `type` ordinal mirrors `formulon::cf::RuleType`:
 *    0 expression, 1 cellIs, 2 colorScale, 3 dataBar, 4 iconSet,
 *    5 top10, 6 aboveAverage, 7 containsText, 8 notContainsText,
 *    9 beginsWith, 10 endsWith, 11 containsBlanks, 12 notContainsBlanks,
 *    13 containsErrors, 14 notContainsErrors, 15 timePeriod,
 *    16 duplicateValues, 17 uniqueValues.
 *
 *  `id`, `type`, `priority`, `stopIfTrue` and `sqref` populate for every
 *  rule kind; the remaining fields are engaged by the kinds noted on
 *  each. A visual rule (`colorScale` / `dataBar` / `iconSet`) carries its
 *  full sub-spec, so an entry read back here can be handed straight to
 *  `addConditionalFormat`.
 */
export interface ConditionalFormatEntry {
  readonly id: string;
  readonly type: number;
  readonly priority: number;
  readonly stopIfTrue: boolean;
  readonly sqref: ReadonlyArray<ConditionalFormatRange>;
  readonly dxfId?: number;
  readonly formula1?: string;
  readonly formula2?: string;
  /** `formulon::cf::CellIsOperator` ordinal: 0 lt, 1 le, 2 eq, 3 ne,
   *   4 ge, 5 gt, 6 between, 7 notBetween. Engaged for `cellIs` rules. */
  readonly op?: number;
  /** Engaged for `top10` rules. */
  readonly rank?: number;
  readonly percent?: boolean;
  readonly bottom?: boolean;
  /** Engaged for `aboveAverage` rules. */
  readonly aboveAverage?: boolean;
  readonly equalAverage?: boolean;
  readonly stdDev?: number;
  /** Engaged for `containsText` / `beginsWith` / `endsWith` /
   *  `notContainsText` rules. */
  readonly text?: string;
  /** `formulon::cf::TimePeriod` ordinal. Engaged for `timePeriod` rules. */
  readonly timePeriod?: number;
  /** Engaged for `colorScale` rules. */
  readonly colorScale?: {
    readonly thresholds: ReadonlyArray<CfValueObjectInput>;
    readonly colors: ReadonlyArray<CfColor>;
  };
  /** Engaged for `dataBar` rules. */
  readonly dataBar?: {
    readonly min: CfValueObjectInput;
    readonly max: CfValueObjectInput;
    readonly fill: CfColor;
    readonly showValue: boolean;
    readonly minLengthPct: number;
    readonly maxLengthPct: number;
    /**
     * `x14` extension payload. Present whenever the rule states the
     * setting; a rule read back from a workbook always carries all six, so
     * the object can be handed straight to `addConditionalFormat`.
     *
     * `axisPosition` is 0 = automatic, 1 = middle, 2 = none.
     */
    readonly gradient?: boolean;
    readonly axisPosition?: number;
    readonly negativeFill?: CfColor;
    readonly border?: CfColor;
    readonly negativeBorder?: CfColor;
    readonly axisColor?: CfColor;
  };
  /** Engaged for `iconSet` rules. */
  readonly iconSet?: {
    readonly name: number;
    readonly thresholds: ReadonlyArray<CfValueObjectInput>;
    readonly reverse: boolean;
    readonly showValue: boolean;
    /** Round-trip only: the `percent` attribute is preserved across load
     *  and save but never consulted during evaluation. Each threshold
     *  carries its own `type`, and that type is what interprets it. */
    readonly percent: boolean;
  };
}

/** Argument shape accepted by `addConditionalFormat(sheet, rule)`.
 *
 *  When `priority` is missing, zero, or negative, the engine assigns
 *  `existing_max + 1`. When `id` is missing or empty, the engine
 *  synthesises one. */
export interface ConditionalFormatInput {
  sqref: ReadonlyArray<ConditionalFormatRange>;
  type: number;
  priority?: number;
  stopIfTrue?: boolean;
  id?: string;
  dxfId?: number;
  formula1?: string;
  formula2?: string;
  op?: number;
  rank?: number;
  percent?: boolean;
  bottom?: boolean;
  aboveAverage?: boolean;
  equalAverage?: boolean;
  stdDev?: number;
  text?: string;
  timePeriod?: number;
  /** Payload for type 2 (`colorScale`). Threshold and color counts must match: 2 or 3. */
  colorScale?: {
    thresholds: ReadonlyArray<CfValueObjectInput>;
    colors: ReadonlyArray<CfColor>;
  };
  /** Payload for type 3 (`dataBar`). */
  dataBar?: {
    min: CfValueObjectInput;
    max: CfValueObjectInput;
    fill: CfColor;
    showValue?: boolean;
    minLengthPct?: number;
    maxLengthPct?: number;
    /**
     * `x14` extension payload. Omit a key to keep the model default:
     * gradient fill on, automatic axis, negative fill equal to `fill`, no
     * border, and a black axis. `axisPosition` is 0 = automatic,
     * 1 = middle, 2 = none.
     */
    gradient?: boolean;
    axisPosition?: number;
    negativeFill?: CfColor;
    border?: CfColor;
    negativeBorder?: CfColor;
    axisColor?: CfColor;
  };
  /** Payload for type 4 (`iconSet`). `name` is `formulon::cf::IconSetName` ordinal. */
  iconSet?: {
    name: number;
    thresholds: ReadonlyArray<CfValueObjectInput>;
    reverse?: boolean;
    showValue?: boolean;
    /** Round-trip only: preserved across load and save but never
     *  consulted during evaluation (each threshold's own `type` is
     *  authoritative). */
    percent?: boolean;
  };
}

/** Return type of `Workbook.getCellXfIndex(sheet, row, col)`. */
export interface CellXfIndexResult {
  status: Status;
  xfIndex: number;
}

/** Return type of `Workbook.getCellXf(xfIndex)`. */
export interface CellXfResult {
  status: Status;
  fontIndex: number;
  fillIndex: number;
  borderIndex: number;
  numFmtId: number;
  horizontalAlign: number;
  verticalAlign: number;
  wrapText: boolean;
  /** OOXML `justifyLastLine`; effective with distributed horizontal alignment. */
  justifyLastLine: boolean;
  /** Whether the source `<xf>` carried an `<alignment>` child. */
  hasAlignment: boolean;
  /** Whether `horizontal="..."` was explicitly present. */
  hasHorizontalAlign: boolean;
  /** Whether `vertical="..."` was explicitly present. */
  hasVerticalAlign: boolean;
  /** Whether `wrapText="..."` was explicitly present. */
  hasWrapText: boolean;
  /** Whether `justifyLastLine="..."` was explicitly present. */
  hasJustifyLastLine: boolean;
  /** OOXML `textRotation`; present only when the source attribute existed. */
  textRotation?: number;
  /** OOXML `indent`; present only when the source attribute existed. */
  indent?: number;
  /** OOXML signed `relativeIndent`; present only when the source attribute existed. */
  relativeIndent?: number;
  /** OOXML `shrinkToFit`; present only when the source attribute existed. */
  shrinkToFit?: boolean;
  /** OOXML `readingOrder`; present only when the source attribute existed. */
  readingOrder?: number;
  /** Index into `<cellStyleXfs>` of the named style this format inherits
   *  from. Always 0 for a record read back from `getCellStyleXf`. */
  xfId?: number;
}

/** How an OOXML `<color>` element expressed its value.
 *  Mirrors `formulon::io::ColorSpec`.
 *
 *  A record read from a workbook carries the original specification here.
 *  When `kind` is non-zero, this selector is authoritative; the sibling
 *  `*Argb` field is not an Excel-rendered value for theme/indexed/auto
 *  selectors. It is literal RGB for `kind === 1`, or a compatibility
 *  fallback otherwise. Leave `kind` at 0 on a record built from scratch and
 *  the writer emits the sibling `*Argb` as `rgb`. */
export interface ColorSpec {
  /** 0=none, 1=rgb, 2=theme, 3=indexed, 4=auto. */
  kind: number;
  /** AARRGGBB; meaningful when `kind` is 1. */
  rgb: number;
  /** Theme index; meaningful when `kind` is 2. */
  theme: number;
  /** Theme tint in -1..1; meaningful when `kind` is 2. */
  tint: number;
  /** Legacy palette index; meaningful when `kind` is 3. */
  indexed: number;
}

/** Plain-data shape of a font record. Mirrors `formulon::io::FontRecord`. */
export interface FontRecord {
  name: string;
  size: number;
  bold: boolean;
  italic: boolean;
  strike: boolean;
  /** Whether the source carried a `<b>` element at all. An absent element
   *  on a differential font means "leave bold unchanged"; `hasBold` with
   *  `bold: false` means "switch bold off". */
  hasBold: boolean;
  /** Whether the source carried an `<i>` element at all. */
  hasItalic: boolean;
  /** Whether the source carried a `<strike>` element at all. */
  hasStrike: boolean;
  /** 0=none, 1=single, 2=double, 3=singleAccounting, 4=doubleAccounting. */
  underline: number;
  /** 0=baseline, 1=superscript, 2=subscript. */
  vertAlign: number;
  /** Whether the source carried a `<family>` element. */
  hasFamily: boolean;
  /** OOXML font-family class (0..5). */
  family: number;
  /** Whether the source carried a `<charset>` element. */
  hasCharset: boolean;
  /** OOXML charset codepage id (e.g. 128 = Shift_JIS). */
  charset: number;
  /** AARRGGBB literal RGB or compatibility fallback; not a resolved
   *  theme/indexed/auto render colour. */
  colorArgb: number;
  /** Original `<color>` specification, preserved for round-tripping. */
  color: ColorSpec;
}

/** Plain-data shape of a fill record. Mirrors `formulon::io::FillRecord`. */
export interface FillRecord {
  /** OOXML pattern ordinal: 0=none, 1=solid, 2..18=standard pattern set. */
  pattern: number;
  /** Foreground literal RGB or compatibility fallback. */
  fgArgb: number;
  /** Background literal RGB or compatibility fallback. */
  bgArgb: number;
  /** Original `<fgColor>` specification, preserved for round-tripping. */
  fg: ColorSpec;
  /** Original `<bgColor>` specification, preserved for round-tripping. */
  bg: ColorSpec;
}

/** One side of a `BorderRecord`. */
export interface BorderSide {
  /** OOXML border-style ordinal: 0=none, 1=thin, ..., 13=slantDashDot. */
  style: number;
  /** AARRGGBB literal RGB or compatibility fallback; not a resolved
   *  theme/indexed/auto render colour. */
  colorArgb: number;
  /** Original `<color>` specification, preserved for round-tripping. */
  color: ColorSpec;
}

/** Plain-data shape of a border record. Mirrors `formulon::io::BorderRecord`. */
export interface BorderRecord {
  left: BorderSide;
  right: BorderSide;
  top: BorderSide;
  bottom: BorderSide;
  diagonal: BorderSide;
  diagonalUp: boolean;
  diagonalDown: boolean;
}

/** Plain-data shape of an `<xf>` record. Mirrors `formulon::io::CellXf`. */
export interface CellXf {
  fontIndex: number;
  fillIndex: number;
  borderIndex: number;
  numFmtId: number;
  /** 0=general, 1=left, 2=center, 3=right, 4=fill, 5=justify, 6=centerContinuous, 7=distributed. */
  horizontalAlign: number;
  /** 0=top, 1=center, 2=bottom, 3=justify, 4=distributed. */
  verticalAlign: number;
  wrapText: boolean;
  /** OOXML `justifyLastLine`; effective with distributed alignment. Defaults to false. */
  justifyLastLine?: boolean;
  /** Whether to preserve an explicit alignment child. Omit to infer from supplied alignment fields. */
  hasAlignment?: boolean;
  /** Whether `horizontalAlign` was explicitly supplied; defaults to inference. */
  hasHorizontalAlign?: boolean;
  /** Whether `verticalAlign` was explicitly supplied; defaults to inference. */
  hasVerticalAlign?: boolean;
  /** Whether `wrapText` was explicitly supplied; defaults to inference. */
  hasWrapText?: boolean;
  /** Whether `justifyLastLine` was explicitly supplied; defaults to inference. */
  hasJustifyLastLine?: boolean;
  /** OOXML `textRotation` (0..180 or 255). Omitted when absent. */
  textRotation?: number;
  /** OOXML `indent` (0..255). Omitted when absent. */
  indent?: number;
  /** OOXML signed `relativeIndent`. Omitted when absent. */
  relativeIndent?: number;
  /** OOXML `shrinkToFit`; omitted when absent so explicit false is retained. */
  shrinkToFit?: boolean;
  /** OOXML `readingOrder` (0=context, 1=LTR, 2=RTL). Omitted when absent. */
  readingOrder?: number;
  /** Index of the named style in `<cellStyleXfs>` this format inherits from.
   *  Only meaningful for `addXf`; must name an entry created by
   *  `addCellStyleXf`. Defaults to 0. */
  xfId?: number;
}

/** Return type of `Workbook.getFont(fontIndex)`. */
export interface FontResult extends FontRecord {
  status: Status;
}

/** Return type of `Workbook.getFill(fillIndex)`. */
export interface FillResult extends FillRecord {
  status: Status;
}

/** Return type of `Workbook.getBorder(borderIndex)`. */
export interface BorderResult extends BorderRecord {
  status: Status;
}

/** Return type of `Workbook.getNumFmt(numFmtId)`. */
export interface NumFmtResult {
  status: Status;
  numFmtId: number;
  formatCode: string;
}

/** Return type of `Workbook.getDxf(dxfIndex)`.
 *  Optional properties are present only when that `<dxf>` child exists. */
export interface DxfRecord {
  font?: FontRecord;
  fill?: FillRecord;
  border?: BorderRecord;
  numFmt?: {
    numFmtId: number;
    formatCode: string;
  };
  /** Serialized OOXML `<alignment .../>` child; lexical formatting may normalize on load. */
  alignmentXml?: string;
  /** Serialized OOXML `<protection .../>` child; lexical formatting may normalize on load. */
  protectionXml?: string;
}

export interface DxfResult extends DxfRecord {
  status: Status;
}

/** Return type of `Workbook.getLambdaText(sheet, row, col)`. The
 *  rendered text never carries a leading `=` and is suitable for
 *  passing back through `setFormula`. `kInvalidArgument` surfaces when
 *  the cell is absent or its cached value is not a lambda. */
export interface LambdaTextResult {
  status: Status;
  /** Excel formula text in `LAMBDA(p1,p2,body)` form. Empty string
   *  when `status` is non-OK. */
  text: string;
}

/** Minimum severity for the engine's structured log stream. Mirrors
 *  `fm_log_level_t`. `Off` discards every record and is the default: an
 *  embedded library must not write to the host's stderr unless asked. */
export enum LogLevel {
  Debug = 0,
  Info = 1,
  Warn = 2,
  Error = 3,
  Off = 4,
}

/** External-link kinds. Mirrors
 *  `formulon::io::ExternalLinkRecord::Kind`. */
export enum ExternalLinkKind {
  Unknown = 0,
  ExternalBook = 1,
  Ole = 2,
  Dde = 3,
}

/** Element type returned by `Workbook.getExternalLinks()`. Mirrors
 *  `formulon::io::ExternalLinkRecord`. The body part itself is not
 *  exposed (it round-trips through the OOXML passthrough mechanism);
 *  this surface only enumerates the cross-workbook references and
 *  their resolved target URLs. */
export interface ExternalLinkRecord {
  /** 1-based document order matching `<externalReferences>` in
   *  `xl/workbook.xml`. */
  index: number;
  /** Workbook-rels Id ("rId3" etc.). */
  relId: string;
  /** Resolved package-relative path of the body part (e.g.
   *  `xl/externalLinks/externalLink1.xml`). */
  partPath: string;
  /** Remote workbook URL (e.g. `file:///path/book.xlsx`,
   *  `https://example/sheet.xlsx`). Empty when the per-link rels file
   *  was missing or unparseable. */
  target: string;
  /** Whether the per-link rels relationship was emitted with
   *  `TargetMode="External"` (the common case). */
  targetExternal: boolean;
  /** One of `ExternalLinkKind.*`. */
  kind: number;
}

/** Return type of `Workbook.getCellStyle(index)`. Mirrors
 *  `formulon::io::CellStyleRecord`. `xfId` indexes into the named-style
 *  xf table reachable via `Workbook.getCellStyleXf(...)`. */
export interface CellStyleResult {
  status: Status;
  /** Display name (e.g. "Normal", "Heading 1", or a user-defined label). */
  name: string;
  /** Index into the `<cellStyleXfs>` table. */
  xfId: number;
  /** OOXML built-in style ordinal (`0..47`), or `0xFFFFFFFF` for custom
   *  entries that did not carry a `builtinId` attribute. */
  builtinId: number;
  /** Outline level for built-in heading styles (0 otherwise). */
  iLevel: number;
  hidden: boolean;
  customBuiltin: boolean;
}

/** Return type of `Workbook.addFont/Fill/Border/Xf(...)`. The
 *  add-functions deduplicate against existing entries via linear
 *  search; `index` is either the matched index or the freshly-appended
 *  index. */
export interface AddStyleResult {
  status: Status;
  index: number;
}

/** Return type of `Workbook.addNumFmt(formatCode)`. The id may reuse any
 *  existing effective mapping (built-in or custom, including a custom record
 *  overriding a built-in slot), or be a freshly-assigned custom id (`>= 164`). */
export interface AddNumFmtResult {
  status: Status;
  numFmtId: number;
}

/** Range used by `partialRecalc`. */
export interface RecalcViewport {
  sheet: number;
  firstRow: number;
  lastRow: number;
  firstCol: number;
  lastCol: number;
}

/** Return type of `Workbook.partialRecalc(viewport)`. */
export interface PartialRecalcResult {
  status: Status;
  /** Number of cells the engine actually evaluated. */
  recomputed: number;
}

/** Return type of `Workbook.getIterative()`. */
export interface IterativeSettingsResult {
  status: Status;
  /** Whether iterative calculation is switched on. */
  enabled: boolean;
  /** Stored iteration cap; meaningful even while `enabled` is false. */
  maxIterations: number;
  /** Stored convergence threshold; meaningful even while `enabled` is false. */
  maxChange: number;
}

/** Result of `Workbook.getSheetAutoFilterXml`. The empty `xml` string means
 * that the sheet has no AutoFilter definition. */
export interface SheetAutoFilterXmlResult {
  status: Status;
  xml: string;
}

/** Result of the five raw print-settings getters. The empty `xml` string
 * means the sheet declares no such element. */
export interface SheetPrintXmlResult {
  status: Status;
  xml: string;
}

/** Result of `Workbook.getSheetPrintArea`. `ranges` is comma-separated A1
 * (`"A1:F8"`, `"A1:B10,D5:E20"`), empty when no print area is declared. A
 * whole-axis area reads back expanded to explicit corners. */
export interface SheetPrintAreaResult {
  status: Status;
  ranges: string;
}

/** Result of `Workbook.getSheetPrintTitles`. Either span is empty when
 * that axis has no repeat setting. */
export interface SheetPrintTitlesResult {
  status: Status;
  /** Whole-row span, e.g. `"1:2"`. */
  repeatRows: string;
  /** Whole-column span, e.g. `"A:A"`. */
  repeatCols: string;
}

/** One manual page break. `id` is the 0-based row / column the break
 * precedes; `min` / `max` bound it on the perpendicular axis. `manual` is
 * false for a break Excel computed rather than one a user placed. */
export interface PageBreak {
  id: number;
  min: number;
  max: number;
  manual: boolean;
}

/** Result of `Workbook.getSheetRowBreaks` / `getSheetColBreaks`. */
export interface SheetPageBreaksResult {
  status: Status;
  breaks: PageBreak[];
}

/** Page orientation as `<pageSetup orientation>` encodes it: `0` leaves the
 * attribute off (the printer default), `1` is portrait, `2` is landscape.
 *
 * Declared as a union rather than an `enum` so it stays a compile-time-only
 * type: there is no runtime table to keep in step across the two packages,
 * and the three values are the whole domain. */
export type Orientation = 0 | 1 | 2;

/** Partial `<pageSetup>` update. Only the keys present are applied; every
 * other attribute in the stored XML — including ones this engine does not
 * model, such as `r:id`, `horizontalDpi` and `copies` — is left alone.
 *
 * `fitToPage` is not a `<pageSetup>` attribute: it is routed to
 * `<sheetPr><pageSetUpPr>`. It selects the fit-to-page mode, and
 * `fitToWidth` / `fitToHeight` state the target, so "fit onto one page"
 * needs all three. */
export interface PageSetupInput {
  orientation?: Orientation;
  /** OOXML paperSize code; 9 = A4. */
  paperSize?: number;
  /** Print scale percentage. Outside `[10, 400]` is rejected, not clamped. */
  scale?: number;
  fitToWidth?: number;
  fitToHeight?: number;
  fitToPage?: boolean;
}

/** Result of `Workbook.getSheetPageSetup`. Every value carries the setting
 * in force; the `*Stated` flags say whether the XML declares it, which is
 * the only way to tell `scale="100"` from an absent attribute. */
export interface SheetPageSetupResult {
  status: Status;
  orientation: Orientation;
  paperSize: number;
  scale: number;
  fitToWidth: number;
  fitToHeight: number;
  fitToPage: boolean;
  orientationStated: boolean;
  paperSizeStated: boolean;
  scaleStated: boolean;
  fitToWidthStated: boolean;
  fitToHeightStated: boolean;
  fitToPageStated: boolean;
}

/** Partial `<pageMargins>` update, in inches. A negative, infinite or NaN
 * margin is rejected: the paginator subtracts these from the paper. */
export interface PageMarginsInput {
  left?: number;
  right?: number;
  top?: number;
  bottom?: number;
  header?: number;
  footer?: number;
}

/** Result of `Workbook.getSheetPageMargins`, in inches. */
export interface SheetPageMarginsResult {
  status: Status;
  left: number;
  right: number;
  top: number;
  bottom: number;
  header: number;
  footer: number;
  leftStated: boolean;
  rightStated: boolean;
  topStated: boolean;
  bottomStated: boolean;
  headerStated: boolean;
  footerStated: boolean;
}

/** Partial `<printOptions>` update. */
export interface PrintOptionsInput {
  gridLines?: boolean;
  headings?: boolean;
  horizontalCentered?: boolean;
  verticalCentered?: boolean;
}

/** Partial `<headerFooter>` update.
 *
 * Each section is tri-state: omit the key to leave it untouched, pass `""`
 * to clear it, pass anything else to replace it. Section text is the
 * decoded string, so Excel's formatting codes go in plainly
 * (`'&C&"MS Gothic"Report &P/&N'`); the engine applies the XML escaping
 * that turns each leading `&` into the `&amp;` Excel stores. */
export interface HeaderFooterInput {
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
  differentOddEven?: boolean;
  differentFirst?: boolean;
  scaleWithDoc?: boolean;
  alignWithMargins?: boolean;
}

/** Iterative-solver progress callback. Receives the current
 *  iteration number, the maximum residual seen, and the configured
 *  iteration cap. Returning `false` (or any falsy value) aborts the
 *  solve; returning `true` (or `undefined`) continues.
 *
 *  Throwing also aborts the solve, and the exception does not reach the
 *  caller: the recalc that ran the callback returns `{ok: false}` with
 *  status 7003 (`kBindingCallbackException`) instead. Cells committed by
 *  the sweeps that already ran keep their values, and the workbook stays
 *  usable — the next recalc behaves as if the aborted one had been
 *  cancelled deliberately. */
export type IterativeProgressCallback = (
  iteration: number,
  maxResidual: number,
  maxIterations: number,
) => boolean | undefined | void;

/** Workbook handle. Always release with `delete()` when finished. */
export interface Workbook {
  /** True when the wrapper holds a live native handle. */
  isValid(): boolean;
  /** Releases the native handle. The instance must not be used afterwards. */
  delete(): void;

  save(): SaveResult;
  /** Serialises using an explicit container `format` (see `WorkbookFormat`). */
  saveAs(format: WorkbookFormat | number): SaveResult;
  /** Serialises using an explicit format and reports loss/defer counters. */
  saveWithDiagnostics(format: WorkbookFormat | number): SaveDiagnosticsResult;
  /** Returns the loss counters captured when this handle was loaded. */
  readDiagnostics(): ReadDiagnosticsResult;
  addSheet(name: string): Status;
  /** Removes the sheet at `index`. */
  removeSheet(index: number): Status;
  /** Renames the sheet at `index`. */
  renameSheet(index: number, newName: string): Status;
  /** Moves the sheet from `fromIdx` to `toIdx` (post-removal index). */
  moveSheet(fromIdx: number, toIdx: number): Status;
  sheetCount(): number;
  sheetName(idx: number): StringResult;

  setNumber(sheet: number, row: number, col: number, value: number): Status;
  setBool(sheet: number, row: number, col: number, value: boolean): Status;
  /** Stores a static Excel error literal; `errorCode` is an ErrorCode ordinal. */
  setError(sheet: number, row: number, col: number, errorCode: number): Status;
  setText(sheet: number, row: number, col: number, text: string): Status;
  /** Stores (or, when empty, clears) the cell's OOXML phonetic guide (`<rPh>`). */
  setCellPhonetic(sheet: number, row: number, col: number, phonetic: string): Status;
  setBlank(sheet: number, row: number, col: number): Status;
  setFormula(sheet: number, row: number, col: number, formula: string): Status;

  getValue(sheet: number, row: number, col: number): CellResult;

  /** Evaluates `formula` as if entered at `(sheet, row, col)` and returns a
   *  single scalar result, without mutating the workbook. Local and
   *  cross-sheet references, defined names, and `ROW()` / `COLUMN()` resolve
   *  relative to the anchor. An array / spill result is reduced to its
   *  top-left element (a pragmatic API shape, not Excel implicit
   *  intersection or spilling; multi-cell results are a Phase 2 follow-up).
   *  Note: a self-reference reads the target cell's cached value rather than
   *  raising `#REF!`, since the ad-hoc formula never joins the dep graph. */
  evaluateFormulaText(sheet: number, row: number, col: number, formula: string): EvalResult;

  /** Evaluates `formula` as if entered at `(sheet, row, col)` and returns the
   *  whole multi-cell result without mutating the workbook. Unlike
   *  `evaluateFormulaText` (which reduces an array to its top-left element),
   *  a dynamic-array formula such as `=SEQUENCE(2,3)` yields the full
   *  `rows` x `cols` grid in `result.cells` (row-major, `cells[r][c]`); a
   *  scalar result is reported as a 1x1 array. Same read-only purity and
   *  self-reference caveat as `evaluateFormulaText`. */
  evaluateFormulaArray(sheet: number, row: number, col: number, formula: string): EvalArrayResult;

  /** Evaluates `formula` as a conditional-formatting predicate anchored at
   *  `(sheet, row, col)`, with relative references written relative to
   *  `(anchorRow, anchorCol)` (the CF-applied range's top-left). The result
   *  is coerced with Excel's CF rules: error / blank / text / numeric-zero
   *  yield `false`, any non-zero number yields `true`. Read-only. */
  evaluateConditionalFormula(
    sheet: number,
    row: number,
    col: number,
    anchorRow: number,
    anchorCol: number,
    formula: string,
  ): EvalResult;

  /** Renders the lambda value stored at `(sheet, row, col)` as Excel
   *  formula text. Returns `kInvalidArgument` when the cell is absent
   *  or its cached value is not a lambda. */
  getLambdaText(sheet: number, row: number, col: number): LambdaTextResult;
  /** Returns the cell's OOXML phonetic guide (`<rPh>`), or an empty string. */
  getCellPhonetic(sheet: number, row: number, col: number): StringResult;

  /** Recalculates all dirty cells serially on the caller thread.
   *
   *  Reports status 7003 (`kBindingCallbackException`) when an installed
   *  `IterativeProgressCallback` threw during the pass. */
  recalc(): Status;
  /**
   * Recalculates all dirty cells using the parallel SCC scheduler and returns
   * per-call telemetry. This call is synchronous: it returns after all
   * workers have joined. `threadCount` must be a finite integer: 0 selects
   * automatic detection (capped at 8), 1 keeps evaluation on the caller and
   * starts no workers, and 2..8 set the worker upper bound. Missing,
   * fractional, non-finite, negative, or above-8 values fail with
   * `kInvalidArgument`.
   * The `uint64_t` counters in `stats` are JavaScript numbers and are exact
   * through `Number.MAX_SAFE_INTEGER` (2^53 - 1).
   * As with `recalc`, a throwing `IterativeProgressCallback` is reported as
   * status 7003 with the counters left at zero.
   */
  recalcParallel(threadCount: number): ParallelRecalcResult;
  /** Recalculates only cells touched by the supplied viewport. Reports
   *  status 7003 with `recomputed` at zero when a throwing
   *  `IterativeProgressCallback` cut the pass short. */
  partialRecalc(viewport: RecalcViewport): PartialRecalcResult;

  setIterative(enabled: boolean, maxIterations: number, maxChange: number): Status;
  /** Reads back the stored iterative-calculation settings. The cap and
   *  threshold are the stored values even while `enabled` is false, so a
   *  host can render Excel's iterative-calculation dialog before the user
   *  has switched iteration on. */
  getIterative(): IterativeSettingsResult;
  /** Installs (or, when passed `null`, clears) a JS callback invoked
   *  after each Gauss-Seidel sweep. The callback belongs to this
   *  `Workbook`: installing one on another instance leaves this one's
   *  registration alone.
   *
   *  A value that is not callable is accepted here and reported when the
   *  solver tries to use it, the same way a throwing callback is. */
  setIterativeProgress(callback: IterativeProgressCallback | null): Status;

  /**
   * Workbook-level calc mode (Excel `<calcPr calcMode>` policy).
   *
   * Returns one of `CalcMode` codes. The engine itself does NOT gate
   * evaluation on this value — every `recalc()` call honours all dirty
   * cells. The mode is preserved as round-trip metadata and surfaced
   * here so the UI can mirror Excel's user-visible state.
   */
  calcMode(): CalcMode;
  setCalcMode(mode: CalcMode): Status;

  /**
   * The workbook's pinned wall-clock reading, or `null` when it follows the
   * host clock (the default).
   *
   * Pinning gives `NOW()`, `TODAY()` and the pivot relative-period filters
   * ("this month", "year to date", ...) one instant to agree on. Without a
   * pin each reads the clock independently, which makes a recalc internally
   * inconsistent across a midnight boundary and makes any such result
   * untestable.
   *
   * Local civil fields are carried rather than a timestamp so the reading has
   * no residual timezone interpretation: the same values reproduce the same
   * results in any zone. This is model state, not file state — `save()` does
   * not record it and a reloaded workbook comes back unpinned.
   *
   * Cached formula values are not recomputed by `setPinnedNow`; drive
   * `recalc()` to see the pin take effect on existing cells.
   *
   * `setPinnedNow` returns `kInvalidArgument` unless `year` is in
   * `[1900, 9999]`, `month` in `[1, 12]`, `day` within that month's real
   * length, `hour` in `[0, 23]`, and `minute` / `second` in `[0, 59]`. The
   * pin is a calendar instant, not a normalising constructor: a month of 13
   * is rejected rather than rolled into the next year.
   */
  pinnedNow(): CivilTime | null;
  setPinnedNow(year: number, month: number, day: number, hour: number, minute: number, second: number): Status;
  clearPinnedNow(): Status;

  /**
   * Full formula-behaviour profile id. Defaults to `win-365-ja_JP`.
   */
  excelProfileId(): ExcelProfileId;
  setExcelProfileId(profileId: ExcelProfileId): Status;

  /** Inserts `count` rows at `row` on `sheet` and rewrites cross-workbook
   *  references to follow the shift. */
  insertRows(sheet: number, row: number, count: number): Status;
  /** Deletes `count` rows starting at `row` on `sheet`. References that
   *  fall inside the deleted interval collapse to `#REF!`. */
  deleteRows(sheet: number, row: number, count: number): Status;
  /** Inserts `count` columns at `col` on `sheet`. */
  insertCols(sheet: number, col: number, count: number): Status;
  /** Deletes `count` columns starting at `col` on `sheet`. */
  deleteCols(sheet: number, col: number, count: number): Status;

  cellCount(sheet: number): number;
  cellAt(sheet: number, idx: number): CellEntry;

  definedNameCount(): number;
  definedNameAt(idx: number): DefinedNameEntry;
  /** Adds, replaces, or (when `formula` is empty) removes a workbook-
   *  scoped defined name. */
  setDefinedName(name: string, formula: string): Status;
  /** Adds, replaces, or removes a defined name in workbook scope (-1)
   *  or a sheet-local scope (0-based sheet index). */
  setDefinedNameScoped(name: string, formula: string, localSheetId: number): Status;

  tableCount(): number;
  tableAt(idx: number): TableEntry;
  createTable(input: TableInput): AddStyleResult;
  /** Partially updates a table. Omitted fields preserve their current
   *  values. A supplied `ref` must keep the table's existing column count;
   *  an empty `styleName` clears the style. */
  updateTable(idx: number, input: TableUpdateInput): Status;
  removeTable(idx: number): Status;

  /** Reads the complete worksheet-level `<autoFilter>` XML fragment. */
  getSheetAutoFilterXml(sheet: number): SheetAutoFilterXmlResult;
  /** Replaces the worksheet-level `<autoFilter>` XML fragment; empty clears it. */
  setSheetAutoFilterXml(sheet: number, xml: string): Status;

  passthroughCount(): number;
  passthroughAt(idx: number): PassthroughEntry;

  /** Returns the number of PivotTables anchored on `sheet`. */
  pivotCount(sheet: number): number;
  /** Evaluates and projects a PivotTable into concrete grid cells. */
  pivotLayout(sheet: number, pivotIndex: number): PivotLayoutResult;

  // ---- PivotCache mutation -----------------------------------------------
  /** Returns the number of pivot caches owned by the workbook. */
  pivotCacheCount(): number;
  /** Returns `{ status, index: cacheId }` for the cache at flat
   *  index `idx`. */
  pivotCacheIdAt(idx: number): AddStyleResult;
  /** Creates a new empty pivot cache. Pass `0` for `requestedId` to
   *  auto-assign. Returns `{ status, index: cacheId }`. */
  pivotCacheCreate(requestedId: number): AddStyleResult;
  /** Removes the pivot cache with id `cacheId`. Fails if any pivot
   *  table still references it. */
  pivotCacheRemove(cacheId: number): Status;
  /** Reads the cache's worksheet source range / defined-name metadata. */
  pivotCacheGetWorksheetSource(cacheId: number): PivotWorksheetSourceResult;
  /** Sets or clears the cache's worksheet source metadata. */
  pivotCacheSetWorksheetSource(cacheId: number, source: PivotWorksheetSource): Status;

  /** Number of fields on the cache identified by `cacheId`. */
  pivotCacheFieldCount(cacheId: number): number;
  /** Reads the name of the cache field at `fieldIdx`. */
  pivotCacheFieldName(cacheId: number, fieldIdx: number): StringResult;
  /** Appends a new field with the given UTF-8 name to the cache.
   *  `index` carries the new field's index. */
  pivotCacheFieldAdd(cacheId: number, name: string): AddStyleResult;
  /** Drops every field (and every record) from the cache. */
  pivotCacheFieldClear(cacheId: number): Status;

  /** Number of shared items configured on cache field `fieldIdx`. */
  pivotCacheFieldSharedItemCount(cacheId: number, fieldIdx: number): number;
  /** Appends a numeric shared item to cache field `fieldIdx`. */
  pivotCacheFieldAddSharedItemNumber(cacheId: number, fieldIdx: number, value: number): Status;
  /** Appends a text shared item to cache field `fieldIdx`. */
  pivotCacheFieldAddSharedItemText(cacheId: number, fieldIdx: number, value: string): Status;
  /** Appends a boolean shared item to cache field `fieldIdx`. */
  pivotCacheFieldAddSharedItemBool(cacheId: number, fieldIdx: number, value: boolean): Status;
  /** Appends a blank shared item to cache field `fieldIdx`. */
  pivotCacheFieldAddSharedItemBlank(cacheId: number, fieldIdx: number): Status;
  /** Appends an Excel error shared item to cache field `fieldIdx`. */
  pivotCacheFieldAddSharedItemError(cacheId: number, fieldIdx: number, errorCode: number): Status;
  /** Drops every shared item from cache field `fieldIdx`. */
  pivotCacheFieldClearSharedItems(cacheId: number, fieldIdx: number): Status;

  /** Returns the number of records on the cache. */
  pivotCacheRecordCount(cacheId: number): number;
  /** Appends a new empty record. `index` carries the new record's index. */
  pivotCacheRecordAdd(cacheId: number): AddStyleResult;
  /** Drops every record from the cache. */
  pivotCacheRecordClear(cacheId: number): Status;
  /** Sets cell `(recordIdx, fieldIdx)` to a numeric value. */
  pivotCacheRecordSetNumber(cacheId: number, recordIdx: number, fieldIdx: number, value: number): Status;
  /** Sets cell `(recordIdx, fieldIdx)` to a UTF-8 text value. */
  pivotCacheRecordSetText(cacheId: number, recordIdx: number, fieldIdx: number, value: string): Status;
  /** Sets cell `(recordIdx, fieldIdx)` to a boolean value. */
  pivotCacheRecordSetBool(cacheId: number, recordIdx: number, fieldIdx: number, value: boolean): Status;
  /** Sets cell `(recordIdx, fieldIdx)` to Blank. */
  pivotCacheRecordSetBlank(cacheId: number, recordIdx: number, fieldIdx: number): Status;
  /** Sets cell `(recordIdx, fieldIdx)` to an Excel error value. */
  pivotCacheRecordSetError(cacheId: number, recordIdx: number, fieldIdx: number, errorCode: number): Status;

  // ---- PivotTable mutation -----------------------------------------------
  /** Creates a new empty pivot table. Returns `{ status, index: pivotIdx }`. */
  pivotCreate(sheet: number, name: string, cacheId: number, anchorRow: number, anchorCol: number): AddStyleResult;
  /** Removes the pivot table at `pivotIdx`. */
  pivotRemove(sheet: number, pivotIdx: number): Status;
  /** Renames the pivot table. */
  pivotSetName(sheet: number, pivotIdx: number, name: string): Status;
  /** Updates the pivot's anchor cell and span. */
  pivotSetAnchor(
    sheet: number,
    pivotIdx: number,
    anchorRow: number,
    anchorCol: number,
    spanRows: number,
    spanCols: number,
  ): Status;
  /** Toggles the row / column grand total bands on the pivot. */
  pivotSetGrandTotals(sheet: number, pivotIdx: number, rowsEnabled: boolean, colsEnabled: boolean): Status;
  /** Reads the pivot's compact / tabular / outline report layout. */
  pivotGetLayout(sheet: number, pivotIdx: number): PivotReportLayoutResult;
  /** Sets the pivot's compact / tabular / outline report layout. */
  pivotSetLayout(sheet: number, pivotIdx: number, layout: PivotReportLayout): Status;

  /** Number of fields configured on the pivot. */
  pivotFieldCount(sheet: number, pivotIdx: number): number;
  /** Appends a new field. Returns `{ status, index: fieldIdx }`. */
  pivotFieldAdd(sheet: number, pivotIdx: number, spec: PivotFieldSpec): AddStyleResult;
  /** Drops every field from the pivot. */
  pivotFieldClear(sheet: number, pivotIdx: number): Status;
  /** Sets the axis of pivot field `fieldIdx`. */
  pivotFieldSetAxis(sheet: number, pivotIdx: number, fieldIdx: number, axis: PivotAxis): Status;
  /** Sets the sort directive on pivot field `fieldIdx`. Pass an empty
   *  `byField` to clear the by-field key. */
  pivotFieldSetSort(sheet: number, pivotIdx: number, fieldIdx: number, ascending: boolean, byField: string): Status;
  /** Sets the `subtotal_top` flag on pivot field `fieldIdx`. */
  pivotFieldSetSubtotalTop(sheet: number, pivotIdx: number, fieldIdx: number, top: boolean): Status;
  /** Appends an aggregation to pivot field `fieldIdx`. */
  pivotFieldAddAggregation(sheet: number, pivotIdx: number, fieldIdx: number, agg: PivotAggregation): Status;
  /** Drops every aggregation from pivot field `fieldIdx`. */
  pivotFieldClearAggregations(sheet: number, pivotIdx: number, fieldIdx: number): Status;
  /** Appends a manual-filter item to pivot field `fieldIdx`, addressed by
   *  its label. The item carries no cache binding, so the filter engine
   *  matches source records by comparing their rendered label against
   *  `name`. An empty `name` therefore cannot express the blank item,
   *  which has no label of its own; use `pivotFieldAddItemAt`. */
  pivotFieldAddItem(sheet: number, pivotIdx: number, fieldIdx: number, name: string, visible: boolean): Status;
  /** Appends a manual-filter item to pivot field `fieldIdx`, addressed by
   *  its position in the bound cache field's shared items -- the same
   *  index space as OOXML `<item x="N">`. The label is resolved from that
   *  shared item, which makes this the form that can express the blank
   *  item. `cacheIndex` is not validated: the cache field may be filled in
   *  after the pivot definition, and an index that resolves to nothing
   *  yields an item with an empty label that filters nothing. */
  pivotFieldAddItemAt(sheet: number, pivotIdx: number, fieldIdx: number, cacheIndex: number, visible: boolean): Status;
  /** Drops every manual-filter item from pivot field `fieldIdx`. */
  pivotFieldClearItems(sheet: number, pivotIdx: number, fieldIdx: number): Status;
  /** Toggles the visibility of item `itemIdx` on field `fieldIdx`. */
  pivotFieldSetItemVisible(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    itemIdx: number,
    visible: boolean,
  ): Status;
  /** Appends a subtotal-fn entry to pivot field `fieldIdx`. */
  pivotFieldAddSubtotalFn(sheet: number, pivotIdx: number, fieldIdx: number, agg: PivotAggregation): Status;
  /** Drops every subtotal-fn entry from pivot field `fieldIdx`. */
  pivotFieldClearSubtotalFns(sheet: number, pivotIdx: number, fieldIdx: number): Status;
  /** Configures date-grouping on pivot field `fieldIdx`. Pass `-1` for
   *  `startYear` / `endYear` to leave the bound unset. */
  pivotFieldSetDateGroup(
    sheet: number,
    pivotIdx: number,
    fieldIdx: number,
    granularity: PivotDateGrouping,
    calendar: PivotCalendar,
    startYear: number,
    endYear: number,
  ): Status;
  /** Removes the date-grouping config from pivot field `fieldIdx`. */
  pivotFieldClearDateGroup(sheet: number, pivotIdx: number, fieldIdx: number): Status;
  /** Sets the OOXML number-format string on pivot field `fieldIdx`.
   *  Pass an empty string to clear. */
  pivotFieldSetNumberFormat(sheet: number, pivotIdx: number, fieldIdx: number, utf8: string): Status;

  /** Replaces the row-axis field order with `indices`. Each entry must
   *  be `< pivotFieldCount`. Pass an empty array to clear. */
  pivotSetRowFieldOrder(sheet: number, pivotIdx: number, indices: ReadonlyArray<number>): Status;
  /** Replaces the column-axis field order. Same contract as rows. */
  pivotSetColFieldOrder(sheet: number, pivotIdx: number, indices: ReadonlyArray<number>): Status;

  /** Number of `<dataField>` entries on the pivot. */
  pivotDataFieldCount(sheet: number, pivotIdx: number): number;
  /** Appends a new data-field entry. */
  pivotDataFieldAdd(sheet: number, pivotIdx: number, spec: PivotDataFieldSpec): AddStyleResult;
  /** Drops every data-field entry from the pivot. */
  pivotDataFieldClear(sheet: number, pivotIdx: number): Status;
  /** Replaces the data-field entry at `dataFieldIdx` in place. */
  pivotDataFieldSet(sheet: number, pivotIdx: number, dataFieldIdx: number, spec: PivotDataFieldSpec): Status;

  /** Number of active (slicer-applied) filters on the pivot. Counts only
   *  the entries this session added -- see {@link PivotFilterResult}. */
  pivotFilterCount(sheet: number, pivotIdx: number): number;
  /** Appends an active filter. The entry is session state and is not
   *  written by `save`; see {@link PivotFilterResult}. */
  pivotFilterAdd(sheet: number, pivotIdx: number, spec: PivotFilterSpec): Status;
  /** Reads the active filter at `filterIdx`. The list is session state:
   *  see {@link PivotFilterResult}. */
  pivotFilterAt(sheet: number, pivotIdx: number, filterIdx: number): PivotFilterResult;
  /** Drops every active filter from the pivot. */
  pivotFilterClear(sheet: number, pivotIdx: number): Status;
  /** Removes the active filter at `filterIdx`. */
  pivotFilterRemoveAt(sheet: number, pivotIdx: number, filterIdx: number): Status;

  /** Evaluates every CF block on `sheet` against the inclusive range
   *  `[(firstRow, firstCol), (lastRow, lastCol)]`. `todaySerial` pins the
   *  date basis for `TimePeriod` rules; pass `NaN` to disable them (no
   *  `TimePeriod` rule then matches). `0` does *not* disable them -- it is
   *  the valid serial for 1899-12-30. */
  evaluateCfRange(
    sheet: number,
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    todaySerial: number,
  ): CfRangeResult;

  /** Computes print-area page breaks and physical page count for `sheet`. */
  paginate(sheet: number): PaginationResult;

  /** Reads the worksheet's `<pageSetup>` fragment; empty when absent. */
  getSheetPageSetupXml(sheet: number): SheetPrintXmlResult;
  /** Replaces `<pageSetup>`; empty removes it and restores the defaults.
   *  A fragment carrying `r:id` is rejected unless the sheet already has a
   *  printerSettings part, since the reference would otherwise dangle. */
  setSheetPageSetupXml(sheet: number, xml: string): Status;
  /** Reads the worksheet's `<pageMargins>` fragment; empty when absent. */
  getSheetPageMarginsXml(sheet: number): SheetPrintXmlResult;
  /** Replaces `<pageMargins>`; empty removes it. */
  setSheetPageMarginsXml(sheet: number, xml: string): Status;
  /** Reads the worksheet's `<printOptions>` fragment; empty when absent. */
  getSheetPrintOptionsXml(sheet: number): SheetPrintXmlResult;
  /** Replaces `<printOptions>`; empty removes it. */
  setSheetPrintOptionsXml(sheet: number, xml: string): Status;
  /** Reads the worksheet's `<headerFooter>` fragment; empty when absent. */
  getSheetHeaderFooterXml(sheet: number): SheetPrintXmlResult;
  /** Replaces `<headerFooter>`; empty removes it. Header codes live in XML
   *  element text, so a hand-built fragment writes them escaped
   *  (`<oddHeader>&amp;C...</oddHeader>`); `setSheetHeaderFooter` takes the
   *  decoded text instead. */
  setSheetHeaderFooterXml(sheet: number, xml: string): Status;
  /** Reads the worksheet's `<sheetPr>` fragment; empty when absent. It also
   *  carries the tab colour and VBA code name, not just print settings. */
  getSheetSheetPrXml(sheet: number): SheetPrintXmlResult;
  /** Replaces `<sheetPr>`; empty removes it. */
  setSheetSheetPrXml(sheet: number, xml: string): Status;
  /** Sets or clears `<sheetPr><pageSetUpPr fitToPage>` without disturbing
   *  anything else `<sheetPr>` carries. */
  setSheetFitToPage(sheet: number, enabled: boolean): Status;

  /** Reads `_xlnm.Print_Area` as comma-separated A1 ranges. */
  getSheetPrintArea(sheet: number): SheetPrintAreaResult;
  /** Writes `_xlnm.Print_Area` from one or more comma-separated A1 ranges
   *  (`"A1:F8"`, `"A1:B10,D5:E20"`, `"A:D"`). Empty removes it. */
  setSheetPrintArea(sheet: number, rangesA1: string): Status;
  /** Reads `_xlnm.Print_Titles` as a row span and a column span. */
  getSheetPrintTitles(sheet: number): SheetPrintTitlesResult;
  /** Writes `_xlnm.Print_Titles`. Both empty removes it. */
  setSheetPrintTitles(sheet: number, repeatRows: string, repeatCols: string): Status;

  /** Upserts a row break before `row`, spanning the whole sheet. */
  addSheetRowBreak(sheet: number, row: number, manual: boolean): Status;
  /** Upserts a column break before `col`, spanning the whole sheet. */
  addSheetColBreak(sheet: number, col: number, manual: boolean): Status;
  /** Removes the row break at `row`. An absent break is a no-op. */
  removeSheetRowBreak(sheet: number, row: number): Status;
  /** Removes the column break at `col`. An absent break is a no-op. */
  removeSheetColBreak(sheet: number, col: number): Status;
  /** Removes every manual break on both axes. */
  clearSheetBreaks(sheet: number): Status;
  /** Returns the row breaks, ascending by row. */
  getSheetRowBreaks(sheet: number): SheetPageBreaksResult;
  /** Returns the column breaks, ascending by column. */
  getSheetColBreaks(sheet: number): SheetPageBreaksResult;

  /** Applies a partial `<pageSetup>` update. */
  setSheetPageSetup(sheet: number, setup: PageSetupInput): Status;
  /** Applies a partial `<pageMargins>` update, in inches. */
  setSheetPageMargins(sheet: number, margins: PageMarginsInput): Status;
  /** Applies a partial `<printOptions>` update. */
  setSheetPrintOptions(sheet: number, options: PrintOptionsInput): Status;
  /** Applies a partial `<headerFooter>` update. */
  setSheetHeaderFooter(sheet: number, headerFooter: HeaderFooterInput): Status;
  /** Reads the effective page setup and which attributes state it. */
  getSheetPageSetup(sheet: number): SheetPageSetupResult;
  /** Reads the effective page margins, in inches. */
  getSheetPageMargins(sheet: number): SheetPageMarginsResult;

  /** Reads the full per-sheet view (zoom, freeze, tab-hidden, and the
   *  display / orientation flags). */
  getSheetView(sheet: number): SheetViewResult;
  /** Sets the sheet zoom percentage (clamped to `[10, 400]`). */
  setSheetZoom(sheet: number, zoomScale: number): Status;
  /** Sets the frozen pane in `(rows, cols)`. */
  setSheetFreeze(sheet: number, freezeRows: number, freezeCols: number): Status;
  /** Sets the sheet tab's hidden flag — the two-state view of
   *  `setSheetVisibility`. `true` on an already very-hidden sheet leaves
   *  it very-hidden rather than demoting it, since "hidden" says nothing
   *  that sheet does not already satisfy; `false` shows it from either
   *  hidden state. */
  setSheetTabHidden(sheet: number, hidden: boolean): Status;
  /** Sets the sheet tab's visibility to one of the three states. The only
   *  way to newly state `SheetVisibility.VeryHidden`, and the only way to
   *  demote a very-hidden sheet to plain hidden; `setSheetTabHidden` can
   *  express neither. A value outside `SheetVisibility` fails with
   *  `kInvalidArgument` and leaves the sheet unchanged. */
  setSheetVisibility(sheet: number, visibility: SheetVisibility): Status;
  /** Sets the sheet's `showGridLines` flag. */
  setSheetShowGridLines(sheet: number, show: boolean): Status;
  /** Sets the sheet's `showRowColHeaders` flag. */
  setSheetShowRowColHeaders(sheet: number, show: boolean): Status;
  /** Sets the sheet's `showZeros` flag. */
  setSheetShowZeros(sheet: number, show: boolean): Status;
  /** Sets the sheet's `rightToLeft` flag. */
  setSheetRightToLeft(sheet: number, rightToLeft: boolean): Status;
  /** Sets the sheet's `tabSelected` flag. */
  setSheetTabSelected(sheet: number, selected: boolean): Status;
  /** Sets the sheet's `<sheetView view="...">` mode: `""`,
   *  `"pageBreakPreview"`, or `"pageLayout"`. Stored verbatim. */
  setSheetViewMode(sheet: number, mode: string): Status;

  /**
   * Reads the sheet's `<sheetProtection>` flags. Strings are
   * deep-copied; the returned object is independent of the
   * workbook's storage.
   */
  getSheetProtection(sheet: number): SheetProtectionResult;
  /**
   * Replaces the sheet's `<sheetProtection>` flags wholesale.
   * Setting `enabled = 0` clears the protection block on save.
   */
  setSheetProtection(sheet: number, protection: SheetProtection): Status;

  /** Returns the column-layout overrides on `sheet` in storage order. */
  getSheetColumns(sheet: number): ColumnsResult;
  /** Sets / replaces the column width override on `[first, last]`. */
  setColumnWidth(sheet: number, first: number, last: number, width: number): Status;
  /** Sets / replaces the column hidden flag on `[first, last]`. */
  setColumnHidden(sheet: number, first: number, last: number, hidden: boolean): Status;
  /** Sets / replaces the column outline level on `[first, last]` (clamped to 0..255). */
  setColumnOutline(sheet: number, first: number, last: number, level: number): Status;

  /** Returns the row-layout overrides on `sheet`. */
  getSheetRowOverrides(sheet: number): RowsResult;
  /** Sets / replaces the row height override at `row`. */
  setRowHeight(sheet: number, row: number, height: number): Status;
  /** Sets / replaces the row hidden flag at `row`. */
  setRowHidden(sheet: number, row: number, hidden: boolean): Status;
  /** Sets / replaces the row outline level at `row` (clamped to 0..255). */
  setRowOutline(sheet: number, row: number, level: number): Status;

  /** Returns `{ status, xfIndex }` for the cell at `(sheet, row, col)`. */
  getCellXfIndex(sheet: number, row: number, col: number): CellXfIndexResult;
  /** Persists `xfIndex` on the cell at `(sheet, row, col)`. */
  setCellXfIndex(sheet: number, row: number, col: number, xfIndex: number): Status;
  /** Applies `xfIndex` across an inclusive rectangle, materialising blank
   *  cells so an empty ruled box renders. Saves one call per cell. */
  setRangeXfIndex(
    sheet: number,
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    xfIndex: number,
  ): Status;
  /** Returns the resolved XF record at `xfIndex`. */
  getCellXf(xfIndex: number): CellXfResult;
  /** Returns the font record at `fontIndex`, including its ColorSpec selector. */
  getFont(fontIndex: number): FontResult;
  /** Returns the fill record at `fillIndex`, including its ColorSpec selectors. */
  getFill(fillIndex: number): FillResult;
  /** Returns the border record at `borderIndex`, including its ColorSpec selectors. */
  getBorder(borderIndex: number): BorderResult;
  /** Returns the format string registered for `numFmtId`. */
  getNumFmt(numFmtId: number): NumFmtResult;
  /** Returns the differential-format record referenced by CF `dxfId`. */
  getDxf(dxfIndex: number): DxfResult;

  /** Adds a font (deduplicating against existing entries) and returns
   *  the resolved index. */
  addFont(record: FontRecord): AddStyleResult;
  /** Adds a fill (deduplicating against existing entries). */
  addFill(record: FillRecord): AddStyleResult;
  /** Adds a border (deduplicating against existing entries). */
  addBorder(record: BorderRecord): AddStyleResult;
  /** Adds a number-format code. An existing effective mapping is reused,
   *  including an existing custom record; a built-in id is reused without
   *  modifying the table only when its current effective mapping matches the
   *  code. Otherwise custom codes are appended at
   *  `max(existing_custom_id, 163) + 1`. Effective mappings use the first
   *  valid custom record for an id in document order, then a non-empty
   *  built-in code. */
  addNumFmt(formatCode: string): AddNumFmtResult;
  /** Adds an `<xf>` record (deduplicating against existing entries).
   *  Out-of-range font/fill/border indices or unregistered `numFmtId`
   *  surface `kInvalidArgument` rather than auto-growing the parallel
   *  tables. */
  addXf(record: CellXf): AddStyleResult;
  /** Adds a `<dxf>` differential-format record for conditional formats. */
  addDxf(record: DxfRecord): AddStyleResult;

  /** Returns the number of font records currently registered. */
  fontCount(): number;
  /** Returns the number of fill records currently registered. */
  fillCount(): number;
  /** Returns the number of border records currently registered. */
  borderCount(): number;
  /** Returns the number of `<xf>` records currently registered. */
  xfCount(): number;
  /** Returns the number of `<dxf>` records available for CF `dxfId`. */
  dxfCount(): number;

  /** Returns the number of named cell styles (`<cellStyle>` entries)
   *  registered. Zero for workbooks that do not declare any named
   *  styles. */
  cellStyleCount(): number;
  /** Returns the number of `<cellStyleXfs>` records — the named-style
   *  xf table referenced by `CellStyleResult.xfId`. Independent of the
   *  per-cell `cellXfs` table. */
  cellStyleXfCount(): number;
  /** Returns the named cell style at `index`. Out-of-range indices
   *  surface `kInvalidArgument` via `status`. */
  getCellStyle(index: number): CellStyleResult;
  /** Returns the named-style xf record at `index`. Output shape mirrors
   *  `getCellXf`. */
  getCellStyleXf(index: number): CellXfResult;
  /** Appends (deduplicating) a named-style xf and returns its
   *  `<cellStyleXfs>` index, which is what `setCellStyle` and
   *  `CellXf.xfId` reference. */
  addCellStyleXf(record: CellXf): AddStyleResult;
  /** Adds or replaces the `<cellStyle>` entry named `name`. `builtinId` is
   *  an OOXML ordinal `0..47`, or `0xFFFFFFFF` for a custom style; it is
   *  required because the binding takes a fixed argument count. */
  setCellStyle(name: string, xfId: number, builtinId: number): Status;

  /** Returns every external-link record carried by the workbook in
   *  `<externalReferences>` document order. Empty for fresh workbooks
   *  and any package whose source archive had no `<externalReferences>`
   *  block. */
  getExternalLinks(): ListResult<ExternalLinkRecord>;

  /** Adds a merge range to `sheet`. */
  addMerge(sheet: number, range: MergeRange): Status;
  /** Removes every merge that overlaps `range` (inclusive). No-op when nothing overlaps. */
  removeMerge(sheet: number, range: MergeRange): Status;
  /** Removes the merge at `index`. Returns kInvalidArgument if `index` is out of range. */
  removeMergeAt(sheet: number, index: number): Status;
  /** Drops every merge on `sheet`. */
  clearMerges(sheet: number): Status;
  /** Returns every merge range on `sheet` as a JS array. */
  getMerges(sheet: number): ListResult<MergeRange>;

  /** Returns the cell comment at `(sheet, row, col)`, or `null` when absent. */
  getComment(sheet: number, row: number, col: number): CommentEntry | null;
  /** Returns a comment lookup result that distinguishes absence from an invalid sheet or handle. */
  getCommentResult(sheet: number, row: number, col: number): CommentResult;
  /** Returns every comment on `sheet`, including comments anchored on
   *  cells that carry no value. */
  getComments(sheet: number): ListResult<SheetCommentEntry>;
  /** Sets / replaces the cell comment. Pass an empty `text` to remove. */
  setComment(sheet: number, row: number, col: number, author: string, text: string): Status;

  /** Appends a hyperlink to `sheet`. For an in-workbook link, pass an empty
   *  `target` and its A1 destination as `location`. */
  addHyperlink(
    sheet: number,
    row: number,
    col: number,
    target: string,
    display: string,
    tooltip: string,
    location: string,
  ): Status;
  /** Appends a hyperlink covering the inclusive rectangle from `(row, col)`
   *  through `(lastRow, lastCol)`. */
  addHyperlinkRange(
    sheet: number,
    row: number,
    col: number,
    lastRow: number,
    lastCol: number,
    target: string,
    display: string,
    tooltip: string,
    location: string,
  ): Status;
  /** Removes every hyperlink anchored at `(row, col)`. No-op when none match. */
  removeHyperlink(sheet: number, row: number, col: number): Status;
  /** Removes the hyperlink at `index`. Returns kInvalidArgument if `index` is out of range. */
  removeHyperlinkAt(sheet: number, index: number): Status;
  /** Drops every hyperlink on `sheet`. */
  clearHyperlinks(sheet: number): Status;
  /** Returns every hyperlink on `sheet` as a JS array. */
  getHyperlinks(sheet: number): ListResult<HyperlinkEntry>;

  /** Returns every data-validation rule on `sheet` in storage order.
   *  Each rule's `ranges` mirrors the OOXML `<dataValidation sqref=...>`
   *  cell-range list; the rule itself surfaces its raw OOXML payload
   *  (the engine does not yet evaluate validation rules). */
  getValidations(sheet: number): ListResult<DataValidationEntry>;
  /** Appends a data-validation rule to `sheet`. */
  addValidation(sheet: number, validation: DataValidationInput): Status;
  /** Removes the validation rule at `index`. Returns `kInvalidArgument`
   *  if `index` is out of range. */
  removeValidationAt(sheet: number, index: number): Status;
  /** Drops every validation rule on `sheet`. */
  clearValidations(sheet: number): Status;

  /** Returns every CF rule on `sheet` in flattened priority order. The
   *  returned entries borrow rule ids from the engine's storage; treat
   *  them as immutable view objects. */
  getConditionalFormats(sheet: number): ListResult<ConditionalFormatEntry>;
  /** Appends a new single-rule `<conditionalFormatting>` block to
   *  `sheet`, including visual rules when their payload object is supplied.
   *  `index` is the new rule's position in the sheet's flattened CF rule
   *  sequence (the same indexing `getConditionalFormats` and
   *  `removeConditionalFormatAt` use); it stays valid until a subsequent
   *  add/remove/clear mutation on the same sheet renumbers the sequence. */
  addConditionalFormat(sheet: number, rule: ConditionalFormatInput): AddStyleResult;
  /** Removes the CF rule at `index` (flattened order). When the
   *  containing block becomes empty it is removed too. */
  removeConditionalFormatAt(sheet: number, index: number): Status;
  /** Drops every CF block on `sheet`. */
  clearConditionalFormats(sheet: number): Status;

  /** Returns the cells that `(sheet, row, col)` directly reads
   *  (1-step precedents) when `depth <= 1`, or every cell reached
   *  within `depth` BFS steps otherwise. `depth` is capped at 32 to
   *  avoid runaway expansion in cyclic graphs. */
  precedents(sheet: number, row: number, col: number, depth: number): ReadonlyArray<CellNode>;
  /** Returns the cells that read `(sheet, row, col)` directly
   *  (1-step dependents). Same depth semantics as `precedents`. */
  dependents(sheet: number, row: number, col: number, depth: number): ReadonlyArray<CellNode>;

  /** Returns metadata for the function `name` (case-insensitive). When
   *  the function is unknown, returns `{ok: false}`. `locale` selects
   *  the catalog locale (`0` = `en-US`, `1` = `ja-JP`) and is validated,
   *  but does not change the result: the description / signature fields
   *  are always `undefined` (see {@link FunctionMetadataResult}). */
  functionMetadata(name: string, locale: number): FunctionMetadataResult;
  /** Returns every registered function's canonical name in ascending
   *  sort order. */
  functionNames(): ReadonlyArray<string>;

  /** Returns the localized display name for the canonical function
   *  `canonicalName` in `locale`. No alias table exists in any locale,
   *  so this always returns the canonical name unchanged; localized
   *  display names belong to the host's provider document. Returns the
   *  empty string when the canonical name does not match a registered
   *  function. */
  localizeFunctionName(canonicalName: string, locale: number): string;
  /** Inverse of `localizeFunctionName`: returns the canonical English
   *  name for the localized function `localizedName`. With no alias
   *  table this is a case-insensitive canonical-name match, in every
   *  locale. Returns the empty string when no function matches. */
  canonicalizeFunctionName(localizedName: string, locale: number): string;

  /** Returns dynamic-array spill info for `(sheet, row, col)`.
   *  When the cell is part of a spill region (anchor or phantom),
   *  `engaged` is `true` and `(anchorRow, anchorCol)` + `(rows, cols)`
   *  describe the region; the per-cell values are read via `getValue`,
   *  which is already spill-aware. When the cell is not part of any
   *  region, `engaged` is `false` and the other fields are zero. */
  spillInfo(sheet: number, row: number, col: number): SpillInfo;
}

/** Static factories on the Workbook class. */
export interface WorkbookCtor {
  /** Workbook with a single default sheet (`"Sheet1"`). */
  createDefault(): Workbook;
  /** Workbook with no sheets. */
  createEmpty(): Workbook;
  /** Loads from an in-memory `.xlsx` byte buffer. The returned wrapper
   *  may be invalid (`!isValid()`) on failure; consult
   *  `lastErrorMessage()` for diagnostics. */
  loadBytes(bytes: Uint8Array): Workbook;
}

/** Type of the resolved Module returned by the factory. */
export interface FormulonModule {
  Workbook: WorkbookCtor;

  /** Convenience: evaluates `formula` read-only against a fresh
   *  single-sheet workbook, anchored at `Sheet1!A1`. Nothing is written
   *  and no recalc runs, so an anchor-referencing formula such as `=A1`
   *  or `=COUNTA(A1)` reads the blank anchor rather than becoming a
   *  self-reference. An array result is reduced to its top-left element;
   *  use `Workbook.evaluateFormulaArray` for the whole grid. */
  evalFormula(formula: string): EvalResult;

  /** Library version string (UTF-8). */
  versionString(): string;

  /** Alias of {@link versionString}, matching the Node binding's name. */
  version(): string;

  /** Static description of `status` (e.g. `"kOk"`). */
  statusString(status: number): string;

  /** Excel display literal for a cell error code (e.g. `"#DIV/0!"`). */
  errorDisplayName(errorCode: number): string;

  /** Most-recent thread-local error message. */
  lastErrorMessage(): string;

  /** Most-recent thread-local error context. */
  lastErrorContext(): string;

  /** Sets the engine's minimum structured-log severity (a `LogLevel`
   *  ordinal). This is **process-wide** state, not per `Workbook`. The
   *  default is `LogLevel.Off`, under which the engine writes nothing to
   *  stderr and invokes no sink. */
  setLogMinLevel(level: number): Status;

  /** Routes structured-log records to `sink`, or restores the (silent at
   *  the default threshold) stderr fallback when `sink` is `null`.
   *  **Process-wide** state, not per `Workbook`.
   *
   *  `sink` receives one complete JSON record as raw bytes -- the C ABI
   *  hands the binding a length-delimited range, never a NUL-terminated
   *  string. The view is only valid for the duration of the call; copy it
   *  (`new Uint8Array(record)`) or decode it (`new TextDecoder().decode(
   *  record)`) before returning.
   *
   *  A sink that throws has its exception dropped and the record lost;
   *  the call that emitted it is unaffected. Records come from arbitrary
   *  depth inside the engine, so there is nowhere to report it. */
  setLogSink(sink: ((record: Uint8Array) => void) | null): Status;
}

/** Optional Emscripten module-init overrides. Pass to the factory to
 *  customise the default heap, stdout/stderr forwarding, or wasm
 *  binary resolution. */
export interface FormulonModuleOptions {
  locateFile?: (path: string, prefix: string) => string;
  print?: (msg: string) => void;
  printErr?: (msg: string) => void;
  noInitialRun?: boolean;
  noExitRuntime?: boolean;
}

/**
 * Default export from `formulon.js`: the Emscripten module factory.
 *
 * Usage (Node, ESM):
 * ```ts
 * import createFormulon from '@libraz/formulon';
 * const Module = await createFormulon();
 * const r = Module.evalFormula('=SUM(1,2,3)');
 * console.log(r.value.number);  // 6
 * ```
 */
export default function createFormulon(opts?: FormulonModuleOptions): Promise<FormulonModule>;
