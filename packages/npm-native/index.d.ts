// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// Hand-written TypeScript declarations for @libraz/formulon-native.
//
// This file is the public surface of the native N-API binding. It
// MUST be kept in sync with `src/node_addon/addon.cc` -- the addon's
// `Init()` registers the JS members declared here. Drift will surface
// at runtime as `undefined` access.
//
// The shape mirrors `packages/npm/dist/formulon.d.ts` (the WASM
// binding) so that JS callers can treat the two packages
// interchangeably for the methods both expose. The two notable shape
// differences are deliberate:
//
//   * Vector-of-value-object types (`CfMatchVector`, `CfCellVector`,
//     `ColumnLayoutVector`, `RowLayoutVector`) only exist in the embind
//     binding because embind cannot return `std::vector<T>` directly to
//     JS. The N-API binding marshals each list as a plain JS Array, so
//     `evaluateCfRange` / `getSheetColumns` / `getSheetRowOverrides`
//     return ReadonlyArray<...> here instead of an iterable handle.
//
//   * The native binding does NOT expose `delete()`. The handle is GC-
//     finalized; consumers hold the reference for the lifetime they
//     need the workbook.

/** `fm_value_kind_t` ordinals (mirror of `fm_value_kind_t`). */
export const ValueKind: Readonly<{
  Blank: 0;
  Number: 1;
  Bool: 2;
  Text: 3;
  Error: 4;
  Array: 5;
  Ref: 6;
  Lambda: 7;
}>;

/** Result envelope returned by every fallible binding call. */
export interface Status {
  /** True when the underlying C ABI returned `kOk`. */
  ok: boolean;
  /** Numeric `fm_status_t`. 0 on success. */
  status: number;
  /** Thread-local last-error message (empty on success). */
  message: string;
  /** Optional thread-local context string (empty on success). */
  context: string;
}

/** Flattened mirror of `fm_value_t`. Only the field selected by `kind`
 *  is meaningful; the others carry default-zero values. */
export interface Value {
  kind: number;
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

/** Return type of `Workbook.save()`. */
export interface SaveResult {
  status: Status;
  /** Freshly-allocated `Uint8Array` on success; `null` on failure. */
  bytes: Uint8Array | null;
}

/** Return type of `Workbook.sheetName(idx)`. */
export interface StringResult {
  status: Status;
  value: string;
}

/** Return type of `Workbook.cellAt(sheet, idx)`. */
export interface CellEntry {
  status: Status;
  row: number;
  col: number;
  /** Raw formula text, or `null` for pure literals. */
  formula: string | null;
  value: Value;
}

/** Return type of `Workbook.definedNameAt(idx)`. */
export interface DefinedNameEntry {
  status: Status;
  name: string;
  formula: string;
}

/** Return type of `Workbook.tableAt(idx)`. */
export interface TableEntry {
  status: Status;
  name: string;
  displayName: string;
  ref: string;
  sheetIndex: number;
}

/** Return type of `Workbook.passthroughAt(idx)`. */
export interface PassthroughEntry {
  status: Status;
  path: string;
}

/** Conditional-format match kind. Mirrors `formulon::cf::CFMatchKind`. */
export const CfMatchKind: Readonly<{
  DifferentialFormat: 0;
  ColorScale: 1;
  DataBar: 2;
  IconSet: 3;
}>;

/** RGBA colour. Channels are 0-255 (sRGB). */
export interface CfColor {
  r: number;
  g: number;
  b: number;
  a: number;
}

/** Resolved CF match. Active fields depend on `kind`; the others carry
 *  default-zero values. */
export interface CfMatch {
  kind: number;
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

/** Per-sheet view: zoom (10..400, default 100), frozen-pane row/col
 *  counts, and tab-hidden flag (0/1). */
export interface SheetView {
  zoomScale: number;
  freezeRows: number;
  freezeCols: number;
  /** Boolean stored as 0/1 to match the embind binding's wire shape. */
  tabHidden: number;
}

/** Return type of `Workbook.getSheetView(sheet)`. */
export interface SheetViewResult {
  status: Status;
  view: SheetView;
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
  /** Absolute or relative target (URL, email, internal ref, ...). */
  target: string;
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
 *  strings. */
export interface DataValidationInput {
  ranges?: ReadonlyArray<DataValidationRange>;
  type: number;
  op?: number;
  errorStyle?: number;
  allowBlank?: boolean;
  showInputMessage?: boolean;
  showErrorMessage?: boolean;
  formula1?: string;
  formula2?: string;
  errorTitle?: string;
  errorMessage?: string;
  promptTitle?: string;
  promptMessage?: string;
}

/** @deprecated Use {@link DataValidationEntry} instead. */
export type ValidationEntry = DataValidationEntry;

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
}

/** Plain-data shape of a font record. Mirrors `formulon::io::FontRecord`. */
export interface FontRecord {
  name: string;
  size: number;
  bold: boolean;
  italic: boolean;
  strike: boolean;
  /** 0=none, 1=single, 2=double, 3=singleAccounting, 4=doubleAccounting. */
  underline: number;
  /** AARRGGBB packed colour. */
  colorArgb: number;
}

/** Plain-data shape of a fill record. */
export interface FillRecord {
  /** OOXML pattern ordinal: 0=none, 1=solid, 2..18=standard pattern set. */
  pattern: number;
  /** Foreground AARRGGBB colour. */
  fgArgb: number;
  /** Background AARRGGBB colour. */
  bgArgb: number;
}

/** One side of a `BorderRecord`. */
export interface BorderSide {
  style: number;
  colorArgb: number;
}

/** Plain-data shape of a border record. */
export interface BorderRecord {
  left: BorderSide;
  right: BorderSide;
  top: BorderSide;
  bottom: BorderSide;
  diagonal: BorderSide;
  diagonalUp: boolean;
  diagonalDown: boolean;
}

/** Plain-data shape of an `<xf>` record. */
export interface CellXf {
  fontIndex: number;
  fillIndex: number;
  borderIndex: number;
  numFmtId: number;
  horizontalAlign: number;
  verticalAlign: number;
  wrapText: boolean;
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

/** Return type of `Workbook.addFont/Fill/Border/Xf(...)`. The
 *  add-functions deduplicate against existing entries via linear
 *  search; `index` is either the matched index or the freshly-appended
 *  index. */
export interface AddStyleResult {
  status: Status;
  index: number;
}

/** Return type of `Workbook.addNumFmt(formatCode)`. The id is either a
 *  matched built-in (`0..163`) or a freshly-assigned custom id (`>= 164`). */
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

/** Iterative-solver progress callback. Receives the current
 *  iteration number, the maximum residual seen, and the configured
 *  iteration cap. Returning `false` (or any falsy value) aborts the
 *  solve; returning `true` (or `undefined`) continues. */
export type IterativeProgressCallback = (
  iteration: number,
  maxResidual: number,
  maxIterations: number,
) => boolean | undefined | void;

/** Workbook handle. The wrapper is GC-finalized; it does NOT expose
 *  an explicit `delete()` step. Hold the reference for the lifetime
 *  you need the workbook. */
export interface Workbook {
  /** True when the wrapper holds a live native handle. False when the
   *  underlying handle has been finalised or could not be constructed. */
  isValid(): boolean;

  // Cell mutation.
  setNumber(sheet: number, row: number, col: number, value: number): Status;
  setBool(sheet: number, row: number, col: number, value: boolean): Status;
  setText(sheet: number, row: number, col: number, text: string): Status;
  setBlank(sheet: number, row: number, col: number): Status;
  setFormula(sheet: number, row: number, col: number, formula: string): Status;

  // Cell read.
  getValue(sheet: number, row: number, col: number): CellResult;

  // Recalc + save.
  recalc(): Status;
  /** Recalculates only cells touched by the supplied viewport. */
  partialRecalc(viewport: RecalcViewport): PartialRecalcResult;
  setIterative(enabled: boolean, maxIterations: number, maxChange: number): Status;
  /** Installs (or, when passed `null`, clears) a JS callback invoked
   *  after each Gauss-Seidel sweep. Only one callback can be active per
   *  addon instance -- installing a new one displaces the previous. */
  setIterativeProgress(callback: IterativeProgressCallback | null): Status;
  save(): SaveResult;

  // Sheet operations.
  addSheet(name: string): Status;
  removeSheet(index: number): Status;
  renameSheet(index: number, name: string): Status;
  /** Moves the sheet from `fromIdx` to `toIdx` (post-removal index). */
  moveSheet(fromIdx: number, toIdx: number): Status;
  sheetCount(): number;
  sheetName(index: number): StringResult;

  // Row / column structural edits.
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

  // Iteration / metadata.
  cellCount(sheet: number): number;
  cellAt(sheet: number, idx: number): CellEntry;
  definedNameCount(): number;
  definedNameAt(idx: number): DefinedNameEntry;
  tableCount(): number;
  tableAt(idx: number): TableEntry;
  passthroughCount(): number;
  passthroughAt(idx: number): PassthroughEntry;

  // Defined names.
  setDefinedName(name: string, formula: string): Status;

  // Conditional formatting.
  /** Evaluates every CF block on `sheet` against the inclusive range
   *  `[(firstRow, firstCol), (lastRow, lastCol)]`. Pass `NaN` for
   *  `todaySerial` to disable `TimePeriod` rules. */
  evaluateCfRange(
    sheet: number,
    firstRow: number,
    firstCol: number,
    lastRow: number,
    lastCol: number,
    todaySerial: number,
  ): CfRangeResult;

  // Sheet view / layout.
  /** Reads the per-sheet view (zoom, freeze, tab-hidden). */
  getSheetView(sheet: number): SheetViewResult;
  /** Sets the sheet zoom percentage (clamped to `[10, 400]`). */
  setSheetZoom(sheet: number, zoomScale: number): Status;
  /** Sets the frozen pane in `(rows, cols)`. */
  setSheetFreeze(sheet: number, freezeRows: number, freezeCols: number): Status;
  /** Sets the sheet tab's hidden flag. */
  setSheetTabHidden(sheet: number, hidden: boolean): Status;

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

  // Styles.
  /** Returns `{ status, xfIndex }` for the cell at `(sheet, row, col)`. */
  getCellXfIndex(sheet: number, row: number, col: number): CellXfIndexResult;
  /** Persists `xfIndex` on the cell at `(sheet, row, col)`. */
  setCellXfIndex(sheet: number, row: number, col: number, xfIndex: number): Status;
  /** Returns the resolved XF record at `xfIndex`. */
  getCellXf(xfIndex: number): CellXfResult;
  /** Returns the resolved font record at `fontIndex`. */
  getFont(fontIndex: number): FontResult;
  /** Returns the resolved fill record at `fillIndex`. */
  getFill(fillIndex: number): FillResult;
  /** Returns the resolved border record at `borderIndex`. */
  getBorder(borderIndex: number): BorderResult;
  /** Returns the format string registered for `numFmtId`. */
  getNumFmt(numFmtId: number): NumFmtResult;

  /** Adds a font (deduplicating against existing entries) and returns
   *  the resolved index. */
  addFont(record: FontRecord): AddStyleResult;
  /** Adds a fill (deduplicating against existing entries). */
  addFill(record: FillRecord): AddStyleResult;
  /** Adds a border (deduplicating against existing entries). */
  addBorder(record: BorderRecord): AddStyleResult;
  /** Adds a number-format code. Built-in matches return the built-in id
   *  without modifying the table; custom codes are appended at
   *  `max(existing_custom_id, 163) + 1`. */
  addNumFmt(formatCode: string): AddNumFmtResult;
  /** Adds an `<xf>` record (deduplicating against existing entries).
   *  Out-of-range font/fill/border indices or unregistered `numFmtId`
   *  surface `kInvalidArgument` rather than auto-growing the parallel
   *  tables. */
  addXf(record: CellXf): AddStyleResult;

  /** Returns the number of font records currently registered. */
  fontCount(): number;
  /** Returns the number of fill records currently registered. */
  fillCount(): number;
  /** Returns the number of border records currently registered. */
  borderCount(): number;
  /** Returns the number of `<xf>` records currently registered. */
  xfCount(): number;

  // Sheet UI features (merges, comments, hyperlinks, validations).
  /** Adds a merge range to `sheet`. */
  addMerge(sheet: number, range: MergeRange): Status;
  /** Removes every merge that overlaps `range` (inclusive). No-op when nothing overlaps. */
  removeMerge(sheet: number, range: MergeRange): Status;
  /** Removes the merge at `index`. Returns kInvalidArgument if `index` is out of range. */
  removeMergeAt(sheet: number, index: number): Status;
  /** Drops every merge on `sheet`. */
  clearMerges(sheet: number): Status;
  /** Returns every merge range on `sheet` as a JS array. */
  getMerges(sheet: number): ReadonlyArray<MergeRange>;
  /** Returns the cell comment at `(sheet, row, col)`, or `null` when absent. */
  getComment(sheet: number, row: number, col: number): CommentEntry | null;
  /** Sets / replaces the cell comment. Pass an empty `text` to remove. */
  setComment(sheet: number, row: number, col: number, author: string, text: string): Status;
  /** Appends a hyperlink to `sheet`. Pass empty strings for `display`
   *  or `tooltip` to mean "use the default" or "no tooltip". The
   *  `location` field is filled implicitly (empty) and the writer mints
   *  a fresh `rId` on save. */
  addHyperlink(sheet: number, row: number, col: number, target: string, display: string, tooltip: string): Status;
  /** Removes every hyperlink anchored at `(row, col)`. No-op when none match. */
  removeHyperlink(sheet: number, row: number, col: number): Status;
  /** Removes the hyperlink at `index`. Returns kInvalidArgument if `index` is out of range. */
  removeHyperlinkAt(sheet: number, index: number): Status;
  /** Drops every hyperlink on `sheet`. */
  clearHyperlinks(sheet: number): Status;
  /** Returns every hyperlink on `sheet` as a JS array. */
  getHyperlinks(sheet: number): ReadonlyArray<HyperlinkEntry>;
  /** Returns every data-validation rule on `sheet` in storage order. */
  getValidations(sheet: number): ReadonlyArray<DataValidationEntry>;
  /** Appends a data-validation rule to `sheet`. */
  addValidation(sheet: number, validation: DataValidationInput): Status;
  /** Removes the validation rule at `index`. Returns `kInvalidArgument`
   *  if `index` is out of range. */
  removeValidationAt(sheet: number, index: number): Status;
  /** Drops every validation rule on `sheet`. */
  clearValidations(sheet: number): Status;
}

/** Static factories on the Workbook class. */
export interface WorkbookCtor {
  /** Workbook with a single default sheet (`"Sheet1"`). */
  createDefault(): Workbook;
  /** Workbook with no sheets. */
  createEmpty(): Workbook;
  /** Loads from an in-memory `.xlsx` byte buffer. The returned wrapper
   *  may be unusable (subsequent calls return `kBindingNullPointer`)
   *  on failure; consult `lastErrorMessage()` for diagnostics. */
  loadBytes(bytes: Uint8Array): Workbook;
}

export const Workbook: WorkbookCtor;

/** Convenience: evaluates a single formula in a fresh workbook
 *  (place at `Sheet1!A1`, recalc, return the cached value). */
export function evalFormula(formula: string): EvalResult;

/** Library version string (UTF-8). */
export function version(): string;

/** Most-recent thread-local error message. */
export function lastErrorMessage(): string;

/** Most-recent thread-local error context. */
export function lastErrorContext(): string;

/** Static description of `status` (e.g. `"kOk"`). */
export function statusString(status: number): string;

declare const _default: {
  Workbook: WorkbookCtor;
  evalFormula: typeof evalFormula;
  version: typeof version;
  lastErrorMessage: typeof lastErrorMessage;
  lastErrorContext: typeof lastErrorContext;
  statusString: typeof statusString;
  ValueKind: typeof ValueKind;
  CfMatchKind: typeof CfMatchKind;
};
export default _default;
