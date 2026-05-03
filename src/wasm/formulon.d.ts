// Copyright 2026 libraz. Licensed under the MIT License.
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
  /** Thread-local last-error message (empty on success). */
  message: string;
  /** Optional thread-local context string (empty on success). */
  context: string;
}

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
export enum CfMatchKind {
  DifferentialFormat = 0,
  ColorScale = 1,
  DataBar = 2,
  IconSet = 3,
}

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

/** Iterable handle backing a `std::vector<CfMatch>` on the C++ side.
 *  Mirrors how embind surfaces `register_vector<T>`. */
export interface CfMatchVector {
  size(): number;
  get(index: number): CfMatch;
  delete(): void;
}

/** One cell's CF result inside a viewport-range evaluation. */
export interface CfCellResult {
  row: number;
  col: number;
  matches: CfMatchVector;
}

/** Iterable handle backing a `std::vector<CfCellResult>`. */
export interface CfCellVector {
  size(): number;
  get(index: number): CfCellResult;
  delete(): void;
}

/** Return type of `Workbook.evaluateCfRange(...)`. `cells` is sparse:
 *  only cells that produced at least one match appear. */
export interface CfRangeResult {
  status: Status;
  cells: CfCellVector;
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

/** Iterable handle backing a `std::vector<ColumnLayout>`. */
export interface ColumnLayoutVector {
  size(): number;
  get(index: number): ColumnLayout;
  delete(): void;
}

/** Return type of `Workbook.getSheetColumns(sheet)`. */
export interface ColumnsResult {
  status: Status;
  columns: ColumnLayoutVector;
}

/** Per-row layout override. */
export interface RowLayout {
  row: number;
  height: number;
  /** Boolean stored as 0/1. */
  hidden: number;
  outlineLevel: number;
}

/** Iterable handle backing a `std::vector<RowLayout>`. */
export interface RowLayoutVector {
  size(): number;
  get(index: number): RowLayout;
  delete(): void;
}

/** Return type of `Workbook.getSheetRowOverrides(sheet)`. */
export interface RowsResult {
  status: Status;
  rows: RowLayoutVector;
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
  /** Absolute or relative target (URL, email, internal ref, …). */
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

/** Sheet validation entry. The shape will firm up when the writeable
 *  surface lands; today the array is always empty. */
export interface ValidationEntry {
  ranges?: ReadonlyArray<MergeRange>;
  type?: string;
  op?: string;
  formula1?: string;
  formula2?: string;
  errorMessage?: string;
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

/** Workbook handle. Always release with `delete()` when finished. */
export interface Workbook {
  /** True when the wrapper holds a live native handle. */
  isValid(): boolean;
  /** Releases the native handle. The instance must not be used afterwards. */
  delete(): void;

  save(): SaveResult;
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
  setText(sheet: number, row: number, col: number, text: string): Status;
  setBlank(sheet: number, row: number, col: number): Status;
  setFormula(sheet: number, row: number, col: number, formula: string): Status;

  getValue(sheet: number, row: number, col: number): CellResult;

  recalc(): Status;
  /** Recalculates only cells touched by the supplied viewport. */
  partialRecalc(viewport: RecalcViewport): PartialRecalcResult;

  setIterative(enabled: boolean, maxIterations: number, maxChange: number): Status;
  /** Installs (or, when passed `null`, clears) a JS callback invoked
   *  after each Gauss-Seidel sweep. Only one callback can be active per
   *  WASM instance — installing a new one displaces the previous. */
  setIterativeProgress(callback: IterativeProgressCallback | null): Status;

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

  tableCount(): number;
  tableAt(idx: number): TableEntry;

  passthroughCount(): number;
  passthroughAt(idx: number): PassthroughEntry;

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

  /** Returns `{ status, xfIndex }` for the cell at `(sheet, row, col)`. */
  getCellXfIndex(sheet: number, row: number, col: number): CellXfIndexResult;
  /** Persists `xfIndex` on the cell at `(sheet, row, col)`. */
  setCellXfIndex(sheet: number, row: number, col: number, xfIndex: number): Status;
  /** Returns the resolved XF record at `xfIndex`. */
  getCellXf(xfIndex: number): CellXfResult;

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

  /** Returns every hyperlink on `sheet` as a JS array. */
  getHyperlinks(sheet: number): ReadonlyArray<HyperlinkEntry>;

  /** Returns every validation entry on `sheet`. Currently always empty:
   *  the writeable surface (and the underlying C ABI iterator) lands in
   *  a follow-up bundle. */
  getValidations(sheet: number): ReadonlyArray<ValidationEntry>;
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

  /** Convenience: evaluates a single formula in a fresh workbook
   *  (place at `Sheet1!A1`, recalc, return the cached value). */
  evalFormula(formula: string): EvalResult;

  /** Library version string (UTF-8). */
  versionString(): string;

  /** Static description of `status` (e.g. `"kOk"`). */
  statusString(status: number): string;

  /** Most-recent thread-local error message. */
  lastErrorMessage(): string;

  /** Most-recent thread-local error context. */
  lastErrorContext(): string;
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
