// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
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
 *  `description` are populated only when the locale metadata table
 *  has an entry for this function and locale. */
export interface FunctionMetadataResult {
  readonly ok: boolean;
  readonly name?: string;
  readonly minArity?: number;
  /** `0xFFFFFFFF` (i.e. `4294967295`) denotes an unbounded variadic. */
  readonly maxArity?: number;
  readonly signatureTemplate?: string;
  readonly description?: string;
}

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
 *  Visual rule kinds (`colorScale` / `dataBar` / `iconSet`) round-trip
 *  through the OOXML reader / writer but their visual sub-spec fields
 *  are not yet surfaced through this read API; only `id`, `type`,
 *  `priority`, `stopIfTrue`, and `sqref` populate for those kinds.
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
}

/** Argument shape accepted by `addConditionalFormat(sheet, rule)`.
 *
 *  Visual rule kinds (`colorScale` / `dataBar` / `iconSet`) are
 *  rejected — their visual sub-specs are not yet creatable through this
 *  API. The OOXML reader / writer continue to round-trip those rules
 *  verbatim if they were authored elsewhere.
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

/** Plain-data shape of a fill record. Mirrors `formulon::io::FillRecord`. */
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
  /** OOXML border-style ordinal: 0=none, 1=thin, ..., 13=slantDashDot. */
  style: number;
  /** AARRGGBB packed colour. */
  colorArgb: number;
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

/** External-link kinds. Mirrors
 *  `formulon::io::ExternalLinkRecord::Kind`. */
export const enum ExternalLinkKind {
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

  /** Renders the lambda value stored at `(sheet, row, col)` as Excel
   *  formula text. Returns `kInvalidArgument` when the cell is absent
   *  or its cached value is not a lambda. */
  getLambdaText(sheet: number, row: number, col: number): LambdaTextResult;

  recalc(): Status;
  /** Recalculates only cells touched by the supplied viewport. */
  partialRecalc(viewport: RecalcViewport): PartialRecalcResult;

  setIterative(enabled: boolean, maxIterations: number, maxChange: number): Status;
  /** Installs (or, when passed `null`, clears) a JS callback invoked
   *  after each Gauss-Seidel sweep. Only one callback can be active per
   *  WASM instance — installing a new one displaces the previous. */
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

  /** Returns every external-link record carried by the workbook in
   *  `<externalReferences>` document order. Empty for fresh workbooks
   *  and any package whose source archive had no `<externalReferences>`
   *  block. */
  getExternalLinks(): ReadonlyArray<ExternalLinkRecord>;

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

  /** Returns every data-validation rule on `sheet` in storage order.
   *  Each rule's `ranges` mirrors the OOXML `<dataValidation sqref=...>`
   *  cell-range list; the rule itself surfaces its raw OOXML payload
   *  (the engine does not yet evaluate validation rules). */
  getValidations(sheet: number): ReadonlyArray<DataValidationEntry>;
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
  getConditionalFormats(sheet: number): ReadonlyArray<ConditionalFormatEntry>;
  /** Appends a new single-rule `<conditionalFormatting>` block to
   *  `sheet`. Visual rule types (`colorScale` / `dataBar` / `iconSet`)
   *  are rejected with `kInvalidArgument` — those payloads are
   *  preserved on round-trip but not yet creatable here. */
  addConditionalFormat(sheet: number, rule: ConditionalFormatInput): Status;
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
   *  the catalog locale (`0` = `en-US`, `1` = `ja-JP`); description /
   *  signature fields are populated only when the locale metadata table
   *  has an entry. */
  functionMetadata(name: string, locale: number): FunctionMetadataResult;
  /** Returns every registered function's canonical name in ascending
   *  sort order. */
  functionNames(): ReadonlyArray<string>;

  /** Returns the localized display name for the canonical function
   *  `canonicalName` in `locale`. Returns the canonical name unchanged
   *  when the locale's alias table is empty (currently always for
   *  non-`en-US` locales). Returns the empty string when the canonical
   *  name does not match a registered function. */
  localizeFunctionName(canonicalName: string, locale: number): string;
  /** Inverse of `localizeFunctionName`: returns the canonical English
   *  name for the localized function `localizedName`. Falls through to
   *  case-insensitive canonical-name matching when no alias is
   *  registered. Returns the empty string when no function matches. */
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
