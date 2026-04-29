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
export const enum ValueKind {
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

/** Workbook handle. Always release with `delete()` when finished. */
export interface Workbook {
  /** True when the wrapper holds a live native handle. */
  isValid(): boolean;
  /** Releases the native handle. The instance must not be used afterwards. */
  delete(): void;

  save(): SaveResult;
  addSheet(name: string): Status;
  sheetCount(): number;
  sheetName(idx: number): StringResult;

  setNumber(sheet: number, row: number, col: number, value: number): Status;
  setBool(sheet: number, row: number, col: number, value: boolean): Status;
  setText(sheet: number, row: number, col: number, text: string): Status;
  setBlank(sheet: number, row: number, col: number): Status;
  setFormula(sheet: number, row: number, col: number, formula: string): Status;

  getValue(sheet: number, row: number, col: number): CellResult;

  recalc(): Status;
  setIterative(enabled: boolean, maxIterations: number, maxChange: number): Status;

  cellCount(sheet: number): number;
  cellAt(sheet: number, idx: number): CellEntry;

  definedNameCount(): number;
  definedNameAt(idx: number): DefinedNameEntry;

  tableCount(): number;
  tableAt(idx: number): TableEntry;

  passthroughCount(): number;
  passthroughAt(idx: number): PassthroughEntry;
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
