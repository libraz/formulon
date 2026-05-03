// Copyright 2026 libraz. Licensed under the MIT License.
//
// Hand-written TypeScript declarations for @libraz/formulon-native.
//
// This file is the public surface of the native N-API binding. It
// MUST be kept in sync with `src/node_addon/addon.cc` — the addon's
// `Init()` registers the JS members declared here. Drift will surface
// at runtime as `undefined` access.
//
// The shape mirrors `packages/npm/dist/formulon.d.ts` (the WASM
// binding) so that JS callers can treat the two packages
// interchangeably for the methods both expose.

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

/** Workbook handle. The wrapper is GC-finalized; it does NOT expose
 *  an explicit `delete()` step. Hold the reference for the lifetime
 *  you need the workbook. */
export interface Workbook {
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
  save(): SaveResult;

  // Sheet operations.
  addSheet(name: string): Status;
  removeSheet(index: number): Status;
  renameSheet(index: number, name: string): Status;
  sheetCount(): number;
  sheetName(index: number): StringResult;

  // Defined names.
  setDefinedName(name: string, formula: string): Status;
}

/** Static factories on the Workbook class. */
export interface WorkbookCtor {
  /** Workbook with a single default sheet (`"Sheet1"`). */
  createDefault(): Workbook;
  /** Workbook with no sheets. */
  createEmpty(): Workbook;
  /** Loads from an in-memory `.xlsx` byte buffer. */
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
};
export default _default;
