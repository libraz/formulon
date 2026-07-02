// Copyright 2026 libraz. Licensed under the Apache License, Version 2.0.
//
// ESM entry point for @libraz/formulon-native.
//
// This shim loads the compiled `.node` addon and re-exports the
// JS-friendly API. The `.node` module itself is a CommonJS binary
// loaded via `createRequire`, then surfaced to ESM consumers verbatim.
// The JS-visible shape is intentionally kept identical to
// `@libraz/formulon` (the WASM package) so consumers can swap binaries
// without code changes.
//
// Binary lookup order at require-time:
//   1. dist/prebuilds/<platform>-<arch>/formulon.node  (shipped artifacts)
//   2. dist/formulon.node                              (local dev fallback)

import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { existsSync } from 'node:fs';
import path from 'node:path';

const require_ = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));

const prebuildSlot = path.join(here, 'prebuilds', `${process.platform}-${process.arch}`, 'formulon.node');
const fallbackSlot = path.join(here, 'formulon.node');

let nativePath;
if (existsSync(prebuildSlot)) {
  nativePath = prebuildSlot;
} else if (existsSync(fallbackSlot)) {
  nativePath = fallbackSlot;
} else {
  throw new Error(
    `@libraz/formulon-native: no prebuild for ${process.platform}-${process.arch}. ` +
      `Looked in:\n  ${prebuildSlot}\n  ${fallbackSlot}\n` +
      'See https://github.com/libraz/formulon for supported platforms.',
  );
}

const native = require_(nativePath);

export const Workbook = native.Workbook;
export const evalFormula = native.evalFormula;
export const version = native.version;
/** Alias of {@link version}, matching the WASM binding's name. */
export const versionString = native.versionString;
export const lastErrorMessage = native.lastErrorMessage;
export const lastErrorContext = native.lastErrorContext;
export const statusString = native.statusString;

/** `fm_value_kind_t` ordinals (mirror of `fm_value_kind_t`). */
export const ValueKind = Object.freeze({
  Blank: 0,
  Number: 1,
  Bool: 2,
  Text: 3,
  Error: 4,
  Array: 5,
  Ref: 6,
  Lambda: 7,
});

/** `fm_cf_match_kind_t` ordinals (mirror of `formulon::cf::CFMatchKind`). */
export const CfMatchKind = Object.freeze({
  DifferentialFormat: 0,
  ColorScale: 1,
  DataBar: 2,
  IconSet: 3,
});

/** `fm_pivot_cell_kind_t` ordinals. */
export const PivotCellKind = Object.freeze({
  Header: 0,
  RowLabel: 1,
  ColLabel: 2,
  Data: 3,
  RowSubtotal: 4,
  ColSubtotal: 5,
  GrandTotal: 6,
  Blank: 7,
});

/** `fm_workbook_format_t` ordinals: container format for `saveEx`. */
export const WorkbookFormat = Object.freeze({
  Unknown: 0,
  Xlsx: 1,
  Xlsb: 2,
});

/** `fm_error_code_t` ordinals (mirror of `formulon::ErrorCode`). */
export const ErrorCode = Object.freeze({
  Null: 0,
  Div0: 1,
  Value: 2,
  Ref: 3,
  Name: 4,
  Num: 5,
  NA: 6,
  GettingData: 7,
  Spill: 8,
  Calc: 9,
  Field: 10,
  Blocked: 11,
  Connect: 12,
  External: 13,
  Busy: 14,
  Python: 15,
  Unknown: 16,
});

export default {
  Workbook,
  evalFormula,
  version,
  versionString,
  lastErrorMessage,
  lastErrorContext,
  statusString,
  ValueKind,
  CfMatchKind,
  PivotCellKind,
  ErrorCode,
  WorkbookFormat,
};
