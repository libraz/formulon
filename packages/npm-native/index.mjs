// Copyright 2026 libraz. Licensed under the MIT License.
//
// ESM entry point for @libraz/formulon-native.
//
// This shim loads the compiled `.node` addon and re-exports the
// JS-friendly API. The `.node` module itself is a CommonJS binary
// loaded via `createRequire`, then surfaced to ESM consumers verbatim.
// The JS-visible shape is intentionally kept identical to
// `@libraz/formulon` (the WASM package) so consumers can swap binaries
// without code changes.

import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

const require_ = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));

// `formulon.node` is staged alongside this file in dist/.
const native = require_(path.join(here, 'formulon.node'));

export const Workbook = native.Workbook;
export const evalFormula = native.evalFormula;
export const version = native.version;
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

export default {
  Workbook,
  evalFormula,
  version,
  lastErrorMessage,
  lastErrorContext,
  statusString,
  ValueKind,
  CfMatchKind,
};
