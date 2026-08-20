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

import { existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

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
export const errorDisplayName = native.errorDisplayName;
export const setLogMinLevel = native.setLogMinLevel;
export const setLogSink = native.setLogSink;

/**
 * Merge a host-supplied metadata entry over the engine's structural
 * `functionMetadata()` result.
 *
 * This is a pure, side-effect-free helper: it does not touch the engine.
 * The engine returns `signatureTemplate` / `description` as
 * `undefined`; a host injects display metadata (see
 * `docs/function-metadata-schema.md`) and merges it here at display time.
 * The metadata is display-only and never affects formula parsing or
 * evaluation.
 *
 * Field precedence (first non-nullish wins):
 *   - signatureTemplate: `entry.localized[locale].signature` ->
 *     `entry.signature` -> `base.signatureTemplate`
 *   - description: `entry.localized[locale].description` ->
 *     `entry.description` -> `base.description`
 *   - localizedName: `entry.aliases[locale]` -> `base.name`
 *
 * @param {object} base A `FunctionMetadataResult` from `functionMetadata()`.
 * @param {object|undefined} entry The provider's `functions[NAME]` entry, or
 *   `undefined`/`null` to leave `base` unchanged (signature/description
 *   stay `undefined`).
 * @param {string} locale A BCP-47 display locale tag (e.g. `"fr-FR"`),
 *   matching the keys in `aliases` / `localized`. Independent of the numeric
 *   locale code passed to `functionMetadata()`.
 * @returns {object} The merged metadata (a new object), or `base` verbatim
 *   when `entry` is absent.
 */
export function mergeFunctionMetadata(base, entry, locale) {
  if (entry === undefined || entry === null) {
    return base;
  }
  const localized = (entry.localized && entry.localized[locale]) || {};
  const signatureTemplate = localized.signature ?? entry.signature ?? base.signatureTemplate;
  const description = localized.description ?? entry.description ?? base.description;
  const aliasName = entry.aliases && entry.aliases[locale];
  const localizedName = aliasName ?? base.name;
  return { ...base, signatureTemplate, description, localizedName };
}

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

/** `fm_pivot_axis_t` ordinals. */
export const PivotAxis = Object.freeze({ Row: 0, Col: 1, Value: 2, Page: 3 });

/** `fm_pivot_aggregation_t` ordinals. */
export const PivotAggregation = Object.freeze({
  Sum: 0,
  Count: 1,
  Average: 2,
  Max: 3,
  Min: 4,
  Product: 5,
  CountNumbers: 6,
  StdDev: 7,
  StdDevP: 8,
  Var: 9,
  VarP: 10,
});

/** `fm_pivot_show_as_t` ordinals. */
export const PivotShowValuesAs = Object.freeze({
  Normal: 0,
  PercentOfRow: 1,
  PercentOfCol: 2,
  PercentOfTotal: 3,
  RunningTotalInRow: 4,
  RunningTotalInCol: 5,
  Index: 6,
  DifferenceFrom: 7,
  PercentDifferenceFrom: 8,
  PercentOfParentRow: 9,
  PercentOfParentCol: 10,
  PercentOfParent: 11,
});

export const PIVOT_SHOW_AS_BASE_PREVIOUS = 1048828;
export const PIVOT_SHOW_AS_BASE_NEXT = 1048829;

/** `fm_pivot_filter_type_t` ordinals. */
export const PivotFilterType = Object.freeze({
  ValueTop10: 0,
  ValueGreaterThan: 1,
  ValueBetween: 2,
  LabelContains: 3,
  LabelBeginsWith: 4,
  LabelDate: 5,
});

/** `fm_pivot_date_grouping_t` ordinals. */
export const PivotDateGrouping = Object.freeze({
  Day: 0,
  Month: 1,
  Quarter: 2,
  Year: 3,
  Week: 4,
  Hour: 5,
  Minute: 6,
  Second: 7,
});

/** `fm_pivot_calendar_t` ordinals. */
export const PivotCalendar = Object.freeze({ Gregorian: 0, Japanese: 1 });

/** `fm_pivot_filter_value_kind_t` ordinals. */
export const PivotFilterValueKind = Object.freeze({ None: -1, Int: 0, Double: 1, Text: 2 });

/** `fm_pivot_layout_t` ordinals. */
export const PivotReportLayout = Object.freeze({ Compact: 0, Tabular: 1, Outline: 2 });

/** `fm_workbook_format_t` ordinals: container format for `saveAs`. */
export const WorkbookFormat = Object.freeze({
  Unknown: 0,
  Xlsx: 1,
  Xlsb: 2,
});

/** `fm_calc_mode_t` ordinals (mirror of `formulon::io::CalcMode`). */
export const CalcMode = Object.freeze({ Auto: 0, Manual: 1, AutoNoTable: 2 });

/** External-link kinds (mirror of `formulon::io::ExternalLinkRecord::Kind`). */
export const ExternalLinkKind = Object.freeze({ Unknown: 0, ExternalBook: 1, Ole: 2, Dde: 3 });

/** `fm_log_level_t` ordinals for `setLogMinLevel`. `Off` is the default. */
export const LogLevel = Object.freeze({ Debug: 0, Info: 1, Warn: 2, Error: 3, Off: 4 });

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
  errorDisplayName,
  setLogMinLevel,
  setLogSink,
  mergeFunctionMetadata,
  ValueKind,
  CfMatchKind,
  PivotCellKind,
  PivotAxis,
  PivotAggregation,
  PivotShowValuesAs,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  PIVOT_SHOW_AS_BASE_NEXT,
  PivotFilterType,
  PivotDateGrouping,
  PivotCalendar,
  PivotFilterValueKind,
  PivotReportLayout,
  LogLevel,
  CalcMode,
  ExternalLinkKind,
  ErrorCode,
  WorkbookFormat,
};
