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

/**
 * `fm_pivot_date_grouping_t` ordinals. `Days` is Excel's "By: Days"
 * grouping (an explicit interval, e.g. a real Excel "week" is `Days`
 * with `intervalDays: 7`); there is no separate week grouping.
 */
export const PivotDateGrouping = Object.freeze({
  Day: 0,
  Month: 1,
  Quarter: 2,
  Year: 3,
  Days: 4,
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

/** `fm_sheet_visibility_t` ordinals (mirror of OOXML `<sheet state>`). */
export const SheetVisibility = Object.freeze({ Visible: 0, Hidden: 1, VeryHidden: 2 });

/** `fm_geometry_mode_t` ordinals: which column-width figure a geometry call uses. */
export const GeometryMode = Object.freeze({ Display: 0, Print: 1 });

/** `fm_display_status_t` ordinals for `getDisplayText` / `formatValue`. */
export const DisplayStatus = Object.freeze({ Ok: 0, Overflow: 1, InvalidFormat: 2 });

/** `fm_filter_kind_t` ordinals: criterion of an AutoFilter column. */
export const FilterKind = Object.freeze({ None: 0, Values: 1, Custom: 2, Top10: 3, Dynamic: 4, Color: 5, Icon: 6 });

/** `fm_filter_operator_t` ordinals: comparison operator of a custom AutoFilter condition. */
export const FilterOperator = Object.freeze({
  Equal: 0,
  LessThan: 1,
  LessThanOrEqual: 2,
  NotEqual: 3,
  GreaterThanOrEqual: 4,
  GreaterThan: 5,
});

/** `fm_dynamic_filter_type_t` ordinals: dynamic AutoFilter type, in `ST_DynamicFilterType` schema order. */
export const DynamicFilterType = Object.freeze({
  Null: 0,
  AboveAverage: 1,
  BelowAverage: 2,
  Tomorrow: 3,
  Today: 4,
  Yesterday: 5,
  NextWeek: 6,
  ThisWeek: 7,
  LastWeek: 8,
  NextMonth: 9,
  ThisMonth: 10,
  LastMonth: 11,
  NextQuarter: 12,
  ThisQuarter: 13,
  LastQuarter: 14,
  NextYear: 15,
  ThisYear: 16,
  LastYear: 17,
  YearToDate: 18,
  Q1: 19,
  Q2: 20,
  Q3: 21,
  Q4: 22,
  M1: 23,
  M2: 24,
  M3: 25,
  M4: 26,
  M5: 27,
  M6: 28,
  M7: 29,
  M8: 30,
  M9: 31,
  M10: 32,
  M11: 33,
  M12: 34,
});

/** `fm_sort_by_t` ordinals: what a sort condition sorts by. */
export const SortBy = Object.freeze({ Value: 0, CellColor: 1, FontColor: 2, Icon: 3 });

/** `fm_sort_method_t` ordinals: sort method of an AutoFilter sort state. */
export const SortMethod = Object.freeze({ None: 0, PinYin: 1, Stroke: 2 });

/** `fm_date_time_grouping_t` ordinals: granularity of an AutoFilter date-group item. */
export const DateTimeGrouping = Object.freeze({ Year: 0, Month: 1, Day: 2, Hour: 3, Minute: 4, Second: 5 });

/** `fm_validation_error_style_t` ordinals: data-validation error style. */
export const ValidationErrorStyle = Object.freeze({ Stop: 0, Warning: 1, Information: 2 });

/** `fm_color_context_t` ordinals: where a colour is used. */
export const ColorContext = Object.freeze({ Font: 0, FillForeground: 1, FillBackground: 2, Border: 3 });

/** `fm_color_resolution_t` ordinals: how a colour was resolved. */
export const ColorResolution = Object.freeze({
  Exact: 0,
  DefaultTheme: 1,
  IndexOutOfRange: 2,
  ThemeUnparseable: 3,
  AutoContext: 4,
});

/** Where `getTheme` read the theme from (the C `out_source` values). */
export const ThemeSource = Object.freeze({ Part: 0, Default: 1, Unparseable: 2 });

/** Where `getEffectiveStyle` took its xf from. */
export const EffectiveStyleSource = Object.freeze({ Cell: 0, Row: 1, Column: 2, Default: 3 });

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
  SheetVisibility,
  GeometryMode,
  DisplayStatus,
  ColorContext,
  FilterKind,
  FilterOperator,
  DynamicFilterType,
  SortBy,
  SortMethod,
  DateTimeGrouping,
  ValidationErrorStyle,
  ColorResolution,
  ThemeSource,
  EffectiveStyleSource,
  ExternalLinkKind,
  ErrorCode,
  WorkbookFormat,
};
