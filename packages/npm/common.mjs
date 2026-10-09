// Value constants and pure helpers declared in formulon.d.ts, shared by the
// `@libraz/formulon` and `@libraz/formulon/threads` entry shims.

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
export const WorkbookFormat = Object.freeze({ Unknown: 0, Xlsx: 1, Xlsb: 2 });
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
export const PivotAxis = Object.freeze({ Row: 0, Col: 1, Value: 2, Page: 3 });
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
export const PivotFilterType = Object.freeze({
  ValueTop10: 0,
  ValueGreaterThan: 1,
  ValueBetween: 2,
  LabelContains: 3,
  LabelBeginsWith: 4,
  LabelDate: 5,
});
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
export const PivotCalendar = Object.freeze({ Gregorian: 0, Japanese: 1 });
export const PivotReportLayout = Object.freeze({ Compact: 0, Tabular: 1, Outline: 2 });
export const PivotFilterValueKind = Object.freeze({ None: -1, Int: 0, Double: 1, Text: 2 });
export const CfMatchKind = Object.freeze({ DifferentialFormat: 0, ColorScale: 1, DataBar: 2, IconSet: 3 });
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
export const CalcMode = Object.freeze({ Auto: 0, Manual: 1, AutoNoTable: 2 });
export const SheetVisibility = Object.freeze({ Visible: 0, Hidden: 1, VeryHidden: 2 });
export const GeometryMode = Object.freeze({ Display: 0, Print: 1 });
export const FilterKind = Object.freeze({ None: 0, Values: 1, Custom: 2, Top10: 3, Dynamic: 4, Color: 5, Icon: 6 });
export const FilterOperator = Object.freeze({
  Equal: 0,
  LessThan: 1,
  LessThanOrEqual: 2,
  NotEqual: 3,
  GreaterThanOrEqual: 4,
  GreaterThan: 5,
});
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
export const SortBy = Object.freeze({ Value: 0, CellColor: 1, FontColor: 2, Icon: 3 });
export const SortMethod = Object.freeze({ None: 0, PinYin: 1, Stroke: 2 });
export const DateTimeGrouping = Object.freeze({ Year: 0, Month: 1, Day: 2, Hour: 3, Minute: 4, Second: 5 });
export const ValidationErrorStyle = Object.freeze({ Stop: 0, Warning: 1, Information: 2 });
export const ColorContext = Object.freeze({ Font: 0, FillForeground: 1, FillBackground: 2, Border: 3 });
export const ColorResolution = Object.freeze({
  Exact: 0,
  DefaultTheme: 1,
  IndexOutOfRange: 2,
  ThemeUnparseable: 3,
  AutoContext: 4,
});
export const ImageFormat = Object.freeze({ Unknown: 0, Png: 1, Jpeg: 2, Gif: 3, Bmp: 4 });
export const DrawingObjectKind = Object.freeze({
  Picture: 0,
  Shape: 1,
  Chart: 2,
  Group: 3,
  Connector: 4,
  GraphicFrame: 5,
  Other: 6,
});
export const AnchorKind = Object.freeze({ OneCell: 0, TwoCell: 1, Absolute: 2 });
export const AnchorEditAs = Object.freeze({ TwoCell: 0, OneCell: 1, Absolute: 2 });
export const ThemeSource = Object.freeze({ Part: 0, Default: 1, Unparseable: 2 });
export const EffectiveStyleSource = Object.freeze({ Cell: 0, Row: 1, Column: 2, Default: 3 });
export const DisplayStatus = Object.freeze({ Ok: 0, Overflow: 1, InvalidFormat: 2 });
export const ExternalLinkKind = Object.freeze({ Unknown: 0, ExternalBook: 1, Ole: 2, Dde: 3 });
export const LogLevel = Object.freeze({ Debug: 0, Info: 1, Warn: 2, Error: 3, Off: 4 });

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
 *   matching the keys in `aliases` / `localized`. Independent of the profile id
 *   passed to the engine locale calls.
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
