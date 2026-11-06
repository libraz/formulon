// Value constants and pure-JS members declared in formulon.d.ts, shared by the
// `@libraz/formulon` and `@libraz/formulon/threads` entry shims and copied
// into `@libraz/formulon-native` by its staging script. Each package therefore
// has its own `FormulonError` class.

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
  PercentOf: 12,
  PercentOfRunningTotal: 13,
  RankAscending: 14,
  RankDescending: 15,
});
export const PIVOT_SHOW_AS_BASE_PREVIOUS = 1048828;
export const PIVOT_SHOW_AS_BASE_NEXT = 1048829;
export const PIVOT_VALUES_FIELD = -2;
export const PivotFilterType = Object.freeze({
  ValueTop10: 0,
  ValueGreaterThan: 1,
  ValueBetween: 2,
  LabelContains: 3,
  LabelBeginsWith: 4,
  LabelDate: 5,
  ValueEqual: 6,
  ValueNotEqual: 7,
  ValueGreaterThanOrEqual: 8,
  ValueLessThan: 9,
  ValueLessThanOrEqual: 10,
  ValueNotBetween: 11,
  LabelEqual: 12,
  LabelNotEqual: 13,
  LabelNotBeginsWith: 14,
  LabelEndsWith: 15,
  LabelNotEndsWith: 16,
  LabelNotContains: 17,
  LabelGreaterThan: 18,
  LabelGreaterThanOrEqual: 19,
  LabelLessThan: 20,
  LabelLessThanOrEqual: 21,
  LabelBetween: 22,
  LabelNotBetween: 23,
  ValueBottomCount: 24,
  ValueTopPercent: 25,
  ValueBottomPercent: 26,
  ValueTopSum: 27,
  ValueBottomSum: 28,
  DateEqual: 29,
  DateNotEqual: 30,
  DateOlderThan: 31,
  DateOlderThanOrEqual: 32,
  DateNewerThan: 33,
  DateNewerThanOrEqual: 34,
  DateNotBetween: 35,
  DateTomorrow: 36,
  DateToday: 37,
  DateYesterday: 38,
  DateNextWeek: 39,
  DateThisWeek: 40,
  DateLastWeek: 41,
  DateNextMonth: 42,
  DateThisMonth: 43,
  DateLastMonth: 44,
  DateNextQuarter: 45,
  DateThisQuarter: 46,
  DateLastQuarter: 47,
  DateNextYear: 48,
  DateThisYear: 49,
  DateLastYear: 50,
  DateYearToDate: 51,
  DateQ1: 52,
  DateQ2: 53,
  DateQ3: 54,
  DateQ4: 55,
  DateM1: 56,
  DateM2: 57,
  DateM3: 58,
  DateM4: 59,
  DateM5: 60,
  DateM6: 61,
  DateM7: 62,
  DateM8: 63,
  DateM9: 64,
  DateM10: 65,
  DateM11: 66,
  DateM12: 67,
});
export const PivotDatePeriod = Object.freeze({
  Seconds: 1,
  Minutes: 2,
  Hours: 4,
  Days: 8,
  Months: 16,
  Quarters: 32,
  Years: 64,
});
export const PivotCalendar = Object.freeze({ Gregorian: 0, Japanese: 1 });
export const PivotSortOrder = Object.freeze({ Descending: 0, Ascending: 1, Manual: 2 });
export const PivotReportLayout = Object.freeze({ Compact: 0, Tabular: 1, Outline: 2 });
export const PivotFilterValueKind = Object.freeze({ Int: 0, Double: 1, Text: 2 });
export const CfMatchKind = Object.freeze({ DifferentialFormat: 0, ColorScale: 1, DataBar: 2, IconSet: 3 });
export const CfDataBarDirection = Object.freeze({ Context: 0, LeftToRight: 1, RightToLeft: 2 });
export const CfDataBarAxisPosition = Object.freeze({ Automatic: 0, Middle: 1, None: 2 });
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
export const RefMode = Object.freeze({ A1: 0, R1C1: 1 });
export const CollationStrength = Object.freeze({ Compare: 0, Dynamic: 1, ClassicExact: 2 });
export const SheetVisibility = Object.freeze({ Visible: 0, Hidden: 1, VeryHidden: 2 });
export const SheetViewMode = Object.freeze({ Normal: 0, PageBreakPreview: 1, PageLayout: 2 });
export const PageOrientation = Object.freeze({ Portrait: 0, Landscape: 1 });
export const LocaleDateOrder = Object.freeze({ Mdy: 0, Ymd: 1, Dmy: 2 });
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
export const ColorKind = Object.freeze({ None: 0, Rgb: 1, Theme: 2, Indexed: 3, Auto: 4 });
export const FunctionAvailability = Object.freeze({
  Implemented: 0,
  ImplementedUnverified: 1,
  EnvironmentBound: 2,
  UnavailableStub: 3,
});
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
export const AnchorEditAs = Object.freeze({ OneCell: 0, TwoCell: 1, Absolute: 2 });
export const ShapePaint = Object.freeze({ Solid: 0, None: 1, Other: 2 });
export const SlicerSortOrder = Object.freeze({ Descending: 0, Ascending: 1 });
export const SlicerCrossFilter = Object.freeze({ ShowItemsWithDataAtTop: 0, ShowItemsWithNoData: 1, None: 2 });
export const SparklineType = Object.freeze({ Line: 0, Column: 1, Stacked: 2 });
export const SparklineAxisType = Object.freeze({ Individual: 0, Group: 1, Custom: 2 });
export const SparklineEmptyCells = Object.freeze({ Gap: 0, Zero: 1, Span: 2 });
export const ChartPlotType = Object.freeze({
  Bar: 0,
  Bar3D: 1,
  Line: 2,
  Line3D: 3,
  Area: 4,
  Area3D: 5,
  Pie: 6,
  Pie3D: 7,
  Doughnut: 8,
  OfPie: 9,
  Scatter: 10,
  Bubble: 11,
  Radar: 12,
  Stock: 13,
  Surface: 14,
  Surface3D: 15,
  Chartex: 16,
});
export const ChartInsertKind = Object.freeze({ Column: 0, Bar: 1, Line: 2, Area: 3, Pie: 4, Scatter: 5, Doughnut: 6 });
export const ChartGrouping = Object.freeze({ Clustered: 0, Stacked: 1, PercentStacked: 2, Standard: 3 });
export const ChartBlanksAs = Object.freeze({ Gap: 0, Zero: 1, Span: 2 });
export const ChartPointState = Object.freeze({ Value: 0, Gap: 1, Connect: 2 });
export const ChartAxisKind = Object.freeze({ Category: 0, Value: 1, Date: 2, Series: 3 });
export const ChartAxisCrosses = Object.freeze({ AutoZero: 0, Min: 1, Max: 2, At: 3 });
export const ChartTitleKind = Object.freeze({ Text: 1, Ref: 2, Auto: 3 });
export const ChartSeriesDirection = Object.freeze({ Auto: 0, Columns: 1, Rows: 2 });
export const ChartLegendPosition = Object.freeze({ Bottom: 0, Top: 1, Left: 2, Right: 3, TopRight: 4 });
export const ChartAxisPosition = Object.freeze({ Bottom: 0, Top: 1, Left: 2, Right: 3 });
export const ChartBarDirection = Object.freeze({ Col: 0, Bar: 1 });
export const ChartLineDash = Object.freeze({ Solid: 0, Dash: 1, Dot: 2, DashDot: 3, Other: 4 });
export const ChartMarkerSymbol = Object.freeze({
  None: 0,
  Auto: 1,
  Circle: 2,
  Square: 3,
  Diamond: 4,
  Triangle: 5,
  X: 6,
  Star: 7,
  Plus: 8,
  Dot: 9,
  Dash: 10,
  Picture: 11,
});
export const ChartLabelPosition = Object.freeze({
  BestFit: 0,
  Center: 1,
  InsideBase: 2,
  InsideEnd: 3,
  OutsideEnd: 4,
  Above: 5,
  Below: 6,
  Left: 7,
  Right: 8,
});
export const ThemeSource = Object.freeze({ Part: 0, Default: 1, Unparseable: 2 });
export const EffectiveStyleSource = Object.freeze({ Cell: 0, Row: 1, Column: 2, Default: 3 });
export const DisplayStatus = Object.freeze({ Ok: 0, Overflow: 1, InvalidFormat: 2 });
export const ExternalLinkKind = Object.freeze({ Unknown: 0, ExternalBook: 1, Ole: 2, Dde: 3 });
export const LogLevel = Object.freeze({ Debug: 0, Info: 1, Warn: 2, Error: 3, Off: 4 });
// <status-codes:begin> generated by tools/codegen/gen_status_codes.py; do not edit
export const StatusCode = Object.freeze({
  Ok: 0,
  UnknownError: 1,
  InvalidArgument: 2,
  NotImplemented: 3,
  OutOfMemory: 4,
  CancelRequested: 5,
  NotFound: 6,
  AlreadyExists: 7,
  InternalError: 8,
  PreconditionFailed: 9,
  ParserUnexpectedToken: 1000,
  ParserUnterminatedString: 1001,
  ParserUnterminatedComment: 1002,
  ParserInvalidNumber: 1003,
  ParserInvalidReference: 1004,
  ParserInvalidSheetName: 1005,
  ParserMismatchedBrackets: 1006,
  ParserExcessiveNesting: 1007,
  ParserUnknownFunction: 1008,
  ParserInvalidArrayLiteral: 1009,
  ParserInvalidStructRef: 1010,
  ParserInvalidR1C1: 1011,
  ParserInvalidDotNotation: 1012,
  ParserBomNotSupported: 1013,
  ParserTooManyErrors: 1014,
  EvalStackOverflow: 2000,
  EvalCircularReference: 2001,
  EvalLambdaArityMismatch: 2002,
  EvalInvalidByteCode: 2003,
  EvalTimeLimit: 2004,
  EvalMemoryLimit: 2005,
  EvalInvalidReference: 2006,
  EvalLambdaNotFound: 2007,
  EvalSpillCollision: 2008,
  EvalPivotMissing: 2009,
  EvalPivotInvalid: 2010,
  FnNotRegistered: 3000,
  FnWrongArgCount: 3001,
  FnArgTypeMismatch: 3002,
  FnDomainError: 3003,
  FnOverflow: 3004,
  FnUnderflow: 3005,
  FnRegexCompile: 3006,
  FnRegexMatchLimit: 3007,
  FnHostNotAvailable: 3008,
  GraphCycleDetected: 4000,
  GraphIterativeDiverged: 4001,
  GraphInvalidNodeRef: 4002,
  GraphScheduleFailed: 4003,
  GraphThreadPoolError: 4004,
  GraphRecalcReentrant: 4005,
  GraphIterationAborted: 4006,
  IoFileNotFound: 5000,
  IoFilePermission: 5001,
  IoFileTooLarge: 5002,
  IoZipCorrupt: 5003,
  IoZipEncrypted: 5004,
  IoZipBomb: 5005,
  IoZipSlip: 5006,
  IoXmlParse: 5007,
  IoXmlDoctype: 5008,
  IoXmlEntityExplosion: 5009,
  IoRelationshipBroken: 5010,
  IoContentTypeInvalid: 5011,
  IoSheetCorrupt: 5012,
  IoUnsupportedVariant: 5013,
  IoXlsbRecordCorrupt: 5014,
  IoCsvEncodingDetect: 5015,
  IoWriteFailed: 5016,
  IoCalcChainMismatch: 5017,
  IoXlsbRecordTruncated: 5018,
  IoXlsbUnknownRecord: 5019,
  IoXlsbUnsupportedPtg: 5020,
  IoXlsbCorrupt: 5021,
  IoXlsbRetainedPartStale: 5022,
  IoImageUnsupported: 5023,
  IoDrawingUnparseable: 5024,
  IoChartUnparseable: 5025,
  IoChartInvalidReference: 5026,
  IoChartUnsupportedKind: 5027,
  CryptoAgileNotSupported: 6000,
  CryptoStandardNotSupported: 6001,
  CryptoBadPassword: 6002,
  CryptoHashMismatch: 6003,
  CryptoKeyDerivationFailed: 6004,
  SecResourceLimit: 6005,
  SecExternalNotAllowed: 6006,
  SecPolicyBlocked: 6007,
  BindingInvalidHandle: 7000,
  BindingNullPointer: 7001,
  BindingUtf8EncodingError: 7002,
  BindingCallbackException: 7003,
  BindingWasmInitFailed: 7004,
  CliInvalidArgs: 7005,
  CliFileNotFound: 7006,
  CliOutputFailed: 7007,
  PivotSourceInvalid: 8000,
  PivotFieldNotFound: 8001,
  PivotAggregationFailed: 8002,
  SlicerNotConnected: 8003,
  SparklineInvalid: 8004,
  DataTableInvalid: 8005,
  AutoFilterInvalid: 8006,
  PrintFontMetricsMissing: 9000,
  PrintLayoutConvergence: 9001,
  PrintInvalidArea: 9002,
  UiViewStateInvalid: 9003,
  UiSnapshotFailed: 9004,
  PrintPageCountOverflow: 9005,
  ThreadedCommentInvalid: 9006,
  SheetIndexOutOfRange: 5050,
  InvalidSheetName: 5051,
  CannotRemoveLastSheet: 5052,
  SheetCountLimitExceeded: 5053,
  IndexOutOfRange: 5054,
});
// <status-codes:end>

/**
 * Merge a host-supplied metadata entry over the engine's structural
 * `getFunctionMetadata()` result.
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
 * @param {object} base A `FunctionMetadata`, as `getFunctionMetadata()` returns it.
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

/**
 * The error every host-side failure throws: an engine call that did not
 * complete, or a call on a disposed workbook. Failures are told apart by
 * `code`, a `StatusCode`; there are no subclasses.
 *
 * `instanceof` holds only for the class of the package that threw, so code
 * handling errors from both packages tests `e.name === 'FormulonError'`.
 */
export class FormulonError extends Error {
  /**
   * @param {number} code A `StatusCode`.
   * @param {string} message What went wrong, in English.
   * @param {string} [context] Structured `key=value` detail; `''` when none.
   */
  constructor(code, message, context = '') {
    super(message);
    this.name = 'FormulonError';
    this.code = code;
    this.context = context;
  }
}

/**
 * Member names of the frozen object `createFormulon()` resolves to, the same
 * on both packages.
 */
export const FORMULON_MEMBERS = Object.freeze([
  'Workbook',
  'FormulonError',
  'ValueKind',
  'StatusCode',
  'WorkbookFormat',
  'PivotCellKind',
  'PivotAxis',
  'PivotAggregation',
  'PivotShowValuesAs',
  'PivotFilterType',
  'PivotDatePeriod',
  'PivotCalendar',
  'PivotSortOrder',
  'PivotReportLayout',
  'PivotFilterValueKind',
  'CfMatchKind',
  'ErrorCode',
  'CalcMode',
  'RefMode',
  'CollationStrength',
  'GeometryMode',
  'DisplayStatus',
  'SheetVisibility',
  'LogLevel',
  'ExternalLinkKind',
  'FilterKind',
  'FilterOperator',
  'DynamicFilterType',
  'SortBy',
  'SortMethod',
  'DateTimeGrouping',
  'ValidationErrorStyle',
  'ColorContext',
  'ImageFormat',
  'DrawingObjectKind',
  'AnchorKind',
  'AnchorEditAs',
  'ShapePaint',
  'SlicerSortOrder',
  'SlicerCrossFilter',
  'SparklineType',
  'SparklineAxisType',
  'SparklineEmptyCells',
  'ChartPlotType',
  'ChartInsertKind',
  'ChartGrouping',
  'ChartBlanksAs',
  'ChartPointState',
  'ChartAxisKind',
  'ChartAxisCrosses',
  'ChartTitleKind',
  'ChartSeriesDirection',
  'ChartLegendPosition',
  'ChartLineDash',
  'ChartMarkerSymbol',
  'ChartLabelPosition',
  'ColorResolution',
  'ThemeSource',
  'EffectiveStyleSource',
  'PageOrientation',
  'SheetViewMode',
  'ChartBarDirection',
  'ChartAxisPosition',
  'LocaleDateOrder',
  'CfDataBarDirection',
  'CfDataBarAxisPosition',
  'ColorKind',
  'FunctionAvailability',
  'PIVOT_SHOW_AS_BASE_PREVIOUS',
  'PIVOT_SHOW_AS_BASE_NEXT',
  'PIVOT_VALUES_FIELD',
  'mergeFunctionMetadata',
  'evalFormula',
  'getVersion',
  'getStatusName',
  'getErrorDisplayName',
  'setLogMinLevel',
  'setLogSink',
]);

// Members the engine runtime (Emscripten module or addon) provides; every
// other member comes from this module. The runtime's raw `Workbook` class is
// not one of them: `buildFormulon` wraps it with `makeFacade`.
const RUNTIME_MEMBERS = new Set([
  'evalFormula',
  'getVersion',
  'getStatusName',
  'getErrorDisplayName',
  'setLogMinLevel',
  'setLogSink',
]);

const FORWARDED_OPTIONS = ['locateFile', 'print', 'printErr'];

/**
 * Builds the frozen `Formulon` object from an engine runtime.
 *
 * @param {object} runtime The Emscripten module or the native addon.
 * @param {object} shared This module's namespace object.
 * @returns {object} The object `createFormulon()` resolves to.
 */
export function buildFormulon(runtime, shared) {
  const formulon = {};
  for (const name of FORMULON_MEMBERS) {
    let value;
    if (name === 'Workbook') {
      value = runtime.Workbook === undefined ? undefined : makeFacade(runtime.Workbook);
    } else {
      value = RUNTIME_MEMBERS.has(name) ? runtime[name] : shared[name];
    }
    if (value === undefined) {
      throw new TypeError(`formulon: the engine runtime does not provide ${name}`);
    }
    formulon[name] = value;
  }
  return Object.freeze(formulon);
}

/**
 * Instantiates an Emscripten module factory and builds the `Formulon` object.
 * Only `locateFile`, `print` and `printErr` reach the factory.
 *
 * @param {Function} factory The Emscripten module factory.
 * @param {object} [options] `FormulonOptions`.
 * @param {object} shared This module's namespace object.
 * @returns {Promise<object>} The `Formulon` object.
 */
export async function loadFormulon(factory, options, shared) {
  const moduleOptions = { FormulonError };
  if (options !== undefined && options !== null) {
    for (const key of FORWARDED_OPTIONS) {
      if (options[key] !== undefined) moduleOptions[key] = options[key];
    }
  }
  const module = await factory(moduleOptions);
  return buildFormulon(module, shared);
}

// <facade-tables>
const WORKBOOK_METHODS = Object.freeze([
  'addBorder',
  'addCellStyleXf',
  'addCellXf',
  'addDxf',
  'addFill',
  'addFont',
  'addNumFmt',
  'addPerson',
  'addPivotCache',
  'addPivotCacheCalculatedField',
  'addPivotCacheCalculatedItem',
  'addPivotCacheField',
  'addPivotCacheFieldSharedItemBlank',
  'addPivotCacheFieldSharedItemBool',
  'addPivotCacheFieldSharedItemDate',
  'addPivotCacheFieldSharedItemError',
  'addPivotCacheFieldSharedItemNumber',
  'addPivotCacheFieldSharedItemText',
  'addPivotCacheRecord',
  'addSheet',
  'addTable',
  'applyTableAutoFilter',
  'canonicalizeFormula',
  'canonicalizeFunctionName',
  'clearPivotCacheFieldSharedItems',
  'clearPivotCacheFields',
  'clearPivotCacheRecords',
  'clearTableAutoFilter',
  'clearTheme',
  'compareText',
  'copySheet',
  'dispose',
  'disposed',
  'evaluateTableAutoFilter',
  'formatReference',
  'formatValue',
  'getActiveSheet',
  'getBorder',
  'getBorderCount',
  'getCalcMode',
  'getCalcOnSave',
  'getCellStyle',
  'getCellStyleXf',
  'getCellStyleXfCount',
  'getCellXf',
  'getCellXfCount',
  'getCollationClasses',
  'getDate1904',
  'getDefinedName',
  'getDiagnostics',
  'getDxf',
  'getDxfCount',
  'getExcelProfileId',
  'getFileName',
  'getFill',
  'getFillCount',
  'getFont',
  'getFontCount',
  'getFullPrecision',
  'getFunctionMetadata',
  'getImageInfo',
  'getIterative',
  'getLocaleFacts',
  'getMemoryUsage',
  'getNumFmtString',
  'getPassthrough',
  'getPerson',
  'getPinnedNow',
  'getPivotCacheFieldCount',
  'getPivotCacheFieldFormula',
  'getPivotCacheFieldName',
  'getPivotCacheFieldSharedItemCount',
  'getPivotCacheFieldSharedItemValue',
  'getPivotCacheOptions',
  'getPivotCacheRecordCount',
  'getPivotCacheWorksheetSource',
  'getRefMode',
  'getTableAutoFilter',
  'getTableStyle',
  'getTableStyleDefaults',
  'getThemeColors',
  'getThemeFonts',
  'getThemeScriptFont',
  'getWorkbookProtection',
  'isPivotCacheFieldSharedItemDate',
  'listCellStyles',
  'listDefinedNames',
  'listExternalLinks',
  'listFunctionNames',
  'listLocaleErrorNames',
  'listPassthroughs',
  'listPeople',
  'listPivotCacheCalculatedItems',
  'listPivotCaches',
  'listSheets',
  'listTableStyles',
  'listTables',
  'localizeFormula',
  'localizeFunctionName',
  'movePivotCacheCalculatedItem',
  'moveSheet',
  'parseReference',
  'recalc',
  'recalcDataTables',
  'recalcParallel',
  'removeCellStyle',
  'removeDefinedName',
  'removePerson',
  'removePivotCache',
  'removePivotCacheCalculatedField',
  'removePivotCacheCalculatedItem',
  'removeSheet',
  'removeTable',
  'removeTableAutoFilter',
  'removeTableStyle',
  'renameSheet',
  'renameTableColumn',
  'resolveColor',
  'save',
  'saveWithDiagnostics',
  'setActiveSheet',
  'setCalcMode',
  'setCalcOnSave',
  'setCellStyle',
  'setDate1904',
  'setDefaultFont',
  'setDefinedName',
  'setExcelProfileId',
  'setFileName',
  'setFont',
  'setFullPrecision',
  'setIterative',
  'setIterativeProgress',
  'setPinnedNow',
  'setPivotCacheCalculatedFieldFormula',
  'setPivotCacheCalculatedItemFormula',
  'setPivotCacheOptions',
  'setPivotCacheRecordBlank',
  'setPivotCacheRecordBool',
  'setPivotCacheRecordDate',
  'setPivotCacheRecordError',
  'setPivotCacheRecordNumber',
  'setPivotCacheRecordText',
  'setPivotCacheWorksheetSource',
  'setRefMode',
  'setTableAutoFilter',
  'setTableColumnTotals',
  'setTableStyle',
  'setTableStyleDefaults',
  'setThemeColors',
  'setThemeFonts',
  'setThemeScriptFont',
  'setWorkbookProtection',
  'sheet',
  'updateTable',
]);
const SHEET_METHODS = Object.freeze([
  'addCfRule',
  'addChart',
  'addChartSeries',
  'addColumnBreak',
  'addDataTable',
  'addHyperlink',
  'addHyperlinkRange',
  'addImage',
  'addMerge',
  'addPivot',
  'addPivotDataField',
  'addPivotField',
  'addPivotFieldItem',
  'addPivotFieldSharedItem',
  'addPivotFieldSubtotalFn',
  'addPivotFilter',
  'addProtectedRange',
  'addRowBreak',
  'addShape',
  'addSlicer',
  'addSparklineGroup',
  'addThreadedComment',
  'addValidation',
  'applyAutoFilter',
  'clearAutoFilter',
  'clearBreaks',
  'clearCellXfIndex',
  'clearCfRules',
  'clearColumnStyle',
  'clearColumnWidth',
  'clearHyperlinks',
  'clearMerges',
  'clearPivotDataFields',
  'clearPivotFieldDateGroup',
  'clearPivotFieldItems',
  'clearPivotFieldSubtotalFns',
  'clearPivotFields',
  'clearPivotFilters',
  'clearRowHeight',
  'clearRowStyle',
  'clearValidations',
  'convertColumnCharsToPt',
  'convertColumnPtToChars',
  'deleteColumns',
  'deleteRows',
  'evaluateAutoFilter',
  'evaluateCfFormula',
  'evaluateCfRange',
  'evaluateFormula',
  'evaluateFormulaArray',
  'getArrayFormula',
  'getAutoFilter',
  'getAutoFilterXml',
  'getCellAt',
  'getCellCount',
  'getCellPhonetic',
  'getCellPhoneticProperties',
  'getCellPhoneticRuns',
  'getCellRectPt',
  'getCellShowPhonetic',
  'getCellXfIndex',
  'getCellsInRange',
  'getChart',
  'getChartData',
  'getChartPreview',
  'getColumnWidthPt',
  'getComment',
  'getCommentVisible',
  'getDataTable',
  'getDependents',
  'getDisplayText',
  'getDrawingObject',
  'getEffectiveStyle',
  'getFormatDefaults',
  'getFormula',
  'getFormulaR1C1',
  'getHeaderFooterXml',
  'getHyperlink',
  'getImage',
  'getInvalidCells',
  'getLambdaText',
  'getMergesInRange',
  'getPageMargins',
  'getPageMarginsXml',
  'getPageSetup',
  'getPageSetupXml',
  'getPagination',
  'getPivotCacheId',
  'getPivotColumnFieldOrder',
  'getPivotCount',
  'getPivotDataFieldDefaultName',
  'getPivotFieldItem',
  'getPivotFieldItemCount',
  'getPivotGrandTotals',
  'getPivotLayout',
  'getPivotName',
  'getPivotOptions',
  'getPivotReportLayout',
  'getPivotRowFieldOrder',
  'getPivotStyleInfo',
  'getPrecedents',
  'getPrintArea',
  'getPrintOptionsXml',
  'getPrintTitles',
  'getRangeStyleBases',
  'getRangeStyleBasesExcluding',
  'getRowHeightPt',
  'getSelection',
  'getShape',
  'getSheetPrXml',
  'getSheetProtection',
  'getSlicer',
  'getSpillInfo',
  'getThreadedComment',
  'getValue',
  'getView',
  'getWidthModel',
  'insertColumns',
  'insertRows',
  'isPivotProjectionPending',
  'listCfRules',
  'listColumnBreaks',
  'listColumnOverrides',
  'listComments',
  'listDrawingObjects',
  'listHyperlinks',
  'listMerges',
  'listPivotDataFields',
  'listPivotFields',
  'listPivotFilters',
  'listProtectedRanges',
  'listRowBreaks',
  'listRowOverrides',
  'listShapes',
  'listSlicers',
  'listSparklineGroups',
  'listThreadedComments',
  'listValidations',
  'projectPivot',
  'recalcViewport',
  'remapRangeStyle',
  'removeAutoFilter',
  'removeCfRule',
  'removeChartSeries',
  'removeColumnBreak',
  'removeComment',
  'removeDataTable',
  'removeDrawingObject',
  'removeHyperlink',
  'removeHyperlinkAt',
  'removeMerge',
  'removeMergeAt',
  'removePivot',
  'removePivotFilter',
  'removeProtectedRange',
  'removeRowBreak',
  'removeShape',
  'removeSlicer',
  'removeSparklineGroup',
  'removeThreadedComment',
  'removeValidation',
  'restoreDrawing',
  'restoreRangeStyle',
  'setArrayFormula',
  'setAutoFilter',
  'setAutoFilterXml',
  'setBlank',
  'setBool',
  'setCellPhonetic',
  'setCellPhoneticProperties',
  'setCellPhoneticRuns',
  'setCellShowPhonetic',
  'setCellXfIndex',
  'setChartAxis',
  'setChartDisplay',
  'setChartFilter',
  'setChartLegend',
  'setChartPlot',
  'setChartPoint',
  'setChartSeries',
  'setChartSource',
  'setChartTitle',
  'setChartType',
  'setColumnCollapsed',
  'setColumnHidden',
  'setColumnOutline',
  'setColumnStyle',
  'setColumnWidth',
  'setComment',
  'setCommentVisible',
  'setDrawingAnchor',
  'setDrawingName',
  'setDrawingZOrder',
  'setError',
  'setFitToPage',
  'setFormatDefaults',
  'setFormula',
  'setFreeze',
  'setHeaderFooter',
  'setHeaderFooterXml',
  'setNumber',
  'setPageMargins',
  'setPageMarginsXml',
  'setPageSetup',
  'setPageSetupXml',
  'setPaneTopLeft',
  'setPivotAnchor',
  'setPivotColumnFieldOrder',
  'setPivotDataField',
  'setPivotFieldAxis',
  'setPivotFieldCustomName',
  'setPivotFieldDateGroup',
  'setPivotFieldDefaultSubtotal',
  'setPivotFieldFillDownLabels',
  'setPivotFieldInsertBlankRow',
  'setPivotFieldItemCaption',
  'setPivotFieldItemShowDetail',
  'setPivotFieldItemVisible',
  'setPivotFieldNumberFormat',
  'setPivotFieldShowAll',
  'setPivotFieldShowDetail',
  'setPivotFieldSort',
  'setPivotFieldSubtotalTop',
  'setPivotGrandTotals',
  'setPivotName',
  'setPivotOptions',
  'setPivotReportLayout',
  'setPivotRowFieldOrder',
  'setPivotStyleInfo',
  'setPrintArea',
  'setPrintOptions',
  'setPrintOptionsXml',
  'setPrintTitles',
  'setRangeXfIndex',
  'setRightToLeft',
  'setRowAutoHeight',
  'setRowCollapsed',
  'setRowHeight',
  'setRowHidden',
  'setRowOutline',
  'setRowStyle',
  'setSelection',
  'setSheetPrXml',
  'setSheetProtection',
  'setShowFormulas',
  'setShowGridLines',
  'setShowOutlineSymbols',
  'setShowRowColHeaders',
  'setShowZeros',
  'setSparklineGroupSparklines',
  'setSplit',
  'setTabSelected',
  'setText',
  'setThreadResolved',
  'setTopLeft',
  'setValue',
  'setViewMode',
  'setVisibility',
  'setZoom',
  'showPivotDetails',
  'snapshotDrawing',
  'snapshotRangeStyle',
  'updateProtectedRange',
  'updateShape',
  'updateSlicer',
  'updateSparklineGroup',
  'updateThreadedComment',
  'validateValue',
]);
const SHEET_STRUCTURE_OPS = Object.freeze([
  'addSheet',
  'copySheet',
  'moveSheet',
  'removeSheet',
  'renameSheet',
  'showPivotDetails',
]);
const LIST_FAMILIES = Object.freeze({
  listCfRules: ['getCfRuleCount', 'getCfRule'],
  listExternalLinks: ['getExternalLinkCount', 'getExternalLink'],
  listPivotCacheCalculatedItems: ['getPivotCacheCalculatedItemCount', 'getPivotCacheCalculatedItem'],
  listPivotDataFields: ['getPivotDataFieldCount', 'getPivotDataField'],
  listPivotFields: ['getPivotFieldCount', 'getPivotField'],
  listPivotFilters: ['getPivotFilterCount', 'getPivotFilter'],
  listProtectedRanges: ['getProtectedRangeCount', 'getProtectedRange'],
  listSparklineGroups: ['getSparklineGroupCount', 'getSparklineGroup'],
  listTables: ['getTableCount', 'getTable'],
  listValidations: ['getValidationCount', 'getValidation'],
});
const HANDWRITTEN = Object.freeze([
  'addHyperlink',
  'addHyperlinkRange',
  'addSheet',
  'copySheet',
  'getActiveSheet',
  'getCellPhoneticRuns',
  'getFormula',
  'getThemeColors',
  'getThemeFonts',
  'listSheets',
  'moveSheet',
  'removeSheet',
  'renameSheet',
  'setActiveSheet',
  'setArrayFormula',
  'setChartSource',
  'setComment',
  'setFormula',
  'setPinnedNow',
  'setPivotFieldDateGroup',
  'setPivotGrandTotals',
  'setThemeColors',
  'setThemeScriptFont',
  'setValue',
  'showPivotDetails',
  'updateThreadedComment',
]);
const CELL_ARGS = Object.freeze({
  addHyperlink: [[0, 1]],
  addPivot: [[2, 3]],
  clearCellXfIndex: [[0, 1]],
  evaluateCfFormula: [
    [0, 1],
    [2, 3],
  ],
  evaluateFormula: [[0, 1]],
  evaluateFormulaArray: [[0, 1]],
  getArrayFormula: [[0, 1]],
  getCellPhonetic: [[0, 1]],
  getCellPhoneticProperties: [[0, 1]],
  getCellPhoneticRuns: [[0, 1]],
  getCellShowPhonetic: [[0, 1]],
  getCellXfIndex: [[0, 1]],
  getComment: [[0, 1]],
  getCommentVisible: [[0, 1]],
  getDataTable: [[0, 1]],
  getDependents: [[0, 1]],
  getDisplayText: [[0, 1]],
  getEffectiveStyle: [[0, 1]],
  getFormula: [[0, 1]],
  getFormulaR1C1: [[0, 1]],
  getHyperlink: [[0, 1]],
  getLambdaText: [[0, 1]],
  getPrecedents: [[0, 1]],
  getSpillInfo: [[0, 1]],
  getValue: [[0, 1]],
  removeComment: [[0, 1]],
  removeDataTable: [[0, 1]],
  removeHyperlink: [[0, 1]],
  setBlank: [[0, 1]],
  setBool: [[0, 1]],
  setCellPhonetic: [[0, 1]],
  setCellPhoneticProperties: [[0, 1]],
  setCellPhoneticRuns: [[0, 1]],
  setCellShowPhonetic: [[0, 1]],
  setCellXfIndex: [[0, 1]],
  setComment: [[0, 1]],
  setCommentVisible: [[0, 1]],
  setError: [[0, 1]],
  setFormula: [[0, 1]],
  setNumber: [[0, 1]],
  setPaneTopLeft: [[0, 1]],
  setText: [[0, 1]],
  setTopLeft: [[0, 1]],
  setValue: [[0, 1]],
  showPivotDetails: [[1, 2]],
  validateValue: [[0, 1]],
});
const RECT_ARGS = Object.freeze({
  addHyperlinkRange: [0],
  addMerge: [0],
  evaluateCfRange: [0],
  getCellRectPt: [0],
  getCellsInRange: [0],
  getMergesInRange: [0],
  getRangeStyleBases: [0],
  getRangeStyleBasesExcluding: [0],
  recalcViewport: [0],
  remapRangeStyle: [0],
  removeMerge: [0],
  setPivotAnchor: [1],
  setPrintArea: [0],
  setRangeXfIndex: [0],
  snapshotRangeStyle: [0],
});
// </facade-tables>

// Arguments the declarations type `T | null` (or optional) where the raw method
// takes the C sentinel: method -> [position in the public call, sentinel,
// whether an omitted argument also means absent].
const WORKBOOK_ABSENT_ARGS = Object.freeze({
  addPivotCache: [0, 0, true],
  getDefinedName: [1, -1],
  removeDefinedName: [1, -1],
  setFileName: [0, ''],
  setWorkbookProtection: [0, {}],
});
const SHEET_ABSENT_ARGS = Object.freeze({
  setAutoFilterXml: [0, ''],
  setCellPhonetic: [2, ''],
  setCellPhoneticRuns: [2, []],
  setChartLegend: [1, -1],
  setHeaderFooterXml: [0, ''],
  setPageMarginsXml: [0, ''],
  setPageSetupXml: [0, ''],
  setPivotFieldAxis: [2, -1],
  setPivotFieldCustomName: [2, ''],
  setPivotFieldItemCaption: [3, ''],
  setPivotFieldNumberFormat: [2, ''],
  setPrintArea: [0, []],
  setPrintOptionsXml: [0, ''],
  setSheetPrXml: [0, ''],
  setSheetProtection: [0, {}],
});

// Patch setters and the non-nullable fields of their patch: method -> [position
// of the patch in the public call, { field: expected type }]. `null` on one of
// these fields is a TypeError; an omitted or `undefined` key keeps the value.
const bools = (...fields) => Object.fromEntries(fields.map((f) => [f, 'boolean']));
const numbers = (...fields) => Object.fromEntries(fields.map((f) => [f, 'number']));
const WORKBOOK_PATCHES = Object.freeze({
  setIterative: [0, { ...bools('enabled'), ...numbers('maxIterations', 'maxChange') }],
  setPivotCacheOptions: [1, bools('saveData', 'refreshOnLoad')],
});
const SHEET_PATCHES = Object.freeze({
  setPageSetup: [0, { ...numbers('paperSize', 'scale', 'fitToWidth', 'fitToHeight'), ...bools('fitToPage') }],
  setPageMargins: [0, numbers('left', 'right', 'top', 'bottom', 'header', 'footer')],
  setFormatDefaults: [0, { ...numbers('baseColWidth'), ...bools('customHeight', 'zeroHeight') }],
  setPivotOptions: [
    1,
    {
      ...bools(
        'showError',
        'showMissing',
        'showHeaders',
        'showDrill',
        'showDataTips',
        'gridDropZones',
        'showEmptyRow',
        'showEmptyCol',
        'showDataDropDown',
        'printDrill',
        'fieldPrintTitles',
        'itemPrintTitles',
        'mergeItem',
        'pageOverThenDown',
        'multipleFieldFilters',
        'preserveFormatting',
        'customListSort',
        'autoFormatting',
        'enableDrill',
      ),
      ...numbers('indent', 'pageWrap'),
    },
  ],
});

function isObject(value) {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function requireObject(value, where) {
  if (!isObject(value)) throw new TypeError(`${where} must be an object`);
  return value;
}

function requireString(value, where) {
  if (typeof value !== 'string') throw new TypeError(`${where} must be a primitive string`);
  return value;
}

function requireBoolean(value, where) {
  if (typeof value !== 'boolean') throw new TypeError(`${where} must be a primitive boolean`);
  return value;
}

function requireArray(value, where) {
  if (!Array.isArray(value)) throw new TypeError(`${where} must be an array`);
  return value;
}

function requireNumber(value, where) {
  if (typeof value !== 'number') throw new TypeError(`${where} must be a primitive number`);
  return value;
}

function requireInteger(value, where) {
  if (typeof value !== 'number') throw new TypeError(`${where} must be a primitive number`);
  if (!Number.isInteger(value)) throw new RangeError(`${where} is outside int32 range`);
  return value;
}

/** A patch field: `undefined` keeps `current`, `null` is refused, anything else must pass `check`. */
function patchField(patch, field, current, check, where) {
  const value = patch[field];
  return value === undefined ? current : check(value, `${where} ${field}`);
}

function invalidAddress(a1) {
  return new FormulonError(
    StatusCode.InvalidArgument,
    `${JSON.stringify(a1)} is not a single-cell A1 address without a sheet name`,
    `address=${a1}`,
  );
}

// The only A1 reader: whatever `parseReference` accepts as one unqualified
// cell is an address; everything else is rejected.
function cellAddress(raw, a1) {
  const ref = raw.parseReference(a1);
  if (ref === null || ref.sheetName !== null) throw invalidAddress(a1);
  const { rect } = ref;
  if (rect.firstRow !== rect.lastRow || rect.firstCol !== rect.lastCol) throw invalidAddress(a1);
  return rect;
}

function requireValueField(value, field, type) {
  if (typeof value[field] !== type) {
    throw new TypeError(`Sheet.setValue argument 3 must carry a ${type} \`${field}\` for its kind`);
  }
  return value[field];
}

function defineMethod(proto, name, fn) {
  Object.defineProperty(proto, name, { value: fn, writable: true, configurable: true });
}

function namedFunction(name, body) {
  return {
    [name](...args) {
      return body.call(this, args);
    },
  }[name];
}

/**
 * Builds the public `Workbook` and `Sheet` classes over a runtime's raw
 * `Workbook` class. Each call returns new classes, so a workbook of one
 * `createFormulon()` result is not an `instanceof` the `Workbook` of another.
 *
 * `Workbook` holds the workbook-scope ops and `Sheet` the sheet-scope ones; a
 * `Sheet` method calls the raw method with the sheet's current index first.
 * The workbook keeps one `Sheet` per sheet in sheet order and updates that list
 * after every op in `SHEET_STRUCTURE_OPS`, so a handle follows its sheet across
 * inserts, moves and renames, and a removed sheet's handle throws `NotFound`.
 *
 * @param {Function} Raw The runtime's raw `Workbook` class.
 * @returns {Function} The public `Workbook` class.
 */
function makeFacade(Raw) {
  const token = Symbol('formulon.facade');
  let rawOf;
  let isDisposed;
  let sheetIndex;
  let sheetState;
  let insertSheet;

  class Sheet {
    #workbook;
    #position;
    #name;

    constructor(key, workbook, position, name) {
      if (key !== token) throw new TypeError('Sheet has no public constructor; use Workbook.sheet');
      this.#workbook = workbook;
      this.#position = position;
      this.#name = name;
    }

    #current() {
      if (isDisposed(this.#workbook)) {
        throw new FormulonError(StatusCode.BindingInvalidHandle, 'the workbook is disposed', '');
      }
      if (this.#position < 0) {
        throw new FormulonError(StatusCode.NotFound, 'the sheet was removed', `sheet_name=${this.#name}`);
      }
      return this.#position;
    }

    static {
      sheetIndex = (sheet) => sheet.#current();
      sheetState = {
        create: (workbook, position, name) => new Sheet(token, workbook, position, name),
        place: (sheet, position) => {
          sheet.#position = position;
        },
        rename: (sheet, name) => {
          sheet.#name = name;
        },
      };
    }

    get workbook() {
      return this.#workbook;
    }

    get index() {
      return this.#current();
    }

    get name() {
      this.#current();
      return this.#name;
    }

    get(a1) {
      const raw = rawOf(this.#workbook);
      const cell = cellAddress(raw, a1);
      return raw.getValue(this.#current(), cell.firstRow, cell.firstCol);
    }

    set(a1, value) {
      const cell = cellAddress(rawOf(this.#workbook), a1);
      this.setValue(cell.firstRow, cell.firstCol, value);
    }

    getFormula(rowOrA1, col) {
      const raw = rawOf(this.#workbook);
      if (typeof rowOrA1 === 'string') {
        const cell = cellAddress(raw, rowOrA1);
        return raw.getFormula(this.#current(), cell.firstRow, cell.firstCol);
      }
      return raw.getFormula(this.#current(), rowOrA1, col);
    }

    setFormula(rowOrA1, ...rest) {
      const raw = rawOf(this.#workbook);
      if (typeof rowOrA1 === 'string') {
        const cell = cellAddress(raw, rowOrA1);
        return raw.setFormula(this.#current(), cell.firstRow, cell.firstCol, rest[0]);
      }
      return raw.setFormula(this.#current(), rowOrA1, rest[0], rest[1]);
    }

    setValue(row, col, value) {
      if (value === null) return this.setBlank(row, col);
      switch (typeof value) {
        case 'number':
          return this.setNumber(row, col, value);
        case 'string':
          return this.setText(row, col, value);
        case 'boolean':
          return this.setBool(row, col, value);
        case 'object':
          if (Array.isArray(value)) break;
          switch (value.kind) {
            case ValueKind.Blank:
              return this.setBlank(row, col);
            case ValueKind.Number:
              return this.setNumber(row, col, requireValueField(value, 'number', 'number'));
            case ValueKind.Bool:
              return this.setBool(row, col, requireValueField(value, 'boolean', 'boolean'));
            case ValueKind.Text:
              return this.setText(row, col, requireValueField(value, 'text', 'string'));
            case ValueKind.Error:
              return this.setError(row, col, requireValueField(value, 'errorCode', 'number'));
          }
          break;
      }
      throw new TypeError('Sheet.setValue argument 3 must be a number, string, boolean, null or a scalar Value');
    }

    showPivotDetails(pivotIndex, row, col) {
      const workbook = this.#workbook;
      const position = rawOf(workbook).showPivotDetails(this.#current(), pivotIndex, row, col);
      return insertSheet(workbook, position);
    }

    setComment(row, col, comment) {
      const where = 'Sheet.setComment';
      requireObject(comment, `${where} argument 3`);
      const author = requireString(comment.author, `${where} author`);
      const text = requireString(comment.text, `${where} text`);
      rawOf(this.#workbook).setComment(this.#current(), row, col, author, text);
    }

    addHyperlink(row, col, link) {
      rawOf(this.#workbook).addHyperlink(this.#current(), row, col, ...hyperlinkArgs(link, 'Sheet.addHyperlink', 3));
    }

    addHyperlinkRange(range, link) {
      rawOf(this.#workbook).addHyperlinkRange(
        this.#current(),
        range,
        ...hyperlinkArgs(link, 'Sheet.addHyperlinkRange', 2),
      );
    }

    updateThreadedComment(id, patch) {
      const where = 'Sheet.updateThreadedComment';
      requireObject(patch, `${where} argument 2`);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const current =
        patch.text === undefined || patch.mentions === undefined ? raw.getThreadedComment(sheet, id) : null;
      const text = patchField(patch, 'text', current?.text ?? '', requireString, where);
      const mentions = patchField(patch, 'mentions', current?.mentions ?? [], requireArray, where);
      raw.updateThreadedComment(sheet, id, text, mentions);
    }

    setPivotGrandTotals(pivotIndex, patch) {
      const where = 'Sheet.setPivotGrandTotals';
      requireObject(patch, `${where} argument 2`);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const current =
        patch.rows === undefined || patch.columns === undefined ? raw.getPivotGrandTotals(sheet, pivotIndex) : {};
      const rows = patchField(patch, 'rows', current.rows, requireBoolean, where);
      const columns = patchField(patch, 'columns', current.columns, requireBoolean, where);
      raw.setPivotGrandTotals(sheet, pivotIndex, rows, columns);
    }

    setArrayFormula(arrayFormula) {
      const where = 'Sheet.setArrayFormula';
      requireObject(arrayFormula, `${where} argument 1`);
      const rect = requireObject(arrayFormula.rect, `${where} rect`);
      const formula = requireString(arrayFormula.formula, `${where} formula`);
      rawOf(this.#workbook).setArrayFormula(this.#current(), rect, formula);
    }

    setChartSource(objectId, source) {
      const where = 'Sheet.setChartSource';
      requireObject(source, `${where} argument 2`);
      const flag = (field) =>
        source[field] === undefined || source[field] === null
          ? undefined
          : requireBoolean(source[field], `${where} ${field}`);
      const args = [
        requireString(source.source, `${where} source`),
        requireInteger(source.direction, `${where} direction`),
      ];
      const headerRow = flag('headerRow');
      const headerColumn = flag('headerColumn');
      if (headerRow !== undefined || headerColumn !== undefined) args.push(headerRow, headerColumn);
      rawOf(this.#workbook).setChartSource(this.#current(), objectId, ...args);
    }

    setPivotFieldDateGroup(pivotIndex, fieldIndex, group) {
      const where = 'Sheet.setPivotFieldDateGroup';
      requireObject(group, `${where} argument 3`);
      const optional = (field) =>
        group[field] === undefined ? undefined : requireNumber(group[field], `${where} ${field}`);
      const args = [
        requireInteger(group.periods, `${where} periods`),
        group.calendar === undefined ? PivotCalendar.Gregorian : requireInteger(group.calendar, `${where} calendar`),
      ];
      const tail = [optional('intervalDays'), optional('start'), optional('end')];
      while (tail.length > 0 && tail.at(-1) === undefined) tail.pop();
      rawOf(this.#workbook).setPivotFieldDateGroup(this.#current(), pivotIndex, fieldIndex, ...args, ...tail);
    }

    setPageSetup(setup) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      refuseNullFields('Sheet', 'setPageSetup', setup, SHEET_PATCHES.setPageSetup[1]);
      // A stated orientation cannot be removed; `null` keeps an unstated one.
      if (isObject(setup) && setup.orientation === null) {
        if (raw.getPageSetup(sheet).orientationStated) {
          throw new FormulonError(
            StatusCode.InvalidArgument,
            'a stated page orientation cannot be removed',
            `sheet_index=${sheet}`,
          );
        }
        setup = { ...setup, orientation: undefined };
      }
      raw.setPageSetup(sheet, setup);
    }

    getCellPhoneticRuns(row, col) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const count = raw.getCellPhoneticRunCount(sheet, row, col);
      const runs = [];
      for (let i = 0; i < count; i++) runs.push(raw.getCellPhoneticRun(sheet, row, col, i));
      return runs;
    }
  }

  function hyperlinkArgs(link, where, position) {
    requireObject(link, `${where} argument ${position}`);
    const optional = (field) => (link[field] === undefined ? '' : requireString(link[field], `${where} ${field}`));
    const target =
      link.target === undefined || link.target === null ? null : requireString(link.target, `${where} target`);
    const location = optional('location');
    if (target === null && link.location === undefined) {
      throw new TypeError(`${where} location is required when target is null`);
    }
    return [target ?? '', optional('display'), optional('tooltip'), location];
  }

  class Workbook {
    #raw;
    #sheets;
    #disposed = false;

    constructor(key, raw) {
      if (key !== token) {
        throw new TypeError(
          'Workbook has no public constructor; use Workbook.load, createDefault, createEmpty or createForProfile',
        );
      }
      this.#raw = raw;
      this.#sheets = [];
      const count = raw.getSheetCount();
      for (let i = 0; i < count; i++) this.#sheets.push(sheetState.create(this, i, raw.getSheetName(i)));
    }

    static {
      rawOf = (workbook) => workbook.#raw;
      isDisposed = (workbook) => workbook.#disposed;
      insertSheet = (workbook, position) => workbook.#insert(position);
    }

    static load(bytes) {
      return new Workbook(token, Raw.load(bytes));
    }

    static createDefault() {
      return new Workbook(token, Raw.createDefault());
    }

    static createEmpty() {
      return new Workbook(token, Raw.createEmpty());
    }

    static createForProfile(profileId) {
      return new Workbook(token, Raw.createForProfile(profileId));
    }

    get disposed() {
      return this.#disposed;
    }

    dispose() {
      this.#raw.dispose();
      this.#disposed = true;
    }

    #alive() {
      if (this.#disposed) throw new FormulonError(StatusCode.BindingInvalidHandle, 'the workbook is disposed', '');
    }

    #renumber() {
      for (const [i, sheet] of this.#sheets.entries()) sheetState.place(sheet, i);
    }

    #insert(position) {
      const sheet = sheetState.create(this, position, this.#raw.getSheetName(position));
      this.#sheets.splice(position, 0, sheet);
      this.#renumber();
      return sheet;
    }

    // Resolves a `SheetRef` the way `sheet(key)` does; `method` and `position` word the errors.
    #resolve(ref, method, position) {
      const where = `Workbook.${method} argument ${position}`;
      if (ref instanceof Sheet) {
        if (ref.workbook !== this) throw new TypeError(`${where} is a Sheet of another workbook`);
        return ref.index;
      }
      if (typeof ref === 'string') {
        this.#alive();
        const index = this.#raw.getSheetIndex(ref);
        if (index === null) {
          throw new FormulonError(StatusCode.NotFound, `no sheet is named ${JSON.stringify(ref)}`, `sheet_name=${ref}`);
        }
        return index;
      }
      if (typeof ref === 'number') {
        if (!Number.isInteger(ref) || ref < 0 || ref > 0xffffffff) {
          throw new RangeError(`${where} is outside uint32 range`);
        }
        this.#alive();
        if (ref >= this.#sheets.length) {
          throw new FormulonError(
            StatusCode.SheetIndexOutOfRange,
            `sheet index ${ref} is out of range`,
            `sheet_index=${ref}`,
          );
        }
        return ref;
      }
      throw new TypeError(`${where} must be a sheet name, a sheet index or a Sheet`);
    }

    sheet(key) {
      return this.#sheets[this.#resolve(key, 'sheet', 1)];
    }

    listSheets() {
      this.#alive();
      return [...this.#sheets];
    }

    getActiveSheet() {
      this.#alive();
      return this.#sheets[this.#raw.getActiveSheet()];
    }

    setActiveSheet(sheet) {
      this.#raw.setActiveSheet(this.#resolve(sheet, 'setActiveSheet', 1));
    }

    addSheet(name) {
      return this.#insert(this.#raw.addSheet(name));
    }

    copySheet(source, insertBefore, newName) {
      const index = this.#resolve(source, 'copySheet', 1);
      const position =
        newName === undefined
          ? this.#raw.copySheet(index, insertBefore)
          : this.#raw.copySheet(index, insertBefore, newName);
      return this.#insert(position);
    }

    moveSheet(sheet, toIndex) {
      const from = this.#resolve(sheet, 'moveSheet', 1);
      this.#raw.moveSheet(from, toIndex);
      const [moved] = this.#sheets.splice(from, 1);
      this.#sheets.splice(toIndex, 0, moved);
      this.#renumber();
    }

    removeSheet(sheet) {
      const index = this.#resolve(sheet, 'removeSheet', 1);
      this.#raw.removeSheet(index);
      const [removed] = this.#sheets.splice(index, 1);
      sheetState.place(removed, -1);
      this.#renumber();
    }

    renameSheet(sheet, name) {
      const index = this.#resolve(sheet, 'renameSheet', 1);
      this.#raw.renameSheet(index, name);
      sheetState.rename(this.#sheets[index], this.#raw.getSheetName(index));
    }

    setPinnedNow(time) {
      if (time === null) return this.#raw.clearPinnedNow();
      const where = 'Workbook.setPinnedNow';
      requireObject(time, `${where} argument 1`);
      const fields = ['year', 'month', 'day', 'hour', 'minute', 'second'];
      return this.#raw.setPinnedNow(...fields.map((f) => requireInteger(time[f], `${where} ${f}`)));
    }

    getThemeFonts() {
      return this.#raw.getThemeFonts();
    }

    getThemeColors() {
      return this.#raw.getThemeColors();
    }

    setThemeColors(colors) {
      const where = 'Workbook.setThemeColors';
      requireObject(colors, `${where} argument 1`);
      this.#raw.setThemeColors(requireArray(colors.colors, `${where} colors`));
    }

    setThemeScriptFont(font) {
      const where = 'Workbook.setThemeScriptFont';
      requireObject(font, `${where} argument 1`);
      this.#raw.setThemeScriptFont(
        requireBoolean(font.major, `${where} major`),
        requireString(font.script, `${where} script`),
        requireString(font.typeface, `${where} typeface`),
      );
    }
  }

  // `using` calls the same function as `dispose()`.
  Object.defineProperty(Workbook.prototype, Symbol.dispose, {
    value: Workbook.prototype.dispose,
    writable: true,
    configurable: true,
  });
  const forWorkbook = [WORKBOOK_METHODS, WORKBOOK_ABSENT_ARGS, WORKBOOK_PATCHES];
  installForwarders(Workbook.prototype, 'Workbook', ...forWorkbook, Raw, (self, name, args) =>
    rawOf(self)[name](...args),
  );
  const forSheet = [SHEET_METHODS, SHEET_ABSENT_ARGS, SHEET_PATCHES];
  installForwarders(Sheet.prototype, 'Sheet', ...forSheet, Raw, (self, name, args) =>
    rawOf(self.workbook)[name](sheetIndex(self), ...args),
  );
  return Workbook;
}

/** Throws `TypeError: <owner>.<name> <field> must be a <type>` for a `null` on a non-nullable patch field. */
function refuseNullFields(owner, name, patch, fields) {
  if (!isObject(patch)) return;
  for (const [field, type] of Object.entries(fields)) {
    if (patch[field] === null) throw new TypeError(`${owner}.${name} ${field} must be a ${type}`);
  }
}

/**
 * Defines the generated forwarders on a facade prototype: each listed method the
 * prototype does not define by hand calls `call(this, rawName, args)`. An op the
 * ledger marks hand-written, or one that changes the sheet list, must already be
 * defined by hand. A list family loops over the raw count and element getter.
 * An `*_ABSENT_ARGS` entry maps `null` (and, for an optional argument,
 * `undefined`) at its position to the raw sentinel, and refuses a negative
 * number where that sentinel is negative; a `*_PATCHES` entry refuses `null`
 * on a non-nullable patch field. `owner` words the errors. A method
 * the raw class does not register (a member only the other package has) is
 * skipped.
 */
function installForwarders(proto, owner, names, absentArgs, patches, Raw, call) {
  for (const name of names) {
    if (Object.hasOwn(proto, name)) continue;
    if (HANDWRITTEN.includes(name) || SHEET_STRUCTURE_OPS.includes(name)) {
      throw new TypeError(`formulon: the facade does not implement the hand-written ${name}`);
    }
    const family = LIST_FAMILIES[name];
    if (family !== undefined) {
      const [count, element] = family;
      defineMethod(
        proto,
        name,
        namedFunction(name, function (args) {
          const n = call(this, count, args);
          const out = [];
          for (let i = 0; i < n; i++) out.push(call(this, element, [...args, i]));
          return out;
        }),
      );
      continue;
    }
    if (typeof Raw.prototype[name] !== 'function') continue;
    const patched = patches[name];
    if (patched !== undefined) {
      const [position, fields] = patched;
      defineMethod(
        proto,
        name,
        namedFunction(name, function (args) {
          refuseNullFields(owner, name, args[position], fields);
          return call(this, name, args);
        }),
      );
      continue;
    }
    const absent = absentArgs[name];
    defineMethod(
      proto,
      name,
      absent === undefined
        ? namedFunction(name, function (args) {
            return call(this, name, args);
          })
        : namedFunction(name, function (args) {
            const [position, sentinel, optional = false] = absent;
            const value = args[position];
            if (value === null || (optional && value === undefined)) {
              args = [...args];
              args[position] = sentinel;
            } else if (typeof sentinel === 'number' && sentinel < 0 && typeof value === 'number' && value < 0) {
              // The C sentinel is not a public value; absence is spelled `null`.
              throw new RangeError(`${owner}.${name} argument ${position + 1} is outside uint32 range`);
            }
            return call(this, name, args);
          }),
    );
  }
}
