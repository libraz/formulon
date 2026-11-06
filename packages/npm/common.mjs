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
export const RefMode = Object.freeze({ A1: 0, R1C1: 1 });
export const CollationStrength = Object.freeze({ Compare: 0, Dynamic: 1, ClassicExact: 2 });
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
export const ShapePaint = Object.freeze({ Unset: 0, Solid: 1, None: 2, Other: 3 });
export const SlicerSortOrder = Object.freeze({ Ascending: 0, Descending: 1 });
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
export const ChartBlanksAs = Object.freeze({ Absent: 0, Gap: 1, Zero: 2, Span: 3 });
export const ChartPointState = Object.freeze({ Value: 0, Gap: 1, Connect: 2 });
export const ChartAxisKind = Object.freeze({ Category: 0, Value: 1, Date: 2, Series: 3 });
export const ChartAxisCrosses = Object.freeze({ Absent: 0, AutoZero: 1, Min: 2, Max: 3, At: 4 });
export const ChartTitleKind = Object.freeze({ None: 0, Text: 1, Ref: 2, Auto: 3 });
export const ChartSeriesDirection = Object.freeze({ Auto: 0, Columns: 1, Rows: 2 });
export const ChartLegendPosition = Object.freeze({ None: 0, Bottom: 1, Top: 2, Left: 3, Right: 4, TopRight: 5 });
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
// other member comes from this module.
const RUNTIME_MEMBERS = new Set([
  'Workbook',
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
    const value = RUNTIME_MEMBERS.has(name) ? runtime[name] : shared[name];
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
  installWorkbookExtensions(module.Workbook);
  return buildFormulon(module, shared);
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
function cellAddress(workbook, a1) {
  const ref = workbook.parseReference(a1);
  if (ref === null || ref.sheetName !== null) throw invalidAddress(a1);
  const { rect } = ref;
  if (rect.firstRow !== rect.lastRow || rect.firstCol !== rect.lastCol) throw invalidAddress(a1);
  return rect;
}

/** A1 view of one worksheet; see `Sheet` in formulon.d.ts. */
class Sheet {
  #workbook;
  #index;

  constructor(workbook, index) {
    this.#workbook = workbook;
    this.#index = index;
  }

  get workbook() {
    return this.#workbook;
  }

  get index() {
    return this.#index;
  }

  get name() {
    return this.#workbook.getSheetName(this.#index);
  }

  get(a1) {
    const cell = cellAddress(this.#workbook, a1);
    return this.#workbook.getValue(this.#index, cell.firstRow, cell.firstCol);
  }

  set(a1, value) {
    const cell = cellAddress(this.#workbook, a1);
    this.#workbook.setValue(this.#index, cell.firstRow, cell.firstCol, value);
  }

  getFormula(a1) {
    const cell = cellAddress(this.#workbook, a1);
    return this.#workbook.getFormula(this.#index, cell.firstRow, cell.firstCol);
  }

  setFormula(a1, formula) {
    const cell = cellAddress(this.#workbook, a1);
    this.#workbook.setFormula(this.#index, cell.firstRow, cell.firstCol, formula);
  }
}

function requireField(value, field, type) {
  if (typeof value[field] !== type) {
    throw new TypeError(`setValue argument 4: a Value of this kind needs a ${type} \`${field}\``);
  }
  return value[field];
}

function setValue(sheet, row, col, value) {
  if (value === null) return this.setBlank(sheet, row, col);
  switch (typeof value) {
    case 'number':
      return this.setNumber(sheet, row, col, value);
    case 'string':
      return this.setText(sheet, row, col, value);
    case 'boolean':
      return this.setBool(sheet, row, col, value);
    case 'object':
      if (Array.isArray(value)) break;
      switch (value.kind) {
        case ValueKind.Blank:
          return this.setBlank(sheet, row, col);
        case ValueKind.Number:
          return this.setNumber(sheet, row, col, requireField(value, 'number', 'number'));
        case ValueKind.Bool:
          return this.setBool(sheet, row, col, requireField(value, 'boolean', 'boolean'));
        case ValueKind.Text:
          return this.setText(sheet, row, col, requireField(value, 'text', 'string'));
        case ValueKind.Error:
          return this.setError(sheet, row, col, requireField(value, 'errorCode', 'number'));
      }
      break;
  }
  throw new TypeError('setValue argument 4 must be a number, string, boolean, null or a scalar Value');
}

function sheet(nameOrIndex) {
  if (typeof nameOrIndex === 'string') {
    const index = this.getSheetIndex(nameOrIndex);
    if (index === null) {
      throw new FormulonError(
        StatusCode.NotFound,
        `no sheet is named ${JSON.stringify(nameOrIndex)}`,
        `sheet_name=${nameOrIndex}`,
      );
    }
    return new Sheet(this, index);
  }
  if (typeof nameOrIndex === 'number') {
    // Validates the index exactly as every other sheet-index argument is validated.
    this.getSheetName(nameOrIndex);
    return new Sheet(this, nameOrIndex);
  }
  throw new TypeError('sheet argument 1 must be a sheet name or index');
}

/**
 * Adds the pure-JS `Workbook` methods (`setValue`, `sheet`) to a runtime's
 * Workbook class. Both packages share this one implementation.
 *
 * @param {Function} WorkbookCtor The runtime's `Workbook` class.
 */
export function installWorkbookExtensions(WorkbookCtor) {
  for (const [name, fn] of [
    ['setValue', setValue],
    ['sheet', sheet],
  ]) {
    Object.defineProperty(WorkbookCtor.prototype, name, { value: fn, writable: true, configurable: true });
  }
}
