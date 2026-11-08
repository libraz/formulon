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
export const CfRuleType = Object.freeze({
  Expression: 0,
  CellIs: 1,
  ColorScale: 2,
  DataBar: 3,
  IconSet: 4,
  Top10: 5,
  AboveAverage: 6,
  ContainsText: 7,
  NotContainsText: 8,
  BeginsWith: 9,
  EndsWith: 10,
  ContainsBlanks: 11,
  NotContainsBlanks: 12,
  ContainsErrors: 13,
  NotContainsErrors: 14,
  TimePeriod: 15,
  DuplicateValues: 16,
  UniqueValues: 17,
});
export const CfOperator = Object.freeze({
  LessThan: 0,
  LessThanOrEqual: 1,
  Equal: 2,
  NotEqual: 3,
  GreaterThanOrEqual: 4,
  GreaterThan: 5,
  Between: 6,
  NotBetween: 7,
});
export const CfTimePeriod = Object.freeze({
  Today: 0,
  Yesterday: 1,
  Tomorrow: 2,
  Last7Days: 3,
  ThisWeek: 4,
  LastWeek: 5,
  NextWeek: 6,
  ThisMonth: 7,
  LastMonth: 8,
  NextMonth: 9,
});
export const CfValueObjectType = Object.freeze({
  Number: 0,
  Percent: 1,
  Percentile: 2,
  Min: 3,
  Max: 4,
  Formula: 5,
  AutoMin: 6,
  AutoMax: 7,
});
export const IconSetName = Object.freeze({
  ThreeArrows: 0,
  ThreeArrowsGray: 1,
  ThreeFlags: 2,
  ThreeTrafficLights1: 3,
  ThreeTrafficLights2: 4,
  ThreeSigns: 5,
  ThreeSymbols: 6,
  ThreeSymbols2: 7,
  FourArrows: 8,
  FourArrowsGray: 9,
  FourRedToBlack: 10,
  FourRating: 11,
  FourTrafficLights: 12,
  FiveArrows: 13,
  FiveArrowsGray: 14,
  FiveRating: 15,
  FiveQuarters: 16,
});
export const ValidationType = Object.freeze({
  None: 0,
  Whole: 1,
  Decimal: 2,
  List: 3,
  Date: 4,
  Time: 5,
  TextLength: 6,
  Custom: 7,
});
export const ValidationOperator = Object.freeze({
  Between: 0,
  NotBetween: 1,
  Equal: 2,
  NotEqual: 3,
  GreaterThan: 4,
  LessThan: 5,
  GreaterThanOrEqual: 6,
  LessThanOrEqual: 7,
});
export const ParseErrorCode = Object.freeze({
  LexerInvalidCharacter: 0,
  LexerUnterminatedString: 1,
  LexerUnterminatedSheetQuote: 2,
  LexerInvalidNumberLiteral: 3,
  LexerInvalidErrorLiteral: 4,
  LexerInvalidEscape: 5,
  LexerInvalidReference: 6,
  LexerExcessiveLength: 7,
  UnexpectedToken: 8,
  UnexpectedEof: 9,
  ExpectedExpression: 10,
  ExpectedCloseParen: 11,
  UnbalancedBraces: 12,
  ArrayRowMismatch: 13,
  ExpectedRparenOrComma: 14,
  ExpectedCommaOrSemiInArray: 15,
  InvalidReference: 16,
  UnsupportedConstruct: 17,
  ExpectedOpenParen: 18,
  ExpectedComma: 19,
  UnbalancedBrackets: 20,
  InvalidRange: 21,
  NestedFormulaTooDeep: 22,
  TooManyErrors: 23,
  LetInvalidName: 24,
  LetWrongArity: 25,
  LambdaInvalidParam: 26,
  LambdaEmpty: 27,
  LambdaDuplicateParam: 28,
  EntryWrongArity: 29,
  EntryReferenceRequired: 30,
  EntryQualifiedBuiltin: 31,
  EntryRefusedWhitespace: 32,
  EntryParamShapedName: 33,
  EntryDoubleEquals: 34,
  EntryBomNotSupported: 35,
  StructuredRefInvalid: 36,
});
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
// <error-display-names:begin> generated by tools/codegen/gen_error_display_names.py; do not edit
export const ErrorDisplayName = Object.freeze({
  [ErrorCode.Null]: '#NULL!',
  [ErrorCode.Div0]: '#DIV/0!',
  [ErrorCode.Value]: '#VALUE!',
  [ErrorCode.Ref]: '#REF!',
  [ErrorCode.Name]: '#NAME?',
  [ErrorCode.Num]: '#NUM!',
  [ErrorCode.NA]: '#N/A',
  [ErrorCode.GettingData]: '#GETTING_DATA',
  [ErrorCode.Spill]: '#SPILL!',
  [ErrorCode.Calc]: '#CALC!',
  [ErrorCode.Field]: '#FIELD!',
  [ErrorCode.Blocked]: '#BLOCKED!',
  [ErrorCode.Connect]: '#CONNECT!',
  [ErrorCode.External]: '#EXTERNAL!',
  [ErrorCode.Busy]: '#BUSY!',
  [ErrorCode.Python]: '#PYTHON!',
  [ErrorCode.Unknown]: '#UNKNOWN!',
});
// <error-display-names:end>
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
export const HorizontalAlign = Object.freeze({
  General: 0,
  Left: 1,
  Center: 2,
  Right: 3,
  Fill: 4,
  Justify: 5,
  CenterContinuous: 6,
  Distributed: 7,
});
export const VerticalAlign = Object.freeze({ Top: 0, Center: 1, Bottom: 2, Justify: 3, Distributed: 4 });
export const ReadingOrder = Object.freeze({ Context: 0, LeftToRight: 1, RightToLeft: 2 });
export const UnderlineStyle = Object.freeze({
  None: 0,
  Single: 1,
  Double: 2,
  SingleAccounting: 3,
  DoubleAccounting: 4,
});
export const FontVerticalAlign = Object.freeze({ Baseline: 0, Superscript: 1, Subscript: 2 });
export const FontScheme = Object.freeze({ None: 0, Major: 1, Minor: 2 });
export const FontFamily = Object.freeze({
  NotApplicable: 0,
  Roman: 1,
  Swiss: 2,
  Modern: 3,
  Script: 4,
  Decorative: 5,
});
export const FillPattern = Object.freeze({
  None: 0,
  Solid: 1,
  MediumGray: 2,
  DarkGray: 3,
  LightGray: 4,
  DarkHorizontal: 5,
  DarkVertical: 6,
  DarkDown: 7,
  DarkUp: 8,
  DarkGrid: 9,
  DarkTrellis: 10,
  LightHorizontal: 11,
  LightVertical: 12,
  LightDown: 13,
  LightUp: 14,
  LightGrid: 15,
  LightTrellis: 16,
  Gray125: 17,
  Gray0625: 18,
});
export const GradientType = Object.freeze({ Linear: 0, Path: 1 });
export const BorderStyle = Object.freeze({
  None: 0,
  Thin: 1,
  Medium: 2,
  Dashed: 3,
  Dotted: 4,
  Thick: 5,
  Double: 6,
  Hair: 7,
  MediumDashed: 8,
  DashDot: 9,
  MediumDashDot: 10,
  DashDotDot: 11,
  MediumDashDotDot: 12,
  SlantDashDot: 13,
});
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
export const ChartSourceKind = Object.freeze({ None: 0, Ref: 1, Literal: 2 });
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
 * An Excel cell error such as `#DIV/0!`, as the native reads return it. `code` is an `ErrorCode`; `toString()`
 * spells it as `ErrorDisplayName` does.
 */
export class ExcelError {
  /** @param {number} code An `ErrorCode`. */
  constructor(code) {
    if (typeof code !== 'number') throw new TypeError('ExcelError argument 1 must be an ErrorCode');
    if (ErrorDisplayName[code] === undefined) throw new RangeError('ExcelError argument 1 must be an ErrorCode');
    this.code = code;
    Object.freeze(this);
  }

  toString() {
    return ErrorDisplayName[this.code];
  }
}

/**
 * The key `value` has in the ordinal table `enumType`; `null` when no key has it.
 *
 * @param {object} enumType A frozen ordinal table such as `CfRuleType`.
 * @param {number} value The ordinal.
 * @returns {string | null}
 */
export function enumName(enumType, value) {
  if (!isObject(enumType)) throw new TypeError('enumName argument 1 must be an ordinal table');
  for (const [key, ordinal] of Object.entries(enumType)) if (ordinal === value) return key;
  return null;
}

const NOT_A_VALUE = 'toNative argument 1 must be a Value';

/**
 * Maps a `Value` to its native JS value: Blank to `null`, Number to a number, Bool to a boolean, Text to a
 * string, Error to an `ExcelError` and an `ArrayValue` to its rows. A `RefValue` or `LambdaValue`, which no
 * read returns, and anything that is not a `Value` throw `TypeError`.
 *
 * @param {object} value A `Value`.
 * @returns {null | number | boolean | string | ExcelError | Array} The native value.
 */
export function toNative(value) {
  if (!isObject(value)) throw new TypeError(NOT_A_VALUE);
  switch (value.kind) {
    case ValueKind.Blank:
      return null;
    case ValueKind.Array: {
      const { rows, cols, cells } = value;
      if (!Number.isInteger(rows) || !Number.isInteger(cols) || !Array.isArray(cells)) throw new TypeError(NOT_A_VALUE);
      if (cells.length !== rows * cols) throw new TypeError(NOT_A_VALUE);
      return Array.from({ length: rows }, (_, r) => cells.slice(r * cols, (r + 1) * cols).map(toNative));
    }
  }
  const [field, type] = SCALAR_VALUE_FIELDS[value.kind] ?? [];
  if (field === undefined || typeof value[field] !== type) throw new TypeError(NOT_A_VALUE);
  return value.kind === ValueKind.Error ? new ExcelError(value.errorCode) : value[field];
}

/**
 * The error every host-side failure throws: an engine call that did not
 * complete, or a call on a disposed workbook. Failures are told apart by
 * `code`, a `StatusCode`; there are no subclasses. A formula that does not
 * parse carries where in `parseError`.
 *
 * `instanceof` holds only for the class of the package that threw, so code
 * handling errors from both packages tests `e.name === 'FormulonError'`.
 */
export class FormulonError extends Error {
  /**
   * @param {number} code A `StatusCode`.
   * @param {string} message What went wrong, in English.
   * @param {string} [context] Structured `key=value` detail; `''` when none.
   * @param {object} [parseError] `{code, start, end, line, column}` of a formula syntax error.
   */
  constructor(code, message, context = '', parseError = undefined) {
    super(message);
    this.name = 'FormulonError';
    this.code = code;
    this.context = context;
    this.parseError = parseError === undefined || parseError === null ? undefined : Object.freeze({ ...parseError });
  }

  toString() {
    return `FormulonError [${enumName(StatusCode, this.code) ?? this.code}]: ${this.message}`;
  }
}

/**
 * Member names of the frozen object `createFormulon()` resolves to, the same
 * on both packages.
 */
export const FORMULON_MEMBERS = Object.freeze([
  'Workbook',
  'FormulonError',
  'ExcelError',
  'ValueKind',
  'StatusCode',
  'ParseErrorCode',
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
  'CfRuleType',
  'CfOperator',
  'CfTimePeriod',
  'CfValueObjectType',
  'IconSetName',
  'ErrorCode',
  'ErrorDisplayName',
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
  'ValidationType',
  'ValidationOperator',
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
  'ChartSourceKind',
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
  'HorizontalAlign',
  'VerticalAlign',
  'ReadingOrder',
  'UnderlineStyle',
  'FontVerticalAlign',
  'FontScheme',
  'FontFamily',
  'FillPattern',
  'GradientType',
  'BorderStyle',
  'FunctionAvailability',
  'PIVOT_SHOW_AS_BASE_PREVIOUS',
  'PIVOT_SHOW_AS_BASE_NEXT',
  'PIVOT_VALUES_FIELD',
  'mergeFunctionMetadata',
  'toNative',
  'enumName',
  'evaluateFormula',
  'evaluateTypedFormula',
  'convertCivilToSerial',
  'convertSerialToCivil',
  'getVersion',
  'getStatusName',
  'getFunctionMetadata',
  'listFunctionNames',
  'localizeFunctionName',
  'canonicalizeFunctionName',
  'localizeFormula',
  'canonicalizeFormula',
  'getLocaleFacts',
  'listLocaleErrorNames',
  'compareText',
  'getCollationClasses',
  'getImageInfo',
  'setLogMinLevel',
  'setLogSink',
]);

// Members the engine runtime (Emscripten module or addon) provides; every
// other member comes from this module. The runtime's raw `Workbook` class is
// not one of them: `buildFormulon` wraps it with `makeFacade`.
const RUNTIME_MEMBERS = new Set([
  'getVersion',
  'getStatusName',
  'setLogMinLevel',
  'setLogSink',
  'getFunctionMetadata',
  'listFunctionNames',
  'localizeFunctionName',
  'canonicalizeFunctionName',
  'localizeFormula',
  'canonicalizeFormula',
  'getLocaleFacts',
  'listLocaleErrorNames',
  'compareText',
  'getCollationClasses',
  'getImageInfo',
]);

const FORWARDED_OPTIONS = ['locateFile', 'print', 'printErr'];

function int32Field(value, where) {
  requireInteger(value, where);
  if (value > INT32_MAX || value < -INT32_MAX - 1) throw new RangeError(`${where} is outside int32 range`);
  return value;
}

function dateSystem(date1904, where) {
  return date1904 === undefined ? false : requireBoolean(date1904, `${where} argument 2`);
}

const CIVIL_FIELDS = Object.freeze(['year', 'month', 'day', 'hour', 'minute', 'second']);
const CIVIL_DATE_FIELDS = Object.freeze(['year', 'month', 'day']);

/** The tagged form of a raw `{rows, cols, cells}` result (`cells` as rows): a 1x1 result is its one value. */
function typedResult(result) {
  if (result.rows === 1 && result.cols === 1) return result.cells[0][0];
  return { kind: ValueKind.Array, rows: result.rows, cols: result.cols, cells: result.cells.flat() };
}

/** A string formula argument; anything else is a `TypeError` naming `where`. */
function formulaText(formula, where) {
  if (typeof formula !== 'string') throw new TypeError(`${where} must be a primitive string`);
  return formula;
}

// Members built over a runtime function, keyed by member name: [the runtime function, the wrapper factory]. The
// civil-date wrappers default `date1904` and the time of day and refuse what the fields cannot represent; the
// engine judges the calendar. The evaluation wrappers fold the raw `{rows, cols, cells}` result.
const RUNTIME_WRAPPERS = Object.freeze({
  convertCivilToSerial: [
    'convertCivilToSerial',
    (convert) =>
      function convertCivilToSerial(civil, date1904) {
        const where = 'Formulon.convertCivilToSerial';
        requireObject(civil, `${where} argument 1`);
        const fields = {};
        for (const field of CIVIL_FIELDS) {
          const value = civil[field];
          fields[field] =
            value === undefined && !CIVIL_DATE_FIELDS.includes(field) ? 0 : int32Field(value, `${where} ${field}`);
        }
        return convert(fields, dateSystem(date1904, where));
      },
  ],
  convertSerialToCivil: [
    'convertSerialToCivil',
    (convert) =>
      function convertSerialToCivil(serial, date1904) {
        const where = 'Formulon.convertSerialToCivil';
        requireNumber(serial, `${where} argument 1`);
        if (!Number.isFinite(serial)) throw new RangeError(`${where} argument 1 must be finite`);
        return convert(serial, dateSystem(date1904, where));
      },
  ],
  evaluateFormula: [
    'evaluateFormula',
    (evaluate) =>
      function evaluateFormula(formula) {
        return toNative(typedResult(evaluate(formulaText(formula, 'Formulon.evaluateFormula argument 1'))));
      },
  ],
  evaluateTypedFormula: [
    'evaluateFormula',
    (evaluate) =>
      function evaluateTypedFormula(formula) {
        return typedResult(evaluate(formulaText(formula, 'Formulon.evaluateTypedFormula argument 1')));
      },
  ],
});

/**
 * Builds the frozen `Formulon` object from an engine runtime.
 *
 * @param {object} runtime The Emscripten module or the native addon.
 * @param {object} shared This module's namespace object.
 * @param {FinalizationRegistry} [registry] Where each workbook registers its raw object; `dispose()`
 *   unregisters it and deletes the raw object, so only a runtime whose raw objects have `delete()` passes one.
 * @returns {object} The object `createFormulon()` resolves to.
 */
export function buildFormulon(runtime, shared, registry) {
  const formulon = {};
  for (const name of FORMULON_MEMBERS) {
    let value;
    if (name === 'Workbook') {
      value = runtime.Workbook === undefined ? undefined : makeFacade(runtime.Workbook, registry, runtime);
    } else if (Object.hasOwn(RUNTIME_WRAPPERS, name)) {
      const [source, wrap] = RUNTIME_WRAPPERS[name];
      value = typeof runtime[source] === 'function' ? wrap(runtime[source]) : undefined;
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

function isBufferSource(value) {
  return ArrayBuffer.isView(value) || value instanceof ArrayBuffer;
}

/**
 * The factory options `FormulonOptions.wasm` adds: a `BufferSource` as `wasmBinary`, a `WebAssembly.Module`
 * as an `instantiateWasm` that instantiates it, whose failure `fail` reports.
 */
function wasmOptions(wasm, fail) {
  if (wasm === undefined || wasm === null) return {};
  if (typeof WebAssembly === 'object' && wasm instanceof WebAssembly.Module) {
    return {
      instantiateWasm(imports, receive) {
        WebAssembly.instantiate(wasm, imports).then((instance) => receive(instance, wasm), fail);
        return {};
      },
    };
  }
  if (!isBufferSource(wasm)) throw new TypeError('FormulonOptions.wasm must be a BufferSource or a WebAssembly.Module');
  if (wasm.byteLength === 0) throw new TypeError('FormulonOptions.wasm must not be empty');
  return { wasmBinary: wasm };
}

/**
 * Instantiates an Emscripten module factory and builds the `Formulon` object.
 * Only `locateFile`, `print`, `printErr` and `wasm` reach the factory.
 *
 * @param {Function} factory The Emscripten module factory.
 * @param {object} [options] `FormulonOptions`.
 * @param {object} shared This module's namespace object.
 * @param {FinalizationRegistry} [registry] See `buildFormulon`.
 * @returns {Promise<object>} The `Formulon` object.
 */
export async function loadFormulon(factory, options, shared, registry) {
  const moduleOptions = { FormulonError };
  // The glue awaits `instantiateWasm` without a rejection path, so a failed instantiation settles here.
  let fail;
  const failed = new Promise((_resolve, reject) => {
    fail = reject;
  });
  if (options !== undefined && options !== null) {
    for (const key of FORWARDED_OPTIONS) {
      if (options[key] !== undefined) moduleOptions[key] = options[key];
    }
    Object.assign(moduleOptions, wasmOptions(options.wasm, fail));
  }
  const module = await Promise.race([factory(moduleOptions), failed]);
  return buildFormulon(module, shared, registry);
}

// <facade-tables>
const WORKBOOK_METHODS = Object.freeze([
  '[Symbol.iterator]',
  'addPerson',
  'addPivotCache',
  'addSheet',
  'addTable',
  'clearTheme',
  'copySheet',
  'dispose',
  'disposed',
  'formatReference',
  'formatValue',
  'getActiveSheet',
  'getCalcMode',
  'getCalcOnSave',
  'getDate1904',
  'getDefinedName',
  'getDiagnostics',
  'getFileName',
  'getFullPrecision',
  'getIterative',
  'getMemoryUsage',
  'getPerson',
  'getPinnedNow',
  'getProfileId',
  'getRefMode',
  'getTableStyle',
  'getTableStyleDefaults',
  'getThemeColors',
  'getThemeFonts',
  'getThemeScriptFont',
  'getWorkbookProtection',
  'hasDirtyCells',
  'listDefinedNames',
  'listExternalLinks',
  'listPeople',
  'listPivotCaches',
  'listSheets',
  'listTableStyles',
  'listTables',
  'parseReference',
  'pivotCache',
  'recalc',
  'recalcDataTables',
  'recalcParallel',
  'removeDefinedName',
  'removePerson',
  'removePivotCache',
  'removeSheet',
  'removeTable',
  'removeTableStyle',
  'resolveColor',
  'save',
  'saveWithDiagnostics',
  'setActiveSheet',
  'setCalcMode',
  'setCalcOnSave',
  'setDate1904',
  'setDefaultFont',
  'setDefinedName',
  'setFileName',
  'setFullPrecision',
  'setIterative',
  'setIterativeProgress',
  'setPinnedNow',
  'setProfileId',
  'setRefMode',
  'setTableStyle',
  'setTableStyleDefaults',
  'setThemeColors',
  'setThemeFonts',
  'setThemeScriptFont',
  'setWorkbookProtection',
  'sheet',
  'table',
]);
const SHEET_METHODS = Object.freeze([
  'addCfRule',
  'addChart',
  'addColumnBreak',
  'addDataTable',
  'addHyperlink',
  'addHyperlinkRange',
  'addMerge',
  'addPicture',
  'addPivot',
  'addProtectedRange',
  'addRowBreak',
  'addShape',
  'addSlicer',
  'addSparklineGroup',
  'addThreadedComment',
  'addValidation',
  'applyAutoFilter',
  'chart',
  'clearAutoFilter',
  'clearBreaks',
  'clearCfRules',
  'clearColumnWidth',
  'clearHyperlinks',
  'clearMerges',
  'clearRowHeight',
  'clearValidations',
  'convertColumnCharsToPt',
  'convertColumnPtToChars',
  'deleteColumns',
  'deleteRows',
  'drawing',
  'evaluateAutoFilter',
  'evaluateCfFormula',
  'evaluateCfRange',
  'evaluateFormula',
  'evaluateFormulaArray',
  'evaluateTypedFormula',
  'evaluateTypedFormulaArray',
  'getArrayFormula',
  'getAutoFilter',
  'getCellFormat',
  'getCellPhonetic',
  'getCellPhoneticProperties',
  'getCellPhoneticRuns',
  'getCellRectPt',
  'getCellShowPhonetic',
  'getCellsInRange',
  'getColumnWidthPt',
  'getComment',
  'getCommentVisible',
  'getDataTable',
  'getDate',
  'getDependents',
  'getDisplayText',
  'getEffectiveStyle',
  'getFormatDefaults',
  'getFormula',
  'getFormulaR1C1',
  'getHeaderFooter',
  'getHyperlink',
  'getInvalidCells',
  'getLambdaText',
  'getMergesInRange',
  'getPageMargins',
  'getPageSetup',
  'getPagination',
  'getPrecedents',
  'getPrintArea',
  'getPrintOptions',
  'getPrintTitles',
  'getRowHeightPt',
  'getSelection',
  'getSheetProtection',
  'getSpillInfo',
  'getTabColor',
  'getThreadedComment',
  'getTypedValue',
  'getTypedValues',
  'getUsedRange',
  'getValue',
  'getValues',
  'getView',
  'getWidthModel',
  'insertColumns',
  'insertRows',
  'iterRows',
  'listCells',
  'listCfRules',
  'listCharts',
  'listColumnBreaks',
  'listColumnOverrides',
  'listComments',
  'listDrawings',
  'listHyperlinks',
  'listMerges',
  'listPictures',
  'listPivots',
  'listProtectedRanges',
  'listRowBreaks',
  'listRowOverrides',
  'listShapes',
  'listSlicers',
  'listSparklineGroups',
  'listThreadedComments',
  'listValidations',
  'picture',
  'pivot',
  'recalcViewport',
  'removeAutoFilter',
  'removeCfRule',
  'removeColumnBreak',
  'removeComment',
  'removeDataTable',
  'removeDrawing',
  'removeHyperlink',
  'removeMerge',
  'removePivot',
  'removeProtectedRange',
  'removeRowBreak',
  'removeSparklineGroup',
  'removeThreadedComment',
  'removeValidation',
  'restoreDrawing',
  'restoreRangeStyle',
  'setArrayFormula',
  'setAutoFilter',
  'setBlank',
  'setBool',
  'setCellFormat',
  'setCellPhonetic',
  'setCellPhoneticProperties',
  'setCellPhoneticRuns',
  'setCellShowPhonetic',
  'setColumnCollapsed',
  'setColumnHidden',
  'setColumnOutline',
  'setColumnWidth',
  'setComment',
  'setCommentVisible',
  'setError',
  'setFormatDefaults',
  'setFormula',
  'setHeaderFooter',
  'setNumber',
  'setPageMargins',
  'setPageSetup',
  'setPrintArea',
  'setPrintOptions',
  'setPrintTitles',
  'setRowAutoHeight',
  'setRowCollapsed',
  'setRowHeight',
  'setRowHidden',
  'setRowOutline',
  'setSelection',
  'setSheetProtection',
  'setTabColor',
  'setText',
  'setValue',
  'setValues',
  'setView',
  'shape',
  'slicer',
  'snapshotRangeStyle',
  'updateProtectedRange',
  'updateSparklineGroup',
  'updateThreadedComment',
  'validateValue',
]);
const OOXML_WORKBOOK_METHODS = Object.freeze([
  'addBorder',
  'addCellStyleXf',
  'addCellXf',
  'addDxf',
  'addFill',
  'addFont',
  'addNumFmt',
  'getBorder',
  'getCellStyle',
  'getCellStyleXf',
  'getCellXf',
  'getDxf',
  'getFill',
  'getFont',
  'getNumFmtString',
  'getPassthrough',
  'listBorders',
  'listCellStyleXfs',
  'listCellStyles',
  'listCellXfs',
  'listDxfs',
  'listFills',
  'listFonts',
  'listPassthroughs',
  'removeCellStyle',
  'setCellStyle',
  'setFont',
]);
const OOXML_SHEET_METHODS = Object.freeze([
  'clearCellXfIndex',
  'clearColumnStyle',
  'clearRowStyle',
  'getAutoFilterXml',
  'getCellXfIndex',
  'getColumnStyle',
  'getHeaderFooterXml',
  'getPageMarginsXml',
  'getPageSetupXml',
  'getPrintOptionsXml',
  'getRangeStyleBases',
  'getRangeStyleBasesExcluding',
  'getRowStyle',
  'getSheetPrXml',
  'remapRangeStyle',
  'setAutoFilterXml',
  'setCellXfIndex',
  'setColumnStyle',
  'setHeaderFooterXml',
  'setPageMarginsXml',
  'setPageSetupXml',
  'setPrintOptionsXml',
  'setRangeXfIndex',
  'setRowStyle',
  'setSheetPrXml',
]);
const SHEET_STRUCTURE_OPS = Object.freeze([
  'Pivot.showDetails',
  'Sheet.move',
  'Sheet.rename',
  'Workbook.addSheet',
  'Workbook.copySheet',
  'Workbook.removeSheet',
]);
const LIST_FAMILIES = Object.freeze({
  'Pivot.listDataFields': ['getPivotDataFieldCount', 'getPivotDataField'],
  'Pivot.listFieldItems': ['getPivotFieldItemCount', 'getPivotFieldItem'],
  'Pivot.listFields': ['getPivotFieldCount', 'getPivotField'],
  'Pivot.listFilters': ['getPivotFilterCount', 'getPivotFilter'],
  'PivotCache.listCalculatedItems': ['getPivotCacheCalculatedItemCount', 'getPivotCacheCalculatedItem'],
  'Sheet.listCfRules': ['getCfRuleCount', 'getCfRule'],
  'Sheet.listProtectedRanges': ['getProtectedRangeCount', 'getProtectedRange'],
  'Sheet.listSparklineGroups': ['getSparklineGroupCount', 'getSparklineGroup'],
  'Sheet.listValidations': ['getValidationCount', 'getValidation'],
  'Workbook.listBorders': ['getBorderCount', 'getBorder'],
  'Workbook.listCellStyleXfs': ['getCellStyleXfCount', 'getCellStyleXf'],
  'Workbook.listCellXfs': ['getCellXfCount', 'getCellXf'],
  'Workbook.listDxfs': ['getDxfCount', 'getDxf'],
  'Workbook.listExternalLinks': ['getExternalLinkCount', 'getExternalLink'],
  'Workbook.listFills': ['getFillCount', 'getFill'],
  'Workbook.listFonts': ['getFontCount', 'getFont'],
});
const HANDWRITTEN = Object.freeze([
  '[Symbol.iterator]',
  'addDataTable',
  'addHyperlink',
  'addHyperlinkRange',
  'addPicture',
  'addSheet',
  'chart',
  'copySheet',
  'drawing',
  'evaluateTypedFormula',
  'evaluateTypedFormulaArray',
  'getActiveSheet',
  'getCellFormat',
  'getCellPhoneticRuns',
  'getColumnStyle',
  'getDate',
  'getFormula',
  'getRowStyle',
  'getThemeColors',
  'getThemeFonts',
  'getTypedValue',
  'getTypedValues',
  'getUsedRange',
  'getValues',
  'iterRows',
  'listCells',
  'listCharts',
  'listDrawings',
  'listPictures',
  'listPivotCaches',
  'listPivots',
  'listShapes',
  'listSheets',
  'listSlicers',
  'listTables',
  'picture',
  'pivot',
  'pivotCache',
  'removeDrawing',
  'removeSheet',
  'setActiveSheet',
  'setArrayFormula',
  'setCellFormat',
  'setComment',
  'setFormula',
  'setPinnedNow',
  'setThemeColors',
  'setThemeScriptFont',
  'setValue',
  'setValues',
  'setView',
  'shape',
  'slicer',
  'table',
  'updateProtectedRange',
  'updateSparklineGroup',
  'updateThreadedComment',
]);
const HANDLE_FORWARDERS = Object.freeze({
  'Chart.addSeries': 'addChartSeries',
  'Chart.getData': 'getChartData',
  'Chart.getPreview': 'getChartPreview',
  'Chart.getProperties': 'getChart',
  'Chart.removeSeries': 'removeChartSeries',
  'Chart.rename': 'setDrawingName',
  'Chart.setAnchor': 'setDrawingAnchor',
  'Chart.setAxis': 'setChartAxis',
  'Chart.setDisplay': 'setChartDisplay',
  'Chart.setFilter': 'setChartFilter',
  'Chart.setLegend': 'setChartLegend',
  'Chart.setPlot': 'setChartPlot',
  'Chart.setPoint': 'setChartPoint',
  'Chart.setSeries': 'setChartSeries',
  'Chart.setSource': null,
  'Chart.setTitle': 'setChartTitle',
  'Chart.setType': 'setChartType',
  'Chart.setZOrder': 'setDrawingZOrder',
  'Chart.snapshot': 'snapshotDrawing',
  'Drawing.getDrawingProperties': 'getDrawingObject',
  'Picture.getProperties': 'getImage',
  'Picture.rename': 'setDrawingName',
  'Picture.setAnchor': 'setDrawingAnchor',
  'Picture.setZOrder': 'setDrawingZOrder',
  'Picture.snapshot': 'snapshotDrawing',
  'Pivot.addDataField': 'addPivotDataField',
  'Pivot.addField': 'addPivotField',
  'Pivot.addFieldItem': 'addPivotFieldItem',
  'Pivot.addFieldSharedItem': null,
  'Pivot.addFieldSubtotalFn': 'addPivotFieldSubtotalFn',
  'Pivot.addFilter': 'addPivotFilter',
  'Pivot.clearDataFields': 'clearPivotDataFields',
  'Pivot.clearFieldDateGroup': 'clearPivotFieldDateGroup',
  'Pivot.clearFieldItems': 'clearPivotFieldItems',
  'Pivot.clearFieldSubtotalFns': 'clearPivotFieldSubtotalFns',
  'Pivot.clearFields': 'clearPivotFields',
  'Pivot.clearFilters': 'clearPivotFilters',
  'Pivot.getCache': 'getPivotCacheId',
  'Pivot.getColumnFieldOrder': 'getPivotColumnFieldOrder',
  'Pivot.getDataFieldDefaultName': 'getPivotDataFieldDefaultName',
  'Pivot.getFieldItem': 'getPivotFieldItem',
  'Pivot.getGrandTotals': 'getPivotGrandTotals',
  'Pivot.getLayout': 'getPivotLayout',
  'Pivot.getOptions': 'getPivotOptions',
  'Pivot.getProperties': null,
  'Pivot.getReportLayout': 'getPivotReportLayout',
  'Pivot.getRowFieldOrder': 'getPivotRowFieldOrder',
  'Pivot.getStyleInfo': 'getPivotStyleInfo',
  'Pivot.isProjectionPending': 'isPivotProjectionPending',
  'Pivot.listDataFields': null,
  'Pivot.listFieldItems': null,
  'Pivot.listFields': null,
  'Pivot.listFilters': null,
  'Pivot.project': 'projectPivot',
  'Pivot.removeFilter': 'removePivotFilter',
  'Pivot.rename': 'setPivotName',
  'Pivot.setAnchor': 'setPivotAnchor',
  'Pivot.setColumnFieldOrder': 'setPivotColumnFieldOrder',
  'Pivot.setDataField': 'setPivotDataField',
  'Pivot.setFieldAxis': 'setPivotFieldAxis',
  'Pivot.setFieldDateGroup': null,
  'Pivot.setFieldShowDetail': 'setPivotFieldShowDetail',
  'Pivot.setFieldSort': 'setPivotFieldSort',
  'Pivot.setGrandTotals': null,
  'Pivot.setOptions': 'setPivotOptions',
  'Pivot.setReportLayout': 'setPivotReportLayout',
  'Pivot.setRowFieldOrder': 'setPivotRowFieldOrder',
  'Pivot.setStyleInfo': 'setPivotStyleInfo',
  'Pivot.showDetails': null,
  'Pivot.updateField': null,
  'Pivot.updateItem': null,
  'PivotCache.addCalculatedField': 'addPivotCacheCalculatedField',
  'PivotCache.addCalculatedItem': 'addPivotCacheCalculatedItem',
  'PivotCache.addField': 'addPivotCacheField',
  'PivotCache.addFieldSharedItem': null,
  'PivotCache.addRecord': 'addPivotCacheRecord',
  'PivotCache.clearFieldSharedItems': 'clearPivotCacheFieldSharedItems',
  'PivotCache.clearFields': 'clearPivotCacheFields',
  'PivotCache.clearRecords': 'clearPivotCacheRecords',
  'PivotCache.getOptions': 'getPivotCacheOptions',
  'PivotCache.getRecordCount': 'getPivotCacheRecordCount',
  'PivotCache.getWorksheetSource': 'getPivotCacheWorksheetSource',
  'PivotCache.listCalculatedItems': null,
  'PivotCache.listFieldSharedItems': null,
  'PivotCache.listFields': null,
  'PivotCache.moveCalculatedItem': 'movePivotCacheCalculatedItem',
  'PivotCache.removeCalculatedField': 'removePivotCacheCalculatedField',
  'PivotCache.removeCalculatedItem': 'removePivotCacheCalculatedItem',
  'PivotCache.setCalculatedFieldFormula': 'setPivotCacheCalculatedFieldFormula',
  'PivotCache.setCalculatedItemFormula': 'setPivotCacheCalculatedItemFormula',
  'PivotCache.setOptions': 'setPivotCacheOptions',
  'PivotCache.setRecordValue': null,
  'PivotCache.setWorksheetSource': null,
  'Shape.getProperties': 'getShape',
  'Shape.rename': 'setDrawingName',
  'Shape.setZOrder': 'setDrawingZOrder',
  'Shape.snapshot': 'snapshotDrawing',
  'Shape.update': null,
  'Sheet.move': null,
  'Sheet.rename': null,
  'Slicer.getProperties': 'getSlicer',
  'Slicer.setZOrder': 'setDrawingZOrder',
  'Slicer.update': null,
  'Table.applyAutoFilter': 'applyTableAutoFilter',
  'Table.clearAutoFilter': 'clearTableAutoFilter',
  'Table.evaluateAutoFilter': 'evaluateTableAutoFilter',
  'Table.getAutoFilter': 'getTableAutoFilter',
  'Table.getProperties': 'getTable',
  'Table.removeAutoFilter': 'removeTableAutoFilter',
  'Table.renameColumn': 'renameTableColumn',
  'Table.setAutoFilter': 'setTableAutoFilter',
  'Table.setColumnTotals': 'setTableColumnTotals',
  'Table.update': null,
});
const HANDLE_PROPERTIES = Object.freeze({
  'Pivot.name': 'getPivotName',
});
const CELL_ARGS = Object.freeze({
  'Pivot.showDetails': [[0, 1]],
  'Sheet.addHyperlink': [[0, 1]],
  'Sheet.clearCellXfIndex': [[0, 1]],
  'Sheet.evaluateCfFormula': [
    [0, 1],
    [2, 3],
  ],
  'Sheet.evaluateFormula': [[0, 1]],
  'Sheet.evaluateFormulaArray': [[0, 1]],
  'Sheet.evaluateTypedFormula': [[0, 1]],
  'Sheet.evaluateTypedFormulaArray': [[0, 1]],
  'Sheet.getArrayFormula': [[0, 1]],
  'Sheet.getCellFormat': [[0, 1]],
  'Sheet.getCellPhonetic': [[0, 1]],
  'Sheet.getCellPhoneticProperties': [[0, 1]],
  'Sheet.getCellPhoneticRuns': [[0, 1]],
  'Sheet.getCellShowPhonetic': [[0, 1]],
  'Sheet.getCellXfIndex': [[0, 1]],
  'Sheet.getComment': [[0, 1]],
  'Sheet.getCommentVisible': [[0, 1]],
  'Sheet.getDataTable': [[0, 1]],
  'Sheet.getDate': [[0, 1]],
  'Sheet.getDependents': [[0, 1]],
  'Sheet.getDisplayText': [[0, 1]],
  'Sheet.getEffectiveStyle': [[0, 1]],
  'Sheet.getFormula': [[0, 1]],
  'Sheet.getFormulaR1C1': [[0, 1]],
  'Sheet.getHyperlink': [[0, 1]],
  'Sheet.getLambdaText': [[0, 1]],
  'Sheet.getPrecedents': [[0, 1]],
  'Sheet.getSpillInfo': [[0, 1]],
  'Sheet.getTypedValue': [[0, 1]],
  'Sheet.getValue': [[0, 1]],
  'Sheet.removeComment': [[0, 1]],
  'Sheet.removeDataTable': [[0, 1]],
  'Sheet.removeHyperlink': [[0, 1]],
  'Sheet.setBlank': [[0, 1]],
  'Sheet.setBool': [[0, 1]],
  'Sheet.setCellPhonetic': [[0, 1]],
  'Sheet.setCellPhoneticProperties': [[0, 1]],
  'Sheet.setCellPhoneticRuns': [[0, 1]],
  'Sheet.setCellShowPhonetic': [[0, 1]],
  'Sheet.setCellXfIndex': [[0, 1]],
  'Sheet.setComment': [[0, 1]],
  'Sheet.setCommentVisible': [[0, 1]],
  'Sheet.setError': [[0, 1]],
  'Sheet.setFormula': [[0, 1]],
  'Sheet.setNumber': [[0, 1]],
  'Sheet.setText': [[0, 1]],
  'Sheet.setValue': [[0, 1]],
  'Sheet.setValues': [[0, 1]],
  'Sheet.validateValue': [[0, 1]],
});
const RECT_ARGS = Object.freeze({
  'Pivot.setAnchor': [0],
  'Sheet.addHyperlinkRange': [0],
  'Sheet.addMerge': [0],
  'Sheet.evaluateCfRange': [0],
  'Sheet.getCellRectPt': [0],
  'Sheet.getCellsInRange': [0],
  'Sheet.getMergesInRange': [0],
  'Sheet.getRangeStyleBases': [0],
  'Sheet.getRangeStyleBasesExcluding': [0],
  'Sheet.getTypedValues': [0],
  'Sheet.getValues': [0],
  'Sheet.iterRows': [0],
  'Sheet.listCells': [0],
  'Sheet.recalcViewport': [0],
  'Sheet.remapRangeStyle': [0],
  'Sheet.removeMerge': [0],
  'Sheet.setCellFormat': [0],
  'Sheet.setPrintArea': [0],
  'Sheet.setRangeXfIndex': [0],
  'Sheet.snapshotRangeStyle': [0],
});
const SPAN_ARGS = Object.freeze({
  'Sheet.clearColumnStyle': [[0, 'col', 'span']],
  'Sheet.clearColumnWidth': [[0, 'col', 'span']],
  'Sheet.clearRowHeight': [[0, 'row', 'span']],
  'Sheet.clearRowStyle': [[0, 'row', 'span']],
  'Sheet.deleteColumns': [[0, 'col', 'start_count']],
  'Sheet.deleteRows': [[0, 'row', 'start_count']],
  'Sheet.insertColumns': [[0, 'col', 'start_count']],
  'Sheet.insertRows': [[0, 'row', 'start_count']],
  'Sheet.setColumnCollapsed': [[0, 'col', 'span']],
  'Sheet.setColumnHidden': [[0, 'col', 'span']],
  'Sheet.setColumnOutline': [[0, 'col', 'span']],
  'Sheet.setColumnStyle': [[0, 'col', 'span']],
  'Sheet.setColumnWidth': [[0, 'col', 'span']],
  'Sheet.setRowAutoHeight': [[0, 'row', 'span']],
  'Sheet.setRowCollapsed': [[0, 'row', 'span']],
  'Sheet.setRowHeight': [[0, 'row', 'span']],
  'Sheet.setRowHidden': [[0, 'row', 'span']],
  'Sheet.setRowOutline': [[0, 'row', 'span']],
  'Sheet.setRowStyle': [[0, 'row', 'span']],
});
const REF_ARGS = Object.freeze({
  'Sheet.addSlicer': [0, 'table'],
  'Sheet.pivot': [0, 'pivot'],
  'Sheet.removeCfRule': [0, 'cfRule'],
  'Sheet.removePivot': [0, 'pivot'],
  'Sheet.removeProtectedRange': [0, 'protectedRange'],
  'Sheet.updateProtectedRange': [0, 'protectedRange'],
  'Workbook.removeTable': [0, 'table'],
  'Workbook.table': [0, 'table'],
});
const PATCH_KEYS = Object.freeze({
  'Chart.setLegend': [0, ['hiddenSeries', 'overlay', 'position']],
  'Pivot.setGrandTotals': [0, ['columns', 'rows']],
  'Pivot.setOptions': [
    0,
    [
      'altText',
      'altTextSummary',
      'autoFormatting',
      'customListSort',
      'enableDrill',
      'errorCaption',
      'fieldPrintTitles',
      'grandTotalCaption',
      'gridDropZones',
      'indent',
      'itemPrintTitles',
      'mergeItem',
      'missingCaption',
      'multipleFieldFilters',
      'pageOverThenDown',
      'pageWrap',
      'preserveFormatting',
      'printDrill',
      'showDataDropDown',
      'showDataTips',
      'showDrill',
      'showEmptyCol',
      'showEmptyRow',
      'showError',
      'showHeaders',
      'showMissing',
    ],
  ],
  'Pivot.setStyleInfo': [
    0,
    ['name', 'showColHeaders', 'showColStripes', 'showLastColumn', 'showRowHeaders', 'showRowStripes'],
  ],
  'Pivot.updateField': [
    1,
    [
      'axis',
      'cacheFieldIndex',
      'customName',
      'dateCalendar',
      'datePeriods',
      'defaultSubtotal',
      'fillDownLabels',
      'groupBaseIndex',
      'insertBlankRow',
      'numberFormat',
      'showAll',
      'sortByField',
      'sortOrder',
      'sourceName',
      'subtotalFns',
      'subtotalTop',
    ],
  ],
  'Pivot.updateItem': [2, ['cacheIndex', 'caption', 'name', 'showDetail', 'visible']],
  'PivotCache.setOptions': [0, ['missingItemsLimit', 'refreshOnLoad', 'saveData']],
  'Shape.update': [
    0,
    [
      'anchorKind',
      'editAs',
      'fillColor',
      'fillKind',
      'fromCol',
      'fromColOffsetEmu',
      'fromRow',
      'fromRowOffsetEmu',
      'heightEmu',
      'hidden',
      'kind',
      'lineColor',
      'lineKind',
      'lineWidthEmu',
      'name',
      'objectId',
      'presetGeometry',
      'rotation',
      'text',
      'textBox',
      'toCol',
      'toColOffsetEmu',
      'toRow',
      'toRowOffsetEmu',
      'widthEmu',
    ],
  ],
  'Sheet.setCellFormat': [1, ['alignment', 'border', 'fill', 'font', 'numberFormat', 'protection']],
  'Sheet.setFormatDefaults': [
    0,
    [
      'baseColWidth',
      'customHeight',
      'defaultColWidth',
      'defaultRowHeight',
      'effectiveDefaultColWidth',
      'effectiveDefaultRowHeight',
      'zeroHeight',
    ],
  ],
  'Sheet.setHeaderFooter': [
    0,
    [
      'alignWithMargins',
      'alignWithMarginsStated',
      'differentFirst',
      'differentFirstStated',
      'differentOddEven',
      'differentOddEvenStated',
      'evenFooter',
      'evenHeader',
      'firstFooter',
      'firstHeader',
      'oddFooter',
      'oddHeader',
      'scaleWithDoc',
      'scaleWithDocStated',
    ],
  ],
  'Sheet.setPageMargins': [
    0,
    [
      'bottom',
      'bottomStated',
      'footer',
      'footerStated',
      'header',
      'headerStated',
      'left',
      'leftStated',
      'right',
      'rightStated',
      'top',
      'topStated',
    ],
  ],
  'Sheet.setPageSetup': [
    0,
    [
      'fitToHeight',
      'fitToHeightStated',
      'fitToPage',
      'fitToPageStated',
      'fitToWidth',
      'fitToWidthStated',
      'orientation',
      'orientationStated',
      'paperSize',
      'paperSizeStated',
      'scale',
      'scaleStated',
    ],
  ],
  'Sheet.setPrintOptions': [
    0,
    [
      'gridLines',
      'gridLinesStated',
      'headings',
      'headingsStated',
      'horizontalCentered',
      'horizontalCenteredStated',
      'verticalCentered',
      'verticalCenteredStated',
    ],
  ],
  'Sheet.setView': [
    0,
    [
      'freezeCols',
      'freezeRows',
      'paneTopLeftCol',
      'paneTopLeftRow',
      'rightToLeft',
      'showFormulas',
      'showGridLines',
      'showOutlineSymbols',
      'showRowColHeaders',
      'showZeros',
      'splitX',
      'splitY',
      'tabSelected',
      'topLeftCol',
      'topLeftRow',
      'viewMode',
      'visibility',
      'zoomScale',
    ],
  ],
  'Sheet.updateProtectedRange': [
    1,
    ['algorithmName', 'hashValue', 'name', 'password', 'ranges', 'saltValue', 'securityDescriptor', 'spinCount'],
  ],
  'Sheet.updateSparklineGroup': [
    1,
    [
      'colorAxis',
      'colorFirst',
      'colorHigh',
      'colorLast',
      'colorLow',
      'colorMarkers',
      'colorNegative',
      'colorSeries',
      'dateAxis',
      'dateAxisRef',
      'displayHidden',
      'displayXAxis',
      'emptyCellsAs',
      'first',
      'high',
      'last',
      'lineWeight',
      'low',
      'manualMax',
      'manualMin',
      'markers',
      'maxAxisType',
      'minAxisType',
      'negative',
      'sparklines',
      'type',
    ],
  ],
  'Sheet.updateThreadedComment': [
    1,
    ['col', 'created', 'done', 'id', 'mentions', 'parentId', 'personId', 'row', 'text'],
  ],
  'Slicer.update': [
    0,
    [
      'anchorKind',
      'cacheName',
      'caption',
      'columnCount',
      'columnName',
      'crossFilter',
      'fromCol',
      'fromColOffsetEmu',
      'fromRow',
      'fromRowOffsetEmu',
      'name',
      'objectId',
      'rowHeightEmu',
      'sortOrder',
      'style',
      'tableName',
      'toCol',
      'toColOffsetEmu',
      'toRow',
      'toRowOffsetEmu',
    ],
  ],
  'Table.update': [
    0,
    [
      'columns',
      'displayName',
      'headerRow',
      'name',
      'range',
      'showColumnStripes',
      'showFirstColumn',
      'showLastColumn',
      'showRowStripes',
      'styleName',
      'totalsRow',
    ],
  ],
  'Workbook.setIterative': [0, ['enabled', 'maxChange', 'maxIterations']],
  'Workbook.setTableStyleDefaults': [0, ['pivotStyle', 'tableStyle']],
});
const PATCH_NESTED_KEYS = Object.freeze({
  'Sheet.setCellFormat': {
    alignment: [
      'horizontal',
      'indent',
      'justifyLastLine',
      'readingOrder',
      'relativeIndent',
      'shrinkToFit',
      'textRotation',
      'vertical',
      'wrapText',
    ],
    border: ['bottom', 'diagonal', 'diagonalDown', 'diagonalUp', 'left', 'right', 'top'],
    'border.bottom': ['color', 'style'],
    'border.diagonal': ['color', 'style'],
    'border.left': ['color', 'style'],
    'border.right': ['color', 'style'],
    'border.top': ['color', 'style'],
    fill: ['bg', 'fg', 'gradient', 'pattern'],
    font: [
      'bold',
      'charset',
      'color',
      'condense',
      'extend',
      'family',
      'italic',
      'name',
      'outline',
      'scheme',
      'shadow',
      'size',
      'strike',
      'underline',
      'vertAlign',
    ],
    protection: ['hidden', 'locked'],
  },
});
// </facade-tables>
// <record-range-fields> generated by tools/codegen/gen_facade_tables.py; do not edit
const RECORD_RANGE_FIELDS = Object.freeze({
  AutoFilterInput: { range: 'CellRect', sort: 'AutoFilterSortStateInput' },
  AutoFilterSortCondition: { ref: 'CellRect' },
  AutoFilterSortConditionInput: { ref: 'CellRect' },
  AutoFilterSortState: { ref: 'CellRect', conditions: 'AutoFilterSortCondition[]' },
  AutoFilterSortStateInput: { ref: 'CellRect', conditions: 'AutoFilterSortConditionInput[]' },
  ConditionalFormatInput: { ranges: 'CellRect[]' },
  DataTableInput: { rect: 'CellRect' },
  DataValidationInput: { ranges: 'CellRect[]' },
  PivotInput: { anchor: 'CellAddress' },
  ReferenceRangeInput: { rect: 'CellRect' },
  SheetSelectionInput: { active: 'CellAddress', ranges: 'CellRect[]' },
  TableInput: { range: 'SheetRect' },
  TablePatch: { range: 'SheetRect' },
});
// </record-range-fields>
// <method-signatures> generated by tools/codegen/gen_facade_tables.py; do not edit
const METHOD_ARITY = Object.freeze({
  'Chart.addSeries': 2,
  'Chart.getData': 0,
  'Chart.getPreview': 1,
  'Chart.getProperties': 0,
  'Chart.removeSeries': 2,
  'Chart.rename': 1,
  'Chart.setAnchor': 1,
  'Chart.setAxis': 2,
  'Chart.setDisplay': 1,
  'Chart.setFilter': 2,
  'Chart.setLegend': 1,
  'Chart.setPlot': 2,
  'Chart.setPoint': 4,
  'Chart.setSeries': 3,
  'Chart.setSource': 1,
  'Chart.setTitle': 1,
  'Chart.setType': 3,
  'Chart.setZOrder': 1,
  'Chart.snapshot': 0,
  'Drawing.getDrawingProperties': 0,
  'Picture.getProperties': 0,
  'Picture.rename': 1,
  'Picture.setAnchor': 1,
  'Picture.setZOrder': 1,
  'Picture.snapshot': 0,
  'Pivot.addDataField': 1,
  'Pivot.addField': 1,
  'Pivot.addFieldItem': 3,
  'Pivot.addFieldSharedItem': 3,
  'Pivot.addFieldSubtotalFn': 2,
  'Pivot.addFilter': 1,
  'Pivot.clearDataFields': 0,
  'Pivot.clearFieldDateGroup': 1,
  'Pivot.clearFieldItems': 1,
  'Pivot.clearFieldSubtotalFns': 1,
  'Pivot.clearFields': 0,
  'Pivot.clearFilters': 0,
  'Pivot.getCache': 0,
  'Pivot.getColumnFieldOrder': 0,
  'Pivot.getDataFieldDefaultName': 3,
  'Pivot.getFieldItem': 2,
  'Pivot.getGrandTotals': 0,
  'Pivot.getLayout': 0,
  'Pivot.getOptions': 0,
  'Pivot.getProperties': 0,
  'Pivot.getReportLayout': 0,
  'Pivot.getRowFieldOrder': 0,
  'Pivot.getStyleInfo': 0,
  'Pivot.isProjectionPending': 0,
  'Pivot.listDataFields': 0,
  'Pivot.listFieldItems': 1,
  'Pivot.listFields': 0,
  'Pivot.listFilters': 0,
  'Pivot.project': 0,
  'Pivot.removeFilter': 1,
  'Pivot.rename': 1,
  'Pivot.setAnchor': 1,
  'Pivot.setColumnFieldOrder': 1,
  'Pivot.setDataField': 2,
  'Pivot.setFieldAxis': 2,
  'Pivot.setFieldDateGroup': 2,
  'Pivot.setFieldShowDetail': 2,
  'Pivot.setFieldSort': 3,
  'Pivot.setGrandTotals': 1,
  'Pivot.setOptions': 1,
  'Pivot.setReportLayout': 1,
  'Pivot.setRowFieldOrder': 1,
  'Pivot.setStyleInfo': 1,
  'Pivot.showDetails': 2,
  'Pivot.updateField': 2,
  'Pivot.updateItem': 3,
  'PivotCache.addCalculatedField': 2,
  'PivotCache.addCalculatedItem': 3,
  'PivotCache.addField': 1,
  'PivotCache.addFieldSharedItem': 3,
  'PivotCache.addRecord': 0,
  'PivotCache.clearFieldSharedItems': 1,
  'PivotCache.clearFields': 0,
  'PivotCache.clearRecords': 0,
  'PivotCache.getOptions': 0,
  'PivotCache.getRecordCount': 0,
  'PivotCache.getWorksheetSource': 0,
  'PivotCache.listCalculatedItems': 0,
  'PivotCache.listFieldSharedItems': 1,
  'PivotCache.listFields': 0,
  'PivotCache.moveCalculatedItem': 2,
  'PivotCache.removeCalculatedField': 1,
  'PivotCache.removeCalculatedItem': 1,
  'PivotCache.setCalculatedFieldFormula': 2,
  'PivotCache.setCalculatedItemFormula': 2,
  'PivotCache.setOptions': 1,
  'PivotCache.setRecordValue': 4,
  'PivotCache.setWorksheetSource': 1,
  'Shape.getProperties': 0,
  'Shape.rename': 1,
  'Shape.setZOrder': 1,
  'Shape.snapshot': 0,
  'Shape.update': 1,
  'Sheet.addCfRule': 1,
  'Sheet.addChart': 1,
  'Sheet.addColumnBreak': 2,
  'Sheet.addDataTable': 1,
  'Sheet.addHyperlink': 3,
  'Sheet.addHyperlinkRange': 2,
  'Sheet.addMerge': 1,
  'Sheet.addPicture': 2,
  'Sheet.addPivot': 1,
  'Sheet.addProtectedRange': 1,
  'Sheet.addRowBreak': 2,
  'Sheet.addShape': 1,
  'Sheet.addSlicer': 3,
  'Sheet.addSparklineGroup': 1,
  'Sheet.addThreadedComment': 1,
  'Sheet.addValidation': 1,
  'Sheet.applyAutoFilter': 0,
  'Sheet.chart': 1,
  'Sheet.clearAutoFilter': 0,
  'Sheet.clearBreaks': 0,
  'Sheet.clearCellXfIndex': 2,
  'Sheet.clearCfRules': 0,
  'Sheet.clearColumnStyle': 2,
  'Sheet.clearColumnWidth': 2,
  'Sheet.clearHyperlinks': 0,
  'Sheet.clearMerges': 0,
  'Sheet.clearRowHeight': 2,
  'Sheet.clearRowStyle': 2,
  'Sheet.clearValidations': 0,
  'Sheet.convertColumnCharsToPt': 2,
  'Sheet.convertColumnPtToChars': 2,
  'Sheet.deleteColumns': 2,
  'Sheet.deleteRows': 2,
  'Sheet.drawing': 1,
  'Sheet.evaluateAutoFilter': 0,
  'Sheet.evaluateCfFormula': 5,
  'Sheet.evaluateCfRange': 2,
  'Sheet.evaluateFormula': 3,
  'Sheet.evaluateFormulaArray': 3,
  'Sheet.evaluateTypedFormula': 3,
  'Sheet.evaluateTypedFormulaArray': 3,
  'Sheet.getArrayFormula': 2,
  'Sheet.getAutoFilter': 0,
  'Sheet.getAutoFilterXml': 0,
  'Sheet.getCellFormat': 2,
  'Sheet.getCellPhonetic': 2,
  'Sheet.getCellPhoneticProperties': 2,
  'Sheet.getCellPhoneticRuns': 2,
  'Sheet.getCellRectPt': 2,
  'Sheet.getCellShowPhonetic': 2,
  'Sheet.getCellXfIndex': 2,
  'Sheet.getCellsInRange': 3,
  'Sheet.getColumnStyle': 1,
  'Sheet.getColumnWidthPt': 2,
  'Sheet.getComment': 2,
  'Sheet.getCommentVisible': 2,
  'Sheet.getDataTable': 2,
  'Sheet.getDate': 2,
  'Sheet.getDependents': 3,
  'Sheet.getDisplayText': 2,
  'Sheet.getEffectiveStyle': 2,
  'Sheet.getFormatDefaults': 0,
  'Sheet.getFormula': 2,
  'Sheet.getFormulaR1C1': 2,
  'Sheet.getHeaderFooter': 0,
  'Sheet.getHeaderFooterXml': 0,
  'Sheet.getHyperlink': 2,
  'Sheet.getInvalidCells': 2,
  'Sheet.getLambdaText': 2,
  'Sheet.getMergesInRange': 1,
  'Sheet.getPageMargins': 0,
  'Sheet.getPageMarginsXml': 0,
  'Sheet.getPageSetup': 0,
  'Sheet.getPageSetupXml': 0,
  'Sheet.getPagination': 0,
  'Sheet.getPrecedents': 3,
  'Sheet.getPrintArea': 0,
  'Sheet.getPrintOptions': 0,
  'Sheet.getPrintOptionsXml': 0,
  'Sheet.getPrintTitles': 0,
  'Sheet.getRangeStyleBases': 1,
  'Sheet.getRangeStyleBasesExcluding': 3,
  'Sheet.getRowHeightPt': 1,
  'Sheet.getRowStyle': 1,
  'Sheet.getSelection': 0,
  'Sheet.getSheetPrXml': 0,
  'Sheet.getSheetProtection': 0,
  'Sheet.getSpillInfo': 2,
  'Sheet.getTabColor': 0,
  'Sheet.getThreadedComment': 1,
  'Sheet.getTypedValue': 2,
  'Sheet.getTypedValues': 1,
  'Sheet.getUsedRange': 0,
  'Sheet.getValue': 2,
  'Sheet.getValues': 1,
  'Sheet.getView': 0,
  'Sheet.getWidthModel': 1,
  'Sheet.insertColumns': 2,
  'Sheet.insertRows': 2,
  'Sheet.iterRows': 1,
  'Sheet.listCells': 1,
  'Sheet.listCfRules': 0,
  'Sheet.listCharts': 0,
  'Sheet.listColumnBreaks': 0,
  'Sheet.listColumnOverrides': 0,
  'Sheet.listComments': 0,
  'Sheet.listDrawings': 0,
  'Sheet.listHyperlinks': 0,
  'Sheet.listMerges': 0,
  'Sheet.listPictures': 0,
  'Sheet.listPivots': 0,
  'Sheet.listProtectedRanges': 0,
  'Sheet.listRowBreaks': 0,
  'Sheet.listRowOverrides': 0,
  'Sheet.listShapes': 0,
  'Sheet.listSlicers': 0,
  'Sheet.listSparklineGroups': 0,
  'Sheet.listThreadedComments': 0,
  'Sheet.listValidations': 0,
  'Sheet.move': 1,
  'Sheet.picture': 1,
  'Sheet.pivot': 1,
  'Sheet.recalcViewport': 1,
  'Sheet.remapRangeStyle': 3,
  'Sheet.removeAutoFilter': 0,
  'Sheet.removeCfRule': 1,
  'Sheet.removeColumnBreak': 1,
  'Sheet.removeComment': 2,
  'Sheet.removeDataTable': 2,
  'Sheet.removeDrawing': 1,
  'Sheet.removeHyperlink': 2,
  'Sheet.removeMerge': 1,
  'Sheet.removePivot': 1,
  'Sheet.removeProtectedRange': 1,
  'Sheet.removeRowBreak': 1,
  'Sheet.removeSparklineGroup': 1,
  'Sheet.removeThreadedComment': 1,
  'Sheet.removeValidation': 1,
  'Sheet.rename': 1,
  'Sheet.restoreDrawing': 2,
  'Sheet.restoreRangeStyle': 1,
  'Sheet.setArrayFormula': 1,
  'Sheet.setAutoFilter': 1,
  'Sheet.setAutoFilterXml': 1,
  'Sheet.setBlank': 2,
  'Sheet.setBool': 3,
  'Sheet.setCellFormat': 2,
  'Sheet.setCellPhonetic': 3,
  'Sheet.setCellPhoneticProperties': 3,
  'Sheet.setCellPhoneticRuns': 3,
  'Sheet.setCellShowPhonetic': 3,
  'Sheet.setCellXfIndex': 3,
  'Sheet.setColumnCollapsed': 3,
  'Sheet.setColumnHidden': 3,
  'Sheet.setColumnOutline': 3,
  'Sheet.setColumnStyle': 3,
  'Sheet.setColumnWidth': 3,
  'Sheet.setComment': 3,
  'Sheet.setCommentVisible': 3,
  'Sheet.setError': 3,
  'Sheet.setFormatDefaults': 1,
  'Sheet.setFormula': 3,
  'Sheet.setHeaderFooter': 1,
  'Sheet.setHeaderFooterXml': 1,
  'Sheet.setNumber': 3,
  'Sheet.setPageMargins': 1,
  'Sheet.setPageMarginsXml': 1,
  'Sheet.setPageSetup': 1,
  'Sheet.setPageSetupXml': 1,
  'Sheet.setPrintArea': 1,
  'Sheet.setPrintOptions': 1,
  'Sheet.setPrintOptionsXml': 1,
  'Sheet.setPrintTitles': 1,
  'Sheet.setRangeXfIndex': 2,
  'Sheet.setRowAutoHeight': 3,
  'Sheet.setRowCollapsed': 3,
  'Sheet.setRowHeight': 3,
  'Sheet.setRowHidden': 3,
  'Sheet.setRowOutline': 3,
  'Sheet.setRowStyle': 3,
  'Sheet.setSelection': 1,
  'Sheet.setSheetPrXml': 1,
  'Sheet.setSheetProtection': 1,
  'Sheet.setTabColor': 1,
  'Sheet.setText': 3,
  'Sheet.setValue': 3,
  'Sheet.setValues': 3,
  'Sheet.setView': 1,
  'Sheet.shape': 1,
  'Sheet.slicer': 1,
  'Sheet.snapshotRangeStyle': 1,
  'Sheet.updateProtectedRange': 2,
  'Sheet.updateSparklineGroup': 2,
  'Sheet.updateThreadedComment': 2,
  'Sheet.validateValue': 3,
  'Slicer.getProperties': 0,
  'Slicer.setZOrder': 1,
  'Slicer.update': 1,
  'Table.applyAutoFilter': 0,
  'Table.clearAutoFilter': 0,
  'Table.evaluateAutoFilter': 0,
  'Table.getAutoFilter': 0,
  'Table.getProperties': 0,
  'Table.removeAutoFilter': 0,
  'Table.renameColumn': 2,
  'Table.setAutoFilter': 1,
  'Table.setColumnTotals': 2,
  'Table.update': 1,
  'Workbook.[Symbol.dispose]': 0,
  'Workbook.[Symbol.iterator]': 0,
  'Workbook.addBorder': 1,
  'Workbook.addCellStyleXf': 1,
  'Workbook.addCellXf': 1,
  'Workbook.addDxf': 1,
  'Workbook.addFill': 1,
  'Workbook.addFont': 1,
  'Workbook.addNumFmt': 1,
  'Workbook.addPerson': 1,
  'Workbook.addPivotCache': 1,
  'Workbook.addSheet': 1,
  'Workbook.addTable': 1,
  'Workbook.clearTheme': 0,
  'Workbook.copySheet': 3,
  'Workbook.dispose': 0,
  'Workbook.formatReference': 2,
  'Workbook.formatValue': 2,
  'Workbook.getActiveSheet': 0,
  'Workbook.getBorder': 1,
  'Workbook.getCalcMode': 0,
  'Workbook.getCalcOnSave': 0,
  'Workbook.getCellStyle': 1,
  'Workbook.getCellStyleXf': 1,
  'Workbook.getCellXf': 1,
  'Workbook.getDate1904': 0,
  'Workbook.getDefinedName': 2,
  'Workbook.getDiagnostics': 0,
  'Workbook.getDxf': 1,
  'Workbook.getFileName': 0,
  'Workbook.getFill': 1,
  'Workbook.getFont': 1,
  'Workbook.getFullPrecision': 0,
  'Workbook.getIterative': 0,
  'Workbook.getMemoryUsage': 0,
  'Workbook.getNumFmtString': 1,
  'Workbook.getPassthrough': 1,
  'Workbook.getPerson': 1,
  'Workbook.getPinnedNow': 0,
  'Workbook.getProfileId': 0,
  'Workbook.getRefMode': 0,
  'Workbook.getTableStyle': 1,
  'Workbook.getTableStyleDefaults': 0,
  'Workbook.getThemeColors': 0,
  'Workbook.getThemeFonts': 0,
  'Workbook.getThemeScriptFont': 1,
  'Workbook.getWorkbookProtection': 0,
  'Workbook.hasDirtyCells': 0,
  'Workbook.listBorders': 0,
  'Workbook.listCellStyleXfs': 0,
  'Workbook.listCellStyles': 0,
  'Workbook.listCellXfs': 0,
  'Workbook.listDefinedNames': 0,
  'Workbook.listDxfs': 0,
  'Workbook.listExternalLinks': 0,
  'Workbook.listFills': 0,
  'Workbook.listFonts': 0,
  'Workbook.listPassthroughs': 0,
  'Workbook.listPeople': 0,
  'Workbook.listPivotCaches': 0,
  'Workbook.listSheets': 0,
  'Workbook.listTableStyles': 0,
  'Workbook.listTables': 0,
  'Workbook.parseReference': 2,
  'Workbook.pivotCache': 1,
  'Workbook.recalc': 0,
  'Workbook.recalcDataTables': 0,
  'Workbook.recalcParallel': 1,
  'Workbook.removeCellStyle': 1,
  'Workbook.removeDefinedName': 2,
  'Workbook.removePerson': 1,
  'Workbook.removePivotCache': 1,
  'Workbook.removeSheet': 1,
  'Workbook.removeTable': 1,
  'Workbook.removeTableStyle': 1,
  'Workbook.resolveColor': 2,
  'Workbook.save': 1,
  'Workbook.saveWithDiagnostics': 1,
  'Workbook.setActiveSheet': 1,
  'Workbook.setCalcMode': 1,
  'Workbook.setCalcOnSave': 1,
  'Workbook.setCellStyle': 1,
  'Workbook.setDate1904': 1,
  'Workbook.setDefaultFont': 1,
  'Workbook.setDefinedName': 1,
  'Workbook.setFileName': 1,
  'Workbook.setFont': 2,
  'Workbook.setFullPrecision': 1,
  'Workbook.setIterative': 1,
  'Workbook.setIterativeProgress': 1,
  'Workbook.setPinnedNow': 1,
  'Workbook.setProfileId': 1,
  'Workbook.setRefMode': 1,
  'Workbook.setTableStyle': 1,
  'Workbook.setTableStyleDefaults': 1,
  'Workbook.setThemeColors': 1,
  'Workbook.setThemeFonts': 1,
  'Workbook.setThemeScriptFont': 1,
  'Workbook.setWorkbookProtection': 1,
  'Workbook.sheet': 1,
  'Workbook.table': 1,
});
const TYPED_ARGS = Object.freeze({
  'Sheet.addCfRule': [[0, 'ConditionalFormatInput']],
  'Sheet.addDataTable': [[0, 'DataTableInput']],
  'Sheet.addPivot': [[0, 'PivotInput']],
  'Sheet.addValidation': [[0, 'DataValidationInput']],
  'Sheet.setAutoFilter': [[0, 'AutoFilterInput']],
  'Sheet.setPrintArea': [[0, 'CellRect[]']],
  'Sheet.setSelection': [[0, 'SheetSelectionInput']],
  'Table.setAutoFilter': [[0, 'AutoFilterInput']],
  'Table.update': [[0, 'TablePatch']],
  'Workbook.addTable': [[0, 'TableInput']],
  'Workbook.formatReference': [[0, 'ReferenceRangeInput']],
});
const DEFAULT_ARGS = Object.freeze({
  'Sheet.evaluateCfRange': [[1, 'today']],
  'Sheet.getDependents': [[2, 1]],
  'Sheet.getPrecedents': [[2, 1]],
  'Workbook.copySheet': [[1, null]],
  'Workbook.recalcParallel': [[0, 0]],
});
// </method-signatures>
// Arguments the declarations type `T | null` (or optional) where the raw method
// takes the C sentinel: `Class.method` -> [position in the public call,
// sentinel, whether an omitted argument also means absent].
const ABSENT_ARGS = Object.freeze({
  'Workbook.setFileName': [0, ''],
  'Workbook.setWorkbookProtection': [0, {}],
  'Sheet.setAutoFilterXml': [0, ''],
  'Sheet.setCellPhonetic': [2, ''],
  'Sheet.setCellPhoneticRuns': [2, []],
  'Sheet.setHeaderFooterXml': [0, ''],
  'Sheet.setPageMarginsXml': [0, ''],
  'Sheet.setPageSetupXml': [0, ''],
  'Sheet.setPrintArea': [0, []],
  'Sheet.setPrintOptionsXml': [0, ''],
  'Sheet.setSheetPrXml': [0, ''],
  'Sheet.setSheetProtection': [0, {}],
  'Pivot.setFieldAxis': [1, -1],
});

// Patch setters and the non-nullable fields of their patch: `Class.method` ->
// { field: expected type }. `null` on one of these fields is a TypeError; an
// omitted or `undefined` key keeps the value. The patch's position and its keys
// are `PATCH_KEYS`.
const bools = (...fields) => Object.fromEntries(fields.map((f) => [f, 'boolean']));
const numbers = (...fields) => Object.fromEntries(fields.map((f) => [f, 'number']));
const strings = (...fields) => Object.fromEntries(fields.map((f) => [f, 'string']));
const objects = (...fields) => Object.fromEntries(fields.map((f) => [f, 'object']));
const PATCH_FIELDS = Object.freeze({
  'Workbook.setIterative': { ...bools('enabled'), ...numbers('maxIterations', 'maxChange') },
  'PivotCache.setOptions': bools('saveData', 'refreshOnLoad'),
  'Sheet.setView': {
    ...numbers(
      'zoomScale',
      'freezeRows',
      'freezeCols',
      'splitX',
      'splitY',
      'topLeftRow',
      'topLeftCol',
      'paneTopLeftRow',
      'paneTopLeftCol',
      'visibility',
      'viewMode',
    ),
    ...bools(
      'showGridLines',
      'showRowColHeaders',
      'showZeros',
      'showFormulas',
      'showOutlineSymbols',
      'rightToLeft',
      'tabSelected',
    ),
  },
  'Sheet.setPrintOptions': bools('gridLines', 'headings', 'horizontalCentered', 'verticalCentered'),
  'Sheet.setHeaderFooter': bools('differentOddEven', 'differentFirst', 'scaleWithDoc', 'alignWithMargins'),
  'Sheet.setPageSetup': { ...numbers('paperSize', 'scale', 'fitToWidth', 'fitToHeight'), ...bools('fitToPage') },
  'Sheet.setPageMargins': numbers('left', 'right', 'top', 'bottom', 'header', 'footer'),
  'Sheet.setFormatDefaults': { ...numbers('baseColWidth'), ...bools('customHeight', 'zeroHeight') },
  'Pivot.updateField': bools('defaultSubtotal', 'insertBlankRow', 'fillDownLabels', 'showAll', 'subtotalTop'),
  'Pivot.updateItem': bools('visible', 'showDetail'),
  'Pivot.setStyleInfo': bools('showRowHeaders', 'showColHeaders', 'showRowStripes', 'showColStripes', 'showLastColumn'),
  'Pivot.setOptions': {
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
  'Chart.setLegend': { position: 'number', overlay: 'boolean', hiddenSeries: 'array' },
  // `range` takes a `SheetRect` or an A1 string, checked by `tableRange`.
  'Table.update': {
    ...strings('name', 'displayName'),
    ...bools('showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes', 'headerRow', 'totalsRow'),
  },
  'Shape.update': {
    ...strings('name', 'presetGeometry', 'text'),
    ...bools('textBox', 'hidden'),
    ...numbers(
      'anchorKind',
      'editAs',
      'fromRow',
      'fromCol',
      'fromRowOffsetEmu',
      'fromColOffsetEmu',
      'toRow',
      'toCol',
      'toRowOffsetEmu',
      'toColOffsetEmu',
      'widthEmu',
      'heightEmu',
      'lineWidthEmu',
      'rotation',
    ),
    ...objects('fillColor', 'lineColor'),
  },
  'Slicer.update': {
    ...strings('caption'),
    ...numbers(
      'columnCount',
      'rowHeightEmu',
      'fromRow',
      'fromCol',
      'fromRowOffsetEmu',
      'fromColOffsetEmu',
      'toRow',
      'toCol',
      'toRowOffsetEmu',
      'toColOffsetEmu',
      'sortOrder',
      'crossFilter',
    ),
  },
  'Sheet.updateSparklineGroup': {
    ...numbers('type', 'minAxisType', 'maxAxisType', 'lineWeight', 'emptyCellsAs'),
    ...objects(
      'colorSeries',
      'colorNegative',
      'colorAxis',
      'colorMarkers',
      'colorFirst',
      'colorLast',
      'colorHigh',
      'colorLow',
    ),
    ...bools('markers', 'high', 'low', 'first', 'last', 'negative', 'displayXAxis', 'displayHidden', 'dateAxis'),
    sparklines: 'array',
  },
  'Sheet.updateProtectedRange': {
    ...strings('name', 'password', 'algorithmName', 'hashValue', 'saltValue', 'securityDescriptor'),
    spinCount: 'number',
    ranges: 'array',
  },
});

// The nullable keys of the hand-written patches the read-modify-write and drawing setters take. A `ShapePatch`
// `fillKind` / `lineKind` of `null` is the value the getter returns for a paint the engine cannot write, and
// leaves the paint unchanged.
const PATCH_NULLABLE = Object.freeze({
  'Shape.update': { ...numbers('fillKind', 'lineKind'), ...objects('fillColor', 'lineColor') },
  'Slicer.update': strings('style'),
  'Sheet.updateSparklineGroup': { ...numbers('manualMin', 'manualMax'), ...strings('dateAxisRef') },
  'Sheet.updateProtectedRange': {},
});

// The anchor keys of a drawing patch, which apply only together with the trigger key that carries the anchor.
const ANCHOR_PATCH_KEYS = Object.freeze({
  'Shape.update': [
    'anchorKind',
    [
      'editAs',
      'fromRow',
      'fromCol',
      'fromRowOffsetEmu',
      'fromColOffsetEmu',
      'toRow',
      'toCol',
      'toRowOffsetEmu',
      'toColOffsetEmu',
      'widthEmu',
      'heightEmu',
    ],
  ],
  'Slicer.update': [
    'fromRow',
    ['fromCol', 'fromRowOffsetEmu', 'fromColOffsetEmu', 'toRow', 'toCol', 'toRowOffsetEmu', 'toColOffsetEmu'],
  ],
});

/** Checks a drawing patch of `where`: its types and nulls, and no anchor key without the trigger key. */
function checkDrawingPatch(where, patch) {
  checkPatch(where, patch, PATCH_FIELDS[where], PATCH_NULLABLE[where]);
  const [trigger, anchorKeys] = ANCHOR_PATCH_KEYS[where];
  if (patch[trigger] !== undefined) return;
  const stray = anchorKeys.find((key) => patch[key] !== undefined);
  if (stray !== undefined) throw new TypeError(`${where} ${stray} needs ${trigger} in the same patch`);
}

const COLOR_KEYS = Object.freeze(['indexed', 'kind', 'rgb', 'theme', 'tint']);
const SPARKLINE_KEYS = Object.freeze(['formula', 'location']);

/**
 * Refuses an unknown key inside the records a `SparklineGroupPatch` nests: each `Color` and each element of
 * `sparklines`. A patch that is not an object is left to `checkPatch`.
 */
function checkSparklinePatch(where, patch) {
  if (!isObject(patch)) return;
  const refuse = (value, accepted, path) => {
    if (!isObject(value)) return;
    for (const [key, item] of Object.entries(value)) {
      if (item !== undefined && !accepted.includes(key)) {
        throw new TypeError(`${where} argument 2 has no field ${path}.${key}`);
      }
    }
  };
  for (const [key, value] of Object.entries(patch)) {
    if (PATCH_FIELDS[where][key] === 'object') refuse(value, COLOR_KEYS, key);
  }
  if (Array.isArray(patch.sparklines)) {
    for (const [i, item] of patch.sparklines.entries()) refuse(item, SPARKLINE_KEYS, `sparklines[${i}]`);
  }
}

/**
 * A `TablePatch.range` as the table's new `CellRect`: an A1 string, unqualified or naming the table's sheet, or
 * a `SheetRect` on that sheet. Another sheet is `InvalidArgument`.
 */
function tableRange(raw, sheet, range, where) {
  let onSheet;
  let rect;
  if (typeof range === 'string') {
    const ref = raw.parseReference(range);
    if (ref === null) {
      throw new FormulonError(
        StatusCode.InvalidArgument,
        `${JSON.stringify(range)} is not an A1 range`,
        `address=${range}`,
      );
    }
    onSheet = ref.sheetName === null ? sheet : raw.getSheetIndex(ref.sheetName);
    rect = ref.rect;
  } else if (isObject(range)) {
    onSheet = requireIndex(range.sheet, `${where} range.sheet`);
    rect = requireObject(range.rect, `${where} range.rect`);
  } else {
    throw new TypeError(`${where} range must be a SheetRect or an A1 string`);
  }
  if (onSheet !== sheet) {
    throw new FormulonError(
      StatusCode.InvalidArgument,
      `${where} range is not on the table's sheet`,
      `sheet_index=${sheet}`,
    );
  }
  return rect;
}

// How a string key of each `*Ref` kind is matched: the element count, the key
// of element `i`, the wording of the errors, and the handle class that also
// names an element. Tables are workbook-wide.
const REF_KINDS = Object.freeze({
  table: {
    noun: 'table displayName',
    handle: 'Table',
    context: 'table_name',
    count: (raw) => raw.getTableCount(),
    key: (raw, _sheet, i) => raw.getTable(i).displayName,
  },
  pivot: {
    noun: 'pivot name',
    handle: 'Pivot',
    context: 'pivot_name',
    count: (raw, sheet) => raw.getPivotCount(sheet),
    key: (raw, sheet, i) => raw.getPivotName(sheet, i),
  },
  cfRule: {
    noun: 'conditional-format rule id',
    context: 'cf_rule_id',
    count: (raw, sheet) => raw.getCfRuleCount(sheet),
    key: (raw, sheet, i) => raw.getCfRule(sheet, i).id,
  },
  protectedRange: {
    noun: 'protected range name',
    context: 'protected_range_name',
    count: (raw, sheet) => raw.getProtectedRangeCount(sheet),
    key: (raw, sheet, i) => raw.getProtectedRange(sheet, i).name,
  },
});

const GRID_ROWS = 1048576;
const GRID_COLS = 16384;
// 2^22 cells: the largest range `getValues` reads.
const MAX_VALUES_CELLS = 4194304;
const BLANK = Object.freeze({ kind: ValueKind.Blank });
const DEFAULT_LEGEND = Object.freeze({ position: ChartLegendPosition.Right, overlay: false, hiddenSeries: [] });
const GEOMETRY_MODES = new Set(Object.values(GeometryMode));

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

// An index or id argument: a number in uint32 range.
function requireIndex(value, where) {
  requireNumber(value, where);
  if (!Number.isInteger(value) || value < 0 || value > 0xffffffff) {
    throw new RangeError(`${where} is outside uint32 range`);
  }
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
function cellAddress(raw, a1, where) {
  const ref = raw.parseReference(a1);
  if (ref !== null && ref.sheetName !== null && where?.startsWith('Sheet.')) {
    const local = a1.slice(a1.lastIndexOf('!') + 1);
    const method = where.slice('Sheet.'.length);
    throw new FormulonError(
      StatusCode.InvalidArgument,
      `${JSON.stringify(a1)} names a sheet; use wb.sheet(${JSON.stringify(ref.sheetName)}).${method}(${JSON.stringify(local)})`,
      `address=${a1}`,
    );
  }
  if (ref === null || ref.sheetName !== null) throw invalidAddress(a1);
  const { rect } = ref;
  if (rect.firstRow !== rect.lastRow || rect.firstCol !== rect.lastCol) throw invalidAddress(a1);
  return rect;
}

// A range argument: any unqualified A1 reference, a single cell being 1x1.
function rangeAddress(raw, a1) {
  const ref = raw.parseReference(a1);
  if (ref === null || ref.sheetName !== null) {
    throw new FormulonError(
      StatusCode.InvalidArgument,
      `${JSON.stringify(a1)} is not an A1 range without a sheet name`,
      `address=${a1}`,
    );
  }
  return ref.rect;
}

// The axis a span argument (`SPAN_ARGS`) runs along: the rectangle it must
// cover whole, and the coordinates that bound it.
const SPAN_AXES = Object.freeze({
  col: {
    noun: 'whole-column',
    whole: (rect) => rect.firstRow === 0 && rect.lastRow === GRID_ROWS - 1,
    first: 'firstCol',
    last: 'lastCol',
  },
  row: {
    noun: 'whole-row',
    whole: (rect) => rect.firstCol === 0 && rect.lastCol === GRID_COLS - 1,
    first: 'firstRow',
    last: 'lastRow',
  },
});

// A span argument: an unqualified A1 reference over whole columns (`B`, `B:D`)
// or whole rows (`3`, `3:5`); the whole sheet is both.
function spanAddress(raw, text, axis) {
  const spec = SPAN_AXES[axis];
  const ref = raw.parseReference(text.includes(':') ? text : `${text}:${text}`);
  if (ref === null || ref.sheetName !== null || !spec.whole(ref.rect)) {
    throw new FormulonError(
      StatusCode.InvalidArgument,
      `${JSON.stringify(text)} is not a ${spec.noun} A1 span without a sheet name`,
      `address=${text}`,
    );
  }
  return [ref.rect[spec.first], ref.rect[spec.last]];
}

/**
 * The numeric form of a call whose span arguments (`SPAN_ARGS`) may be one A1
 * span string: `(first, last)` for a `span` form, `(first, count)` for a
 * `start_count` one.
 */
function expandSpans(raw, spans, args) {
  let out = args;
  for (const [at, axis, form] of spans) {
    if (typeof out[at] !== 'string') continue;
    const [first, last] = spanAddress(raw, out[at], axis);
    out = [...out.slice(0, at), first, form === 'span' ? last : last - first + 1, ...out.slice(at + 1)];
  }
  return out;
}

/**
 * The coordinate form of a call whose cell pairs (`CELL_ARGS`) may be written
 * as A1 strings. The type at the first pair's position picks the form, and the
 * form holds for every pair of the call.
 */
function expandCells(raw, where, pairs, args) {
  if (typeof args[pairs[0][0]] !== 'string') {
    for (const [i, j] of pairs) {
      if (typeof args[i] === 'string' || typeof args[j] === 'string') {
        throw new TypeError(`${where} argument ${i + 1} must be a row number when argument ${pairs[0][0] + 1} is one`);
      }
    }
    return args;
  }
  const out = [];
  let from = 0;
  pairs.forEach(([i], k) => {
    const at = i - k;
    out.push(...args.slice(from, at));
    if (typeof args[at] !== 'string') {
      throw new TypeError(`${where} argument ${at + 1} must be an A1 address when argument ${pairs[0][0] + 1} is one`);
    }
    const cell = cellAddress(raw, args[at], where);
    out.push(cell.firstRow, cell.firstCol);
    from = at + 1;
  });
  out.push(...args.slice(from));
  return out;
}

// `CellRect` arguments (`RECT_ARGS`) given as A1 ranges, alone or in an array.
function expandRects(raw, positions, args) {
  let out = args;
  const rect = (value) => (typeof value === 'string' ? rangeAddress(raw, value) : value);
  for (const at of positions) {
    const value = args[at];
    const next = Array.isArray(value) ? value.map(rect) : rect(value);
    if (next !== value) {
      if (out === args) out = [...args];
      out[at] = next;
    }
  }
  return out;
}

// One A1 reference of a record field or typed argument: `[sheet name or null, rect]`.
function parsedRange(raw, value, label, type) {
  if (typeof value !== 'string') throw new TypeError(`${label} must be a ${type} or an A1 string`);
  const ref = raw.parseReference(value);
  if (ref === null) {
    throw new FormulonError(StatusCode.InvalidArgument, `${label} ${JSON.stringify(value)} is not an A1 range`, '');
  }
  return [ref.sheetName, ref.rect];
}

function unqualifiedRect(raw, value, label) {
  const [sheetName, rect] = parsedRange(raw, value, label, 'CellRect');
  if (sheetName !== null) {
    throw new FormulonError(
      StatusCode.InvalidArgument,
      `${label} ${JSON.stringify(value)} names a sheet; the range is on the sheet it is passed to`,
      `address=${value}`,
    );
  }
  return rect;
}

const RANGE_KINDS = new Set(['CellRect', 'SheetRect', 'CellAddress']);

// The value of a range-typed field or argument in raw form; `sheet` resolves an unqualified `SheetRect`.
function normalizeAddress(raw, kind, value, sheet, label) {
  if (value === null || value === undefined) return value;
  if (kind.endsWith('[]')) {
    const element = kind.slice(0, -2);
    const range = RANGE_KINDS.has(element);
    if (range && typeof value === 'string') return [normalizeAddress(raw, element, value, sheet, `${label}[0]`)];
    if (!Array.isArray(value)) {
      throw new TypeError(`${label} must be an array${range ? ' or one A1 string' : ''}`);
    }
    return value.map((item, i) => normalizeAddress(raw, element, item, sheet, `${label}[${i}]`));
  }
  if (!RANGE_KINDS.has(kind)) return normalizeRecordRanges(raw, value, kind, sheet, label);
  if (isObject(value)) return value;
  if (kind === 'CellRect') return unqualifiedRect(raw, value, label);
  if (kind === 'SheetRect') {
    const [sheetName, rect] = parsedRange(raw, value, label, kind);
    if (sheetName !== null) return { sheet: raw.getSheetIndex(sheetName), rect };
    if (sheet === undefined) {
      throw new FormulonError(
        StatusCode.InvalidArgument,
        `${label} ${JSON.stringify(value)} has no sheet name and the call has no sheet`,
        `address=${value}`,
      );
    }
    return { sheet, rect };
  }
  const rect = unqualifiedRect(raw, value, label);
  if (rect.firstRow !== rect.lastRow || rect.firstCol !== rect.lastCol) {
    throw new FormulonError(StatusCode.InvalidArgument, `${label} ${JSON.stringify(value)} is not one cell`, '');
  }
  return { row: rect.firstRow, col: rect.firstCol };
}

/**
 * A plain copy of every field the raw layer would read from `record`: it reads through the prototype chain, so
 * a getter or a non-enumerable field counts as well as an own enumerable one.
 */
function readableCopy(record) {
  const out = {};
  for (let owner = record; owner !== null && owner !== Object.prototype; owner = Object.getPrototypeOf(owner)) {
    for (const key of Object.getOwnPropertyNames(owner)) {
      if (key !== 'constructor' && !Object.hasOwn(out, key)) out[key] = record[key];
    }
  }
  return out;
}

/**
 * `record` (of input record `type`) in raw form, by `RECORD_RANGE_FIELDS`: a `CellRect` field takes an
 * unqualified A1 range, a `CellRect[]` field an array of them or one string, a `SheetRect` field a qualified
 * range (an unqualified one resolves on `sheet`) and a nested record its own fields. A `CellAddress` field `x`
 * takes one A1 cell and reaches the raw method as `xRow` / `xCol`. Anything but an object passes as it is.
 * `label` words the errors: a field `f` is `<label>f` when it ends in a space, `<label>.f` otherwise.
 */
function normalizeRecordRanges(raw, record, type, sheet, label) {
  const schema = RECORD_RANGE_FIELDS[type];
  if (schema === undefined || !isObject(record)) return record;
  let out = record;
  for (const [field, kind] of Object.entries(schema)) {
    const value = record[field];
    const next = normalizeAddress(raw, kind, value, sheet, label.endsWith(' ') ? label + field : `${label}.${field}`);
    if (kind === 'CellAddress' && isObject(next)) {
      if (out === record) out = readableCopy(record);
      delete out[field];
      out[`${field}Row`] = next.row;
      out[`${field}Col`] = next.col;
    } else if (next !== value) {
      if (out === record) out = readableCopy(record);
      out[field] = next;
    }
  }
  return out;
}

// The arguments `TYPED_ARGS` lists for `where`, in raw form.
function normalizeTypedArgs(raw, sheet, where, typed, args) {
  let out = args;
  for (const [at, kind] of typed) {
    const value = args[at];
    const label = RANGE_KINDS.has(kind.replace(/\[\]$/, '')) ? `${where} argument ${at + 1}` : `${where} `;
    const next = normalizeAddress(raw, kind, value, sheet, label);
    if (next !== value) {
      if (out === args) out = [...args];
      out[at] = next;
    }
  }
  return out;
}

// `args` with each omitted `DEFAULT_ARGS` slot filled; `'today'` is the pinned date, else the host's local date.
function fillDefaults(runtime, raw, defaults, args) {
  let out = args;
  for (const [at, value] of defaults) {
    if (args[at] !== undefined) continue;
    if (out === args) out = [...args];
    out[at] = value === 'today' ? todaySerial(runtime, raw) : value;
  }
  return out;
}

function todaySerial(runtime, raw) {
  const now = new Date();
  const { year, month, day } = raw.getPinnedNow() ?? {
    year: now.getFullYear(),
    month: now.getMonth() + 1,
    day: now.getDate(),
  };
  const civil = { year, month, day, hour: 0, minute: 0, second: 0 };
  return runtime.convertCivilToSerial(civil, raw.getDate1904());
}

// Whether `text` reads as one unqualified A1 cell (`"T2"`), which a table name may not be.
function namesACell(raw, text) {
  let ref;
  try {
    ref = raw.parseReference(text);
  } catch (e) {
    if (e?.name === 'FormulonError' && e.parseError !== undefined) return false;
    throw e;
  }
  if (ref === null || ref.sheetName !== null) return false;
  const { rect } = ref;
  return rect.firstRow === rect.lastRow && rect.firstCol === rect.lastCol;
}

// The identity key of a name under simple case folding, as the engine compares
// sheet names: each code point maps one-to-one, and an expanding mapping
// (`ß` -> `SS`) leaves it as it is.
function foldName(name) {
  let out = '';
  for (const ch of name) {
    const upper = ch.toUpperCase();
    const one = [...upper].length === 1 ? upper : ch;
    const lower = one.toLowerCase();
    out += [...lower].length === 1 ? lower : one;
  }
  return out;
}

/**
 * The index form of a call whose `*Ref` argument (`REF_ARGS`) may be a string or
 * a handle: a string is the first element whose key matches it
 * case-insensitively, as sheet names match, listed afresh on every call; a
 * handle is its element's current index (`handleIndex`, undefined for a value
 * that is not one).
 */
function resolveRef(raw, sheet, where, [at, kind], args, handleIndex) {
  const value = args[at];
  if (typeof value === 'number') return args;
  const spec = REF_KINDS[kind];
  const label = `${where} argument ${at + 1}`;
  const out = [...args];
  if (typeof value !== 'string') {
    const index = spec.handle === undefined ? undefined : handleIndex(kind, value, label);
    if (index === undefined) {
      const handle = spec.handle === undefined ? '' : ` or a ${spec.handle}`;
      throw new TypeError(`${label} must be an index, a ${spec.noun}${handle}`);
    }
    out[at] = index;
    return out;
  }
  const want = foldName(value);
  const n = spec.count(raw, sheet);
  for (let i = 0; i < n; i++) {
    if (foldName(spec.key(raw, sheet, i)) === want) {
      out[at] = i;
      return out;
    }
  }
  throw new FormulonError(
    StatusCode.NotFound,
    `no element has the ${spec.noun} ${JSON.stringify(value)}`,
    `${spec.context}=${value}`,
  );
}

/**
 * `patch` without each flag `x` it passes back beside `xStated: false` at the
 * value the current record (`read()`) reads while `x` is unstated there too, so
 * that `set(get())` leaves an unstated flag unstated.
 */
function withoutUnstated(patch, read) {
  if (!isObject(patch)) return patch;
  const unstated = Object.keys(patch).filter((key) => key.endsWith('Stated') && patch[key] === false);
  if (unstated.length === 0) return patch;
  const current = read();
  const out = { ...patch };
  for (const key of unstated) {
    const field = key.slice(0, -'Stated'.length);
    if (current[key] === false && out[field] === current[field]) delete out[field];
  }
  return out;
}

/**
 * Throws `TypeError: <where> argument <n> has no field <path>` for a key that
 * neither the patch type of `where` (`Class.method`) nor its getter's record
 * declares (`PATCH_KEYS`), and, at each path `PATCH_NESTED_KEYS` lists, for a
 * key of the object there outside its list; `undefined` is ignored.
 */
function refuseUnknownKeys(where, patch) {
  const [position, keys] = PATCH_KEYS[where];
  const nested = PATCH_NESTED_KEYS[where] ?? {};
  const visit = (value, accepted, prefix) => {
    for (const [key, item] of Object.entries(value)) {
      if (item === undefined) continue;
      const path = `${prefix}${key}`;
      if (!accepted.includes(key)) throw new TypeError(`${where} argument ${position + 1} has no field ${path}`);
      if (Object.hasOwn(nested, path) && isObject(item)) visit(item, nested[path], `${path}.`);
    }
  };
  visit(patch, keys, '');
}

/**
 * Checks a hand-written patch of `where` (`Class.method`): an object whose keys
 * are in `PATCH_KEYS`, where a field of `fields` refuses `null`, one of
 * `nullable` takes `null`, and each typed value has its type.
 */
function checkPatch(where, patch, fields, nullable = {}) {
  requireObject(patch, `${where} argument ${PATCH_KEYS[where][0] + 1}`);
  refuseUnknownKeys(where, patch);
  for (const [key, value] of Object.entries(patch)) {
    const type = fields[key] ?? nullable[key];
    if (value === undefined || type === undefined) continue;
    if (value === null && Object.hasOwn(nullable, key)) continue;
    const mismatch =
      type === 'array' ? !Array.isArray(value) : type === 'object' ? !isObject(value) : typeof value !== type;
    if (mismatch) {
      const noun = type === 'array' || type === 'object' ? `an ${type}` : `a ${type}`;
      throw new TypeError(`${where} ${key} must be ${noun}`);
    }
  }
  return patch;
}

/** The keys of `patch` a hand-written setter applies: those of `fields` and `nullable` it gives. */
function appliedFields(patch, fields, nullable = {}) {
  const out = {};
  for (const [key, value] of Object.entries(patch)) {
    if (value !== undefined && (Object.hasOwn(fields, key) || Object.hasOwn(nullable, key))) out[key] = value;
  }
  return out;
}

/**
 * Refuses a key outside `PATCH_KEYS` for `where` (`Class.method`), and `null` on
 * a field of `fields` (`TypeError: <where> <field> must be a <type>`). A
 * non-object patch is left to the raw method.
 */
function checkForwardedPatch(where, patch, fields = {}) {
  if (!isObject(patch)) return;
  refuseUnknownKeys(where, patch);
  for (const [field, type] of Object.entries(fields)) {
    if (patch[field] === null) throw new TypeError(`${where} ${field} must be a ${type}`);
  }
}

function geometryMode(mode, where, position) {
  if (mode === undefined) return GeometryMode.Display;
  requireNumber(mode, `${where} argument ${position}`);
  if (!GEOMETRY_MODES.has(mode)) throw new RangeError(`${where} argument ${position} must be a GeometryMode`);
  return mode;
}

/** The typed raw suffix and arguments for a pivot-cache value (`date` only with a number). */
function cacheValue(where, position, value, options) {
  let date = false;
  if (options !== undefined && options !== null) {
    requireObject(options, `${where} argument ${position + 1}`);
    if (options.date !== undefined) date = requireBoolean(options.date, `${where} date`);
  }
  if (date && typeof value !== 'number') throw new TypeError(`${where} date applies only to a number`);
  if (value === null) return ['Blank', []];
  switch (typeof value) {
    case 'number':
      return [date ? 'Date' : 'Number', [value]];
    case 'string':
      return ['Text', [value]];
    case 'boolean':
      return ['Bool', [value]];
    case 'object':
      if (value.kind === ValueKind.Error && typeof value.errorCode === 'number') return ['Error', [value.errorCode]];
      break;
  }
  throw new TypeError(`${where} argument ${position} must be a number, string, boolean, null or an ErrorValue`);
}

// The payload field and its type of each scalar `Value` kind but Blank.
const SCALAR_VALUE_FIELDS = Object.freeze({
  [ValueKind.Number]: ['number', 'number'],
  [ValueKind.Bool]: ['boolean', 'boolean'],
  [ValueKind.Text]: ['text', 'string'],
  [ValueKind.Error]: ['errorCode', 'number'],
});

/** Throws unless `value` is something `Sheet.setValue` writes as one cell. */
function checkCellValue(value, where) {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return;
  if (value instanceof ExcelError) return;
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) throw new RangeError(`${where} is an invalid Date`);
    return;
  }
  let number = value;
  if (isObject(value)) {
    if (value.kind === ValueKind.Blank) return;
    const [field, type] = SCALAR_VALUE_FIELDS[value.kind] ?? [];
    if (field === undefined || typeof value[field] !== type) {
      throw new TypeError(`${where} must be a number, string, boolean, null, Date, ExcelError or a scalar Value`);
    }
    if (value.kind !== ValueKind.Number) return;
    number = value.number;
  } else if (typeof value !== 'number') {
    throw new TypeError(`${where} must be a number, string, boolean, null, Date, ExcelError or a scalar Value`);
  }
  if (!Number.isFinite(number)) {
    throw new FormulonError(StatusCode.InvalidArgument, `${where} is not a finite number`, `value=${number}`);
  }
}

/** `error` with the cell `row` / `col` in its context, unless the context already names a row. */
function withCell(error, row, col) {
  if (error?.name !== 'FormulonError' || /(?:^| )row=/.test(error.context)) return error;
  const context = [error.context, `row=${row} col=${col}`].filter((part) => part !== '').join(' ');
  return new FormulonError(error.code, error.message, context, error.parseError);
}

/** The native form of a value a read returned; a reference or `LAMBDA` value never reaches a read. */
function nativeOf(value) {
  if (value.kind === ValueKind.Ref || value.kind === ValueKind.Lambda) {
    throw new Error(`formulon: internal error: a read returned a value of kind ${value.kind}`);
  }
  return toNative(value);
}

/** A `RangeCell` from a raw cell entry. */
function rangeCell({ row, col, formula, value }) {
  return { row, col, formula, value: nativeOf(value), kind: value.kind };
}

const WHOLE_SHEET = Object.freeze({ firstRow: 0, firstCol: 0, lastRow: GRID_ROWS - 1, lastCol: GRID_COLS - 1 });
// The cells one `iterRows` block reads: one `getCellsInRange` page.
const ITER_ROWS_CELLS = 65536;
const MS_PER_DAY = 86400000;
const DATE_FORMAT = 'm/d/yyyy';
const DATE_TIME_FORMAT = 'm/d/yyyy h:mm';
// The built-in number formats the en-US codes above name.
const BUILTIN_NUM_FMT_IDS = Object.freeze({ [DATE_FORMAT]: 14, [DATE_TIME_FORMAT]: 22 });

/**
 * The serial of `date`'s UTC fields under the date system `date1904`, and whether it has a time of day. The
 * engine converts the calendar date; the time of day is added here.
 */
function dateSerial(runtime, date, date1904, where) {
  if (Number.isNaN(date.getTime())) throw new RangeError(`${where} is an invalid Date`);
  const civil = {
    year: date.getUTCFullYear(),
    month: date.getUTCMonth() + 1,
    day: date.getUTCDate(),
    hour: 0,
    minute: 0,
    second: 0,
  };
  const seconds = date.getUTCHours() * 3600 + date.getUTCMinutes() * 60 + date.getUTCSeconds();
  const ms = date.getUTCMilliseconds();
  const serial = runtime.convertCivilToSerial(civil, date1904) + seconds / 86400 + ms / MS_PER_DAY;
  return [serial, seconds !== 0 || ms !== 0];
}

/** The `Date` whose UTC fields carry `serial` under the date system `date1904`, rounded to the millisecond. */
function serialDate(runtime, serial, date1904) {
  let days = Math.floor(serial);
  let ms = Math.round((serial - days) * MS_PER_DAY);
  if (ms === MS_PER_DAY) {
    days += 1;
    ms = 0;
  }
  const { year, month, day } = runtime.convertSerialToCivil(days, date1904);
  const date = new Date(0);
  date.setUTCFullYear(year, month - 1, day);
  date.setUTCHours(0, 0, 0, ms);
  return date;
}

// Record keys whose value is a `ColorInput`; a colour string there becomes a `Color`.
const COLOR_INPUT_KEYS = new Set([
  'color',
  'fg',
  'bg',
  'colors',
  'fill',
  'negativeFill',
  'border',
  'negativeBorder',
  'axisColor',
  'fillColor',
  'lineColor',
  'colorSeries',
  'colorNegative',
  'colorAxis',
  'colorMarkers',
  'colorFirst',
  'colorLast',
  'colorHigh',
  'colorLow',
  'fontColor',
]);

const asColor = (value, label) => (typeof value === 'string' ? formatColor(value, label) : value);

/** `value` with every colour string at a `COLOR_INPUT_KEYS` key as a `Color`; `value` itself when it has none. */
function colorInputs(value, label) {
  if (Array.isArray(value)) {
    let out = value;
    value.forEach((item, i) => {
      const next = colorInputs(item, `${label}[${i}]`);
      if (next !== item) {
        if (out === value) out = [...value];
        out[i] = next;
      }
    });
    return out;
  }
  if (!isObject(value) || Object.getPrototypeOf(value) !== Object.prototype) return value;
  let out = value;
  for (const [key, item] of Object.entries(value)) {
    const path = `${label}.${key}`;
    let next;
    if (COLOR_INPUT_KEYS.has(key) && typeof item === 'string') next = formatColor(item, path);
    else if (key === 'colors' && Array.isArray(item)) next = item.map((c, i) => asColor(c, `${path}[${i}]`));
    else next = colorInputs(item, path);
    if (next !== item && !(Array.isArray(next) && next.every((c, i) => c === item[i]))) {
      if (out === value) out = { ...value };
      out[key] = next;
    }
  }
  return out;
}

/** `args` with the colour strings of its record arguments as `Color`s. */
function colorArgs(args, where) {
  let out = args;
  args.forEach((arg, i) => {
    const next = colorInputs(arg, `${where} argument ${i + 1}`);
    if (next !== arg) {
      if (out === args) out = [...args];
      out[i] = next;
    }
  });
  return out;
}

// Every populated cell of `rect`, read in pages.
function forEachCellIn(raw, sheet, rect, visit) {
  let page = raw.getCellsInRange(sheet, rect, 0);
  for (;;) {
    for (const cell of page.cells) visit(cell);
    if (page.nextCursor === null) return;
    page = raw.getCellsInRange(sheet, rect, page.nextCursor);
  }
}

function requireValueField(value, field, type) {
  if (typeof value[field] !== type) {
    throw new TypeError(`Sheet.setValue argument 3 must carry a ${type} \`${field}\` for its kind`);
  }
  return value[field];
}

// ---------------------------------------------------------------------------
// Cell formats: `Sheet.getCellFormat` / `Sheet.setCellFormat` over the style
// records. tools/api/cell_format_cases.json pins the behaviour for the JS and
// Python facades alike.
// ---------------------------------------------------------------------------

const enumField = (table, nullable = false) => ({ kind: 'enum', max: Math.max(...Object.values(table)), nullable });
const FORMAT_BOOLEAN = Object.freeze({ kind: 'boolean' });
const FORMAT_FLAG = Object.freeze({ kind: 'boolean', nullable: true });
const FORMAT_COLOR = Object.freeze({ kind: 'color' });
const FORMAT_SIDE = Object.freeze({ kind: 'side' });

// How each key of each group of a `CellFormatPatch` is checked; a nullable key
// takes `null` to clear the attribute. The accepted keys are `PATCH_NESTED_KEYS`.
const FORMAT_FIELDS = Object.freeze({
  font: {
    name: { kind: 'string' },
    size: { kind: 'size', nullable: true },
    bold: FORMAT_FLAG,
    italic: FORMAT_FLAG,
    strike: FORMAT_FLAG,
    outline: FORMAT_FLAG,
    shadow: FORMAT_FLAG,
    condense: FORMAT_FLAG,
    extend: FORMAT_FLAG,
    underline: enumField(UnderlineStyle),
    vertAlign: enumField(FontVerticalAlign),
    family: enumField(FontFamily, true),
    charset: { kind: 'enum', max: 255, nullable: true },
    scheme: enumField(FontScheme),
    color: FORMAT_COLOR,
  },
  fill: {
    pattern: enumField(FillPattern),
    fg: FORMAT_COLOR,
    bg: FORMAT_COLOR,
    gradient: { kind: 'object' },
  },
  border: {
    left: FORMAT_SIDE,
    right: FORMAT_SIDE,
    top: FORMAT_SIDE,
    bottom: FORMAT_SIDE,
    diagonal: FORMAT_SIDE,
    diagonalUp: FORMAT_BOOLEAN,
    diagonalDown: FORMAT_BOOLEAN,
  },
  side: { style: enumField(BorderStyle), color: FORMAT_COLOR },
  alignment: {
    horizontal: enumField(HorizontalAlign),
    vertical: enumField(VerticalAlign),
    wrapText: FORMAT_BOOLEAN,
    shrinkToFit: FORMAT_BOOLEAN,
    indent: { kind: 'enum', max: 255 },
    textRotation: { kind: 'rotation' },
    readingOrder: enumField(ReadingOrder),
    justifyLastLine: FORMAT_BOOLEAN,
    relativeIndent: { kind: 'signed' },
  },
  protection: { locked: FORMAT_BOOLEAN, hidden: FORMAT_BOOLEAN },
});

// The `CellXfInput` field each alignment key writes, and the presence flag it sets when it has one.
const ALIGNMENT_XF = Object.freeze({
  horizontal: ['horizontalAlign', 'hasHorizontalAlign'],
  vertical: ['verticalAlign', 'hasVerticalAlign'],
  wrapText: ['wrapText', 'hasWrapText'],
  justifyLastLine: ['justifyLastLine', 'hasJustifyLastLine'],
  shrinkToFit: ['shrinkToFit'],
  indent: ['indent'],
  textRotation: ['textRotation'],
  readingOrder: ['readingOrder'],
  relativeIndent: ['relativeIndent'],
});
const OPTIONAL_XF_FIELDS = Object.freeze(['textRotation', 'indent', 'shrinkToFit', 'readingOrder', 'relativeIndent']);
const APPLY_FLAGS = Object.freeze({
  numberFormat: 'applyNumberFormat',
  font: 'applyFont',
  fill: 'applyFill',
  border: 'applyBorder',
  alignment: 'applyAlignment',
  protection: 'applyProtection',
});
const COLOR_TEXT = /^#(?:[0-9a-fA-F]{6}|[0-9a-fA-F]{8})$/;
const INT32_MAX = 0x7fffffff;

function outOfRange(label, value) {
  return new FormulonError(StatusCode.InvalidArgument, `${label} is out of range`, `value=${value}`);
}

// A non-negative integer: a fraction, a negative number or NaN is a RangeError.
function formatCount(value, label) {
  requireNumber(value, label);
  if (!Number.isInteger(value) || value < 0) throw new RangeError(`${label} must be a non-negative integer`);
  return value;
}

/** A `ColorInput` as a `Color`: `#RRGGBB` (opaque) or `#AARRGGBB`, or a `Color` record. */
function formatColor(value, label) {
  if (typeof value === 'string') {
    if (!COLOR_TEXT.test(value)) throw new TypeError(`${label} must be "#RRGGBB" or "#AARRGGBB"`);
    const hex = value.length === 7 ? `FF${value.slice(1)}` : value.slice(1);
    return { kind: ColorKind.Rgb, rgb: Number.parseInt(hex, 16), theme: 0, tint: 0, indexed: 0 };
  }
  if (!isObject(value)) throw new TypeError(`${label} must be a colour string or a Color`);
  const kind = formatCount(value.kind, `${label}.kind`);
  if (kind > ColorKind.Auto) throw outOfRange(`${label}.kind`, kind);
  return value;
}

function formatValue(spec, value, label) {
  if (value === null) {
    if (spec.nullable) return null;
    throw new TypeError(`${label} must not be null`);
  }
  switch (spec.kind) {
    case 'string':
      return requireString(value, label);
    case 'boolean':
      return requireBoolean(value, label);
    case 'object':
      return requireObject(value, label);
    case 'color':
      return formatColor(value, label);
    case 'side':
      return formatGroup('side', value, label);
    case 'size':
      requireNumber(value, label);
      if (!Number.isFinite(value) || value < 0) throw new RangeError(`${label} must be finite and not negative`);
      return value;
    case 'enum':
      if (formatCount(value, label) > spec.max) throw outOfRange(label, value);
      return value;
    case 'rotation':
      if (formatCount(value, label) > 180 && value !== 255) throw outOfRange(label, value);
      return value;
    case 'signed':
      requireNumber(value, label);
      if (!Number.isInteger(value)) throw new RangeError(`${label} must be an integer`);
      if (value > INT32_MAX || value < -INT32_MAX - 1) throw outOfRange(label, value);
      return value;
  }
  return value;
}

// One group of a `CellFormatPatch`, checked: its defined keys, colours as `Color` records. Its keys were
// checked against `PATCH_NESTED_KEYS` with the whole patch.
function formatGroup(group, value, label) {
  requireObject(value, label);
  const fields = FORMAT_FIELDS[group];
  const out = {};
  for (const [key, item] of Object.entries(value)) {
    if (item === undefined) continue;
    out[key] = fields[key] === undefined ? item : formatValue(fields[key], item, `${label}.${key}`);
  }
  return out;
}

/** Checks a whole `CellFormatPatch` before anything is written; the result holds only its given keys. */
function checkCellFormatPatch(patch) {
  const where = 'Sheet.setCellFormat';
  requireObject(patch, `${where} argument 2`);
  refuseUnknownKeys(where, patch);
  const out = {};
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined) continue;
    if (key === 'numberFormat') {
      requireString(value, `${where} numberFormat`);
      if (value === '') {
        throw new FormulonError(StatusCode.InvalidArgument, `${where} numberFormat is empty`, 'number_format=');
      }
      out.numberFormat = value;
      continue;
    }
    out[key] = formatGroup(key, value, `${where} ${key}`);
  }
  return out;
}

function mergeFont(current, patch) {
  return { ...current, ...patch };
}

// Setting a colour on a fill without a pattern makes it solid, Excel's "fill
// colour"; any fill key drops a gradient unless the patch names one.
function mergeFill(current, patch) {
  const out = { ...current };
  if (patch.gradient === undefined) delete out.gradient;
  else out.gradient = patch.gradient;
  for (const key of ['pattern', 'fg', 'bg']) if (patch[key] !== undefined) out[key] = patch[key];
  const colored = patch.fg !== undefined || patch.bg !== undefined;
  if (colored && patch.pattern === undefined && out.pattern === FillPattern.None) out.pattern = FillPattern.Solid;
  return out;
}

function mergeBorder(current, patch) {
  const out = { ...current };
  for (const [key, value] of Object.entries(patch)) {
    out[key] = key === 'diagonalUp' || key === 'diagonalDown' ? value : { ...current[key], ...value };
  }
  return out;
}

const present = (value) => value !== null && value !== undefined;

/**
 * The `CellXfInput` that rebuilds `xf` as `addCellXf` reads it: an absent
 * attribute (`null`) keeps its input default with its presence flag false, and
 * an optional alignment attribute is copied only when the record has it.
 */
function cellXfInput(xf) {
  const input = {
    fontIndex: xf.fontIndex,
    fillIndex: xf.fillIndex,
    borderIndex: xf.borderIndex,
    numFmtId: xf.numFmtId,
    horizontalAlign: xf.horizontalAlign ?? HorizontalAlign.General,
    hasHorizontalAlign: present(xf.horizontalAlign),
    verticalAlign: xf.verticalAlign ?? VerticalAlign.Bottom,
    hasVerticalAlign: present(xf.verticalAlign),
    wrapText: xf.wrapText ?? false,
    hasWrapText: present(xf.wrapText),
    justifyLastLine: xf.justifyLastLine ?? false,
    hasJustifyLastLine: present(xf.justifyLastLine),
    applyNumberFormat: xf.applyNumberFormat,
    applyFont: xf.applyFont,
    applyFill: xf.applyFill,
    applyBorder: xf.applyBorder,
    applyAlignment: xf.applyAlignment,
    applyProtection: xf.applyProtection,
    quotePrefix: xf.quotePrefix,
    hasProtection: present(xf.locked) || present(xf.hidden),
    locked: xf.locked ?? true,
    hidden: xf.hidden ?? false,
  };
  for (const key of OPTIONAL_XF_FIELDS) if (present(xf[key])) input[key] = xf[key];
  if (present(xf.xfId)) input.xfId = xf.xfId;
  input.hasAlignment =
    input.hasHorizontalAlign ||
    input.hasVerticalAlign ||
    input.hasWrapText ||
    input.hasJustifyLastLine ||
    OPTIONAL_XF_FIELDS.some((key) => key in input);
  return input;
}

/** `xf` with `patch` merged in, as the `CellXfInput` of the new record. */
function patchedXf(raw, xf, patch, numFmtId) {
  const input = cellXfInput(xf);
  for (const group of Object.keys(patch)) input[APPLY_FLAGS[group]] = true;
  if (numFmtId !== undefined) input.numFmtId = numFmtId;
  if (patch.font !== undefined) input.fontIndex = raw.addFont(mergeFont(raw.getFont(xf.fontIndex), patch.font));
  if (patch.fill !== undefined) input.fillIndex = raw.addFill(mergeFill(raw.getFill(xf.fillIndex), patch.fill));
  if (patch.border !== undefined) {
    input.borderIndex = raw.addBorder(mergeBorder(raw.getBorder(xf.borderIndex), patch.border));
  }
  if (patch.alignment !== undefined) {
    for (const [key, value] of Object.entries(patch.alignment)) {
      const [field, flag] = ALIGNMENT_XF[key];
      input[field] = value;
      if (flag !== undefined) input[flag] = true;
    }
    input.hasAlignment = true;
  }
  if (patch.protection !== undefined) {
    input.locked = patch.protection.locked ?? input.locked;
    input.hidden = patch.protection.hidden ?? input.hidden;
    input.hasProtection = true;
  }
  return input;
}

/** The `CellFormat` of the cell XF record `xf`. */
function cellFormatOf(raw, xf) {
  const code = xf.numFmtId === 0 ? null : raw.getNumFmtString(xf.numFmtId);
  return {
    numberFormat: code ?? 'General',
    font: raw.getFont(xf.fontIndex),
    fill: raw.getFill(xf.fillIndex),
    border: raw.getBorder(xf.borderIndex),
    alignment: {
      horizontal: xf.horizontalAlign ?? HorizontalAlign.General,
      vertical: xf.verticalAlign ?? VerticalAlign.Bottom,
      wrapText: xf.wrapText ?? false,
      shrinkToFit: xf.shrinkToFit ?? false,
      indent: xf.indent ?? 0,
      textRotation: xf.textRotation ?? 0,
      readingOrder: xf.readingOrder ?? ReadingOrder.Context,
      justifyLastLine: xf.justifyLastLine ?? false,
      relativeIndent: xf.relativeIndent ?? 0,
    },
    protection: { locked: xf.locked ?? true, hidden: xf.hidden ?? false },
  };
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

// A raw call that resolves a runtime id, with its `NotFound` reworded as the handle's object being gone.
function resolveRuntimeId(resolve, message) {
  try {
    return resolve();
  } catch (e) {
    if (e?.name === 'FormulonError' && e.code === StatusCode.NotFound) {
      throw new FormulonError(StatusCode.NotFound, message, '');
    }
    throw e;
  }
}

/**
 * Builds the public `Workbook` and `Sheet` classes and the handle classes
 * (`Table`, `Pivot`, `PivotCache`, `Drawing` and its `Picture`, `Chart`,
 * `Shape` and `Slicer` subclasses) over a runtime's raw `Workbook` class. Each
 * call returns new classes, so a workbook of one `createFormulon()` result is
 * not an `instanceof` the `Workbook` of another.
 *
 * `Workbook` holds the workbook-scope ops and `Sheet` the sheet-scope ones; a
 * `Sheet` method calls the raw method with the sheet's current index first. The
 * workbook keeps one `Sheet` per sheet in sheet order and updates that list
 * after every op in `SHEET_STRUCTURE_OPS`, so a handle follows its sheet across
 * inserts, moves and renames, and a removed sheet's handle throws `NotFound`.
 *
 * A handle method calls its raw method with the arguments the handle supplies
 * first. Tables and pivots are named by the engine's runtime id, resolved to the
 * current index on every call; pivot caches and drawing objects by their cache
 * and object ids. The parent keeps one handle per id and never drops it, so
 * every lookup of an object returns the same handle, and a removed object's
 * handle throws `NotFound`.
 *
 * Cell, range, span and `*Ref` arguments also take their string forms (see
 * `installArgumentForms`).
 *
 * @param {Function} Raw The runtime's raw `Workbook` class.
 * @param {FinalizationRegistry} [registry] See `buildFormulon`.
 * @param {object} runtime The runtime, for its civil-date conversions.
 * @returns {Function} The public `Workbook` class.
 */
function makeFacade(Raw, registry, runtime) {
  const token = Symbol('formulon.facade');
  let rawOf;
  let liveRaw;
  let isDisposed;
  // Facade calls in flight per workbook; `dispose()` refuses while one runs.
  const activeCalls = new WeakMap();
  let sheetIndex;
  let sheetState;
  let insertSheet;
  let moveSheet;
  let pivotCacheOf;
  let tableState;
  let pivotState;
  let cacheState;
  let drawingState;
  let ooxmlState;

  class Table {
    #workbook;
    #id;

    constructor(key, workbook, id) {
      if (key !== token) throw new TypeError('Table has no public constructor; use Workbook.table');
      this.#workbook = workbook;
      this.#id = id;
    }

    static {
      tableState = {
        create: (workbook, id) => new Table(token, workbook, id),
        locate: (table) => table.#locate(),
      };
    }

    // The raw workbook and the table's current index.
    #locate() {
      const raw = liveRaw(this.#workbook);
      return [raw, resolveRuntimeId(() => raw.resolveTableRuntimeId(this.#id), 'the table was removed')];
    }

    get workbook() {
      return this.#workbook;
    }

    get index() {
      return this.#locate()[1];
    }

    get name() {
      const [raw, index] = this.#locate();
      return raw.getTable(index).name;
    }

    get displayName() {
      const [raw, index] = this.#locate();
      return raw.getTable(index).displayName;
    }

    // `range` reaches the raw method as its `rect`; a `styleName` of `null` as the empty name that removes it.
    update(patch) {
      const where = 'Table.update';
      const fields = PATCH_FIELDS[where];
      const nullable = { styleName: 'string' };
      checkPatch(where, patch, fields, nullable);
      if (patch.range === null) throw new TypeError(`${where} range must be a SheetRect or an A1 string`);
      const [raw, index] = this.#locate();
      const input = appliedFields(patch, fields, nullable);
      if (input.styleName === null) input.styleName = '';
      if (patch.range !== undefined) input.rect = tableRange(raw, raw.getTable(index).range.sheet, patch.range, where);
      raw.updateTable(index, input);
    }
  }

  class Pivot {
    #sheet;
    #id;

    constructor(key, sheet, id) {
      if (key !== token) throw new TypeError('Pivot has no public constructor; use Sheet.pivot');
      this.#sheet = sheet;
      this.#id = id;
    }

    static {
      pivotState = {
        create: (sheet, id) => new Pivot(token, sheet, id),
        locate: (pivot) => pivot.#locate(),
      };
    }

    // The raw workbook, the sheet's current index and the pivot's current index on it.
    #locate() {
      const sheet = sheetIndex(this.#sheet);
      const raw = rawOf(this.#sheet.workbook);
      return [raw, sheet, resolveRuntimeId(() => raw.resolvePivotRuntimeId(sheet, this.#id), 'the pivot was removed')];
    }

    get sheet() {
      return this.#sheet;
    }

    get index() {
      return this.#locate()[2];
    }

    getProperties() {
      const [raw, sheet, pivot] = this.#locate();
      return { name: raw.getPivotName(sheet, pivot), cacheId: raw.getPivotCacheId(sheet, pivot) };
    }

    getCache() {
      const [raw, sheet, pivot] = this.#locate();
      return pivotCacheOf(this.#sheet.workbook, raw.getPivotCacheId(sheet, pivot));
    }

    showDetails(row, col) {
      const [raw, sheet, pivot] = this.#locate();
      return insertSheet(this.#sheet.workbook, raw.showPivotDetails(sheet, pivot, row, col));
    }

    updateField(field, patch) {
      const where = 'Pivot.updateField';
      checkPatch(where, patch, PATCH_FIELDS[where], { customName: 'string', numberFormat: 'string' });
      const [raw, sheet, pivot] = this.#locate();
      const setters = [
        ['customName', 'setPivotFieldCustomName'],
        ['numberFormat', 'setPivotFieldNumberFormat'],
        ['defaultSubtotal', 'setPivotFieldDefaultSubtotal'],
        ['insertBlankRow', 'setPivotFieldInsertBlankRow'],
        ['fillDownLabels', 'setPivotFieldFillDownLabels'],
        ['showAll', 'setPivotFieldShowAll'],
        ['subtotalTop', 'setPivotFieldSubtotalTop'],
      ];
      for (const [key, setter] of setters) {
        if (patch[key] !== undefined) raw[setter](sheet, pivot, field, patch[key] ?? '');
      }
    }

    updateItem(field, item, patch) {
      const where = 'Pivot.updateItem';
      checkPatch(where, patch, PATCH_FIELDS[where], { caption: 'string' });
      const [raw, sheet, pivot] = this.#locate();
      const setters = [
        ['visible', 'setPivotFieldItemVisible'],
        ['showDetail', 'setPivotFieldItemShowDetail'],
        ['caption', 'setPivotFieldItemCaption'],
      ];
      for (const [key, setter] of setters) {
        if (patch[key] !== undefined) raw[setter](sheet, pivot, field, item, patch[key] ?? '');
      }
    }

    setStyleInfo(patch) {
      const where = 'Pivot.setStyleInfo';
      const fields = PATCH_FIELDS[where];
      checkPatch(where, patch, fields, { name: 'string' });
      const [raw, sheet, pivot] = this.#locate();
      const current = raw.getPivotStyleInfo(sheet, pivot) ?? {};
      const info = {};
      const name = patch.name === undefined ? current.name || undefined : (patch.name ?? undefined);
      if (name !== undefined) info.name = name;
      for (const key of Object.keys(fields)) info[key] = patch[key] ?? current[key] ?? false;
      raw.setPivotStyleInfo(sheet, pivot, info);
    }

    setFieldSort(field, order, options) {
      const where = 'Pivot.setFieldSort';
      let byField = '';
      if (options !== undefined && options !== null) {
        requireObject(options, `${where} argument 3`);
        if (options.byField !== undefined) byField = requireString(options.byField, `${where} byField`);
      }
      const [raw, sheet, pivot] = this.#locate();
      raw.setPivotFieldSort(sheet, pivot, field, order, byField);
    }

    addFieldItem(field, name, options) {
      const where = 'Pivot.addFieldItem';
      let visible = true;
      if (options !== undefined && options !== null) {
        requireObject(options, `${where} argument 3`);
        if (options.visible !== undefined) visible = requireBoolean(options.visible, `${where} visible`);
      }
      const [raw, sheet, pivot] = this.#locate();
      return raw.addPivotFieldItem(sheet, pivot, field, name, visible);
    }

    addFieldSharedItem(field, cacheIndex, options) {
      const where = 'Pivot.addFieldSharedItem';
      let visible = true;
      if (options !== undefined && options !== null) {
        requireObject(options, `${where} argument 3`);
        if (options.visible !== undefined) visible = requireBoolean(options.visible, `${where} visible`);
      }
      const [raw, sheet, pivot] = this.#locate();
      return raw.addPivotFieldSharedItem(sheet, pivot, field, cacheIndex, visible);
    }

    setGrandTotals(patch) {
      const where = 'Pivot.setGrandTotals';
      requireObject(patch, `${where} argument 1`);
      refuseUnknownKeys(where, patch);
      const [raw, sheet, pivot] = this.#locate();
      const current =
        patch.rows === undefined || patch.columns === undefined ? raw.getPivotGrandTotals(sheet, pivot) : {};
      const rows = patchField(patch, 'rows', current.rows, requireBoolean, where);
      const columns = patchField(patch, 'columns', current.columns, requireBoolean, where);
      raw.setPivotGrandTotals(sheet, pivot, rows, columns);
    }

    setFieldDateGroup(field, group) {
      const where = 'Pivot.setFieldDateGroup';
      requireObject(group, `${where} argument 2`);
      const optional = (key) => (group[key] === undefined ? undefined : requireNumber(group[key], `${where} ${key}`));
      const args = [
        requireInteger(group.periods, `${where} periods`),
        group.calendar === undefined ? PivotCalendar.Gregorian : requireInteger(group.calendar, `${where} calendar`),
      ];
      const tail = [optional('intervalDays'), optional('start'), optional('end')];
      while (tail.length > 0 && tail.at(-1) === undefined) tail.pop();
      const [raw, sheet, pivot] = this.#locate();
      raw.setPivotFieldDateGroup(sheet, pivot, field, ...args, ...tail);
    }
  }

  class PivotCache {
    #workbook;
    #cacheId;

    constructor(key, workbook, cacheId) {
      if (key !== token) throw new TypeError('PivotCache has no public constructor; use Workbook.pivotCache');
      this.#workbook = workbook;
      this.#cacheId = cacheId;
    }

    static {
      cacheState = {
        create: (workbook, cacheId) => new PivotCache(token, workbook, cacheId),
        raw: (cache) => cache.#raw(),
      };
    }

    // The raw workbook, once the cache is known to exist: the raw methods report a missing cache id as an
    // invalid argument, not as a removed cache.
    #raw() {
      const raw = liveRaw(this.#workbook);
      if (!raw.listPivotCaches().includes(this.#cacheId)) {
        throw new FormulonError(StatusCode.NotFound, 'the pivot cache was removed', `cache_id=${this.#cacheId}`);
      }
      return raw;
    }

    get workbook() {
      return this.#workbook;
    }

    get cacheId() {
      return this.#cacheId;
    }

    listFields() {
      const raw = this.#raw();
      const id = this.#cacheId;
      const n = raw.getPivotCacheFieldCount(id);
      const fields = [];
      for (let i = 0; i < n; i++) {
        fields.push({ name: raw.getPivotCacheFieldName(id, i), formula: raw.getPivotCacheFieldFormula(id, i) });
      }
      return fields;
    }

    listFieldSharedItems(field) {
      const raw = this.#raw();
      const id = this.#cacheId;
      const n = raw.getPivotCacheFieldSharedItemCount(id, field);
      const items = [];
      for (let i = 0; i < n; i++) {
        items.push({
          value: raw.getPivotCacheFieldSharedItemValue(id, field, i),
          isDate: raw.isPivotCacheFieldSharedItemDate(id, field, i),
        });
      }
      return items;
    }

    addFieldSharedItem(field, value, options) {
      const [kind, args] = cacheValue('PivotCache.addFieldSharedItem', 2, value, options);
      this.#raw()[`addPivotCacheFieldSharedItem${kind}`](this.#cacheId, field, ...args);
    }

    setRecordValue(record, field, value, options) {
      const [kind, args] = cacheValue('PivotCache.setRecordValue', 3, value, options);
      this.#raw()[`setPivotCacheRecord${kind}`](this.#cacheId, record, field, ...args);
    }

    // `null` removes the source; an object writes one with the attributes it gives (`null` is one left off).
    setWorksheetSource(source) {
      const where = 'PivotCache.setWorksheetSource';
      if (source === null) return this.#raw().setPivotCacheWorksheetSource(this.#cacheId, { present: false });
      requireObject(source, `${where} argument 1`);
      const keys = ['ref', 'sheet', 'name'];
      const input = { present: true };
      // Each value is read once.
      for (const [key, value] of Object.entries(source)) {
        if (value === undefined) continue;
        if (!keys.includes(key)) throw new TypeError(`${where} argument 1 has no field ${key}`);
        if (value !== null) input[key] = requireString(value, `${where} ${key}`);
      }
      this.#raw().setPivotCacheWorksheetSource(this.#cacheId, input);
    }
  }

  class Drawing {
    #sheet;
    #objectId;
    #kind;

    constructor(key, sheet, objectId, kind) {
      if (key !== token) throw new TypeError('Drawing has no public constructor; use Sheet.drawing');
      this.#sheet = sheet;
      this.#objectId = objectId;
      this.#kind = kind;
    }

    static {
      drawingState = {
        create: (cls, sheet, objectId, kind) => new cls(token, sheet, objectId, kind),
        locate: (drawing) => drawing.#locate(),
      };
    }

    // The raw workbook, the sheet's current index and the object id, once the object is known to exist: the raw
    // methods report a missing object id as an invalid argument, not as a removed object.
    #locate() {
      const sheet = sheetIndex(this.#sheet);
      const raw = rawOf(this.#sheet.workbook);
      if (raw.getDrawingObject(sheet, this.#objectId) === null) {
        throw new FormulonError(StatusCode.NotFound, 'the drawing was removed', `object_id=${this.#objectId}`);
      }
      return [raw, sheet, this.#objectId];
    }

    get sheet() {
      return this.#sheet;
    }

    get objectId() {
      return this.#objectId;
    }

    get kind() {
      return this.#kind;
    }
  }

  class Picture extends Drawing {}

  class Chart extends Drawing {
    setLegend(patch) {
      const where = 'Chart.setLegend';
      const [raw, sheet, id] = drawingState.locate(this);
      if (patch === null) return raw.setChartLegend(sheet, id, -1, false);
      checkPatch(where, patch, PATCH_FIELDS[where]);
      const current = raw.getChart(sheet, id)?.legend ?? DEFAULT_LEGEND;
      raw.setChartLegend(
        sheet,
        id,
        patch.position ?? current.position,
        patch.overlay ?? current.overlay,
        patch.hiddenSeries ?? current.hiddenSeries,
      );
    }

    setSource(source) {
      const where = 'Chart.setSource';
      requireObject(source, `${where} argument 1`);
      const flag = (field) =>
        source[field] === undefined || source[field] === null
          ? undefined
          : requireBoolean(source[field], `${where} ${field}`);
      const args = [
        requireString(source.source, `${where} source`),
        source.direction === undefined
          ? ChartSeriesDirection.Auto
          : requireInteger(source.direction, `${where} direction`),
      ];
      const headerRow = flag('headerRow');
      const headerColumn = flag('headerColumn');
      if (headerRow !== undefined || headerColumn !== undefined) args.push(headerRow, headerColumn);
      const [raw, sheet, id] = drawingState.locate(this);
      raw.setChartSource(sheet, id, ...args);
    }
  }

  // The raw update applies exactly the keys it is given; the record keys it cannot set are dropped first.
  class Shape extends Drawing {
    update(patch) {
      const where = 'Shape.update';
      patch = colorInputs(patch, `${where} argument 1`);
      checkDrawingPatch(where, patch);
      const input = appliedFields(patch, PATCH_FIELDS[where], PATCH_NULLABLE[where]);
      for (const key of ['fillKind', 'lineKind', 'fillColor', 'lineColor']) if (input[key] === null) delete input[key];
      const [raw, sheet, id] = drawingState.locate(this);
      raw.updateShape(sheet, id, input);
    }
  }

  class Slicer extends Drawing {
    update(patch) {
      const where = 'Slicer.update';
      checkDrawingPatch(where, patch);
      const input = appliedFields(patch, PATCH_FIELDS[where], PATCH_NULLABLE[where]);
      const [raw, sheet, id] = drawingState.locate(this);
      raw.updateSlicer(sheet, id, input);
    }
  }

  // The handle class of a drawing object, from its `DrawingProperties` record.
  function drawingClass(raw, sheet, record) {
    switch (record.kind) {
      case DrawingObjectKind.Picture:
        return Picture;
      case DrawingObjectKind.Chart:
        return Chart;
      case DrawingObjectKind.Shape:
      case DrawingObjectKind.Connector:
      case DrawingObjectKind.Group:
        return Shape;
    }
    return raw.getSlicer(sheet, record.objectId) === null ? Drawing : Slicer;
  }

  // The current index of a `TableRef` / `PivotRef` handle `value` passed to a method of `self`; undefined for
  // anything else.
  function handleIndex(self, kind, value, label) {
    if (kind === 'table' && value instanceof Table) {
      const workbook = self instanceof Workbook ? self : self.workbook;
      if (value.workbook !== workbook) {
        throw new FormulonError(StatusCode.InvalidArgument, `${label} is a Table of another workbook`, '');
      }
      return value.index;
    }
    if (kind === 'pivot' && value instanceof Pivot) {
      if (value.sheet !== self) {
        throw new FormulonError(StatusCode.InvalidArgument, `${label} is a Pivot of another sheet`, '');
      }
      return value.index;
    }
    return undefined;
  }

  class Sheet {
    #workbook;
    #position;
    #name;
    #pivots = new Map();
    #drawings = new Map();
    #ooxml;

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

    rename(newName) {
      const raw = rawOf(this.#workbook);
      const index = this.#current();
      raw.renameSheet(index, newName);
      this.#name = raw.getSheetName(index);
    }

    move(toIndex) {
      moveSheet(this.#workbook, this.#current(), toIndex);
    }

    get ooxml() {
      this.#current();
      this.#ooxml ??= ooxmlState.sheet(this);
      return this.#ooxml;
    }

    getFormula(row, col) {
      return rawOf(this.#workbook).getFormula(this.#current(), row, col);
    }

    setFormula(row, col, formula) {
      return rawOf(this.#workbook).setFormula(this.#current(), row, col, formula);
    }

    getValue(row, col) {
      return nativeOf(rawOf(this.#workbook).getValue(this.#current(), row, col));
    }

    getTypedValue(row, col) {
      return rawOf(this.#workbook).getValue(this.#current(), row, col);
    }

    getValues(range) {
      return this.#typedGrid(range, 'Sheet.getValues').map((line) => line.map(nativeOf));
    }

    getTypedValues(range) {
      return this.#typedGrid(range, 'Sheet.getTypedValues');
    }

    // The tagged values of `range` as rows, a blank cell `BLANK`.
    #typedGrid(range, where) {
      requireObject(range, `${where} argument 1`);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const rows = range.lastRow - range.firstRow + 1;
      const cols = range.lastCol - range.firstCol + 1;
      if (rows * cols > MAX_VALUES_CELLS) {
        throw new FormulonError(
          StatusCode.InvalidArgument,
          `${where} reads at most ${MAX_VALUES_CELLS} cells`,
          `cells=${rows * cols}`,
        );
      }
      // Allocated once the engine has accepted the rectangle.
      const blank = () => Array.from({ length: rows }, () => new Array(cols).fill(BLANK));
      let values = null;
      forEachCellIn(raw, sheet, range, (cell) => {
        values ??= blank();
        values[cell.row - range.firstRow][cell.col - range.firstCol] = cell.value;
      });
      return values ?? blank();
    }

    getCellsInRange(range, ...rest) {
      const page = rawOf(this.#workbook).getCellsInRange(this.#current(), range, ...rest);
      return { cells: page.cells.map(rangeCell), nextCursor: page.nextCursor };
    }

    listCells(range) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const rect = range ?? { firstRow: 0, firstCol: 0, lastRow: GRID_ROWS - 1, lastCol: GRID_COLS - 1 };
      const cells = [];
      forEachCellIn(raw, sheet, rect, (cell) => cells.push(rangeCell(cell)));
      return cells;
    }

    getUsedRange() {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      let used = null;
      forEachCellIn(raw, sheet, WHOLE_SHEET, ({ row, col }) => {
        if (used === null) {
          used = { firstRow: row, firstCol: col, lastRow: row, lastCol: col };
          return;
        }
        // Cells arrive row-major, so only the columns and the last row can widen.
        used.firstCol = Math.min(used.firstCol, col);
        used.lastCol = Math.max(used.lastCol, col);
        used.lastRow = row;
      });
      return used;
    }

    // Reads the range in blocks of whole rows, each within one `getCellsInRange` page.
    *iterRows(range) {
      const rect = range === undefined || range === null ? this.getUsedRange() : range;
      if (rect === null) return;
      requireObject(rect, 'Sheet.iterRows argument 1');
      const cols = rect.lastCol - rect.firstCol + 1;
      const block = Math.max(1, Math.floor(ITER_ROWS_CELLS / cols));
      for (let first = rect.firstRow; first <= rect.lastRow; first += block) {
        const last = Math.min(rect.lastRow, first + block - 1);
        const grid = this.#typedGrid({ ...rect, firstRow: first, lastRow: last }, 'Sheet.iterRows');
        for (const line of grid) yield line.map(nativeOf);
      }
    }

    getDate(row, col) {
      const raw = rawOf(this.#workbook);
      const value = raw.getValue(this.#current(), row, col);
      if (value.kind === ValueKind.Blank) return null;
      if (value.kind !== ValueKind.Number) {
        throw new FormulonError(
          StatusCode.InvalidArgument,
          'Sheet.getDate reads a cell that holds no number',
          `row=${row} col=${col}`,
        );
      }
      return serialDate(runtime, value.number, raw.getDate1904());
    }

    evaluateFormula(row, col, formula) {
      return toNative(typedResult(this.#evaluate(row, col, formula)));
    }

    evaluateTypedFormula(row, col, formula) {
      return typedResult(this.#evaluate(row, col, formula));
    }

    evaluateFormulaArray(row, col, formula) {
      return this.#evaluate(row, col, formula).cells.map((line) => line.map(nativeOf));
    }

    evaluateTypedFormulaArray(row, col, formula) {
      return this.#evaluate(row, col, formula).cells;
    }

    #evaluate(row, col, formula) {
      return rawOf(this.#workbook).evaluateFormulaArray(this.#current(), row, col, formula);
    }

    getTabColor() {
      return rawOf(this.#workbook).getTabColor(this.#current());
    }

    setTabColor(color) {
      const value = color === null || color === undefined ? color : formatColor(color, 'Sheet.setTabColor argument 1');
      rawOf(this.#workbook).setTabColor(this.#current(), value);
    }

    listColumnOverrides() {
      return rawOf(this.#workbook)
        .listColumnOverrides(this.#current())
        .map(({ styleXf, ...layout }) => layout);
    }

    listRowOverrides() {
      return rawOf(this.#workbook)
        .listRowOverrides(this.#current())
        .map(({ styleXf, ...layout }) => layout);
    }

    // Each row is written in order; a failure leaves the cells before it written and names its cell.
    setValues(row, col, values) {
      const where = 'Sheet.setValues';
      requireInteger(row, `${where} argument 1`);
      requireInteger(col, `${where} argument 2`);
      requireArray(values, `${where} argument 3`);
      const width = values.length === 0 ? 0 : requireArray(values[0], `${where} row 1`).length;
      for (const [i, line] of values.entries()) {
        requireArray(line, `${where} row ${i + 1}`);
        if (line.length !== width) {
          throw new FormulonError(StatusCode.InvalidArgument, `${where} rows differ in length`, `row=${i}`);
        }
        for (let j = 0; j < width; j++) {
          if (!(j in line) || line[j] === undefined) throw new TypeError(`${where} row ${i + 1} has a hole at ${j}`);
          checkCellValue(line[j], `${where} row ${i + 1} element ${j + 1}`);
        }
      }
      if (width === 0) return;
      if (row < 0 || col < 0 || row + values.length > GRID_ROWS || col + width > GRID_COLS) {
        throw new FormulonError(
          StatusCode.InvalidArgument,
          `${where} writes past the grid`,
          `row=${row} col=${col} rows=${values.length} cols=${width}`,
        );
      }
      for (const [i, line] of values.entries()) {
        for (let j = 0; j < width; j++) {
          try {
            this.setValue(row + i, col + j, line[j]);
          } catch (e) {
            throw withCell(e, row + i, col + j);
          }
        }
      }
    }

    getCellFormat(row, col) {
      const raw = rawOf(this.#workbook);
      const { xfIndex } = raw.getEffectiveStyle(this.#current(), row, col);
      return cellFormatOf(raw, raw.getCellXf(xfIndex));
    }

    // Every argument is checked first; the sheet changes only in the final remap.
    setCellFormat(range, patch) {
      const checked = checkCellFormatPatch(patch);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const bases = raw.getRangeStyleBases(sheet, range);
      if (Object.keys(checked).length === 0) return;
      let numFmtId;
      if (checked.numberFormat === 'General') numFmtId = 0;
      else if (checked.numberFormat !== undefined) numFmtId = raw.addNumFmt(checked.numberFormat);
      const xfs = bases.map((base) => raw.addCellXf(patchedXf(raw, raw.getCellXf(base), checked, numFmtId)));
      raw.remapRangeStyle(sheet, range, bases, xfs);
    }

    // The handle of the pivot at `index` on the sheet `sheet` is.
    #pivotAt(raw, sheet, index) {
      const id = raw.getPivotRuntimeId(sheet, index);
      let pivot = this.#pivots.get(id);
      if (pivot === undefined) {
        pivot = pivotState.create(this, id);
        this.#pivots.set(id, pivot);
      }
      return pivot;
    }

    listPivots() {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const n = raw.getPivotCount(sheet);
      const pivots = [];
      for (let i = 0; i < n; i++) pivots.push(this.#pivotAt(raw, sheet, i));
      return pivots;
    }

    pivot(index) {
      requireIndex(index, 'Sheet.pivot argument 1');
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      if (index >= raw.getPivotCount(sheet)) {
        throw new FormulonError(
          StatusCode.IndexOutOfRange,
          `pivot index ${index} is out of range`,
          `pivot_index=${index}`,
        );
      }
      return this.#pivotAt(raw, sheet, index);
    }

    addPivot(input) {
      const where = 'Sheet.addPivot';
      requireObject(input, `${where} argument 1`);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const index = raw.addPivot(
        sheet,
        requireString(input.name, `${where} name`),
        requireNumber(input.cacheId, `${where} cacheId`),
        requireNumber(input.anchorRow, `${where} anchorRow`),
        requireNumber(input.anchorCol, `${where} anchorCol`),
      );
      return this.#pivotAt(raw, sheet, index);
    }

    // The handle of drawing object `objectId`, whose `DrawingProperties` is `record`; `cls` overrides the class
    // its kind selects.
    #drawingFor(raw, sheet, objectId, record, cls) {
      let drawing = this.#drawings.get(objectId);
      if (drawing === undefined) {
        const of = cls ?? drawingClass(raw, sheet, record);
        drawing = drawingState.create(of, this, objectId, record.kind);
        this.#drawings.set(objectId, drawing);
      }
      return drawing;
    }

    // The object id a `number | Drawing` argument names.
    #drawingId(ref, where) {
      if (ref instanceof Drawing) {
        if (ref.sheet !== this) {
          throw new FormulonError(StatusCode.InvalidArgument, `${where} is a Drawing of another sheet`, '');
        }
        return ref.objectId;
      }
      if (typeof ref !== 'number') throw new TypeError(`${where} must be an object id or a Drawing`);
      return requireIndex(ref, where);
    }

    // The `DrawingProperties` of `objectId`; `NotFound` when the sheet's drawing has no such object.
    #drawingRecord(raw, sheet, objectId) {
      const record = raw.getDrawingObject(sheet, objectId);
      if (record === null) {
        throw new FormulonError(
          StatusCode.NotFound,
          `the sheet's drawing has no object ${objectId}`,
          `object_id=${objectId}`,
        );
      }
      return record;
    }

    #lookupDrawing(ref, method) {
      const id = this.#drawingId(ref, `Sheet.${method} argument 1`);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      return this.#drawingFor(raw, sheet, id, this.#drawingRecord(raw, sheet, id));
    }

    #typedDrawing(ref, cls, method) {
      const drawing = this.#lookupDrawing(ref, method);
      if (!(drawing instanceof cls)) {
        throw new FormulonError(
          StatusCode.InvalidArgument,
          `drawing object ${drawing.objectId} is not a ${cls.name}`,
          `object_id=${drawing.objectId}`,
        );
      }
      return drawing;
    }

    // The handle of an object just added with id `objectId`.
    #added(raw, sheet, objectId) {
      return this.#drawingFor(raw, sheet, objectId, this.#drawingRecord(raw, sheet, objectId));
    }

    drawing(objectId) {
      return this.#lookupDrawing(objectId, 'drawing');
    }

    picture(objectId) {
      return this.#typedDrawing(objectId, Picture, 'picture');
    }

    chart(objectId) {
      return this.#typedDrawing(objectId, Chart, 'chart');
    }

    shape(objectId) {
      return this.#typedDrawing(objectId, Shape, 'shape');
    }

    slicer(objectId) {
      return this.#typedDrawing(objectId, Slicer, 'slicer');
    }

    listDrawings() {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      return raw.listDrawingObjects(sheet).map((record) => this.#drawingFor(raw, sheet, record.objectId, record));
    }

    listPictures() {
      return this.listDrawings().filter((drawing) => drawing instanceof Picture);
    }

    listCharts() {
      return this.listDrawings().filter((drawing) => drawing instanceof Chart);
    }

    listShapes() {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      return raw.listShapes(sheet).map((record) => this.#drawingFor(raw, sheet, record.objectId, record, Shape));
    }

    listSlicers() {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      return raw.listSlicers(sheet).map(({ objectId }) => {
        const known = this.#drawings.get(objectId);
        return known ?? this.#drawingFor(raw, sheet, objectId, this.#drawingRecord(raw, sheet, objectId), Slicer);
      });
    }

    addPicture(bytes, input) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const id = input === undefined ? raw.addImage(sheet, bytes) : raw.addImage(sheet, bytes, input);
      return this.#added(raw, sheet, id);
    }

    addChart(options) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      return this.#added(raw, sheet, raw.addChart(sheet, colorInputs(options, 'Sheet.addChart argument 1')));
    }

    addShape(spec) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      return this.#added(raw, sheet, raw.addShape(sheet, colorInputs(spec, 'Sheet.addShape argument 1')));
    }

    addSlicer(table, column, spec) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      return this.#added(raw, sheet, raw.addSlicer(sheet, table, column, spec));
    }

    restoreDrawing(bytes, opts) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const id = opts === undefined ? raw.restoreDrawing(sheet, bytes) : raw.restoreDrawing(sheet, bytes, opts);
      return this.#added(raw, sheet, id);
    }

    // Each kind has its own raw remover; any other object cannot be removed.
    removeDrawing(objectId) {
      const drawing = this.#lookupDrawing(objectId, 'removeDrawing');
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const id = drawing.objectId;
      if (drawing instanceof Picture || drawing instanceof Chart) return raw.removeDrawingObject(sheet, id);
      if (drawing instanceof Shape) return raw.removeShape(sheet, id);
      if (drawing instanceof Slicer) return raw.removeSlicer(sheet, id);
      throw new FormulonError(
        StatusCode.InvalidArgument,
        `drawing object ${id} is not a picture, chart, shape or table slicer`,
        `object_id=${id}`,
      );
    }

    // Freeze or split, pane origin, scroll origin, display flags, zoom, view
    // mode, direction, tab selection, then visibility. Only the two refusals
    // checked up front leave the view untouched.
    setView(patch) {
      const where = 'Sheet.setView';
      const fields = PATCH_FIELDS[where];
      checkPatch(where, patch, fields);
      const given = (key) => patch[key] !== undefined;
      const nonZero = (...keys) => keys.some((key) => given(key) && patch[key] !== 0);
      if (nonZero('freezeRows', 'freezeCols') && nonZero('splitX', 'splitY')) {
        throw new TypeError(`${where} cannot freeze and split the view at once`);
      }
      if (!Object.keys(fields).some(given)) return;
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const view = raw.getView(sheet);
      const merged = { ...view };
      for (const key of Object.keys(fields)) if (given(key)) merged[key] = patch[key];
      const freeze = merged.freezeRows !== view.freezeRows || merged.freezeCols !== view.freezeCols;
      const split = merged.splitX !== view.splitX || merged.splitY !== view.splitY;
      // The panes once applied: a freeze replaces a split and a split a freeze.
      let panes = [view.freezeRows, view.freezeCols, view.splitX, view.splitY];
      if (freeze) panes = [merged.freezeRows, merged.freezeCols, 0, 0];
      if (split) {
        const cleared = merged.splitX === 0 && merged.splitY === 0;
        panes = cleared ? [panes[0], panes[1], 0, 0] : [0, 0, merged.splitX, merged.splitY];
      }
      // The origin a view without panes reads (0, 0) passes back unrefused.
      if (nonZero('paneTopLeftRow', 'paneTopLeftCol') && panes.every((n) => n === 0)) {
        throw new FormulonError(
          StatusCode.InvalidArgument,
          `${where} sets a pane origin on a view without panes`,
          `sheet_index=${sheet}`,
        );
      }
      if (merged.visibility !== SheetVisibility.Visible && view.visibility === SheetVisibility.Visible) {
        const count = raw.getSheetCount();
        let visible = false;
        for (let i = 0; i < count && !visible; i++) {
          visible = i !== sheet && raw.getView(i).visibility === SheetVisibility.Visible;
        }
        if (!visible) {
          throw new FormulonError(
            StatusCode.InvalidArgument,
            `${where} hides the only visible sheet`,
            `sheet_index=${sheet}`,
          );
        }
      }
      if (freeze) raw.setFreeze(sheet, merged.freezeRows, merged.freezeCols);
      if (split) raw.setSplit(sheet, merged.splitX, merged.splitY);
      // Changing the panes resets their origins, so partners are read afresh.
      const current = freeze || split ? raw.getView(sheet) : view;
      if (given('paneTopLeftRow') || given('paneTopLeftCol')) {
        raw.setPaneTopLeft(
          sheet,
          patch.paneTopLeftRow ?? current.paneTopLeftRow,
          patch.paneTopLeftCol ?? current.paneTopLeftCol,
        );
      }
      if (given('topLeftRow') || given('topLeftCol')) {
        raw.setTopLeft(sheet, patch.topLeftRow ?? current.topLeftRow, patch.topLeftCol ?? current.topLeftCol);
      }
      const singles = [
        ['showGridLines', 'setShowGridLines'],
        ['showRowColHeaders', 'setShowRowColHeaders'],
        ['showZeros', 'setShowZeros'],
        ['showFormulas', 'setShowFormulas'],
        ['showOutlineSymbols', 'setShowOutlineSymbols'],
        ['zoomScale', 'setZoom'],
        ['viewMode', 'setViewMode'],
        ['rightToLeft', 'setRightToLeft'],
        ['tabSelected', 'setTabSelected'],
        ['visibility', 'setVisibility'],
      ];
      for (const [key, setter] of singles) if (given(key)) raw[setter](sheet, patch[key]);
    }

    setPrintOptions(patch) {
      const where = 'Sheet.setPrintOptions';
      checkPatch(where, patch, PATCH_FIELDS[where]);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      raw.setPrintOptions(
        sheet,
        withoutUnstated(patch, () => raw.getPrintOptions(sheet)),
      );
    }

    // The C layer has no third state for a section: `null` and `''` both clear it.
    setHeaderFooter(patch) {
      const where = 'Sheet.setHeaderFooter';
      const sections = ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'];
      const nullable = Object.fromEntries(sections.map((key) => [key, 'string']));
      checkPatch(where, patch, PATCH_FIELDS[where], nullable);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const input = { ...withoutUnstated(patch, () => raw.getHeaderFooter(sheet)) };
      for (const key of sections) if (input[key] === null) input[key] = '';
      raw.setHeaderFooter(sheet, input);
    }

    getPrintArea() {
      const areas = rawOf(this.#workbook).getPrintArea(this.#current());
      return areas.length === 0 ? null : areas;
    }

    getColumnWidthPt(col, mode) {
      mode = geometryMode(mode, 'Sheet.getColumnWidthPt', 2);
      return rawOf(this.#workbook).getColumnWidthPt(this.#current(), col, mode);
    }

    convertColumnCharsToPt(chars, mode) {
      mode = geometryMode(mode, 'Sheet.convertColumnCharsToPt', 2);
      return rawOf(this.#workbook).convertColumnCharsToPt(this.#current(), mode, chars);
    }

    convertColumnPtToChars(pt, mode) {
      mode = geometryMode(mode, 'Sheet.convertColumnPtToChars', 2);
      return rawOf(this.#workbook).convertColumnPtToChars(this.#current(), mode, pt);
    }

    getCellRectPt(range, mode) {
      mode = geometryMode(mode, 'Sheet.getCellRectPt', 2);
      return rawOf(this.#workbook).getCellRectPt(this.#current(), range, mode);
    }

    getWidthModel(mode) {
      mode = geometryMode(mode, 'Sheet.getWidthModel', 1);
      return rawOf(this.#workbook).getWidthModel(this.#current(), mode);
    }

    // The raw update writes every field, so the patch lands on the group read back; without `sparklines` the
    // raw call keeps the group's own.
    updateSparklineGroup(groupIndex, patch) {
      const where = 'Sheet.updateSparklineGroup';
      patch = colorInputs(patch, `${where} argument 2`);
      checkSparklinePatch(where, patch);
      this.#patchOver(
        where,
        patch,
        'sparklines',
        (raw, sheet) => raw.getSparklineGroup(sheet, groupIndex),
        (raw, sheet, group) => raw.updateSparklineGroup(sheet, groupIndex, group),
      );
    }

    // As `updateSparklineGroup`; without `ranges` the raw call keeps the stored ranges, unparsed ones included.
    updateProtectedRange(protectedRange, patch) {
      const where = 'Sheet.updateProtectedRange';
      this.#patchOver(
        where,
        patch,
        'ranges',
        (raw, sheet) => raw.getProtectedRange(sheet, protectedRange),
        (raw, sheet, range) => raw.updateProtectedRange(sheet, protectedRange, range),
      );
    }

    // Lays the checked patch of `where` over the record `read` returns and hands it to `write`; the record's
    // `list` key reaches `write` only when the patch gives it. An empty patch writes nothing.
    #patchOver(where, patch, list, read, write) {
      const fields = PATCH_FIELDS[where];
      const nullable = PATCH_NULLABLE[where];
      checkPatch(where, patch, fields, nullable);
      const given = appliedFields(patch, fields, nullable);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const record = { ...read(raw, sheet), ...given };
      if (Object.keys(given).length === 0) return;
      if (given[list] === undefined) delete record[list];
      write(raw, sheet, record);
    }

    addDataTable(input) {
      const where = 'Sheet.addDataTable';
      requireObject(input, `${where} argument 1`);
      const raw = rawOf(this.#workbook);
      const spec = { rect: typeof input.rect === 'string' ? rangeAddress(raw, input.rect) : input.rect };
      for (const key of ['rowInput', 'colInput']) {
        const cell = input[key];
        if (typeof cell === 'string') {
          const { firstRow, firstCol } = cellAddress(raw, cell);
          spec[key] = { row: firstRow, col: firstCol };
        } else if (present(cell)) {
          spec[key] = cell;
        }
      }
      raw.addDataTable(this.#current(), spec);
    }

    setValue(row, col, value) {
      if (value === null) return this.setBlank(row, col);
      switch (typeof value) {
        case 'number':
          return this.setNumber(row, col, value);
        case 'string':
          return value.startsWith('=') ? this.setFormula(row, col, value) : this.setText(row, col, value);
        case 'boolean':
          return this.setBool(row, col, value);
        case 'object':
          if (value instanceof ExcelError) return this.setError(row, col, value.code);
          if (value instanceof Date) return this.#setDate(row, col, value);
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
      throw new TypeError(
        'Sheet.setValue argument 3 must be a number, string, boolean, null, Date, ExcelError or a scalar Value',
      );
    }

    // The serial of the date's UTC fields; a General cell also gets the built-in date format.
    #setDate(row, col, date) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const [serial, timed] = dateSerial(runtime, date, raw.getDate1904(), 'Sheet.setValue argument 3');
      this.setNumber(row, col, serial);
      const { xfIndex } = raw.getEffectiveStyle(sheet, row, col);
      if (raw.getCellXf(xfIndex).numFmtId !== 0) return;
      const rect = { firstRow: row, firstCol: col, lastRow: row, lastCol: col };
      const patch = { numberFormat: timed ? DATE_TIME_FORMAT : DATE_FORMAT };
      const bases = raw.getRangeStyleBases(sheet, rect);
      const xfs = bases.map((base) =>
        raw.addCellXf(patchedXf(raw, raw.getCellXf(base), patch, BUILTIN_NUM_FMT_IDS[patch.numberFormat])),
      );
      raw.remapRangeStyle(sheet, rect, bases, xfs);
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

    // Only the keys the patch gives reach the engine; a record's other fields are ignored.
    updateThreadedComment(id, patch) {
      const where = 'Sheet.updateThreadedComment';
      requireObject(patch, `${where} argument 2`);
      refuseUnknownKeys(where, patch);
      const input = {};
      const checks = [
        ['text', requireString],
        ['mentions', requireArray],
        ['done', requireBoolean],
      ];
      for (const [key, check] of checks) {
        if (patch[key] !== undefined) input[key] = check(patch[key], `${where} ${key}`);
      }
      rawOf(this.#workbook).updateThreadedComment(this.#current(), id, input);
    }

    setArrayFormula(arrayFormula) {
      const where = 'Sheet.setArrayFormula';
      requireObject(arrayFormula, `${where} argument 1`);
      const rect = requireObject(arrayFormula.rect, `${where} rect`);
      const formula = requireString(arrayFormula.formula, `${where} formula`);
      rawOf(this.#workbook).setArrayFormula(this.#current(), rect, formula);
    }

    setPageSetup(setup) {
      const where = 'Sheet.setPageSetup';
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      checkForwardedPatch(where, setup, PATCH_FIELDS[where]);
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
      raw.setPageSetup(
        sheet,
        withoutUnstated(setup, () => raw.getPageSetup(sheet)),
      );
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
    const optional = (field) =>
      link[field] === undefined || link[field] === null ? '' : requireString(link[field], `${where} ${field}`);
    const target =
      link.target === undefined || link.target === null ? null : requireString(link.target, `${where} target`);
    const location = optional('location');
    if (target === null && (link.location === undefined || link.location === null)) {
      throw new TypeError(`${where} location is required when target is null`);
    }
    return [target ?? '', optional('display'), optional('tooltip'), location];
  }

  class Workbook {
    #handle;
    #sheets;
    #tables = new Map();
    #pivotCaches = new Map();
    #disposed = false;
    #deleted = false;
    #ooxml;
    // The registry's unregister token: an object of its own, so whoever sees the token cannot reach the workbook.
    #registration = {};

    constructor(key, raw) {
      if (key !== token) {
        throw new TypeError(
          'Workbook has no public constructor; use Workbook.load, createDefault, createEmpty or createForProfile',
        );
      }
      this.#handle = raw;
      this.#sheets = [];
      registry?.register(this, raw, this.#registration);
      const count = raw.getSheetCount();
      for (let i = 0; i < count; i++) this.#sheets.push(sheetState.create(this, i, raw.getSheetName(i)));
    }

    static {
      rawOf = (workbook) => workbook.#raw;
      liveRaw = (workbook) => {
        workbook.#alive();
        return workbook.#raw;
      };
      isDisposed = (workbook) => workbook.#disposed;
      insertSheet = (workbook, position) => workbook.#insert(position);
      moveSheet = (workbook, from, toIndex) => workbook.#move(from, toIndex);
      pivotCacheOf = (workbook, cacheId) => workbook.#cacheHandle(cacheId);
    }

    static load(bytes) {
      if (bytes instanceof ArrayBuffer) bytes = new Uint8Array(bytes);
      else if (ArrayBuffer.isView(bytes) && !(bytes instanceof Uint8Array)) {
        bytes = new Uint8Array(bytes.buffer, bytes.byteOffset, bytes.byteLength);
      }
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

    get ooxml() {
      this.#alive();
      this.#ooxml ??= ooxmlState.workbook(this);
      return this.#ooxml;
    }

    *[Symbol.iterator]() {
      yield* this.listSheets();
    }

    // With a registry (the WASM entries), the raw object is deleted here as a collection would delete it.
    dispose() {
      if (this.#disposed) return;
      if (activeCalls.get(this) > 0) {
        throw new FormulonError(
          StatusCode.GraphRecalcReentrant,
          'cannot dispose Workbook during an active method call (kGraphRecalcReentrant, 4005)',
          '',
        );
      }
      this.#handle.dispose();
      this.#disposed = true;
      if (registry !== undefined) {
        registry.unregister(this.#registration);
        this.#handle.delete();
        this.#deleted = true;
      }
    }

    // The raw workbook. Once `dispose()` has deleted it, every access throws `BindingInvalidHandle`; a raw
    // object that survives its disposal answers for itself.
    get #raw() {
      if (this.#deleted) this.#alive();
      return this.#handle;
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

    #move(from, toIndex) {
      this.#raw.moveSheet(from, toIndex);
      const [moved] = this.#sheets.splice(from, 1);
      this.#sheets.splice(toIndex, 0, moved);
      this.#renumber();
    }

    // Resolves a `SheetRef` the way `sheet(key)` does; `method` and `position` word the errors.
    // `position` is an argument number, or the name of the record field that holds `ref`.
    #resolve(ref, method, position) {
      const where = `Workbook.${method} ${typeof position === 'number' ? `argument ${position}` : position}`;
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
      const before = insertBefore === null ? this.#sheets.length : insertBefore;
      const position =
        newName === undefined ? this.#raw.copySheet(index, before) : this.#raw.copySheet(index, before, newName);
      return this.#insert(position);
    }

    removeSheet(sheet) {
      const index = this.#resolve(sheet, 'removeSheet', 1);
      this.#raw.removeSheet(index);
      const [removed] = this.#sheets.splice(index, 1);
      sheetState.place(removed, -1);
      this.#renumber();
    }

    // The handle of the table at workbook-wide `index`.
    #tableAt(index) {
      const id = this.#raw.getTableRuntimeId(index);
      let table = this.#tables.get(id);
      if (table === undefined) {
        table = tableState.create(this, id);
        this.#tables.set(id, table);
      }
      return table;
    }

    listTables() {
      this.#alive();
      const n = this.#raw.getTableCount();
      const tables = [];
      for (let i = 0; i < n; i++) tables.push(this.#tableAt(i));
      return tables;
    }

    table(index) {
      requireIndex(index, 'Workbook.table argument 1');
      this.#alive();
      if (index >= this.#raw.getTableCount()) {
        throw new FormulonError(
          StatusCode.IndexOutOfRange,
          `table index ${index} is out of range`,
          `table_index=${index}`,
        );
      }
      return this.#tableAt(index);
    }

    // Omitted `columns` reach the raw method empty, which takes the names from the header cells.
    addTable(input) {
      requireObject(input, 'Workbook.addTable argument 1');
      for (const [field, context] of [
        ['name', 'name'],
        ['displayName', 'display_name'],
      ]) {
        const text = input[field];
        if (typeof text === 'string' && namesACell(this.#raw, text)) {
          throw new FormulonError(
            StatusCode.InvalidArgument,
            `invalid table ${field}: ${JSON.stringify(text)} is a cell address`,
            `${context}=${text}`,
          );
        }
      }
      const spec = input.columns === undefined ? { ...input, columns: [] } : input;
      return this.#tableAt(this.#raw.addTable(spec));
    }

    // The handle of the pivot cache `cacheId`, whether or not the workbook has it now.
    #cacheHandle(cacheId) {
      let cache = this.#pivotCaches.get(cacheId);
      if (cache === undefined) {
        cache = cacheState.create(this, cacheId);
        this.#pivotCaches.set(cacheId, cache);
      }
      return cache;
    }

    // The cache id a `number | PivotCache` argument names.
    #cacheId(ref, where) {
      if (ref instanceof PivotCache) {
        if (ref.workbook !== this) {
          throw new FormulonError(StatusCode.InvalidArgument, `${where} is a PivotCache of another workbook`, '');
        }
        return ref.cacheId;
      }
      if (typeof ref !== 'number') throw new TypeError(`${where} must be a cache id or a PivotCache`);
      return requireIndex(ref, where);
    }

    listPivotCaches() {
      this.#alive();
      return this.#raw.listPivotCaches().map((cacheId) => this.#cacheHandle(cacheId));
    }

    pivotCache(cacheId) {
      const id = this.#cacheId(cacheId, 'Workbook.pivotCache argument 1');
      this.#alive();
      if (!this.#raw.listPivotCaches().includes(id)) {
        throw new FormulonError(StatusCode.NotFound, `the workbook has no pivot cache ${id}`, `cache_id=${id}`);
      }
      return this.#cacheHandle(id);
    }

    addPivotCache(requestedId) {
      const requested = requestedId === undefined || requestedId === null ? 0 : requestedId;
      return this.#cacheHandle(this.#raw.addPivotCache(requested));
    }

    removePivotCache(cacheId) {
      this.#raw.removePivotCache(this.#cacheId(cacheId, 'Workbook.removePivotCache argument 1'));
    }

    setPinnedNow(time) {
      if (time === null) return this.#raw.clearPinnedNow();
      const where = 'Workbook.setPinnedNow';
      if (time instanceof Date) {
        if (Number.isNaN(time.getTime())) throw new RangeError(`${where} argument 1 is an invalid Date`);
        return this.#raw.setPinnedNow(
          time.getUTCFullYear(),
          time.getUTCMonth() + 1,
          time.getUTCDate(),
          time.getUTCHours(),
          time.getUTCMinutes(),
          time.getUTCSeconds(),
        );
      }
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

    getThemeScriptFont(key) {
      const where = 'Workbook.getThemeScriptFont';
      requireObject(key, `${where} argument 1`);
      return this.#raw.getThemeScriptFont(
        requireBoolean(key.major, `${where} major`),
        requireString(key.script, `${where} script`),
      );
    }

    // A `null` or omitted scope is workbook scope; the raw sentinel is -1.
    #scope(scope, method) {
      return scope === undefined || scope === null ? -1 : this.#resolve(scope, method, 2);
    }

    getDefinedName(name, scope) {
      return this.#raw.getDefinedName(name, this.#scope(scope, 'getDefinedName'));
    }

    removeDefinedName(name, scope) {
      return this.#raw.removeDefinedName(name, this.#scope(scope, 'removeDefinedName'));
    }

    setDefinedName(definedName) {
      requireObject(definedName, 'Workbook.setDefinedName argument 1');
      const { scope, ...fields } = definedName;
      if (scope === undefined || scope === null) return this.#raw.setDefinedName(fields);
      return this.#raw.setDefinedName({ ...fields, localSheetId: this.#resolve(scope, 'setDefinedName', 'scope') });
    }

    setIterative(patch) {
      const where = 'Workbook.setIterative';
      checkForwardedPatch(where, patch, PATCH_FIELDS[where]);
      if (isObject(patch)) {
        const { maxIterations, maxChange } = patch;
        if (maxIterations !== undefined && !Number.isInteger(maxIterations)) {
          throw new TypeError(`${where} maxIterations must be an integer`);
        }
        if (typeof maxChange === 'number' && maxChange < 0)
          throw new RangeError(`${where} maxChange must not be negative`);
      }
      this.#raw.setIterative(patch);
    }
  }

  class WorkbookOoxml {
    #workbook;

    constructor(key, workbook) {
      if (key !== token) throw new TypeError('WorkbookOoxml has no public constructor; use Workbook.ooxml');
      this.#workbook = workbook;
    }

    static {
      ooxmlState = { ...ooxmlState, workbookOf: (ooxml) => ooxml.#workbook };
    }
  }

  // The row and column XF indices are read from the layout overrides, whose facade records leave them out.
  class SheetOoxml {
    #sheet;

    constructor(key, sheet) {
      if (key !== token) throw new TypeError('SheetOoxml has no public constructor; use Sheet.ooxml');
      this.#sheet = sheet;
    }

    static {
      ooxmlState = { ...ooxmlState, sheetOf: (ooxml) => ooxml.#sheet };
    }

    getColumnStyle(col) {
      requireIndex(col, 'Sheet.getColumnStyle argument 1');
      const raw = rawOf(this.#sheet.workbook);
      const span = raw.listColumnOverrides(sheetIndex(this.#sheet)).find((c) => c.first <= col && col <= c.last);
      return span?.styleXf ?? null;
    }

    getRowStyle(row) {
      requireIndex(row, 'Sheet.getRowStyle argument 1');
      const raw = rawOf(this.#sheet.workbook);
      const override = raw.listRowOverrides(sheetIndex(this.#sheet)).find((r) => r.row === row);
      return override?.styleXf ?? null;
    }
  }

  ooxmlState = {
    ...ooxmlState,
    workbook: (workbook) => new WorkbookOoxml(token, workbook),
    sheet: (sheet) => new SheetOoxml(token, sheet),
  };

  const sheetCall = (self, name, args) => rawOf(self.workbook)[name](sheetIndex(self), ...args);
  const tableCall = (self, name, args) => {
    const [raw, index] = tableState.locate(self);
    return raw[name](index, ...args);
  };
  const pivotCall = (self, name, args) => {
    const [raw, sheet, index] = pivotState.locate(self);
    return raw[name](sheet, index, ...args);
  };
  const cacheCall = (self, name, args) => cacheState.raw(self)[name](self.cacheId, ...args);
  // A drawing read the raw method answers with `null` for an object without that part (a picture's media).
  const drawingRead = /^(?:Drawing|Picture|Chart|Shape|Slicer)\.get(?:Drawing)?Properties$/;
  const drawingReads = new Set(
    Object.entries(HANDLE_FORWARDERS)
      .filter(([key, raw]) => raw !== null && drawingRead.test(key))
      .map(([, raw]) => raw),
  );
  const drawingCall = (self, name, args) => {
    const [raw, sheet, id] = drawingState.locate(self);
    const out = raw[name](sheet, id, ...args);
    if (out === null && drawingReads.has(name)) {
      throw new FormulonError(StatusCode.NotFound, 'the drawing object has no such part', `object_id=${id}`);
    }
    return out;
  };

  const workbookCall = (self, name, args) => rawOf(self)[name](...args);
  const facadeClasses = [
    ['Workbook', Workbook, workbookCall, WORKBOOK_METHODS],
    ['Sheet', Sheet, sheetCall, SHEET_METHODS],
    [
      'Workbook',
      WorkbookOoxml,
      (self, name, args) => workbookCall(ooxmlState.workbookOf(self), name, args),
      OOXML_WORKBOOK_METHODS,
    ],
    ['Sheet', SheetOoxml, (self, name, args) => sheetCall(ooxmlState.sheetOf(self), name, args), OOXML_SHEET_METHODS],
  ];
  for (const [owner, cls, call, names] of facadeClasses) {
    installForwarders(
      cls.prototype,
      owner,
      names.map((name) => [name, name]),
      Raw,
      call,
    );
  }
  const handleClasses = [
    ['Sheet', Sheet, sheetCall],
    ['Table', Table, tableCall],
    ['Pivot', Pivot, pivotCall],
    ['PivotCache', PivotCache, cacheCall],
    ['Drawing', Drawing, drawingCall],
    ['Picture', Picture, drawingCall],
    ['Chart', Chart, drawingCall],
    ['Shape', Shape, drawingCall],
    ['Slicer', Slicer, drawingCall],
  ];
  for (const [owner, cls, call] of handleClasses) {
    const prefix = `${owner}.`;
    const entries = Object.entries(HANDLE_FORWARDERS)
      .filter(([key]) => key.startsWith(prefix))
      .map(([key, raw]) => [key.slice(prefix.length), raw]);
    installForwarders(cls.prototype, owner, entries, Raw, call);
    for (const [key, getter] of Object.entries(HANDLE_PROPERTIES)) {
      if (!key.startsWith(prefix)) continue;
      Object.defineProperty(cls.prototype, key.slice(prefix.length), {
        get() {
          return call(this, getter, []);
        },
        configurable: true,
      });
    }
  }

  const tableSheet = (self) => {
    const [raw, index] = tableState.locate(self);
    return raw.getTable(index).range.sheet;
  };
  const argumentForms = [
    ['Workbook', Workbook, rawOf, () => undefined],
    ['Sheet', Sheet, (self) => rawOf(self.workbook), sheetIndex],
    ['Pivot', Pivot, (self) => rawOf(self.sheet.workbook), (self) => sheetIndex(self.sheet)],
    ['Table', Table, (self) => tableState.locate(self)[0], tableSheet],
    [
      'Sheet',
      SheetOoxml,
      (self) => rawOf(ooxmlState.sheetOf(self).workbook),
      (self) => sheetIndex(ooxmlState.sheetOf(self)),
    ],
  ];
  for (const [owner, cls, rawFor, sheetFor] of argumentForms) {
    installArgumentForms(cls.prototype, owner, rawFor, sheetFor, handleIndex, runtime);
  }
  // Under the arity guard: the handle is checked before a call reads or validates its input, and the workbook
  // cannot be disposed from inside the call.
  const workbookOfSheet = (sheet) => sheet.workbook;
  const guardedClasses = [
    [Workbook, (self) => self],
    [Sheet, workbookOfSheet],
    [Table, workbookOfSheet],
    [PivotCache, workbookOfSheet],
    [Pivot, (self) => self.sheet.workbook],
    [Drawing, (self) => self.sheet.workbook],
    [Picture, (self) => self.sheet.workbook],
    [Chart, (self) => self.sheet.workbook],
    [Shape, (self) => self.sheet.workbook],
    [Slicer, (self) => self.sheet.workbook],
    [WorkbookOoxml, (self) => ooxmlState.workbookOf(self)],
    [SheetOoxml, (self) => ooxmlState.sheetOf(self).workbook],
  ];
  for (const [cls, workbookOf] of guardedClasses) installCallGuards(cls.prototype, workbookOf, liveRaw, activeCalls);
  const guarded = [['Workbook', Workbook], ['Workbook', WorkbookOoxml], ['Sheet', SheetOoxml], ...handleClasses];
  for (const [owner, cls] of guarded) installArityGuards(cls.prototype, owner);
  // `using` calls the same function as `dispose()`.
  Object.defineProperty(Workbook.prototype, Symbol.dispose, {
    value: Workbook.prototype.dispose,
    writable: true,
    configurable: true,
  });
  return Workbook;
}

/**
 * Wraps every defined method of `owner` listed in `CELL_ARGS`, `RECT_ARGS`,
 * `SPAN_ARGS`, `REF_ARGS`, `TYPED_ARGS` or `DEFAULT_ARGS` so that A1 addresses, A1
 * ranges, A1 spans, `*Ref` strings, `*Ref` handles and the A1 text of record fields
 * (`normalizeRecordRanges`) reach it in coordinate, index and raw form, with an
 * omitted trailing argument given its ledger default. `rawOf(self)` and
 * `sheetOf(self)` give the raw workbook and the sheet index a string key is
 * resolved against; `handleIndex(self, kind, value, label)` the index a handle
 * names; `runtime` converts the `'today'` default.
 */
function installArgumentForms(proto, owner, rawOf, sheetOf, handleIndex, runtime) {
  const prefix = `${owner}.`;
  const tables = [CELL_ARGS, RECT_ARGS, SPAN_ARGS, REF_ARGS, TYPED_ARGS, DEFAULT_ARGS];
  const keys = new Set(tables.flatMap((table) => Object.keys(table)));
  for (const where of keys) {
    if (!where.startsWith(prefix)) continue;
    const name = where.slice(prefix.length);
    if (!Object.hasOwn(proto, name)) continue;
    const inner = proto[name];
    const pairs = CELL_ARGS[where];
    const rects = RECT_ARGS[where];
    const spans = SPAN_ARGS[where];
    const ref = REF_ARGS[where];
    const typed = TYPED_ARGS[where];
    const defaults = DEFAULT_ARGS[where];
    defineMethod(
      proto,
      name,
      namedFunction(name, function (args) {
        if (spans !== undefined) args = expandSpans(rawOf(this), spans, args);
        if (pairs !== undefined) args = expandCells(rawOf(this), where, pairs, args);
        if (defaults !== undefined) args = fillDefaults(runtime, rawOf(this), defaults, args);
        if (typed !== undefined) args = normalizeTypedArgs(rawOf(this), sheetOf(this), where, typed, args);
        if (rects !== undefined) args = expandRects(rawOf(this), rects, args);
        if (ref !== undefined && typeof args[ref[0]] !== 'number') {
          const named = (kind, value, label) => handleIndex(this, kind, value, label);
          args = resolveRef(rawOf(this), sheetOf(this), where, ref, args, named);
        }
        return inner.apply(this, args);
      }),
    );
  }
}

/**
 * Wraps every defined method of `owner` listed in `METHOD_ARITY` so that a call
 * with more positional arguments than the declarations allow throws
 * `TypeError: <owner>.<method> takes at most N arguments`. A call in the string
 * form (an A1 cell for each `CELL_ARGS` pair, an A1 span for a `SPAN_ARGS` one)
 * takes one argument fewer per string. `[Symbol.dispose]` is an alias defined
 * afterwards.
 */
function installArityGuards(proto, owner) {
  const prefix = `${owner}.`;
  for (const [where, most] of Object.entries(METHOD_ARITY)) {
    if (!where.startsWith(prefix)) continue;
    const name = where.slice(prefix.length);
    if (name === '[Symbol.dispose]') continue;
    const key = name === '[Symbol.iterator]' ? Symbol.iterator : name;
    if (!Object.hasOwn(proto, key)) continue;
    const inner = proto[key];
    const pairs = CELL_ARGS[where];
    const spans = SPAN_ARGS[where];
    defineMethod(
      proto,
      key,
      namedFunction(name, function (args) {
        let limit = most;
        if (pairs !== undefined && typeof args[pairs[0][0]] === 'string') limit -= pairs.length;
        for (const [at] of spans ?? []) if (typeof args[at] === 'string') limit -= 1;
        if (args.length > limit) {
          throw new TypeError(`${where} takes at most ${limit} argument${limit === 1 ? '' : 's'}`);
        }
        return inner.apply(this, args);
      }),
    );
  }
}

/**
 * Wraps every method `proto` defines itself (not `dispose`) so that the workbook's handle is checked before the
 * call reads or validates its arguments, and so that the workbook counts the call as active until it returns.
 */
function installCallGuards(proto, workbookOf, liveRaw, activeCalls) {
  for (const name of Object.getOwnPropertyNames(proto)) {
    if (name === 'constructor' || name === 'dispose') continue;
    const desc = Object.getOwnPropertyDescriptor(proto, name);
    if (typeof desc.value !== 'function') continue;
    const inner = desc.value;
    defineMethod(
      proto,
      name,
      namedFunction(name, function (args) {
        const workbook = workbookOf(this);
        liveRaw(workbook);
        activeCalls.set(workbook, (activeCalls.get(workbook) ?? 0) + 1);
        try {
          return inner.apply(this, args);
        } finally {
          activeCalls.set(workbook, activeCalls.get(workbook) - 1);
        }
      }),
    );
  }
}

/**
 * Defines the generated forwarders on a facade prototype: each `[name, rawName]`
 * of `entries` the prototype does not define by hand calls `call(this, rawName,
 * args)`. An op the ledger marks hand-written (a `Workbook` / `Sheet` name in
 * `HANDWRITTEN`, a handle method whose raw name is null), or one that changes the
 * sheet list, must already be defined by hand. A list family loops over the raw
 * count and element getter. An `ABSENT_ARGS` entry maps `null` (and, for an
 * optional argument, `undefined`) at its position to the raw sentinel, and
 * refuses a negative number where that sentinel is negative; a patch op
 * (`PATCH_KEYS`) refuses a key outside its table and, per `PATCH_FIELDS`, `null`
 * on a non-nullable field. `owner` words the errors. A method the raw class does
 * not register (a member only the other package has) is skipped.
 */
function installForwarders(proto, owner, entries, Raw, call) {
  const facade = owner === 'Workbook' || owner === 'Sheet';
  for (const [name, rawName] of entries) {
    if (Object.hasOwn(proto, name === '[Symbol.iterator]' ? Symbol.iterator : name)) continue;
    const where = `${owner}.${name}`;
    const family = LIST_FAMILIES[where];
    if (family !== undefined) {
      const [count, element] = family;
      defineMethod(
        proto,
        name,
        namedFunction(name, function (args) {
          args = colorArgs(args, where);
          const n = call(this, count, args);
          const out = [];
          for (let i = 0; i < n; i++) out.push(call(this, element, [...args, i]));
          return out;
        }),
      );
      continue;
    }
    if (rawName === null || (facade && HANDWRITTEN.includes(name)) || SHEET_STRUCTURE_OPS.includes(where)) {
      throw new TypeError(`formulon: the facade does not implement the hand-written ${where}`);
    }
    if (typeof Raw.prototype[rawName] !== 'function') continue;
    if (PATCH_KEYS[where] !== undefined) {
      const [position, keys] = PATCH_KEYS[where];
      const fields = PATCH_FIELDS[where];
      // A patch that carries `*Stated` keys passes back the record its `get*` twin reads.
      const stated = keys.some((key) => key.endsWith('Stated'));
      const reader = rawName.replace(/^set/, 'get');
      defineMethod(
        proto,
        name,
        namedFunction(name, function (args) {
          args = colorArgs(args, where);
          checkForwardedPatch(where, args[position], fields);
          if (stated) {
            args = [...args];
            args[position] = withoutUnstated(args[position], () => call(this, reader, args.slice(0, position)));
          }
          return call(this, rawName, args);
        }),
      );
      continue;
    }
    const absent = ABSENT_ARGS[where];
    defineMethod(
      proto,
      name,
      absent === undefined
        ? namedFunction(name, function (args) {
            return call(this, rawName, colorArgs(args, where));
          })
        : namedFunction(name, function (args) {
            args = colorArgs(args, where);
            const [position, sentinel, optional = false] = absent;
            const value = args[position];
            if (value === null || (optional && value === undefined)) {
              args = [...args];
              args[position] = sentinel;
            } else if (typeof sentinel === 'number' && sentinel < 0 && typeof value === 'number' && value < 0) {
              // The C sentinel is not a public value; absence is spelled `null`.
              throw new RangeError(`${where} argument ${position + 1} is outside uint32 range`);
            }
            return call(this, rawName, args);
          }),
    );
  }
}
