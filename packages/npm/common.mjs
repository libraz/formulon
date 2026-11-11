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
export const WorkbookFormat = Object.freeze({ Xlsx: 0, Xlsb: 1 });
export const CellShift = Object.freeze({ Down: 0, Right: 1, Up: 2, Left: 3 });
export const SpillCause = Object.freeze({ Occupied: 0, Merged: 1, BeyondEdge: 2 });
export const TableTotalsRowFunction = Object.freeze({
  None: 0,
  Sum: 1,
  Min: 2,
  Max: 3,
  Average: 4,
  Count: 5,
  CountNumbers: 6,
  StdDev: 7,
  Var: 8,
  Custom: 9,
});
export const TableStyleElementType = Object.freeze({
  WholeTable: 0,
  HeaderRow: 1,
  TotalRow: 2,
  FirstColumn: 3,
  LastColumn: 4,
  FirstRowStripe: 5,
  SecondRowStripe: 6,
  FirstColumnStripe: 7,
  SecondColumnStripe: 8,
  FirstHeaderCell: 9,
  LastHeaderCell: 10,
  FirstTotalCell: 11,
  LastTotalCell: 12,
  FirstSubtotalColumn: 13,
  SecondSubtotalColumn: 14,
  ThirdSubtotalColumn: 15,
  FirstSubtotalRow: 16,
  SecondSubtotalRow: 17,
  ThirdSubtotalRow: 18,
  BlankRow: 19,
  FirstColumnSubheading: 20,
  SecondColumnSubheading: 21,
  ThirdColumnSubheading: 22,
  FirstRowSubheading: 23,
  SecondRowSubheading: 24,
  ThirdRowSubheading: 25,
  PageFieldLabels: 26,
  PageFieldValues: 27,
});
export const PhoneticType = Object.freeze({
  HalfwidthKatakana: 0,
  FullwidthKatakana: 1,
  Hiragana: 2,
  NoConversion: 3,
});
export const PhoneticAlignment = Object.freeze({
  NoControl: 0,
  Left: 1,
  Center: 2,
  Distributed: 3,
});
export const PageOrder = Object.freeze({
  DownThenOver: 0,
  OverThenDown: 1,
});
export const PaperSize = Object.freeze({
  Letter: 1,
  Legal: 5,
  A3: 8,
  A4: 9,
  A5: 11,
  B4: 12,
  B5: 13,
  A6: 70,
});
export const WidthPlatform = Object.freeze({
  Win: 0,
  Mac: 1,
});
export const ChartTickLabelPosition = Object.freeze({
  NextTo: 0,
  High: 1,
  Low: 2,
  None: 3,
});
export const ChartDomain = Object.freeze({
  Category: 0,
  Xy: 1,
});
export const PivotCellKind = Object.freeze({
  Header: 0,
  RowLabel: 1,
  ColumnLabel: 2,
  Data: 3,
  RowSubtotal: 4,
  ColumnSubtotal: 5,
  GrandTotal: 6,
  Blank: 7,
});
export const PivotAxis = Object.freeze({ Row: 0, Column: 1, Value: 2, Page: 3 });
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
  PercentOfColumn: 2,
  PercentOfTotal: 3,
  RunningTotalInRow: 4,
  RunningTotalInColumn: 5,
  Index: 6,
  DifferenceFrom: 7,
  PercentDifferenceFrom: 8,
  PercentOfParentRow: 9,
  PercentOfParentColumn: 10,
  PercentOfParent: 11,
  PercentOf: 12,
  PercentOfRunningTotal: 13,
  RankAscending: 14,
  RankDescending: 15,
});
export const PivotShowAsBase = Object.freeze({
  Previous: 1048828,
  Next: 1048829,
});
export const PIVOT_VALUES_FIELD = -2;
export const PivotFilterKind = Object.freeze({
  ValueTopCount: 0,
  ValueGreaterThan: 1,
  ValueBetween: 2,
  LabelContains: 3,
  LabelBeginsWith: 4,
  DateBetween: 5,
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
  Tomorrow: 36,
  Today: 37,
  Yesterday: 38,
  NextWeek: 39,
  ThisWeek: 40,
  LastWeek: 41,
  NextMonth: 42,
  ThisMonth: 43,
  LastMonth: 44,
  NextQuarter: 45,
  ThisQuarter: 46,
  LastQuarter: 47,
  NextYear: 48,
  ThisYear: 49,
  LastYear: 50,
  YearToDate: 51,
  Q1: 52,
  Q2: 53,
  Q3: 54,
  Q4: 55,
  M1: 56,
  M2: 57,
  M3: 58,
  M4: 59,
  M5: 60,
  M6: 61,
  M7: 62,
  M8: 63,
  M9: 64,
  M10: 65,
  M11: 66,
  M12: 67,
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
export const PivotSortOrder = Object.freeze({
  Manual: 0,
  Ascending: 1,
  Descending: 2,
});
export const PivotReportLayout = Object.freeze({ Compact: 0, Tabular: 1, Outline: 2 });
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
export const DatePeriod = Object.freeze({
  Today: 0,
  Yesterday: 1,
  Tomorrow: 2,
  Last7Days: 3,
  ThisMonth: 4,
  LastMonth: 5,
  NextMonth: 6,
  ThisWeek: 7,
  LastWeek: 8,
  NextWeek: 9,
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
  ThreeStars: 17,
  ThreeTriangles: 18,
  FiveBoxes: 19,
  NoIcons: 20,
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
export const ComparisonOperator = Object.freeze({
  Between: 0,
  NotBetween: 1,
  Equal: 2,
  NotEqual: 3,
  LessThan: 4,
  LessThanOrEqual: 5,
  GreaterThan: 6,
  GreaterThanOrEqual: 7,
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
export const AnchorKind = Object.freeze({ Default: 0, OneCell: 1, TwoCell: 2, Absolute: 3 });
export const AnchorEditAs = Object.freeze({ Default: 0, OneCell: 1, TwoCell: 2, Absolute: 3 });
export const SlicerSortOrder = Object.freeze({
  Ascending: 0,
  Descending: 1,
});
export const SlicerCrossFilter = Object.freeze({ ShowItemsWithDataAtTop: 0, ShowItemsWithNoData: 1, None: 2 });
export const SparklineType = Object.freeze({ Line: 0, Column: 1, Stacked: 2 });
export const SparklineAxisType = Object.freeze({ Individual: 0, Group: 1, Custom: 2 });
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
  Extended: 16,
});
export const ChartExtendedLayout = Object.freeze({
  BoxWhisker: 0,
  ClusteredColumn: 1,
  Funnel: 2,
  ParetoLine: 3,
  RegionMap: 4,
  Sunburst: 5,
  Treemap: 6,
  Waterfall: 7,
});
export const ChartRole = Object.freeze({
  Name: 0,
  Categories: 1,
  Values: 2,
  X: 3,
  Y: 4,
  Size: 5,
  Title: 6,
  AxisTitle: 7,
});
export const ChartGrouping = Object.freeze({ Clustered: 0, Stacked: 1, PercentStacked: 2, Standard: 3 });
export const BlanksAs = Object.freeze({
  Zero: 0,
  Gap: 1,
  Span: 2,
});
export const ChartPointState = Object.freeze({ Value: 0, Gap: 1, Connect: 2 });
export const ChartAxisKind = Object.freeze({ Category: 0, Value: 1, Date: 2, Series: 3 });
export const ChartAxisCrosses = Object.freeze({ AutoZero: 0, Min: 1, Max: 2, At: 3 });
export const ChartTitleKind = Object.freeze({ Text: 1, Ref: 2, Auto: 3 });
export const ChartSeriesDirection = Object.freeze({ Auto: 0, Columns: 1, Rows: 2 });
export const ChartSourceKind = Object.freeze({ None: 0, Ref: 1, Literal: 2 });
export const ChartLegendPosition = Object.freeze({
  Right: 0,
  Bottom: 1,
  TopRight: 2,
  Left: 3,
  Top: 4,
});
export const ChartAxisPosition = Object.freeze({ Bottom: 0, Top: 1, Left: 2, Right: 3 });
export const ChartBarDirection = Object.freeze({ Column: 0, Bar: 1 });
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
  InvalidArgument: 2,
  NotImplemented: 3,
  OutOfMemory: 4,
  NotFound: 6,
  AlreadyExists: 7,
  InternalError: 8,
  PreconditionFailed: 9,
  SheetIndexOutOfRange: 10,
  InvalidSheetName: 11,
  SheetCountLimitExceeded: 12,
  IndexOutOfRange: 13,
  ParserUnexpectedToken: 1000,
  ParserUnterminatedString: 1001,
  ParserInvalidNumber: 1003,
  ParserInvalidReference: 1004,
  ParserInvalidSheetName: 1005,
  ParserMismatchedBrackets: 1006,
  ParserExcessiveNesting: 1007,
  ParserInvalidArrayLiteral: 1009,
  ParserInvalidStructRef: 1010,
  ParserBomNotSupported: 1013,
  ParserTooManyErrors: 1014,
  EvalPivotMissing: 2009,
  EvalPivotInvalid: 2010,
  FunctionOverflow: 3004,
  GraphScheduleFailed: 4003,
  GraphRecalcReentrant: 4005,
  IoFileNotFound: 5000,
  IoFileTooLarge: 5002,
  IoZipCorrupt: 5003,
  IoZipEncrypted: 5004,
  IoZipBomb: 5005,
  IoZipSlip: 5006,
  IoXmlParse: 5007,
  IoRelationshipBroken: 5010,
  IoContentTypeInvalid: 5011,
  IoSheetCorrupt: 5012,
  IoXlsbRecordCorrupt: 5014,
  IoWriteFailed: 5016,
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
  SecurityResourceLimit: 6005,
  BindingInvalidHandle: 7000,
  BindingNullPointer: 7001,
  BindingUtf8EncodingError: 7002,
  BindingCallbackException: 7003,
  DataTableInvalid: 8005,
  AutoFilterInvalid: 8006,
  PrintInvalidArea: 9002,
  PrintPageCountOverflow: 9005,
});
// <status-codes:end>

/**
 * Merge a host-supplied metadata entry over the engine's structural
 * `getFunctionMetadata()` result.
 *
 * This is a pure, side-effect-free helper: it does not touch the engine.
 * The engine returns `signature` / `description` as
 * `undefined`; a host injects display metadata (see
 * `docs/function-metadata-schema.md`) and merges it here at display time.
 * The metadata is display-only and never affects formula parsing or
 * evaluation.
 *
 * Field precedence (first non-nullish wins):
 *   - signature: `entry.localized[locale].signature` ->
 *     `entry.signature` -> `base.signature`
 *   - description: `entry.localized[locale].description` ->
 *     `entry.description` -> `base.description`
 *   - localizedName: `entry.aliases[locale]` -> `base.name`
 *
 * @param {object} base A `FunctionMetadata`, as `getFunctionMetadata()` returns it.
 * @param {object|undefined} entry The provider's `functions[NAME]` entry, or
 *   `undefined`/`null` for none (signature/description stay `undefined`, and
 *   `localizedName` is `base.name`).
 * @param {string} locale A BCP-47 display locale tag (e.g. `"fr-FR"`),
 *   matching the keys in `aliases` / `localized`. Independent of the profile id
 *   passed to the engine locale calls.
 * @returns {object} The merged metadata (a new object).
 */
export function mergeFunctionMetadata(base, entry, locale) {
  if (entry === undefined || entry === null) {
    return { ...base, localizedName: base.name };
  }
  const localized = (entry.localized && entry.localized[locale]) || {};
  const signature = localized.signature ?? entry.signature ?? base.signature;
  const description = localized.description ?? entry.description ?? base.description;
  const aliasName = entry.aliases && entry.aliases[locale];
  const localizedName = aliasName ?? base.name;
  return { ...base, signature, description, localizedName };
}

/**
 * An Excel cell error such as `#DIV/0!`, as the native reads return it. `code` is an `ErrorCode`; `toString()`
 * spells it as `ErrorDisplayName` does.
 */
export class ExcelError {
  /** @param {number} code An `ErrorCode`. */
  constructor(code) {
    if (typeof code !== 'number') throw new TypeError('ExcelError code must be an ErrorCode');
    if (ErrorDisplayName[code] === undefined) throw new RangeError('ExcelError code must be an ErrorCode');
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
  if (!isObject(enumType)) throw new TypeError('Formulon.enumName enumType must be an object');
  for (const [key, ordinal] of Object.entries(enumType)) if (ordinal === value) return key;
  return null;
}

/**
 * Maps a `Value` to its native JS value: Blank to `null`, Number to a number, Bool to a boolean, Text to a
 * string, Error to an `ExcelError` and an `ArrayValue` to its rows. A `RefValue` or `LambdaValue`, which no
 * read returns, and anything that is not a `Value` throw `TypeError`.
 *
 * @param {object} value A `Value`.
 * @returns {null | number | boolean | string | ExcelError | Array} The native value.
 */
export function toNative(value) {
  return nativeOfValue(requireValue(value, 'Formulon.toNative value'));
}

/** `toNative` of a value `requireValue` accepted. */
function nativeOfValue(value) {
  switch (value.kind) {
    case ValueKind.Blank:
      return null;
    case ValueKind.Array: {
      const { rows, cols, cells } = value;
      return Array.from({ length: rows }, (_, r) => cells.slice(r * cols, (r + 1) * cols).map(nativeOfValue));
    }
  }
  return value.kind === ValueKind.Error ? new ExcelError(value.errorCode) : value[SCALAR_VALUE_FIELDS[value.kind][0]];
}

/**
 * `value` when it is a `Value` of a kind a read returns, its payload of the kind's type (an `ArrayValue` with
 * `rows * cols` such cells); otherwise a `TypeError` "`<label> must be a Value`".
 */
function requireValue(value, label) {
  if (!isObject(value)) throw new TypeError(`${label} must be a Value`);
  if (value.kind === ValueKind.Blank) return value;
  if (value.kind === ValueKind.Array) {
    const { rows, cols, cells } = value;
    if (!Number.isInteger(rows) || !Number.isInteger(cols) || !Array.isArray(cells) || cells.length !== rows * cols) {
      throw new TypeError(`${label} must be a Value`);
    }
    for (const [i, cell] of cells.entries()) requireValue(cell, `${label}.cells[${i}]`);
    return value;
  }
  const [field, type] = SCALAR_VALUE_FIELDS[value.kind] ?? [];
  if (field === undefined || typeof value[field] !== type) throw new TypeError(`${label} must be a Value`);
  return value;
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
  'CellShift',
  'SpillCause',
  'PivotCellKind',
  'PivotAxis',
  'PivotAggregation',
  'PivotShowAsBase',
  'PivotShowValuesAs',
  'PivotFilterKind',
  'PivotDatePeriod',
  'PivotCalendar',
  'PivotSortOrder',
  'PivotReportLayout',
  'CfMatchKind',
  'CfRuleType',
  'DatePeriod',
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
  'ComparisonOperator',
  'DynamicFilterType',
  'SortBy',
  'SortMethod',
  'DateTimeGrouping',
  'ValidationErrorStyle',
  'ValidationType',
  'ColorContext',
  'ImageFormat',
  'DrawingObjectKind',
  'AnchorKind',
  'AnchorEditAs',
  'SlicerSortOrder',
  'SlicerCrossFilter',
  'SparklineType',
  'SparklineAxisType',
  'ChartPlotType',
  'ChartExtendedLayout',
  'ChartRole',
  'ChartGrouping',
  'BlanksAs',
  'PhoneticType',
  'PhoneticAlignment',
  'PageOrder',
  'PaperSize',
  'TableTotalsRowFunction',
  'TableStyleElementType',
  'ChartTickLabelPosition',
  'ChartDomain',
  'WidthPlatform',
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
  'PIVOT_VALUES_FIELD',
  'mergeFunctionMetadata',
  'toNative',
  'enumName',
  'evaluateFormula',
  'evaluateTypedFormula',
  'convertCivilToSerial',
  'convertSerialToCivil',
  'getVersion',
  'getFunctionMetadata',
  'listFunctionNames',
  'localizeFunctionName',
  'canonicalizeFunctionName',
  'localizeFormula',
  'canonicalizeFormula',
  'getLocaleFacts',
  'listProfileIds',
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
  'setLogMinLevel',
  'getFunctionMetadata',
  'listFunctionNames',
  'localizeFunctionName',
  'canonicalizeFunctionName',
  'getLocaleFacts',
  'listProfileIds',
  'compareText',
  'getImageInfo',
]);

const FORWARDED_OPTIONS = ['locateFile', 'print', 'printErr'];
const FORMULON_OPTION_KEYS = Object.freeze([...FORWARDED_OPTIONS, 'wasm']);

function dateSystem(options, where) {
  if (options === undefined) return false;
  const label = argLabel(where, 1);
  requireObject(options, label);
  return options.date1904 === undefined ? false : requireBoolean(options.date1904, `${label}.date1904`);
}

const CIVIL_FIELDS = Object.freeze(['year', 'month', 'day', 'hour', 'minute', 'second']);
const CIVIL_DATE_FIELDS = Object.freeze(['year', 'month', 'day']);

/** The tagged form of a raw `{rows, cols, cells}` result (`cells` as rows): a 1x1 result is its one value. */
function typedResult(result) {
  if (result.rows === 1 && result.cols === 1) return result.cells[0][0];
  return { kind: ValueKind.Array, rows: result.rows, cols: result.cols, cells: result.cells.flat() };
}

/** A string formula argument; anything else is a `TypeError` naming `label`. */
function formulaText(formula, label) {
  return requireString(formula, label);
}

const LOG_LEVELS = Object.freeze({
  debug: LogLevel.Debug,
  info: LogLevel.Info,
  warn: LogLevel.Warn,
  error: LogLevel.Error,
});

/** The `LogEvent` a log line of the engine's JSON carries: `level` as its `LogLevel` ordinal. */
function logRecord(bytes) {
  const record = JSON.parse(new TextDecoder().decode(bytes));
  return { ...record, level: LOG_LEVELS[record.level] ?? record.level };
}

// Members built over a runtime function, keyed by member name: [the runtime function, the wrapper factory]. The
// civil-date wrappers default `date1904` and the time of day and refuse what the fields cannot represent; the
// engine judges the calendar. The evaluation wrappers fold the raw `{rows, cols, cells}` result. The formula-text
// wrappers always pass the optional options argument, which the runtime reads.
const RUNTIME_WRAPPERS = Object.freeze({
  canonicalizeFormula: [
    'canonicalizeFormula',
    (canonicalize) =>
      function canonicalizeFormula(formula, profileId, options) {
        return canonicalize(formula, profileId, options);
      },
  ],
  localizeFormula: [
    'localizeFormula',
    (localize) =>
      function localizeFormula(formula, profileId, options) {
        return localize(formula, profileId, options);
      },
  ],
  convertCivilToSerial: [
    'convertCivilToSerial',
    (convert) =>
      function convertCivilToSerial(date, options) {
        const where = 'Formulon.convertCivilToSerial';
        const civil = toCivil(date, where, argumentPath(where, 0));
        return convert(civil, dateSystem(options, where));
      },
  ],
  getCollationClasses: [
    'getCollationClasses',
    (classify) =>
      function getCollationClasses(values, strength, profileId) {
        const label = argLabel('Formulon.getCollationClasses', 0);
        requireArray(values, label);
        for (const [i, value] of values.entries()) requireValue(value, `${label}[${i}]`);
        return classify(values, strength, profileId);
      },
  ],
  getImageInfo: [
    'getImageInfo',
    (info) =>
      function getImageInfo(bytes) {
        return info(requireBytes(bytes, 'Formulon.getImageInfo', 'bytes'));
      },
  ],
  convertSerialToCivil: [
    'convertSerialToCivil',
    (convert) =>
      function convertSerialToCivil(serial, options) {
        const where = 'Formulon.convertSerialToCivil';
        const label = argLabel(where, 0);
        requireNumber(serial, label);
        if (!Number.isFinite(serial)) throw new RangeError(`${label} must be finite`);
        return convert(serial, dateSystem(options, where));
      },
  ],
  evaluateFormula: [
    'evaluateFormula',
    (evaluate) =>
      function evaluateFormula(formula) {
        return toNative(typedResult(evaluate(formulaText(formula, argLabel('Formulon.evaluateFormula', 0)))));
      },
  ],
  evaluateTypedFormula: [
    'evaluateFormula',
    (evaluate) =>
      function evaluateTypedFormula(formula) {
        return typedResult(evaluate(formulaText(formula, argLabel('Formulon.evaluateTypedFormula', 0))));
      },
  ],
  setLogSink: [
    'setLogSink',
    (setSink) =>
      function setLogSink(sink) {
        if (sink !== null && typeof sink !== 'function') {
          throw new TypeError(argumentMessage('Formulon.setLogSink', 'sink', 'must be a function'));
        }
        return setSink(sink === null ? null : recordThrown((bytes) => sink(logRecord(bytes))));
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
  installInputGuards(formulon, 'Formulon');
  installArityGuards(formulon, 'Formulon');
  installErrorLabels(formulon, 'Formulon', (name) => Object.hasOwn(METHOD_ARITY, `Formulon.${name}`));
  return Object.freeze(formulon);
}

function isBufferSource(value) {
  return ArrayBuffer.isView(value) || value instanceof ArrayBuffer;
}

/**
 * Refuses a `createFormulon` options bag the factory cannot take: not an object, a key outside `locateFile`,
 * `print`, `printErr` and `wasm`, a callback that is not a function, or a `wasm` that is not a `BufferSource` or
 * a `WebAssembly.Module`. `null` and `undefined` stand for no options, and a `null` value for an omitted key.
 *
 * @param {unknown} options What `createFormulon` was called with.
 */
export function checkFormulonOptions(options) {
  const where = 'createFormulon';
  if (options === undefined || options === null) return;
  if (!isObject(options)) throw new TypeError(argumentMessage(where, 'options', 'must be an object'));
  for (const [key, value] of Object.entries(options)) {
    if (!FORMULON_OPTION_KEYS.includes(key)) {
      throw new TypeError(argumentMessage(where, 'options', `has no field ${key}`));
    }
    if (value === undefined || value === null) continue;
    const path = `options.${key}`;
    if (key !== 'wasm') {
      if (typeof value !== 'function') throw new TypeError(argumentMessage(where, path, 'must be a function'));
    } else if (!(typeof WebAssembly === 'object' && value instanceof WebAssembly.Module)) {
      if (!isBufferSource(value)) {
        throw new TypeError(argumentMessage(where, path, 'must be a BufferSource or a WebAssembly.Module'));
      }
      if (value.byteLength === 0) throw new TypeError(argumentMessage(where, path, 'must not be empty'));
    }
  }
}

/**
 * The factory options `FormulonOptions.wasm` adds: a `BufferSource` as `wasmBinary`, a `WebAssembly.Module`
 * as an `instantiateWasm` that instantiates it, whose failure `fail` reports. `wasm` was checked by
 * `checkFormulonOptions`.
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
  checkFormulonOptions(options);
  const moduleOptions = { FormulonError };
  // The glue awaits `instantiateWasm` without a rejection path, so a failed instantiation settles here.
  let fail;
  const failed = new Promise((_resolve, reject) => {
    fail = reject;
  });
  if (options !== undefined && options !== null) {
    for (const key of FORWARDED_OPTIONS) {
      if (options[key] !== undefined && options[key] !== null) {
        moduleOptions[key] = key === 'locateFile' ? options[key] : recordThrown(options[key]);
      }
    }
    Object.assign(moduleOptions, wasmOptions(options.wasm, fail));
  }
  const module = await Promise.race([factory(moduleOptions), failed]);
  return buildFormulon(module, shared, registry);
}

// <facade-tables>
const WORKBOOK_METHODS = Object.freeze([
  '[Symbol.iterator]',
  'activeSheet',
  'addPerson',
  'addPivotCache',
  'addSheet',
  'copySheet',
  'dispose',
  'disposed',
  'formatReference',
  'formatValue',
  'getCalcMode',
  'getCalcOnSave',
  'getCellStyle',
  'getDate1904',
  'getDefaultFont',
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
  'hasSheet',
  'listCellStyles',
  'listDefinedNames',
  'listExternalLinks',
  'listPeople',
  'listPivotCaches',
  'listSheets',
  'listTableStyles',
  'listThemeScriptFonts',
  'ooxml',
  'parseReference',
  'pivotCache',
  'recalc',
  'recalcDataTables',
  'recalcParallel',
  'removeCellStyle',
  'removeDefinedName',
  'removePerson',
  'removePivotCache',
  'removeSheet',
  'removeTableStyle',
  'resetTheme',
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
  'setFileName',
  'setFullPrecision',
  'setIterativeProgress',
  'setPinnedNow',
  'setProfileId',
  'setRefMode',
  'setTableStyle',
  'setThemeColors',
  'setThemeFonts',
  'setThemeScriptFont',
  'setWorkbookProtection',
  'sheet',
  'sheetCount',
  'table',
  'updateIterative',
  'updateTableStyleDefaults',
]);
const SHEET_METHODS = Object.freeze([
  'addCfRule',
  'addChart',
  'addDataTable',
  'addMerge',
  'addPicture',
  'addPivot',
  'addProtectedRange',
  'addShape',
  'addSlicer',
  'addSparklineGroup',
  'addTable',
  'addThreadedComment',
  'addValidation',
  'applyAutoFilter',
  'chart',
  'clearAutoFilterCriteria',
  'clearBreaks',
  'clearCfRules',
  'clearHyperlinks',
  'clearMerges',
  'clearValidations',
  'convertColumnCharsToPt',
  'convertColumnPtToChars',
  'deleteCells',
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
  'getCfRule',
  'getColumnWidthPt',
  'getComment',
  'getCommentVisible',
  'getDataTable',
  'getDate',
  'getDependents',
  'getDisplayText',
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
  'getProtectedRange',
  'getRowHeightPt',
  'getSelection',
  'getSheetProtection',
  'getShowPhoneticCells',
  'getSparklineGroup',
  'getSpill',
  'getTabColor',
  'getThreadedComment',
  'getTypedValue',
  'getTypedValues',
  'getUsedRange',
  'getValidation',
  'getValue',
  'getValues',
  'getView',
  'getVisibility',
  'getWidthModel',
  'insertCells',
  'insertColumns',
  'insertHyperlinkAt',
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
  'listTables',
  'listThreadedComments',
  'listValidations',
  'picture',
  'pivot',
  'recalcViewport',
  'removeCfRule',
  'removeDataTable',
  'removeDrawing',
  'removeHyperlinkAt',
  'removeMerge',
  'removePivot',
  'removeProtectedRange',
  'removeSparklineGroup',
  'removeTable',
  'removeThreadedComment',
  'removeValidation',
  'resetColumnWidth',
  'resetRowHeight',
  'restoreDrawing',
  'setArrayFormula',
  'setAutoFilter',
  'setBlank',
  'setBool',
  'setCellPhonetic',
  'setCellPhoneticProperties',
  'setCellPhoneticRuns',
  'setCellShowPhonetic',
  'setColumnBreak',
  'setColumnCollapsed',
  'setColumnHidden',
  'setColumnOutline',
  'setColumnWidth',
  'setComment',
  'setCommentVisible',
  'setError',
  'setFormula',
  'setHyperlink',
  'setNumber',
  'setPrintArea',
  'setPrintTitles',
  'setRowAutoHeight',
  'setRowBreak',
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
  'setVisibility',
  'shape',
  'slicer',
  'table',
  'updateCellFormat',
  'updateCfRule',
  'updateFormatDefaults',
  'updateHeaderFooter',
  'updatePageMargins',
  'updatePageSetup',
  'updatePrintOptions',
  'updateProtectedRange',
  'updateSheetProtection',
  'updateSparklineGroup',
  'updateThreadedComment',
  'updateValidation',
  'updateView',
  'validateValue',
]);
const OOXML_WORKBOOK_METHODS = Object.freeze([
  'addBorder',
  'addCellStyleXf',
  'addCellXf',
  'addDxf',
  'addFill',
  'addFont',
  'addNumberFormat',
  'getBorder',
  'getCellStyleXf',
  'getCellXf',
  'getDxf',
  'getFill',
  'getFont',
  'getNumberFormat',
  'hasPassthrough',
  'listBorders',
  'listCellStyleXfs',
  'listCellXfs',
  'listDxfs',
  'listFills',
  'listFonts',
  'listNumberFormats',
  'listPassthroughs',
  'setBorder',
  'setCellStyleXf',
  'setCellXf',
  'setDxf',
  'setFill',
  'setFont',
]);
const OOXML_SHEET_METHODS = Object.freeze([
  'getAutoFilterXml',
  'getCellXfIndex',
  'getCellXfIndices',
  'getColumnXfIndex',
  'getEffectiveStyle',
  'getHeaderFooterXml',
  'getPageMarginsXml',
  'getPageSetupXml',
  'getPrintOptionsXml',
  'getRangeXfBases',
  'getRangeXfBasesExcluding',
  'getRowXfIndex',
  'getSheetPrXml',
  'remapRangeXfs',
  'resetCellXfIndex',
  'resetColumnXfIndex',
  'resetRowXfIndex',
  'restoreRangeXfs',
  'setAutoFilterXml',
  'setCellXfIndex',
  'setColumnXfIndex',
  'setHeaderFooterXml',
  'setPageMarginsXml',
  'setPageSetupXml',
  'setPrintOptionsXml',
  'setRangeXfIndex',
  'setRowXfIndex',
  'setSheetPrXml',
  'snapshotRangeXfs',
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
  'activeSheet',
  'addCfRule',
  'addDataTable',
  'addPicture',
  'addSheet',
  'chart',
  'copySheet',
  'drawing',
  'evaluateFormula',
  'evaluateTypedFormula',
  'evaluateTypedFormulaArray',
  'formatReference',
  'getCellFormat',
  'getCellPhoneticRuns',
  'getCellsInRange',
  'getColumnXfIndex',
  'getDate',
  'getDefaultFont',
  'getFormula',
  'getInvalidCells',
  'getRangeXfBasesExcluding',
  'getRowXfIndex',
  'getThemeColors',
  'getThemeFonts',
  'getTypedValue',
  'getTypedValues',
  'getUsedRange',
  'getValues',
  'hasSheet',
  'insertHyperlinkAt',
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
  'listThreadedComments',
  'picture',
  'pivot',
  'pivotCache',
  'remapRangeXfs',
  'removeDrawing',
  'removeSheet',
  'setActiveSheet',
  'setArrayFormula',
  'setComment',
  'setFormula',
  'setHyperlink',
  'setPinnedNow',
  'setThemeColors',
  'setThemeScriptFont',
  'setValue',
  'setValues',
  'shape',
  'slicer',
  'table',
  'updateCellFormat',
  'updateCfRule',
  'updateProtectedRange',
  'updateSparklineGroup',
  'updateThreadedComment',
  'updateValidation',
  'updateView',
]);
const HANDLE_FORWARDERS = Object.freeze({
  'Chart.addSeries': 'addChartSeries',
  'Chart.getData': 'getChartData',
  'Chart.getPreview': 'getChartPreview',
  'Chart.getProperties': 'getChart',
  'Chart.removeSeries': 'removeChartSeries',
  'Chart.setFilter': 'setChartFilter',
  'Chart.setLegend': 'setChartLegend',
  'Chart.setPoint': null,
  'Chart.setSource': null,
  'Chart.setTitle': 'setChartTitle',
  'Chart.setType': 'setChartType',
  'Chart.setZOrder': 'setDrawingZOrder',
  'Chart.snapshot': 'snapshotDrawing',
  'Chart.updateAxis': 'updateChartAxis',
  'Chart.updateDisplay': 'updateChartDisplay',
  'Chart.updateLegend': 'updateChartLegend',
  'Chart.updatePlot': 'updateChartPlot',
  'Chart.updatePoint': 'updateChartPoint',
  'Chart.updateSeries': 'updateChartSeries',
  'Drawing.getDrawingProperties': 'getDrawingObject',
  'Drawing.update': null,
  'Picture.getProperties': 'getPicture',
  'Picture.setZOrder': 'setDrawingZOrder',
  'Picture.snapshot': 'snapshotDrawing',
  'Pivot.addDataField': 'addPivotDataField',
  'Pivot.addFieldItem': 'addPivotFieldItem',
  'Pivot.addFilter': 'addPivotFilter',
  'Pivot.cache': 'getPivotCacheId',
  'Pivot.clearDataFields': 'clearPivotDataFields',
  'Pivot.clearFieldItems': 'clearPivotFieldItems',
  'Pivot.clearFields': 'clearPivotFields',
  'Pivot.clearFilters': 'clearPivotFilters',
  'Pivot.getColumnFieldOrder': 'getPivotColumnFieldOrder',
  'Pivot.getDataField': 'getPivotDataField',
  'Pivot.getDataFieldDefaultName': 'getPivotDataFieldDefaultName',
  'Pivot.getField': 'getPivotField',
  'Pivot.getFieldDateGroup': null,
  'Pivot.getFieldItem': 'getPivotFieldItem',
  'Pivot.getFieldNumericGroup': 'getPivotFieldNumericGroup',
  'Pivot.getFilter': 'getPivotFilter',
  'Pivot.getGrandTotals': 'getPivotGrandTotals',
  'Pivot.getLayout': 'getPivotLayout',
  'Pivot.getOptions': 'getPivotOptions',
  'Pivot.getPageFieldOrder': 'getPivotPageFieldOrder',
  'Pivot.getProperties': null,
  'Pivot.getReportLayout': 'getPivotReportLayout',
  'Pivot.getRowFieldOrder': 'getPivotRowFieldOrder',
  'Pivot.getSpill': 'getPivotSpill',
  'Pivot.getStyleOptions': 'getPivotStyleOptions',
  'Pivot.isProjectionPending': 'isPivotProjectionPending',
  'Pivot.listDataFields': null,
  'Pivot.listFieldItems': null,
  'Pivot.listFields': null,
  'Pivot.listFilters': null,
  'Pivot.project': 'projectPivot',
  'Pivot.removeFilter': 'removePivotFilter',
  'Pivot.rename': 'renamePivot',
  'Pivot.setAnchor': 'setPivotAnchor',
  'Pivot.setColumnFieldOrder': 'setPivotColumnFieldOrder',
  'Pivot.setDataField': 'setPivotDataField',
  'Pivot.setFieldAxis': 'setPivotFieldAxis',
  'Pivot.setFieldDateGroup': null,
  'Pivot.setFieldNumericGroup': null,
  'Pivot.setFieldSubtotalFunctions': 'setPivotFieldSubtotalFunctions',
  'Pivot.setPageFieldOrder': 'setPivotPageFieldOrder',
  'Pivot.setReportLayout': 'setPivotReportLayout',
  'Pivot.setRowFieldOrder': 'setPivotRowFieldOrder',
  'Pivot.showDetails': null,
  'Pivot.updateField': null,
  'Pivot.updateFieldItem': null,
  'Pivot.updateGrandTotals': null,
  'Pivot.updateOptions': 'updatePivotOptions',
  'Pivot.updateStyleOptions': 'updatePivotStyleOptions',
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
  'PivotCache.getRecordValue': 'getPivotCacheRecordValue',
  'PivotCache.getWorksheetSource': 'getPivotCacheWorksheetSource',
  'PivotCache.listCalculatedItems': null,
  'PivotCache.listFieldSharedItems': null,
  'PivotCache.listFields': null,
  'PivotCache.moveCalculatedItem': 'movePivotCacheCalculatedItem',
  'PivotCache.refresh': 'refreshPivotCache',
  'PivotCache.removeCalculatedField': 'removePivotCacheCalculatedField',
  'PivotCache.removeCalculatedItem': 'removePivotCacheCalculatedItem',
  'PivotCache.setCalculatedFieldFormula': 'setPivotCacheCalculatedFieldFormula',
  'PivotCache.setCalculatedItemFormula': 'setPivotCacheCalculatedItemFormula',
  'PivotCache.setRecordValue': null,
  'PivotCache.setWorksheetSource': null,
  'PivotCache.updateOptions': 'updatePivotCacheOptions',
  'Shape.getProperties': 'getShape',
  'Shape.setZOrder': 'setDrawingZOrder',
  'Shape.snapshot': 'snapshotDrawing',
  'Shape.update': null,
  'Sheet.move': null,
  'Sheet.rename': null,
  'Slicer.getProperties': 'getSlicer',
  'Slicer.setZOrder': 'setDrawingZOrder',
  'Slicer.update': null,
  'Table.applyAutoFilter': 'applyTableAutoFilter',
  'Table.clearAutoFilterCriteria': 'clearTableAutoFilterCriteria',
  'Table.evaluateAutoFilter': 'evaluateTableAutoFilter',
  'Table.getAutoFilter': 'getTableAutoFilter',
  'Table.getProperties': 'getTable',
  'Table.getStyleOptions': 'getTableStyleOptions',
  'Table.insertColumns': null,
  'Table.insertRows': null,
  'Table.rename': null,
  'Table.renameColumn': 'renameTableColumn',
  'Table.setAutoFilter': 'setTableAutoFilter',
  'Table.update': null,
  'Table.updateStyleOptions': 'updateTableStyleOptions',
});
const HANDLE_PROPERTIES = Object.freeze({
  'Pivot.name': 'getPivotName',
});
const CELL_ARGS = Object.freeze({
  'Pivot.showDetails': [[0, 1]],
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
  'Sheet.getSpill': [[0, 1]],
  'Sheet.getTypedValue': [[0, 1]],
  'Sheet.getValue': [[0, 1]],
  'Sheet.removeDataTable': [[0, 1]],
  'Sheet.resetCellXfIndex': [[0, 1]],
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
  'Sheet.setHyperlink': [[0, 1]],
  'Sheet.setNumber': [[0, 1]],
  'Sheet.setText': [[0, 1]],
  'Sheet.setValue': [[0, 1]],
  'Sheet.setValues': [[0, 1]],
  'Sheet.validateValue': [[0, 1]],
});
const RECT_ARGS = Object.freeze({
  'Sheet.addMerge': [0],
  'Sheet.deleteCells': [0],
  'Sheet.evaluateCfRange': [0],
  'Sheet.getCellRectPt': [0],
  'Sheet.getCellXfIndices': [0],
  'Sheet.getCellsInRange': [0],
  'Sheet.getMergesInRange': [0],
  'Sheet.getRangeXfBases': [0],
  'Sheet.getRangeXfBasesExcluding': [0],
  'Sheet.getShowPhoneticCells': [0],
  'Sheet.getTypedValues': [0],
  'Sheet.getValues': [0],
  'Sheet.insertCells': [0],
  'Sheet.iterRows': [0],
  'Sheet.listCells': [0],
  'Sheet.recalcViewport': [0],
  'Sheet.remapRangeXfs': [0],
  'Sheet.removeMerge': [0],
  'Sheet.setPrintArea': [0],
  'Sheet.setRangeXfIndex': [0],
  'Sheet.snapshotRangeXfs': [0],
  'Sheet.updateCellFormat': [0],
  'Workbook.formatReference': [0],
});
const SPAN_ARGS = Object.freeze({
  'Sheet.deleteColumns': [[0, 'col', 'span']],
  'Sheet.deleteRows': [[0, 'row', 'span']],
  'Sheet.getColumnWidthPt': [[0, 'col', 'single']],
  'Sheet.getColumnXfIndex': [[0, 'col', 'single']],
  'Sheet.getRowHeightPt': [[0, 'row', 'single']],
  'Sheet.getRowXfIndex': [[0, 'row', 'single']],
  'Sheet.insertColumns': [[0, 'col', 'span']],
  'Sheet.insertRows': [[0, 'row', 'span']],
  'Sheet.resetColumnWidth': [[0, 'col', 'span']],
  'Sheet.resetColumnXfIndex': [[0, 'col', 'span']],
  'Sheet.resetRowHeight': [[0, 'row', 'span']],
  'Sheet.resetRowXfIndex': [[0, 'row', 'span']],
  'Sheet.setColumnCollapsed': [[0, 'col', 'span']],
  'Sheet.setColumnHidden': [[0, 'col', 'span']],
  'Sheet.setColumnOutline': [[0, 'col', 'span']],
  'Sheet.setColumnWidth': [[0, 'col', 'span']],
  'Sheet.setColumnXfIndex': [[0, 'col', 'span']],
  'Sheet.setRowAutoHeight': [[0, 'row', 'span']],
  'Sheet.setRowCollapsed': [[0, 'row', 'span']],
  'Sheet.setRowHeight': [[0, 'row', 'span']],
  'Sheet.setRowHidden': [[0, 'row', 'span']],
  'Sheet.setRowOutline': [[0, 'row', 'span']],
  'Sheet.setRowXfIndex': [[0, 'row', 'span']],
});
const REF_ARGS = Object.freeze({
  'Sheet.getCfRule': [0, 'cfRule'],
  'Sheet.getProtectedRange': [0, 'protectedRange'],
  'Sheet.getSparklineGroup': [0, 'sparklineGroup'],
  'Sheet.getValidation': [0, 'validation'],
  'Sheet.pivot': [0, 'pivot'],
  'Sheet.removeCfRule': [0, 'cfRule'],
  'Sheet.removePivot': [0, 'pivot'],
  'Sheet.removeProtectedRange': [0, 'protectedRange'],
  'Sheet.removeSparklineGroup': [0, 'sparklineGroup'],
  'Sheet.removeTable': [0, 'table'],
  'Sheet.removeValidation': [0, 'validation'],
  'Sheet.table': [0, 'table'],
  'Sheet.updateCfRule': [0, 'cfRule'],
  'Sheet.updateProtectedRange': [0, 'protectedRange'],
  'Sheet.updateSparklineGroup': [0, 'sparklineGroup'],
  'Sheet.updateValidation': [0, 'validation'],
  'Workbook.pivotCache': [0, 'pivotCache'],
  'Workbook.removePivotCache': [0, 'pivotCache'],
});
const PATCH_KEYS = Object.freeze({
  'Chart.updateAxis': [
    1,
    [
      'crossAxis',
      'crosses',
      'crossesAt',
      'id',
      'kind',
      'logBase',
      'majorGridlines',
      'majorUnit',
      'max',
      'min',
      'minorGridlines',
      'minorUnit',
      'numberFormat',
      'position',
      'reversed',
      'sourceLinked',
      'tickLabelPosition',
      'title',
      'visible',
    ],
  ],
  'Chart.updateDisplay': [
    0,
    [
      'autoTitleDeleted',
      'axes',
      'blanksAs',
      'date1904',
      'filter',
      'isExtended',
      'isPivotChart',
      'legend',
      'naAsBlank',
      'objectId',
      'onChartsheet',
      'plotVisibleOnly',
      'plots',
      'source',
      'style',
      'title',
    ],
  ],
  'Chart.updateLegend': [0, ['hiddenEntries', 'overlay', 'position']],
  'Chart.updatePlot': [
    1,
    [
      'axisIds',
      'barDirection',
      'extendedLayout',
      'firstSliceAngleDeg',
      'gapWidthPct',
      'grouping',
      'holeSizePct',
      'overlapPct',
      'series',
      'type',
      'varyColors',
    ],
  ],
  'Chart.updatePoint': [3, ['fill', 'index', 'line', 'marker']],
  'Chart.updateSeries': [
    2,
    [
      'categories',
      'fill',
      'filtered',
      'labels',
      'line',
      'marker',
      'name',
      'order',
      'points',
      'seriesId',
      'size',
      'smooth',
      'values',
      'x',
      'y',
    ],
  ],
  'Drawing.update': [0, ['anchor', 'description', 'hidden', 'kind', 'name', 'objectId']],
  'Pivot.updateField': [
    1,
    [
      'axis',
      'cacheFieldIndex',
      'customName',
      'defaultSubtotal',
      'fillDownLabels',
      'groupBaseIndex',
      'insertBlankRow',
      'numberFormat',
      'showAll',
      'showDetail',
      'sortByField',
      'sortOrder',
      'sourceName',
      'subtotalFunctions',
      'subtotalTop',
    ],
  ],
  'Pivot.updateFieldItem': [2, ['cacheIndex', 'caption', 'name', 'showDetail', 'visible']],
  'Pivot.updateGrandTotals': [0, ['columns', 'rows']],
  'Pivot.updateOptions': [
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
      'showEmptyColumn',
      'showEmptyRow',
      'showError',
      'showHeaders',
      'showMissing',
    ],
  ],
  'Pivot.updateStyleOptions': [
    0,
    ['showColumnHeaders', 'showColumnStripes', 'showLastColumn', 'showRowHeaders', 'showRowStripes', 'styleName'],
  ],
  'PivotCache.updateOptions': [0, ['missingItemsLimit', 'refreshOnLoad', 'saveData']],
  'Shape.update': [
    0,
    [
      'anchor',
      'description',
      'fill',
      'hidden',
      'kind',
      'line',
      'name',
      'objectId',
      'presetGeometry',
      'rotationDeg',
      'text',
      'textBox',
    ],
  ],
  'Sheet.updateCellFormat': [1, ['alignment', 'border', 'fill', 'font', 'numberFormat', 'protection', 'source']],
  'Sheet.updateCfRule': [
    1,
    [
      'aboveAverage',
      'bottom',
      'colorScale',
      'dataBar',
      'dxfIndex',
      'equalAverage',
      'formula1',
      'formula2',
      'iconSet',
      'id',
      'operator',
      'percent',
      'priority',
      'ranges',
      'rank',
      'stdDev',
      'stopIfTrue',
      'text',
      'timePeriod',
      'type',
    ],
  ],
  'Sheet.updateFormatDefaults': [
    0,
    [
      'baseColumnWidthChars',
      'customHeight',
      'defaultColumnWidthChars',
      'defaultRowHeightPt',
      'effectiveDefaultColumnWidthChars',
      'effectiveDefaultRowHeightPt',
      'zeroHeight',
    ],
  ],
  'Sheet.updateHeaderFooter': [
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
  'Sheet.updatePageMargins': [
    0,
    [
      'bottomIn',
      'bottomInStated',
      'footerIn',
      'footerInStated',
      'headerIn',
      'headerInStated',
      'leftIn',
      'leftInStated',
      'rightIn',
      'rightInStated',
      'topIn',
      'topInStated',
    ],
  ],
  'Sheet.updatePageSetup': [
    0,
    [
      'fitToHeight',
      'fitToHeightStated',
      'fitToPage',
      'fitToPageStated',
      'fitToWidth',
      'fitToWidthStated',
      'orientation',
      'paperSize',
      'paperSizeStated',
      'scalePct',
      'scalePctStated',
    ],
  ],
  'Sheet.updatePrintOptions': [
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
  'Sheet.updateProtectedRange': [
    1,
    [
      'algorithmName',
      'hashValue',
      'legacyPasswordHash',
      'name',
      'ranges',
      'saltValue',
      'securityDescriptor',
      'spinCount',
    ],
  ],
  'Sheet.updateSheetProtection': [
    0,
    [
      'algorithmName',
      'allowAutoFilter',
      'allowDeleteColumns',
      'allowDeleteRows',
      'allowEditObjects',
      'allowEditScenarios',
      'allowFormatCells',
      'allowFormatColumns',
      'allowFormatRows',
      'allowInsertColumns',
      'allowInsertHyperlinks',
      'allowInsertRows',
      'allowPivotTables',
      'allowSelectLockedCells',
      'allowSelectUnlockedCells',
      'allowSort',
      'hashValue',
      'legacyPasswordHash',
      'saltValue',
      'spinCount',
    ],
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
      'dateAxisRange',
      'displayHidden',
      'displayXAxis',
      'emptyCellsAs',
      'first',
      'high',
      'id',
      'last',
      'lineWidthPt',
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
  'Sheet.updateThreadedComment': [1, ['cell', 'created', 'done', 'id', 'mentions', 'parentId', 'personId', 'text']],
  'Sheet.updateValidation': [
    1,
    [
      'allowBlank',
      'errorMessage',
      'errorStyle',
      'errorTitle',
      'formula1',
      'formula2',
      'id',
      'operator',
      'promptMessage',
      'promptTitle',
      'ranges',
      'showErrorMessage',
      'showInCellDropdown',
      'showInputMessage',
      'type',
    ],
  ],
  'Sheet.updateView': [
    0,
    [
      'freezeCols',
      'freezeRows',
      'paneTopLeft',
      'rightToLeft',
      'showFormulas',
      'showGridLines',
      'showOutlineSymbols',
      'showRowColumnHeaders',
      'showZeros',
      'splitXTwips',
      'splitYTwips',
      'tabSelected',
      'topLeft',
      'viewMode',
      'zoomPct',
    ],
  ],
  'Slicer.update': [
    0,
    [
      'anchor',
      'cacheName',
      'caption',
      'columnCount',
      'columnName',
      'crossFilter',
      'description',
      'hidden',
      'kind',
      'objectId',
      'rowHeightPt',
      'sortOrder',
      'styleName',
      'tableName',
    ],
  ],
  'Table.update': [0, ['columns', 'headerRow', 'name', 'range', 'totalsRow']],
  'Table.updateStyleOptions': [
    0,
    ['showColumnStripes', 'showFirstColumn', 'showLastColumn', 'showRowStripes', 'styleName'],
  ],
  'Workbook.updateIterative': [0, ['enabled', 'maxChange', 'maxIterations']],
  'Workbook.updateTableStyleDefaults': [0, ['pivotStyleName', 'tableStyleName']],
});
const PATCH_NESTED_KEYS = Object.freeze({
  'Chart.updatePoint': {
    fill: ['color', 'none', 'other', 'transparencyPct'],
    line: ['color', 'dash', 'none', 'other', 'transparencyPct', 'widthPt'],
    marker: ['fill', 'line', 'sizePt', 'symbol'],
    'marker.fill': ['color', 'none', 'other', 'transparencyPct'],
    'marker.line': ['color', 'dash', 'none', 'other', 'transparencyPct', 'widthPt'],
  },
  'Chart.updateSeries': {
    fill: ['color', 'none', 'other', 'transparencyPct'],
    labels: [
      'deleted',
      'fontColor',
      'fontSizePt',
      'numberFormat',
      'position',
      'separator',
      'showCategory',
      'showLegendKey',
      'showPercent',
      'showSeries',
      'showValue',
      'sourceLinked',
    ],
    line: ['color', 'dash', 'none', 'other', 'transparencyPct', 'widthPt'],
    marker: ['fill', 'line', 'sizePt', 'symbol'],
    'marker.fill': ['color', 'none', 'other', 'transparencyPct'],
    'marker.line': ['color', 'dash', 'none', 'other', 'transparencyPct', 'widthPt'],
  },
  'Shape.update': {
    fill: ['color', 'none', 'other', 'transparencyPct'],
    line: ['color', 'dash', 'none', 'other', 'transparencyPct', 'widthPt'],
  },
  'Sheet.updateCellFormat': {
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
    fill: ['background', 'foreground', 'gradient', 'pattern'],
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
      'sizePt',
      'strike',
      'underline',
      'vertAlign',
    ],
    protection: ['hidden', 'locked'],
  },
  'Table.update': {
    'columns[]': ['calculatedFormula', 'name', 'totalsRowFormula', 'totalsRowFunction', 'totalsRowLabel'],
  },
});
// </facade-tables>
// <record-range-fields> generated by tools/codegen/gen_facade_tables.py; do not edit
const RECORD_RANGE_FIELDS = Object.freeze({
  ArrayFormulaInput: { range: 'CellRect' },
  AutoFilterInput: { range: 'CellRect', sort: 'AutoFilterSortStateInput' },
  AutoFilterSortCondition: { range: 'CellRect' },
  AutoFilterSortConditionInput: { range: 'CellRect' },
  AutoFilterSortState: { range: 'CellRect', conditions: 'AutoFilterSortCondition[]' },
  AutoFilterSortStateInput: { range: 'CellRect', conditions: 'AutoFilterSortConditionInput[]' },
  CfAboveAverageRuleInput: { ranges: 'CellRect[]' },
  CfCellIsRuleInput: { ranges: 'CellRect[]' },
  CfColorScaleRuleInput: { ranges: 'CellRect[]' },
  CfDataBarRuleInput: { ranges: 'CellRect[]' },
  CfExpressionRuleInput: { ranges: 'CellRect[]' },
  CfIconSetRuleInput: { ranges: 'CellRect[]' },
  CfPlainRuleInput: { ranges: 'CellRect[]' },
  CfTextRuleInput: { ranges: 'CellRect[]' },
  CfTimePeriodRuleInput: { ranges: 'CellRect[]' },
  CfTop10RuleInput: { ranges: 'CellRect[]' },
  ChartInput: { range: 'SheetRect' },
  ChartPreviewOverridesInput: { range: 'SheetRect' },
  ChartSourceInput: { range: 'SheetRect' },
  ConditionalFormatInput: { ranges: 'CellRect[]' },
  DataTableInput: { range: 'CellRect', rowInput: 'CellAddress', columnInput: 'CellAddress' },
  DataValidationInput: { ranges: 'CellRect[]' },
  HyperlinkInput: { range: 'CellRect' },
  PivotInput: { anchor: 'CellAddress' },
  PivotWorksheetSource: { range: 'SheetRect' },
  PivotWorksheetSourceInput: { range: 'SheetRect' },
  ProtectedRangeInput: { ranges: 'CellRect[]' },
  SheetSelectionInput: { active: 'CellAddress', ranges: 'CellRect[]' },
  SparklineGroupInput: { dateAxisRange: 'SheetRect', sparklines: 'SparklineInput[]' },
  SparklineInput: { source: 'SheetRect', location: 'CellAddress' },
  TableInput: { range: 'CellRect' },
  ThreadedCommentInput: { cell: 'CellAddress' },
});
// </record-range-fields>
// <input-keys> generated by tools/codegen/gen_facade_tables.py; do not edit
const INPUT_KEYS = Object.freeze({
  'Chart.addSeries': [
    1,
    [
      'categories',
      'fill',
      'filtered',
      'labels',
      'line',
      'marker',
      'name',
      'order',
      'points',
      'seriesId',
      'size',
      'smooth',
      'values',
      'x',
      'y',
    ],
  ],
  'Chart.getPreview': [
    0,
    ['blanksAs', 'direction', 'headerColumn', 'headerRow', 'naAsBlank', 'plotVisibleOnly', 'range'],
  ],
  'Chart.setFilter': [0, ['hiddenCategories', 'hiddenSeries']],
  'Chart.setLegend': [0, ['hiddenEntries', 'overlay', 'position']],
  'Chart.setPoint': [3, ['fill', 'index', 'line', 'marker']],
  'Chart.setSource': [0, ['direction', 'headerColumn', 'headerRow', 'range']],
  'Chart.setTitle': [0, ['formula', 'kind', 'text']],
  'Formulon.canonicalizeFormula': [2, ['baseCol', 'baseRow', 'r1c1']],
  'Formulon.convertCivilToSerial': [1, ['date1904']],
  'Formulon.convertSerialToCivil': [1, ['date1904']],
  'Formulon.localizeFormula': [2, ['baseCol', 'baseRow', 'r1c1']],
  'Pivot.addDataField': [
    0,
    ['aggregation', 'fieldIndex', 'name', 'numberFormat', 'showAs', 'showAsBaseField', 'showAsBaseItem'],
  ],
  'Pivot.addFilter': [0, ['dataFieldIndex', 'fieldIndex', 'kind', 'value', 'valueHigh']],
  'Pivot.setDataField': [
    1,
    ['aggregation', 'fieldIndex', 'name', 'numberFormat', 'showAs', 'showAsBaseField', 'showAsBaseItem'],
  ],
  'Pivot.setFieldDateGroup': [1, ['calendar', 'end', 'intervalDays', 'periods', 'start']],
  'Pivot.setFieldNumericGroup': [1, ['end', 'interval', 'start']],
  'PivotCache.setWorksheetSource': [0, ['name', 'range', 'sheet']],
  'Sheet.addCfRule': [
    0,
    [
      'aboveAverage',
      'bottom',
      'colorScale',
      'dataBar',
      'dxfIndex',
      'equalAverage',
      'format',
      'formula1',
      'formula2',
      'iconSet',
      'id',
      'operator',
      'percent',
      'priority',
      'ranges',
      'rank',
      'stdDev',
      'stopIfTrue',
      'text',
      'timePeriod',
      'type',
    ],
  ],
  'Sheet.addChart': [
    0,
    [
      'anchor',
      'barDirection',
      'description',
      'direction',
      'grouping',
      'headerColumn',
      'headerRow',
      'hidden',
      'legend',
      'name',
      'range',
      'title',
      'type',
    ],
  ],
  'Sheet.addDataTable': [0, ['columnInput', 'range', 'rowInput', 'rowOriented', 'twoDimensional']],
  'Sheet.addPicture': [1, ['anchor', 'description', 'hidden', 'name']],
  'Sheet.addPivot': [0, ['anchor', 'cacheId', 'name']],
  'Sheet.addProtectedRange': [
    0,
    [
      'algorithmName',
      'hashValue',
      'legacyPasswordHash',
      'name',
      'ranges',
      'saltValue',
      'securityDescriptor',
      'spinCount',
    ],
  ],
  'Sheet.addShape': [
    0,
    [
      'anchor',
      'description',
      'fill',
      'hidden',
      'kind',
      'line',
      'name',
      'objectId',
      'presetGeometry',
      'rotationDeg',
      'text',
      'textBox',
    ],
  ],
  'Sheet.addSlicer': [
    2,
    [
      'anchor',
      'cacheName',
      'caption',
      'columnCount',
      'columnName',
      'crossFilter',
      'description',
      'hidden',
      'name',
      'objectId',
      'rowHeightPt',
      'sortOrder',
      'styleName',
      'tableName',
    ],
  ],
  'Sheet.addSparklineGroup': [
    0,
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
      'dateAxisRange',
      'displayHidden',
      'displayXAxis',
      'emptyCellsAs',
      'first',
      'high',
      'id',
      'last',
      'lineWidthPt',
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
  'Sheet.addTable': [
    0,
    [
      'columns',
      'headerRow',
      'name',
      'range',
      'showColumnStripes',
      'showFirstColumn',
      'showLastColumn',
      'showRowStripes',
      'styleName',
      'styled',
      'totalsRow',
    ],
  ],
  'Sheet.addThreadedComment': [0, ['cell', 'created', 'done', 'id', 'mentions', 'parentId', 'personId', 'text']],
  'Sheet.addValidation': [
    0,
    [
      'allowBlank',
      'errorMessage',
      'errorStyle',
      'errorTitle',
      'formula1',
      'formula2',
      'id',
      'operator',
      'promptMessage',
      'promptTitle',
      'ranges',
      'showErrorMessage',
      'showInCellDropdown',
      'showInputMessage',
      'type',
    ],
  ],
  'Sheet.evaluateCfRange': [1, ['timePeriodRules', 'today']],
  'Sheet.getCellsInRange': [1, ['cursor', 'limit']],
  'Sheet.getInvalidCells': [0, ['cursor', 'limit']],
  'Sheet.insertHyperlinkAt': [1, ['display', 'location', 'range', 'target', 'tooltip']],
  'Sheet.restoreDrawing': [1, ['assignNewId']],
  'Sheet.setArrayFormula': [0, ['formula', 'range']],
  'Sheet.setAutoFilter': [0, ['columns', 'range', 'sort']],
  'Sheet.setCellPhoneticProperties': [2, ['alignment', 'fontIndex', 'stated', 'type']],
  'Sheet.setColumnBreak': [1, ['col', 'first', 'last', 'manual', 'row']],
  'Sheet.setComment': [2, ['author', 'text']],
  'Sheet.setHyperlink': [2, ['display', 'location', 'range', 'target', 'tooltip']],
  'Sheet.setPrintTitles': [0, ['columns', 'rows']],
  'Sheet.setRowBreak': [1, ['col', 'first', 'last', 'manual', 'row']],
  'Sheet.setSelection': [0, ['active', 'ranges']],
  'Sheet.setSheetProtection': [
    0,
    [
      'algorithmName',
      'allowAutoFilter',
      'allowDeleteColumns',
      'allowDeleteRows',
      'allowEditObjects',
      'allowEditScenarios',
      'allowFormatCells',
      'allowFormatColumns',
      'allowFormatRows',
      'allowInsertColumns',
      'allowInsertHyperlinks',
      'allowInsertRows',
      'allowPivotTables',
      'allowSelectLockedCells',
      'allowSelectUnlockedCells',
      'allowSort',
      'hashValue',
      'legacyPasswordHash',
      'saltValue',
      'spinCount',
    ],
  ],
  'Table.setAutoFilter': [0, ['columns', 'range', 'sort']],
  'Workbook.addBorder': [0, ['bottom', 'diagonal', 'diagonalDown', 'diagonalUp', 'left', 'right', 'top']],
  'Workbook.addCellStyleXf': [
    0,
    [
      'applyAlignment',
      'applyBorder',
      'applyFill',
      'applyFont',
      'applyNumberFormat',
      'applyProtection',
      'borderIndex',
      'fillIndex',
      'fontIndex',
      'hidden',
      'horizontal',
      'indent',
      'justifyLastLine',
      'locked',
      'numFmtId',
      'quotePrefix',
      'readingOrder',
      'relativeIndent',
      'shrinkToFit',
      'styleXfIndex',
      'textRotation',
      'vertical',
      'wrapText',
    ],
  ],
  'Workbook.addCellXf': [
    0,
    [
      'applyAlignment',
      'applyBorder',
      'applyFill',
      'applyFont',
      'applyNumberFormat',
      'applyProtection',
      'borderIndex',
      'fillIndex',
      'fontIndex',
      'hidden',
      'horizontal',
      'indent',
      'justifyLastLine',
      'locked',
      'numFmtId',
      'quotePrefix',
      'readingOrder',
      'relativeIndent',
      'shrinkToFit',
      'styleXfIndex',
      'textRotation',
      'vertical',
      'wrapText',
    ],
  ],
  'Workbook.addDxf': [0, ['alignment', 'border', 'fill', 'font', 'numberFormat', 'protection']],
  'Workbook.addFill': [0, ['background', 'foreground', 'gradient', 'pattern']],
  'Workbook.addFont': [
    0,
    [
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
      'sizePt',
      'strike',
      'underline',
      'vertAlign',
    ],
  ],
  'Workbook.addPerson': [0, ['displayName', 'id', 'providerId', 'userId']],
  'Workbook.addSheet': [1, ['insertBefore']],
  'Workbook.copySheet': [1, ['insertBefore', 'newName']],
  'Workbook.formatReference': [1, ['absolute', 'sheetName']],
  'Workbook.parseReference': [1, ['baseCol', 'baseRow', 'r1c1']],
  'Workbook.resolveColor': [0, ['argb', 'indexed', 'kind', 'theme', 'tint']],
  'Workbook.setBorder': [1, ['bottom', 'diagonal', 'diagonalDown', 'diagonalUp', 'left', 'right', 'top']],
  'Workbook.setCellStyle': [0, ['builtinId', 'customBuiltin', 'hidden', 'name', 'outlineLevel', 'styleXfIndex']],
  'Workbook.setCellStyleXf': [
    1,
    [
      'applyAlignment',
      'applyBorder',
      'applyFill',
      'applyFont',
      'applyNumberFormat',
      'applyProtection',
      'borderIndex',
      'fillIndex',
      'fontIndex',
      'hidden',
      'horizontal',
      'indent',
      'justifyLastLine',
      'locked',
      'numFmtId',
      'quotePrefix',
      'readingOrder',
      'relativeIndent',
      'shrinkToFit',
      'styleXfIndex',
      'textRotation',
      'vertical',
      'wrapText',
    ],
  ],
  'Workbook.setCellXf': [
    1,
    [
      'applyAlignment',
      'applyBorder',
      'applyFill',
      'applyFont',
      'applyNumberFormat',
      'applyProtection',
      'borderIndex',
      'fillIndex',
      'fontIndex',
      'hidden',
      'horizontal',
      'indent',
      'justifyLastLine',
      'locked',
      'numFmtId',
      'quotePrefix',
      'readingOrder',
      'relativeIndent',
      'shrinkToFit',
      'styleXfIndex',
      'textRotation',
      'vertical',
      'wrapText',
    ],
  ],
  'Workbook.setDefaultFont': [
    0,
    [
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
      'sizePt',
      'strike',
      'underline',
      'vertAlign',
    ],
  ],
  'Workbook.setDefinedName': [0, ['comment', 'formula', 'hidden', 'name', 'scope']],
  'Workbook.setDxf': [1, ['alignment', 'border', 'fill', 'font', 'numberFormat', 'protection']],
  'Workbook.setFill': [1, ['background', 'foreground', 'gradient', 'pattern']],
  'Workbook.setFont': [
    1,
    [
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
      'sizePt',
      'strike',
      'underline',
      'vertAlign',
    ],
  ],
  'Workbook.setTableStyle': [0, ['elements', 'name', 'pivot', 'table']],
  'Workbook.setThemeColors': [0, ['colors', 'source']],
  'Workbook.setThemeFonts': [0, ['majorEastAsian', 'majorLatin', 'minorEastAsian', 'minorLatin']],
  'Workbook.setThemeScriptFont': [0, ['scheme', 'script', 'typeface']],
  'Workbook.setWorkbookProtection': [
    0,
    [
      'lockRevision',
      'lockStructure',
      'lockWindows',
      'revisionsAlgorithmName',
      'revisionsHashValue',
      'revisionsLegacyPasswordHash',
      'revisionsSaltValue',
      'revisionsSpinCount',
      'workbookAlgorithmName',
      'workbookHashValue',
      'workbookLegacyPasswordHash',
      'workbookSaltValue',
      'workbookSpinCount',
    ],
  ],
});
const INPUT_NESTED_KEYS = Object.freeze({
  'Chart.addSeries': {
    fill: ['color', 'none', 'other', 'transparencyPct'],
    labels: [
      'deleted',
      'fontColor',
      'fontSizePt',
      'numberFormat',
      'position',
      'separator',
      'showCategory',
      'showLegendKey',
      'showPercent',
      'showSeries',
      'showValue',
      'sourceLinked',
    ],
    line: ['color', 'dash', 'none', 'other', 'transparencyPct', 'widthPt'],
    marker: ['fill', 'line', 'sizePt', 'symbol'],
    'marker.fill': ['color', 'none', 'other', 'transparencyPct'],
    'marker.line': ['color', 'dash', 'none', 'other', 'transparencyPct', 'widthPt'],
  },
  'Chart.setPoint': {
    fill: ['color', 'none', 'other', 'transparencyPct'],
    'fill.color': ['argb', 'indexed', 'kind', 'theme', 'tint'],
    line: ['color', 'dash', 'none', 'other', 'transparencyPct', 'widthPt'],
    'line.color': ['argb', 'indexed', 'kind', 'theme', 'tint'],
    marker: ['fill', 'line', 'sizePt', 'symbol'],
    'marker.fill': ['color', 'none', 'other', 'transparencyPct'],
    'marker.fill.color': ['argb', 'indexed', 'kind', 'theme', 'tint'],
    'marker.line': ['color', 'dash', 'none', 'other', 'transparencyPct', 'widthPt'],
    'marker.line.color': ['argb', 'indexed', 'kind', 'theme', 'tint'],
  },
  'Sheet.addCfRule': {
    colorScale: ['colors', 'thresholds'],
    'colorScale.thresholds[]': ['gte', 'type', 'value'],
    dataBar: [
      'axisColor',
      'axisPosition',
      'border',
      'direction',
      'fill',
      'gradient',
      'max',
      'maxLengthPct',
      'min',
      'minLengthPct',
      'negativeBorder',
      'negativeFill',
      'showValue',
    ],
    'dataBar.max': ['gte', 'type', 'value'],
    'dataBar.min': ['gte', 'type', 'value'],
    format: ['alignment', 'border', 'fill', 'font', 'numberFormat', 'protection', 'source'],
    'format.alignment': [
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
    'format.border': ['bottom', 'diagonal', 'diagonalDown', 'diagonalUp', 'left', 'right', 'top'],
    'format.border.bottom': ['color', 'style'],
    'format.border.diagonal': ['color', 'style'],
    'format.border.left': ['color', 'style'],
    'format.border.right': ['color', 'style'],
    'format.border.top': ['color', 'style'],
    'format.fill': ['background', 'foreground', 'gradient', 'pattern'],
    'format.fill.gradient': ['angleDeg', 'bottom', 'left', 'right', 'stops', 'top', 'type'],
    'format.fill.gradient.stops[]': ['color', 'position'],
    'format.font': [
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
      'sizePt',
      'strike',
      'underline',
      'vertAlign',
    ],
    'format.protection': ['hidden', 'locked'],
    iconSet: ['floor', 'icons', 'name', 'percent', 'reverse', 'showValue', 'thresholds'],
    'iconSet.floor': ['gte', 'type', 'value'],
    'iconSet.icons[]': ['iconIndex', 'iconSet'],
    'iconSet.thresholds[]': ['gte', 'type', 'value'],
  },
  'Sheet.addChart': {
    anchor: [
      'editAs',
      'fromCol',
      'fromColOffsetEmu',
      'fromRow',
      'fromRowOffsetEmu',
      'heightEmu',
      'kind',
      'toCol',
      'toColOffsetEmu',
      'toRow',
      'toRowOffsetEmu',
      'widthEmu',
    ],
  },
  'Sheet.addDataTable': {
    columnInput: ['col', 'deleted', 'row'],
    rowInput: ['col', 'deleted', 'row'],
  },
  'Sheet.addPicture': {
    anchor: [
      'editAs',
      'fromCol',
      'fromColOffsetEmu',
      'fromRow',
      'fromRowOffsetEmu',
      'heightEmu',
      'kind',
      'toCol',
      'toColOffsetEmu',
      'toRow',
      'toRowOffsetEmu',
      'widthEmu',
    ],
  },
  'Sheet.addShape': {
    anchor: [
      'editAs',
      'fromCol',
      'fromColOffsetEmu',
      'fromRow',
      'fromRowOffsetEmu',
      'heightEmu',
      'kind',
      'toCol',
      'toColOffsetEmu',
      'toRow',
      'toRowOffsetEmu',
      'widthEmu',
    ],
    fill: ['color', 'none', 'other', 'transparencyPct'],
    line: ['color', 'dash', 'none', 'other', 'transparencyPct', 'widthPt'],
  },
  'Sheet.addSlicer': {
    anchor: [
      'editAs',
      'fromCol',
      'fromColOffsetEmu',
      'fromRow',
      'fromRowOffsetEmu',
      'heightEmu',
      'kind',
      'toCol',
      'toColOffsetEmu',
      'toRow',
      'toRowOffsetEmu',
      'widthEmu',
    ],
  },
  'Sheet.addSparklineGroup': {
    'sparklines[]': ['location', 'source'],
  },
  'Sheet.addThreadedComment': {
    'mentions[]': ['end', 'mentionId', 'personId', 'start'],
  },
  'Sheet.setAutoFilter': {
    'columns[]': [
      'cellColor',
      'columnOffset',
      'conditions',
      'customAnd',
      'dateGroups',
      'dxfIndex',
      'dynamicMaxValue',
      'dynamicType',
      'dynamicValue',
      'filterBlank',
      'filterValue',
      'hiddenButton',
      'iconIndex',
      'iconSet',
      'kind',
      'maxValueIso',
      'percent',
      'showButton',
      'top',
      'topValue',
      'valueIso',
      'values',
    ],
    'columns[].conditions[]': ['operator', 'value'],
    'columns[].dateGroups[]': ['day', 'grouping', 'hour', 'minute', 'month', 'second', 'year'],
    sort: ['caseSensitive', 'columnSort', 'conditions', 'range', 'sortMethod'],
    'sort.conditions[]': ['customList', 'descending', 'dxfIndex', 'iconIndex', 'iconSet', 'range', 'sortBy'],
  },
  'Table.setAutoFilter': {
    'columns[]': [
      'cellColor',
      'columnOffset',
      'conditions',
      'customAnd',
      'dateGroups',
      'dxfIndex',
      'dynamicMaxValue',
      'dynamicType',
      'dynamicValue',
      'filterBlank',
      'filterValue',
      'hiddenButton',
      'iconIndex',
      'iconSet',
      'kind',
      'maxValueIso',
      'percent',
      'showButton',
      'top',
      'topValue',
      'valueIso',
      'values',
    ],
    'columns[].conditions[]': ['operator', 'value'],
    'columns[].dateGroups[]': ['day', 'grouping', 'hour', 'minute', 'month', 'second', 'year'],
    sort: ['caseSensitive', 'columnSort', 'conditions', 'range', 'sortMethod'],
    'sort.conditions[]': ['customList', 'descending', 'dxfIndex', 'iconIndex', 'iconSet', 'range', 'sortBy'],
  },
  'Workbook.addBorder': {
    bottom: ['color', 'style'],
    diagonal: ['color', 'style'],
    left: ['color', 'style'],
    right: ['color', 'style'],
    top: ['color', 'style'],
  },
  'Workbook.addDxf': {
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
    fill: ['background', 'foreground', 'gradient', 'pattern'],
    'fill.gradient': ['angleDeg', 'bottom', 'left', 'right', 'stops', 'top', 'type'],
    'fill.gradient.stops[]': ['color', 'position'],
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
      'sizePt',
      'strike',
      'underline',
      'vertAlign',
    ],
    protection: ['hidden', 'locked'],
  },
  'Workbook.addFill': {
    gradient: ['angleDeg', 'bottom', 'left', 'right', 'stops', 'top', 'type'],
    'gradient.stops[]': ['color', 'position'],
  },
  'Workbook.setBorder': {
    bottom: ['color', 'style'],
    diagonal: ['color', 'style'],
    left: ['color', 'style'],
    right: ['color', 'style'],
    top: ['color', 'style'],
  },
  'Workbook.setDxf': {
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
    fill: ['background', 'foreground', 'gradient', 'pattern'],
    'fill.gradient': ['angleDeg', 'bottom', 'left', 'right', 'stops', 'top', 'type'],
    'fill.gradient.stops[]': ['color', 'position'],
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
      'sizePt',
      'strike',
      'underline',
      'vertAlign',
    ],
    protection: ['hidden', 'locked'],
  },
  'Workbook.setFill': {
    gradient: ['angleDeg', 'bottom', 'left', 'right', 'stops', 'top', 'type'],
    'gradient.stops[]': ['color', 'position'],
  },
  'Workbook.setTableStyle': {
    'elements[]': ['dxfIndex', 'size', 'type'],
  },
});
const READ_ONLY_KEYS = Object.freeze({
  'Chart.addSeries': ['filtered', 'order', 'points', 'seriesId'],
  'Chart.updateAxis': ['crossAxis', 'id', 'kind'],
  'Chart.updateDisplay': [
    'autoTitleDeleted',
    'axes',
    'date1904',
    'filter',
    'isExtended',
    'isPivotChart',
    'legend',
    'objectId',
    'onChartsheet',
    'plots',
    'source',
    'title',
  ],
  'Chart.updatePlot': ['axisIds', 'barDirection', 'extendedLayout', 'grouping', 'series', 'type'],
  'Chart.updatePoint': ['index'],
  'Chart.updateSeries': ['filtered', 'order', 'points', 'seriesId'],
  'Drawing.update': ['kind', 'objectId'],
  'Pivot.updateField': ['axis', 'cacheFieldIndex', 'groupBaseIndex', 'sourceName', 'subtotalFunctions'],
  'Pivot.updateFieldItem': ['cacheIndex', 'name'],
  'Shape.update': ['kind', 'objectId'],
  'Sheet.addDataTable': ['columnInput.deleted', 'rowInput.deleted', 'rowOriented', 'twoDimensional'],
  'Sheet.addShape': ['kind', 'objectId'],
  'Sheet.addSlicer': ['cacheName', 'columnName', 'objectId', 'tableName'],
  'Sheet.addSparklineGroup': ['id'],
  'Sheet.addValidation': ['id'],
  'Sheet.setCellPhoneticProperties': ['stated'],
  'Sheet.setColumnBreak': ['col', 'row'],
  'Sheet.setRowBreak': ['col', 'row'],
  'Sheet.updateCellFormat': ['source'],
  'Sheet.updateCfRule': ['id', 'type'],
  'Sheet.updateFormatDefaults': ['effectiveDefaultColumnWidthChars', 'effectiveDefaultRowHeightPt'],
  'Sheet.updateHeaderFooter': [
    'alignWithMarginsStated',
    'differentFirstStated',
    'differentOddEvenStated',
    'scaleWithDocStated',
  ],
  'Sheet.updatePageMargins': [
    'bottomInStated',
    'footerInStated',
    'headerInStated',
    'leftInStated',
    'rightInStated',
    'topInStated',
  ],
  'Sheet.updatePageSetup': [
    'fitToHeightStated',
    'fitToPageStated',
    'fitToWidthStated',
    'paperSizeStated',
    'scalePctStated',
  ],
  'Sheet.updatePrintOptions': [
    'gridLinesStated',
    'headingsStated',
    'horizontalCenteredStated',
    'verticalCenteredStated',
  ],
  'Sheet.updateSparklineGroup': ['id'],
  'Sheet.updateThreadedComment': ['cell', 'created', 'id', 'parentId', 'personId'],
  'Sheet.updateValidation': ['id'],
  'Slicer.update': ['cacheName', 'columnName', 'kind', 'objectId', 'tableName'],
  'Table.update': ['name'],
  'Workbook.setThemeColors': ['source'],
});
// </input-keys>
// <patch-nullable> generated by tools/codegen/gen_facade_tables.py; do not edit
const PATCH_NULLABLE = Object.freeze({
  'Chart.updateAxis': {
    logBase: 'number',
    majorUnit: 'number',
    max: 'number',
    min: 'number',
    minorUnit: 'number',
    numberFormat: 'string',
    tickLabelPosition: 'number',
    title: 'object',
  },
  'Chart.updateDisplay': { blanksAs: 'number', naAsBlank: 'boolean', plotVisibleOnly: 'boolean', style: 'number' },
  'Chart.updatePoint': { fill: 'object', line: 'object', marker: 'object' },
  'Chart.updateSeries': {
    categories: 'object',
    fill: 'object',
    labels: 'object',
    line: 'object',
    marker: 'object',
    name: 'object',
    x: 'object',
  },
  'Pivot.updateField': { customName: 'string', numberFormat: 'string', sortByField: 'string' },
  'Pivot.updateFieldItem': { caption: 'string' },
  'Pivot.updateOptions': {
    altText: 'string',
    altTextSummary: 'string',
    errorCaption: 'string',
    grandTotalCaption: 'string',
    missingCaption: 'string',
  },
  'Pivot.updateStyleOptions': { styleName: 'string' },
  'PivotCache.updateOptions': { missingItemsLimit: 'number' },
  'Shape.update': { fill: 'object', line: 'object' },
  'Sheet.updateCfRule': { dxfIndex: 'number', formula1: 'string', formula2: 'string' },
  'Sheet.updateFormatDefaults': { defaultColumnWidthChars: 'number', defaultRowHeightPt: 'number' },
  'Sheet.updateHeaderFooter': {
    evenFooter: 'string',
    evenHeader: 'string',
    firstFooter: 'string',
    firstHeader: 'string',
    oddFooter: 'string',
    oddHeader: 'string',
  },
  'Sheet.updatePageSetup': { orientation: 'number' },
  'Sheet.updateProtectedRange': {
    algorithmName: 'string',
    hashValue: 'string',
    legacyPasswordHash: 'string',
    saltValue: 'string',
    securityDescriptor: 'string',
  },
  'Sheet.updateSheetProtection': {
    algorithmName: 'string',
    hashValue: 'string',
    legacyPasswordHash: 'string',
    saltValue: 'string',
    spinCount: 'number',
  },
  'Sheet.updateSparklineGroup': {
    colorAxis: 'object',
    colorFirst: 'object',
    colorHigh: 'object',
    colorLast: 'object',
    colorLow: 'object',
    colorMarkers: 'object',
    colorNegative: 'object',
    colorSeries: 'object',
    dateAxisRange: 'object',
    manualMax: 'number',
    manualMin: 'number',
  },
  'Sheet.updateValidation': {
    errorMessage: 'string',
    errorTitle: 'string',
    formula1: 'string',
    formula2: 'string',
    promptMessage: 'string',
    promptTitle: 'string',
  },
  'Slicer.update': { styleName: 'string' },
  'Table.updateStyleOptions': { styleName: 'string' },
  'Workbook.updateTableStyleDefaults': { pivotStyleName: 'string', tableStyleName: 'string' },
});
// </patch-nullable>
// <method-signatures> generated by tools/codegen/gen_facade_tables.py; do not edit
const METHOD_ARITY = Object.freeze({
  'Chart.addSeries': 2,
  'Chart.getData': 0,
  'Chart.getDrawingProperties': 0,
  'Chart.getPreview': 1,
  'Chart.getProperties': 0,
  'Chart.removeSeries': 2,
  'Chart.setFilter': 1,
  'Chart.setLegend': 1,
  'Chart.setPoint': 4,
  'Chart.setSource': 1,
  'Chart.setTitle': 1,
  'Chart.setType': 3,
  'Chart.setZOrder': 1,
  'Chart.snapshot': 0,
  'Chart.update': 1,
  'Chart.updateAxis': 2,
  'Chart.updateDisplay': 1,
  'Chart.updateLegend': 1,
  'Chart.updatePlot': 2,
  'Chart.updatePoint': 4,
  'Chart.updateSeries': 3,
  'Drawing.getDrawingProperties': 0,
  'Drawing.update': 1,
  'Formulon.canonicalizeFormula': 3,
  'Formulon.canonicalizeFunctionName': 2,
  'Formulon.compareText': 4,
  'Formulon.convertCivilToSerial': 2,
  'Formulon.convertSerialToCivil': 2,
  'Formulon.evaluateFormula': 1,
  'Formulon.evaluateTypedFormula': 1,
  'Formulon.getCollationClasses': 3,
  'Formulon.getFunctionMetadata': 1,
  'Formulon.getImageInfo': 1,
  'Formulon.getLocaleFacts': 1,
  'Formulon.getVersion': 0,
  'Formulon.listFunctionNames': 0,
  'Formulon.listProfileIds': 0,
  'Formulon.localizeFormula': 3,
  'Formulon.localizeFunctionName': 2,
  'Formulon.setLogMinLevel': 1,
  'Formulon.setLogSink': 1,
  'Picture.getDrawingProperties': 0,
  'Picture.getProperties': 0,
  'Picture.setZOrder': 1,
  'Picture.snapshot': 0,
  'Picture.update': 1,
  'Pivot.addDataField': 1,
  'Pivot.addFieldItem': 3,
  'Pivot.addFilter': 1,
  'Pivot.cache': 0,
  'Pivot.clearDataFields': 0,
  'Pivot.clearFieldItems': 1,
  'Pivot.clearFields': 0,
  'Pivot.clearFilters': 0,
  'Pivot.getColumnFieldOrder': 0,
  'Pivot.getDataField': 1,
  'Pivot.getDataFieldDefaultName': 3,
  'Pivot.getField': 1,
  'Pivot.getFieldDateGroup': 1,
  'Pivot.getFieldItem': 2,
  'Pivot.getFieldNumericGroup': 1,
  'Pivot.getFilter': 1,
  'Pivot.getGrandTotals': 0,
  'Pivot.getLayout': 0,
  'Pivot.getOptions': 0,
  'Pivot.getPageFieldOrder': 0,
  'Pivot.getProperties': 0,
  'Pivot.getReportLayout': 0,
  'Pivot.getRowFieldOrder': 0,
  'Pivot.getSpill': 0,
  'Pivot.getStyleOptions': 0,
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
  'Pivot.setFieldNumericGroup': 2,
  'Pivot.setFieldSubtotalFunctions': 2,
  'Pivot.setPageFieldOrder': 1,
  'Pivot.setReportLayout': 1,
  'Pivot.setRowFieldOrder': 1,
  'Pivot.showDetails': 2,
  'Pivot.updateField': 2,
  'Pivot.updateFieldItem': 3,
  'Pivot.updateGrandTotals': 1,
  'Pivot.updateOptions': 1,
  'Pivot.updateStyleOptions': 1,
  'PivotCache.addCalculatedField': 2,
  'PivotCache.addCalculatedItem': 3,
  'PivotCache.addField': 1,
  'PivotCache.addFieldSharedItem': 2,
  'PivotCache.addRecord': 0,
  'PivotCache.clearFieldSharedItems': 1,
  'PivotCache.clearFields': 0,
  'PivotCache.clearRecords': 0,
  'PivotCache.getOptions': 0,
  'PivotCache.getRecordCount': 0,
  'PivotCache.getRecordValue': 2,
  'PivotCache.getWorksheetSource': 0,
  'PivotCache.listCalculatedItems': 0,
  'PivotCache.listFieldSharedItems': 1,
  'PivotCache.listFields': 0,
  'PivotCache.moveCalculatedItem': 2,
  'PivotCache.refresh': 0,
  'PivotCache.removeCalculatedField': 1,
  'PivotCache.removeCalculatedItem': 1,
  'PivotCache.setCalculatedFieldFormula': 2,
  'PivotCache.setCalculatedItemFormula': 2,
  'PivotCache.setRecordValue': 3,
  'PivotCache.setWorksheetSource': 1,
  'PivotCache.updateOptions': 1,
  'Shape.getDrawingProperties': 0,
  'Shape.getProperties': 0,
  'Shape.setZOrder': 1,
  'Shape.snapshot': 0,
  'Shape.update': 1,
  'Sheet.addCfRule': 1,
  'Sheet.addChart': 1,
  'Sheet.addDataTable': 1,
  'Sheet.addMerge': 1,
  'Sheet.addPicture': 2,
  'Sheet.addPivot': 1,
  'Sheet.addProtectedRange': 1,
  'Sheet.addShape': 1,
  'Sheet.addSlicer': 3,
  'Sheet.addSparklineGroup': 1,
  'Sheet.addTable': 1,
  'Sheet.addThreadedComment': 1,
  'Sheet.addValidation': 1,
  'Sheet.applyAutoFilter': 0,
  'Sheet.chart': 1,
  'Sheet.clearAutoFilterCriteria': 0,
  'Sheet.clearBreaks': 0,
  'Sheet.clearCfRules': 0,
  'Sheet.clearHyperlinks': 0,
  'Sheet.clearMerges': 0,
  'Sheet.clearValidations': 0,
  'Sheet.convertColumnCharsToPt': 2,
  'Sheet.convertColumnPtToChars': 2,
  'Sheet.deleteCells': 2,
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
  'Sheet.getCellXfIndices': 1,
  'Sheet.getCellsInRange': 2,
  'Sheet.getCfRule': 1,
  'Sheet.getColumnWidthPt': 2,
  'Sheet.getColumnXfIndex': 1,
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
  'Sheet.getInvalidCells': 1,
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
  'Sheet.getProtectedRange': 1,
  'Sheet.getRangeXfBases': 1,
  'Sheet.getRangeXfBasesExcluding': 2,
  'Sheet.getRowHeightPt': 1,
  'Sheet.getRowXfIndex': 1,
  'Sheet.getSelection': 0,
  'Sheet.getSheetPrXml': 0,
  'Sheet.getSheetProtection': 0,
  'Sheet.getShowPhoneticCells': 1,
  'Sheet.getSparklineGroup': 1,
  'Sheet.getSpill': 2,
  'Sheet.getTabColor': 0,
  'Sheet.getThreadedComment': 1,
  'Sheet.getTypedValue': 2,
  'Sheet.getTypedValues': 1,
  'Sheet.getUsedRange': 0,
  'Sheet.getValidation': 1,
  'Sheet.getValue': 2,
  'Sheet.getValues': 1,
  'Sheet.getView': 0,
  'Sheet.getVisibility': 0,
  'Sheet.getWidthModel': 1,
  'Sheet.insertCells': 2,
  'Sheet.insertColumns': 2,
  'Sheet.insertHyperlinkAt': 2,
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
  'Sheet.listTables': 0,
  'Sheet.listThreadedComments': 2,
  'Sheet.listValidations': 0,
  'Sheet.move': 1,
  'Sheet.picture': 1,
  'Sheet.pivot': 1,
  'Sheet.recalcViewport': 1,
  'Sheet.remapRangeXfs': 2,
  'Sheet.removeCfRule': 1,
  'Sheet.removeDataTable': 2,
  'Sheet.removeDrawing': 1,
  'Sheet.removeHyperlinkAt': 1,
  'Sheet.removeMerge': 1,
  'Sheet.removePivot': 1,
  'Sheet.removeProtectedRange': 1,
  'Sheet.removeSparklineGroup': 1,
  'Sheet.removeTable': 1,
  'Sheet.removeThreadedComment': 1,
  'Sheet.removeValidation': 1,
  'Sheet.rename': 1,
  'Sheet.resetCellXfIndex': 2,
  'Sheet.resetColumnWidth': 2,
  'Sheet.resetColumnXfIndex': 2,
  'Sheet.resetRowHeight': 2,
  'Sheet.resetRowXfIndex': 2,
  'Sheet.restoreDrawing': 2,
  'Sheet.restoreRangeXfs': 1,
  'Sheet.setArrayFormula': 1,
  'Sheet.setAutoFilter': 1,
  'Sheet.setAutoFilterXml': 1,
  'Sheet.setBlank': 2,
  'Sheet.setBool': 3,
  'Sheet.setCellPhonetic': 3,
  'Sheet.setCellPhoneticProperties': 3,
  'Sheet.setCellPhoneticRuns': 3,
  'Sheet.setCellShowPhonetic': 3,
  'Sheet.setCellXfIndex': 3,
  'Sheet.setColumnBreak': 2,
  'Sheet.setColumnCollapsed': 3,
  'Sheet.setColumnHidden': 3,
  'Sheet.setColumnOutline': 3,
  'Sheet.setColumnWidth': 3,
  'Sheet.setColumnXfIndex': 3,
  'Sheet.setComment': 3,
  'Sheet.setCommentVisible': 3,
  'Sheet.setError': 3,
  'Sheet.setFormula': 3,
  'Sheet.setHeaderFooterXml': 1,
  'Sheet.setHyperlink': 3,
  'Sheet.setNumber': 3,
  'Sheet.setPageMarginsXml': 1,
  'Sheet.setPageSetupXml': 1,
  'Sheet.setPrintArea': 1,
  'Sheet.setPrintOptionsXml': 1,
  'Sheet.setPrintTitles': 1,
  'Sheet.setRangeXfIndex': 2,
  'Sheet.setRowAutoHeight': 3,
  'Sheet.setRowBreak': 2,
  'Sheet.setRowCollapsed': 3,
  'Sheet.setRowHeight': 3,
  'Sheet.setRowHidden': 3,
  'Sheet.setRowOutline': 3,
  'Sheet.setRowXfIndex': 3,
  'Sheet.setSelection': 1,
  'Sheet.setSheetPrXml': 1,
  'Sheet.setSheetProtection': 1,
  'Sheet.setTabColor': 1,
  'Sheet.setText': 3,
  'Sheet.setValue': 3,
  'Sheet.setValues': 3,
  'Sheet.setVisibility': 1,
  'Sheet.shape': 1,
  'Sheet.slicer': 1,
  'Sheet.snapshotRangeXfs': 1,
  'Sheet.table': 1,
  'Sheet.updateCellFormat': 2,
  'Sheet.updateCfRule': 2,
  'Sheet.updateFormatDefaults': 1,
  'Sheet.updateHeaderFooter': 1,
  'Sheet.updatePageMargins': 1,
  'Sheet.updatePageSetup': 1,
  'Sheet.updatePrintOptions': 1,
  'Sheet.updateProtectedRange': 2,
  'Sheet.updateSheetProtection': 1,
  'Sheet.updateSparklineGroup': 2,
  'Sheet.updateThreadedComment': 2,
  'Sheet.updateValidation': 2,
  'Sheet.updateView': 1,
  'Sheet.validateValue': 3,
  'Slicer.getDrawingProperties': 0,
  'Slicer.getProperties': 0,
  'Slicer.setZOrder': 1,
  'Slicer.update': 1,
  'Table.applyAutoFilter': 0,
  'Table.clearAutoFilterCriteria': 0,
  'Table.evaluateAutoFilter': 0,
  'Table.getAutoFilter': 0,
  'Table.getProperties': 0,
  'Table.getStyleOptions': 0,
  'Table.insertColumns': 2,
  'Table.insertRows': 2,
  'Table.rename': 1,
  'Table.renameColumn': 2,
  'Table.setAutoFilter': 1,
  'Table.update': 1,
  'Table.updateStyleOptions': 1,
  'Workbook.[Symbol.dispose]': 0,
  'Workbook.[Symbol.iterator]': 0,
  'Workbook.activeSheet': 0,
  'Workbook.addBorder': 1,
  'Workbook.addCellStyleXf': 1,
  'Workbook.addCellXf': 1,
  'Workbook.addDxf': 1,
  'Workbook.addFill': 1,
  'Workbook.addFont': 1,
  'Workbook.addNumberFormat': 1,
  'Workbook.addPerson': 1,
  'Workbook.addPivotCache': 1,
  'Workbook.addSheet': 2,
  'Workbook.copySheet': 2,
  'Workbook.createDefault': 0,
  'Workbook.createEmpty': 0,
  'Workbook.createForProfile': 1,
  'Workbook.dispose': 0,
  'Workbook.formatReference': 2,
  'Workbook.formatValue': 2,
  'Workbook.getBorder': 1,
  'Workbook.getCalcMode': 0,
  'Workbook.getCalcOnSave': 0,
  'Workbook.getCellStyle': 1,
  'Workbook.getCellStyleXf': 1,
  'Workbook.getCellXf': 1,
  'Workbook.getDate1904': 0,
  'Workbook.getDefaultFont': 0,
  'Workbook.getDefinedName': 2,
  'Workbook.getDiagnostics': 0,
  'Workbook.getDxf': 1,
  'Workbook.getFileName': 0,
  'Workbook.getFill': 1,
  'Workbook.getFont': 1,
  'Workbook.getFullPrecision': 0,
  'Workbook.getIterative': 0,
  'Workbook.getMemoryUsage': 0,
  'Workbook.getNumberFormat': 1,
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
  'Workbook.hasPassthrough': 1,
  'Workbook.hasSheet': 1,
  'Workbook.listBorders': 0,
  'Workbook.listCellStyleXfs': 0,
  'Workbook.listCellStyles': 0,
  'Workbook.listCellXfs': 0,
  'Workbook.listDefinedNames': 0,
  'Workbook.listDxfs': 0,
  'Workbook.listExternalLinks': 0,
  'Workbook.listFills': 0,
  'Workbook.listFonts': 0,
  'Workbook.listNumberFormats': 0,
  'Workbook.listPassthroughs': 0,
  'Workbook.listPeople': 0,
  'Workbook.listPivotCaches': 0,
  'Workbook.listSheets': 0,
  'Workbook.listTableStyles': 0,
  'Workbook.listThemeScriptFonts': 1,
  'Workbook.load': 1,
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
  'Workbook.removeTableStyle': 1,
  'Workbook.resetTheme': 0,
  'Workbook.resolveColor': 2,
  'Workbook.save': 1,
  'Workbook.saveWithDiagnostics': 1,
  'Workbook.setActiveSheet': 1,
  'Workbook.setBorder': 2,
  'Workbook.setCalcMode': 1,
  'Workbook.setCalcOnSave': 1,
  'Workbook.setCellStyle': 1,
  'Workbook.setCellStyleXf': 2,
  'Workbook.setCellXf': 2,
  'Workbook.setDate1904': 1,
  'Workbook.setDefaultFont': 1,
  'Workbook.setDefinedName': 1,
  'Workbook.setDxf': 2,
  'Workbook.setFileName': 1,
  'Workbook.setFill': 2,
  'Workbook.setFont': 2,
  'Workbook.setFullPrecision': 1,
  'Workbook.setIterativeProgress': 1,
  'Workbook.setPinnedNow': 1,
  'Workbook.setProfileId': 1,
  'Workbook.setRefMode': 1,
  'Workbook.setTableStyle': 1,
  'Workbook.setThemeColors': 1,
  'Workbook.setThemeFonts': 1,
  'Workbook.setThemeScriptFont': 1,
  'Workbook.setWorkbookProtection': 1,
  'Workbook.sheet': 1,
  'Workbook.table': 1,
  'Workbook.updateIterative': 1,
  'Workbook.updateTableStyleDefaults': 1,
});
const TYPED_ARGS = Object.freeze({
  'Chart.getPreview': [[0, 'ChartPreviewOverridesInput']],
  'Chart.setSource': [[0, 'ChartSourceInput']],
  'Pivot.setAnchor': [[0, 'CellAddress']],
  'PivotCache.setWorksheetSource': [[0, 'PivotWorksheetSourceInput']],
  'Sheet.addCfRule': [[0, 'ConditionalFormatInput']],
  'Sheet.addChart': [[0, 'ChartInput']],
  'Sheet.addDataTable': [[0, 'DataTableInput']],
  'Sheet.addPivot': [[0, 'PivotInput']],
  'Sheet.addProtectedRange': [[0, 'ProtectedRangeInput']],
  'Sheet.addSparklineGroup': [[0, 'SparklineGroupInput']],
  'Sheet.addTable': [[0, 'TableInput']],
  'Sheet.addThreadedComment': [[0, 'ThreadedCommentInput']],
  'Sheet.addValidation': [[0, 'DataValidationInput']],
  'Sheet.getRangeXfBasesExcluding': [[1, 'CellAddress[]']],
  'Sheet.insertHyperlinkAt': [[1, 'HyperlinkInput']],
  'Sheet.setArrayFormula': [[0, 'ArrayFormulaInput']],
  'Sheet.setAutoFilter': [[0, 'AutoFilterInput']],
  'Sheet.setHyperlink': [
    [1, 'HyperlinkInput'],
    [2, 'HyperlinkInput'],
  ],
  'Sheet.setPrintArea': [[0, 'CellRect[]']],
  'Sheet.setSelection': [[0, 'SheetSelectionInput']],
  'Table.setAutoFilter': [[0, 'AutoFilterInput']],
});
const DEFAULT_ARGS = Object.freeze({
  'Sheet.getCellXfIndices': [[0, null]],
  'Sheet.getCellsInRange': [[0, null]],
  'Sheet.getDependents': [[2, 1]],
  'Sheet.getPrecedents': [[2, 1]],
  'Sheet.getShowPhoneticCells': [[0, null]],
  'Sheet.getTypedValues': [[0, null]],
  'Sheet.getValues': [[0, null]],
  'Sheet.iterRows': [[0, null]],
  'Sheet.listCells': [[0, null]],
  'Sheet.move': [[0, null]],
  'Table.insertColumns': [
    [0, null],
    [1, null],
  ],
  'Table.insertRows': [
    [0, null],
    [1, null],
  ],
  'Workbook.addSheet': [[0, null]],
  'Workbook.recalcParallel': [[0, 0]],
});
const ARG_ROLES = Object.freeze({
  'Chart.addSeries': [[0, 'position']],
  'Chart.removeSeries': [
    [0, 'position'],
    [1, 'quantity'],
  ],
  'Chart.setType': [[0, 'position']],
  'Chart.setZOrder': [[0, 'position']],
  'Chart.updatePlot': [[0, 'position']],
  'Chart.updatePoint': [
    [0, 'position'],
    [1, 'quantity'],
    [2, 'position'],
  ],
  'Chart.updateSeries': [
    [0, 'position'],
    [1, 'quantity'],
  ],
  'Picture.setZOrder': [[0, 'position']],
  'Pivot.getDataField': [[0, 'position']],
  'Pivot.getDataFieldDefaultName': [[2, 'quantity']],
  'Pivot.getFieldItem': [[1, 'position']],
  'Pivot.getFilter': [[0, 'position']],
  'Pivot.removeFilter': [[0, 'position']],
  'Pivot.setDataField': [[0, 'position']],
  'PivotCache.getRecordValue': [[0, 'position']],
  'PivotCache.moveCalculatedItem': [[1, 'position']],
  'Shape.setZOrder': [[0, 'position']],
  'Sheet.addSlicer': [[1, 'position']],
  'Sheet.deleteColumns': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.deleteRows': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.evaluateCfFormula': [
    [0, 'coordinate'],
    [1, 'coordinate'],
    [2, 'coordinate'],
    [3, 'coordinate'],
  ],
  'Sheet.evaluateFormulaArray': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getArrayFormula': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getCellPhonetic': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getCellPhoneticProperties': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getCellShowPhonetic': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getCellXfIndex': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getColumnWidthPt': [[0, 'coordinate']],
  'Sheet.getComment': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getCommentVisible': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getDataTable': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getDependents': [
    [0, 'coordinate'],
    [1, 'coordinate'],
    [2, 'quantity'],
  ],
  'Sheet.getDisplayText': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getEffectiveStyle': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getFormulaR1C1': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getHyperlink': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getLambdaText': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getPrecedents': [
    [0, 'coordinate'],
    [1, 'coordinate'],
    [2, 'quantity'],
  ],
  'Sheet.getRowHeightPt': [[0, 'coordinate']],
  'Sheet.getSpill': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.getValue': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.insertColumns': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.insertRows': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.removeDataTable': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.removeHyperlinkAt': [[0, 'position']],
  'Sheet.resetCellXfIndex': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.resetColumnWidth': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.resetColumnXfIndex': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.resetRowHeight': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.resetRowXfIndex': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setBlank': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setBool': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setCellPhonetic': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setCellPhoneticProperties': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setCellPhoneticRuns': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setCellShowPhonetic': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setCellXfIndex': [
    [0, 'coordinate'],
    [1, 'coordinate'],
    [2, 'position'],
  ],
  'Sheet.setColumnBreak': [[0, 'coordinate']],
  'Sheet.setColumnCollapsed': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setColumnHidden': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setColumnOutline': [
    [0, 'coordinate'],
    [1, 'coordinate'],
    [2, 'quantity'],
  ],
  'Sheet.setColumnWidth': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setColumnXfIndex': [
    [0, 'coordinate'],
    [1, 'coordinate'],
    [2, 'position'],
  ],
  'Sheet.setCommentVisible': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setError': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setNumber': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setRangeXfIndex': [[1, 'position']],
  'Sheet.setRowAutoHeight': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setRowBreak': [[0, 'coordinate']],
  'Sheet.setRowCollapsed': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setRowHeight': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setRowHidden': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.setRowOutline': [
    [0, 'coordinate'],
    [1, 'coordinate'],
    [2, 'quantity'],
  ],
  'Sheet.setRowXfIndex': [
    [0, 'coordinate'],
    [1, 'coordinate'],
    [2, 'position'],
  ],
  'Sheet.setText': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Sheet.validateValue': [
    [0, 'coordinate'],
    [1, 'coordinate'],
  ],
  'Slicer.setZOrder': [[0, 'position']],
  'Table.renameColumn': [[0, 'position']],
  'Workbook.addPivotCache': [[0, 'quantity']],
  'Workbook.getBorder': [[0, 'position']],
  'Workbook.getCellStyleXf': [[0, 'position']],
  'Workbook.getCellXf': [[0, 'position']],
  'Workbook.getDxf': [[0, 'position']],
  'Workbook.getFill': [[0, 'position']],
  'Workbook.getFont': [[0, 'position']],
  'Workbook.getNumberFormat': [[0, 'quantity']],
  'Workbook.recalcParallel': [[0, 'quantity']],
  'Workbook.setBorder': [[0, 'position']],
  'Workbook.setCellStyleXf': [[0, 'position']],
  'Workbook.setCellXf': [[0, 'position']],
  'Workbook.setDxf': [[0, 'position']],
  'Workbook.setFill': [[0, 'position']],
  'Workbook.setFont': [[0, 'position']],
});
// </method-signatures>
// <param-names> generated by tools/codegen/gen_facade_tables.py; do not edit
const PARAM_NAMES = Object.freeze({
  'Chart.addSeries': ['plot', 'input'],
  'Chart.getPreview': ['overrides'],
  'Chart.removeSeries': ['plot', 'seriesId'],
  'Chart.setFilter': ['filter'],
  'Chart.setLegend': ['legend'],
  'Chart.setPoint': ['plot', 'seriesId', 'pointIndex', 'point'],
  'Chart.setSource': ['source'],
  'Chart.setTitle': ['title'],
  'Chart.setType': ['plot', 'type', 'options'],
  'Chart.setZOrder': ['zOrderIndex'],
  'Chart.update': ['patch'],
  'Chart.updateAxis': ['axis', 'patch'],
  'Chart.updateDisplay': ['patch'],
  'Chart.updateLegend': ['patch'],
  'Chart.updatePlot': ['plot', 'patch'],
  'Chart.updatePoint': ['plot', 'seriesId', 'pointIndex', 'patch'],
  'Chart.updateSeries': ['plot', 'seriesId', 'patch'],
  'Drawing.update': ['patch'],
  'Formulon.canonicalizeFormula': ['formula', 'profileId', 'options'],
  'Formulon.canonicalizeFunctionName': ['localizedName', 'profileId'],
  'Formulon.compareText': ['a', 'b', 'strength', 'profileId'],
  'Formulon.convertCivilToSerial': ['date', 'options'],
  'Formulon.convertSerialToCivil': ['serial', 'options'],
  'Formulon.evaluateFormula': ['formula'],
  'Formulon.evaluateTypedFormula': ['formula'],
  'Formulon.getCollationClasses': ['values', 'strength', 'profileId'],
  'Formulon.getFunctionMetadata': ['name'],
  'Formulon.getImageInfo': ['bytes'],
  'Formulon.getLocaleFacts': ['profileId'],
  'Formulon.localizeFormula': ['formula', 'profileId', 'options'],
  'Formulon.localizeFunctionName': ['canonicalName', 'profileId'],
  'Formulon.setLogMinLevel': ['level'],
  'Formulon.setLogSink': ['sink'],
  'Picture.setZOrder': ['zOrderIndex'],
  'Picture.update': ['patch'],
  'Pivot.addDataField': ['spec'],
  'Pivot.addFieldItem': ['fieldIndex', 'item', 'options'],
  'Pivot.addFilter': ['spec'],
  'Pivot.clearFieldItems': ['fieldIndex'],
  'Pivot.getDataField': ['dataFieldIndex'],
  'Pivot.getDataFieldDefaultName': ['fieldIndex', 'aggregation', 'excludeDataField'],
  'Pivot.getField': ['fieldIndex'],
  'Pivot.getFieldDateGroup': ['fieldIndex'],
  'Pivot.getFieldItem': ['fieldIndex', 'itemIndex'],
  'Pivot.getFieldNumericGroup': ['fieldIndex'],
  'Pivot.getFilter': ['filterIndex'],
  'Pivot.listFieldItems': ['fieldIndex'],
  'Pivot.removeFilter': ['filterIndex'],
  'Pivot.rename': ['name'],
  'Pivot.setAnchor': ['cell'],
  'Pivot.setColumnFieldOrder': ['fields'],
  'Pivot.setDataField': ['dataFieldIndex', 'spec'],
  'Pivot.setFieldAxis': ['fieldIndex', 'axis'],
  'Pivot.setFieldDateGroup': ['fieldIndex', 'group'],
  'Pivot.setFieldNumericGroup': ['fieldIndex', 'group'],
  'Pivot.setFieldSubtotalFunctions': ['fieldIndex', 'aggregations'],
  'Pivot.setPageFieldOrder': ['fields'],
  'Pivot.setReportLayout': ['layout'],
  'Pivot.setRowFieldOrder': ['fields'],
  'Pivot.showDetails': ['row', 'col'],
  'Pivot.updateField': ['fieldIndex', 'patch'],
  'Pivot.updateFieldItem': ['fieldIndex', 'itemIndex', 'patch'],
  'Pivot.updateGrandTotals': ['patch'],
  'Pivot.updateOptions': ['patch'],
  'Pivot.updateStyleOptions': ['patch'],
  'PivotCache.addCalculatedField': ['name', 'formula'],
  'PivotCache.addCalculatedItem': ['field', 'name', 'formula'],
  'PivotCache.addField': ['name'],
  'PivotCache.addFieldSharedItem': ['field', 'value'],
  'PivotCache.clearFieldSharedItems': ['field'],
  'PivotCache.getRecordValue': ['recordIndex', 'field'],
  'PivotCache.listFieldSharedItems': ['field'],
  'PivotCache.moveCalculatedItem': ['calculatedItemIndex', 'newIndex'],
  'PivotCache.removeCalculatedField': ['field'],
  'PivotCache.removeCalculatedItem': ['calculatedItemIndex'],
  'PivotCache.setCalculatedFieldFormula': ['field', 'formula'],
  'PivotCache.setCalculatedItemFormula': ['calculatedItemIndex', 'formula'],
  'PivotCache.setRecordValue': ['recordIndex', 'field', 'value'],
  'PivotCache.setWorksheetSource': ['source'],
  'PivotCache.updateOptions': ['patch'],
  'Shape.setZOrder': ['zOrderIndex'],
  'Shape.update': ['patch'],
  'Sheet.addCfRule': ['rule'],
  'Sheet.addChart': ['input'],
  'Sheet.addDataTable': ['spec'],
  'Sheet.addMerge': ['range'],
  'Sheet.addPicture': ['bytes', 'input'],
  'Sheet.addPivot': ['input'],
  'Sheet.addProtectedRange': ['input'],
  'Sheet.addShape': ['spec'],
  'Sheet.addSlicer': ['table', 'column', 'spec'],
  'Sheet.addSparklineGroup': ['spec'],
  'Sheet.addTable': ['input'],
  'Sheet.addThreadedComment': ['comment'],
  'Sheet.addValidation': ['validation'],
  'Sheet.chart': ['objectId'],
  'Sheet.convertColumnCharsToPt': ['chars', 'mode'],
  'Sheet.convertColumnPtToChars': ['pt', 'mode'],
  'Sheet.deleteCells': ['range', 'shift'],
  'Sheet.deleteColumns': ['first', 'last'],
  'Sheet.deleteRows': ['first', 'last'],
  'Sheet.drawing': ['objectId'],
  'Sheet.evaluateCfFormula': ['row', 'col', 'anchorRow', 'anchorCol', 'formula'],
  'Sheet.evaluateCfRange': ['range', 'options'],
  'Sheet.evaluateFormula': ['row', 'col', 'formula'],
  'Sheet.evaluateFormulaArray': ['row', 'col', 'formula'],
  'Sheet.evaluateTypedFormula': ['row', 'col', 'formula'],
  'Sheet.evaluateTypedFormulaArray': ['row', 'col', 'formula'],
  'Sheet.getArrayFormula': ['row', 'col'],
  'Sheet.getCellFormat': ['row', 'col'],
  'Sheet.getCellPhonetic': ['row', 'col'],
  'Sheet.getCellPhoneticProperties': ['row', 'col'],
  'Sheet.getCellPhoneticRuns': ['row', 'col'],
  'Sheet.getCellRectPt': ['range', 'mode'],
  'Sheet.getCellShowPhonetic': ['row', 'col'],
  'Sheet.getCellsInRange': ['range', 'options'],
  'Sheet.getCfRule': ['cfRule'],
  'Sheet.getColumnWidthPt': ['col', 'mode'],
  'Sheet.getComment': ['row', 'col'],
  'Sheet.getCommentVisible': ['row', 'col'],
  'Sheet.getDataTable': ['row', 'col'],
  'Sheet.getDate': ['row', 'col'],
  'Sheet.getDependents': ['row', 'col', 'depth'],
  'Sheet.getDisplayText': ['row', 'col'],
  'Sheet.getFormula': ['row', 'col'],
  'Sheet.getFormulaR1C1': ['row', 'col'],
  'Sheet.getHyperlink': ['row', 'col'],
  'Sheet.getInvalidCells': ['options'],
  'Sheet.getLambdaText': ['row', 'col'],
  'Sheet.getMergesInRange': ['range'],
  'Sheet.getPrecedents': ['row', 'col', 'depth'],
  'Sheet.getProtectedRange': ['protectedRange'],
  'Sheet.getRowHeightPt': ['row'],
  'Sheet.getShowPhoneticCells': ['range'],
  'Sheet.getSparklineGroup': ['sparklineGroup'],
  'Sheet.getSpill': ['row', 'col'],
  'Sheet.getThreadedComment': ['id'],
  'Sheet.getTypedValue': ['row', 'col'],
  'Sheet.getTypedValues': ['range'],
  'Sheet.getValidation': ['validation'],
  'Sheet.getValue': ['row', 'col'],
  'Sheet.getValues': ['range'],
  'Sheet.getWidthModel': ['mode'],
  'Sheet.insertCells': ['range', 'shift'],
  'Sheet.insertColumns': ['first', 'last'],
  'Sheet.insertHyperlinkAt': ['index', 'link'],
  'Sheet.insertRows': ['first', 'last'],
  'Sheet.iterRows': ['range'],
  'Sheet.listCells': ['range'],
  'Sheet.move': ['insertBefore'],
  'Sheet.picture': ['objectId'],
  'Sheet.pivot': ['pivot'],
  'Sheet.recalcViewport': ['viewport'],
  'Sheet.removeCfRule': ['cfRule'],
  'Sheet.removeDataTable': ['row', 'col'],
  'Sheet.removeDrawing': ['objectId'],
  'Sheet.removeHyperlinkAt': ['index'],
  'Sheet.removeMerge': ['range'],
  'Sheet.removePivot': ['pivot'],
  'Sheet.removeProtectedRange': ['protectedRange'],
  'Sheet.removeSparklineGroup': ['sparklineGroup'],
  'Sheet.removeTable': ['table'],
  'Sheet.removeThreadedComment': ['id'],
  'Sheet.removeValidation': ['validation'],
  'Sheet.rename': ['newName'],
  'Sheet.resetColumnWidth': ['first', 'last'],
  'Sheet.resetRowHeight': ['first', 'last'],
  'Sheet.restoreDrawing': ['bytes', 'opts'],
  'Sheet.setArrayFormula': ['arrayFormula'],
  'Sheet.setAutoFilter': ['filter'],
  'Sheet.setBlank': ['row', 'col'],
  'Sheet.setBool': ['row', 'col', 'value'],
  'Sheet.setCellPhonetic': ['row', 'col', 'phonetic'],
  'Sheet.setCellPhoneticProperties': ['row', 'col', 'properties'],
  'Sheet.setCellPhoneticRuns': ['row', 'col', 'runs'],
  'Sheet.setCellShowPhonetic': ['row', 'col', 'show'],
  'Sheet.setColumnBreak': ['col', 'input'],
  'Sheet.setColumnCollapsed': ['first', 'last', 'collapsed'],
  'Sheet.setColumnHidden': ['first', 'last', 'hidden'],
  'Sheet.setColumnOutline': ['first', 'last', 'level'],
  'Sheet.setColumnWidth': ['first', 'last', 'widthChars'],
  'Sheet.setComment': ['row', 'col', 'comment'],
  'Sheet.setCommentVisible': ['row', 'col', 'visible'],
  'Sheet.setError': ['row', 'col', 'error'],
  'Sheet.setFormula': ['row', 'col', 'formula'],
  'Sheet.setHyperlink': ['row', 'col', 'link'],
  'Sheet.setNumber': ['row', 'col', 'value'],
  'Sheet.setPrintArea': ['ranges'],
  'Sheet.setPrintTitles': ['titles'],
  'Sheet.setRowAutoHeight': ['first', 'last', 'heightPt'],
  'Sheet.setRowBreak': ['row', 'input'],
  'Sheet.setRowCollapsed': ['first', 'last', 'collapsed'],
  'Sheet.setRowHeight': ['first', 'last', 'heightPt'],
  'Sheet.setRowHidden': ['first', 'last', 'hidden'],
  'Sheet.setRowOutline': ['first', 'last', 'level'],
  'Sheet.setSelection': ['selection'],
  'Sheet.setSheetProtection': ['protection'],
  'Sheet.setTabColor': ['color'],
  'Sheet.setText': ['row', 'col', 'text'],
  'Sheet.setValue': ['row', 'col', 'value'],
  'Sheet.setValues': ['row', 'col', 'values'],
  'Sheet.setVisibility': ['visibility'],
  'Sheet.shape': ['objectId'],
  'Sheet.slicer': ['objectId'],
  'Sheet.table': ['table'],
  'Sheet.updateCellFormat': ['range', 'patch'],
  'Sheet.updateCfRule': ['cfRule', 'patch'],
  'Sheet.updateFormatDefaults': ['patch'],
  'Sheet.updateHeaderFooter': ['patch'],
  'Sheet.updatePageMargins': ['patch'],
  'Sheet.updatePageSetup': ['patch'],
  'Sheet.updatePrintOptions': ['patch'],
  'Sheet.updateProtectedRange': ['protectedRange', 'patch'],
  'Sheet.updateSheetProtection': ['patch'],
  'Sheet.updateSparklineGroup': ['sparklineGroup', 'patch'],
  'Sheet.updateThreadedComment': ['id', 'patch'],
  'Sheet.updateValidation': ['validation', 'patch'],
  'Sheet.updateView': ['patch'],
  'Sheet.validateValue': ['row', 'col', 'value'],
  'SheetOoxml.getCellXfIndex': ['row', 'col'],
  'SheetOoxml.getCellXfIndices': ['range'],
  'SheetOoxml.getColumnXfIndex': ['col'],
  'SheetOoxml.getEffectiveStyle': ['row', 'col'],
  'SheetOoxml.getRangeXfBases': ['range'],
  'SheetOoxml.getRangeXfBasesExcluding': ['range', 'excluded'],
  'SheetOoxml.getRowXfIndex': ['row'],
  'SheetOoxml.remapRangeXfs': ['range', 'mapping'],
  'SheetOoxml.resetCellXfIndex': ['row', 'col'],
  'SheetOoxml.resetColumnXfIndex': ['first', 'last'],
  'SheetOoxml.resetRowXfIndex': ['first', 'last'],
  'SheetOoxml.restoreRangeXfs': ['bytes'],
  'SheetOoxml.setAutoFilterXml': ['xml'],
  'SheetOoxml.setCellXfIndex': ['row', 'col', 'xfIndex'],
  'SheetOoxml.setColumnXfIndex': ['first', 'last', 'xfIndex'],
  'SheetOoxml.setHeaderFooterXml': ['xml'],
  'SheetOoxml.setPageMarginsXml': ['xml'],
  'SheetOoxml.setPageSetupXml': ['xml'],
  'SheetOoxml.setPrintOptionsXml': ['xml'],
  'SheetOoxml.setRangeXfIndex': ['range', 'xfIndex'],
  'SheetOoxml.setRowXfIndex': ['first', 'last', 'xfIndex'],
  'SheetOoxml.setSheetPrXml': ['xml'],
  'SheetOoxml.snapshotRangeXfs': ['range'],
  'Slicer.setZOrder': ['zOrderIndex'],
  'Slicer.update': ['patch'],
  'Table.insertColumns': ['first', 'last'],
  'Table.insertRows': ['first', 'last'],
  'Table.rename': ['name'],
  'Table.renameColumn': ['column', 'name'],
  'Table.setAutoFilter': ['filter'],
  'Table.update': ['patch'],
  'Table.updateStyleOptions': ['patch'],
  'Workbook.addPerson': ['person'],
  'Workbook.addPivotCache': ['requestedId'],
  'Workbook.addSheet': ['name', 'options'],
  'Workbook.copySheet': ['source', 'options'],
  'Workbook.createForProfile': ['profileId'],
  'Workbook.formatReference': ['range', 'options'],
  'Workbook.formatValue': ['value', 'numberFormat'],
  'Workbook.getCellStyle': ['name'],
  'Workbook.getDefinedName': ['name', 'scope'],
  'Workbook.getPerson': ['id'],
  'Workbook.getTableStyle': ['name'],
  'Workbook.getThemeScriptFont': ['key'],
  'Workbook.hasSheet': ['ref'],
  'Workbook.listThemeScriptFonts': ['scheme'],
  'Workbook.load': ['bytes'],
  'Workbook.parseReference': ['text', 'options'],
  'Workbook.pivotCache': ['cacheId'],
  'Workbook.recalcParallel': ['threadCount'],
  'Workbook.removeCellStyle': ['name'],
  'Workbook.removeDefinedName': ['name', 'scope'],
  'Workbook.removePerson': ['id'],
  'Workbook.removePivotCache': ['cacheId'],
  'Workbook.removeSheet': ['sheet'],
  'Workbook.removeTableStyle': ['name'],
  'Workbook.resolveColor': ['spec', 'context'],
  'Workbook.save': ['format'],
  'Workbook.saveWithDiagnostics': ['format'],
  'Workbook.setActiveSheet': ['sheet'],
  'Workbook.setCalcMode': ['mode'],
  'Workbook.setCalcOnSave': ['calcOnSave'],
  'Workbook.setCellStyle': ['record'],
  'Workbook.setDate1904': ['date1904'],
  'Workbook.setDefaultFont': ['record'],
  'Workbook.setDefinedName': ['definedName'],
  'Workbook.setFileName': ['name'],
  'Workbook.setFullPrecision': ['fullPrecision'],
  'Workbook.setIterativeProgress': ['callback'],
  'Workbook.setPinnedNow': ['time'],
  'Workbook.setProfileId': ['profileId'],
  'Workbook.setRefMode': ['mode'],
  'Workbook.setTableStyle': ['style'],
  'Workbook.setThemeColors': ['colors'],
  'Workbook.setThemeFonts': ['fonts'],
  'Workbook.setThemeScriptFont': ['font'],
  'Workbook.setWorkbookProtection': ['protection'],
  'Workbook.sheet': ['key'],
  'Workbook.table': ['name'],
  'Workbook.updateIterative': ['patch'],
  'Workbook.updateTableStyleDefaults': ['patch'],
  'WorkbookOoxml.addBorder': ['record'],
  'WorkbookOoxml.addCellStyleXf': ['record'],
  'WorkbookOoxml.addCellXf': ['record'],
  'WorkbookOoxml.addDxf': ['input'],
  'WorkbookOoxml.addFill': ['record'],
  'WorkbookOoxml.addFont': ['record'],
  'WorkbookOoxml.addNumberFormat': ['numberFormat'],
  'WorkbookOoxml.getBorder': ['borderIndex'],
  'WorkbookOoxml.getCellStyleXf': ['cellStyleXfIndex'],
  'WorkbookOoxml.getCellXf': ['xfIndex'],
  'WorkbookOoxml.getDxf': ['dxfIndex'],
  'WorkbookOoxml.getFill': ['fillIndex'],
  'WorkbookOoxml.getFont': ['fontIndex'],
  'WorkbookOoxml.getNumberFormat': ['numFmtId'],
  'WorkbookOoxml.hasPassthrough': ['path'],
  'WorkbookOoxml.setBorder': ['borderIndex', 'record'],
  'WorkbookOoxml.setCellStyleXf': ['cellStyleXfIndex', 'record'],
  'WorkbookOoxml.setCellXf': ['xfIndex', 'record'],
  'WorkbookOoxml.setDxf': ['dxfIndex', 'input'],
  'WorkbookOoxml.setFill': ['fillIndex', 'record'],
  'WorkbookOoxml.setFont': ['fontIndex', 'record'],
});
// </param-names>
// <boolean-args> generated by tools/codegen/gen_facade_tables.py; do not edit
const BOOLEAN_ARGS = Object.freeze({
  'Sheet.setBool': [2],
  'Sheet.setCellShowPhonetic': [2],
  'Sheet.setColumnCollapsed': [2],
  'Sheet.setColumnHidden': [2],
  'Sheet.setCommentVisible': [2],
  'Sheet.setRowCollapsed': [2],
  'Sheet.setRowHidden': [2],
  'Workbook.setCalcOnSave': [0],
  'Workbook.setDate1904': [0],
  'Workbook.setFullPrecision': [0],
});
// </boolean-args>
// <c-field-names> generated by tools/codegen/gen_facade_tables.py; do not edit
const C_FIELD_NAMES = Object.freeze({
  date_axis_ref: { js: 'dateAxisRange', py: 'date_axis_range' },
  has_bottom_in: { js: 'bottomInStated', py: 'bottom_in_stated' },
  has_footer_in: { js: 'footerInStated', py: 'footer_in_stated' },
  has_header_in: { js: 'headerInStated', py: 'header_in_stated' },
  has_left_in: { js: 'leftInStated', py: 'left_in_stated' },
  has_right_in: { js: 'rightInStated', py: 'right_in_stated' },
  has_scale_pct: { js: 'scalePctStated', py: 'scale_pct_stated' },
  has_top_in: { js: 'topInStated', py: 'top_in_stated' },
  items: { js: 'sparklines', py: 'sparklines' },
});
// </c-field-names>
// Arguments the declarations type `T | null` (or optional) where the raw method
// takes the C sentinel: `Class.method` -> [position in the public call,
// sentinel, whether an omitted argument also means absent].
const ABSENT_ARGS = Object.freeze({
  'Workbook.setFileName': [0, ''],
  'Sheet.setAutoFilterXml': [0, ''],
  'Sheet.setCellPhonetic': [2, ''],
  'Sheet.setCellPhoneticRuns': [2, []],
  'Sheet.setHeaderFooterXml': [0, ''],
  'Sheet.setPageMarginsXml': [0, ''],
  'Sheet.setPageSetupXml': [0, ''],
  'Sheet.setPrintArea': [0, []],
  'Sheet.setPrintOptionsXml': [0, ''],
  'Sheet.setSheetPrXml': [0, ''],
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
const LEGEND_FIELDS = Object.freeze({ position: 'number', overlay: 'boolean', hiddenEntries: 'array' });
const PATCH_FIELDS = Object.freeze({
  'Workbook.updateIterative': { ...bools('enabled'), ...numbers('maxIterations', 'maxChange') },
  'PivotCache.updateOptions': bools('saveData', 'refreshOnLoad'),
  'Sheet.updateValidation': {
    ...numbers('type', 'operator', 'errorStyle'),
    ...bools('allowBlank', 'showInputMessage', 'showErrorMessage', 'showInCellDropdown'),
  },
  'Sheet.updateCfRule': {
    ...numbers('type', 'priority', 'operator', 'rank', 'stdDev', 'timePeriod'),
    ...bools('stopIfTrue', 'percent', 'bottom', 'aboveAverage', 'equalAverage'),
    ...strings('text'),
    ...objects('colorScale', 'dataBar', 'iconSet'),
  },
  'Sheet.updateView': {
    ...numbers('zoomPct', 'freezeRows', 'freezeCols', 'splitXTwips', 'splitYTwips', 'viewMode'),
    ...bools(
      'showGridLines',
      'showRowColumnHeaders',
      'showZeros',
      'showFormulas',
      'showOutlineSymbols',
      'rightToLeft',
      'tabSelected',
    ),
  },
  'Sheet.updatePrintOptions': bools('gridLines', 'headings', 'horizontalCentered', 'verticalCentered'),
  'Sheet.updateHeaderFooter': bools('differentOddEven', 'differentFirst', 'scaleWithDoc', 'alignWithMargins'),
  'Sheet.updatePageSetup': { ...numbers('paperSize', 'scalePct', 'fitToWidth', 'fitToHeight'), ...bools('fitToPage') },
  'Sheet.updatePageMargins': numbers('leftIn', 'rightIn', 'topIn', 'bottomIn', 'headerIn', 'footerIn'),
  'Sheet.updateFormatDefaults': { ...numbers('baseColumnWidthChars'), ...bools('customHeight', 'zeroHeight') },
  'Sheet.updateSheetProtection': bools(
    'allowEditObjects',
    'allowEditScenarios',
    'allowFormatCells',
    'allowFormatColumns',
    'allowFormatRows',
    'allowInsertColumns',
    'allowInsertRows',
    'allowInsertHyperlinks',
    'allowDeleteColumns',
    'allowDeleteRows',
    'allowSelectLockedCells',
    'allowSelectUnlockedCells',
    'allowSort',
    'allowAutoFilter',
    'allowPivotTables',
  ),
  'Pivot.updateField': {
    ...bools('defaultSubtotal', 'insertBlankRow', 'fillDownLabels', 'showAll', 'subtotalTop', 'showDetail'),
    sortOrder: 'number',
  },
  'Pivot.updateFieldItem': bools('visible', 'showDetail'),
  'Pivot.updateStyleOptions': bools(
    'showRowHeaders',
    'showColumnHeaders',
    'showRowStripes',
    'showColumnStripes',
    'showLastColumn',
  ),
  'Pivot.updateOptions': {
    ...bools(
      'showError',
      'showMissing',
      'showHeaders',
      'showDrill',
      'showDataTips',
      'gridDropZones',
      'showEmptyRow',
      'showEmptyColumn',
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
  'Chart.setLegend': LEGEND_FIELDS,
  'Chart.updateLegend': LEGEND_FIELDS,
  'Chart.updateAxis': { ...bools('minorGridlines', 'sourceLinked'), position: 'number' },
  // A nested key is named by its path; `labels` itself takes `null`, its flags do not.
  'Chart.updateSeries': bools('labels.sourceLinked', 'labels.deleted'),
  // `range` takes a `CellRect`, a `SheetRect` or an A1 string, checked by `tableRange`.
  'Table.update': bools('headerRow', 'totalsRow'),
  'Table.updateStyleOptions': bools('showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes'),
  'Shape.update': {
    ...strings('presetGeometry', 'text'),
    ...bools('textBox'),
    rotationDeg: 'number',
  },
  'Slicer.update': {
    ...strings('caption'),
    ...numbers('columnCount', 'rowHeightPt', 'sortOrder', 'crossFilter'),
  },
  'Sheet.updateSparklineGroup': {
    ...numbers('type', 'minAxisType', 'maxAxisType', 'lineWidthPt', 'emptyCellsAs'),
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
    ...strings('name', 'algorithmName', 'hashValue', 'saltValue', 'securityDescriptor'),
    spinCount: 'number',
    ranges: 'array',
  },
});

// The keys of a drawing patch every kind takes, and the keys of its `anchor`.
const DRAWING_PATCH_KEYS = Object.freeze({
  name: 'string',
  description: 'string',
  hidden: 'boolean',
  anchor: 'object',
});
const DRAWING_ANCHOR_KEYS = Object.freeze([
  'kind',
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
]);

// The keys of a conditional-format patch each rule kind uses, beyond `type` and `id`.
const CF_PATCH_BASE = Object.freeze(['ranges', 'priority', 'stopIfTrue']);
const CF_PATCH_FORMAT = Object.freeze([...CF_PATCH_BASE, 'dxfIndex']);
const CF_PATCH_KEYS = Object.freeze({
  [CfRuleType.Expression]: [...CF_PATCH_FORMAT, 'formula1'],
  [CfRuleType.CellIs]: [...CF_PATCH_FORMAT, 'operator', 'formula1', 'formula2'],
  [CfRuleType.ColorScale]: [...CF_PATCH_BASE, 'colorScale'],
  [CfRuleType.DataBar]: [...CF_PATCH_BASE, 'dataBar'],
  [CfRuleType.IconSet]: [...CF_PATCH_BASE, 'iconSet'],
  [CfRuleType.Top10]: [...CF_PATCH_FORMAT, 'rank', 'percent', 'bottom'],
  [CfRuleType.AboveAverage]: [...CF_PATCH_FORMAT, 'aboveAverage', 'equalAverage', 'stdDev'],
  [CfRuleType.ContainsText]: [...CF_PATCH_FORMAT, 'text', 'formula1'],
  [CfRuleType.NotContainsText]: [...CF_PATCH_FORMAT, 'text', 'formula1'],
  [CfRuleType.BeginsWith]: [...CF_PATCH_FORMAT, 'text', 'formula1'],
  [CfRuleType.EndsWith]: [...CF_PATCH_FORMAT, 'text', 'formula1'],
  [CfRuleType.TimePeriod]: [...CF_PATCH_FORMAT, 'timePeriod', 'formula1'],
  [CfRuleType.ContainsBlanks]: [...CF_PATCH_FORMAT, 'formula1'],
  [CfRuleType.NotContainsBlanks]: [...CF_PATCH_FORMAT, 'formula1'],
  [CfRuleType.ContainsErrors]: [...CF_PATCH_FORMAT, 'formula1'],
  [CfRuleType.NotContainsErrors]: [...CF_PATCH_FORMAT, 'formula1'],
  [CfRuleType.DuplicateValues]: [...CF_PATCH_FORMAT, 'formula1'],
  [CfRuleType.UniqueValues]: [...CF_PATCH_FORMAT, 'formula1'],
});

// Where a sparkline whose source text does not parse keeps that text, so the group it is read in writes it back.
const STORED_SOURCE = Symbol('formulon.sparklineSource');

/** The public `ThreadedComment` of a raw record: its anchor as `cell`, its stamp as a civil time. */
function threadedCommentRead({ row, col, created, ...comment }) {
  return { ...comment, cell: { row, col }, created: stampCivil(created) };
}

// A1 text of the one cell `{ row, col }`.
function cellText(raw, { row, col }) {
  return raw.formatReference({ rect: { firstRow: row, firstCol: col, lastRow: row, lastCol: col } });
}

/**
 * The raw record of a sparkline group (or of the keys of a patch) given as a public `SparklineGroupInput`:
 * each sparkline's `source` and `location`, and `dateAxisRange`, become the reference text the raw layer takes.
 * `sheet` resolves an unqualified source.
 */
function encodeSparklineGroup(raw, sheet, group, where) {
  const { sparklines, dateAxisRange, ...rest } = group;
  const out = { ...rest };
  if (sparklines !== undefined) {
    out.sparklines = requireArray(sparklines, `${where} sparklines`).map((item, i) => {
      const label = `${where} sparklines[${i}]`;
      requireObject(item, label);
      for (const key of Object.keys(item)) {
        if (item[key] !== undefined && !SPARKLINE_KEYS.includes(key)) {
          throw new TypeError(`${where} has no field sparklines[${i}].${key}`);
        }
      }
      const formula = present(item.source)
        ? sheetRectText(raw, item.source, sheet, `${label}.source`)
        : (item[STORED_SOURCE] ?? '');
      return { formula, location: cellText(raw, addressOf(raw, item.location, `${label}.location`)) };
    });
  }
  if (dateAxisRange !== undefined) {
    out.dateAxisRef =
      dateAxisRange === null ? null : sheetRectText(raw, dateAxisRange, sheet, `${where} dateAxisRange`);
  }
  return out;
}

/** The public `SparklineGroup` of a raw group record: sources and the date range as `SheetRect`, locations as cells. */
function decodeSparklineGroup(raw, group) {
  const { dateAxisRef, sparklines, ...rest } = group;
  const withoutEquals = (text) => (typeof text === 'string' && text.startsWith('=') ? text.slice(1) : text);
  return {
    ...rest,
    dateAxisRange: parseSheetRect(raw, withoutEquals(dateAxisRef)),
    sparklines: sparklines.map(({ formula, location }) => {
      const text = withoutEquals(formula);
      const source = parseSheetRect(raw, text);
      const cell = parseReferenceOrNull(raw, location);
      const sparkline = {
        source,
        location: cell === null ? { row: 0, col: 0 } : { row: cell.rect.firstRow, col: cell.rect.firstCol },
      };
      if (source === null && text !== '') Object.defineProperty(sparkline, STORED_SOURCE, { value: formula });
      return sparkline;
    }),
  };
}

/**
 * Splits a drawing patch of `where` into the keys every kind takes (`DRAWING_PATCH_KEYS`, checked here) and the
 * kind's own keys, `undefined` keys and the read-only keys of `Drawing.update` dropped. `description` takes
 * `null`, no other common key or anchor key does.
 */
function splitDrawingPatch(where, patch) {
  const label = argLabel(where, 0);
  requireObject(patch, label);
  const common = {};
  const rest = {};
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined) continue;
    if (Object.hasOwn(DRAWING_PATCH_KEYS, key)) {
      const type = DRAWING_PATCH_KEYS[key];
      if (value === null && key === 'description') {
        common[key] = value;
        continue;
      }
      const mismatch = type === 'object' ? !isObject(value) : typeof value !== type;
      if (mismatch) throw new TypeError(`${label}.${key} must be ${typeNoun(type)}`);
      common[key] = value;
    } else if (!READ_ONLY_KEYS['Drawing.update'].includes(key)) {
      rest[key] = value;
    }
  }
  if (common.anchor !== undefined) {
    for (const [key, value] of Object.entries(common.anchor)) {
      if (value === undefined) continue;
      if (!DRAWING_ANCHOR_KEYS.includes(key)) throw new TypeError(`${label}.anchor has no field ${key}`);
      if (typeof value !== 'number') throw new TypeError(`${label}.anchor.${key} must be a number`);
    }
  }
  return [common, rest];
}

/** The raw drawing patch of the common keys `common`: `description: null` is the empty description. */
function rawDrawingPatch(common) {
  return common.description === null ? { ...common, description: '' } : common;
}

/** The axis id a `ChartAxisRef` of chart `id` names: a number is the id, a `{ kind, plot? }` key is looked up. */
function chartAxisId(raw, sheet, id, axis, label) {
  if (typeof axis === 'number') {
    const known = raw.getChart(sheet, id);
    if (known !== null && !known.axes.some((a) => a.id === axis)) {
      const kind = enumName(ChartAxisKind, axis) ?? 'Value';
      throw new FormulonError(
        StatusCode.NotFound,
        `${label} has no axis with id ${axis}. Pass { kind: ChartAxisKind.${kind} } to address an axis by kind.`,
        `axis_id=${axis}`,
      );
    }
    return axis;
  }
  if (!isObject(axis)) throw new TypeError(`${label} must be a number or an object`);
  for (const key of Object.keys(axis)) {
    if (axis[key] !== undefined && key !== 'kind' && key !== 'plot') {
      throw new TypeError(`${label} has no field ${key}`);
    }
  }
  const kind = requireInteger(axis.kind, `${label}.kind`);
  const plot = axis.plot === undefined ? 0 : requireIndex(axis.plot, `${label}.plot`, 'position');
  const properties = raw.getChart(sheet, id);
  if (properties === null) {
    throw new FormulonError(StatusCode.NotFound, 'the drawing object has no such part', `object_id=${id}`);
  }
  const group = properties.plots[plot];
  if (group === undefined) {
    throw new FormulonError(
      StatusCode.IndexOutOfRange,
      `${label}.plot ${plot} ${pastCount(properties.plots.length, 'item')}`,
      `plot=${plot}`,
    );
  }
  const found = group.axisIds.find((axisId) => properties.axes.find((a) => a.id === axisId)?.kind === kind);
  if (found === undefined) {
    throw new FormulonError(
      StatusCode.NotFound,
      `plot ${plot} has no axis of kind ${kind}`,
      `plot=${plot};kind=${kind}`,
    );
  }
  return found;
}

const COLOR_KEYS = Object.freeze(['argb', 'indexed', 'kind', 'theme', 'tint']);
const SPARKLINE_KEYS = Object.freeze(['source', 'location']);

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
        throw new TypeError(`${argLabel(where, 1)}.${path} has no field ${key}`);
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

/** The `[first, last]` of an insert call whose arguments may be omitted: both omitted appends one at `count`. */
function insertSpan(where, first, last, count) {
  const given = (value) => value !== undefined && value !== null;
  if (!given(first)) {
    if (given(last)) {
      throw new FormulonError(
        StatusCode.InvalidArgument,
        `${argLabel(where, 1)} needs ${argumentPath(where, 0)}`,
        `last=${last}`,
      );
    }
    return [count, count];
  }
  const from = requireIndex(first, argLabel(where, 0), 'coordinate');
  return [from, given(last) ? requireIndex(last, argLabel(where, 1), 'coordinate') : from];
}

/**
 * A `TablePatch.range` as the table's new `CellRect`: an A1 string, unqualified or naming the table's sheet, a
 * `CellRect`, or a `SheetRect` on that sheet. Another sheet is `InvalidArgument`.
 */
function tableRange(raw, sheet, range, where) {
  let onSheet;
  let rect;
  if (typeof range === 'string') {
    const ref = raw.parseReference(range);
    onSheet = ref.sheetName === null ? sheet : raw.getSheetIndex(ref.sheetName);
    rect = ref.rect;
  } else if (isObject(range) && range.rect === undefined) {
    onSheet = sheet;
    rect = range;
  } else if (isObject(range)) {
    onSheet = requireIndex(range.sheet, `${where} range.sheet`, 'sheet');
    rect = requireObject(range.rect, `${where} range.rect`);
  } else {
    throw new TypeError(`${where} range must be a CellRect, a SheetRect or an A1 string`);
  }
  if (onSheet !== sheet) {
    throw new FormulonError(
      StatusCode.InvalidArgument,
      `${argLabel(where, 0, 'range')} must be on the table's sheet`,
      `sheet_index=${sheet}`,
    );
  }
  return rect;
}

// How a string key of each `*Ref` kind is matched: the element count, the key
// of element `i`, the wording of the errors, and the handle class that also
// names an element.
// A `keyed` kind reaches the raw call as its key: a string passes through for the engine to match (an unknown
// one is `NotFound`), a position is read as the key stored there.
const REF_KINDS = Object.freeze({
  table: {
    noun: 'table name',
    handle: 'Table',
    context: 'table_name',
    count: (raw, sheet) => raw.getTableCount(sheet),
    key: (raw, sheet, i) => raw.getTable(sheet, i).name,
  },
  pivot: {
    noun: 'pivot name',
    handle: 'Pivot',
    context: 'pivot_name',
    count: (raw, sheet) => raw.getPivotCount(sheet),
    key: (raw, sheet, i) => raw.getPivotName(sheet, i),
  },
  cfRule: {
    keyed: true,
    noun: 'conditional-format rule id',
    context: 'cf_rule_id',
    count: (raw, sheet) => raw.getCfRuleCount(sheet),
    key: (raw, sheet, i) => raw.getCfRule(sheet, i).id,
  },
  protectedRange: {
    keyed: true,
    noun: 'protected range name',
    context: 'protected_range_name',
    count: (raw, sheet) => raw.getProtectedRangeCount(sheet),
    key: (raw, sheet, i) => raw.getProtectedRange(sheet, i).name,
  },
  validation: {
    keyed: true,
    noun: 'validation id',
    context: 'validation_id',
    count: (raw, sheet) => raw.getValidationCount(sheet),
    key: (raw, sheet, i) => raw.getValidation(sheet, i).id,
  },
  sparklineGroup: {
    keyed: true,
    noun: 'sparkline group id',
    context: 'sparkline_group_id',
    count: (raw, sheet) => raw.getSparklineGroupCount(sheet),
    key: (raw, sheet, i) => raw.getSparklineGroup(sheet, i).id,
  },
});

const GRID_ROWS = 1048576;
const GRID_COLS = 16384;
// 2^22 cells: the largest range `getValues` reads.
const MAX_VALUES_CELLS = 4194304;
const BLANK = Object.freeze({ kind: ValueKind.Blank });
const DEFAULT_LEGEND = Object.freeze({ position: ChartLegendPosition.Right, overlay: false, hiddenEntries: [] });
const GEOMETRY_MODES = new Set(Object.values(GeometryMode));

function isObject(value) {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

// ---- Argument messages: `<Class>.<method> <path> <constraint>` ----

/** `<where> <path> <constraint>`; an empty `path` gives `<where> <constraint>`. */
function argumentMessage(where, path, constraint) {
  return path === '' ? `${where} ${constraint}` : `${where} ${path} ${constraint}`;
}

// The classes whose methods share another class's table rows: a lookup that misses falls back along these.
const KEY_FALLBACKS = Object.freeze({
  Workbook: ['WorkbookOoxml'],
  Sheet: ['SheetOoxml'],
  WorkbookOoxml: ['Workbook'],
  SheetOoxml: ['Sheet'],
  Picture: ['Drawing'],
  Chart: ['Drawing'],
  Shape: ['Drawing'],
  Slicer: ['Drawing'],
});

/** The row of `table` for `where` (`Class.method`), or the row its fallback class shares. */
function tableEntry(table, where) {
  if (Object.hasOwn(table, where)) return table[where];
  const dot = where.indexOf('.');
  for (const owner of KEY_FALLBACKS[where.slice(0, dot)] ?? []) {
    const key = `${owner}${where.slice(dot)}`;
    if (Object.hasOwn(table, key)) return table[key];
  }
  return undefined;
}

/**
 * The path of argument `position` of `where`, as the declaration names it, followed by `fields`: a string is a
 * `.field`, a number an `[index]`.
 */
function argumentPath(where, position, ...fields) {
  let path = tableEntry(PARAM_NAMES, where)?.[position] ?? 'value';
  for (const field of fields) path += typeof field === 'number' ? `[${field}]` : `.${field}`;
  return path;
}

/** `<where> <path>` of argument `position`: the label the `require*` checks append a constraint to. */
function argLabel(where, position, ...fields) {
  return `${where} ${argumentPath(where, position, ...fields)}`;
}

/** `is out of range (<count> <noun>)`, the noun plural unless `count` is 1: the wording every binding uses. */
function pastCount(count, noun) {
  return `is out of range (${count} ${count === 1 ? noun : `${noun}s`})`;
}

/** The label of `fields` of the record (patch or input) argument of `where`; `source` is the row of its keys. */
function recordLabel(where, fields = [], source = where) {
  const position = (tableEntry(PATCH_KEYS, source) ?? tableEntry(INPUT_KEYS, source))?.[0] ?? 0;
  return argLabel(where, position, ...fields);
}

const CALL_TOKEN = /^[A-Z][A-Za-z0-9]*\.[A-Za-z0-9]+(?= )/;
const C_PATH = /^[a-z][a-z0-9_]*(?:\[\d+\])*(?:\.[a-z][a-z0-9_]*(?:\[\d+\])*)*$/;
const camelCase = (snake) => snake.replace(/_([a-z0-9])/g, (_, c) => c.toUpperCase());

/**
 * Whether `cPath` leads with a name `where` takes: one of its parameters (`sheet_index` names `sheet`), a key of
 * its input or patch record, or a C field a record publishes under another name. A message that leads with any
 * other word is prose and keeps its wording.
 */
function namesArgument(where, cPath) {
  const root = cPath.split(/[.[]/)[0];
  if (Object.hasOwn(C_FIELD_NAMES, root)) return true;
  const params = tableEntry(PARAM_NAMES, where) ?? [];
  if ([root, root.replace(/_index$/, '')].some((name) => params.includes(camelCase(name)))) return true;
  const keys = (tableEntry(PATCH_KEYS, where) ?? tableEntry(INPUT_KEYS, where))?.[1] ?? [];
  return keys.includes(camelCase(root));
}

/** The public path of an engine body's C path: renamed fields, camelCase, rooted at the parameter it names. */
function publicPath(where, cPath) {
  const segments = cPath.split('.').map((segment) => {
    const [, name, indices] = /^([a-z0-9_]+)((?:\[\d+\])*)$/.exec(segment);
    return (C_FIELD_NAMES[name]?.js ?? camelCase(name)) + indices;
  });
  const params = tableEntry(PARAM_NAMES, where) ?? [];
  const root = cPath.split(/[.[]/)[0];
  const named = [camelCase(root), camelCase(root.replace(/_index$/, ''))].find((name) => params.includes(name));
  if (named !== undefined) return [named + segments[0].replace(/^[^[]*/, ''), ...segments.slice(1)].join('.');
  const position = (tableEntry(PATCH_KEYS, where) ?? tableEntry(INPUT_KEYS, where))?.[0];
  return position === undefined || params[position] === undefined
    ? segments.join('.')
    : [params[position], ...segments].join('.');
}

/** Values a user callback threw: they reach the caller as they were thrown. */
const calledBackErrors = new WeakSet();

/** `fn` as a function that records what it throws, so the error labels leave it alone. */
function recordThrown(fn) {
  if (typeof fn !== 'function') return fn;
  return function (...args) {
    try {
      return fn.apply(this, args);
    } catch (e) {
      if (isObject(e)) calledBackErrors.add(e);
      throw e;
    }
  };
}

/** `error` as the same class, code and context with `message`. */
function withMessage(error, message) {
  if (error.name === 'FormulonError') return new FormulonError(error.code, message, error.context, error.parseError);
  return new error.constructor(message);
}

/**
 * `error` as the call `where` (`Class.method`) words it: a binding-layer message that leads with another call is
 * rewritten to lead with `where`, and an engine `FormulonError` that leads with no call is labelled with `where`;
 * when its first token is a C path naming one of the call's arguments (`<c_path> <rest>`), that token becomes
 * the public path. Anything else passes as it is.
 */
function relabelArgumentError(error, where) {
  if (calledBackErrors.has(error)) return error;
  if (typeof error?.message !== 'string') return error;
  const { message } = error;
  const isFormulon = error.name === 'FormulonError';
  if (!isFormulon && !(error instanceof TypeError) && !(error instanceof RangeError)) return error;
  const call = CALL_TOKEN.exec(message);
  if (call !== null) return call[0] === where ? error : withMessage(error, where + message.slice(call[0].length));
  if (!isFormulon) return error;
  const space = message.indexOf(' ');
  const head = space < 0 ? '' : message.slice(0, space);
  if (!C_PATH.test(head) || !namesArgument(where, head)) return withMessage(error, `${where} ${message}`);
  return withMessage(error, `${where} ${publicPath(where, head)}${message.slice(space)}`);
}

/**
 * Wraps every method `proto` defines itself (`accept(name)` choosing, when given) so that a thrown argument error
 * is worded as the call the consumer made: `owner` is the class name, or a function of the receiver giving it.
 * Installed outermost, after every guard.
 */
function installErrorLabels(proto, owner, accept = () => true) {
  const classOf = typeof owner === 'function' ? owner : () => owner;
  for (const name of Object.getOwnPropertyNames(proto)) {
    if (name === 'constructor' || name === 'dispose' || !accept(name)) continue;
    const desc = Object.getOwnPropertyDescriptor(proto, name);
    if (typeof desc.value !== 'function') continue;
    const inner = desc.value;
    defineMethod(
      proto,
      name,
      namedFunction(name, function (args) {
        try {
          return inner.apply(this, args);
        } catch (e) {
          throw relabelArgumentError(e, `${classOf(this)}.${name}`);
        }
      }),
    );
  }
}

const typeNoun = (type) => (type === 'array' || type === 'object' ? `an ${type}` : `a ${type}`);

function requireObject(value, label) {
  if (!isObject(value)) throw new TypeError(`${label} must be an object`);
  return value;
}

function requireString(value, label) {
  if (typeof value !== 'string') throw new TypeError(`${label} must be a string`);
  return value;
}

function requireBoolean(value, label) {
  if (typeof value !== 'boolean') throw new TypeError(`${label} must be a boolean`);
  return value;
}

function requireArray(value, label) {
  if (!Array.isArray(value)) throw new TypeError(`${label} must be an array`);
  return value;
}

function requireNumber(value, label) {
  if (typeof value !== 'number') throw new TypeError(`${label} must be a number`);
  return value;
}

/** A signed 32-bit integer: a non-number is a `TypeError`, a fraction or a value outside int32 a `RangeError`. */
function requireInteger(value, label) {
  requireNumber(value, label);
  if (!Number.isInteger(value)) throw new RangeError(`${label} must be an integer`);
  if (value > INT32_MAX || value < -INT32_MAX - 1) throw new RangeError(`${label} is too large`);
  return value;
}

// An unsigned integer by role. A negative value is `IndexOutOfRange` for a `position`,
// `SheetIndexOutOfRange` for a `sheet`, and `InvalidArgument` for a `coordinate` (also at or above 2^31) and a
// `quantity`; a fraction, NaN or an infinity, and a value beyond 32 bits, is a `RangeError`.
function requireIndex(value, label, role) {
  requireNumber(value, label);
  if (!Number.isInteger(value)) throw new RangeError(`${label} must be an integer`);
  if (value > 0xffffffff) throw new RangeError(`${label} is too large`);
  if (value < 0) {
    switch (role) {
      case 'sheet':
        throw new FormulonError(
          StatusCode.SheetIndexOutOfRange,
          `${label} ${value} is out of range`,
          `sheet_index=${value}`,
        );
      case 'position':
        throw new FormulonError(StatusCode.IndexOutOfRange, `${label} ${value} is out of range`, `index=${value}`);
      case 'coordinate':
        throw new FormulonError(
          StatusCode.InvalidArgument,
          `${label} ${value} is outside the grid`,
          `coordinate=${value}`,
        );
      default:
        throw new FormulonError(StatusCode.InvalidArgument, `${label} must not be negative`, `value=${value}`);
    }
  }
  if (role === 'coordinate' && value >= 0x80000000) {
    throw new FormulonError(StatusCode.InvalidArgument, `${label} ${value} is outside the grid`, `coordinate=${value}`);
  }
  return value;
}

/** A byte-input argument as a `Uint8Array`: an `ArrayBuffer`, a `Uint8Array` or any `ArrayBufferView`. */
function requireBytes(value, where, param) {
  if (value instanceof Uint8Array) return value;
  if (value instanceof ArrayBuffer) return new Uint8Array(value);
  if (ArrayBuffer.isView(value)) return new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
  throw new TypeError(argumentMessage(where, param, 'must be a Uint8Array'));
}

/** A patch field: `undefined` keeps `current`, `null` is refused, anything else must pass `check`. */
function patchField(patch, field, current, check, where) {
  const value = patch[field];
  return value === undefined ? current : check(value, recordLabel(where, [field]));
}

function invalidAddress(a1, label) {
  return new FormulonError(StatusCode.InvalidArgument, `${label} must name one cell`, `address=${a1}`);
}

// The only A1 reader: whatever `parseReference` accepts as one unqualified
// cell is an address; everything else is rejected. `label` is the `<call> <path>` of the argument.
function cellAddress(raw, a1, label) {
  const ref = raw.parseReference(a1);
  if (ref.sheetName !== null && label.startsWith('Sheet.')) {
    const local = a1.slice(a1.lastIndexOf('!') + 1);
    const method = label.slice('Sheet.'.length, label.indexOf(' '));
    throw new FormulonError(
      StatusCode.InvalidArgument,
      `${label} must not name a sheet; use wb.sheet(${JSON.stringify(ref.sheetName)}).${method}(${JSON.stringify(local)})`,
      `address=${a1}`,
    );
  }
  if (ref.sheetName !== null) throw invalidAddress(a1, label);
  const { rect } = ref;
  if (rect.firstRow !== rect.lastRow || rect.firstCol !== rect.lastCol) throw invalidAddress(a1, label);
  return rect;
}

// A range argument: any unqualified A1 reference, a single cell being 1x1.
function rangeAddress(raw, a1, label) {
  const ref = raw.parseReference(a1);
  if (ref.sheetName !== null) {
    throw new FormulonError(StatusCode.InvalidArgument, `${label} must not name a sheet`, `address=${a1}`);
  }
  return ref.rect;
}

// The axis a span argument (`SPAN_ARGS`) runs along: the rectangle it must
// cover whole, and the coordinates that bound it.
const SPAN_AXES = Object.freeze({
  col: {
    noun: 'column',
    whole: (rect) => rect.firstRow === 0 && rect.lastRow === GRID_ROWS - 1,
    first: 'firstCol',
    last: 'lastCol',
  },
  row: {
    noun: 'row',
    whole: (rect) => rect.firstCol === 0 && rect.lastCol === GRID_COLS - 1,
    first: 'firstRow',
    last: 'lastRow',
  },
});

// A span argument: an unqualified A1 reference over whole columns (`B`, `B:D`)
// or whole rows (`3`, `3:5`); the whole sheet is both. `refusal` words what it must name.
function spanAddress(raw, text, axis, label, refusal = `must name whole ${SPAN_AXES[axis].noun}s`) {
  const spec = SPAN_AXES[axis];
  const ref = raw.parseReference(text.includes(':') ? text : `${text}:${text}`);
  if (ref.sheetName !== null) {
    throw new FormulonError(StatusCode.InvalidArgument, `${label} must not name a sheet`, `address=${text}`);
  }
  if (!spec.whole(ref.rect)) {
    throw new FormulonError(StatusCode.InvalidArgument, `${label} ${refusal}`, `address=${text}`);
  }
  return [ref.rect[spec.first], ref.rect[spec.last]];
}

// The one row or column an A1 string of a `single` span argument names.
function singleAxis(raw, where, text, axis) {
  const label = `${where} a1`;
  const refusal = `must name one ${SPAN_AXES[axis].noun}`;
  const [first, last] = spanAddress(raw, text, axis, label, refusal);
  if (first !== last) throw new FormulonError(StatusCode.InvalidArgument, `${label} ${refusal}`, `address=${text}`);
  return first;
}

/**
 * The numeric form of a call whose span arguments (`SPAN_ARGS`) may be one A1
 * span string: `(first, last)`, or the one row or column of a `single` span.
 */
function expandSpans(raw, where, spans, args) {
  let out = args;
  for (const [at, axis, form] of spans) {
    if (typeof out[at] !== 'string') continue;
    if (form === 'single') {
      out = [...out.slice(0, at), singleAxis(raw, where, out[at], axis), ...out.slice(at + 1)];
      continue;
    }
    const [first, last] = spanAddress(raw, out[at], axis, `${where} a1`);
    out = [...out.slice(0, at), first, last, ...out.slice(at + 1)];
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
        throw new TypeError(`${argLabel(where, typeof args[i] === 'string' ? i : j)} must be a number`);
      }
    }
    return args;
  }
  const out = [];
  let from = 0;
  pairs.forEach(([i], k) => {
    const at = i - k;
    out.push(...args.slice(from, at));
    if (typeof args[at] !== 'string') throw new TypeError(`${argLabel(where, i)} must be a string`);
    const cell = cellAddress(raw, args[at], `${where} a1`);
    out.push(cell.firstRow, cell.firstCol);
    from = at + 1;
  });
  out.push(...args.slice(from));
  return out;
}

// `CellRect` arguments (`RECT_ARGS`) given as A1 ranges, alone or in an array.
function expandRects(raw, where, positions, args) {
  let out = args;
  for (const at of positions) {
    const value = args[at];
    const label = argLabel(where, at);
    const rect = (item) =>
      typeof item === 'string'
        ? rangeAddress(raw, item, label)
        : isObject(item)
          ? checkRangeFields('CellRect', item, label)
          : item;
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
  return [ref.sheetName, ref.rect];
}

function unqualifiedRect(raw, value, label) {
  const [sheetName, rect] = parsedRange(raw, value, label, 'CellRect');
  if (sheetName !== null) {
    throw new FormulonError(StatusCode.InvalidArgument, `${label} must not name a sheet`, `address=${value}`);
  }
  return rect;
}

const RANGE_KINDS = new Set(['CellRect', 'SheetRect', 'CellAddress']);

// The value of a range-typed field or argument in raw form; `sheet` resolves an unqualified `SheetRect`.
function normalizeAddress(raw, kind, value, sheet, label) {
  if (value === null || value === undefined) return value;
  if (kind.endsWith('[]')) {
    const element = kind.slice(0, -2);
    // Only a list of rectangles also takes one A1 string (`'A1:B2'`); a list of cells is always an array.
    const range = element === 'CellRect';
    if (range && typeof value === 'string') return [normalizeAddress(raw, element, value, sheet, `${label}[0]`)];
    if (!Array.isArray(value)) {
      throw new TypeError(`${label} must be an array${range ? ' or one A1 string' : ''}`);
    }
    return value.map((item, i) => normalizeAddress(raw, element, item, sheet, `${label}[${i}]`));
  }
  if (!RANGE_KINDS.has(kind)) return normalizeRecordRanges(raw, value, kind, sheet, label);
  if (isObject(value)) return checkRangeFields(kind, value, label);
  if (kind === 'CellRect') return unqualifiedRect(raw, value, label);
  if (kind === 'SheetRect') {
    const [sheetName, rect] = parsedRange(raw, value, label, kind);
    if (sheetName !== null) {
      const index = raw.getSheetIndex(sheetName);
      if (index === null || index === undefined) {
        throw new FormulonError(
          StatusCode.NotFound,
          `${label} ${JSON.stringify(value)}: no sheet is named ${JSON.stringify(sheetName)}`,
          `sheet_name=${sheetName}`,
        );
      }
      return { sheet: index, rect };
    }
    if (sheet === undefined) {
      throw new FormulonError(StatusCode.InvalidArgument, `${label} must name a sheet`, `address=${value}`);
    }
    return { sheet, rect };
  }
  const rect = unqualifiedRect(raw, value, label);
  if (rect.firstRow !== rect.lastRow || rect.firstCol !== rect.lastCol) {
    throw new FormulonError(StatusCode.InvalidArgument, `${label} must name one cell`, '');
  }
  return { row: rect.firstRow, col: rect.firstCol };
}

const RECT_FIELDS = ['firstRow', 'firstCol', 'lastRow', 'lastCol'];

/** Checks the coordinate fields of a `CellRect`, `CellAddress` or `SheetRect`; absent ones are left to the call. */
function checkRangeFields(kind, value, label) {
  const check = (record, fields, path) => {
    for (const field of fields) {
      if (record[field] !== undefined) requireIndex(record[field], `${path}.${field}`, 'coordinate');
    }
  };
  if (kind === 'CellAddress') check(value, ['row', 'col'], label);
  else if (kind === 'SheetRect') {
    if (value.sheet !== undefined) requireIndex(value.sheet, `${label}.sheet`, 'sheet');
    if (isObject(value.rect)) check(value.rect, RECT_FIELDS, `${label}.rect`);
  } else check(value, RECT_FIELDS, label);
  return value;
}

/** The `{ row, col }` of a `CellAddress` or a one-cell A1 string. */
function addressOf(raw, value, label) {
  if (typeof value === 'string') return normalizeAddress(raw, 'CellAddress', value, undefined, label);
  requireObject(value, label);
  return checkRangeFields('CellAddress', value, label);
}

/** `record` with its key `from` named `to`; anything but an object (a `null` read) passes as it is. */
function renamed(record, from, to) {
  if (!isObject(record) || !(from in record)) return record;
  const { [from]: value, ...rest } = record;
  return { ...rest, [to]: value };
}

/** `record` with each `[from, to]` of `pairs` renamed. */
function renamedAll(record, pairs) {
  return pairs.reduce((out, [from, to]) => renamed(out, from, to), record);
}

/** The text `value` (a `SheetRect`, or an A1 string resolved on `sheet`) is as a sheet-qualified reference. */
function sheetRectText(raw, value, sheet, label) {
  const { sheet: at, rect } = normalizeAddress(raw, 'SheetRect', value, sheet, label);
  return raw.formatReference({ sheetName: raw.getSheetName(at), rect });
}

/** The raw `{ sheetName, rect }` of a stored reference text; `null` when it is empty or does not parse. */
function parseReferenceOrNull(raw, text) {
  if (typeof text !== 'string' || text === '') return null;
  try {
    return raw.parseReference(text);
  } catch (e) {
    if (e?.name === 'FormulonError' && e.parseError !== undefined) return null;
    throw e;
  }
}

/** The `SheetRect` a stored qualified reference text names; `null` when it is empty, unqualified or unparseable. */
function parseSheetRect(raw, text) {
  const ref = parseReferenceOrNull(raw, text);
  if (ref === null || ref.sheetName === null) return null;
  const sheet = raw.getSheetIndex(ref.sheetName);
  return sheet === null || sheet === undefined ? null : { sheet, rect: ref.rect };
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
 * range (an unqualified one resolves on `sheet`), a `CellAddress` field one A1 cell (read as `{ row, col }`) and a
 * nested record its own fields. Anything but an object passes as it is.
 * `label` words the errors: a field `f` is `<label>.f`.
 */
function normalizeRecordRanges(raw, record, type, sheet, label) {
  const schema = RECORD_RANGE_FIELDS[type];
  if (schema === undefined || !isObject(record)) return record;
  let out = record;
  for (const [field, kind] of Object.entries(schema)) {
    const value = record[field];
    const next = normalizeAddress(raw, kind, value, sheet, `${label}.${field}`);
    if (next !== value) {
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
    const next = normalizeAddress(raw, kind, value, sheet, argLabel(where, at));
    if (next !== value) {
      if (out === args) out = [...args];
      out[at] = next;
    }
  }
  return out;
}

// `args` with each omitted `DEFAULT_ARGS` slot filled.
function fillDefaults(defaults, args) {
  let out = args;
  for (const [at, value] of defaults) {
    if (args[at] !== undefined) continue;
    if (out === args) out = [...args];
    out[at] = value;
  }
  return out;
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
  if (ref.sheetName !== null) return false;
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

// The `Pivot` methods whose first argument is a `PivotFieldRef`.
const PIVOT_FIELD_ARGUMENTS = Object.freeze([
  'getField',
  'updateField',
  'setFieldAxis',
  'listFieldItems',
  'getFieldItem',
  'updateFieldItem',
  'addFieldItem',
  'clearFieldItems',
  'setFieldSubtotalFunctions',
  'setFieldDateGroup',
  'getFieldDateGroup',
  'setFieldNumericGroup',
  'getFieldNumericGroup',
  'getDataFieldDefaultName',
]);

// The `PivotCache` methods taking a `PivotFieldRef`, with its argument position.
const PIVOT_CACHE_FIELD_ARGUMENTS = Object.freeze([
  ['setCalculatedFieldFormula', 0],
  ['removeCalculatedField', 0],
  ['addCalculatedItem', 0],
  ['listFieldSharedItems', 0],
  ['addFieldSharedItem', 0],
  ['clearFieldSharedItems', 0],
  ['getRecordValue', 1],
  ['setRecordValue', 1],
]);

/** The field index a `PivotFieldRef` names: a number is the index, a string the source name matched as names are. */
function pivotFieldIndex(raw, sheet, pivot, ref, label) {
  if (typeof ref === 'number') return requireIndex(ref, label, 'position');
  return cacheFieldIndex(raw, raw.getPivotCacheId(sheet, pivot), ref, label);
}

/** `pivotFieldIndex` against the fields of the cache `cacheId`. */
function cacheFieldIndex(raw, cacheId, ref, label) {
  if (typeof ref === 'number') return requireIndex(ref, label, 'position');
  if (typeof ref !== 'string') throw new TypeError(`${label} must be a field index or a field source name`);
  const want = foldName(ref);
  for (let i = 0, n = raw.getPivotCacheFieldCount(cacheId); i < n; i++) {
    if (foldName(raw.getPivotCacheFieldName(cacheId, i)) === want) return i;
  }
  throw new FormulonError(
    StatusCode.NotFound,
    `no field has the source name ${JSON.stringify(ref)}`,
    `field_name=${ref}`,
  );
}

/** `patch` with its `ranges` (A1 text or rectangles) as `CellRect`s. */
function withRanges(raw, where, patch) {
  const label = recordLabel(where, ['ranges']);
  return { ...patch, ranges: normalizeAddress(raw, 'CellRect[]', patch.ranges, undefined, label) };
}

/** `ranges` is not clearable: a patch's `null` there is a TypeError, as on any other non-nullable key. */
function refuseNullRanges(where, patch) {
  if (isObject(patch) && patch.ranges === null) {
    throw new TypeError(`${recordLabel(where, ['ranges'])} must be an array`);
  }
}

/**
 * The index form of a call whose `*Ref` argument (`REF_ARGS`) may be a string or
 * a handle: a string is the first element whose key matches it
 * case-insensitively, as sheet names match, listed afresh on every call; a
 * handle is its element's current index (`handleIndex`, undefined for a value
 * that is not one). A kind without a `REF_KINDS` entry (a pivot cache, keyed by
 * id) has no name form and is left to the method.
 */
function resolveRef(raw, sheet, where, [at, kind], args, handleIndex) {
  const value = args[at];
  const spec = REF_KINDS[kind];
  if (spec === undefined || (typeof value === 'number' && spec.keyed !== true)) return args;
  const label = argLabel(where, at);
  if (spec.keyed === true) {
    if (typeof value === 'string') return args;
    if (typeof value !== 'number') throw new TypeError(`${label} must be a position or a ${spec.noun}`);
    requireIndex(value, label, 'position');
    const count = spec.count(raw, sheet);
    if (value >= count) {
      throw new FormulonError(
        StatusCode.IndexOutOfRange,
        `${label} ${value} ${pastCount(count, 'item')}`,
        `position=${value}`,
      );
    }
    const out = [...args];
    out[at] = spec.key(raw, sheet, value);
    return out;
  }
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

// The key tables of a patch argument and of an input record argument.
const PATCH_TABLES = Object.freeze({ keys: PATCH_KEYS, nested: PATCH_NESTED_KEYS });
const INPUT_TABLES = Object.freeze({ keys: INPUT_KEYS, nested: INPUT_NESTED_KEYS });

/**
 * Throws `TypeError: <where> <path> has no field <key>` for a key that
 * neither the type of `where` (`Class.method`) nor its getter's record
 * declares (`tables.keys`), and, at each path `tables.nested` lists, for a
 * key of the object there outside its list (a path segment ending in `[]`
 * names each element of an array); `undefined` is ignored. `source` names the table entry when
 * `where` words the error for a method whose record is the entry's twin.
 */
function refuseUnknownKeys(where, patch, source = where, tables = PATCH_TABLES) {
  const [position, keys] = tableEntry(tables.keys, source);
  const nested = tableEntry(tables.nested, source) ?? {};
  const visit = (value, accepted, prefix, trail) => {
    for (const [key, item] of Object.entries(value)) {
      if (item === undefined) continue;
      const path = `${prefix}${key}`;
      if (!accepted.includes(key)) throw new TypeError(`${trail} has no field ${key}`);
      if (Object.hasOwn(nested, path) && isObject(item)) visit(item, nested[path], `${path}.`, `${trail}.${key}`);
      const each = `${path}[]`;
      if (Object.hasOwn(nested, each) && Array.isArray(item)) {
        for (const [i, element] of item.entries()) {
          if (isObject(element)) visit(element, nested[each], `${each}.`, `${trail}.${key}[${i}]`);
        }
      }
    }
  };
  visit(patch, keys, '', argLabel(where, position));
}

/** Refuses a key outside the input record of `where` (`INPUT_KEYS`) that `value` is; anything but an object passes. */
function refuseUnknownInputKeys(where, value) {
  if (isObject(value)) refuseUnknownKeys(where, value, where, INPUT_TABLES);
}

/**
 * Checks a hand-written patch of `where` (`Class.method`): an object whose keys
 * are in `PATCH_KEYS`, where a field of `fields` refuses `null`, one of
 * `nullable` takes `null`, and each typed value has its type.
 */
function checkPatch(where, patch, fields, nullable = {}, source = where) {
  requireObject(patch, argLabel(where, tableEntry(PATCH_KEYS, source)[0]));
  refuseUnknownKeys(where, patch, source);
  for (const [key, value] of Object.entries(patch)) {
    const type = fields[key] ?? nullable[key];
    if (value === undefined || type === undefined) continue;
    if (value === null && Object.hasOwn(nullable, key)) continue;
    const mismatch =
      type === 'array' ? !Array.isArray(value) : type === 'object' ? !isObject(value) : typeof value !== type;
    if (mismatch) throw new TypeError(`${recordLabel(where, [key], source)} must be ${typeNoun(type)}`);
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
 * a field of `fields` (`TypeError: <where> <path> must be a <type>`). A
 * non-object patch is left to the raw method.
 */
function checkForwardedPatch(where, patch, fields = {}) {
  if (!isObject(patch)) return;
  refuseUnknownKeys(where, patch);
  for (const [field, type] of Object.entries(fields)) {
    const [head, tail] = field.split('.');
    const value = tail === undefined ? patch[head] : isObject(patch[head]) ? patch[head][tail] : undefined;
    if (value === null) throw new TypeError(`${recordLabel(where, [field])} must be ${typeNoun(type)}`);
  }
}

/**
 * The positional `[cursor?, limit?]` the raw paging calls take, from a `PageOptions` argument at `position`: an
 * omitted or `null` key is left to the engine, whose maximum page a limit of 0 asks for.
 */
function pageArguments(where, position, options) {
  if (options === undefined || options === null) return [];
  const label = argLabel(where, position);
  requireObject(options, label);
  refuseOtherKeys(label, options, ['cursor', 'limit']);
  const cursor = options.cursor ?? undefined;
  const limit = options.limit ?? undefined;
  if (cursor !== undefined) {
    requireNumber(cursor, `${label}.cursor`);
    if (!Number.isInteger(cursor)) throw new RangeError(`${label}.cursor must be an integer`);
    if (cursor < 0) {
      throw new FormulonError(StatusCode.InvalidArgument, `${label}.cursor must not be negative`, `cursor=${cursor}`);
    }
    if (cursor > Number.MAX_SAFE_INTEGER) throw new RangeError(`${label}.cursor is too large`);
  }
  if (limit !== undefined) requireIndex(limit, `${label}.limit`, 'quantity');
  if (limit === undefined) return cursor === undefined ? [] : [cursor];
  return [cursor ?? 0, limit];
}

function geometryMode(mode, where, position) {
  if (mode === undefined) return GeometryMode.Display;
  const label = argLabel(where, position);
  requireNumber(mode, label);
  if (!GEOMETRY_MODES.has(mode)) {
    throw new FormulonError(StatusCode.InvalidArgument, `${label} ${mode} is not a GeometryMode`, `mode=${mode}`);
  }
  return mode;
}

/** The typed raw suffix and arguments for a pivot-cache value; `dateSerialOf(date)` is a `Date`'s serial. */
function cacheValue(where, position, value, dateSerialOf) {
  if (value === null) return ['Blank', []];
  // A `PivotCacheValue` read back: a date keeps its serial and its date flag.
  if (isObject(value) && typeof value.isDate === 'boolean' && Object.hasOwn(value, 'value')) {
    if (value.isDate && typeof value.value === 'number') return ['Date', [value.value]];
    return cacheValue(where, position, value.value, dateSerialOf);
  }
  if (value instanceof Date) return ['Date', [dateSerialOf(value)]];
  switch (typeof value) {
    case 'number':
      return ['Number', [value]];
    case 'string':
      return ['Text', [value]];
    case 'boolean':
      return ['Bool', [value]];
    case 'object':
      if (value.kind === ValueKind.Error && typeof value.errorCode === 'number') return ['Error', [value.errorCode]];
      break;
  }
  throw new TypeError(`${argLabel(where, position)} must be a pivot cache value`);
}

// The payload field and its type of each scalar `Value` kind but Blank.
const SCALAR_VALUE_FIELDS = Object.freeze({
  [ValueKind.Number]: ['number', 'number'],
  [ValueKind.Bool]: ['boolean', 'boolean'],
  [ValueKind.Text]: ['text', 'string'],
  [ValueKind.Error]: ['errorCode', 'number'],
});

/** Throws unless `value` is something `Sheet.setValue` writes as one cell. */
function checkCellValue(value, label) {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return;
  if (value instanceof ExcelError) return;
  if (value instanceof Date) {
    if (Number.isNaN(value.getTime())) throw new RangeError(`${label} must be a valid Date`);
    return;
  }
  let number = value;
  if (isObject(value)) {
    if (value.kind === ValueKind.Blank) return;
    const [field, type] = SCALAR_VALUE_FIELDS[value.kind] ?? [];
    if (field === undefined || typeof value[field] !== type) throw new TypeError(`${label} must be a cell value`);
    if (value.kind !== ValueKind.Number) return;
    number = value.number;
  } else if (typeof value !== 'number') {
    throw new TypeError(`${label} must be a cell value`);
  }
  if (!Number.isFinite(number)) {
    throw new FormulonError(StatusCode.InvalidArgument, `${label} must be finite`, `value=${number}`);
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
 * A `CivilTime` from a date-time input: a civil record needs `year`, `month` and `day` and takes 0 for an omitted
 * time of day; a `Date` is read through its local getters, or its UTC getters under `utc` (milliseconds dropped),
 * and an Invalid Date is a `RangeError`. `where` and `param` (the argument's path) word the errors.
 */
function toCivil(input, where, param, { utc = false } = {}) {
  const label = `${where} ${param}`;
  if (input instanceof Date) {
    if (Number.isNaN(input.getTime())) throw new RangeError(`${label} must be a valid Date`);
    return utc
      ? {
          year: input.getUTCFullYear(),
          month: input.getUTCMonth() + 1,
          day: input.getUTCDate(),
          hour: input.getUTCHours(),
          minute: input.getUTCMinutes(),
          second: input.getUTCSeconds(),
        }
      : {
          year: input.getFullYear(),
          month: input.getMonth() + 1,
          day: input.getDate(),
          hour: input.getHours(),
          minute: input.getMinutes(),
          second: input.getSeconds(),
        };
  }
  requireObject(input, label);
  for (const [key, value] of Object.entries(input)) {
    if (value !== undefined && !CIVIL_FIELDS.includes(key)) throw new TypeError(`${label} has no field ${key}`);
  }
  const civil = {};
  for (const field of CIVIL_FIELDS) {
    const value = input[field];
    if (value === undefined && CIVIL_DATE_FIELDS.includes(field)) throw new TypeError(`${label}.${field} is required`);
    civil[field] = value === undefined ? 0 : requireInteger(value, `${label}.${field}`);
  }
  return civil;
}

/** The serial of `civil` under the date system `date1904`. */
const civilSerial = (runtime, civil, date1904) => runtime.convertCivilToSerial(civil, date1904);

/** The serial of a date-time input under the date system `date1904`. */
const dateInputSerial = (runtime, input, date1904, where, param) =>
  civilSerial(runtime, toCivil(input, where, param), date1904);

/** The `YYYY-MM-DDTHH:MM:SS.00` stamp of a civil time, as the C threaded comment carries `created`. */
function civilStamp(civil) {
  const pad = (value, width) => String(value).padStart(width, '0');
  const date = `${pad(civil.year, 4)}-${pad(civil.month, 2)}-${pad(civil.day, 2)}`;
  return `${date}T${pad(civil.hour, 2)}:${pad(civil.minute, 2)}:${pad(civil.second, 2)}.00`;
}

/** The `CivilTime` a C threaded-comment stamp names, hundredths dropped; `null` when it names none. */
function stampCivil(text) {
  const match = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/.exec(text);
  if (match === null) return null;
  const [year, month, day, hour, minute, second] = match.slice(1).map(Number);
  return { year, month, day, hour, minute, second };
}

/** Throws `TypeError: <label> has no field <key>` for a key of `object` outside `accepted`. */
function refuseOtherKeys(label, object, accepted) {
  for (const [key, value] of Object.entries(object)) {
    if (value !== undefined && !accepted.includes(key)) throw new TypeError(`${label} has no field ${key}`);
  }
}

/**
 * The serial of `date`'s local fields under the date system `date1904`, and whether it has a time of day. The
 * engine converts the calendar date; the time of day is added here.
 */
function dateSerial(runtime, date, date1904, where) {
  if (Number.isNaN(date.getTime())) throw new RangeError(`${where} must be a valid Date`);
  const civil = {
    year: date.getFullYear(),
    month: date.getMonth() + 1,
    day: date.getDate(),
    hour: 0,
    minute: 0,
    second: 0,
  };
  const seconds = date.getHours() * 3600 + date.getMinutes() * 60 + date.getSeconds();
  const ms = date.getMilliseconds();
  const serial = runtime.convertCivilToSerial(civil, date1904) + seconds / 86400 + ms / MS_PER_DAY;
  return [serial, seconds !== 0 || ms !== 0];
}

/**
 * The tagged `Value` that the cell-value argument `value` (labelled `label`) stands for: a string is always text and
 * a `Date` its serial under the workbook's date system.
 */
function valueRecord(runtime, raw, value, label) {
  checkCellValue(value, label);
  if (value === null) return BLANK;
  if (typeof value === 'number') return { kind: ValueKind.Number, number: value };
  if (typeof value === 'string') return { kind: ValueKind.Text, text: value };
  if (typeof value === 'boolean') return { kind: ValueKind.Bool, boolean: value };
  if (value instanceof ExcelError) return { kind: ValueKind.Error, errorCode: value.code };
  if (value instanceof Date) {
    return { kind: ValueKind.Number, number: dateSerial(runtime, value, raw.getDate1904(), label)[0] };
  }
  return value;
}

/** The `Date` whose local fields carry `serial` under the date system `date1904`, rounded to the millisecond. */
function serialDate(runtime, serial, date1904) {
  let days = Math.floor(serial);
  let ms = Math.round((serial - days) * MS_PER_DAY);
  if (ms === MS_PER_DAY) {
    days += 1;
    ms = 0;
  }
  const { year, month, day } = runtime.convertSerialToCivil(days, date1904);
  const date = new Date(0);
  date.setFullYear(year, month - 1, day);
  date.setHours(Math.floor(ms / 3600000), Math.floor(ms / 60000) % 60, Math.floor(ms / 1000) % 60, ms % 1000);
  return date;
}

// Record keys whose value is a `ColorInput`; a colour string there becomes a `Color`.
const CHART_SOURCE_KEYS = new Set(['name', 'categories', 'values', 'x', 'y', 'size']);

/** A `ChartDataSource` record as the raw layer reads it: reference text, or the record itself for a literal one. */
function chartSourceInput(source) {
  if (source.kind === ChartSourceKind.Literal) return source;
  return source.kind === ChartSourceKind.Ref ? source.formula : '';
}

/**
 * The raw options `{ barDirection?, grouping? }` of a chart type write, `null` entries as omitted; a bar direction
 * on a type other than `Bar` is a `TypeError`.
 */
function chartTypeOptions(where, position, type, options) {
  const out = {};
  if (options === undefined || options === null) return out;
  const label = argLabel(where, position);
  requireObject(options, label);
  refuseOtherKeys(label, options, ['barDirection', 'grouping']);
  for (const key of ['barDirection', 'grouping']) if (present(options[key])) out[key] = options[key];
  if (out.barDirection !== undefined && type !== ChartPlotType.Bar) {
    throw new TypeError(`${label}.barDirection applies to the Bar type only`);
  }
  return out;
}

const COLOR_INPUT_KEYS = new Set([
  'color',
  'foreground',
  'background',
  'colors',
  'fill',
  'negativeFill',
  'border',
  'negativeBorder',
  'axisColor',
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
    else if (CHART_SOURCE_KEYS.has(key) && isObject(item) && typeof item.kind === 'number' && 'formula' in item)
      next = chartSourceInput(item);
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
    const next = colorInputs(arg, argLabel(where, i));
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

function requireValueField(value, field, type, label) {
  if (typeof value[field] !== type) throw new TypeError(`${label}.${field} must be ${typeNoun(type)}`);
  return value[field];
}

// ---------------------------------------------------------------------------
// Cell formats: `Sheet.getCellFormat` / `Sheet.updateCellFormat` over the style
// records. tools/api/cell_format_cases.json pins the behaviour for the JS and
// Python facades alike.
// ---------------------------------------------------------------------------

const enumField = (table, nullable = false) => ({ kind: 'enum', max: Math.max(...Object.values(table)), nullable });
const FORMAT_BOOLEAN = Object.freeze({ kind: 'boolean' });
const FORMAT_FLAG = Object.freeze({ kind: 'boolean', nullable: true });
const FORMAT_COLOR = Object.freeze({ kind: 'color' });
const FORMAT_COLOR_OR_NULL = Object.freeze({ kind: 'color', nullable: true });
const FORMAT_SIDE = Object.freeze({ kind: 'side' });

// How each key of each group of a `CellFormatPatch` is checked; a nullable key
// takes `null` to clear the attribute. The accepted keys are `PATCH_NESTED_KEYS`.
const FORMAT_FIELDS = Object.freeze({
  font: {
    name: { kind: 'string', nullable: true },
    sizePt: { kind: 'size', nullable: true },
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
    color: FORMAT_COLOR_OR_NULL,
  },
  fill: {
    pattern: enumField(FillPattern),
    foreground: FORMAT_COLOR,
    background: FORMAT_COLOR,
    gradient: { kind: 'object', nullable: true },
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
  if (!Number.isInteger(value)) throw new RangeError(`${label} must be an integer`);
  if (value < 0) throw new RangeError(`${label} must not be negative`);
  return value;
}

const COLOR_FIELDS = Object.freeze(['kind', 'argb', 'theme', 'tint', 'indexed']);

/** A `ColorInput` as a `Color`: `#RRGGBB` (opaque) or `#AARRGGBB`, or a `Color` record. */
function formatColor(value, label) {
  if (typeof value === 'string') {
    if (!COLOR_TEXT.test(value)) throw new TypeError(`${label} must be "#RRGGBB" or "#AARRGGBB"`);
    const hex = value.length === 7 ? `FF${value.slice(1)}` : value.slice(1);
    return { kind: ColorKind.Rgb, argb: Number.parseInt(hex, 16), theme: 0, tint: 0, indexed: 0 };
  }
  if (!isObject(value)) throw new TypeError(`${label} must be a colour string or a Color`);
  for (const [key, item] of Object.entries(value)) {
    if (item !== undefined && !COLOR_FIELDS.includes(key)) throw new TypeError(`${label} has no field ${key}`);
  }
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
      return requireNumber(value, label);
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
  const where = 'Sheet.updateCellFormat';
  requireObject(patch, argLabel(where, 1));
  refuseUnknownKeys(where, patch);
  const out = {};
  for (const [key, value] of Object.entries(patch)) {
    if (value === undefined || READ_ONLY_KEYS[where].includes(key)) continue;
    if (key === 'numberFormat') {
      const label = recordLabel(where, [key]);
      requireString(value, label);
      if (value === '') {
        throw new FormulonError(StatusCode.InvalidArgument, `${label} must not be empty`, 'number_format=');
      }
      out.numberFormat = value;
      continue;
    }
    out[key] = formatGroup(key, value, recordLabel(where, [key]));
  }
  return out;
}

function mergeFont(current, patch) {
  return { ...current, ...patch };
}

// Setting a colour on a fill without a pattern makes it solid, Excel's "fill
// colour"; any fill key drops a gradient unless the patch names one, and a `null` gradient drops it.
function mergeFill(current, patch) {
  const out = { ...current };
  if (patch.gradient == null) delete out.gradient;
  else out.gradient = patch.gradient;
  for (const key of ['pattern', 'foreground', 'background']) if (patch[key] !== undefined) out[key] = patch[key];
  const colored = patch.foreground !== undefined || patch.background !== undefined;
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
 * The `CellXfInput` that rebuilds `xf` as `addCellXf` reads it: an absent attribute is `null` or omitted, so the
 * record copies, and an optional attribute is copied only when the record has it.
 */
function cellXfInput(xf) {
  const input = {
    fontIndex: xf.fontIndex,
    fillIndex: xf.fillIndex,
    borderIndex: xf.borderIndex,
    numFmtId: xf.numFmtId,
    applyNumberFormat: xf.applyNumberFormat,
    applyFont: xf.applyFont,
    applyFill: xf.applyFill,
    applyBorder: xf.applyBorder,
    applyAlignment: xf.applyAlignment,
    applyProtection: xf.applyProtection,
    quotePrefix: xf.quotePrefix,
  };
  for (const key of ['horizontal', 'vertical', 'wrapText', 'justifyLastLine', 'locked', 'hidden']) {
    if (present(xf[key])) input[key] = xf[key];
  }
  for (const key of OPTIONAL_XF_FIELDS) if (present(xf[key])) input[key] = xf[key];
  if (present(xf.styleXfIndex)) input.styleXfIndex = xf.styleXfIndex;
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
  if (patch.alignment !== undefined) Object.assign(input, patch.alignment);
  if (patch.protection !== undefined) {
    input.locked = patch.protection.locked ?? input.locked ?? true;
    input.hidden = patch.protection.hidden ?? input.hidden ?? false;
  }
  return input;
}

/** The `CellFormat` of the cell XF record `xf`, the effective XF read from layer `source`. */
function cellFormatOf(raw, xf, source) {
  const code = xf.numFmtId === 0 ? null : raw.getNumberFormat(xf.numFmtId);
  return {
    source,
    numberFormat: code ?? 'General',
    font: raw.getFont(xf.fontIndex),
    fill: raw.getFill(xf.fillIndex),
    border: raw.getBorder(xf.borderIndex),
    alignment: {
      horizontal: xf.horizontal ?? HorizontalAlign.General,
      vertical: xf.vertical ?? VerticalAlign.Bottom,
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
    #sheet;
    #id;

    constructor(key, sheet, id) {
      if (key !== token) throw new TypeError('Table has no public constructor; use Sheet.table or Workbook.table');
      this.#sheet = sheet;
      this.#id = id;
    }

    static {
      tableState = {
        create: (sheet, id) => new Table(token, sheet, id),
        locate: (table) => table.#locate(),
      };
    }

    // The raw workbook, the sheet's current index and the table's current index on it.
    #locate() {
      const sheet = sheetIndex(this.#sheet);
      const raw = rawOf(this.#sheet.workbook);
      const at = resolveRuntimeId(() => raw.resolveTableRuntimeId(this.#id), 'the table was removed');
      return [raw, sheet, at.index];
    }

    get workbook() {
      return this.#sheet.workbook;
    }

    get sheet() {
      return this.#sheet;
    }

    get index() {
      return this.#locate()[2];
    }

    get name() {
      const [raw, sheet, index] = this.#locate();
      return raw.getTable(sheet, index).name;
    }

    // `range` reaches the raw method as its `rect`; the `name` of a read record is ignored.
    update(patch) {
      const where = 'Table.update';
      const fields = PATCH_FIELDS[where];
      checkPatch(where, patch, fields);
      if (patch.range === null) {
        throw new TypeError(`${argLabel(where, 0, 'range')} must be a CellRect, a SheetRect or an A1 string`);
      }
      const [raw, sheet, index] = this.#locate();
      const input = appliedFields(patch, fields);
      if (patch.range !== undefined) input.rect = tableRange(raw, sheet, patch.range, where);
      if (patch.columns !== undefined) input.columns = patch.columns;
      raw.updateTable(sheet, index, input);
    }

    rename(name) {
      requireString(name, argLabel('Table.rename', 0));
      const [raw, sheet, index] = this.#locate();
      raw.renameTable(sheet, index, name);
    }

    // Both omitted appends one at the current count; `first` alone is one at `first`.
    insertRows(first, last) {
      const [raw, sheet, index] = this.#locate();
      const table = raw.getTable(sheet, index);
      const { firstRow, lastRow } = table.range.rect;
      const count = lastRow - firstRow + 1 - (table.headerRow ? 1 : 0) - (table.totalsRow ? 1 : 0);
      const span = insertSpan('Table.insertRows', first, last, count);
      raw.insertTableRows(sheet, index, span[0], span[1]);
    }

    insertColumns(first, last) {
      const [raw, sheet, index] = this.#locate();
      const span = insertSpan('Table.insertColumns', first, last, raw.getTable(sheet, index).columns.length);
      raw.insertTableColumns(sheet, index, span[0], span[1]);
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
      return {
        name: raw.getPivotName(sheet, pivot),
        cacheId: raw.getPivotCacheId(sheet, pivot),
        reportLayout: raw.getPivotReportLayout(sheet, pivot),
        grandTotals: raw.getPivotGrandTotals(sheet, pivot),
        styleOptions: raw.getPivotStyleOptions(sheet, pivot),
      };
    }

    cache() {
      const [raw, sheet, pivot] = this.#locate();
      return pivotCacheOf(this.#sheet.workbook, raw.getPivotCacheId(sheet, pivot));
    }

    getLayout() {
      const [raw, sheet, pivot] = this.#locate();
      const layout = raw.getPivotLayout(sheet, pivot);
      const cells = layout.cells.map(({ row, col, value, kind, depth, fieldName, numberFormat }) => ({
        row,
        col,
        value: toNative(value),
        kind,
        depth,
        fieldName,
        numberFormat,
      }));
      return { range: layout.rect, cells };
    }

    // `value` / `valueHigh` reach the raw method as they are; it maps them onto the C operand union.
    addFilter(spec) {
      const where = 'Pivot.addFilter';
      const label = argLabel(where, 0);
      requireObject(spec, label);
      refuseOtherKeys(label, spec, ['fieldIndex', 'kind', 'dataFieldIndex', 'value', 'valueHigh']);
      const [raw, sheet, pivot] = this.#locate();
      const fieldIndex = pivotFieldIndex(raw, sheet, pivot, spec.fieldIndex, `${label}.fieldIndex`);
      return raw.addPivotFilter(sheet, pivot, { ...spec, fieldIndex });
    }

    showDetails(row, col) {
      const [raw, sheet, pivot] = this.#locate();
      return insertSheet(this.#sheet.workbook, raw.showPivotDetails(sheet, pivot, row, col));
    }

    // The whole patch is checked here and applied by one raw call; the record keys it cannot write are dropped.
    updateField(field, patch) {
      const where = 'Pivot.updateField';
      const fields = PATCH_FIELDS[where];
      const nullable = PATCH_NULLABLE[where];
      checkPatch(where, patch, fields, nullable);
      const [raw, sheet, pivot] = this.#locate();
      raw.updatePivotField(sheet, pivot, field, appliedFields(patch, fields, nullable));
    }

    updateFieldItem(field, item, patch) {
      const where = 'Pivot.updateFieldItem';
      const fields = PATCH_FIELDS[where];
      const nullable = PATCH_NULLABLE[where];
      checkPatch(where, patch, fields, nullable);
      const [raw, sheet, pivot] = this.#locate();
      raw.updatePivotFieldItem(sheet, pivot, field, item, appliedFields(patch, fields, nullable));
    }

    updateStyleOptions(patch) {
      const where = 'Pivot.updateStyleOptions';
      const fields = PATCH_FIELDS[where];
      const nullable = PATCH_NULLABLE[where];
      checkPatch(where, patch, fields, nullable);
      const [raw, sheet, pivot] = this.#locate();
      raw.updatePivotStyleOptions(sheet, pivot, appliedFields(patch, fields, nullable));
    }

    // `item` is a new item's name, or `{ cacheIndex }` to bring in the cache's shared item at that position.
    addFieldItem(field, item, options) {
      const where = 'Pivot.addFieldItem';
      let visible = true;
      if (options !== undefined && options !== null) {
        const label = argLabel(where, 2);
        requireObject(options, label);
        refuseOtherKeys(label, options, ['visible']);
        if (options.visible !== undefined) visible = requireBoolean(options.visible, `${label}.visible`);
      }
      if (typeof item !== 'string') {
        const label = argLabel(where, 1);
        requireObject(item, label);
        refuseOtherKeys(label, item, ['cacheIndex']);
        requireIndex(item.cacheIndex, `${label}.cacheIndex`, 'position');
        item = { cacheIndex: item.cacheIndex };
      }
      const [raw, sheet, pivot] = this.#locate();
      return raw.addPivotFieldItem(sheet, pivot, field, item, visible);
    }

    updateGrandTotals(patch) {
      const where = 'Pivot.updateGrandTotals';
      requireObject(patch, argLabel(where, 0));
      refuseUnknownKeys(where, patch);
      const [raw, sheet, pivot] = this.#locate();
      const current =
        patch.rows === undefined || patch.columns === undefined ? raw.getPivotGrandTotals(sheet, pivot) : {};
      const rows = patchField(patch, 'rows', current.rows, requireBoolean, where);
      const columns = patchField(patch, 'columns', current.columns, requireBoolean, where);
      raw.updatePivotGrandTotals(sheet, pivot, rows, columns);
    }

    setFieldDateGroup(field, group) {
      const where = 'Pivot.setFieldDateGroup';
      const [raw, sheet, pivot] = this.#locate();
      if (group === null) {
        raw.setPivotFieldDateGroup(sheet, pivot, field, null);
        return;
      }
      const label = argLabel(where, 1);
      requireObject(group, label);
      const bound = (key) =>
        group[key] === undefined
          ? undefined
          : dateInputSerial(runtime, group[key], raw.getDate1904(), where, `${argumentPath(where, 1)}.${key}`);
      const args = [
        requireInteger(group.periods, `${label}.periods`),
        group.calendar === undefined ? PivotCalendar.Gregorian : requireInteger(group.calendar, `${label}.calendar`),
      ];
      const intervalDays =
        group.intervalDays === undefined ? undefined : requireNumber(group.intervalDays, `${label}.intervalDays`);
      // The raw call is positional: a bound after an omitted slot needs that slot's default spelled out.
      const tail = [intervalDays ?? 1, bound('start') ?? Number.NaN, bound('end') ?? Number.NaN];
      while (tail.length > 0 && (tail.length === 1 ? tail[0] === 1 : Number.isNaN(tail.at(-1)))) tail.pop();
      raw.setPivotFieldDateGroup(sheet, pivot, field, ...args, ...tail);
    }

    // The stored bounds read back as `Date`s, so a read group can be written back through `setFieldDateGroup`.
    getFieldDateGroup(field) {
      const [raw, sheet, pivot] = this.#locate();
      const group = raw.getPivotFieldDateGroup(sheet, pivot, field);
      if (group === null) return null;
      const date1904 = raw.getDate1904();
      return {
        ...group,
        start: serialDate(runtime, group.start, date1904),
        end: serialDate(runtime, group.end, date1904),
      };
    }

    setFieldNumericGroup(field, group) {
      const where = 'Pivot.setFieldNumericGroup';
      const [raw, sheet, pivot] = this.#locate();
      if (group === null) {
        raw.setPivotFieldNumericGroup(sheet, pivot, field, null);
        return;
      }
      const label = argLabel(where, 1);
      requireObject(group, label);
      const interval = requireNumber(group.interval, `${label}.interval`);
      const bound = (key) => (group[key] === undefined ? Number.NaN : requireNumber(group[key], `${label}.${key}`));
      // The raw call is positional: a NaN bound is an omitted one.
      raw.setPivotFieldNumericGroup(sheet, pivot, field, interval, bound('start'), bound('end'));
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

    #raw() {
      return liveRaw(this.#workbook);
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
        fields.push({
          name: raw.getPivotCacheFieldName(id, i),
          formula: raw.getPivotCacheFieldFormula(id, i),
          ...raw.getPivotCacheFieldTypes(id, i),
        });
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
          value: toNative(raw.getPivotCacheFieldSharedItemValue(id, field, i)),
          isDate: raw.isPivotCacheFieldSharedItemDate(id, field, i),
        });
      }
      return items;
    }

    addFieldSharedItem(field, value) {
      const where = 'PivotCache.addFieldSharedItem';
      const raw = this.#raw();
      const [kind, args] = cacheValue(where, 1, value, (date) =>
        dateInputSerial(runtime, date, raw.getDate1904(), where, argumentPath(where, 1)),
      );
      return raw[`addPivotCacheFieldSharedItem${kind}`](this.#cacheId, field, ...args);
    }

    setRecordValue(record, field, value) {
      const where = 'PivotCache.setRecordValue';
      requireIndex(record, argLabel(where, 0), 'position');
      requireIndex(field, argLabel(where, 1), 'position');
      const raw = this.#raw();
      const [kind, args] = cacheValue(where, 2, value, (date) =>
        dateInputSerial(runtime, date, raw.getDate1904(), where, argumentPath(where, 2)),
      );
      raw[`setPivotCacheRecord${kind}`](this.#cacheId, record, field, ...args);
    }

    // The read record folds the source's sheet and reference into one `SheetRect`; it is `null` for a source
    // without a reference or one on a sheet the workbook does not have.
    getWorksheetSource() {
      const raw = this.#raw();
      const source = raw.getPivotCacheWorksheetSource(this.#cacheId);
      if (source === null || source === undefined) return null;
      let range = null;
      if (typeof source.sheet === 'string') {
        const sheet = raw.getSheetIndex(source.sheet);
        const ref = typeof source.ref === 'string' ? parseReferenceOrNull(raw, source.ref) : null;
        if (ref !== null && ref.sheetName === null && sheet !== null && sheet !== undefined) {
          range = { sheet, rect: ref.rect };
        }
      }
      return { range, name: source.name ?? null };
    }

    // `null` removes the source; an object writes one with the attributes it gives (`null` is one left off).
    // `sheet` qualifies an unqualified `range`, and may not disagree with a qualified one.
    setWorksheetSource(source) {
      const where = 'PivotCache.setWorksheetSource';
      const raw = this.#raw();
      if (source === null) return raw.setPivotCacheWorksheetSource(this.#cacheId, null);
      const label = argLabel(where, 0);
      requireObject(source, label);
      refuseOtherKeys(label, source, ['range', 'sheet', 'name']);
      const input = {};
      let sheet = present(source.sheet) ? sheetRefIndex(this.#workbook, source.sheet, `${label}.sheet`) : undefined;
      if (present(source.range)) {
        let rect;
        let qualified;
        if (typeof source.range === 'string') {
          const [sheetName, parsed] = parsedRange(raw, source.range, `${label}.range`, 'SheetRect');
          rect = parsed;
          qualified = sheetName === null ? undefined : sheetRefIndex(this.#workbook, sheetName, `${label}.range`);
        } else {
          requireObject(source.range, `${label}.range`);
          ({ rect, sheet: qualified } = source.range);
        }
        if (qualified !== undefined && sheet !== undefined && qualified !== sheet) {
          throw new FormulonError(
            StatusCode.InvalidArgument,
            `${label}.sheet must match the sheet of range`,
            `sheet=${sheet}`,
          );
        }
        sheet ??= qualified;
        if (sheet === undefined) {
          throw new FormulonError(StatusCode.InvalidArgument, `${label}.range must name a sheet`, '');
        }
        input.ref = raw.formatReference({ rect });
      }
      if (sheet !== undefined) input.sheet = raw.getSheetName(sheet);
      if (present(source.name)) input.name = requireString(source.name, `${label}.name`);
      raw.setPivotCacheWorksheetSource(this.#cacheId, input);
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

    #locate() {
      return [rawOf(this.#sheet.workbook), sheetIndex(this.#sheet), this.#objectId];
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

    // Read from the sheet's drawing on each access.
    get name() {
      const [raw, sheet, id] = this.#locate();
      const record = raw.getDrawingObject(sheet, id);
      if (record === null) {
        throw new FormulonError(StatusCode.NotFound, `the sheet's drawing has no object ${id}`, `object_id=${id}`);
      }
      return record.name;
    }

    // The keys every kind takes; the kind-specific handles accept their own keys beside them.
    update(patch) {
      const [common, rest] = splitDrawingPatch('Drawing.update', patch);
      const [stray] = Object.keys(rest);
      if (stray !== undefined) throw new TypeError(`${argLabel('Drawing.update', 0)} has no field ${stray}`);
      const [raw, sheet, id] = this.#locate();
      raw.updateDrawing(sheet, id, rawDrawingPatch(common));
    }
  }

  class Picture extends Drawing {}

  class Chart extends Drawing {
    // The record replaces the legend, a key it leaves out taking Excel's default; `null` removes the legend.
    setLegend(legend) {
      const where = 'Chart.setLegend';
      const [raw, sheet, id] = drawingState.locate(this);
      if (legend === null) return raw.setChartLegend(sheet, id, null);
      checkPatch(where, legend, PATCH_FIELDS[where], {}, 'Chart.updateLegend');
      raw.setChartLegend(sheet, id, {
        position: legend.position ?? DEFAULT_LEGEND.position,
        overlay: legend.overlay ?? DEFAULT_LEGEND.overlay,
        hiddenEntries: legend.hiddenEntries ?? DEFAULT_LEGEND.hiddenEntries,
      });
    }

    // The record replaces the point's formatting (its `index` is ignored); `null` removes the formatting.
    setPoint(plot, seriesIndex, pointIndex, point) {
      const where = 'Chart.setPoint';
      const [raw, sheet, id] = drawingState.locate(this);
      if (point === null) return raw.setChartPoint(sheet, id, plot, seriesIndex, pointIndex, null);
      const label = argLabel(where, 3);
      requireObject(point, label);
      refuseOtherKeys(label, point, ['index', 'fill', 'line', 'marker']);
      raw.setChartPoint(sheet, id, plot, seriesIndex, pointIndex, colorInputs(point, label));
    }

    // `range` is a `SheetRect` or an A1 string (an unqualified one is on the chart's sheet); the raw call takes
    // the qualified text.
    setSource(source) {
      const where = 'Chart.setSource';
      const label = argLabel(where, 0);
      requireObject(source, label);
      refuseOtherKeys(label, source, ['range', 'direction', 'headerRow', 'headerColumn']);
      const [raw, sheet, id] = drawingState.locate(this);
      const flag = (field) =>
        source[field] === undefined || source[field] === null
          ? undefined
          : requireBoolean(source[field], `${label}.${field}`);
      const args = [
        sheetRectText(raw, source.range, sheet, `${label}.range`),
        source.direction === undefined
          ? ChartSeriesDirection.Auto
          : requireInteger(source.direction, `${label}.direction`),
      ];
      const headerRow = flag('headerRow');
      const headerColumn = flag('headerColumn');
      if (headerRow !== undefined || headerColumn !== undefined) args.push(headerRow, headerColumn);
      raw.setChartSource(sheet, id, ...args);
    }

    // A bar direction belongs to the `Bar` type alone, and only an explicit one is refused elsewhere.
    setType(plot, type, options) {
      const where = 'Chart.setType';
      const checked = chartTypeOptions(where, 2, type, options);
      const [raw, sheet, id] = drawingState.locate(this);
      raw.setChartType(sheet, id, plot, type, checked);
    }

    // The record names the series and categories the filter hides; a key it leaves out hides none.
    setFilter(filter) {
      const where = 'Chart.setFilter';
      const label = argLabel(where, 0);
      requireObject(filter, label);
      refuseOtherKeys(label, filter, ['hiddenSeries', 'hiddenCategories']);
      const indices = (key) => requireArray(filter[key] ?? [], `${label}.${key}`);
      const [raw, sheet, id] = drawingState.locate(this);
      raw.setChartFilter(sheet, id, indices('hiddenSeries'), indices('hiddenCategories'));
    }

    // `axis` is an axis id or the axis of a kind in a plot group (`plot` 0 by default).
    updateAxis(axis, patch) {
      const where = 'Chart.updateAxis';
      const [raw, sheet, id] = drawingState.locate(this);
      checkForwardedPatch(where, patch, PATCH_FIELDS[where]);
      raw.updateChartAxis(sheet, id, chartAxisId(raw, sheet, id, axis, argLabel(where, 0)), patch);
    }

    // The raw source carries qualified reference text, which reads as a `SheetRect` here.
    getProperties() {
      const [raw, sheet, id] = drawingState.locate(this);
      const properties = raw.getChart(sheet, id);
      if (properties === null) {
        throw new FormulonError(StatusCode.NotFound, 'the drawing object has no such part', `object_id=${id}`);
      }
      const { source } = properties;
      if (source === null || source === undefined) return properties;
      const range = parseSheetRect(raw, source.range);
      return { ...properties, source: range === null ? null : { ...source, range } };
    }

    // `overrides.range` is `source` text on the raw call.
    getPreview(overrides) {
      const where = 'Chart.getPreview';
      const label = argLabel(where, 0);
      requireObject(overrides, label);
      if (overrides.source !== undefined) throw new TypeError(`${label} has no field source`);
      const [raw, sheet, id] = drawingState.locate(this);
      const { range, ...rest } = colorInputs(overrides, label);
      const source = range === undefined ? undefined : sheetRectText(raw, range, sheet, `${label}.range`);
      const input = source === undefined ? rest : { ...rest, source };
      return raw.getChartPreview(sheet, id, input);
    }
  }

  // The common keys and the shape's own keys are checked, then applied by one raw call. `fill: null` /
  // `line: null` restore the paint the shape's style gives.
  class Shape extends Drawing {
    update(patch) {
      const where = 'Shape.update';
      patch = colorInputs(patch, argLabel(where, 0));
      const [common, rest] = splitDrawingPatch(where, patch);
      checkPatch(where, rest, PATCH_FIELDS[where], PATCH_NULLABLE[where]);
      const input = appliedFields(rest, PATCH_FIELDS[where], PATCH_NULLABLE[where]);
      const [raw, sheet, id] = drawingState.locate(this);
      raw.updateShape(sheet, id, { ...rawDrawingPatch(common), ...input });
    }
  }

  // As `Shape.update`; a `name` is refused by the engine, which keys the slicer part by it.
  class Slicer extends Drawing {
    update(patch) {
      const where = 'Slicer.update';
      const [common, rest] = splitDrawingPatch(where, patch);
      checkPatch(where, rest, PATCH_FIELDS[where], PATCH_NULLABLE[where]);
      const input = appliedFields(rest, PATCH_FIELDS[where], PATCH_NULLABLE[where]);
      const [raw, sheet, id] = drawingState.locate(this);
      raw.updateSlicer(sheet, id, { ...rawDrawingPatch(common), ...input });
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

  // The index a `SheetRef` of `workbook` names; `label` words the errors.
  function sheetRefIndex(workbook, ref, label) {
    if (ref instanceof Sheet) {
      if (ref.workbook !== workbook) {
        throw new FormulonError(StatusCode.InvalidArgument, `${label} must be a Sheet of this workbook`, '');
      }
      return sheetIndex(ref);
    }
    if (typeof ref === 'string') {
      const index = rawOf(workbook).getSheetIndex(ref);
      if (index === null || index === undefined) {
        throw new FormulonError(StatusCode.NotFound, `no sheet is named ${JSON.stringify(ref)}`, `sheet_name=${ref}`);
      }
      return index;
    }
    if (typeof ref === 'number') return requireIndex(ref, label, 'sheet');
    throw new TypeError(`${label} must be a sheet name, a sheet index or a Sheet`);
  }

  // The current index of a `TableRef` / `PivotRef` handle `value` passed to a method of the sheet `self`;
  // undefined for anything else.
  function handleIndex(self, kind, value, label) {
    if (kind === 'table' && value instanceof Table) {
      if (value.sheet !== self) {
        throw new FormulonError(StatusCode.InvalidArgument, `${label} must be a Table of this sheet`, '');
      }
      return value.index;
    }
    if (kind === 'pivot' && value instanceof Pivot) {
      if (value.sheet !== self) {
        throw new FormulonError(StatusCode.InvalidArgument, `${label} must be a Pivot of this sheet`, '');
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
    #tables = new Map();
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
        tableHandleAt: (sheet, index) => sheet.#tableAt(rawOf(sheet.#workbook), sheet.#current(), index),
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

    // `insertBefore` is the sheet the move lands before (a `SheetRef`), or `null` to append.
    move(insertBefore) {
      moveSheet(this.#workbook, this.#current(), insertBefore);
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

    // The tagged values of `range` (the used range when it is omitted or `null`) as rows, a blank cell `BLANK`.
    #typedGrid(range, where) {
      if (range === undefined || range === null) range = this.getUsedRange();
      if (range === null) return [];
      requireObject(range, argLabel(where, 0));
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const rows = range.lastRow - range.firstRow + 1;
      const cols = range.lastCol - range.firstCol + 1;
      if (rows * cols > MAX_VALUES_CELLS) {
        throw new FormulonError(
          StatusCode.InvalidArgument,
          `${argLabel(where, 0)} must name at most ${MAX_VALUES_CELLS} cells`,
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

    // The cells of `range` (the used range when it is omitted or `null`), one page at a time.
    getCellsInRange(range, options) {
      const paging = pageArguments('Sheet.getCellsInRange', 1, options);
      const rect = range ?? this.getUsedRange();
      if (rect === null) return { cells: [], nextCursor: null };
      const page = rawOf(this.#workbook).getCellsInRange(this.#current(), rect, ...paging);
      return { cells: page.cells.map(rangeCell), nextCursor: page.nextCursor };
    }

    getInvalidCells(options) {
      const paging = pageArguments('Sheet.getInvalidCells', 0, options);
      const page = rawOf(this.#workbook).getInvalidCells(this.#current(), ...paging);
      return { cells: page.cells.map(rangeCell), nextCursor: page.nextCursor };
    }

    listCells(range) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const rect = range ?? this.getUsedRange();
      if (rect === null) return [];
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
      requireObject(rect, argLabel('Sheet.iterRows', 0));
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
      const label = argLabel('Sheet.setTabColor', 0);
      const value = color === null || color === undefined ? color : formatColor(color, label);
      rawOf(this.#workbook).setTabColor(this.#current(), value);
    }

    listColumnOverrides() {
      return rawOf(this.#workbook)
        .listColumnOverrides(this.#current())
        .map(({ xfIndex, ...layout }) => layout);
    }

    listRowOverrides() {
      return rawOf(this.#workbook)
        .listRowOverrides(this.#current())
        .map(({ xfIndex, ...layout }) => layout);
    }

    // Each row is written in order; a failure leaves the cells before it written and names its cell.
    setValues(row, col, values) {
      const where = 'Sheet.setValues';
      requireInteger(row, argLabel(where, 0));
      requireInteger(col, argLabel(where, 1));
      const label = argLabel(where, 2);
      requireArray(values, label);
      const width = values.length === 0 ? 0 : requireArray(values[0], `${label}[0]`).length;
      for (const [i, line] of values.entries()) {
        requireArray(line, `${label}[${i}]`);
        if (line.length !== width) {
          throw new FormulonError(StatusCode.InvalidArgument, `${label}[${i}] must have ${width} elements`, `row=${i}`);
        }
        for (let j = 0; j < width; j++) {
          if (!(j in line) || line[j] === undefined) throw new TypeError(`${label}[${i}][${j}] is required`);
          checkCellValue(line[j], `${label}[${i}][${j}]`);
        }
      }
      if (width === 0) return;
      if (row < 0 || col < 0 || row + values.length > GRID_ROWS || col + width > GRID_COLS) {
        throw new FormulonError(
          StatusCode.InvalidArgument,
          `${label} ${values.length}x${width} at (${row}, ${col}) is outside the grid`,
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
      const { xfIndex, source } = raw.getEffectiveStyle(this.#current(), row, col);
      return cellFormatOf(raw, raw.getCellXf(xfIndex), source);
    }

    // Every argument is checked first; the sheet changes only in the final remap.
    updateCellFormat(range, patch) {
      const checked = checkCellFormatPatch(patch);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const bases = raw.getRangeXfBases(sheet, range);
      if (Object.keys(checked).length === 0) return;
      let numFmtId;
      if (checked.numberFormat === 'General') numFmtId = 0;
      else if (checked.numberFormat !== undefined) numFmtId = raw.addNumberFormat(checked.numberFormat);
      const xfs = bases.map((base) => raw.addCellXf(patchedXf(raw, raw.getCellXf(base), checked, numFmtId)));
      raw.remapRangeXfs(sheet, range, bases, xfs);
    }

    // The handle of the table at sheet-local `index` on the sheet `sheet` is.
    #tableAt(raw, sheet, index) {
      const id = raw.getTableRuntimeId(sheet, index);
      let table = this.#tables.get(id);
      if (table === undefined) {
        table = tableState.create(this, id);
        this.#tables.set(id, table);
      }
      return table;
    }

    listTables() {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const n = raw.getTableCount(sheet);
      const tables = [];
      for (let i = 0; i < n; i++) tables.push(this.#tableAt(raw, sheet, i));
      return tables;
    }

    // A position, a name matched case-insensitively among this sheet's tables, or the sheet's own handle.
    table(ref) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const named = (kind, value, label) => handleIndex(this, kind, value, label);
      const [index] = resolveRef(raw, sheet, 'Sheet.table', [0, 'table'], [ref], named);
      requireIndex(index, argLabel('Sheet.table', 0), 'position');
      if (index >= raw.getTableCount(sheet)) {
        throw new FormulonError(
          StatusCode.IndexOutOfRange,
          `${argLabel('Sheet.table', 0)} ${index} ${pastCount(raw.getTableCount(sheet), 'item')}`,
          `table_index=${index}`,
        );
      }
      return this.#tableAt(raw, sheet, index);
    }

    // Omitted `columns` reach the raw method empty, which takes the names from the header cells.
    addTable(input) {
      requireObject(input, argLabel('Sheet.addTable', 0));
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      if (typeof input.name === 'string' && namesACell(raw, input.name)) {
        throw new FormulonError(
          StatusCode.InvalidArgument,
          `${argLabel('Sheet.addTable', 0, 'name')} must not be a cell address`,
          `name=${input.name}`,
        );
      }
      const spec = input.columns === undefined ? { ...input, columns: [] } : input;
      return this.#tableAt(raw, sheet, raw.addTable(sheet, spec));
    }

    // An inline `format` is registered as a new differential format and reaches the raw method as its `dxfIndex`.
    addCfRule(input) {
      const where = 'Sheet.addCfRule';
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const label = argLabel(where, 0);
      input = colorInputs(input, label);
      if (isObject(input) && input.format !== undefined) {
        if (present(input.dxfIndex)) throw new TypeError(`${label}.format must not be given with dxfIndex`);
        const format = requireObject(input.format, `${label}.format`);
        refuseOtherKeys(`${label}.format`, format, DXF_INPUT_KEYS);
        const spec = readableCopy(input);
        delete spec.format;
        spec.dxfIndex = raw.addDxf(format);
        input = spec;
      }
      return raw.addCfRule(sheet, input);
    }

    // A key the rule's kind does not use, `format` included (a patch sets `dxfIndex`), is refused before anything
    // is applied.
    updateCfRule(cfRule, patch) {
      const where = 'Sheet.updateCfRule';
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      patch = colorInputs(patch, argLabel(where, 1));
      const [id] = resolveRef(raw, sheet, where, [0, 'cfRule'], [cfRule, patch]);
      refuseNullRanges(where, patch);
      if (isObject(patch)) {
        const { type } = this.#cfRuleById(raw, sheet, id, where);
        const used = CF_PATCH_KEYS[type] ?? [];
        // `null` on a key the rule's type does not use is "not specified", so a read record passes back.
        const specified = {};
        for (const [key, value] of Object.entries(patch)) {
          if (value === undefined || (value === null && !used.includes(key))) continue;
          if (!READ_ONLY_KEYS[where].includes(key) && !used.includes(key)) {
            throw new TypeError(`${argLabel(where, 1)} has no field ${key}`);
          }
          specified[key] = value;
        }
        patch = specified;
        if (typeof patch.type === 'number' && patch.type !== type) {
          throw new FormulonError(
            StatusCode.InvalidArgument,
            `${argLabel(where, 1, 'type')} must equal the type of the rule; remove the rule and add one`,
            `type=${patch.type}`,
          );
        }
        if (present(patch.ranges)) {
          patch = withRanges(raw, where, patch);
        }
      }
      this.#patch(where, patch, (raw, sheet, given) => raw.updateCfRule(sheet, id, given), undefined, ['ranges']);
    }

    // The rule whose id matches `id` as ids are matched everywhere; an unknown id is `NotFound`.
    #cfRuleById(raw, sheet, id, where) {
      const want = foldName(id);
      const n = raw.getCfRuleCount(sheet);
      for (let i = 0; i < n; i++) {
        const rule = raw.getCfRule(sheet, i);
        if (foldName(rule.id) === want) return rule;
      }
      throw new FormulonError(
        StatusCode.NotFound,
        `${where} found no rule with the ${REF_KINDS.cfRule.noun} ${JSON.stringify(id)}`,
        `${REF_KINDS.cfRule.context}=${id}`,
      );
    }

    // The raw call addresses a keyed element by its key: a position names the key read there.
    #idOf(ref, where, kind) {
      if (typeof ref === 'string') return ref;
      if (typeof ref !== 'number') {
        throw new TypeError(`${argLabel(where, 0)} must be a position or a ${REF_KINDS[kind].noun}`);
      }
      requireIndex(ref, argLabel(where, 0), 'position');
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const count = REF_KINDS[kind].count(raw, sheet);
      if (ref >= count) {
        throw new FormulonError(
          StatusCode.IndexOutOfRange,
          `${argLabel(where, 0)} ${ref} ${pastCount(count, 'item')}`,
          `position=${ref}`,
        );
      }
      return REF_KINDS[kind].key(raw, sheet, ref);
    }

    listValidations() {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const n = raw.getValidationCount(sheet);
      const out = [];
      for (let i = 0; i < n; i++) out.push(raw.getValidation(sheet, i));
      return out;
    }

    // A position reads the element there and a key the element it names; an unknown key reads `null`.
    getValidation(validation) {
      const id = this.#idOf(validation, 'Sheet.getValidation', 'validation');
      return rawOf(this.#workbook).getValidation(this.#current(), id);
    }

    getCfRule(cfRule) {
      const id = this.#idOf(cfRule, 'Sheet.getCfRule', 'cfRule');
      return rawOf(this.#workbook).getCfRule(this.#current(), id);
    }

    getProtectedRange(protectedRange) {
      const name = this.#idOf(protectedRange, 'Sheet.getProtectedRange', 'protectedRange');
      return rawOf(this.#workbook).getProtectedRange(this.#current(), name);
    }

    removeValidation(validation) {
      const id = this.#idOf(validation, 'Sheet.removeValidation', 'validation');
      rawOf(this.#workbook).removeValidation(this.#current(), id);
    }

    // `ranges` is an array of `CellRect` / A1 text, or one A1 string; `null` clears the formulas and texts the type
    // allows.
    updateValidation(validation, patch) {
      const where = 'Sheet.updateValidation';
      const raw = rawOf(this.#workbook);
      refuseNullRanges(where, patch);
      if (isObject(patch) && present(patch.ranges)) {
        patch = withRanges(raw, where, patch);
      }
      this.#patch(
        where,
        patch,
        (raw, sheet, given) => raw.updateValidation(sheet, this.#idOf(validation, where, 'validation'), given),
        undefined,
        ['ranges'],
      );
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
      requireIndex(index, argLabel('Sheet.pivot', 0), 'position');
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      if (index >= raw.getPivotCount(sheet)) {
        throw new FormulonError(
          StatusCode.IndexOutOfRange,
          `${argLabel('Sheet.pivot', 0)} ${index} ${pastCount(raw.getPivotCount(sheet), 'item')}`,
          `pivot_index=${index}`,
        );
      }
      return this.#pivotAt(raw, sheet, index);
    }

    addPivot(input) {
      const where = 'Sheet.addPivot';
      requireObject(input, argLabel(where, 0));
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const label = argLabel(where, 0);
      const anchor = addressOf(raw, input.anchor, `${label}.anchor`);
      let { cacheId } = input;
      if (cacheId instanceof PivotCache) {
        if (cacheId.workbook !== this.#workbook) {
          throw new FormulonError(
            StatusCode.InvalidArgument,
            `${label}.cacheId must be a PivotCache of this workbook`,
            '',
          );
        }
        cacheId = cacheId.cacheId;
      }
      const index = raw.addPivot(
        sheet,
        requireString(input.name, `${label}.name`),
        requireNumber(cacheId, `${label}.cacheId`),
        requireNumber(anchor.row, `${label}.anchor.row`),
        requireNumber(anchor.col, `${label}.anchor.col`),
      );
      return this.#pivotAt(raw, sheet, index);
    }

    // The handle of drawing object `objectId`, whose `DrawingProperties` is `record`.
    #drawingFor(raw, sheet, objectId, record) {
      let drawing = this.#drawings.get(objectId);
      if (drawing === undefined) {
        drawing = drawingState.create(drawingClass(raw, sheet, record), this, objectId, record.kind);
        this.#drawings.set(objectId, drawing);
      }
      return drawing;
    }

    // The object id a `number | Drawing` argument names.
    #drawingId(ref, where) {
      if (ref instanceof Drawing) {
        if (ref.sheet !== this) {
          throw new FormulonError(StatusCode.InvalidArgument, `${where} must be a Drawing of this sheet`, '');
        }
        return ref.objectId;
      }
      if (typeof ref !== 'number') throw new TypeError(`${where} must be a number or a Drawing`);
      return requireIndex(ref, where, 'quantity');
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
      const id = this.#drawingId(ref, argLabel(`Sheet.${method}`, 0));
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      return this.#drawingFor(raw, sheet, id, this.#drawingRecord(raw, sheet, id));
    }

    #typedDrawing(ref, cls, method) {
      const drawing = this.#lookupDrawing(ref, method);
      if (!(drawing instanceof cls)) {
        throw new FormulonError(
          StatusCode.InvalidArgument,
          `${argLabel(`Sheet.${method}`, 0)} ${drawing.objectId} is not a ${cls.name}`,
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
      return this.listDrawings().filter((drawing) => drawing instanceof Shape);
    }

    listSlicers() {
      return this.listDrawings().filter((drawing) => drawing instanceof Slicer);
    }

    addPicture(bytes, input) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const id = input === undefined ? raw.addImage(sheet, bytes) : raw.addImage(sheet, bytes, input);
      return this.#added(raw, sheet, id);
    }

    // `range` is a `SheetRect` or an A1 string (an unqualified one is on this sheet); the raw call takes the
    // qualified text as `source`.
    addChart(options) {
      const where = 'Sheet.addChart';
      const label = argLabel(where, 0);
      requireObject(options, label);
      if (options.source !== undefined) throw new TypeError(`${label} has no field source`);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      if (options.title === '') {
        throw new FormulonError(StatusCode.InvalidArgument, `${label}.title must not be empty`, 'argument=title');
      }
      const { range, ...rest } = colorInputs(options, label);
      if (present(rest.barDirection)) {
        chartTypeOptions(where, 0, rest.type, { barDirection: rest.barDirection });
      }
      const source = range === undefined ? undefined : sheetRectText(raw, range, sheet, `${label}.range`);
      const input = source === undefined ? rest : { ...rest, source };
      return this.#added(raw, sheet, raw.addChart(sheet, input));
    }

    addShape(spec) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      spec = colorInputs(spec, argLabel('Sheet.addShape', 0));
      return this.#added(raw, sheet, raw.addShape(sheet, spec));
    }

    // The table may live on another sheet, so the raw call takes its sheet and its sheet-local index. A string
    // `column` is the table's column name, matched as sheet names are.
    addSlicer(table, column, spec) {
      const where = 'Sheet.addSlicer';
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      let tableSheet;
      let tableIndex;
      if (typeof table === 'string') {
        const found = raw.findTable(table);
        if (found === null) {
          throw new FormulonError(
            StatusCode.NotFound,
            `no table has the table name ${JSON.stringify(table)}`,
            `table_name=${table}`,
          );
        }
        ({ sheet: tableSheet, index: tableIndex } = found);
      } else if (table instanceof Table) {
        if (table.workbook !== this.#workbook) {
          throw new FormulonError(
            StatusCode.InvalidArgument,
            `${argLabel(where, 0)} must be a Table of this workbook`,
            '',
          );
        }
        [, tableSheet, tableIndex] = tableState.locate(table);
      } else {
        throw new TypeError(`${argLabel(where, 0)} must be a table name or a Table`);
      }
      if (typeof column === 'string') {
        const want = foldName(column);
        const at = raw.getTable(tableSheet, tableIndex).columns.findIndex((c) => foldName(c.name) === want);
        if (at < 0) {
          throw new FormulonError(
            StatusCode.NotFound,
            `the table has no column named ${JSON.stringify(column)}`,
            `column_name=${column}`,
          );
        }
        column = at;
      }
      return this.#added(raw, sheet, raw.addSlicer(sheet, tableSheet, tableIndex, column, spec));
    }

    restoreDrawing(bytes, opts) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const id = opts === undefined ? raw.restoreDrawing(sheet, bytes) : raw.restoreDrawing(sheet, bytes, opts);
      return this.#added(raw, sheet, id);
    }

    // One raw remover takes an object of any kind.
    removeDrawing(objectId) {
      const drawing = this.#lookupDrawing(objectId, 'removeDrawing');
      return rawOf(this.#workbook).removeDrawingObject(this.#current(), drawing.objectId);
    }

    // The scroll origin and the pane origin read as cell addresses.
    getView() {
      const { topLeftRow, topLeftCol, paneTopLeftRow, paneTopLeftCol, ...view } = rawOf(this.#workbook).getView(
        this.#current(),
      );
      return {
        ...view,
        topLeft: { row: topLeftRow, col: topLeftCol },
        paneTopLeft: { row: paneTopLeftRow, col: paneTopLeftCol },
      };
    }

    // One raw call: the C layer merges the patch with the current view and refuses the whole patch together.
    updateView(patch) {
      const where = 'Sheet.updateView';
      const fields = PATCH_FIELDS[where];
      checkPatch(where, patch, fields);
      const raw = rawOf(this.#workbook);
      // `topLeft` and `paneTopLeft` reach the raw layer as their row / column pairs.
      patch = { ...patch };
      for (const [key, rowKey, colKey] of [
        ['topLeft', 'topLeftRow', 'topLeftCol'],
        ['paneTopLeft', 'paneTopLeftRow', 'paneTopLeftCol'],
      ]) {
        const value = patch[key];
        delete patch[key];
        if (value === null) throw new TypeError(`${where} ${key} must be a CellAddress or an A1 string`);
        if (value === undefined) continue;
        const cell = addressOf(raw, value, `${where} ${key}`);
        patch[rowKey] = requireIndex(cell.row, `${where} ${key}.row`, 'coordinate');
        patch[colKey] = requireIndex(cell.col, `${where} ${key}.col`, 'coordinate');
      }
      raw.updateView(this.#current(), patch);
    }

    updatePrintOptions(patch) {
      const where = 'Sheet.updatePrintOptions';
      checkPatch(where, patch, PATCH_FIELDS[where]);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      raw.updatePrintOptions(
        sheet,
        withoutUnstated(patch, () => raw.getPrintOptions(sheet)),
      );
    }

    // `null` and `''` both clear a section.
    updateHeaderFooter(patch) {
      const where = 'Sheet.updateHeaderFooter';
      const sections = ['oddHeader', 'oddFooter', 'evenHeader', 'evenFooter', 'firstHeader', 'firstFooter'];
      const nullable = Object.fromEntries(sections.map((key) => [key, 'string']));
      checkPatch(where, patch, PATCH_FIELDS[where], nullable);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      raw.updateHeaderFooter(
        sheet,
        withoutUnstated(patch, () => raw.getHeaderFooter(sheet)),
      );
    }

    getPrintArea() {
      const areas = rawOf(this.#workbook).getPrintArea(this.#current());
      return areas.length === 0 ? null : areas;
    }

    getColumnWidthPt(col, mode) {
      mode = geometryMode(mode, 'Sheet.getColumnWidthPt', 1);
      return rawOf(this.#workbook).getColumnWidthPt(this.#current(), col, mode);
    }

    convertColumnCharsToPt(chars, mode) {
      mode = geometryMode(mode, 'Sheet.convertColumnCharsToPt', 1);
      return rawOf(this.#workbook).convertColumnCharsToPt(this.#current(), mode, chars);
    }

    convertColumnPtToChars(pt, mode) {
      mode = geometryMode(mode, 'Sheet.convertColumnPtToChars', 1);
      return rawOf(this.#workbook).convertColumnPtToChars(this.#current(), mode, pt);
    }

    getCellRectPt(range, mode) {
      mode = geometryMode(mode, 'Sheet.getCellRectPt', 1);
      return rawOf(this.#workbook).getCellRectPt(this.#current(), range, mode);
    }

    getWidthModel(mode) {
      mode = geometryMode(mode, 'Sheet.getWidthModel', 0);
      return rawOf(this.#workbook).getWidthModel(this.#current(), mode);
    }

    addSparklineGroup(spec) {
      const where = 'Sheet.addSparklineGroup';
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      spec = colorInputs(spec, argLabel(where, 0));
      requireObject(spec, argLabel(where, 0));
      return raw.addSparklineGroup(sheet, encodeSparklineGroup(raw, sheet, spec, where));
    }

    listSparklineGroups() {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const n = raw.getSparklineGroupCount(sheet);
      const groups = [];
      for (let i = 0; i < n; i++) groups.push(decodeSparklineGroup(raw, raw.getSparklineGroup(sheet, i)));
      return groups;
    }

    getSparklineGroup(sparklineGroup) {
      const raw = rawOf(this.#workbook);
      const id = this.#idOf(sparklineGroup, 'Sheet.getSparklineGroup', 'sparklineGroup');
      const group = raw.getSparklineGroup(this.#current(), id);
      return group === null ? null : decodeSparklineGroup(raw, group);
    }

    removeSparklineGroup(sparklineGroup) {
      const id = this.#idOf(sparklineGroup, 'Sheet.removeSparklineGroup', 'sparklineGroup');
      rawOf(this.#workbook).removeSparklineGroup(this.#current(), id);
    }

    // The raw update merges the patch with the stored group; without `sparklines` it keeps the group's own.
    updateSparklineGroup(sparklineGroup, patch) {
      const where = 'Sheet.updateSparklineGroup';
      patch = colorInputs(patch, argLabel(where, 1));
      checkSparklinePatch(where, patch);
      // `dateAxisRange` takes a record or text, which the one type word of the table cannot say.
      const range = isObject(patch) ? patch.dateAxisRange : undefined;
      if (range !== undefined && range !== null && typeof range !== 'string' && !isObject(range)) {
        throw new TypeError(`${argLabel(where, 1, 'dateAxisRange')} must be a SheetRect, an A1 string or null`);
      }
      this.#patch(
        where,
        patch,
        (raw, sheet, given) =>
          raw.updateSparklineGroup(sheet, this.#idOf(sparklineGroup, where, 'sparklineGroup'), given),
        (raw, sheet, given) => encodeSparklineGroup(raw, sheet, given, where),
        ['dateAxisRange'],
      );
    }

    // As `updateSparklineGroup`; without `ranges` the raw call keeps the stored ranges, unparsed ones included.
    // `ranges` is an array of `CellRect` / A1 text, or one A1 string.
    updateProtectedRange(protectedRange, patch) {
      const where = 'Sheet.updateProtectedRange';
      const raw = rawOf(this.#workbook);
      if (isObject(patch) && present(patch.ranges)) {
        patch = withRanges(raw, where, patch);
      }
      this.#patch(where, patch, (raw, sheet, given) => raw.updateProtectedRange(sheet, protectedRange, given));
    }

    // Checks the patch of `where` and hands its given keys to `write`; the raw call merges them with the stored
    // record. `encode` turns the given keys into the raw record's; `loose` names keys the caller has typed, which
    // `checkPatch` leaves alone.
    #patch(where, patch, write, encode = (_raw, _sheet, given) => given, loose = []) {
      const fields = PATCH_FIELDS[where];
      const nullable = PATCH_NULLABLE[where];
      const untyped = Object.fromEntries(loose.map((key) => [key, undefined]));
      checkPatch(where, isObject(patch) ? { ...patch, ...untyped } : patch, fields, nullable);
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const typed = { ...nullable, ...Object.fromEntries(loose.map((key) => [key, 'object'])) };
      write(raw, sheet, encode(raw, sheet, appliedFields(patch, fields, typed)));
    }

    // `range` and an input cell are `CellRect` / `CellAddress` or A1 text; the raw call takes `rect`.
    addDataTable(input) {
      const where = 'Sheet.addDataTable';
      const label = argLabel(where, 0);
      requireObject(input, label);
      const raw = rawOf(this.#workbook);
      const range = typeof input.range === 'string' ? rangeAddress(raw, input.range, `${label}.range`) : input.range;
      const spec = { rect: requireObject(range, `${label}.range`) };
      for (const key of ['rowInput', 'columnInput']) {
        if (!present(input[key])) continue;
        const { row, col } = addressOf(raw, input[key], `${label}.${key}`);
        spec[key] = { row, col };
      }
      return raw.addDataTable(this.#current(), spec);
    }

    // `range` is a `CellRect` or A1 text (expanded before this runs); the merged range comes back as stored.
    addMerge(range) {
      return rawOf(this.#workbook).addMerge(this.#current(), range);
    }

    setValue(row, col, value) {
      const valueLabel = argLabel('Sheet.setValue', 2);
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
              return this.setNumber(row, col, requireValueField(value, 'number', 'number', valueLabel));
            case ValueKind.Bool:
              return this.setBool(row, col, requireValueField(value, 'boolean', 'boolean', valueLabel));
            case ValueKind.Text:
              return this.setText(row, col, requireValueField(value, 'text', 'string', valueLabel));
            case ValueKind.Error:
              return this.setError(row, col, requireValueField(value, 'errorCode', 'number', valueLabel));
          }
          break;
      }
      throw new TypeError(`${valueLabel} must be a cell value`);
    }

    // The serial of the date's local fields; a General cell also gets the built-in date format.
    #setDate(row, col, date) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const [serial, timed] = dateSerial(runtime, date, raw.getDate1904(), argLabel('Sheet.setValue', 2));
      this.setNumber(row, col, serial);
      const { xfIndex } = raw.getEffectiveStyle(sheet, row, col);
      if (raw.getCellXf(xfIndex).numFmtId !== 0) return;
      const rect = { firstRow: row, firstCol: col, lastRow: row, lastCol: col };
      const patch = { numberFormat: timed ? DATE_TIME_FORMAT : DATE_FORMAT };
      const bases = raw.getRangeXfBases(sheet, rect);
      const xfs = bases.map((base) =>
        raw.addCellXf(patchedXf(raw, raw.getCellXf(base), patch, BUILTIN_NUM_FMT_IDS[patch.numberFormat])),
      );
      raw.remapRangeXfs(sheet, rect, bases, xfs);
    }

    // An omitted or `null` `author` means none; `null` removes the comment and does nothing when there is none.
    setComment(row, col, comment) {
      const where = 'Sheet.setComment';
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      if (comment === null) return raw.setComment(sheet, row, col, null);
      const label = argLabel(where, 2);
      requireObject(comment, label);
      const author = comment.author == null ? '' : requireString(comment.author, `${label}.author`);
      const text = requireString(comment.text, `${label}.text`);
      raw.setComment(sheet, row, col, { author, text });
    }

    // `null` removes the hyperlink anchored at the cell and does nothing when there is none.
    setHyperlink(row, col, link) {
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      raw.setHyperlink(sheet, row, col, link === null ? null : hyperlinkInput(raw, link, 'Sheet.setHyperlink', 2));
    }

    // Replaces nothing; the anchor is the top-left of `link.range`, which is therefore required.
    insertHyperlinkAt(index, link) {
      const where = 'Sheet.insertHyperlinkAt';
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      const at = requireIndex(index, argLabel(where, 0), 'position');
      const input = hyperlinkInput(raw, link, where, 1);
      if (!present(input.range)) {
        throw new FormulonError(
          StatusCode.InvalidArgument,
          `${argLabel(where, 1)}.range is required`,
          'link.range=undefined',
        );
      }
      raw.insertHyperlinkAt(sheet, at, input);
    }

    insertCells(range, shift) {
      this.#editCells('Sheet.insertCells', 'insertCells', range, shift);
    }

    deleteCells(range, shift) {
      this.#editCells('Sheet.deleteCells', 'deleteCells', range, shift);
    }

    #editCells(where, rawName, range, shift) {
      const raw = rawOf(this.#workbook);
      const rect = normalizeAddress(raw, 'CellRect', range, undefined, argLabel(where, 0));
      requireObject(rect, argLabel(where, 0));
      requireInteger(shift, argLabel(where, 1));
      raw[rawName](this.#current(), rect, shift);
    }

    setRowBreak(row, input) {
      this.#setBreak('Sheet.setRowBreak', 'setRowBreak', row, input);
    }

    setColumnBreak(col, input) {
      this.#setBreak('Sheet.setColumnBreak', 'setColumnBreak', col, input);
    }

    // `input` is a `PageBreakInput` or the `RowBreak` / `ColumnBreak` read back, whose `row` / `col` is ignored;
    // `null` removes the break.
    #setBreak(where, rawName, position, input) {
      let record = null;
      if (input !== null) {
        const label = argLabel(where, 1);
        requireObject(input, label);
        refuseOtherKeys(label, input, [...READ_ONLY_KEYS[where], 'first', 'last', 'manual']);
        record = { ...input };
        for (const read of READ_ONLY_KEYS[where]) delete record[read];
      }
      rawOf(this.#workbook)[rawName](this.#current(), position, record);
    }

    // `null` removes the protection; a record sets the whole protection.
    setSheetProtection(protection) {
      const where = 'Sheet.setSheetProtection';
      const source = 'Sheet.updateSheetProtection';
      if (protection !== null) {
        const fields = { ...PATCH_FIELDS[source], spinCount: 'number' };
        const nullable = strings('algorithmName', 'hashValue', 'saltValue', 'legacyPasswordHash');
        checkPatch(where, protection, fields, nullable, source);
      }
      rawOf(this.#workbook).setSheetProtection(this.#current(), protection);
    }

    // `cell` is a `CellAddress` or A1 text, `created` a `DateInput` or `null`; the raw record carries the row and
    // column and the C stamp.
    addThreadedComment(comment) {
      const where = 'Sheet.addThreadedComment';
      requireObject(comment, argLabel(where, 0));
      const { cell, created, ...rest } = comment;
      const input = { ...rest };
      if (present(cell)) {
        const { row, col } = addressOf(rawOf(this.#workbook), cell, argLabel(where, 0, 'cell'));
        input.row = row;
        input.col = col;
      }
      if (present(created)) {
        input.created = civilStamp(toCivil(created, where, argumentPath(where, 0, 'created'), { utc: true }));
      }
      return rawOf(this.#workbook).addThreadedComment(this.#current(), input);
    }

    // With a cell (A1 text, or `row` and `col`) only the threads anchored there: the opener and its replies.
    listThreadedComments(...cell) {
      const where = 'Sheet.listThreadedComments';
      const raw = rawOf(this.#workbook);
      const comments = raw.listThreadedComments(this.#current()).map(threadedCommentRead);
      if (cell.length === 0 || cell[0] === undefined) return comments;
      const at =
        typeof cell[0] === 'string'
          ? cellAddress(raw, cell[0], `${where} a1`)
          : {
              firstRow: requireIndex(cell[0], `${where} row`, 'coordinate'),
              firstCol: requireIndex(cell[1], `${where} col`, 'coordinate'),
            };
      const ids = new Set();
      return comments.filter((comment) => {
        const reply = present(comment.parentId) && comment.parentId !== '';
        const kept = reply
          ? ids.has(comment.parentId)
          : comment.cell.row === at.firstRow && comment.cell.col === at.firstCol;
        if (kept) ids.add(comment.id);
        return kept;
      });
    }

    getThreadedComment(id) {
      const comment = rawOf(this.#workbook).getThreadedComment(this.#current(), id);
      return comment === null ? null : threadedCommentRead(comment);
    }

    // `today` is a `DateInput`; omitted, the engine measures from the pinned clock, else the current date.
    evaluateCfRange(range, options) {
      const where = 'Sheet.evaluateCfRange';
      const raw = rawOf(this.#workbook);
      const input = {};
      let timePeriodRules = true;
      if (options !== undefined) {
        const label = argLabel(where, 1);
        requireObject(options, label);
        refuseOtherKeys(label, options, ['today', 'timePeriodRules']);
        if (options.timePeriodRules !== undefined) {
          timePeriodRules = requireBoolean(options.timePeriodRules, `${label}.timePeriodRules`);
        }
        if (options.today !== undefined) {
          if (!timePeriodRules) {
            throw new FormulonError(
              StatusCode.InvalidArgument,
              `${label}.today must be omitted when timePeriodRules is false`,
              'option=today',
            );
          }
          const today = toCivil(options.today, where, argumentPath(where, 1, 'today'));
          const civil = { ...today, hour: 0, minute: 0, second: 0 };
          input.todaySerial = civilSerial(runtime, civil, raw.getDate1904());
        }
      }
      input.timePeriodRules = timePeriodRules;
      return raw.evaluateCfRange(this.#current(), range, input);
    }

    // An `ExcelError` names its `ErrorCode`.
    setError(row, col, error) {
      const code = error instanceof ExcelError ? error.code : error;
      rawOf(this.#workbook).setError(this.#current(), row, col, code);
    }

    // The candidate is coerced as `setValue` coerces a value, except that a string is always text.
    validateValue(row, col, value) {
      const where = 'Sheet.validateValue';
      const raw = rawOf(this.#workbook);
      const proposed = valueRecord(runtime, raw, value, argLabel(where, 2));
      return raw.validateValue(this.#current(), row, col, proposed);
    }

    // Only the keys the patch gives reach the engine; a record's other fields are ignored.
    updateThreadedComment(id, patch) {
      const where = 'Sheet.updateThreadedComment';
      requireObject(patch, argLabel(where, 1));
      refuseUnknownKeys(where, patch);
      const input = {};
      const checks = [
        ['text', requireString],
        ['mentions', requireArray],
        ['done', requireBoolean],
      ];
      for (const [key, check] of checks) {
        if (patch[key] !== undefined) input[key] = check(patch[key], argLabel(where, 1, key));
      }
      rawOf(this.#workbook).updateThreadedComment(this.#current(), id, input);
    }

    // `range` is a `CellRect` or an A1 range of this sheet.
    setArrayFormula(arrayFormula) {
      const where = 'Sheet.setArrayFormula';
      const label = argLabel(where, 0);
      requireObject(arrayFormula, label);
      const raw = rawOf(this.#workbook);
      const { range } = arrayFormula;
      const rect =
        typeof range === 'string' ? rangeAddress(raw, range, `${label}.range`) : requireObject(range, `${label}.range`);
      const formula = requireString(arrayFormula.formula, `${label}.formula`);
      raw.setArrayFormula(this.#current(), rect, formula);
    }

    // Each axis is a `Span` or an A1 span (`'1:2'`, `'$3'`, `'B'`, `'A:B'`), or `null` for none.
    setPrintTitles(titles) {
      const where = 'Sheet.setPrintTitles';
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      if (titles === null) return raw.setPrintTitles(sheet, null);
      requireObject(titles, argLabel(where, 0));
      const span = (axis, key) => {
        const value = titles[key];
        if (typeof value !== 'string') return value;
        try {
          const [first, last] = spanAddress(raw, value, axis, argLabel(where, 0, key));
          return { first, last };
        } catch (e) {
          if (e?.name !== 'FormulonError' || e.parseError === undefined) throw e;
          throw new FormulonError(StatusCode.InvalidArgument, `${argLabel(where, 0, key)} must be an A1 span`, '');
        }
      };
      raw.setPrintTitles(sheet, { ...titles, rows: span('row', 'rows'), columns: span('col', 'columns') });
    }

    updatePageSetup(setup) {
      const where = 'Sheet.updatePageSetup';
      const raw = rawOf(this.#workbook);
      const sheet = this.#current();
      checkForwardedPatch(where, setup, PATCH_FIELDS[where]);
      raw.updatePageSetup(
        sheet,
        withoutUnstated(setup, () => raw.getPageSetup(sheet)),
      );
    }

    // `runs` is an array of `{ start, end, text }` records, all three keys required, or `null` for none.
    setCellPhoneticRuns(row, col, runs) {
      const where = 'Sheet.setCellPhoneticRuns';
      const label = argLabel(where, 2);
      const given = runs === null ? [] : requireArray(runs, label);
      const checked = given.map((run, i) => {
        const runLabel = `${label}[${i}]`;
        requireObject(run, runLabel);
        refuseOtherKeys(runLabel, run, ['start', 'end', 'text']);
        for (const key of ['start', 'end', 'text']) {
          if (run[key] === undefined) throw new TypeError(`${runLabel}.${key} is required`);
        }
        return {
          start: requireInteger(run.start, `${runLabel}.start`),
          end: requireInteger(run.end, `${runLabel}.end`),
          text: requireString(run.text, `${runLabel}.text`),
        };
      });
      rawOf(this.#workbook).setCellPhoneticRuns(this.#current(), row, col, checked);
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

  // The raw record of a hyperlink: the text fields checked, an A1 `range` as a `CellRect`.
  function hyperlinkInput(raw, link, where, position) {
    const label = argLabel(where, position);
    requireObject(link, label);
    for (const field of ['target', 'location', 'display', 'tooltip']) {
      if (present(link[field])) requireString(link[field], `${label}.${field}`);
    }
    if (!present(link.target) && !present(link.location)) {
      throw new TypeError(`${label}.location is required when target is null`);
    }
    if (link.range === undefined || link.range === null) return link;
    return { ...link, range: normalizeAddress(raw, 'CellRect', link.range, undefined, `${label}.range`) };
  }

  class Workbook {
    #handle;
    #sheets;
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
      moveSheet = (workbook, from, insertBefore) => workbook.#move(from, insertBefore);
      pivotCacheOf = (workbook, cacheId) => workbook.#cacheHandle(cacheId);
    }

    static load(bytes) {
      return new Workbook(token, Raw.load(requireBytes(bytes, 'Workbook.load', 'bytes')));
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

    // `insertBefore` is a `SheetRef` of the list before the move, or `null` for the end.
    #move(from, insertBefore) {
      const before =
        insertBefore === null
          ? this.#sheets.length
          : this.#insertBefore(insertBefore, 'move', argumentPath('Sheet.move', 0), 'Sheet');
      this.#raw.moveSheet(from, before);
      const to = before > from ? before - 1 : before;
      if (to === from) return;
      const [moved] = this.#sheets.splice(from, 1);
      this.#sheets.splice(to, 0, moved);
      this.#renumber();
    }

    // The position an `insertBefore` option names: a `SheetRef`, or `null` / omitted for the end.
    #insertPosition(options, method, position) {
      const where = `Workbook.${method}`;
      if (options === undefined) return this.#sheets.length;
      const label = argLabel(where, position);
      requireObject(options, label);
      refuseOtherKeys(label, options, method === 'addSheet' ? ['insertBefore'] : ['insertBefore', 'newName']);
      const { insertBefore } = options;
      return insertBefore === undefined || insertBefore === null
        ? this.#sheets.length
        : this.#insertBefore(insertBefore, method, argumentPath(where, position, 'insertBefore'));
    }

    // A `SheetRef` the sheet lands before, where the sheet count (one past the last index) appends.
    #insertBefore(ref, method, path, owner = 'Workbook') {
      if (typeof ref === 'number' && ref === this.#sheets.length) {
        this.#alive();
        return ref;
      }
      return this.#resolve(ref, method, path, owner);
    }

    // Resolves a `SheetRef` the way `sheet(key)` does; `owner`, `method` and `path` (that of the argument or record
    // field holding `ref`) word the errors.
    #resolve(ref, method, path, owner = 'Workbook') {
      const where = `${owner}.${method} ${path}`;
      if (ref instanceof Sheet) {
        if (ref.workbook !== this) {
          throw new FormulonError(StatusCode.InvalidArgument, `${where} must be a Sheet of this workbook`, '');
        }
        return ref.index;
      }
      if (typeof ref === 'string') {
        this.#alive();
        const index = this.#raw.getSheetIndex(ref);
        if (index === null) {
          throw new FormulonError(
            StatusCode.NotFound,
            `${where} ${JSON.stringify(ref)} names no sheet`,
            `sheet_name=${ref}`,
          );
        }
        return index;
      }
      if (typeof ref === 'number') {
        requireIndex(ref, where, 'sheet');
        this.#alive();
        if (ref >= this.#sheets.length) {
          throw new FormulonError(
            StatusCode.SheetIndexOutOfRange,
            `${where} ${ref} ${pastCount(this.#sheets.length, 'sheet')}`,
            `sheet_index=${ref}`,
          );
        }
        return ref;
      }
      throw new TypeError(`${where} must be a sheet name, a sheet index or a Sheet`);
    }

    sheet(key) {
      return this.#sheets[this.#resolve(key, 'sheet', argumentPath('Workbook.sheet', 0))];
    }

    listSheets() {
      this.#alive();
      return [...this.#sheets];
    }

    get sheetCount() {
      this.#alive();
      return this.#sheets.length;
    }

    // A well-typed `SheetRef` that names no sheet of this workbook is `false`; any other argument is a `TypeError`.
    hasSheet(ref) {
      this.#alive();
      if (ref instanceof Sheet) return this.#sheets.includes(ref);
      if (typeof ref === 'string') return this.#raw.getSheetIndex(ref) !== null;
      if (typeof ref === 'number') {
        if (!Number.isInteger(ref)) throw new RangeError(`${argLabel('Workbook.hasSheet', 0)} must be an integer`);
        return ref >= 0 && ref < this.#sheets.length;
      }
      throw new TypeError(`${argLabel('Workbook.hasSheet', 0)} must be a sheet name, a sheet index or a Sheet`);
    }

    activeSheet() {
      this.#alive();
      return this.#sheets[this.#raw.activeSheet()];
    }

    setActiveSheet(sheet) {
      this.#raw.setActiveSheet(this.#resolve(sheet, 'setActiveSheet', argumentPath('Workbook.setActiveSheet', 0)));
    }

    // An omitted or `null` name takes the profile's next default sheet name.
    addSheet(name, options) {
      const where = 'Workbook.addSheet';
      if (name !== undefined && name !== null) requireString(name, argLabel(where, 0));
      const before = this.#insertPosition(options, 'addSheet', 1);
      return this.#insert(this.#raw.addSheet(name ?? null, before));
    }

    copySheet(source, options) {
      const index = this.#resolve(source, 'copySheet', argumentPath('Workbook.copySheet', 0));
      const before = this.#insertPosition(options, 'copySheet', 1);
      const newName = options?.newName;
      if (newName !== undefined) requireString(newName, argLabel('Workbook.copySheet', 1, 'newName'));
      const position =
        newName === undefined ? this.#raw.copySheet(index, before) : this.#raw.copySheet(index, before, newName);
      return this.#insert(position);
    }

    removeSheet(sheet) {
      const index = this.#resolve(sheet, 'removeSheet', argumentPath('Workbook.removeSheet', 0));
      this.#raw.removeSheet(index);
      const [removed] = this.#sheets.splice(index, 1);
      sheetState.place(removed, -1);
      this.#renumber();
    }

    // The one workbook-wide table lookup: tables are added, listed and removed on their `Sheet`.
    table(name) {
      requireString(name, argLabel('Workbook.table', 0));
      this.#alive();
      const found = this.#raw.findTable(name);
      if (found === null) {
        throw new FormulonError(
          StatusCode.NotFound,
          `no table has the table name ${JSON.stringify(name)}`,
          `table_name=${name}`,
        );
      }
      return sheetState.tableHandleAt(this.#sheets[found.sheet], found.index);
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
          throw new FormulonError(StatusCode.InvalidArgument, `${where} must be a PivotCache of this workbook`, '');
        }
        return ref.cacheId;
      }
      if (typeof ref !== 'number') throw new TypeError(`${where} must be a number or a PivotCache`);
      return requireIndex(ref, where, 'quantity');
    }

    listPivotCaches() {
      this.#alive();
      return this.#raw.listPivotCaches().map((cacheId) => this.#cacheHandle(cacheId));
    }

    pivotCache(cacheId) {
      const id = this.#cacheId(cacheId, argLabel('Workbook.pivotCache', 0));
      this.#alive();
      if (!this.#raw.listPivotCaches().includes(id)) {
        throw new FormulonError(StatusCode.NotFound, `the workbook has no pivot cache ${id}`, `cache_id=${id}`);
      }
      return this.#cacheHandle(id);
    }

    addPivotCache(requestedId) {
      let requested = 0;
      if (requestedId !== undefined && requestedId !== null) {
        requireNumber(requestedId, argLabel('Workbook.addPivotCache', 0));
        if (requestedId === 0) {
          throw new FormulonError(
            StatusCode.InvalidArgument,
            `${argLabel('Workbook.addPivotCache', 0)} must not be 0; omit it to assign an id`,
            'requested_id=0',
          );
        }
        requested = requestedId;
      }
      return this.#cacheHandle(this.#raw.addPivotCache(requested));
    }

    removePivotCache(cacheId) {
      this.#raw.removePivotCache(this.#cacheId(cacheId, argLabel('Workbook.removePivotCache', 0)));
    }

    setPinnedNow(time) {
      if (time === null) return this.#raw.setPinnedNow(null);
      const where = 'Workbook.setPinnedNow';
      const civil = toCivil(time, where, argumentPath(where, 0));
      return this.#raw.setPinnedNow(...CIVIL_FIELDS.map((field) => civil[field]));
    }

    // The value is coerced as `Sheet.setValue` coerces one (a string is always text), a `Date` read under this
    // workbook's date system.
    formatValue(value, numberFormat) {
      const raw = this.#raw;
      const label = argLabel('Workbook.formatValue', 0);
      if (value === undefined || Array.isArray(value)) throw new TypeError(`${label} must be a cell value`);
      return raw.formatValue(valueRecord(runtime, raw, value, label), numberFormat);
    }

    // `range` is an A1 range, a `CellRect` or a `ParsedReference`, whose sheet name `options.sheetName` replaces.
    formatReference(range, options) {
      const where = 'Workbook.formatReference';
      const raw = this.#raw;
      const label = argLabel(where, 0);
      let rect = range;
      let sheetName = null;
      if (isObject(range) && Object.hasOwn(range, 'range')) {
        rect = range.range;
        sheetName = range.sheetName ?? null;
      }
      if (typeof rect === 'string') rect = rangeAddress(raw, rect, `${label}.range`);
      requireObject(rect, label);
      let absolute = false;
      if (options !== undefined) {
        const optionsLabel = argLabel(where, 1);
        requireObject(options, optionsLabel);
        refuseOtherKeys(optionsLabel, options, ['sheetName', 'absolute']);
        if (options.absolute !== undefined) absolute = requireBoolean(options.absolute, `${optionsLabel}.absolute`);
        if (options.sheetName !== undefined) {
          sheetName = options.sheetName === null ? null : requireString(options.sheetName, `${optionsLabel}.sheetName`);
        }
      }
      const reference = sheetName === null || sheetName === '' ? { rect } : { sheetName, rect };
      return absolute ? raw.formatReference(reference, { absolute }) : raw.formatReference(reference);
    }

    getThemeFonts() {
      return this.#raw.getThemeFonts();
    }

    getThemeColors() {
      return this.#raw.getThemeColors();
    }

    setThemeColors(colors) {
      const where = 'Workbook.setThemeColors';
      requireObject(colors, argLabel(where, 0));
      const input = colorInputs(colors, argLabel(where, 0));
      this.#raw.setThemeColors(requireArray(input.colors, argLabel(where, 0, 'colors')));
    }

    getDefaultFont() {
      return this.#raw.getDefaultFont();
    }

    setThemeScriptFont(font) {
      const where = 'Workbook.setThemeScriptFont';
      const label = argLabel(where, 0);
      requireObject(font, label);
      refuseOtherKeys(label, font, ['scheme', 'script', 'typeface']);
      this.#raw.setThemeScriptFont(
        requireInteger(font.scheme, `${label}.scheme`),
        requireString(font.script, `${label}.script`),
        requireString(font.typeface, `${label}.typeface`),
      );
    }

    getThemeScriptFont(key) {
      const where = 'Workbook.getThemeScriptFont';
      const label = argLabel(where, 0);
      requireObject(key, label);
      refuseOtherKeys(label, key, ['scheme', 'script']);
      return this.#raw.getThemeScriptFont(
        requireInteger(key.scheme, `${label}.scheme`),
        requireString(key.script, `${label}.script`),
      );
    }

    // A `null` or omitted scope is workbook scope; the raw sentinel is -1.
    #scope(scope, method) {
      return scope === undefined || scope === null
        ? -1
        : this.#resolve(scope, method, argumentPath(`Workbook.${method}`, 1));
    }

    getDefinedName(name, scope) {
      return this.#raw.getDefinedName(name, this.#scope(scope, 'getDefinedName'));
    }

    removeDefinedName(name, scope) {
      return this.#raw.removeDefinedName(name, this.#scope(scope, 'removeDefinedName'));
    }

    setDefinedName(definedName) {
      requireObject(definedName, argLabel('Workbook.setDefinedName', 0));
      const { scope, ...fields } = definedName;
      if (scope === undefined || scope === null) return this.#raw.setDefinedName(fields);
      const path = argumentPath('Workbook.setDefinedName', 0, 'scope');
      return this.#raw.setDefinedName({ ...fields, scope: this.#resolve(scope, 'setDefinedName', path) });
    }

    updateIterative(patch) {
      const where = 'Workbook.updateIterative';
      checkForwardedPatch(where, patch, PATCH_FIELDS[where]);
      this.#raw.updateIterative(patch);
    }
  }

  const DXF_INPUT_KEYS = Object.freeze(['font', 'fill', 'border', 'numberFormat', 'alignment', 'protection']);

  class WorkbookOoxml {
    #workbook;

    constructor(key, workbook) {
      if (key !== token) throw new TypeError('WorkbookOoxml has no public constructor; use Workbook.ooxml');
      this.#workbook = workbook;
    }

    addDxf(input) {
      const where = 'Workbook.addDxf';
      const label = argLabel(where, 0);
      requireObject(input, label);
      refuseOtherKeys(label, input, DXF_INPUT_KEYS);
      return rawOf(this.#workbook).addDxf(colorInputs(input, label));
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

    getColumnXfIndex(col) {
      requireIndex(col, argLabel('Sheet.getColumnXfIndex', 0), 'coordinate');
      const raw = rawOf(this.#sheet.workbook);
      const span = raw.listColumnOverrides(sheetIndex(this.#sheet)).find((c) => c.first <= col && col <= c.last);
      return span?.xfIndex ?? null;
    }

    getRowXfIndex(row) {
      requireIndex(row, argLabel('Sheet.getRowXfIndex', 0), 'coordinate');
      const raw = rawOf(this.#sheet.workbook);
      const override = raw.listRowOverrides(sheetIndex(this.#sheet)).find((r) => r.row === row);
      return override?.xfIndex ?? null;
    }

    // The raw call takes the excluded cells as parallel row and column arrays.
    getRangeXfBasesExcluding(range, excluded) {
      const where = 'Sheet.getRangeXfBasesExcluding';
      const raw = rawOf(this.#sheet.workbook);
      requireArray(excluded, argLabel(where, 1));
      const rows = [];
      const cols = [];
      for (const [i, cell] of excluded.entries()) {
        const { row, col } = addressOf(raw, cell, argLabel(where, 1, i));
        rows.push(requireIndex(row, argLabel(where, 1, i, 'row'), 'coordinate'));
        cols.push(requireIndex(col, argLabel(where, 1, i, 'col'), 'coordinate'));
      }
      return raw.getRangeXfBasesExcluding(sheetIndex(this.#sheet), range, rows, cols);
    }

    // The raw call takes the mapping as parallel base and target arrays.
    remapRangeXfs(range, mapping) {
      const where = 'Sheet.remapRangeXfs';
      const label = argLabel(where, 1);
      if (!(mapping instanceof Map)) throw new TypeError(`${label} must be a Map`);
      const from = [];
      const to = [];
      const isIndex = (value) => Number.isInteger(value) && value >= 0 && value <= 0xffffffff;
      for (const [base, target] of mapping) {
        if (!isIndex(base) || !isIndex(target)) {
          throw new TypeError(`${label} must map an XF index to an XF index`);
        }
        from.push(base);
        to.push(target);
      }
      rawOf(this.#sheet.workbook).remapRangeXfs(sheetIndex(this.#sheet), range, from, to);
    }
  }

  ooxmlState = {
    ...ooxmlState,
    workbook: (workbook) => new WorkbookOoxml(token, workbook),
    sheet: (sheet) => new SheetOoxml(token, sheet),
  };

  const sheetCall = (self, name, args) => rawOf(self.workbook)[name](sheetIndex(self), ...args);
  const tableCall = (self, name, args) => {
    const [raw, sheet, index] = tableState.locate(self);
    return raw[name](sheet, index, ...args);
  };
  const pivotCall = (self, name, args) => {
    const [raw, sheet, index] = pivotState.locate(self);
    return raw[name](sheet, index, ...args);
  };
  const cacheCall = (self, name, args) => cacheState.raw(self)[name](self.cacheId, ...args);
  // A drawing read the raw method answers with `null` for an object without that part (a picture's media,
  // a chart that is gone).
  const drawingRead =
    /^(?:(?:Drawing|Picture|Chart|Shape|Slicer)\.get(?:Drawing)?Properties|Chart\.get(?:Data|Preview))$/;
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

  // A `PivotFieldRef` argument or spec key of a `Pivot` method reaches the raw call as the field index.
  const fieldArgument = (name) => (self, args) => {
    if (typeof args[0] === 'number') {
      return [requireIndex(args[0], argLabel(`Pivot.${name}`, 0), 'position'), ...args.slice(1)];
    }
    if (typeof args[0] !== 'string') return args;
    const [raw, sheet, pivot] = pivotState.locate(self);
    return [pivotFieldIndex(raw, sheet, pivot, args[0], argLabel(`Pivot.${name}`, 0)), ...args.slice(1)];
  };
  // `fieldIndex` and `showAsBaseField` of a data field spec.
  const dataFieldSpec = (name, at) => (self, args) => {
    const spec = args[at];
    if (!isObject(spec)) return args;
    const keys = ['fieldIndex', 'showAsBaseField'].filter((key) => typeof spec[key] === 'string');
    if (keys.length === 0) return args;
    const [raw, sheet, pivot] = pivotState.locate(self);
    const resolved = { ...spec };
    for (const key of keys) {
      resolved[key] = pivotFieldIndex(raw, sheet, pivot, spec[key], argLabel(`Pivot.${name}`, at, key));
    }
    return args.map((arg, i) => (i === at ? resolved : arg));
  };
  // An axis order: each name becomes its field index; a number (`PIVOT_VALUES_FIELD` included) is left to the call.
  const fieldOrder = (name) => (self, args) => {
    if (!Array.isArray(args[0]) || !args[0].some((ref) => typeof ref === 'string')) return args;
    const [raw, sheet, pivot] = pivotState.locate(self);
    const fields = args[0].map((ref, i) =>
      typeof ref === 'string' ? pivotFieldIndex(raw, sheet, pivot, ref, argLabel(`Pivot.${name}`, 0, i)) : ref,
    );
    return [fields, ...args.slice(1)];
  };
  const pivotFieldRefs = [
    ...PIVOT_FIELD_ARGUMENTS.map((name) => [name, fieldArgument(name)]),
    ['addDataField', dataFieldSpec('addDataField', 0)],
    ['setDataField', dataFieldSpec('setDataField', 1)],
    ...['setRowFieldOrder', 'setColumnFieldOrder', 'setPageFieldOrder'].map((name) => [name, fieldOrder(name)]),
  ];
  for (const [name, convert] of pivotFieldRefs) {
    const inner = Pivot.prototype[name];
    if (typeof inner !== 'function') continue;
    defineMethod(
      Pivot.prototype,
      name,
      namedFunction(name, function (args) {
        return inner.apply(this, convert(this, args));
      }),
    );
  }

  // A `PivotFieldRef` argument of a `PivotCache` method reaches the raw call as the cache field index; a negative
  // index is `IndexOutOfRange`, as on `Pivot`.
  for (const [name, at] of PIVOT_CACHE_FIELD_ARGUMENTS) {
    const inner = PivotCache.prototype[name];
    if (typeof inner !== 'function') continue;
    defineMethod(
      PivotCache.prototype,
      name,
      namedFunction(name, function (args) {
        if (typeof args[at] !== 'string' && typeof args[at] !== 'number') return inner.apply(this, args);
        const label = argLabel(`PivotCache.${name}`, at);
        const out = [...args];
        out[at] = cacheFieldIndex(cacheState.raw(this), this.cacheId, args[at], label);
        return inner.apply(this, out);
      }),
    );
  }

  const argumentForms = [
    ['Workbook', Workbook, rawOf, () => undefined],
    ['Workbook', WorkbookOoxml, (self) => rawOf(ooxmlState.workbookOf(self)), () => undefined],
    ['Sheet', Sheet, (self) => rawOf(self.workbook), sheetIndex],
    ['Pivot', Pivot, (self) => rawOf(self.sheet.workbook), (self) => sheetIndex(self.sheet)],
    ['Table', Table, (self) => tableState.locate(self)[0], (self) => tableState.locate(self)[1]],
    ...['Drawing', 'Picture', 'Chart', 'Shape', 'Slicer'].map((owner) => [
      owner,
      { Drawing, Picture, Chart, Shape, Slicer }[owner],
      (self) => rawOf(self.sheet.workbook),
      (self) => sheetIndex(self.sheet),
    ]),
    [
      'Sheet',
      SheetOoxml,
      (self) => rawOf(ooxmlState.sheetOf(self).workbook),
      (self) => sheetIndex(ooxmlState.sheetOf(self)),
    ],
  ];
  for (const [owner, cls, rawFor, sheetFor] of argumentForms) {
    installArgumentForms(cls.prototype, owner, rawFor, sheetFor, handleIndex);
  }
  // A pivot cache's methods resolve their own record arguments; only the roles of their integer arguments apply.
  for (const [where, roles] of Object.entries(ARG_ROLES)) {
    const name = where.startsWith('PivotCache.') ? where.slice('PivotCache.'.length) : undefined;
    if (name === undefined || !Object.hasOwn(PivotCache.prototype, name)) continue;
    const inner = PivotCache.prototype[name];
    defineMethod(
      PivotCache.prototype,
      name,
      namedFunction(name, function (args) {
        checkArgumentRoles(where, roles, args);
        return inner.apply(this, args);
      }),
    );
  }
  // Under the arity guard: the handle is checked before a call reads or validates its input, and the workbook
  // cannot be disposed from inside the call.
  const workbookOfSheet = (sheet) => sheet.workbook;
  const guardedClasses = [
    [Workbook, (self) => self],
    [Sheet, workbookOfSheet],
    [Table, (self) => self.workbook],
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
  const guarded = [['Workbook', Workbook], ['Workbook', WorkbookOoxml], ['Sheet', SheetOoxml], ...handleClasses];
  for (const [owner, cls] of guarded) installInputGuards(cls.prototype, owner);
  for (const [cls, workbookOf] of guardedClasses) installCallGuards(cls.prototype, workbookOf, liveRaw, activeCalls);
  for (const [owner, cls] of guarded) installArityGuards(cls.prototype, owner);
  // The static factories take their arguments from the same table.
  installArityGuards(Workbook, 'Workbook');
  // Outermost: an argument error reads as the call the consumer made, on the class of the receiver.
  const drawingClassName = (self) => {
    for (const [name, cls] of [
      ['Picture', Picture],
      ['Chart', Chart],
      ['Shape', Shape],
      ['Slicer', Slicer],
    ]) {
      if (self instanceof cls) return name;
    }
    return 'Drawing';
  };
  for (const [owner, cls] of [
    ['Workbook', Workbook],
    ['Sheet', Sheet],
    ['Table', Table],
    ['Pivot', Pivot],
    ['PivotCache', PivotCache],
    ['WorkbookOoxml', WorkbookOoxml],
    ['SheetOoxml', SheetOoxml],
  ]) {
    installErrorLabels(cls.prototype, owner);
  }
  for (const cls of [Drawing, Picture, Chart, Shape, Slicer]) installErrorLabels(cls.prototype, drawingClassName);
  installErrorLabels(Workbook, 'Workbook');
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
 * names.
 */
function installArgumentForms(proto, owner, rawOf, sheetOf, handleIndex) {
  const prefix = `${owner}.`;
  const tables = [
    CELL_ARGS,
    RECT_ARGS,
    SPAN_ARGS,
    REF_ARGS,
    TYPED_ARGS,
    DEFAULT_ARGS,
    ARG_ROLES,
    ARG_TYPES,
    BOOLEAN_ARGS,
  ];
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
    const roles = ARG_ROLES[where];
    const types = ARG_TYPES[where];
    const booleans = BOOLEAN_ARGS[where];
    defineMethod(
      proto,
      name,
      namedFunction(name, function (args) {
        if (spans !== undefined) args = expandSpans(rawOf(this), where, spans, args);
        if (pairs !== undefined) args = expandCells(rawOf(this), where, pairs, args);
        if (defaults !== undefined) args = fillDefaults(defaults, args);
        if (typed !== undefined) args = normalizeTypedArgs(rawOf(this), sheetOf(this), where, typed, args);
        if (rects !== undefined) args = expandRects(rawOf(this), where, rects, args);
        if (ref !== undefined && (typeof args[ref[0]] !== 'number' || REF_KINDS[ref[1]]?.keyed === true)) {
          const named = (kind, value, label) => handleIndex(this, kind, value, label);
          args = resolveRef(rawOf(this), sheetOf(this), where, ref, args, named);
        }
        if (types !== undefined || booleans !== undefined) args = checkArgumentTypes(where, types, args, booleans);
        if (roles !== undefined) checkArgumentRoles(where, roles, args);
        return inner.apply(this, args);
      }),
    );
  }
}

/**
 * Normalises the `bytes` arguments of `types` and refuses a `value` argument that is not a record, and a
 * `boolean` argument (the positions of `booleans`) that is not a boolean; an omitted argument is left to the call.
 */
function checkArgumentTypes(where, types, args, booleans) {
  let out = args;
  for (const [position, type] of types ?? []) {
    const value = out[position];
    if (value === undefined) continue;
    if (type === 'bytes') {
      const bytes = requireBytes(value, where, argumentPath(where, position));
      if (bytes !== value) {
        out = [...out];
        out[position] = bytes;
      }
    } else if (!isObject(value)) {
      throw new TypeError(`${argLabel(where, position)} must be an object`);
    }
  }
  for (const position of booleans ?? []) {
    if (out[position] !== undefined) requireBoolean(out[position], argLabel(where, position));
  }
  return out;
}

/**
 * Checks each number at a position of `roles` (`ARG_ROLES`, positions of the coordinate form) by its role; other
 * types are left to the call.
 */
function checkArgumentRoles(where, roles, args) {
  for (const [position, role] of roles) {
    const value = args[position];
    if (typeof value === 'number') requireIndex(value, argLabel(where, position), role);
  }
}

/**
 * Wraps every defined method of `owner` listed in `METHOD_ARITY` so that a call
 * with more positional arguments than the declarations allow throws
 * `TypeError: <owner>.<method> takes at most N arguments`. A call in the string
 * form (an A1 cell for each `CELL_ARGS` pair, an A1 span for a `SPAN_ARGS` one that
 * is not `single`) takes one argument fewer per string. `[Symbol.dispose]` is an alias
 * defined afterwards.
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
        for (const [at, , form] of spans ?? []) if (form !== 'single' && typeof args[at] === 'string') limit -= 1;
        if (args.length > limit) {
          throw new TypeError(`${where} takes at most ${limit} argument${limit === 1 ? '' : 's'}`);
        }
        return inner.apply(this, args);
      }),
    );
  }
}

/**
 * Wraps every defined method of `owner` listed in `INPUT_KEYS` so that a key of its input record argument
 * outside the record (and the getter record it passes back) is a `TypeError` before the call reads anything
 * else. The table position is that of the A1-free form: a call in the string form of a cell pair
 * (`CELL_ARGS`) or a span (`SPAN_ARGS`) has the record that many arguments earlier.
 */
function installInputGuards(proto, owner) {
  const prefix = `${owner}.`;
  for (const [where, [at]] of Object.entries(INPUT_KEYS)) {
    if (!where.startsWith(prefix)) continue;
    const name = where.slice(prefix.length);
    if (!Object.hasOwn(proto, name)) continue;
    const inner = proto[name];
    const pairs = CELL_ARGS[where];
    const spans = SPAN_ARGS[where];
    defineMethod(
      proto,
      name,
      namedFunction(name, function (args) {
        let position = at;
        if (pairs !== undefined && typeof args[pairs[0][0]] === 'string') position -= pairs.length;
        for (const [span, , form] of spans ?? []) {
          if (form !== 'single' && typeof args[span] === 'string') position -= 1;
        }
        // A plain record with getters is copied once, so the check and the call read each getter a single time.
        const record = args[position];
        if (
          isObject(record) &&
          Object.getPrototypeOf(record) === Object.prototype &&
          Object.values(Object.getOwnPropertyDescriptors(record)).some((d) => d.get !== undefined)
        ) {
          args = [...args];
          // Own properties, enumerable or not, each read once; a spread would drop a non-enumerable getter.
          const copy = {};
          for (const key of Object.getOwnPropertyNames(record)) copy[key] = record[key];
          args[position] = copy;
        }
        refuseUnknownInputKeys(where, args[position]);
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

const rangeKey = (record) => renamed(record, 'rect', 'range');
const sortRead = (sort) =>
  sort === null || sort === undefined
    ? sort
    : { ...renamed(sort, 'ref', 'range'), conditions: sort.conditions.map((c) => renamed(c, 'ref', 'range')) };
const autoFilterRead = (filter) => (isObject(filter) ? { ...filter, sort: sortRead(filter.sort) } : filter);

// Arguments the generated forwarders pass on whole: `Class.method` -> [position in the public call, type]. A
// `bytes` argument is normalised to a `Uint8Array`; a `value` argument must be a tagged `Value` record.
const ARG_TYPES = Object.freeze({
  'Sheet.addPicture': [[0, 'bytes']],
  'Sheet.restoreDrawing': [[0, 'bytes']],
  'Sheet.restoreRangeXfs': [[0, 'bytes']],
});

// Reads whose raw record names a key other than the public one, by `Class.method`: the result as the public record.
const READ_ADAPTERS = Object.freeze({
  'Sheet.getArrayFormula': rangeKey,
  'Sheet.getAutoFilter': autoFilterRead,
  'Sheet.getDataTable': rangeKey,
  'Sheet.getHyperlink': rangeKey,
  'Sheet.getPagination': (pagination) => ({ ...pagination, pages: pagination.pages.map(rangeKey) }),
  'Sheet.getSpill': rangeKey,
  'Sheet.listHyperlinks': (links) => links.map(rangeKey),
  'Pivot.getSpill': rangeKey,
  'PivotCache.getRecordValue': (record) => ({ value: nativeOf(record.value), isDate: record.isDate }),
  'Table.getAutoFilter': autoFilterRead,
  'Workbook.parseReference': rangeKey,
});

const autoFilterInput = (filter) =>
  isObject(filter) && isObject(filter.sort)
    ? {
        ...filter,
        sort: {
          ...renamed(filter.sort, 'range', 'ref'),
          ...(Array.isArray(filter.sort.conditions)
            ? { conditions: filter.sort.conditions.map((c) => renamed(c, 'range', 'ref')) }
            : {}),
        },
      }
    : filter;

// Inputs whose raw record names a key other than the public one, by `Class.method`: the arguments as the raw
// method takes them.
const ARG_ADAPTERS = Object.freeze({
  'Sheet.setAutoFilter': (args) => [autoFilterInput(args[0]), ...args.slice(1)],
  'Table.setAutoFilter': (args) => [autoFilterInput(args[0]), ...args.slice(1)],
});

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
    const adaptRead = READ_ADAPTERS[where] ?? ((result) => result);
    const adaptOwn = ARG_ADAPTERS[where] ?? ((args) => args);
    const adaptArgs = (args) => adaptOwn(args).map(recordThrown);
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
          return adaptRead(out);
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
      const reader = rawName.replace(/^(?:set|update)/, 'get');
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
          return adaptRead(call(this, rawName, adaptArgs(args)));
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
            return adaptRead(call(this, rawName, adaptArgs(colorArgs(args, where))));
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
              throw new FormulonError(
                StatusCode.InvalidArgument,
                `${argLabel(where, position)} must not be negative; pass null for none`,
                `value=${value}`,
              );
            }
            return adaptRead(call(this, rawName, adaptArgs(args)));
          }),
    );
  }
}
