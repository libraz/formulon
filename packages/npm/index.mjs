// Public ESM surface for @libraz/formulon: the single-threaded build, which
// loads without cross-origin isolation. The generated Emscripten module stays
// an implementation detail: callers see only the `Formulon` object.

import * as common from './common.js';
import createModule from './formulon_core.js';

// A workbook dropped without `dispose()` frees its engine object once collected (best-effort). The held
// value is the raw object alone, which holds no reference back to its facade.
const registry = new FinalizationRegistry((raw) => raw.delete());

/**
 * Loads the engine and resolves to the frozen `Formulon` object.
 *
 * @param {object} [options] `FormulonOptions`: `locateFile`, `print`, `printErr`, `wasm`.
 * @returns {Promise<object>}
 */
export function createFormulon(options) {
  return common.loadFormulon(createModule, options, common, registry);
}
export default createFormulon;

// The value tables, the error classes and `enumName`; everything else is a member of the
// `Formulon` object.
export {
  AnchorEditAs,
  AnchorKind,
  BorderStyle,
  CalcMode,
  CfDataBarAxisPosition,
  CfDataBarDirection,
  CfMatchKind,
  CfOperator,
  CfRuleType,
  CfTimePeriod,
  CfValueObjectType,
  ChartAxisCrosses,
  ChartAxisKind,
  ChartAxisPosition,
  ChartBarDirection,
  ChartBlanksAs,
  ChartGrouping,
  ChartInsertKind,
  ChartLabelPosition,
  ChartLegendPosition,
  ChartLineDash,
  ChartMarkerSymbol,
  ChartPlotType,
  ChartPointState,
  ChartSeriesDirection,
  ChartSourceKind,
  ChartTitleKind,
  CollationStrength,
  ColorContext,
  ColorKind,
  ColorResolution,
  DateTimeGrouping,
  DisplayStatus,
  DrawingObjectKind,
  DynamicFilterType,
  EffectiveStyleSource,
  ErrorCode,
  ErrorDisplayName,
  ExcelError,
  ExternalLinkKind,
  enumName,
  FillPattern,
  FilterKind,
  FilterOperator,
  FontFamily,
  FontScheme,
  FontVerticalAlign,
  FormulonError,
  FunctionAvailability,
  GeometryMode,
  GradientType,
  HorizontalAlign,
  IconSetName,
  ImageFormat,
  LocaleDateOrder,
  LogLevel,
  PageOrientation,
  ParseErrorCode,
  PIVOT_SHOW_AS_BASE_NEXT,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  PIVOT_VALUES_FIELD,
  PivotAggregation,
  PivotAxis,
  PivotCalendar,
  PivotCellKind,
  PivotDatePeriod,
  PivotFilterType,
  PivotFilterValueKind,
  PivotReportLayout,
  PivotShowValuesAs,
  PivotSortOrder,
  ReadingOrder,
  RefMode,
  ShapePaint,
  SheetViewMode,
  SheetVisibility,
  SlicerCrossFilter,
  SlicerSortOrder,
  SortBy,
  SortMethod,
  SparklineAxisType,
  SparklineEmptyCells,
  SparklineType,
  StatusCode,
  ThemeSource,
  UnderlineStyle,
  ValidationErrorStyle,
  ValidationOperator,
  ValidationType,
  ValueKind,
  VerticalAlign,
  WorkbookFormat,
} from './common.js';
