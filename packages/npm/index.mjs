// Public ESM surface for @libraz/formulon: the single-threaded build, which
// loads without cross-origin isolation. The generated Emscripten module stays
// an implementation detail: callers see only the `Formulon` object.

import * as common from './common.js';
import createModule from './formulon_core.js';

/**
 * Loads the engine and resolves to the frozen `Formulon` object.
 *
 * @param {object} [options] `FormulonOptions`: `locateFile`, `print`, `printErr`.
 * @returns {Promise<object>}
 */
export function createFormulon(options) {
  return common.loadFormulon(createModule, options, common);
}
export default createFormulon;

// The value tables and the error class; everything else is a member of the
// `Formulon` object.
export {
  AnchorEditAs,
  AnchorKind,
  CalcMode,
  CfDataBarAxisPosition,
  CfDataBarDirection,
  CfMatchKind,
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
  ExternalLinkKind,
  FilterKind,
  FilterOperator,
  FormulonError,
  FunctionAvailability,
  GeometryMode,
  ImageFormat,
  LocaleDateOrder,
  LogLevel,
  PageOrientation,
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
  ValidationErrorStyle,
  ValueKind,
  WorkbookFormat,
} from './common.js';
