// Public ESM surface for @libraz/formulon/threads: the pthread build, whose
// `recalcParallel` runs on Web Workers. It needs SharedArrayBuffer, so a
// browser page must be cross-origin isolated (COOP/COEP). The pthread
// workers load formulon_threads_core.js directly, not this shim (see
// scripts/stage.mjs).

import * as common from './common.js';
import createModule from './formulon_threads_core.js';

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
  CfMatchKind,
  ChartAxisCrosses,
  ChartAxisKind,
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
  GeometryMode,
  ImageFormat,
  LogLevel,
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
