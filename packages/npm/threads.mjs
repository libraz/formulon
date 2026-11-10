// Public ESM surface for @libraz/formulon/threads: the pthread build, whose
// `recalcParallel` runs on Web Workers. It needs SharedArrayBuffer, so a
// browser page must be cross-origin isolated (COOP/COEP). The pthread
// workers load formulon_threads_core.js directly, not this shim (see
// scripts/stage.mjs).

import * as common from './common.js';
import createModule from './formulon_threads_core.js';

// A workbook dropped without `dispose()` frees its engine object once collected (best-effort). The held
// value is the raw object alone, which holds no reference back to its facade.
const registry = new FinalizationRegistry((raw) => raw.delete());

/**
 * Loads the engine and resolves to the frozen `Formulon` object.
 *
 * @param {object} [options] `FormulonOptions`: `locateFile`, `print`, `printErr`, `wasm`; any other key, or a
 *   key of the wrong type, rejects with a `TypeError`.
 * @returns {Promise<object>}
 */
export function createFormulon(options) {
  return common.loadFormulon(createModule, options, common, registry);
}
export default createFormulon;

// The value tables, the error classes, `enumName`, `toNative` and `mergeFunctionMetadata`;
// everything else is a member of the `Formulon` object.
export {
  AnchorEditAs,
  AnchorKind,
  BlanksAs,
  BorderStyle,
  CalcMode,
  CellShift,
  CfDataBarAxisPosition,
  CfDataBarDirection,
  CfMatchKind,
  CfRuleType,
  CfValueObjectType,
  ChartAxisCrosses,
  ChartAxisKind,
  ChartAxisPosition,
  ChartBarDirection,
  ChartDomain,
  ChartExtendedLayout,
  ChartGrouping,
  ChartLabelPosition,
  ChartLegendPosition,
  ChartLineDash,
  ChartMarkerSymbol,
  ChartPlotType,
  ChartPointState,
  ChartRole,
  ChartSeriesDirection,
  ChartSourceKind,
  ChartTickLabelPosition,
  ChartTitleKind,
  CollationStrength,
  ColorContext,
  ColorKind,
  ColorResolution,
  ComparisonOperator,
  DatePeriod,
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
  mergeFunctionMetadata,
  PageOrder,
  PageOrientation,
  PaperSize,
  ParseErrorCode,
  PhoneticAlignment,
  PhoneticType,
  PIVOT_VALUES_FIELD,
  PivotAggregation,
  PivotAxis,
  PivotCalendar,
  PivotCellKind,
  PivotDatePeriod,
  PivotFilterKind,
  PivotReportLayout,
  PivotShowAsBase,
  PivotShowValuesAs,
  PivotSortOrder,
  ReadingOrder,
  RefMode,
  SheetViewMode,
  SheetVisibility,
  SlicerCrossFilter,
  SlicerSortOrder,
  SortBy,
  SortMethod,
  SparklineAxisType,
  SparklineType,
  SpillCause,
  StatusCode,
  TableStyleElementType,
  TableTotalsRowFunction,
  ThemeSource,
  toNative,
  UnderlineStyle,
  ValidationErrorStyle,
  ValidationType,
  ValueKind,
  VerticalAlign,
  WidthPlatform,
  WorkbookFormat,
} from './common.js';
