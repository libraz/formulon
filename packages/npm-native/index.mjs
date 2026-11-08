//
// ESM entry point for @libraz/formulon-native.
//
// This shim loads the compiled `.node` addon (a CommonJS binary loaded via
// `createRequire`) and exposes it through the same `createFormulon()` entry
// as `@libraz/formulon` (the WASM package), so consumers can swap binaries
// without code changes.
//
// Binary lookup order at require-time:
//   1. dist/prebuilds/<platform>-<arch>/formulon.node  (shipped artifacts)
//   2. dist/formulon.node                              (local dev fallback)

import { existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as common from '../npm/common.mjs';

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
native.__setErrorClass(common.FormulonError);

/**
 * Resolves to a new frozen `Formulon` object, whose `Workbook` is a class of
 * its own, as each `@libraz/formulon` call yields. `options` is accepted for
 * parity with `@libraz/formulon` and not used.
 *
 * @param {object} [_options] `FormulonOptions`.
 * @returns {Promise<object>}
 */
export function createFormulon(_options) {
  return Promise.resolve(common.buildFormulon(native, common));
}
export default createFormulon;

// The value tables, the error classes and `enumName`; everything else is a member of the
// `Formulon` object.
export {
  AnchorEditAs,
  AnchorKind,
  BlanksAs,
  BorderStyle,
  CalcMode,
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
  PageOrder,
  PageOrientation,
  PaperSize,
  ParseErrorCode,
  PhoneticAlignment,
  PhoneticType,
  PIVOT_SHOW_AS_BASE_NEXT,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  PIVOT_VALUES_FIELD,
  PivotAggregation,
  PivotAxis,
  PivotCalendar,
  PivotCellKind,
  PivotDatePeriod,
  PivotFilterKind,
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
  SparklineType,
  StatusCode,
  TableStyleElementType,
  TableTotalsRowFunction,
  ThemeSource,
  UnderlineStyle,
  ValidationErrorStyle,
  ValidationType,
  ValueKind,
  VerticalAlign,
  WidthPlatform,
  WorkbookFormat,
} from '../npm/common.mjs';
