//
// ESM entry point for @libraz/formulon-native.
//
// This shim loads the compiled `.node` addon (a CommonJS binary loaded via
// `createRequire`) and exposes it through the same `createFormulon()` entry
// as `@libraz/formulon` (the WASM package), so consumers can swap binaries
// without code changes.
//
// Binary lookup order, on the first `createFormulon()` call:
//   1. prebuilds/<platform>-<arch>/formulon.node  (shipped artifacts)
//   2. formulon.node                              (local dev fallback)
// `FORMULON_NATIVE_ROOT` replaces the directory both are looked up in.

import { existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import * as common from '../npm/common.mjs';

const require_ = createRequire(import.meta.url);
const here = path.dirname(fileURLToPath(import.meta.url));

let native;

/** Loads the addon once; a missing prebuild is a plain host `Error`, not a `FormulonError`. */
function loadNative() {
  if (native !== undefined) return native;
  const root = process.env.FORMULON_NATIVE_ROOT || here;
  const prebuildSlot = path.join(root, 'prebuilds', `${process.platform}-${process.arch}`, 'formulon.node');
  const fallbackSlot = path.join(root, 'formulon.node');
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
  const loaded = require_(nativePath);
  loaded.__setErrorClass(common.FormulonError);
  native = loaded;
  return native;
}

/**
 * Resolves to a new frozen `Formulon` object, whose `Workbook` is a class of
 * its own, as each `@libraz/formulon` call yields. Rejects with a plain `Error`
 * when no prebuild matches the host. `options` is accepted for parity with
 * `@libraz/formulon` and not used.
 *
 * @param {object} [_options] `FormulonOptions`.
 * @returns {Promise<object>}
 */
export async function createFormulon(_options) {
  return common.buildFormulon(loadNative(), common);
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
} from '../npm/common.mjs';
