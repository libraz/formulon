//
// ESM entry point for @libraz/formulon-native.
//
// This shim loads the compiled `.node` addon and re-exports the
// JS-friendly API. The `.node` module itself is a CommonJS binary
// loaded via `createRequire`, then surfaced to ESM consumers verbatim.
// The JS-visible shape is intentionally kept identical to
// `@libraz/formulon` (the WASM package) so consumers can swap binaries
// without code changes.
//
// Binary lookup order at require-time:
//   1. dist/prebuilds/<platform>-<arch>/formulon.node  (shipped artifacts)
//   2. dist/formulon.node                              (local dev fallback)

import { existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import {
  AnchorEditAs,
  AnchorKind,
  CalcMode,
  CfMatchKind,
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
  GeometryMode,
  ImageFormat,
  LogLevel,
  mergeFunctionMetadata,
  PIVOT_SHOW_AS_BASE_NEXT,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  PivotAggregation,
  PivotAxis,
  PivotCalendar,
  PivotCellKind,
  PivotDateGrouping,
  PivotFilterType,
  PivotFilterValueKind,
  PivotReportLayout,
  PivotShowValuesAs,
  SheetVisibility,
  SortBy,
  SortMethod,
  ThemeSource,
  ValidationErrorStyle,
  ValueKind,
  WorkbookFormat,
} from '../npm/common.mjs';

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

export const Workbook = native.Workbook;
export const evalFormula = native.evalFormula;
export const version = native.version;
/** Alias of {@link version}, matching the WASM binding's name. */
export const versionString = native.versionString;
export const lastErrorMessage = native.lastErrorMessage;
export const lastErrorContext = native.lastErrorContext;
export const statusString = native.statusString;
export const errorDisplayName = native.errorDisplayName;
export const setLogMinLevel = native.setLogMinLevel;
export const setLogSink = native.setLogSink;

// Keep the native entry point's explicit export list in lockstep with the
// WASM package while taking the shared values and helper from one canonical
// module. The default export below deliberately remains explicit as well so
// adding an internal common export cannot change this package's public shape.
export {
  AnchorEditAs,
  AnchorKind,
  CalcMode,
  CfMatchKind,
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
  GeometryMode,
  ImageFormat,
  LogLevel,
  mergeFunctionMetadata,
  PIVOT_SHOW_AS_BASE_NEXT,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  PivotAggregation,
  PivotAxis,
  PivotCalendar,
  PivotCellKind,
  PivotDateGrouping,
  PivotFilterType,
  PivotFilterValueKind,
  PivotReportLayout,
  PivotShowValuesAs,
  SheetVisibility,
  SortBy,
  SortMethod,
  ThemeSource,
  ValidationErrorStyle,
  ValueKind,
  WorkbookFormat,
};

export default {
  Workbook,
  evalFormula,
  version,
  versionString,
  lastErrorMessage,
  lastErrorContext,
  statusString,
  errorDisplayName,
  setLogMinLevel,
  setLogSink,
  mergeFunctionMetadata,
  ValueKind,
  CfMatchKind,
  PivotCellKind,
  PivotAxis,
  PivotAggregation,
  PivotShowValuesAs,
  PIVOT_SHOW_AS_BASE_PREVIOUS,
  PIVOT_SHOW_AS_BASE_NEXT,
  PivotFilterType,
  PivotDateGrouping,
  PivotCalendar,
  PivotFilterValueKind,
  PivotReportLayout,
  LogLevel,
  CalcMode,
  SheetVisibility,
  GeometryMode,
  DisplayStatus,
  ColorContext,
  ImageFormat,
  DrawingObjectKind,
  AnchorKind,
  AnchorEditAs,
  FilterKind,
  FilterOperator,
  DynamicFilterType,
  SortBy,
  SortMethod,
  DateTimeGrouping,
  ValidationErrorStyle,
  ColorResolution,
  ThemeSource,
  EffectiveStyleSource,
  ExternalLinkKind,
  ErrorCode,
  WorkbookFormat,
};
