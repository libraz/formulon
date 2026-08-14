// Public ESM surface for @libraz/formulon. The generated Emscripten module
// remains an implementation detail so this shim can expose the value
// constants declared in formulon.d.ts alongside its default factory.

import createFormulon from './formulon_core.js';

export default createFormulon;

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
});
export const PIVOT_SHOW_AS_BASE_PREVIOUS = 1048828;
export const PIVOT_SHOW_AS_BASE_NEXT = 1048829;
export const PivotFilterType = Object.freeze({
  ValueTop10: 0,
  ValueGreaterThan: 1,
  ValueBetween: 2,
  LabelContains: 3,
  LabelBeginsWith: 4,
  LabelDate: 5,
});
export const PivotDateGrouping = Object.freeze({
  Day: 0,
  Month: 1,
  Quarter: 2,
  Year: 3,
  Week: 4,
  Hour: 5,
  Minute: 6,
  Second: 7,
});
export const PivotCalendar = Object.freeze({ Gregorian: 0, Japanese: 1 });
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
export const ExternalLinkKind = Object.freeze({ Unknown: 0, ExternalBook: 1, Ole: 2, Dde: 3 });
export const LogLevel = Object.freeze({ Debug: 0, Info: 1, Warn: 2, Error: 3, Off: 4 });
