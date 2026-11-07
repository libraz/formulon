// Compile-time checks of the public declarations. Never executed: `make npm-typecheck` runs tsc over it
// against the staged d.ts. A line marked `@ts-expect-error` must fail to compile; it is the rejected form.
import createFormulon, {
  CfOperator,
  CfRuleType,
  ChartSourceKind,
  ExcelError,
  FormulonError,
  IconSetName,
  ParseErrorCode,
  ValueKind,
  enumName,
  type ChartDataSource,
  type CivilTime,
  type Color,
  type ConditionalFormat,
  type ConditionalFormatInput,
  type Formulon,
  type FormulonOptions,
  type NativeValue,
  type ParseError,
  type ProfileId,
  type RangeCell,
  type Pivot,
  type PivotCache,
  type Shape,
  type ShapePatch,
  type ShapeProperties,
  type Sheet,
  type Slicer,
  type SlicerPatch,
  type SlicerProperties,
  type Table,
  type TablePatch,
  type TableProperties,
  type Value,
  type Workbook,
} from '@libraz/formulon';

type Equal<A, B> = (<T>() => T extends A ? 1 : 2) extends <T>() => T extends B ? 1 : 2 ? true : false;
declare function assertType<T extends true>(): T;

declare const formulon: Formulon;
declare const wb: Workbook;
declare const sheet: Sheet;
declare const table: Table;
declare const tableProps: TableProperties;
declare const shape: Shape;
declare const shapeProps: ShapeProperties;
declare const slicer: Slicer;
declare const slicerProps: SlicerProperties;
declare const pivot: Pivot;
declare const cache: PivotCache;
declare const source: ChartDataSource;
declare const rect: { firstRow: 0; firstCol: 0; lastRow: 1; lastCol: 1 };
declare const bytes: Uint8Array;
declare const compiled: WebAssembly.Module;

export function tablePatch(): void {
  table.update({ name: 'Sales', styleName: null, headerRow: true, totalsRow: false });
  table.update({ range: 'A1:C9' });
  table.update({ range: { sheet: 0, rect } });
  table.update(tableProps);
  // @ts-expect-error `range` is a rectangle or an A1 string
  table.update({ range: 5 });
  // @ts-expect-error the count fields became booleans named headerRow / totalsRow
  table.update({ headerRowCount: 1 });
  // @ts-expect-error unknown key
  table.update({ bogus: 1 });
  const header: boolean = tableProps.headerRow;
  const totals: boolean = tableProps.totalsRow;
  void [header, totals];
  assertType<Equal<TablePatch['range'], import('@libraz/formulon').SheetRect | string | undefined>>();
}

export function slicerPatch(): void {
  slicer.update({ caption: 'Region', style: null, columnCount: 2 });
  slicer.update({ fromRow: 0, fromCol: 0, toRow: 4, toCol: 4 });
  slicer.update(slicerProps);
  // @ts-expect-error `name` cannot be updated
  slicer.update({ name: 'renamed' });
  // @ts-expect-error unknown key
  slicer.update({ bogus: true });
  const patch: SlicerPatch = {};
  void patch;
}

export function shapePatch(): void {
  shape.update({ text: 'hi', hidden: false });
  // `null` on fillKind / lineKind means unchanged, so a read-back value passes as is.
  shape.update({ fillKind: null, lineKind: null });
  shape.update(shapeProps);
  // @ts-expect-error unknown key
  shape.update({ bogus: 1 });
  // @ts-expect-error a string is not an anchor kind
  shape.update({ anchorKind: 'oneCell' });
  const patch: ShapePatch = { rotation: 90 };
  void patch;
  shape.rename('Callout');
  shape.setZOrder(0);
  const snapshot: Uint8Array = shape.snapshot();
  void snapshot;
  slicer.setZOrder(0);
  // @ts-expect-error a slicer has no snapshot
  slicer.snapshot();
  // @ts-expect-error a slicer cannot be renamed
  slicer.rename('x');
}

export function chartSourceKind(): void {
  assertType<Equal<ChartDataSource['kind'], ChartSourceKind>>();
  const none: ChartSourceKind = ChartSourceKind.None;
  const ref: ChartSourceKind = ChartSourceKind.Ref;
  const literal: ChartSourceKind = ChartSourceKind.Literal;
  void [none, ref, literal];
  if (source.kind === ChartSourceKind.Ref) void source.formula;
  // @ts-expect-error the kind is an ordinal, not a string
  const text: ChartSourceKind = 'ref';
  void text;
  // @ts-expect-error out of range
  const outside: ChartSourceKind = 3;
  void outside;
}

export async function options(): Promise<void> {
  await createFormulon();
  await createFormulon({ wasm: bytes });
  await createFormulon({ wasm: new ArrayBuffer(8) });
  await createFormulon({ wasm: compiled });
  const withLocate: FormulonOptions = { locateFile: (path) => path, print: () => {}, printErr: () => {} };
  void withLocate;
}

export function convertFunctions(): void {
  const civil: CivilTime = formulon.convertSerialToCivil(45352.5);
  const serial: number = formulon.convertCivilToSerial(civil);
  const serial1904: number = formulon.convertCivilToSerial(civil, true);
  const back: CivilTime = formulon.convertSerialToCivil(serial1904, true);
  void [serial, back];
  // @ts-expect-error date1904 is a boolean
  formulon.convertCivilToSerial(civil, 1);
  // @ts-expect-error the serial is a number
  formulon.convertSerialToCivil('45352');
}

export function pivotCaches(): void {
  cache.setWorksheetSource(null);
  cache.setWorksheetSource({});
  cache.setWorksheetSource({ ref: 'A1:C9', sheet: 'Data' });
  cache.setWorksheetSource({ name: 'SalesTable' });
  // @ts-expect-error clearing is `null`, not `{ present: false }`
  cache.setWorksheetSource({ present: false });
  // @ts-expect-error the argument is required
  cache.setWorksheetSource();
}

export function pivotSharedItems(): void {
  pivot.addFieldSharedItem(0, 1);
  pivot.addFieldSharedItem(0, 1, {});
  pivot.addFieldSharedItem(0, 1, { visible: false });
  // @ts-expect-error the third argument is an options bag
  pivot.addFieldSharedItem(0, 1, false);
  pivot.setAnchor('A1');
}

export function geometryModes(): void {
  sheet.getCellRectPt('A1:B2');
  sheet.getWidthModel();
}

export function readonlyArrays(): void {
  const row: readonly (number | string)[] = [1, 'a'];
  const grid: readonly (readonly (number | string)[])[] = [row];
  sheet.setValues('A1', grid);
  sheet.setSelection({ active: { row: 0, col: 0 }, ranges: [rect] as readonly (typeof rect)[] });
  const sheets: readonly Sheet[] = wb.listSheets();
  void sheets;
  // @ts-expect-error a returned list is read-only
  wb.listSheets().push(sheet);
}

export function autoFilterInput(): void {
  sheet.setAutoFilter({ range: rect });
  sheet.setAutoFilter({ range: rect, columns: [{ colId: 0 }], sort: null });
  table.setAutoFilter({ range: rect, columns: [{ colId: 0, kind: 3, topVal: 2 }], sort: { ref: rect } });
  // @ts-expect-error a column needs its colId
  sheet.setAutoFilter({ range: rect, columns: [{ kind: 3 }] });
  // @ts-expect-error the range is required
  sheet.setAutoFilter({ columns: [] });
  const current = sheet.getAutoFilter();
  if (current !== null) {
    sheet.setAutoFilter(current);
    table.setAutoFilter(current);
  }
}

export function nativeReads(): void {
  const v: NativeValue = sheet.getValue('A1');
  const typed: Value = sheet.getTypedValue(0, 0);
  const grid: readonly (readonly NativeValue[])[] = sheet.getValues('A1:B2');
  const typedGrid: readonly (readonly Value[])[] = sheet.getTypedValues('A1:B2');
  sheet.setValues('A1', typedGrid);
  if (v instanceof ExcelError) void v.code;
  if (typed.kind === ValueKind.Array) void [typed.rows, typed.cols, typed.cells];
  // @ts-expect-error a native read carries no kind tag
  void v.kind;
  const cells: readonly RangeCell[] = sheet.listCells();
  const kind: ValueKind = cells[0].kind;
  const first: NativeValue = cells[0].value;
  const page = sheet.getCellsInRange('A1:C9');
  const paged: NativeValue = page.cells[0].value;
  const scalar: NativeValue = formulon.evaluateFormula('=1+1');
  const typedScalar: Value = formulon.evaluateTypedFormula('=1+1');
  const rows: readonly (readonly NativeValue[])[] = sheet.evaluateFormulaArray('A1', '=SEQUENCE(2)');
  const typedRows: readonly (readonly Value[])[] = sheet.evaluateTypedFormulaArray(0, 0, '=SEQUENCE(2)');
  const date: Date | null = sheet.getDate('A1');
  const used = sheet.getUsedRange();
  for (const row of sheet.iterRows()) void row;
  void [grid, kind, first, paged, scalar, typedScalar, rows, typedRows, date, used];
}

export function cellWrites(): void {
  sheet.setValue('A1', '=SUM(B1:B3)');
  sheet.setValue('A2', new Date(Date.UTC(2024, 2, 1)));
  sheet.setValue('A3', new ExcelError(1));
  sheet.setValues('A4', [[1, 'x', null, true, new Date()]]);
  // @ts-expect-error undefined is not a cell value
  sheet.setValue('A5', undefined);
}

export function workbookSurface(): void {
  for (const ws of wb) void ws.name;
  wb.copySheet(0, null, 'Copy');
  wb.copySheet('Sheet1');
  wb.recalcParallel();
  sheet.getPrecedents('A1');
  sheet.evaluateCfRange('A1:B9');
  const id: ProfileId = wb.getProfileId();
  wb.setProfileId(id);
  // @ts-expect-error renamed to getProfileId
  wb.getExcelProfileId();
  const usage: number = wb.getMemoryUsage();
  wb.setPinnedNow(new Date());
  wb.setDefinedName({ name: 'Rate', formula: '0.05', scope: 'Sheet1' });
  // @ts-expect-error the scope is `scope`, not a sheet index field
  wb.setDefinedName({ name: 'Rate', formula: '0.05', localSheetId: 0 });
  formulon.Workbook.load(new ArrayBuffer(8));
  formulon.Workbook.load(new DataView(new ArrayBuffer(8)));
  formulon.convertCivilToSerial({ year: 2024, month: 3, day: 1 });
  void usage;
}

export function ooxmlLayer(): void {
  const font = wb.ooxml.getFont(0);
  const fontIndex: number = wb.ooxml.addFont(font);
  const xf: number = sheet.ooxml.getCellXfIndex('A1');
  sheet.ooxml.setColumnStyle('B', xf);
  const columnStyle: number | null = sheet.ooxml.getColumnStyle(1);
  const sheetPr: string | null = sheet.ooxml.getSheetPrXml();
  wb.setDefaultFont(font);
  // @ts-expect-error the style tables moved to `ooxml`
  wb.getFont(0);
  // @ts-expect-error the XF index accessors moved to `ooxml`
  sheet.getCellXfIndex('A1');
  // @ts-expect-error a layout record no longer carries the style XF
  void sheet.listColumnOverrides()[0].styleXf;
  void [fontIndex, columnStyle, sheetPr];
}

export function colors(): void {
  sheet.setTabColor('#FF0000');
  sheet.setTabColor(null);
  const tab: Color | null = sheet.getTabColor();
  const font = wb.ooxml.getFont(0);
  const color: Color = font.color;
  // @ts-expect-error the ARGB sibling fields are gone
  void font.colorArgb;
  shape.update({ fillKind: 1, fillColor: '#336699' });
  const fill: Color | null = shapeProps.fillColor;
  void [tab, color, fill];
}

export function conditionalFormats(): void {
  sheet.addCfRule({ type: CfRuleType.CellIs, ranges: 'A1:A9', op: CfOperator.GreaterThan, formula1: '10', dxfId: 0 });
  sheet.addCfRule({ type: CfRuleType.Expression, ranges: ['A1:A9', rect], formula1: '=A1>0' });
  sheet.addCfRule({
    type: CfRuleType.ColorScale,
    ranges: [rect],
    colorScale: { thresholds: [{ type: 3 }, { type: 4 }], colors: ['#FF0000', '#00FF00'] },
  });
  sheet.addCfRule({
    type: CfRuleType.IconSet,
    ranges: [rect],
    iconSet: { name: IconSetName.ThreeArrows, thresholds: [{ type: 1, value: '33' }, { type: 1, value: '67' }] },
  });
  // @ts-expect-error a CellIs rule needs its operator
  sheet.addCfRule({ type: CfRuleType.CellIs, ranges: [rect], formula1: '1' });
  // @ts-expect-error an expression rule carries no colour scale
  sheet.addCfRule({ type: CfRuleType.Expression, ranges: [rect], formula1: '1', colorScale: { thresholds: [], colors: [] } });
  const rules: readonly ConditionalFormat[] = sheet.listCfRules();
  sheet.addCfRule(rules[0]);
  const input: ConditionalFormatInput = { type: CfRuleType.Top10, ranges: [rect], rank: 3 };
  if (input.type === CfRuleType.Top10) void input.rank;
  const name: string | null = enumName(CfRuleType, rules[0].type);
  const key: keyof typeof CfRuleType | null = formulon.enumName(CfRuleType, 3);
  void [name, key];
}

export function parseErrors(error: unknown): void {
  if (error instanceof FormulonError && error.parseError !== undefined) {
    const parse: ParseError = error.parseError;
    const late: boolean = parse.code === ParseErrorCode.EntryDoubleEquals;
    void [parse.start, parse.end, parse.line, parse.column, late];
  }
  const built = new FormulonError(1000, 'syntax', '', { code: 8, start: 0, end: 1, line: 1, column: 1 });
  const text: string = built.toString();
  void text;
}

export function recordAddresses(): void {
  sheet.addPivot({ name: 'P', cacheId: 0, anchor: 'E1' });
  sheet.addPivot({ name: 'P', cacheId: 0, anchor: { row: 0, col: 4 } });
  sheet.addValidation({ ranges: 'B2:B9', type: 3, formula1: '"a,b"' });
  sheet.setAutoFilter({ range: 'A1:C9' });
  sheet.setPrintArea('A1:F40');
  wb.addTable({ range: 'Sheet1!A1:B5', name: 'T' });
  // @ts-expect-error the anchor pair became `anchor`
  sheet.addPivot({ name: 'P', cacheId: 0, anchorRow: 0, anchorCol: 4 });
}
