// Compile-time checks of the public declarations. Never executed: `make npm-typecheck` runs tsc over it
// against the staged d.ts. A line marked `@ts-expect-error` must fail to compile; it is the rejected form.
import createFormulon, {
  CfOperator,
  CfRuleType,
  ChartExtendedLayout,
  ChartLegendPosition,
  ChartPlotType,
  ChartRole,
  ChartSourceKind,
  ExcelError,
  FormulonError,
  IconSetName,
  ParseErrorCode,
  PivotFilterKind,
  ValueKind,
  enumName,
  type Chart,
  type ChartAxis,
  type ChartData,
  type ChartDataSource,
  type ChartPlot,
  type CivilTime,
  type Color,
  type ColumnBreak,
  type ConditionalFormat,
  type ConditionalFormatInput,
  type DataValidation,
  type Dxf,
  type DxfInput,
  type EffectiveStyle,
  type Formulon,
  type Font,
  type CellStyle,
  type FormulonOptions,
  type NativeValue,
  type Pagination,
  type ParseError,
  type ProfileId,
  type RangeCell,
  type Pivot,
  type PivotCache,
  type PivotStyleOptions,
  type RowBreak,
  type Shape,
  type ShapePatch,
  type ShapeProperties,
  type Sheet,
  type SheetProtection,
  type Slicer,
  type SlicerPatch,
  type SlicerProperties,
  type Table,
  type TablePatch,
  type TableProperties,
  type TableStyleOptions,
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
  table.update({ headerRow: true, totalsRow: false });
  table.update({ range: 'A1:C9' });
  table.update({ range: rect });
  table.update({ range: { sheet: 0, rect } });
  table.update(tableProps);
  // @ts-expect-error `range` is a rectangle or an A1 string
  table.update({ range: 5 });
  // @ts-expect-error the count fields became booleans named headerRow / totalsRow
  table.update({ headerRowCount: 1 });
  // @ts-expect-error unknown key
  table.update({ bogus: 1 });
  // @ts-expect-error renaming goes through `Table.rename`
  table.update({ name: 'Sales' });
  // @ts-expect-error the style moved to `Table.updateStyleOptions`
  table.update({ styleName: null });
  const header: boolean = tableProps.headerRow;
  const totals: boolean = tableProps.totalsRow;
  void [header, totals];
  // @ts-expect-error the table has one name
  void tableProps.displayName;
  // @ts-expect-error the style moved to `Table.getStyleOptions`
  void tableProps.styleName;
  assertType<
    Equal<
      TablePatch['range'],
      import('@libraz/formulon').CellRect | import('@libraz/formulon').SheetRect | string | undefined
    >
  >();
}

export function tableStyleOptions(): void {
  const info: TableStyleOptions | null = table.getStyleOptions();
  if (info !== null) {
    const name: string | null = info.styleName;
    const stripes: boolean = info.showRowStripes;
    void [name, stripes];
  }
  table.updateStyleOptions({ styleName: 'TableStyleLight9', showColumnStripes: true });
  table.updateStyleOptions({ styleName: null });
  // @ts-expect-error unknown key
  table.updateStyleOptions({ bogus: true });
  // @ts-expect-error renamed to `getStyleOptions`
  table.getStyleInfo();
  // @ts-expect-error renamed to `updateStyleOptions`
  table.updateStyleInfo({ showRowStripes: true });
  table.rename('Revenue');
  const name: string = table.name;
  const owner: Sheet = table.sheet;
  void [name, owner];
  // @ts-expect-error the table has one name
  void table.displayName;
}

export function sheetTables(): void {
  const created: Table = sheet.addTable({ range: 'A1:B5', name: 'T', columns: ['Item', 'Qty'] });
  sheet.addTable({ range: rect, name: 'U', styleName: '' });
  const byIndex: Table = sheet.table(0);
  const byName: Table = sheet.table('T');
  const byHandle: Table = sheet.table(created);
  const listed: readonly Table[] = sheet.listTables();
  sheet.removeTable('T');
  const found: Table = wb.table('T');
  void [byIndex, byName, byHandle, listed, found];
  // @ts-expect-error the table has one name
  sheet.addTable({ range: 'A1:B5', name: 'T', displayName: 'T' });
  // @ts-expect-error the workbook-wide lookup takes a name only
  wb.table(0);
  // @ts-expect-error tables are added on their sheet
  wb.addTable({ range: 'A1:B5', name: 'T' });
  // @ts-expect-error tables are listed on their sheet
  wb.listTables();
  // @ts-expect-error tables are removed on their sheet
  wb.removeTable('T');
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
  const patch: ShapePatch = { rotationDeg: 90, lineWidthPt: 1.5 };
  void patch;
  // @ts-expect-error rotation is in degrees, `rotationDeg`
  shape.update({ rotation: 90 });
  // @ts-expect-error line width is in points, `lineWidthPt`
  shape.update({ lineWidthEmu: 12700 });
  const rotation: number = shapeProps.rotationDeg;
  void rotation;
  sheet.addShape({ anchorKind: 0, presetGeometry: 'rect' });
  // @ts-expect-error an insert needs `anchorKind`
  sheet.addShape({ presetGeometry: 'rect' });
  sheet.addSlicer('T', 0, { fromRow: 0, fromCol: 0, toRow: 4, toCol: 4 });
  sheet.addSlicer(table, 0, { fromRow: 0 });
  // @ts-expect-error an insert needs the anchor (`fromRow`)
  sheet.addSlicer('T', 0, { caption: 'Region' });
  // @ts-expect-error a table index is sheet-local, so the slicer takes a handle or a name
  sheet.addSlicer(0, 0, { fromRow: 0 });
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

export function pivotFieldItems(): void {
  pivot.addFieldItem(0, { cacheIndex: 1 });
  pivot.addFieldItem(0, 'East', {});
  pivot.addFieldItem(0, { cacheIndex: 1 }, { visible: false });
  // @ts-expect-error the third argument is an options bag
  pivot.addFieldItem(0, 'East', false);
  // @ts-expect-error a cache index is passed as `{ cacheIndex }`
  pivot.addFieldItem(0, 1);
  // @ts-expect-error merged into `addFieldItem(field, { cacheIndex })`
  pivot.addFieldSharedItem(0, 1);
  cache.addFieldSharedItem(0, 1);
  pivot.setAnchor('A1');
  pivot.setAnchor({ row: 0, col: 4 });
  // @ts-expect-error the anchor is one cell
  pivot.setAnchor(rect);
}

export function pivotUpdates(): void {
  pivot.updateGrandTotals({ rows: false });
  pivot.updateOptions({ showError: true, errorCaption: null });
  pivot.updateStyleOptions({ styleName: null, showRowStripes: true });
  pivot.updateFieldItem(0, 0, { visible: false });
  pivot.addFieldSubtotalFunction(0, 0);
  pivot.clearFieldSubtotalFunctions(0);
  cache.updateOptions({ refreshOnLoad: true });
  const style: PivotStyleOptions | null = pivot.getStyleOptions();
  if (style !== null) void [style.styleName, style.showColumnHeaders, style.showColumnStripes];
  pivot.updateStyleOptions({ showColumnHeaders: true, showColumnStripes: false });
  // @ts-expect-error spelled `showColumnStripes`
  pivot.updateStyleOptions({ showColStripes: true });
  const functions: readonly number[] = pivot.listFields()[0].subtotalFunctions;
  void functions;
  // @ts-expect-error renamed to `updateGrandTotals`
  pivot.setGrandTotals({ rows: false });
  // @ts-expect-error renamed to `updateOptions`
  pivot.setOptions({ showError: true });
  // @ts-expect-error renamed to `updateStyleOptions`
  pivot.setStyleInfo({ showRowStripes: true });
  // @ts-expect-error renamed to `updateStyleOptions`
  pivot.updateStyleInfo({ showRowStripes: true });
  // @ts-expect-error renamed to `getStyleOptions`
  pivot.getStyleInfo();
  // @ts-expect-error renamed to `updateFieldItem`
  pivot.updateItem(0, 0, { visible: false });
  // @ts-expect-error renamed to `addFieldSubtotalFunction`
  pivot.addFieldSubtotalFn(0, 0);
  // @ts-expect-error renamed to `updateOptions`
  cache.setOptions({ refreshOnLoad: true });
  // @ts-expect-error the style name is `styleName`
  void style?.name;
  const top: PivotFilterKind = PivotFilterKind.ValueTopCount;
  const between: PivotFilterKind = PivotFilterKind.DateBetween;
  void [top, between];
  // @ts-expect-error renamed to `ValueTopCount`
  void PivotFilterKind.ValueTop10;
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
  sheet.setAutoFilter({ range: rect, columns: [{ columnOffset: 0 }], sort: null });
  table.setAutoFilter({ range: rect, columns: [{ columnOffset: 0, kind: 3, topValue: 2 }], sort: { ref: rect } });
  sheet.setAutoFilter({ range: rect, columns: [{ columnOffset: 1, operator1: 0, value1: '5', dynamicValue: null }] });
  // @ts-expect-error a column needs its columnOffset
  sheet.setAutoFilter({ range: rect, columns: [{ kind: 3 }] });
  // @ts-expect-error renamed to `columnOffset`
  sheet.setAutoFilter({ range: rect, columns: [{ colId: 0 }] });
  sheet.setAutoFilter(null);
  table.setAutoFilter(null);
  sheet.clearAutoFilterCriteria();
  table.clearAutoFilterCriteria();
  // @ts-expect-error removal is `setAutoFilter(null)`
  sheet.removeAutoFilter();
  // @ts-expect-error removal is `setAutoFilter(null)`
  table.removeAutoFilter();
  // @ts-expect-error renamed to `clearAutoFilterCriteria`
  sheet.clearAutoFilter();
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
  sheet.ooxml.setColumnXfIndex('B', xf);
  const columnStyle: number | null = sheet.ooxml.getColumnXfIndex(1);
  sheet.ooxml.setRowXfIndex('3', xf);
  sheet.ooxml.resetColumnXfIndex('B');
  sheet.ooxml.resetRowXfIndex('3');
  sheet.ooxml.resetCellXfIndex('A1');
  const bases: readonly number[] = sheet.ooxml.getRangeXfBases('A1:B2');
  sheet.ooxml.remapRangeXfs('A1:B2', bases, bases);
  sheet.ooxml.restoreRangeXfs(sheet.ooxml.snapshotRangeXfs('A1:B2'));
  const effective: EffectiveStyle = sheet.ooxml.getEffectiveStyle('A1');
  const locked: boolean = effective.protection.locked;
  const code: string = effective.numberFormat;
  const passthrough: boolean = wb.ooxml.hasPassthrough('xl/vbaProject.bin');
  const numberFormat: string | null = wb.ooxml.getNumberFormat(wb.ooxml.addNumberFormat('0.0%'));
  void [locked, code, passthrough, numberFormat];
  // @ts-expect-error renamed to `getColumnXfIndex`
  sheet.ooxml.getColumnStyle(1);
  // @ts-expect-error renamed to `resetCellXfIndex`
  sheet.ooxml.clearCellXfIndex('A1');
  // @ts-expect-error moved to `ooxml`
  sheet.getEffectiveStyle('A1');
  // @ts-expect-error moved to `ooxml` as `snapshotRangeXfs`
  sheet.snapshotRangeStyle('A1:B2');
  // @ts-expect-error the font colour is `fontColor`
  void effective.font;
  // @ts-expect-error renamed to `getNumberFormat`
  wb.ooxml.getNumFmtString(14);
  const sheetPr: string | null = sheet.ooxml.getSheetPrXml();
  wb.setDefaultFont(font);
  const defaultFont: Font = wb.getDefaultFont();
  void defaultFont;
  const styles: readonly CellStyle[] = wb.listCellStyles();
  wb.setCellStyle({ name: 'Accent', styleXfIndex: 0 });
  wb.removeCellStyle('Accent');
  void [styles, wb.getCellStyle('Normal')];
  // @ts-expect-error the cell-style registry moved to `Workbook`
  wb.ooxml.listCellStyles();
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
  sheet.addCfRule({ type: CfRuleType.CellIs, ranges: 'A1:A9', operator: CfOperator.GreaterThan, formula1: '10', dxfIndex: 0 });
  sheet.addCfRule({ type: CfRuleType.Expression, ranges: 'A1:A9', formula1: '=A1>0', format: { font: { bold: true } } });
  // @ts-expect-error the operator is `operator`
  sheet.addCfRule({ type: CfRuleType.CellIs, ranges: 'A1:A9', op: CfOperator.GreaterThan, formula1: '10' });
  // @ts-expect-error the format index is `dxfIndex`
  sheet.addCfRule({ type: CfRuleType.Expression, ranges: 'A1:A9', formula1: '=A1>0', dxfId: 0 });
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
  sheet.addTable({ range: 'A1:B5', name: 'T' });
  // @ts-expect-error the anchor pair became `anchor`
  sheet.addPivot({ name: 'P', cacheId: 0, anchorRow: 0, anchorCol: 4 });
}

export function sheetProtection(): void {
  sheet.setSheetProtection({ allowFormatCells: true });
  sheet.setSheetProtection({});
  sheet.setSheetProtection({ legacyPasswordHash: 'CC1A' });
  sheet.setSheetProtection(null);
  sheet.updateSheetProtection({ allowSort: true, spinCount: null });
  const current: SheetProtection | null = sheet.getSheetProtection();
  if (current !== null) {
    sheet.setSheetProtection(current);
    const select: boolean = current.allowSelectLockedCells;
    const hash: string | null = current.legacyPasswordHash;
    void [select, hash];
    // @ts-expect-error presence is the record itself (`null` when unprotected)
    void current.enabled;
    // @ts-expect-error the flags are `allow*`
    void current.formatCells;
  }
  // @ts-expect-error presence is the record itself, not an `enabled` key
  sheet.setSheetProtection({ enabled: true });
  // @ts-expect-error the legacy hash is `legacyPasswordHash`
  sheet.setSheetProtection({ legacyPassword: 'CC1A' });
  const range = sheet.listProtectedRanges()[0];
  const rangeHash: string | null = range.legacyPasswordHash;
  sheet.updateProtectedRange(0, { legacyPasswordHash: null });
  // @ts-expect-error absent is `null`, not a string
  const rangeHashText: string = range.legacyPasswordHash;
  const book = wb.getWorkbookProtection();
  const bookHash: string | null = book.workbookLegacyPasswordHash;
  void [rangeHash, rangeHashText, bookHash];
}

export function dataValidations(): void {
  sheet.addValidation({ ranges: 'B1:B9', type: 3, formula1: '"a,b"', showInCellDropdown: false, operator: 0 });
  const rule: DataValidation = sheet.listValidations()[0];
  const flags: boolean[] = [rule.allowBlank, rule.showInCellDropdown, rule.showInputMessage, rule.showErrorMessage];
  void flags;
  // @ts-expect-error renamed to `showInCellDropdown`
  sheet.addValidation({ ranges: 'B1:B9', type: 3, showDropDown: false });
  // @ts-expect-error the operator is `operator`
  void rule.op;
}

export function pageBreaks(): void {
  sheet.setRowBreak(10, {});
  sheet.setRowBreak(10, { first: 0, last: 5, manual: false });
  sheet.setRowBreak(10, null);
  sheet.setColumnBreak(3, { manual: true });
  sheet.setColumnBreak(3, null);
  const rows: readonly RowBreak[] = sheet.listRowBreaks();
  const columns: readonly ColumnBreak[] = sheet.listColumnBreaks();
  sheet.setRowBreak(rows[0].row, rows[0]);
  void [rows[0].first, rows[0].last, columns[0].col];
  const pages: Pagination = sheet.getPagination();
  const paged: readonly RowBreak[] = pages.rowBreaks;
  const pagedColumns: readonly ColumnBreak[] = pages.columnBreaks;
  const scale: number = pages.scalePct;
  const margin: number = pages.margins.leftIn;
  void [paged, pagedColumns, scale, margin];
  // @ts-expect-error the manual flag lives on each break
  void pages.rowBreakManual;
  // @ts-expect-error a break is set with `setRowBreak`
  sheet.addRowBreak(10, true);
  // @ts-expect-error removal is `setRowBreak(row, null)`
  sheet.removeRowBreak(10);
  // @ts-expect-error removal is `setColumnBreak(col, null)`
  sheet.removeColumnBreak(3);
  // @ts-expect-error the break fields are `row` / `first` / `last`
  void rows[0].id;
}

export function hyperlinks(): void {
  sheet.setHyperlink('A1', { target: 'https://example.com' });
  sheet.setHyperlink(0, 0, { target: null, location: 'Sheet2!A1' });
  sheet.setHyperlink('A1', null);
  sheet.setHyperlinkRange('A1:B2', { target: 'https://example.com' });
  sheet.setHyperlinkRange('A1:B2', null);
  // @ts-expect-error a hyperlink is set with `setHyperlink`
  sheet.addHyperlink('A1', { target: 'https://example.com' });
  // @ts-expect-error a hyperlink range is set with `setHyperlinkRange`
  sheet.addHyperlinkRange('A1:B2', { target: 'https://example.com' });
}

export function updateMembers(): void {
  sheet.updatePageSetup({ scalePct: 80, orientation: null });
  sheet.updatePageMargins({ leftIn: 0.5, headerIn: 0.3 });
  sheet.updatePrintOptions({ gridLines: true });
  sheet.updateHeaderFooter({ oddHeader: '&CReport' });
  sheet.updateView({ zoomPct: 120, splitXTwips: 0 });
  sheet.updateFormatDefaults({ defaultColWidthChars: 10, defaultRowHeightPt: null, baseColWidthChars: 8 });
  sheet.updateCellFormat('A1:C1', { font: { bold: true, sizePt: 12 }, fill: { foreground: '#FF0000' } });
  wb.updateIterative({ enabled: true });
  wb.updateTableStyleDefaults({ tableStyle: 'TableStyleLight9' });
  const setup = sheet.getPageSetup();
  const stated: boolean = setup.scalePctStated;
  const margins = sheet.getPageMargins();
  const left: number = margins.leftIn;
  const leftStated: boolean = margins.leftInStated;
  const view = sheet.getView();
  const zoom: number = view.zoomPct;
  const split: number = view.splitYTwips;
  const defaults = sheet.getFormatDefaults();
  const width: number = defaults.effectiveDefaultColWidthChars;
  const height: number = defaults.effectiveDefaultRowHeightPt;
  void [stated, left, leftStated, zoom, split, width, height];
  // @ts-expect-error renamed to `updatePageSetup`
  sheet.setPageSetup({ scalePct: 80 });
  // @ts-expect-error the scale is a percentage, `scalePct`
  sheet.updatePageSetup({ scale: 80 });
  // @ts-expect-error margins are in inches, `leftIn`
  sheet.updatePageMargins({ left: 0.5 });
  // @ts-expect-error renamed to `updateView`
  sheet.setView({ zoomPct: 120 });
  // @ts-expect-error the zoom is a percentage, `zoomPct`
  sheet.updateView({ zoomScale: 120 });
  // @ts-expect-error renamed to `updateCellFormat`
  sheet.setCellFormat('A1', {});
  // @ts-expect-error renamed to `updateFormatDefaults`
  sheet.setFormatDefaults({});
  // @ts-expect-error the default row height is in points, `defaultRowHeightPt`
  sheet.updateFormatDefaults({ defaultRowHeight: 18 });
  // @ts-expect-error the default column width is in characters, `defaultColWidthChars`
  sheet.updateFormatDefaults({ defaultColWidth: 10 });
  // @ts-expect-error the base column width is in characters, `baseColWidthChars`
  sheet.updateFormatDefaults({ baseColWidth: 8 });
}

export function resetMembers(): void {
  sheet.setColumnWidth('B', 12);
  sheet.setColumnWidth(1, 3, 12);
  sheet.setRowHeight('3', 20);
  sheet.setRowAutoHeight(2, 4, 15);
  sheet.resetRowHeight('3');
  sheet.resetColumnWidth(1, 3);
  const column: number | null = sheet.listColumnOverrides()[0].widthChars;
  const row: number | null = sheet.listRowOverrides()[0].heightPt;
  const box = sheet.getCellRectPt('A1:B2');
  const model = sheet.getWidthModel();
  void [column, row, box.widthPt, box.heightPt, model.normalFontSizePt];
  // @ts-expect-error renamed to `resetRowHeight`
  sheet.clearRowHeight('3');
  // @ts-expect-error renamed to `resetColumnWidth`
  sheet.clearColumnWidth('B');
  // @ts-expect-error the width is in characters, `widthChars`
  void sheet.listColumnOverrides()[0].width;
}

export function workbookMembers(): void {
  const active: Sheet = wb.activeSheet();
  void active;
  wb.resetTheme();
  wb.setPinnedNow(null);
  wb.setThemeColors({ colors: ['#000000', { kind: 1, argb: 0xffffffff, theme: 0, tint: 0, indexed: 0 }] });
  // @ts-expect-error renamed to `activeSheet`
  wb.getActiveSheet();
  // @ts-expect-error renamed to `resetTheme`
  wb.clearTheme();
  // @ts-expect-error the pin is released with `setPinnedNow(null)`
  wb.clearPinnedNow();
}

export function chartMembers(chart: Chart): void {
  chart.setLegend({ position: ChartLegendPosition.Right, overlay: false, hiddenSeries: [] });
  chart.setLegend(null);
  chart.updateLegend({ overlay: true });
  chart.updateAxis(1, { visible: false });
  chart.updateDisplay({ plotVisibleOnly: true });
  chart.updateSeries(0, 0, { fill: { transparencyPct: 50 }, marker: { sizePt: 7 } });
  chart.updatePlot(0, { gapWidthPct: 150, firstSliceAngleDeg: 90 });
  chart.updatePoint(0, 0, 1, { fill: { color: '#336699' } });
  chart.setPoint(0, 0, 1, null);
  const series = chart.getProperties().plots[0].series[0];
  chart.setPoint(0, 0, 1, series.points[0]);
  chart.addSeries(0, { values: 'Sheet1!B1:B5' });
  chart.addSeries(0, { y: 'Sheet1!B1:B5', x: 'Sheet1!A1:A5' });
  chart.addSeries(0, { values: series.values, name: series.name });
  chart.updateSeries(0, 0, { name: series.name, categories: series.categories });
  // @ts-expect-error `setLegend` takes the whole legend; `updateLegend` merges
  chart.setLegend({ overlay: true });
  // @ts-expect-error renamed to `updateAxis`
  chart.setAxis(1, { visible: false });
  // @ts-expect-error renamed to `updatePlot`
  chart.setPlot(0, {});
  // @ts-expect-error renamed to `updateSeries`
  chart.setSeries(0, 0, {});
  // @ts-expect-error renamed to `updateDisplay`
  chart.setDisplay({});
  // @ts-expect-error the gap width is a percentage, `gapWidthPct`
  chart.updatePlot(0, { gapWidth: 150 });
  // @ts-expect-error the first slice angle is in degrees, `firstSliceAngleDeg`
  chart.updatePlot(0, { firstSliceAngle: 90 });
  // @ts-expect-error the transparency is a percentage, `transparencyPct`
  chart.updateSeries(0, 0, { fill: { transparency: 50 } });
  // @ts-expect-error the transparency is a percentage, `transparencyPct`
  chart.updateSeries(0, 0, { line: { transparency: 50 } });
  // @ts-expect-error the marker size is in points, `sizePt`
  chart.updateSeries(0, 0, { marker: { size: 7 } });
  const props = chart.getProperties();
  const extended: boolean = props.isExtended;
  const plot: ChartPlot = props.plots[0];
  const layout: ChartExtendedLayout | null = plot.extendedLayout;
  const axis: ChartAxis = props.axes[0];
  const visible: boolean = axis.visible;
  void [extended, layout, visible, plot.gapWidthPct, plot.firstSliceAngleDeg];
  const isExtended: boolean = plot.type === ChartPlotType.Extended && layout === ChartExtendedLayout.Waterfall;
  void isExtended;
  // @ts-expect-error renamed to `isExtended`
  void props.isChartex;
  // @ts-expect-error an axis reads `visible`
  void axis.deleted;
  // @ts-expect-error renamed to `Extended`
  void ChartPlotType.Chartex;
  const data: ChartData = chart.getData();
  const role: ChartRole = data.unresolved[0].source;
  const titled: boolean = role === ChartRole.AxisTitle;
  const format: string = data.series[0].numberFormat;
  void [titled, format, data.categories.numberFormat];
  // @ts-expect-error the source of an unresolved entry is a `ChartRole`
  const sourceText: string = data.unresolved[0].source;
  void sourceText;
}

export function styleRecords(dxf: Dxf): void {
  const code: string | null = dxf.numberFormat;
  const alignment = dxf.alignment?.horizontal;
  const locked = dxf.protection?.locked;
  void [code, alignment, locked];
  // @ts-expect-error the format code is `numberFormat`
  void dxf.numFmt;
  const added: number = wb.ooxml.addDxf({ font: wb.ooxml.getFont(0) });
  wb.ooxml.addDxf({ numberFormat: '0.0%' });
  wb.ooxml.addDxf(dxf);
  const input: DxfInput = { numberFormat: null };
  void [added, input];
  // @ts-expect-error the format code is `numberFormat`
  wb.ooxml.addDxf({ numFmt: { numFmtId: 164, formatCode: '0.0%' } });
  // @ts-expect-error the alignment is a record, `alignment`
  void dxf.alignmentXml;
  const font = wb.ooxml.getFont(0);
  const size: number | null = font.sizePt;
  const fill = wb.ooxml.getFill(0);
  const foreground: Color = fill.foreground;
  const argb: number = foreground.argb;
  const xf = wb.ooxml.getCellXf(0);
  const horizontal = xf.horizontal;
  const parent: number | undefined = xf.styleXfIndex;
  const format = sheet.getCellFormat('A1');
  const xfIndex: number = sheet.ooxml.getCellXfIndex('A1');
  const run = sheet.getCellPhoneticRuns('A1')[0];
  const span: number = run.end - run.start;
  sheet.setCellPhoneticProperties('A1', { fontIndex: 0, type: 1, alignment: 1 });
  // @ts-expect-error the font is `fontIndex`
  sheet.setCellPhoneticProperties('A1', { fontId: 0, type: 1, alignment: 1 });
  const xfBase = {
    fontIndex: 0,
    fillIndex: 0,
    borderIndex: 0,
    numFmtId: 0,
    horizontal: 1,
    vertical: 2,
    wrapText: false,
  } as const;
  wb.ooxml.addCellXf({ ...xfBase, hasHorizontal: true, hasVertical: true });
  // @ts-expect-error renamed to `hasHorizontal`
  wb.ooxml.addCellXf({ ...xfBase, hasHorizontalAlign: true });
  // @ts-expect-error renamed to `hasVertical`
  wb.ooxml.addCellXf({ ...xfBase, hasVerticalAlign: true });
  const rule: ConditionalFormat = sheet.listCfRules()[0];
  const dxfIndex: number | undefined = rule.dxfIndex;
  void [size, argb, horizontal, parent, xfIndex, span, dxfIndex, format.source];
  // @ts-expect-error the font size is in points, `sizePt`
  void font.size;
  // @ts-expect-error renamed to `foreground`
  void fill.fg;
  // @ts-expect-error renamed to `argb`
  void foreground.rgb;
  // @ts-expect-error renamed to `horizontal`
  void xf.horizontalAlign;
  // @ts-expect-error the XF index is read through `Sheet.ooxml`
  void format.xfIndex;
  // @ts-expect-error renamed to `start` / `end`
  void run.sb;
}

export function drawingRecords(): void {
  const props = sheet.listDrawings()[0].getDrawingProperties();
  const format: number = props.format;
  const info = sheet.listPictures()[0].getProperties();
  const width: number = info.widthPx;
  sheet.restoreDrawing(bytes, { assignNewId: true });
  const metadata = formulon.getFunctionMetadata('SUM');
  const signature: string | undefined = metadata?.signature;
  void [format, width, signature];
  // @ts-expect-error the media path is not public
  void props.mediaPath;
  // @ts-expect-error renamed to `format`
  void props.imageFormat;
  // @ts-expect-error renamed to `widthPx`
  void info.pxWidth;
  // @ts-expect-error renamed to `assignNewId`
  sheet.restoreDrawing(bytes, { newId: true });
}

export function columnSpellings(): void {
  pivot.updateOptions({ showEmptyColumn: true });
  const empty: boolean = pivot.getOptions().showEmptyColumn;
  sheet.addDataTable({ rect: 'B2:D6', columnInput: 'A1' });
  const input = sheet.getDataTable('B2')?.columnInput;
  void [empty, input];
  // @ts-expect-error spelled `showEmptyColumn`
  pivot.updateOptions({ showEmptyCol: true });
  // @ts-expect-error spelled `columnInput`
  sheet.addDataTable({ rect: 'B2:D6', colInput: 'A1' });
  // @ts-expect-error spelled `columnInput`
  void sheet.getDataTable('B2')?.colInput;
}
