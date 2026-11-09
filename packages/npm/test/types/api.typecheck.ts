// Compile-time checks of the public declarations. Never executed: `make npm-typecheck` runs tsc over it
// against the staged d.ts. A line marked `@ts-expect-error` must fail to compile; it is the rejected form.
import createFormulon, {
  AnchorKind,
  BlanksAs,
  CfRuleType,
  ChartDomain,
  ChartExtendedLayout,
  ChartAxisKind,
  ChartLegendPosition,
  ChartTickLabelPosition,
  ComparisonOperator,
  DatePeriod,
  FontScheme,
  PageOrder,
  PaperSize,
  PhoneticAlignment,
  PhoneticType,
  PivotShowAsBase,
  PivotSortOrder,
  ShapePaint,
  SlicerSortOrder,
  TableStyleElementType,
  TableTotalsRowFunction,
  WidthPlatform,
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
  mergeFunctionMetadata,
  toNative,
  type Chart,
  type ChartAxis,
  type ChartAxisRef,
  type ChartCategories,
  type ChartData,
  type ChartDataSource,
  type ChartPlot,
  type CivilTime,
  type Color,
  type ColumnBreak,
  type ConditionalFormat,
  type ConditionalFormatInput,
  type ConditionalFormatPatch,
  type AutoFilterColumnInput,
  type AutoFilterDateGroup,
  type AutoFilterSortCondition,
  type CfCellIsRuleInput,
  type CfColorScale,
  type CfColorScaleInput,
  type CfColorScaleRuleInput,
  type CfDataBar,
  type CfDataBarInput,
  type CfDataBarRuleInput,
  type CfExpressionRuleInput,
  type CfIconSet,
  type CfIconSetInput,
  type CfIconSetRuleInput,
  type CfAboveAverageRuleInput,
  type CfPlainRuleInput,
  type CfTextRuleInput,
  type CfTimePeriodRuleInput,
  type CfTop10RuleInput,
  type EffectiveBorders,
  type HeaderFooterInput,
  type HeaderFooterPatch,
  type PrintOptionsInput,
  type PrintOptionsPatch,
  type ResolvedColor,
  type SparklineGroupId,
  type SparklineGroupRef,
  type ThemeScriptFont,
  type ThemeScriptFontKey,
  type ValidationId,
  type ValidationRef,
  type DataValidation,
  type DateInput,
  type DrawingPatch,
  type DrawingProperties,
  type CellValueInput,
  type PivotDatePeriodMask,
  type Dxf,
  type DxfInput,
  type EffectiveStyle,
  type Formulon,
  type Font,
  type GradientFill,
  type CellStyle,
  type FormulonOptions,
  type NativeValue,
  type Pagination,
  type ParseError,
  type ProfileId,
  type RangeCell,
  type Pivot,
  type PivotCache,
  type PivotCacheRef,
  type PivotFieldRef,
  type PivotFilter,
  type PivotGrandTotals,
  type PivotReportLayout,
  type PivotStyleOptions,
  type RowBreak,
  type Shape,
  type ShapePatch,
  type ShapeProperties,
  type Sheet,
  type SheetVisibility,
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
  slicer.update({ anchor: { fromRow: 0, fromCol: 0, toRow: 4, toCol: 4 }, description: 'd', hidden: true });
  slicer.update({ anchor: { toCol: 6 } });
  slicer.update(slicerProps);
  // @ts-expect-error `name` cannot be updated
  slicer.update({ name: 'renamed' });
  // @ts-expect-error the anchor keys are nested under `anchor`
  slicer.update({ fromRow: 0 });
  // @ts-expect-error unknown key
  slicer.update({ bogus: true });
  const patch: SlicerPatch = {};
  void patch;
}

export function shapePatch(): void {
  shape.update({ text: 'hi', hidden: false });
  shape.update({ name: 'Callout', description: null, anchor: { kind: AnchorKind.TwoCell, toCol: 6 } });
  // `Default` removes the shape's own paint.
  shape.update({ fillKind: ShapePaint.Default, lineKind: ShapePaint.Default });
  // @ts-expect-error a shape patch has no clearable key
  shape.update({ fillKind: null });
  shape.update({ presetGeometry: shapeProps.presetGeometry, rotationDeg: shapeProps.rotationDeg });
  // @ts-expect-error unknown key
  shape.update({ bogus: 1 });
  // @ts-expect-error a string is not an anchor kind
  shape.update({ anchor: { kind: 'oneCell' } });
  // @ts-expect-error the anchor keys are nested under `anchor`
  shape.update({ anchorKind: AnchorKind.OneCell });
  const patch: ShapePatch = { rotationDeg: 90, lineWidthPt: 1.5 };
  void patch;
  // @ts-expect-error rotation is in degrees, `rotationDeg`
  shape.update({ rotation: 90 });
  // @ts-expect-error line width is in points, `lineWidthPt`
  shape.update({ lineWidthEmu: 12700 });
  const rotation: number = shapeProps.rotationDeg;
  void rotation;
  sheet.addShape({ anchor: { fromRow: 1, fromCol: 1, toRow: 4, toCol: 4 }, presetGeometry: 'rect' });
  sheet.addShape({ presetGeometry: 'rect', hidden: true, description: 'd' });
  // @ts-expect-error the anchor keys are nested under `anchor`
  sheet.addShape({ fromRow: 1, presetGeometry: 'rect' });
  sheet.addSlicer('T', 0, { anchor: { fromRow: 0, fromCol: 0, toRow: 4, toCol: 4 } });
  sheet.addSlicer(table, 'Region', { anchor: { fromRow: 0, fromCol: 0, toRow: 4, toCol: 4 } });
  sheet.addSlicer('T', 0, { caption: 'Region' });
  // @ts-expect-error a table index is sheet-local, so the slicer takes a handle or a name
  sheet.addSlicer(0, 0, { caption: 'x' });
  // @ts-expect-error a column is a position or a name
  sheet.addSlicer(table, true, {});
  // @ts-expect-error a shape is renamed through `update`
  shape.rename('Callout');
  shape.setZOrder(0);
  const snapshot: Uint8Array = shape.snapshot();
  void snapshot;
  slicer.setZOrder(0);
  // @ts-expect-error a slicer has no snapshot
  slicer.snapshot();
  // @ts-expect-error a slicer cannot be renamed
  slicer.rename('x');
  const name: string = shape.name;
  const drawingPatch: DrawingPatch = { description: 'x', hidden: true, anchor: { fromRow: 3 } };
  shape.update(drawingPatch);
  void name;
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
  const serial1904: number = formulon.convertCivilToSerial(civil, { date1904: true });
  const back: CivilTime = formulon.convertSerialToCivil(serial1904, { date1904: true });
  void [serial, back];
  // @ts-expect-error the date system is an options object
  formulon.convertCivilToSerial(civil, 1);
  // @ts-expect-error the serial is a number
  formulon.convertSerialToCivil('45352');
}

export function pivotCaches(): void {
  cache.setWorksheetSource(null);
  cache.setWorksheetSource({});
  cache.setWorksheetSource({ range: 'A1:C9', sheet: 'Data' });
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

export function pivotRefs(): void {
  const byName: PivotField = pivot.getField('Item');
  const byIndex: PivotField = pivot.getField(0);
  const dataField: PivotDataField = pivot.getDataField(0);
  const filter: PivotFilter = pivot.getFilter(0);
  const recordValue: NativeValue = cache.getRecordValue(0, 1);
  const ref: PivotCacheRef = cache;
  const numberRef: PivotCacheRef = 3;
  const fieldRef: PivotFieldRef = 'Item';
  wb.pivotCache(cache);
  wb.removePivotCache(numberRef);
  sheet.addPivot({ name: 'P', cacheId: cache, anchor: 'E1' });
  pivot.updateField(fieldRef, { showDetail: false });
  pivot.setDataField(0, { fieldIndex: 'Sales' });
  pivot.setFieldAxis('Item', null);
  const props = pivot.getProperties();
  const layout: PivotReportLayout = props.reportLayout;
  const totals: PivotGrandTotals = props.grandTotals;
  const style: PivotStyleOptions | null = props.styleOptions;
  // @ts-expect-error a cache handle is not a field reference
  pivot.updateField(cache, {});
  void [byName, byIndex, dataField, filter, recordValue, ref, layout, totals, style];
}

export function pivotUpdates(): void {
  pivot.updateGrandTotals({ rows: false });
  pivot.updateOptions({ showError: true, errorCaption: null });
  pivot.updateStyleOptions({ styleName: null, showRowStripes: true });
  pivot.updateFieldItem(0, 0, { visible: false });
  pivot.setFieldSubtotalFunctions(0, [0, 1]);
  pivot.setFieldSubtotalFunctions('Item', []);
  // @ts-expect-error the subtotal functions are a set, replaced as a whole
  pivot.addFieldSubtotalFunction(0, 0);
  // @ts-expect-error an empty list clears
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
  // @ts-expect-error renamed to `setFieldSubtotalFunctions`
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
  table.setAutoFilter({ range: rect, columns: [{ columnOffset: 0, kind: 3, topValue: 2 }], sort: { range: rect } });
  sheet.setAutoFilter({ range: rect, columns: [{ columnOffset: 1, conditions: [{ operator: 0, value: '5' }], dynamicValue: null }] });
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
  wb.copySheet(0, { newName: 'Copy' });
  wb.copySheet('Sheet1');
  wb.copySheet('Sheet1', { insertBefore: sheet, newName: 'Copy' });
  wb.copySheet(0, { insertBefore: null });
  // @ts-expect-error the copy's position and name are options keys
  wb.copySheet(0, null, 'Copy');
  const added: Sheet = wb.addSheet();
  wb.addSheet(null, { insertBefore: 0 });
  wb.addSheet('X', { insertBefore: 'Sheet1' });
  wb.addSheet('X', { insertBefore: sheet });
  wb.addSheet('X', { insertBefore: null });
  // @ts-expect-error the position is an options key
  wb.addSheet('X', 0);
  const count: number = wb.sheetCount;
  const has: boolean = wb.hasSheet('Sheet1') && wb.hasSheet(0) && wb.hasSheet(sheet);
  const profiles: readonly ProfileId[] = formulon.listProfileIds();
  void [added, count, has, profiles];
  sheet.move(2);
  sheet.move('Sheet1');
  sheet.move(sheet);
  sheet.move(null);
  // @ts-expect-error the target is required (`null` appends)
  sheet.move();
  const vis: SheetVisibility = sheet.getVisibility();
  sheet.setVisibility(vis);
  // @ts-expect-error visibility left the view
  sheet.updateView({ visibility: 0 });
  // @ts-expect-error visibility left the view
  void sheet.getView().visibility;
  sheet.insertRows(3, 5);
  sheet.insertRows('3:5');
  sheet.deleteColumns('C:D');
  sheet.deleteColumns(2, 3);
  sheet.insertColumns(1, 1);
  sheet.deleteRows(0, 0);
  const newId: string = sheet.addThreadedComment({ text: 'x', personId: 'p' });
  void newId;
  // @ts-expect-error personId is required
  sheet.addThreadedComment({ text: 'x' });
  const png = new Uint8Array(8);
  formulon.getImageInfo(png);
  formulon.getImageInfo(png.buffer);
  formulon.getImageInfo(new DataView(png.buffer));
  sheet.addPicture(png.buffer);
  // @ts-expect-error a string is not bytes
  formulon.getImageInfo('abc');
  const book: number = wb.listExternalLinks()[0].bookNumber;
  void book;
  // @ts-expect-error `index` became `bookNumber`
  void wb.listExternalLinks()[0].index;
  void [toNative, mergeFunctionMetadata];
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
  const xf: number = sheet.ooxml.getCellXfIndex('A1') ?? 0;
  sheet.ooxml.setColumnXfIndex('B', xf);
  const columnStyle: number | null = sheet.ooxml.getColumnXfIndex(1);
  sheet.ooxml.setRowXfIndex('3', xf);
  sheet.ooxml.resetColumnXfIndex('B');
  sheet.ooxml.resetRowXfIndex('3');
  sheet.ooxml.resetCellXfIndex('A1');
  const bases: readonly number[] = sheet.ooxml.getRangeXfBases('A1:B2');
  sheet.ooxml.remapRangeXfs('A1:B2', new Map(bases.map((index) => [index, index])));
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
  sheet.addCfRule({ type: CfRuleType.CellIs, ranges: 'A1:A9', operator: ComparisonOperator.GreaterThan, formula1: '10', dxfIndex: 0 });
  sheet.addCfRule({ type: CfRuleType.Expression, ranges: 'A1:A9', formula1: '=A1>0', format: { font: { bold: true } } });
  // @ts-expect-error the operator is `operator`
  sheet.addCfRule({ type: CfRuleType.CellIs, ranges: 'A1:A9', op: ComparisonOperator.GreaterThan, formula1: '10' });
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
  const id: ValidationId = sheet.addValidation({ ranges: 'B1:B9', type: 3, formula1: '"a,b"' });
  const ref: ValidationRef = id;
  sheet.updateValidation(ref, { errorMessage: 'x', formula1: null, ranges: 'C1:C9' });
  sheet.updateValidation(0, { allowBlank: false, formula2: null, promptTitle: null });
  sheet.updateValidation(rule.id, rule);
  sheet.removeValidation(id);
  sheet.removeValidation(0);
  const readId: ValidationId = rule.id;
  void readId;
  // @ts-expect-error `type` cannot be cleared
  sheet.updateValidation(0, { type: null });
  // @ts-expect-error `ranges` cannot be cleared
  sheet.updateValidation(0, { ranges: null });
  // @ts-expect-error unknown key
  sheet.updateValidation(0, { bogus: 1 });
  // @ts-expect-error the id is a string or a position
  sheet.removeValidation(true);
}

export function cfRuleUpdates(): void {
  const patch: ConditionalFormatPatch = { dxfIndex: null, formula1: '=A1>1' };
  sheet.updateCfRule(0, patch);
  sheet.updateCfRule('id-1', { dxfIndex: 2 });
  sheet.updateCfRule(0, { rank: 5, bottom: true });
  sheet.updateCfRule(0, { colorScale: { thresholds: [{ type: 3 }, { type: 4 }], colors: ['#FF0000', '#00FF00'] } });
  sheet.updateCfRule(0, { priority: 2, stopIfTrue: true, ranges: 'A1:A3' });
  // @ts-expect-error the rule's type cannot change
  sheet.updateCfRule(0, { type: CfRuleType.Top10 });
  // @ts-expect-error the id cannot change
  sheet.updateCfRule(0, { id: 'other' });
  // @ts-expect-error `priority` cannot be cleared
  sheet.updateCfRule(0, { priority: null });
  // @ts-expect-error unknown key
  sheet.updateCfRule(0, { bogus: 1 });
  const read: ConditionalFormat = sheet.listCfRules()[0];
  const scale: CfColorScale | null = read.colorScale;
  const bar: CfDataBar | null = read.dataBar;
  const icons: CfIconSet | null = read.iconSet;
  const scaleInput: CfColorScaleInput | null = scale;
  const barInput: CfDataBarInput | null = bar;
  const iconsInput: CfIconSetInput | null = icons;
  const rule: CfCellIsRuleInput = { type: CfRuleType.CellIs, ranges: 'A1', operator: ComparisonOperator.Equal, formula1: '1' };
  const plain: CfPlainRuleInput = { type: CfRuleType.UniqueValues, ranges: 'A1' };
  const others: CfExpressionRuleInput | CfTop10RuleInput | CfTextRuleInput | CfTimePeriodRuleInput | CfAboveAverageRuleInput =
    { type: CfRuleType.Top10, ranges: 'A1' };
  const visual: CfColorScaleRuleInput | CfDataBarRuleInput | CfIconSetRuleInput = {
    type: CfRuleType.ColorScale,
    ranges: 'A1',
    colorScale: { thresholds: [{ type: 3 }, { type: 4 }], colors: ['#FF0000', '#00FF00'] },
  };
  void [scaleInput, barInput, iconsInput, rule, plain, others, visual];
  const sparklineId: SparklineGroupId = sheet.addSparklineGroup({ sparklines: [] });
  const groupRef: SparklineGroupRef = sparklineId;
  sheet.updateSparklineGroup(groupRef, { lineWeight: 2 });
  sheet.updateSparklineGroup(0, { manualMin: null });
  sheet.removeSparklineGroup(sparklineId);
  sheet.removeSparklineGroup(0);
  const groupId: SparklineGroupId = sheet.listSparklineGroups()[0].id;
  void groupId;
  // @ts-expect-error a sparkline group id is a string or a position
  sheet.removeSparklineGroup(true);
}

export function styleRecordSetters(): void {
  const ooxml = wb.ooxml;
  ooxml.setFill(1, ooxml.getFill(1));
  ooxml.setBorder(1, ooxml.getBorder(1));
  ooxml.setCellXf(1, ooxml.getCellXf(1));
  ooxml.setDxf(0, ooxml.getDxf(0));
  ooxml.setFont(0, ooxml.getFont(0));
  // @ts-expect-error the index comes first
  ooxml.setFill(ooxml.getFill(1), 1);
  // @ts-expect-error a border record is not a fill
  ooxml.setFill(1, ooxml.getBorder(1));
  const borders: EffectiveBorders = sheet.ooxml.getEffectiveStyle('A1').borders;
  const left: ResolvedColor = borders.left;
  // @ts-expect-error a returned record is read-only
  borders.left = left;
  const key: ThemeScriptFontKey = { scheme: FontScheme.Major, script: 'Jpan' };
  const font: ThemeScriptFont | null = wb.getThemeScriptFont(key);
  const listed: readonly ThemeScriptFont[] = wb.listThemeScriptFonts(FontScheme.Major);
  void [font, listed];
  // @ts-expect-error the argument is the font set
  wb.listThemeScriptFonts('major');
  const options: PrintOptionsInput = { gridLines: true, headings: false, horizontalCentered: false, verticalCentered: false };
  sheet.updatePrintOptions(options);
  const patchOptions: PrintOptionsPatch = { headings: true };
  const footer: HeaderFooterPatch = { oddFooter: null, scaleWithDoc: true };
  void [patchOptions, footer];
  // @ts-expect-error a flag has no cleared form
  sheet.updatePrintOptions({ gridLines: null });
  // @ts-expect-error the input requires every key
  const incomplete: HeaderFooterInput = { oddHeader: 'x' };
  void incomplete;
}

export function autoFilterRecordsAreReadOnly(): void {
  const filter = sheet.getAutoFilter();
  if (!filter) return;
  const column: AutoFilterColumn = filter.columns[0];
  // @ts-expect-error a returned column is read-only
  column.columnOffset = 1;
  const group: AutoFilterDateGroup = column.dateGroups[0];
  // @ts-expect-error a returned date group is read-only
  group.year = 2024;
  if (filter.sort) {
    // @ts-expect-error a returned sort state is read-only
    filter.sort.columnSort = true;
    const condition: AutoFilterSortCondition = filter.sort.conditions[0];
    // @ts-expect-error a returned sort condition is read-only
    condition.descending = true;
  }
  const input: AutoFilterColumnInput = { columnOffset: 0 };
  input.hiddenButton = true;
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
  chart.setLegend({ position: ChartLegendPosition.Right, overlay: false, hiddenEntries: [] });
  chart.setLegend(null);
  chart.updateLegend({ overlay: true });
  chart.updateAxis(1, { visible: false });
  chart.updateAxis({ kind: ChartAxisKind.Value }, { visible: false });
  chart.updateAxis({ kind: ChartAxisKind.Category, plot: 1 }, { visible: false });
  const axisRef: ChartAxisRef = 1;
  chart.updateAxis(axisRef, {});
  chart.updateAxis(ChartAxisKind.Value, {});
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
  const categories: ChartCategories = data.categories;
  // @ts-expect-error the resolved data is readonly
  categories.numberFormat = 'General';
  const seriesId: number = data.series[0].seriesId;
  const unresolvedId: number = data.unresolved[0].seriesId;
  void [seriesId, unresolvedId];
  // @ts-expect-error the series key is `seriesId`
  void data.series[0].index;
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
  const xfIndex: number | null = sheet.ooxml.getCellXfIndex('A1');
  const run = sheet.getCellPhoneticRuns('A1')[0];
  const span: number = run.end - run.start;
  sheet.setCellPhoneticProperties('A1', { fontIndex: 0, type: PhoneticType.Hiragana, alignment: PhoneticAlignment.Left });
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
  wb.ooxml.addCellXf({ ...xfBase });
  // @ts-expect-error the has* flags are gone
  wb.ooxml.addCellXf({ ...xfBase, hasHorizontalAlign: true });
  // @ts-expect-error the has* flags are gone
  wb.ooxml.addCellXf({ ...xfBase, hasVerticalAlign: true });
  const rule: ConditionalFormat = sheet.listCfRules()[0];
  const dxfIndex: number | null = rule.dxfIndex;
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
  const drawingProps: DrawingProperties = props;
  const anchorKind: AnchorKind = props.anchor.kind;
  const hidden: boolean = props.hidden;
  const format: number = sheet.listPictures()[0].getProperties().format;
  const info = sheet.listPictures()[0].getProperties();
  const width: number = info.widthPx;
  sheet.restoreDrawing(bytes, { assignNewId: true });
  const metadata = formulon.getFunctionMetadata('SUM');
  const signature: string | null | undefined = metadata?.signature;
  void [format, width, signature, drawingProps, anchorKind, hidden];
  // @ts-expect-error the format is a picture field
  void props.format;
  // @ts-expect-error the anchor is nested under `anchor`
  void props.fromRow;
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
  sheet.addDataTable({ range: 'B2:D6', columnInput: 'A1' });
  const input = sheet.getDataTable('B2')?.columnInput;
  void [empty, input];
  // @ts-expect-error spelled `showEmptyColumn`
  pivot.updateOptions({ showEmptyCol: true });
  // @ts-expect-error spelled `columnInput`
  sheet.addDataTable({ rect: 'B2:D6', colInput: 'A1' });
  // @ts-expect-error spelled `columnInput`
  void sheet.getDataTable('B2')?.colInput;
}

export function sharedOrdinalTables(): void {
  sheet.addCfRule({ type: CfRuleType.CellIs, ranges: 'A1:A9', operator: ComparisonOperator.Between, formula1: '1', formula2: '9' });
  sheet.addValidation({ ranges: 'B1:B9', type: 1, operator: ComparisonOperator.NotBetween, formula1: '1', formula2: '9' });
  sheet.setAutoFilter({ range: rect, columns: [{ columnOffset: 0, conditions: [{ operator: ComparisonOperator.GreaterThanOrEqual, value: '5' }] }] });
  sheet.addCfRule({ type: CfRuleType.TimePeriod, ranges: 'A1:A9', timePeriod: DatePeriod.ThisMonth });
  // @ts-expect-error the retired validation operator table is gone
  void formulon.ValidationOperator;
  // @ts-expect-error the retired CF operator table is gone
  void formulon.CfOperator;
  // @ts-expect-error the retired sparkline table is gone
  void formulon.SparklineEmptyCells;
  // @ts-expect-error renamed to DatePeriod
  void formulon.CfTimePeriod;
  // @ts-expect-error the pivot operand discriminator is gone
  void formulon.PivotFilterOperandKind;
  // @ts-expect-error 8 is past the last operator
  const operator: ComparisonOperator = 8;
  void operator;
  assertType<Equal<(typeof ComparisonOperator)['GreaterThanOrEqual'], 7>>();
  assertType<Equal<(typeof DatePeriod)['ThisMonth'], 4>>();
  assertType<Equal<(typeof BlanksAs)['Zero'], 0>>();
  assertType<Equal<(typeof PivotSortOrder)['Manual'], 0>>();
  assertType<Equal<(typeof SlicerSortOrder)['Ascending'], 0>>();
  assertType<Equal<(typeof ChartLegendPosition)['Right'], 0>>();
  sheet.addSparklineGroup({ emptyCellsAs: BlanksAs.Span });
  // @ts-expect-error a sparkline blank mode is a BlanksAs ordinal
  sheet.addSparklineGroup({ emptyCellsAs: 'gap' });
  chart.updateDisplay({ blanksAs: BlanksAs.Gap });
  // @ts-expect-error 3 is not a BlanksAs
  chart.updateDisplay({ blanksAs: 3 });
}

declare const chart: Chart;

export function stringUnionsBecameTables(): void {
  table.setColumnTotals(0, { totalsRowFunction: TableTotalsRowFunction.CountNumbers, totalsRowLabel: 'n' });
  // @ts-expect-error the totals function is an ordinal, not its token
  table.setColumnTotals(0, { totalsRowFunction: 'sum' });
  const fn: TableTotalsRowFunction = tableProps.columns[0].totalsRowFunction;
  // @ts-expect-error the totals function reads back as an ordinal
  const token: string = tableProps.columns[0].totalsRowFunction;
  void [fn, token];
  wb.setTableStyle({ name: 'S', elements: [{ type: TableStyleElementType.HeaderRow, dxfIndex: 0 }] });
  // @ts-expect-error the element type is an ordinal, not its token
  wb.setTableStyle({ name: 'S', elements: [{ type: 'headerRow' }] });
  chart.updateAxis(1, { tickLabelPosition: ChartTickLabelPosition.Low });
  // @ts-expect-error the tick label position is an ordinal, not its token
  chart.updateAxis(1, { tickLabelPosition: 'low' });
  const position: ChartTickLabelPosition | null = chart.getProperties().axes[0].tickLabelPosition;
  void position;
  const domain: ChartDomain = chart.getData().domain;
  const isXy: boolean = domain === ChartDomain.Xy;
  // @ts-expect-error the domain is an ordinal, not its token
  const named: 'xy' | 'category' = chart.getData().domain;
  void [isXy, named];
  const platform: WidthPlatform = sheet.getWidthModel().platform;
  const isMac: boolean = platform === WidthPlatform.Mac;
  // @ts-expect-error the platform is an ordinal, not its token
  const platformName: 'win' | 'mac' = sheet.getWidthModel().platform;
  void [isMac, platformName];
  sheet.addCfRule({
    type: CfRuleType.IconSet,
    ranges: [rect],
    iconSet: {
      name: IconSetName.ThreeArrows,
      thresholds: [{ type: 1, value: '33', gte: true }, { type: 1, value: '67', gte: true }],
      icons: [{ iconSet: IconSetName.NoIcons, iconIndex: 0 }, { iconSet: IconSetName.ThreeStars, iconIndex: 2 }],
    },
  });
  sheet.addCfRule({
    type: CfRuleType.IconSet,
    ranges: [rect],
    iconSet: {
      name: IconSetName.ThreeArrows,
      thresholds: [{ type: 1, value: '33', gte: true }, { type: 1, value: '67', gte: true }],
      // @ts-expect-error a custom icon level names its set by IconSetName, not by its token
      icons: [{ iconSet: '3Arrows', iconIndex: 0 }],
    },
  });
}

export function bareNumbersBecameTables(): void {
  sheet.setCellPhoneticProperties('A1', { fontIndex: 0, type: PhoneticType.NoConversion, alignment: PhoneticAlignment.Distributed });
  // @ts-expect-error 4 is not a PhoneticType
  sheet.setCellPhoneticProperties('A1', { fontIndex: 0, type: 4, alignment: 0 });
  // @ts-expect-error a number outside the table is not a PhoneticAlignment
  sheet.setCellPhoneticProperties('A1', { fontIndex: 0, type: 0, alignment: 4 });
  const order: PageOrder = sheet.getPagination().pageOrder;
  const overThen: boolean = order === PageOrder.OverThenDown;
  void overThen;
  // @ts-expect-error the page order is a PageOrder, not a bare boolean
  const flag: boolean = sheet.getPagination().pageOrder;
  void flag;
  sheet.updatePageSetup({ paperSize: PaperSize.A4 });
  sheet.updatePageSetup({ paperSize: 118 });
  const code: number = sheet.getPageSetup().paperSize;
  void code;
  assertType<Equal<(typeof PaperSize)['A6'], 70>>();
}

export function themeScriptFonts(): void {
  wb.setThemeScriptFont({ scheme: FontScheme.Major, script: 'Jpan', typeface: 'Yu Gothic' });
  // @ts-expect-error the font set is `scheme`, not a `major` flag
  wb.setThemeScriptFont({ major: true, script: 'Jpan', typeface: 'Yu Gothic' });
  wb.getThemeScriptFont({ scheme: FontScheme.Minor, script: 'Jpan' });
  const scheme: FontScheme | undefined = wb.getThemeScriptFont({ scheme: FontScheme.Minor, script: 'Jpan' })?.scheme;
  void scheme;
}

export function nativeBooleansAndDates(): void {
  const matches: boolean = sheet.evaluateCfFormula('A1', 'A1', 'A1>0');
  // @ts-expect-error the predicate result is a boolean, not a Value
  const tagged: Value = sheet.evaluateCfFormula('A1', 'A1', 'A1>0');
  void [matches, tagged];
  const match = sheet.evaluateCfRange('A1:A9')[0].matches[0];
  const engaged: boolean = match.kind === CfMatchKind.DataBar && match.bar.isNegative && match.bar.border !== null && match.bar.gradient;
  if (match.kind === CfMatchKind.DataBar) {
    // @ts-expect-error the flags are booleans
    const asNumber: number = match.bar.gradient;
    void asNumber;
  }
  void engaged;
  const pivotCell = pivot.getLayout().cells[0];
  const cellValue: NativeValue = pivotCell.value;
  // @ts-expect-error a pivot cell has no isDate; its numberFormat tells a date label
  void pivotCell.isDate;
  void cellValue;
  const shared = cache.listFieldSharedItems(0)[0];
  const sharedValue: NativeValue = shared.value;
  void sharedValue;

  const civil: DateInput = { year: 2024, month: 1, day: 31, hour: 0, minute: 0, second: 0 };
  const date: DateInput = new Date(2024, 0, 31);
  wb.setPinnedNow(date);
  wb.setPinnedNow(civil);
  wb.setPinnedNow(null);
  // @ts-expect-error a serial number is not a DateInput
  wb.setPinnedNow(45322);
  sheet.evaluateCfRange('A1:A9', { today: new Date() });
  sheet.evaluateCfRange('A1:A9', { today: new Date(), timePeriodRules: false });
  // @ts-expect-error the date basis is `options.today`, not a serial
  sheet.evaluateCfRange('A1:A9', 45322);
  sheet.addThreadedComment({ id: '{A}', text: 't', personId: 'p', created: new Date() });
  sheet.addThreadedComment({ id: '{A}', text: 't', personId: 'p', created: civil });
  // @ts-expect-error `created` is a DateInput, not a timestamp string
  sheet.addThreadedComment({ id: '{A}', text: 't', personId: 'p', created: '2024-01-31T00:00:00.00' });
  const read: CivilTime | null = sheet.listThreadedComments()[0].created;
  void read;
  cache.setRecordValue(0, 0, new Date(2024, 0, 31));
  cache.addFieldSharedItem(0, new Date(2024, 0, 31));
  // @ts-expect-error a Date is a date; there is no flag
  cache.setRecordValue(0, 0, 45322, { date: true });
  pivot.setFieldDateGroup(0, { periods: 16 | 64, start: new Date(2024, 0, 1), end: civil });
  pivot.setFieldDateGroup(0, null);
  // @ts-expect-error the grouping is removed with `setFieldDateGroup(field, null)`
  pivot.clearFieldDateGroup(0);
  const mask: PivotDatePeriodMask = pivot.listFields()[0].datePeriods;
  void mask;
  // @ts-expect-error a bound is a DateInput, not a serial
  pivot.setFieldDateGroup(0, { periods: 16, start: 45292 });
}

export function absenceReads(): void {
  const spill = sheet.getSpill('A1');
  const blocked: boolean | undefined = spill?.blocked;
  const spillRange: CellRect | undefined = spill?.range;
  // @ts-expect-error the read is `getSpill`
  sheet.getSpillInfo('A1');
  const rank: number | null = sheet.listCfRules()[0].rank;
  const gradient: GradientFill | null = wb.ooxml.getDxf(0).fill?.gradient ?? null;
  const rotation: number | null = wb.ooxml.getCellXf(0).textRotation;
  const colour: Color | null = sheet.listSparklineGroups()[0].colorHigh;
  const fieldName: string | null = pivot.getLayout().cells[0].fieldName;
  wb.addPivotCache(null);
  wb.addPivotCache(7);
  const previous: number = PivotShowAsBase.Previous;
  const next: PivotShowAsBase = PivotShowAsBase.Next;
  pivot.setDataField(0, { fieldIndex: 1, showAs: 7, showAsBaseField: 0, showAsBaseItem: PivotShowAsBase.Previous });
  // @ts-expect-error the sentinel is a member of `PivotShowAsBase`
  void formulon.PIVOT_SHOW_AS_BASE_PREVIOUS;
  void [blocked, spillRange, rank, gradient, rotation, colour, fieldName, previous, next];
}

export function writerUnions(): void {
  const input: CellValueInput = new ExcelError(1);
  sheet.validateValue('B1', 'Open');
  sheet.validateValue(0, 1, 42);
  sheet.validateValue('B1', input);
  sheet.setError('A1', new ExcelError(2));
  sheet.setError(0, 0, 2);
  // @ts-expect-error a string is neither an ErrorCode nor an ExcelError
  sheet.setError('A1', '#N/A');
}

export function pivotFilterOperands(): void {
  pivot.addFilter({ fieldIndex: 'Item', type: PivotFilterKind.ValueTopCount, value: 10 });
  pivot.addFilter({ fieldIndex: 'Item', type: PivotFilterKind.LabelContains, value: 'a' });
  pivot.addFilter({ fieldIndex: 'Item', type: PivotFilterKind.ValueBetween, value: 1, valueHigh: 9 });
  pivot.addFilter({ fieldIndex: 'Item', type: PivotFilterKind.LabelBetween, value: 'a', valueHigh: null });
  // @ts-expect-error the operand discriminator and per-kind payload fields are gone
  pivot.addFilter({ fieldIndex: 'Item', type: PivotFilterKind.ValueTopCount, valueKind: 0, valueInt: 10 });
  // @ts-expect-error the text payload is `value`
  pivot.addFilter({ fieldIndex: 'Item', type: PivotFilterKind.LabelContains, valueText: 'a' });
  pivot.addFilter({ fieldIndex: 2, type: PivotFilterKind.LabelContains, value: 'a' });
  // @ts-expect-error the filter's axis is its field's
  pivot.addFilter({ axis: 0, fieldIndex: 2, type: PivotFilterKind.LabelContains, value: 'a' });
  const filter = pivot.listFilters()[0];
  const filterField: number = filter.fieldIndex;
  void filterField;
  const value: number | string | null = filter.value;
  const high: number | string | null = filter.valueHigh;
  void [value, high, PivotFilterKind.Today, PivotFilterKind.M12];
  // @ts-expect-error the relative-period members dropped their Date prefix
  void PivotFilterKind.DateToday;
}

import {
  ChartBarDirection,
  CfMatchKind,
  FillPattern,
  FilterKind,
  PivotAggregation,
  PivotAxis,
  type ArrayFormula,
  type ArrayFormulaInput,
  type AutoFilterColumn,
  type AutoFilterCondition,
  type AutoFilterConditionInput,
  type BorderInput,
  type CellAddress,
  type CellRect,
  type CellStyleInput,
  type CellXf,
  type CellXfInput,
  type CfColorScaleMatch,
  type CfDataBarMatch,
  type CfFormatMatch,
  type CfIcon,
  type CfIconSetMatch,
  type CfMatch,
  type CfMatchBar,
  type CfMatchIcon,
  type CfValueObject,
  type ChartFilter,
  type ChartFilterInput,
  type ChartInsertInput,
  type ChartPointPatch,
  type ChartProperties,
  type ChartSource,
  type ChartSourceInput,
  type Comment,
  type CommentInput,
  type DataTable,
  type DataTableInput,
  type DateSystemOptions,
  type DefinedName,
  type DefinedNameInput,
  type FillInput,
  type FontInput,
  type FontPatch,
  type FillPatch,
  type MergedFunctionMetadata,
  type Person,
  type PersonInput,
  type PivotDataField,
  type PivotDataFieldInput,
  type PivotField,
  type PivotFieldInput,
  type PivotFieldPatch,
  type PivotWorksheetSource,
  type PivotWorksheetSourceInput,
  type PrintTitles,
  type PrintTitlesInput,
  type ProtectedRange,
  type ProtectedRangeInput,
  type SheetRect,
  type SheetSelection,
  type SheetSelectionInput,
  type SheetView,
  type SheetViewPatch,
  type Sparkline,
  type SparklineGroup,
  type SparklineGroupInput,
  type SparklineInput,
  type TableColumn,
  type TableColumnTotalsInput,
  type TableStyle,
  type TableStyleInput,
  type ThreadedComment,
  type ThreadedCommentInput,
  type WorkbookProtection,
  type WorkbookProtectionInput,
} from '@libraz/formulon';

declare const dateRect: CellRect;
declare const sheetRect: SheetRect;
declare const address: CellAddress;
declare const font: Font;
declare const fill: import('@libraz/formulon').Fill;
declare const border: import('@libraz/formulon').Border;
declare const xf: CellXf;
declare const person: Person;
declare const comment: Comment;
declare const protectedRange: ProtectedRange;
declare const protection: WorkbookProtection;
declare const tableStyle: TableStyle;
declare const cellStyle: CellStyle;
declare const arrayFormula: ArrayFormula;
declare const printTitles: PrintTitles;
declare const selection: SheetSelection;
declare const view: SheetView;
declare const threaded: ThreadedComment;
declare const sparklineGroup: SparklineGroup;
declare const sparkline: Sparkline;
declare const dataTable: DataTable;
declare const definedName: DefinedName;
declare const chartProps: ChartProperties;
declare const pivotField: PivotField;
declare const dataField: PivotDataField;
declare const worksheetSource: PivotWorksheetSource;
declare const tableColumn: TableColumn;
declare const autoFilterColumn: AutoFilterColumn;
declare const match: CfMatch;

/** Every record a getter returns is assignable to the input the matching setter takes. */
export function readRecordsPassBackAsInput(): void {
  const asFont: FontInput = font;
  const asFill: FillInput = fill;
  const asBorder: BorderInput = border;
  const asXf: CellXfInput = xf;
  const asPerson: PersonInput = person;
  const asComment: CommentInput = comment;
  const asProtected: ProtectedRangeInput = protectedRange;
  const asProtection: WorkbookProtectionInput = protection;
  const asTableStyle: TableStyleInput = tableStyle;
  const asCellStyle: CellStyleInput = cellStyle;
  const asArray: ArrayFormulaInput = arrayFormula;
  const asTitles: PrintTitlesInput = printTitles;
  const asSelection: SheetSelectionInput = selection;
  const asView: SheetViewPatch = view;
  const asThread: ThreadedCommentInput = threaded;
  const asSparkline: SparklineInput = sparkline;
  const asGroup: SparklineGroupInput = sparklineGroup;
  const asTable: DataTableInput = dataTable;
  const asName: DefinedNameInput = definedName;
  const asSource: ChartSourceInput = chartProps.source!;
  const asFilter: ChartFilterInput = chartProps.filter;
  const asField: PivotFieldPatch = pivotField;
  const asDataField: PivotDataFieldInput = dataField;
  const asWorksheetSource: PivotWorksheetSourceInput = worksheetSource;
  const asTotals: TableColumnTotalsInput = tableColumn;
  const asAutoFilterCondition: AutoFilterConditionInput = autoFilterColumn.conditions[0];
  void [asFont, asFill, asBorder, asXf, asPerson, asComment, asProtected, asProtection, asTableStyle, asCellStyle];
  void [asArray, asTitles, asSelection, asView, asThread, asSparkline, asGroup, asTable, asName, asSource, asFilter];
  void [asField, asDataField, asWorksheetSource, asTotals, asAutoFilterCondition];
  wb.ooxml.addFont(font);
  wb.ooxml.addFill(fill);
  wb.ooxml.addBorder(border);
  wb.ooxml.addCellXf(xf);
  wb.addPerson(person);
  wb.setWorkbookProtection(protection);
  wb.setTableStyle(tableStyle);
  wb.setCellStyle(cellStyle);
  wb.setDefinedName(definedName);
  sheet.setComment('A1', comment);
  sheet.addProtectedRange(protectedRange);
  sheet.setArrayFormula(arrayFormula);
  sheet.setPrintTitles(printTitles);
  sheet.setSelection(selection);
  sheet.updateView(view);
  sheet.addThreadedComment(threaded);
  sheet.addSparklineGroup(sparklineGroup);
  sheet.addDataTable(dataTable);
  table.setColumnTotals(0, tableColumn);
  chart.setSource(chartProps.source!);
  chart.setFilter(chartProps.filter);
  pivot.updateField(0, pivotField);
  pivot.setDataField(0, dataField);
  cache.setWorksheetSource(worksheetSource);
}

export function writableOnlyInputs(): void {
  wb.ooxml.addFont({ name: 'Arial', color: '#FF0000' });
  wb.ooxml.addFill({ pattern: FillPattern.Solid, foreground: '#FFC7CE' });
  wb.ooxml.addDxf({ fill: { pattern: FillPattern.Solid, foreground: '#FFC7CE' } });
  wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 0 });
  // @ts-expect-error the has* flags are gone: an omitted or null attribute is absent
  wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 0, hasAlignment: true });
  wb.addPerson({ id: '{00000000-0000-4000-8000-000000000001}', displayName: 'A' });
  sheet.setComment('B2', { text: 'check this' });
  sheet.addProtectedRange({ name: 'p', ranges: 'A1:B2' });
  sheet.addProtectedRange({ name: 'p', ranges: ['A1', dateRect] });
  wb.setWorkbookProtection({ lockStructure: true });
  wb.setTableStyle({ name: 'S' });
  wb.setCellStyle({ name: 'Mine' } as CellStyleInput);
  sheet.setArrayFormula({ range: 'A5:A6', formula: '=SEQUENCE(2)' });
  sheet.setPrintTitles({ rows: '1:2', columns: 'A:B' });
  sheet.setPrintTitles(null);
  sheet.setSelection({ active: 'C3', ranges: 'A1:D5' });
  sheet.setSelection({ active: address });
  sheet.updateView({ topLeft: 'C5', paneTopLeft: address });
  sheet.addThreadedComment({ id: 'g', text: 't', personId: 'p', cell: 'B2' });
  sheet.addSparklineGroup({ sparklines: [{ source: 'Sheet1!A1:D1', location: 'E1' }], dateAxisRange: sheetRect });
  sheet.addDataTable({ range: 'B2:D6', rowInput: { row: 0, col: 0 }, columnInput: 'A1' });
  table.setColumnTotals(0, { totalsRowFunction: 0, totalsRowLabel: 'Total' });
  cache.setWorksheetSource({ range: 'A1:C9', sheet: 'Data' });
  cache.setWorksheetSource({ range: sheetRect });
  cache.setWorksheetSource(null);
  pivot.addDataField({ name: 'n', fieldIndex: 0 });
  // @ts-expect-error the old keys of the renamed records are unknown
  sheet.addDataTable({ rect: 'B2:D6', rowInput: 'A1' });
  // @ts-expect-error `function` became totalsRowFunction
  table.setColumnTotals(0, { function: 0 });
  // @ts-expect-error `ref` became `range`
  cache.setWorksheetSource({ ref: 'A1:C9' });
  // @ts-expect-error activeRow / activeCol became `active`
  sheet.setSelection({ activeRow: 0, activeCol: 0 });
  // @ts-expect-error row / col became `cell`
  sheet.addThreadedComment({ id: 'g', text: 't', row: 0, col: 0 });
}

export function requiredInputKeys(): void {
  // @ts-expect-error the pivot axis is required
  const noAxis: PivotFieldInput = { fieldIndex: 'Product' };
  sheet.addSlicer(table, 0, { anchor: { fromRow: 1, fromCol: 3, toRow: 9, toCol: 5 } });
  void [noAxis, PivotAxis.Row, PivotAggregation.Sum];
}

export function renamedReadFields(): void {
  const a: CellRect = arrayFormula.range;
  const b: CellAddress = view.topLeft;
  const c: CellAddress = view.paneTopLeft;
  const d: CellAddress = threaded.cell;
  const e: CellAddress = selection.active;
  const f: SheetRect | null = sparkline.source;
  const g: CellAddress = sparkline.location;
  const h: SheetRect | null = sparklineGroup.dateAxisRange;
  const i: number | null = definedName.scope;
  const j: SheetRect | null = worksheetSource.range;
  const k: CellRect = dataTable.range;
  const l: SheetRect = chartProps.source!.range;
  const m: readonly number[] = chartProps.legend!.hiddenEntries;
  const n: ChartFilter = chartProps.filter;
  const o: readonly AutoFilterCondition[] = autoFilterColumn.conditions;
  const p: boolean = pivotField.showDetail;
  const q: string | null = pivotField.sortByField;
  void [a, b, c, d, e, f, g, h, i, j, k, l, m, n, o, p, q];
  // @ts-expect-error `rect` became `range`
  void arrayFormula.rect;
  // @ts-expect-error `topLeftRow` became part of `topLeft`
  void view.topLeftRow;
  // @ts-expect-error `row` became part of `cell`
  void threaded.row;
  // @ts-expect-error `activeRow` became part of `active`
  void selection.activeRow;
  // @ts-expect-error `formula` became `source`
  void sparkline.formula;
  // @ts-expect-error `localSheetId` became `scope`
  void definedName.localSheetId;
  // @ts-expect-error the read record has no `present`
  void worksheetSource.present;
  // @ts-expect-error `hiddenSeries` of the legend became `hiddenEntries`
  void chartProps.legend!.hiddenSeries;
  // @ts-expect-error the filter is one record
  void chartProps.hiddenCategories;
  // @ts-expect-error the five fixed condition fields became `conditions`
  void autoFilterColumn.operator1;
}

export function chartVocabulary(): void {
  const input: ChartInsertInput = { type: ChartPlotType.Bar, barDirection: ChartBarDirection.Bar, range: 'Sheet1!A1:B3' };
  sheet.addChart(input);
  sheet.addChart({ type: ChartPlotType.Line, range: sheetRect });
  chart.setType(0, ChartPlotType.Bar, { barDirection: ChartBarDirection.Bar });
  chart.setType(0, ChartPlotType.Line);
  chart.setSource({ range: 'Sheet1!A1:B3' });
  chart.setSource({ range: sheetRect, headerRow: null });
  chart.setFilter({ hiddenSeries: [1] });
  chart.updateLegend({ hiddenEntries: [0] });
  chart.updateAxis(0, { minorUnit: null, minorGridlines: true, position: 0, sourceLinked: false });
  chart.updateDisplay({ style: 26 });
  chart.updateDisplay({ style: null });
  // @ts-expect-error `kind` is retired
  sheet.addChart({ kind: 0, range: 'Sheet1!A1:B3' });
  // @ts-expect-error `source` became `range`
  chart.setSource({ source: 'Sheet1!A1:B3' });
  // @ts-expect-error `hiddenSeries` of the legend became `hiddenEntries`
  chart.updateLegend({ hiddenSeries: [0] });
  // @ts-expect-error a flag key of the axis patch has no cleared form
  chart.updateAxis(0, { minorGridlines: null });
  // @ts-expect-error the position of the axis has no cleared form
  chart.updateAxis(0, { position: null });
  // @ts-expect-error a labels patch flag has no cleared form
  chart.updateSeries(0, 0, { labels: { deleted: null } });
  const pointPatch: ChartPointPatch = {};
  chart.updatePoint(0, 0, 0, pointPatch);
  chart.setPoint(0, 0, 0, null);
  void (null as unknown as ChartSource);
}

export function patchClearableKeys(): void {
  // @ts-expect-error the sheet view has no clearable key
  sheet.updateView({ zoomScale: null });
  // @ts-expect-error a table patch has no clearable key
  table.update({ headerRow: null });
  slicer.update({ style: null });
  pivot.updateField(0, { customName: null, numberFormat: null, sortByField: null });
  pivot.updateField(0, { showDetail: false, sortOrder: 1 });
  // @ts-expect-error a field flag has no cleared form
  pivot.updateField(0, { showDetail: null });
  // @ts-expect-error the sort order has no cleared form
  pivot.updateField(0, { sortOrder: null });
  pivot.updateStyleOptions({ styleName: null });
  // @ts-expect-error a pivot style flag has no cleared form
  pivot.updateStyleOptions({ showRowStripes: null });
  pivot.updateFieldItem(0, 0, { caption: null });
  // @ts-expect-error an item flag has no cleared form
  pivot.updateFieldItem(0, 0, { visible: null });
  // @ts-expect-error the per-property writers are gone
  pivot.setFieldSort(0, 1);
  // @ts-expect-error the per-property writers are gone
  pivot.setFieldShowDetail(0, false);
}

export function formatPatchKeys(): void {
  sheet.updateCellFormat('A1', { font: { family: null, scheme: FontScheme.Minor } });
  sheet.updateCellFormat('A1', { font: { outline: true, shadow: false, condense: null, extend: null, charset: 1 } });
  sheet.updateCellFormat('A1', { fill: { gradient: null } });
  const fontPatch: FontPatch = { family: null };
  const fillPatch: FillPatch = { gradient: null };
  void [fontPatch, fillPatch];
}

export function matchNarrowing(): void {
  if (match.kind === CfMatchKind.DataBar) {
    const bar: CfMatchBar = match.bar;
    const border: import('@libraz/formulon').Color | null = bar.border;
    void [bar.lengthPct, bar.isNegative, bar.gradient, bar.direction, border];
  } else if (match.kind === CfMatchKind.IconSet) {
    const icon: CfMatchIcon = match.icon;
    void [icon.setName, icon.index];
  } else if (match.kind === CfMatchKind.ColorScale) {
    const color: import('@libraz/formulon').Color = match.color;
    void color;
  } else {
    const dxf: number | null = match.dxfIndex;
    void dxf;
  }
  // @ts-expect-error `bar` exists only on a data-bar match until the kind is narrowed
  void match.bar;
  // @ts-expect-error `icon` exists only on an icon-set match until the kind is narrowed
  void match.icon;
  const parts: [CfFormatMatch, CfColorScaleMatch, CfDataBarMatch, CfIconSetMatch] = null as unknown as [
    CfFormatMatch,
    CfColorScaleMatch,
    CfDataBarMatch,
    CfIconSetMatch,
  ];
  const union: CfMatch = parts[0];
  void union;
}

export function readonlyThresholds(): void {
  const rule = sheet.listCfRules()[0];
  const threshold: CfValueObject = rule.iconSet!.thresholds[0];
  const icon: CfIcon = rule.iconSet!.icons![0];
  void [threshold, icon];
  sheet.addCfRule(rule);
}

export function autoFilterConditions(): void {
  const condition: AutoFilterConditionInput = { operator: ComparisonOperator.GreaterThan, value: '3' };
  sheet.setAutoFilter({ range: 'A1:B5', columns: [{ columnOffset: 0, kind: FilterKind.Custom, conditions: [condition] }] });
  sheet.setAutoFilter({
    range: 'A1:B5',
    // @ts-expect-error customCount and the numbered operator / value fields are gone
    columns: [{ columnOffset: 0, kind: FilterKind.Custom, customCount: 1, operator1: ComparisonOperator.Equal, value1: '3' }],
  });
}

export function dateSystemAndMetadata(): void {
  const options: DateSystemOptions = { date1904: true };
  formulon.convertCivilToSerial({ year: 1904, month: 1, day: 2 }, options);
  formulon.convertSerialToCivil(1, { date1904: true });
  // @ts-expect-error the flag is an options object
  formulon.convertSerialToCivil(1, true);
  // @ts-expect-error the flag is an options object
  formulon.convertCivilToSerial({ year: 1904, month: 1, day: 2 }, true);
  const merged: MergedFunctionMetadata = formulon.mergeFunctionMetadata(formulon.getFunctionMetadata('SUM')!, undefined, 'fr-FR');
  const localized: string = merged.localizedName;
  void localized;
}

export function xfRemapAndRangeDefaults(): void {
  sheet.ooxml.getRangeXfBasesExcluding('A1:C3', [{ row: 0, col: 0 }, 'B2']);
  sheet.ooxml.getRangeXfBasesExcluding(dateRect, [address]);
  sheet.ooxml.remapRangeXfs('A1:B2', new Map([[0, 1]]));
  // @ts-expect-error the mapping is a Map
  sheet.ooxml.remapRangeXfs('A1:B2', { 0: 1 });
  // @ts-expect-error the parallel arrays are gone
  sheet.ooxml.remapRangeXfs('A1:B2', [0], [1]);
  sheet.getValues();
  sheet.getValues(null);
  sheet.getTypedValues();
  sheet.getCellsInRange();
  sheet.getCellsInRange(null, null, 10);
  sheet.listCells();
  sheet.listCells(null);
  [...sheet.iterRows(null)];
  sheet.listThreadedComments();
  sheet.listThreadedComments('B2');
  sheet.listThreadedComments(1, 1);
}
