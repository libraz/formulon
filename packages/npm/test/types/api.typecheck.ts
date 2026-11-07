// Compile-time checks of the public declarations. Never executed: `make npm-typecheck` runs tsc over it
// against the staged d.ts. A line marked `@ts-expect-error` must fail to compile; it is the rejected form.
import createFormulon, {
  ChartSourceKind,
  type ChartDataSource,
  type CivilTime,
  type Formulon,
  type FormulonOptions,
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
  sheet.setSelection({ activeRow: 0, activeCol: 0, ranges: [rect] as readonly (typeof rect)[] });
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
