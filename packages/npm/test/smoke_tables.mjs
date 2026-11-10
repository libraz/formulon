import test from 'node:test';
import { assert, expectFailure, getModule } from './smoke_support.mjs';

const release = (wb) => wb.dispose();
const failsWith = (code) => (e) => e.name === 'FormulonError' && e.code === code;

// ---- Worksheet tables and custom table styles ---------------------------
//
// Table style toggles, column totals, rename and `<tableStyles>` reach the
// engine through the staged bundle like the calc core, so each gets a
// round trip here.

function salesWorkbook(mod, rect = { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 }, totalsRow = false) {
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setText(0, 0, 'Product');
  wb.sheet(0).setText(0, 1, 'Amount');
  wb.sheet(0).setNumber(1, 1, 10);
  wb.sheet(0).setNumber(2, 1, 20);
  const created = wb.sheet(0).addTable({
    range: rect,
    name: 'Sales',
    columns: ['Product', 'Amount'],
    styleName: 'TableStyleLight9',
    totalsRow,
  });
  assert.equal(created.index, 0);
  assert.equal(wb.sheet(0).table(0), created);
  assert.equal(wb.table('Sales'), created);
  return wb;
}

test('table properties and style options read the style toggles, header and totals flags and columns', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  try {
    const table = wb.sheet(0).listTables()[0].getProperties();
    const style = wb.sheet(0).listTables()[0].getStyleOptions();
    assert.equal(style.styleName, 'TableStyleLight9');
    assert.equal(style.showRowStripes, true);
    assert.equal(style.showFirstColumn, false);
    for (const key of ['styleName', 'showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes']) {
      assert.equal(key in table, false, key);
    }
    assert.equal('displayName' in table, false);
    assert.equal(table.headerRow, true);
    assert.equal(table.totalsRow, false);
    assert.deepEqual(
      table.columns.map((c) => [c.name, c.totalsRowFunction, c.totalsRowLabel, c.totalsRowFormula]),
      [
        ['Product', mod.TableTotalsRowFunction.None, null, null],
        ['Amount', mod.TableTotalsRowFunction.None, null, null],
      ],
    );
  } finally {
    release(wb);
  }
});

test('updateStyleOptions keeps the toggles across a style name change', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  let bytes;
  try {
    wb.sheet(0).table(0).updateStyleOptions({ showFirstColumn: true, showRowStripes: false, showColumnStripes: true });
    wb.sheet(0).table(0).updateStyleOptions({ styleName: 'TableStyleMedium4' });
    const saved = wb.save();
    bytes = saved;
  } finally {
    release(wb);
  }
  const reloaded = mod.Workbook.load(bytes);
  try {
    const table = reloaded.sheet(0).listTables()[0].getStyleOptions();
    assert.equal(table.styleName, 'TableStyleMedium4');
    assert.equal(table.showFirstColumn, true);
    assert.equal(table.showRowStripes, false);
    assert.equal(table.showColumnStripes, true);
  } finally {
    release(reloaded);
  }
});

test('a column totals patch writes the totals-row cells', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod, { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 }, true);
  try {
    wb.sheet(0)
      .table(0)
      .update({ columns: [{ totalsRowLabel: 'Total' }] });
    wb.sheet(0)
      .table(0)
      .update({ columns: [{}, { totalsRowFunction: mod.TableTotalsRowFunction.Sum }] });
    assert.equal(wb.sheet(0).getFormula(3, 1), '=SUBTOTAL(109,[Amount])');
    wb.recalc();
    assert.equal(wb.sheet(0).getTypedValue(3, 1).number, 30);
    assert.equal(wb.sheet(0).getTypedValue(3, 0).text, 'Total');
    assert.throws(
      () =>
        wb
          .sheet(0)
          .table(0)
          .update({ columns: [{}, { totalsRowFunction: 'median' }] }),
      TypeError,
    );
    expectFailure(() =>
      wb
        .sheet(0)
        .table(0)
        .update({ columns: [{}, { totalsRowFunction: 99 }] }),
    );
    const columns = wb.sheet(0).listTables()[0].getProperties().columns;
    assert.equal(columns[0].totalsRowLabel, 'Total');
    assert.equal(columns[1].totalsRowFunction, mod.TableTotalsRowFunction.Sum);
    wb.sheet(0)
      .table(0)
      .update({
        columns: [
          {},
          {
            totalsRowFunction: mod.TableTotalsRowFunction.Custom,
            totalsRowFormula: '=MAX(Sales[Amount])',
          },
        ],
      });
    assert.equal(wb.sheet(0).listTables()[0].getProperties().columns[1].totalsRowFormula, '=MAX([Amount])');
  } finally {
    release(wb);
  }
});

test('rename renames the table and its structured references', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  try {
    wb.sheet(0).setFormula(0, 3, '=SUM(Sales[Amount])');
    wb.sheet(0).table(0).rename('Revenue');
    const table = wb.sheet(0).listTables()[0].getProperties();
    assert.equal(table.name, 'Revenue');
    assert.equal(wb.sheet(0).table(0).name, 'Revenue');
    assert.equal(wb.table('Revenue'), wb.sheet(0).table(0));
    assert.equal(wb.sheet(0).getFormula(0, 3), '=SUM(Revenue[Amount])');
    const other = wb.sheet(0).addTable({
      range: { firstRow: 0, firstCol: 5, lastRow: 1, lastCol: 5 },
      name: 'Other',
      columns: ['X'],
    });
    // A clash is refused and leaves the table as it was.
    expectFailure(() => other.rename('revenue'));
    assert.equal(other.name, 'Other');
    other.update({ range: { firstRow: 0, firstCol: 5, lastRow: 4, lastCol: 5 } });
    assert.deepEqual(other.getProperties().range.rect, { firstRow: 0, firstCol: 5, lastRow: 4, lastCol: 5 });
  } finally {
    release(wb);
  }
});

test('renameTableColumn renames the column and its structured references', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  try {
    wb.sheet(0).setFormula(0, 3, '=SUM(Sales[Amount])');
    const renamed = wb.sheet(0).table(0).renameColumn(1, 'Revenue');
    assert.equal(renamed, 'Revenue');
    assert.equal(wb.sheet(0).getFormula(0, 3), '=SUM(Sales[Revenue])');
    assert.equal(wb.sheet(0).getTypedValue(0, 1).text, 'Revenue');
    assert.equal(wb.sheet(0).table(0).renameColumn(1, ''), 'Column1');
    expectFailure(() => wb.sheet(0).table(0).renameColumn(9, 'X'));
  } finally {
    release(wb);
  }
});

test('custom table styles round-trip and a table can use one', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  let bytes;
  try {
    const dxf = wb.ooxml.addDxf({ font: { name: 'Calibri', bold: true } });
    const style = {
      name: 'Corporate',
      pivot: false,
      table: true,
      elements: [
        { type: mod.TableStyleElementType.WholeTable, size: 1, dxfIndex: dxf },
        { type: mod.TableStyleElementType.FirstRowStripe, size: 2, dxfIndex: dxf },
        { type: mod.TableStyleElementType.TotalRow, size: 1 },
      ],
    };
    wb.setTableStyle(style);
    expectFailure(() => wb.setTableStyle({ name: 'TableStyleMedium2' }));
    assert.throws(() => wb.setTableStyle({ name: 'Bad', elements: [{ type: 'bodyRow' }] }), TypeError);
    expectFailure(() => wb.setTableStyle({ name: 'Bad', elements: [{ type: 99 }] }));
    assert.equal(wb.listTableStyles().length, 1);
    wb.sheet(0).table(0).updateStyleOptions({ styleName: 'Corporate' });
    const defaults = wb.getTableStyleDefaults();
    assert.equal(defaults.tableStyleName, 'TableStyleMedium2');
    assert.equal(defaults.pivotStyleName, 'PivotStyleLight16');
    const saved = wb.save();
    bytes = saved;
  } finally {
    release(wb);
  }
  const reloaded = mod.Workbook.load(bytes);
  try {
    const styles = reloaded.listTableStyles();
    const got = styles[0];
    const byName = reloaded.getTableStyle('corporate');
    assert.equal(byName.name, 'Corporate');
    assert.equal(reloaded.getTableStyle('Other'), null);
    assert.equal(got.name, 'Corporate');
    assert.equal(got.pivot, false);
    assert.equal(got.table, true);
    assert.deepEqual(
      got.elements.map((e) => [e.type, e.size, e.dxfIndex]),
      [
        [mod.TableStyleElementType.WholeTable, 1, got.elements[0].dxfIndex],
        [mod.TableStyleElementType.FirstRowStripe, 2, got.elements[0].dxfIndex],
        [mod.TableStyleElementType.TotalRow, 1, null],
      ],
    );
    assert.equal(reloaded.sheet(0).listTables()[0].getStyleOptions().styleName, 'Corporate');
    reloaded.updateTableStyleDefaults({ tableStyleName: 'Corporate' });
    assert.equal(reloaded.getTableStyleDefaults().tableStyleName, 'Corporate');
    reloaded.removeTableStyle('corporate');
    assert.equal(reloaded.listTableStyles().length, 0);
    assert.equal(reloaded.getTableStyleDefaults().tableStyleName, 'TableStyleMedium2');
  } finally {
    release(reloaded);
  }
});

test('a table created without a styleName takes the workbook default style with row stripes', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    wb.sheet(0).setText(0, 0, 'X');
    wb.sheet(0).setText(0, 1, 'Y');
    const plain = wb.sheet(0).addTable({ range: 'A1:B3', name: 'Plain' });
    assert.deepEqual(plain.getStyleOptions(), {
      styleName: 'TableStyleMedium2',
      showFirstColumn: false,
      showLastColumn: false,
      showRowStripes: true,
      showColumnStripes: false,
    });

    wb.updateTableStyleDefaults({ tableStyleName: 'TableStyleLight1' });
    wb.sheet(0).setText(0, 3, 'X');
    wb.sheet(0).setText(0, 4, 'Y');
    const light = wb.sheet(0).addTable({ range: 'D1:E3', name: 'Light' });
    assert.equal(light.getStyleOptions().styleName, 'TableStyleLight1');
    assert.equal(light.getStyleOptions().showRowStripes, true);

    // styled: false creates the table without a style element.
    wb.sheet(0).setText(0, 6, 'X');
    wb.sheet(0).setText(0, 7, 'Y');
    const bare = wb.sheet(0).addTable({ range: 'G1:H3', name: 'Bare', styled: false });
    assert.equal(bare.getStyleOptions(), null);
  } finally {
    release(wb);
  }
});

test('Sheet.addTable throws TypeError for a missing or malformed range', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const rect = { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 };
    const columns = ['Product', 'Amount'];
    assert.throws(() => wb.sheet(0).addTable({ name: 'T', columns }), TypeError);
    assert.throws(() => wb.sheet(0).addTable({ range: {}, name: 'T', columns }), TypeError);
    assert.throws(() => wb.sheet(0).addTable({ range: { sheet: 0, rect }, name: 'T', columns }), TypeError);
    assert.equal('addTable' in wb, false);
    assert.equal(wb.sheet(0).listTables().length, 0);
    wb.sheet(0).addTable({ range: rect, name: 'T', columns });
    assert.equal(wb.sheet(0).listTables().length, 1);
  } finally {
    release(wb);
  }
});

test('tables live on their sheet: sheet-local index, one workbook-wide lookup by name, removal', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const other = wb.addSheet('Other');
    const a = wb.sheet(0).addTable({ range: 'A1:B3', name: 'First', columns: ['X', 'Y'] });
    const b = wb.sheet(0).addTable({ range: 'D1:E3', name: 'Second', columns: ['X', 'Y'] });
    const c = other.addTable({ range: 'A1:B3', name: 'Third', columns: ['X', 'Y'] });
    assert.deepEqual([a.index, b.index, c.index], [0, 1, 0]);
    assert.deepEqual(
      wb
        .sheet(0)
        .listTables()
        .map((t) => t.name),
      ['First', 'Second'],
    );
    assert.equal(other.table(0), c);
    assert.equal(wb.table('third'), c);
    assert.equal(c.sheet, other);
    // A name is looked up on the receiving sheet only; a position past the end fails.
    assert.throws(() => wb.sheet(0).table('Third'), failsWith(mod.StatusCode.NotFound));
    assert.throws(() => wb.sheet(0).table(5), failsWith(mod.StatusCode.IndexOutOfRange));
    // A table name is unique in the workbook.
    assert.throws(
      () => other.addTable({ range: 'D1:E3', name: 'FIRST', columns: ['X', 'Y'] }),
      failsWith(mod.StatusCode.AlreadyExists),
    );

    wb.sheet(0).removeTable('First');
    assert.equal(b.index, 0);
    assert.equal(wb.sheet(0).table(0), b);
    assert.throws(() => a.getProperties(), failsWith(mod.StatusCode.NotFound));
    assert.throws(() => wb.table('First'), failsWith(mod.StatusCode.NotFound));
    // A handle of another sheet is refused.
    assert.throws(() => wb.sheet(0).removeTable(c), failsWith(mod.StatusCode.InvalidArgument));
    other.removeTable(c);
    assert.equal(other.listTables().length, 0);
  } finally {
    release(wb);
  }
});

// ---- Header row, totals row, structural deletes and header-only creation ----

/** A1:B3 `Product` / `Amount` with data x/10 and y/20, as the table `Sales`. */
function headerWorkbook(mod, build = true) {
  const wb = mod.Workbook.createDefault();
  const ws = wb.sheet(0);
  ws.setText('A1', 'Product');
  ws.setText('B1', 'Amount');
  ws.setText('A2', 'x');
  ws.setNumber('B2', 10);
  ws.setText('A3', 'y');
  ws.setNumber('B3', 20);
  if (build) ws.addTable({ range: 'A1:B3', name: 'Sales' });
  return wb;
}

const rectOf = (table) => ({ ...table.getProperties().range.rect });
const rectLike = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });

test('header row off deletes the header cells and keeps the names; on writes them back', async () => {
  const mod = await getModule();
  const wb = headerWorkbook(mod);
  try {
    const ws = wb.sheet(0);
    const table = ws.table(0);
    ws.setFormula('D1', '=Sales[[#Headers],[Amount]]');
    wb.recalc();
    assert.equal(ws.getValue('D1'), 'Amount');

    table.update({ headerRow: false });
    wb.recalc();
    const off = table.getProperties();
    assert.equal(off.headerRow, false);
    assert.deepEqual(rectOf(table), rectLike(1, 0, 2, 1));
    assert.deepEqual(
      off.columns.map((c) => c.name),
      ['Product', 'Amount'],
    );
    assert.equal(ws.getValue('A1'), null);
    assert.equal(ws.getValue('B1'), null);
    assert.equal(ws.getValue('A2'), 'x');
    assert.equal(ws.getValue('D1').code, mod.ErrorCode.Ref);

    table.update({ headerRow: true });
    wb.recalc();
    assert.equal(table.getProperties().headerRow, true);
    assert.deepEqual(rectOf(table), rectLike(0, 0, 2, 1));
    assert.equal(ws.getValue('A1'), 'Product');
    assert.equal(ws.getValue('B1'), 'Amount');
    assert.equal(ws.getValue('B3'), 20);
    assert.equal(ws.getValue('D1'), 'Amount');
  } finally {
    release(wb);
  }
});

test('header row on shifts the data down when the row above it is occupied', async () => {
  const mod = await getModule();
  const wb = headerWorkbook(mod, false);
  try {
    const ws = wb.sheet(0);
    const table = ws.addTable({ range: 'A1:B3', name: 'Sales' });
    table.update({ headerRow: false });
    ws.setText('A1', 'note');
    table.update({ headerRow: true });
    assert.deepEqual(rectOf(table), rectLike(1, 0, 3, 1));
    assert.equal(ws.getValue('A1'), 'note');
    assert.equal(ws.getValue('A2'), 'Product');
    assert.equal(ws.getValue('B2'), 'Amount');
    assert.equal(ws.getValue('A3'), 'x');
    assert.equal(ws.getValue('B4'), 20);
  } finally {
    release(wb);
  }
});

test('totals row on writes defaults once and restores the saved settings afterwards', async () => {
  const mod = await getModule();
  const wb = headerWorkbook(mod);
  try {
    const ws = wb.sheet(0);
    const table = ws.table(0);
    table.update({ totalsRow: true });
    wb.recalc();
    let columns = table.getProperties().columns;
    assert.deepEqual(rectOf(table), rectLike(0, 0, 3, 1));
    assert.equal(columns[0].totalsRowLabel, 'Total');
    assert.equal(columns[0].totalsRowFunction, mod.TableTotalsRowFunction.None);
    assert.equal(columns[1].totalsRowFunction, mod.TableTotalsRowFunction.Sum);
    assert.equal(ws.getValue('A4'), 'Total');
    assert.equal(ws.getValue('B4'), 30);

    table.update({ columns: [{}, { totalsRowFunction: mod.TableTotalsRowFunction.Max }] });
    table.update({ totalsRow: false });
    table.update({ totalsRow: true });
    wb.recalc();
    columns = table.getProperties().columns;
    assert.equal(columns[1].totalsRowFunction, mod.TableTotalsRowFunction.Max);
    assert.equal(ws.getValue('B4'), 20);
  } finally {
    release(wb);
  }
});

test('totals row defaults count a text last column and give a one-column table the function only', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const ws = wb.sheet(0);
    ws.setValues('A1', [
      ['Item', 'Note', 'Qty'],
      ['a', 'p', 'q'],
      ['b', 'r', 's'],
    ]);
    ws.setValues('E1', [['Solo'], [1], [2]]);
    const wide = ws.addTable({ range: 'A1:C3', name: 'Wide' });
    const solo = ws.addTable({ range: 'E1:E3', name: 'Solo' });
    wide.update({ totalsRow: true });
    solo.update({ totalsRow: true });
    const last = wide.getProperties().columns[2];
    assert.equal(last.totalsRowFunction, mod.TableTotalsRowFunction.Count);
    const only = solo.getProperties().columns[0];
    assert.equal(only.totalsRowFunction, mod.TableTotalsRowFunction.Sum);
    assert.equal(only.totalsRowLabel, null);
  } finally {
    release(wb);
  }
});

test('deleting the totals row clears the totals settings, and turning it on again writes defaults', async () => {
  const mod = await getModule();
  const wb = headerWorkbook(mod);
  try {
    const ws = wb.sheet(0);
    const table = ws.table(0);
    table.update({ totalsRow: true });
    table.update({ columns: [{}, { totalsRowFunction: mod.TableTotalsRowFunction.Max }] });
    ws.deleteRows('4');
    const props = table.getProperties();
    assert.equal(props.totalsRow, false);
    assert.deepEqual(rectOf(table), rectLike(0, 0, 2, 1));
    for (const column of props.columns) {
      assert.equal(column.totalsRowFunction, mod.TableTotalsRowFunction.None);
      assert.equal(column.totalsRowLabel, null);
      assert.equal(column.totalsRowFormula, null);
    }
    table.update({ totalsRow: true });
    assert.equal(table.getProperties().columns[1].totalsRowFunction, mod.TableTotalsRowFunction.Sum);
    assert.equal(table.getProperties().columns[0].totalsRowLabel, 'Total');
  } finally {
    release(wb);
  }
});

test('a row delete over the header row is refused and every data row delete keeps the first, empty', async () => {
  const mod = await getModule();
  const wb = headerWorkbook(mod);
  try {
    const ws = wb.sheet(0);
    const table = ws.table(0);
    assert.throws(() => ws.deleteRows('1'), failsWith(mod.StatusCode.PreconditionFailed));
    assert.throws(() => ws.deleteRows('1:2'), failsWith(mod.StatusCode.PreconditionFailed));
    assert.deepEqual(rectOf(table), rectLike(0, 0, 2, 1));
    assert.equal(ws.getValue('A1'), 'Product');
    assert.equal(ws.getValue('A2'), 'x');

    ws.deleteRows('2:3');
    assert.equal(ws.listTables().length, 1);
    assert.deepEqual(rectOf(table), rectLike(0, 0, 1, 1));
    assert.equal(ws.getValue('A2'), null);
    assert.equal(ws.getValue('B2'), null);
  } finally {
    release(wb);
  }
});

test('a header-only range creates the header and one empty data row', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const ws = wb.sheet(0);
    ws.setValues('A1', [['Item', 'Qty']]);
    const table = ws.addTable({ range: 'A1:B1', name: 'Solo' });
    assert.equal(table.getProperties().headerRow, true);
    assert.deepEqual(rectOf(table), rectLike(0, 0, 1, 1));
    assert.equal(ws.getValue('A2'), null);

    ws.setValues('D1', [
      ['Item', 'Qty'],
      ['kept', 7],
    ]);
    const shifted = ws.addTable({ range: 'D1:E1', name: 'Shifted' });
    assert.deepEqual(rectOf(shifted), rectLike(0, 3, 1, 4));
    assert.equal(ws.getValue('D2'), null);
    assert.equal(ws.getValue('D3'), 'kept');
    assert.equal(ws.getValue('E3'), 7);
  } finally {
    release(wb);
  }
});

test('a header-only range on the last grid row, or with header and totals on one row, is refused', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const ws = wb.sheet(0);
    ws.setValues('A1', [['Item', 'Qty']]);
    ws.setValues('A1048576', [['Item', 'Qty']]);
    assert.throws(
      () => ws.addTable({ range: 'A1048576:B1048576', name: 'Last' }),
      failsWith(mod.StatusCode.PreconditionFailed),
    );
    assert.throws(
      () => ws.addTable({ range: 'A1:B1', name: 'Both', headerRow: true, totalsRow: true }),
      failsWith(mod.StatusCode.InvalidArgument),
    );
    assert.equal(ws.listTables().length, 0);
  } finally {
    release(wb);
  }
});

test('a range with a toggle applies the toggle first and refuses a first row that disagrees', async () => {
  const mod = await getModule();
  const wb = headerWorkbook(mod);
  try {
    const ws = wb.sheet(0);
    const table = ws.table(0);
    assert.throws(() => table.update({ headerRow: false, range: 'A1:B3' }), failsWith(mod.StatusCode.InvalidArgument));
    assert.equal(table.getProperties().headerRow, true);
    assert.deepEqual(rectOf(table), rectLike(0, 0, 2, 1));
    assert.equal(ws.getValue('A1'), 'Product');

    table.update({ headerRow: false, range: 'A2:B3' });
    assert.equal(table.getProperties().headerRow, false);
    assert.deepEqual(rectOf(table), rectLike(1, 0, 2, 1));
    assert.equal(ws.getValue('A1'), null);

    table.update({ headerRow: true, totalsRow: true, range: 'A1:B4' });
    const props = table.getProperties();
    assert.equal(props.headerRow, true);
    assert.equal(props.totalsRow, true);
    assert.deepEqual(rectOf(table), rectLike(0, 0, 3, 1));
    assert.equal(ws.getValue('A1'), 'Product');
  } finally {
    release(wb);
  }
});
