import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

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

test('setTableColumnTotals writes the totals-row cells', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod, { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 1 }, true);
  try {
    wb.sheet(0).table(0).setColumnTotals(0, { totalsRowLabel: 'Total' });
    wb.sheet(0).table(0).setColumnTotals(1, { totalsRowFunction: mod.TableTotalsRowFunction.Sum });
    assert.equal(wb.sheet(0).getFormula(3, 1), '=SUBTOTAL(109,[Amount])');
    wb.recalc();
    assert.equal(wb.sheet(0).getTypedValue(3, 1).number, 30);
    assert.equal(wb.sheet(0).getTypedValue(3, 0).text, 'Total');
    assert.throws(() => wb.sheet(0).table(0).setColumnTotals(1, { totalsRowFunction: 'median' }), TypeError);
    assert.throws(() => wb.sheet(0).table(0).setColumnTotals(1, { totalsRowFunction: 99 }), isFormulonError);
    const columns = wb.sheet(0).listTables()[0].getProperties().columns;
    assert.equal(columns[0].totalsRowLabel, 'Total');
    assert.equal(columns[1].totalsRowFunction, mod.TableTotalsRowFunction.Sum);
    wb.sheet(0).table(0).setColumnTotals(1, {
      totalsRowFunction: mod.TableTotalsRowFunction.Custom,
      totalsRowFormula: '=MAX(Sales[Amount])',
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
    assert.throws(() => other.rename('revenue'), isFormulonError);
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
    assert.throws(() => wb.sheet(0).table(0).renameColumn(9, 'X'), isFormulonError);
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
    assert.throws(() => wb.setTableStyle({ name: 'TableStyleMedium2' }), isFormulonError);
    assert.throws(() => wb.setTableStyle({ name: 'Bad', elements: [{ type: 'bodyRow' }] }), TypeError);
    assert.throws(() => wb.setTableStyle({ name: 'Bad', elements: [{ type: 99 }] }), isFormulonError);
    assert.equal(wb.listTableStyles().length, 1);
    wb.sheet(0).table(0).updateStyleOptions({ styleName: 'Corporate' });
    const defaults = wb.getTableStyleDefaults();
    assert.equal(defaults.tableStyle, 'TableStyleMedium2');
    assert.equal(defaults.pivotStyle, 'PivotStyleLight16');
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
    reloaded.updateTableStyleDefaults({ tableStyle: 'Corporate' });
    assert.equal(reloaded.getTableStyleDefaults().tableStyle, 'Corporate');
    reloaded.removeTableStyle('corporate');
    assert.equal(reloaded.listTableStyles().length, 0);
    assert.equal(reloaded.getTableStyleDefaults().tableStyle, 'TableStyleMedium2');
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

    wb.updateTableStyleDefaults({ tableStyle: 'TableStyleLight1' });
    wb.sheet(0).setText(0, 3, 'X');
    wb.sheet(0).setText(0, 4, 'Y');
    const light = wb.sheet(0).addTable({ range: 'D1:E3', name: 'Light' });
    assert.equal(light.getStyleOptions().styleName, 'TableStyleLight1');
    assert.equal(light.getStyleOptions().showRowStripes, true);

    // An empty styleName creates the table without a style element.
    wb.sheet(0).setText(0, 6, 'X');
    wb.sheet(0).setText(0, 7, 'Y');
    const bare = wb.sheet(0).addTable({ range: 'G1:H3', name: 'Bare', styleName: '' });
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
