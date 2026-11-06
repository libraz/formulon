import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

const release = (wb) => wb.dispose();

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
  const created = wb.addTable({
    range: { sheet: 0, rect },
    name: 'Sales',
    columns: ['Product', 'Amount'],
    styleName: 'TableStyleLight9',
    totalsRow,
  });
  assert.equal(created, 0);
  return wb;
}

test('tableAt reads the style toggles, row counts and columns', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  try {
    const table = wb.listTables()[0];
    assert.equal(table.styleName, 'TableStyleLight9');
    assert.equal(table.showRowStripes, true);
    assert.equal(table.showFirstColumn, false);
    assert.equal(table.headerRowCount, 1);
    assert.equal(table.totalsRowCount, 0);
    assert.deepEqual(
      table.columns.map((c) => [c.name, c.totalsRowFunction, c.totalsRowLabel, c.totalsRowFormula]),
      [
        ['Product', 'none', '', ''],
        ['Amount', 'none', '', ''],
      ],
    );
  } finally {
    release(wb);
  }
});

test('updateTable keeps the toggles across a style name change', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  let bytes;
  try {
    wb.updateTable(0, { showFirstColumn: true, showRowStripes: false, showColumnStripes: true });
    wb.updateTable(0, { styleName: 'TableStyleMedium4' });
    const saved = wb.save();
    bytes = saved;
  } finally {
    release(wb);
  }
  const reloaded = mod.Workbook.load(bytes);
  try {
    const table = reloaded.listTables()[0];
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
    wb.setTableColumnTotals(0, 0, { label: 'Total' });
    wb.setTableColumnTotals(0, 1, { function: 'sum' });
    assert.equal(wb.sheet(0).getFormula(3, 1), '=SUBTOTAL(109,[Amount])');
    wb.recalc();
    assert.equal(wb.sheet(0).getValue(3, 1).number, 30);
    assert.equal(wb.sheet(0).getValue(3, 0).text, 'Total');
    assert.throws(() => wb.setTableColumnTotals(0, 1, { function: 'median' }), isFormulonError);
    const columns = wb.listTables()[0].columns;
    assert.equal(columns[0].totalsRowLabel, 'Total');
    assert.equal(columns[1].totalsRowFunction, 'sum');
    wb.setTableColumnTotals(0, 1, { function: 'custom', formula: '=MAX(Sales[Amount])' });
    assert.equal(wb.listTables()[0].columns[1].totalsRowFormula, 'MAX([Amount])');
  } finally {
    release(wb);
  }
});

test('updateTable renames the table and its structured references', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  try {
    wb.sheet(0).setFormula(0, 3, '=SUM(Sales[Amount])');
    wb.updateTable(0, { name: 'Revenue' });
    const table = wb.listTables()[0];
    assert.equal(table.name, 'Revenue');
    assert.equal(table.displayName, 'Revenue');
    assert.equal(wb.sheet(0).getFormula(0, 3), '=SUM(Revenue[Amount])');
    const other = wb.addTable({
      range: { sheet: 0, rect: { firstRow: 0, firstCol: 5, lastRow: 1, lastCol: 5 } },
      name: 'Other',
      columns: ['X'],
    });
    // A clash is refused and leaves the rest of the update unapplied.
    assert.throws(
      () => wb.updateTable(other, { name: 'revenue', rect: { firstRow: 0, firstCol: 5, lastRow: 4, lastCol: 5 } }),
      isFormulonError,
    );
    assert.deepEqual(wb.listTables()[other].range.rect, { firstRow: 0, firstCol: 5, lastRow: 1, lastCol: 5 });
  } finally {
    release(wb);
  }
});

test('renameTableColumn renames the column and its structured references', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  try {
    wb.sheet(0).setFormula(0, 3, '=SUM(Sales[Amount])');
    const renamed = wb.renameTableColumn(0, 1, 'Revenue');
    assert.equal(renamed, 'Revenue');
    assert.equal(wb.sheet(0).getFormula(0, 3), '=SUM(Sales[Revenue])');
    assert.equal(wb.sheet(0).getValue(0, 1).text, 'Revenue');
    assert.equal(wb.renameTableColumn(0, 1, ''), 'Column1');
    assert.throws(() => wb.renameTableColumn(0, 9, 'X'), isFormulonError);
  } finally {
    release(wb);
  }
});

test('custom table styles round-trip and a table can use one', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  let bytes;
  try {
    const dxf = wb.addDxf({ font: { name: 'Calibri', bold: true } });
    const style = {
      name: 'Corporate',
      pivot: false,
      table: true,
      elements: [
        { type: 'wholeTable', size: 1, dxfId: dxf },
        { type: 'firstRowStripe', size: 2, dxfId: dxf },
        { type: 'totalRow', size: 1 },
      ],
    };
    wb.setTableStyle(style);
    assert.throws(() => wb.setTableStyle({ name: 'TableStyleMedium2' }), isFormulonError);
    assert.throws(() => wb.setTableStyle({ name: 'Bad', elements: [{ type: 'bodyRow' }] }), isFormulonError);
    assert.equal(wb.listTableStyles().length, 1);
    wb.updateTable(0, { styleName: 'Corporate' });
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
      got.elements.map((e) => [e.type, e.size, e.dxfId]),
      [
        ['wholeTable', 1, got.elements[0].dxfId],
        ['firstRowStripe', 2, got.elements[0].dxfId],
        ['totalRow', 1, undefined],
      ],
    );
    assert.equal(reloaded.listTables()[0].styleName, 'Corporate');
    reloaded.setTableStyleDefaults({ tableStyle: 'Corporate' });
    assert.equal(reloaded.getTableStyleDefaults().tableStyle, 'Corporate');
    reloaded.removeTableStyle('corporate');
    assert.equal(reloaded.listTableStyles().length, 0);
    assert.equal(reloaded.getTableStyleDefaults().tableStyle, 'TableStyleMedium2');
  } finally {
    release(reloaded);
  }
});

test('addTable throws TypeError for a missing range or range.rect', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const rect = { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 };
    const columns = ['Product', 'Amount'];
    assert.throws(() => wb.addTable({ name: 'T', columns }), TypeError);
    assert.throws(() => wb.addTable({ range: { sheet: 0 }, name: 'T', columns }), TypeError);
    assert.throws(() => wb.addTable({ range: { rect }, name: 'T', columns }), TypeError);
    assert.equal(wb.listTables().length, 0);
    wb.addTable({ range: { sheet: 0, rect }, name: 'T', columns });
  } finally {
    release(wb);
  }
});
