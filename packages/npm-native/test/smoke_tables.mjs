import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

const release = (wb) => wb.dispose();

// ---- Worksheet tables and custom table styles ---------------------------
//
// Table style toggles, column totals, rename and `<tableStyles>` reach the
// engine through the staged bundle like the calc core, so each gets a
// round trip here.

function salesWorkbook(mod, ref = 'A1:B3', totalsRow = false) {
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setText(0, 0, 0, 'Product').ok);
  assert.ok(wb.setText(0, 0, 1, 'Amount').ok);
  assert.ok(wb.setNumber(0, 1, 1, 10).ok);
  assert.ok(wb.setNumber(0, 2, 1, 20).ok);
  const created = wb.createTable({
    sheetIndex: 0,
    ref,
    name: 'Sales',
    columns: ['Product', 'Amount'],
    styleName: 'TableStyleLight9',
    totalsRow,
  });
  assert.ok(created.status.ok, JSON.stringify(created.status));
  assert.equal(created.index, 0);
  return wb;
}

test('tableAt reads the style toggles, row counts and columns', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  try {
    const table = wb.tableAt(0);
    assert.ok(table.status.ok, JSON.stringify(table.status));
    assert.equal(table.styleName, 'TableStyleLight9');
    assert.equal(table.hasStyleInfo, true);
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
    assert.ok(wb.updateTable(0, { showFirstColumn: true, showRowStripes: false, showColumnStripes: true }).ok);
    assert.ok(wb.updateTable(0, { styleName: 'TableStyleMedium4' }).ok);
    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    bytes = saved.bytes;
  } finally {
    release(wb);
  }
  const reloaded = mod.Workbook.loadBytes(bytes);
  try {
    const table = reloaded.tableAt(0);
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
  const wb = salesWorkbook(mod, 'A1:B4', true);
  try {
    assert.ok(wb.setTableColumnTotals(0, 0, { label: 'Total' }).ok);
    assert.ok(wb.setTableColumnTotals(0, 1, { function: 'sum' }).ok);
    assert.equal(wb.getFormula(0, 3, 1).formula, '=SUBTOTAL(109,Sales[Amount])');
    assert.ok(wb.recalc().ok);
    assert.equal(wb.getValue(0, 3, 1).value.number, 30);
    assert.equal(wb.getValue(0, 3, 0).value.text, 'Total');
    assert.equal(wb.setTableColumnTotals(0, 1, { function: 'median' }).ok, false);
    const columns = wb.tableAt(0).columns;
    assert.equal(columns[0].totalsRowLabel, 'Total');
    assert.equal(columns[1].totalsRowFunction, 'sum');
    assert.ok(wb.setTableColumnTotals(0, 1, { function: 'custom', formula: '=MAX(Sales[Amount])' }).ok);
    assert.equal(wb.tableAt(0).columns[1].totalsRowFormula, 'MAX(Sales[Amount])');
  } finally {
    release(wb);
  }
});

test('updateTable renames the table and its structured references', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  try {
    assert.ok(wb.setFormula(0, 0, 3, '=SUM(Sales[Amount])').ok);
    assert.ok(wb.updateTable(0, { name: 'Revenue' }).ok);
    const table = wb.tableAt(0);
    assert.equal(table.name, 'Revenue');
    assert.equal(table.displayName, 'Revenue');
    assert.equal(wb.getFormula(0, 0, 3).formula, '=SUM(Revenue[Amount])');
    const other = wb.createTable({ sheetIndex: 0, ref: 'F1:F2', name: 'Other', columns: ['X'] });
    assert.ok(other.status.ok);
    // A clash is refused and leaves the rest of the update unapplied.
    assert.equal(wb.updateTable(other.index, { name: 'revenue', ref: 'F1:F5' }).ok, false);
    assert.equal(wb.tableAt(other.index).ref, 'F1:F2');
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
    assert.ok(dxf.status.ok, JSON.stringify(dxf.status));
    const style = {
      name: 'Corporate',
      pivot: false,
      table: true,
      elements: [
        { type: 'wholeTable', size: 1, dxfId: dxf.index },
        { type: 'firstRowStripe', size: 2, dxfId: dxf.index },
        { type: 'totalRow', size: 1 },
      ],
    };
    assert.ok(wb.setTableStyle(style).ok);
    assert.equal(wb.setTableStyle({ name: 'TableStyleMedium2' }).ok, false);
    assert.equal(wb.setTableStyle({ name: 'Bad', elements: [{ type: 'bodyRow' }] }).ok, false);
    assert.equal(wb.tableStyleCount().value, 1);
    assert.ok(wb.updateTable(0, { styleName: 'Corporate' }).ok);
    const defaults = wb.getTableStyleDefaults();
    assert.equal(defaults.tableStyle, 'TableStyleMedium2');
    assert.equal(defaults.pivotStyle, 'PivotStyleLight16');
    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    bytes = saved.bytes;
  } finally {
    release(wb);
  }
  const reloaded = mod.Workbook.loadBytes(bytes);
  try {
    const got = reloaded.getTableStyle(0);
    assert.ok(got.status.ok, JSON.stringify(got.status));
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
    assert.equal(reloaded.tableAt(0).styleName, 'Corporate');
    assert.ok(reloaded.setTableStyleDefaults({ tableStyle: 'Corporate' }).ok);
    assert.equal(reloaded.getTableStyleDefaults().tableStyle, 'Corporate');
    assert.ok(reloaded.removeTableStyle('corporate').ok);
    assert.equal(reloaded.tableStyleCount().value, 0);
    assert.equal(reloaded.getTableStyleDefaults().tableStyle, 'TableStyleMedium2');
  } finally {
    release(reloaded);
  }
});
