import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';
import { rewritePackage } from './zip_rewrite.mjs';

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
    assert.throws(
      () =>
        wb
          .sheet(0)
          .table(0)
          .update({ columns: [{}, { totalsRowFunction: 99 }] }),
      isFormulonError,
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

function addScratchPivot(wb, name = 'ByProduct', anchor = 'E10') {
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ range: 'A1:B3', sheet: 'Sheet1' });
  cache.addField('Product');
  cache.addField('Amount');
  for (const [product, amount] of [
    ['Widget', 10],
    ['Gadget', 20],
  ]) {
    const record = cache.addRecord();
    cache.setRecordValue(record, 0, product);
    cache.setRecordValue(record, 1, amount);
  }
  return wb.sheet(0).addPivot({ name, cacheId: cache.cacheId, anchor });
}

/** `bytes` with the custom style `Both` flagged as `flag` (for example `table="0"`), as a loaded file may carry it. */
function flagStyle(bytes, flag) {
  return rewritePackage(bytes, {
    'xl/styles.xml': (xml) => xml.replace('<tableStyle name="Both"', `<tableStyle name="Both" ${flag}`),
  });
}

test('addTable refuses a style that excludes tables and a slicer style, and creates no table', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  try {
    wb.setTableStyle({ name: 'PivotOnly', table: false });
    for (const styleName of ['PivotOnly', 'SlicerStyleLight1']) {
      assert.throws(
        () => wb.sheet(0).addTable({ range: 'D1:E3', name: 'Other', columns: ['A', 'B'], styleName }),
        failsWith(mod.StatusCode.InvalidArgument),
      );
    }
    assert.equal(wb.sheet(0).listTables().length, 1);
    wb.setTableStyle({ name: 'TableOnly', pivot: false });
    wb.sheet(0).addTable({ range: 'D1:E3', name: 'Other', columns: ['A', 'B'], styleName: 'TableOnly' });
    assert.equal(wb.sheet(0).listTables().length, 2);
  } finally {
    release(wb);
  }
});

test('a pivot style and the style defaults refuse a style that excludes the kind', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  try {
    const pivot = addScratchPivot(wb);
    wb.setTableStyle({ name: 'PivotOnly', table: false });
    wb.setTableStyle({ name: 'TableOnly', pivot: false });
    assert.throws(
      () => pivot.updateStyleOptions({ styleName: 'TableOnly' }),
      failsWith(mod.StatusCode.InvalidArgument),
    );
    assert.throws(
      () => pivot.updateStyleOptions({ styleName: 'SlicerStyleDark2' }),
      failsWith(mod.StatusCode.InvalidArgument),
    );
    assert.notEqual(pivot.getStyleOptions().styleName, 'TableOnly');
    pivot.updateStyleOptions({ styleName: 'PivotOnly' });
    assert.throws(
      () => wb.updateTableStyleDefaults({ tableStyleName: 'PivotOnly' }),
      failsWith(mod.StatusCode.InvalidArgument),
    );
    assert.throws(
      () => wb.updateTableStyleDefaults({ pivotStyleName: 'TableOnly' }),
      failsWith(mod.StatusCode.InvalidArgument),
    );
    wb.updateTableStyleDefaults({ tableStyleName: 'TableOnly', pivotStyleName: 'PivotOnly' });
    assert.equal(wb.getTableStyleDefaults().tableStyleName, 'TableOnly');
  } finally {
    release(wb);
  }
});

test('a loaded default that excludes its kind makes addTable and addPivot fail with PreconditionFailed', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  let saved;
  try {
    wb.setTableStyle({ name: 'Both' });
    wb.updateTableStyleDefaults({ tableStyleName: 'Both', pivotStyleName: 'Both' });
    saved = wb.save();
  } finally {
    release(wb);
  }
  const noTables = mod.Workbook.load(flagStyle(saved, 'table="0"'));
  try {
    assert.throws(
      () => noTables.sheet(0).addTable({ range: 'D1:E3', name: 'Other', columns: ['A', 'B'] }),
      failsWith(mod.StatusCode.PreconditionFailed),
    );
    assert.equal(noTables.sheet(0).listTables().length, 1);
    noTables.sheet(0).addTable({ range: 'D1:E3', name: 'Other', columns: ['A', 'B'], styleName: 'TableStyleLight1' });
    noTables.updateTableStyleDefaults({ tableStyleName: 'TableStyleLight1' });
  } finally {
    release(noTables);
  }
  const noPivots = mod.Workbook.load(flagStyle(saved, 'pivot="0"'));
  try {
    assert.throws(() => addScratchPivot(noPivots), failsWith(mod.StatusCode.PreconditionFailed));
    assert.equal(noPivots.sheet(0).listPivots().length, 0);
  } finally {
    release(noPivots);
  }
});

test('removeTableStyle re-points every user to the default of its own kind', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  let saved;
  try {
    wb.setTableStyle({ name: 'Both' });
    wb.sheet(0).table(0).updateStyleOptions({ styleName: 'Both' });
    addScratchPivot(wb).updateStyleOptions({ styleName: 'Both' });
    saved = wb.save();
  } finally {
    release(wb);
  }
  for (const flag of ['table="0"', 'pivot="0"']) {
    const loaded = mod.Workbook.load(flagStyle(saved, flag));
    try {
      assert.equal(loaded.sheet(0).table(0).getStyleOptions().styleName, 'Both');
      assert.equal(loaded.sheet(0).pivot(0).getStyleOptions().styleName, 'Both');
      loaded.removeTableStyle('Both');
      assert.equal(loaded.sheet(0).table(0).getStyleOptions().styleName, 'TableStyleMedium2');
      assert.equal(loaded.sheet(0).pivot(0).getStyleOptions().styleName, 'PivotStyleLight16');
    } finally {
      release(loaded);
    }
  }
});

test('setTableStyle refuses a built-in style name and a loaded custom style with one is dropped', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  let saved;
  try {
    for (const name of ['SlicerStyleLight1', 'TableStyleMedium9', 'PivotStyleDark3']) {
      assert.throws(() => wb.setTableStyle({ name }), failsWith(mod.StatusCode.InvalidArgument));
    }
    wb.setTableStyle({ name: 'Mine' });
    wb.sheet(0).table(0).updateStyleOptions({ styleName: 'Mine' });
    saved = wb.save();
  } finally {
    release(wb);
  }
  for (const builtin of ['SlicerStyleLight1', 'TableStyleMedium9']) {
    const renamed = rewritePackage(saved, {
      'xl/styles.xml': (xml) => xml.replace('<tableStyle name="Mine"', `<tableStyle name="${builtin}"`),
      'xl/tables/table1.xml': (xml) => xml.replace('name="Mine"', `name="${builtin}"`),
    });
    const loaded = mod.Workbook.load(renamed);
    try {
      assert.equal(loaded.listTableStyles().length, 0);
      assert.equal(loaded.sheet(0).table(0).getStyleOptions().styleName, builtin);
      const again = mod.Workbook.load(loaded.save());
      try {
        assert.equal(again.sheet(0).table(0).getStyleOptions().styleName, builtin);
      } finally {
        release(again);
      }
    } finally {
      release(loaded);
    }
  }
});

test('a table or pivot style name is refused for a kind the style excludes and for a slicer style', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  try {
    wb.setTableStyle({ name: 'PivotOnly', table: false });
    assert.throws(
      () => wb.sheet(0).table(0).updateStyleOptions({ styleName: 'PivotOnly' }),
      failsWith(mod.StatusCode.InvalidArgument),
    );
    assert.throws(
      () => wb.sheet(0).table(0).updateStyleOptions({ styleName: 'SlicerStyleLight1' }),
      failsWith(mod.StatusCode.InvalidArgument),
    );
    assert.notEqual(wb.sheet(0).table(0).getStyleOptions().styleName, 'PivotOnly');
    wb.sheet(0).table(0).updateStyleOptions({ styleName: 'PivotStyleLight16' });
    wb.setTableStyle({ name: 'TableOnly', pivot: false });
    wb.sheet(0).table(0).updateStyleOptions({ styleName: 'TableOnly' });
    assert.throws(
      () => wb.setTableStyle({ name: 'TableOnly', table: false }),
      failsWith(mod.StatusCode.PreconditionFailed),
    );
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
      got.elements.map((e) => [e.type, e.size]),
      [
        [mod.TableStyleElementType.WholeTable, 1],
        [mod.TableStyleElementType.FirstRowStripe, 2],
        [mod.TableStyleElementType.TotalRow, 1],
      ],
    );
    // A save writes one dxf per user, as Excel does: the two elements that shared
    // a dxf now name two equal records, written in reverse element order.
    const [whole, stripe, total] = got.elements;
    assert.equal(whole.dxfIndex, stripe.dxfIndex + 1);
    assert.deepEqual(reloaded.ooxml.getDxf(whole.dxfIndex), reloaded.ooxml.getDxf(stripe.dxfIndex));
    assert.equal(total.dxfIndex, null);
    assert.equal(reloaded.sheet(0).listTables()[0].getStyleOptions().styleName, 'Corporate');
    reloaded.updateTableStyleDefaults({ tableStyleName: 'Corporate' });
    assert.equal(reloaded.getTableStyleDefaults().tableStyleName, 'Corporate');
    reloaded.removeTableStyle('corporate');
    assert.equal(reloaded.listTableStyles().length, 0);
    assert.equal(reloaded.getTableStyleDefaults().tableStyleName, 'TableStyleMedium2');
    // The table that used the removed style takes the reset default.
    assert.equal(reloaded.sheet(0).listTables()[0].getStyleOptions().styleName, 'TableStyleMedium2');
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
