// Table.insertRows / Table.insertColumns through the public entry: the omission forms, where the cells move,
// column naming and the argument and refusal mapping.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

async function withWorkbook(fn, profile) {
  const f = await getModule();
  const wb = profile ? f.Workbook.createForProfile(profile) : f.Workbook.createDefault();
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

function rejectsWith(fn, code, label, message) {
  assert.throws(
    fn,
    (e) => {
      assert.equal(e.name, 'FormulonError', `${label}: expected a FormulonError, got ${e}`);
      assert.equal(e.code, code, `${label}: code ${e.code}`);
      if (message !== undefined) assert.equal(e.message, message, label);
      return true;
    },
    label,
  );
}

// Item, Qty, Double over A1:C3 holding 1/2 and 10/20 in the first two columns.
function salesTable(ws, extra = {}) {
  ws.setValues('A1', [
    ['Item', 'Qty', 'Double'],
    [1, 10, 0],
    [2, 20, 0],
  ]);
  return ws.addTable({ range: 'A1:C3', name: 'Sales', ...extra });
}

const names = (table) => table.getProperties().columns.map((c) => c.name);
const lastRow = (table) => table.getProperties().range.rect.lastRow;
const lastCol = (table) => table.getProperties().range.rect.lastCol;

test('insertRows with no argument appends one row and moves only the table columns', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    ws.setValue('A4', 99);
    ws.setValue('D4', 77);
    assert.equal(table.insertRows(), undefined);
    assert.equal(lastRow(table), 3);
    assert.equal(ws.getValue('A4'), null);
    assert.equal(ws.getValue('A5'), 99);
    assert.equal(ws.getValue('D4'), 77);
  });
});

test('insertRows appends above the totals row', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setValues('A1', [
      ['Item', 'Qty'],
      [1, 10],
      [2, 20],
      ['Total', 30],
    ]);
    const table = ws.addTable({ range: 'A1:B4', name: 'Sales', totalsRow: true });
    table.insertRows();
    assert.equal(lastRow(table), 4);
    assert.equal(table.getProperties().totalsRow, true);
    assert.equal(ws.getValue('A3'), 2);
    assert.equal(ws.getValue('A4'), null);
  });
});

test('insertRows with first alone inserts one row there, and first with last a span', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    table.insertRows(0);
    assert.equal(lastRow(table), 3);
    assert.equal(ws.getValue('A2'), null);
    assert.equal(ws.getValue('A3'), 1);
    table.insertRows(1, 2);
    assert.equal(lastRow(table), 5);
    assert.equal(ws.getValue('A2'), null);
    assert.equal(ws.getValue('A3'), null);
    assert.equal(ws.getValue('A4'), null);
    assert.equal(ws.getValue('A5'), 1);
    assert.equal(ws.getValue('A6'), 2);
  });
});

test('insertRows treats undefined and null arguments as omitted', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    table.insertRows(undefined, undefined);
    table.insertRows(null, null);
    assert.equal(lastRow(table), 4);
    table.insertRows(2, null);
    assert.equal(lastRow(table), 5);
  });
});

test('insertColumns appends and inserts in the middle with names from the left neighbour', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    ws.setText('C1', 'Q2');
    ws.setValue('A4', 5);
    table.insertColumns();
    assert.deepEqual(names(table), ['Item', 'Qty', 'Q2', 'Q3']);
    assert.equal(ws.getValue('D1'), 'Q3');
    table.insertColumns(1, 2);
    assert.deepEqual(names(table), ['Item', 'Column1', 'Column2', 'Qty', 'Q2', 'Q3']);
    assert.equal(lastCol(table), 5);
    assert.equal(ws.getValue('D2'), 10);
    assert.equal(ws.getValue('A4'), 5);
  });
});

test('insertColumns at position 0 takes the profile word', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    table.insertColumns(0);
    assert.deepEqual(names(table), ['列1', 'Item', 'Qty', 'Double']);
    assert.equal(ws.getValue('A1'), '列1');
    assert.equal(ws.getValue('B2'), 1);
    table.insertColumns(1);
    assert.deepEqual(names(table), ['列1', '列2', 'Item', 'Qty', 'Double']);
  }, 'mac-365-ja_JP');
});

test('the argument forms are validated like the sheet-level inserts', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    for (const method of ['insertRows', 'insertColumns']) {
      assert.throws(() => table[method](1.5), RangeError, `${method} fractional`);
      assert.throws(() => table[method](0, Number.NaN), RangeError, `${method} NaN`);
      assert.throws(() => table[method](Number.POSITIVE_INFINITY), RangeError, `${method} infinite`);
      assert.throws(() => table[method](2 ** 32), RangeError, `${method} above 32 bits`);
      assert.throws(() => table[method]('1'), TypeError, `${method} string`);
      rejectsWith(() => table[method](-1), f.StatusCode.InvalidArgument, `${method} negative`);
      rejectsWith(
        () => table[method](undefined, 1),
        f.StatusCode.InvalidArgument,
        `${method} last without first`,
        `Table.${method} last needs first`,
      );
      rejectsWith(
        () => table[method](4),
        f.StatusCode.InvalidArgument,
        `${method} first past the count`,
        `Table.${method} first 4 is out of range (${method === 'insertRows' ? '2 rows' : '3 columns'})`,
      );
      rejectsWith(
        () => table[method](2, 1),
        f.StatusCode.InvalidArgument,
        `${method} last below first`,
        `Table.${method} last 1 is before first 2`,
      );
      assert.throws(() => table[method](1.5), { message: `Table.${method} first must be an integer` });
      assert.throws(() => table[method](0, 1.5), { message: `Table.${method} last must be an integer` });
    }
    assert.equal(lastRow(table), 2);
    assert.equal(lastCol(table), 2);
  });
});

test('a span past the grid and content pushed off it are refused', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    rejectsWith(
      () => table.insertRows(0, 1048576),
      f.StatusCode.InvalidArgument,
      'span past the last row',
      'Table.insertRows last 1048576 is outside the grid',
    );
    rejectsWith(
      () => table.insertColumns(0, 16384),
      f.StatusCode.InvalidArgument,
      'span past the last column',
      'Table.insertColumns last 16384 is outside the grid',
    );
    ws.setValue('B1048576', 1);
    rejectsWith(() => table.insertRows(), f.StatusCode.PreconditionFailed, 'value pushed off the grid');
    assert.equal(lastRow(table), 2);
  });
});

test('the handle of a removed table throws NotFound', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    ws.removeTable(table);
    rejectsWith(() => table.insertRows(), f.StatusCode.NotFound, 'insertRows');
    rejectsWith(() => table.insertColumns(), f.StatusCode.NotFound, 'insertColumns');
  });
});
