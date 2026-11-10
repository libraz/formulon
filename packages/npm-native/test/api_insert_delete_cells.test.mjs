// Insert / delete cells through the public entry: every rule family and the refusal status mapping.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

async function withWorkbook(fn) {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

const rect = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });

function rejectsWith(fn, code, label) {
  assert.throws(
    fn,
    (e) => {
      assert.equal(e.name, 'FormulonError', `${label}: expected a FormulonError, got ${e}`);
      assert.equal(e.code, code, `${label}: code ${e.code}`);
      return true;
    },
    label,
  );
}

const sorted = (rects) => rects.map((r) => JSON.stringify(rect(r.firstRow, r.firstCol, r.lastRow, r.lastCol))).sort();

test('CellShift names the four directions', async () => {
  await withWorkbook((wb, f) => {
    assert.deepEqual({ ...f.CellShift }, { Down: 0, Right: 1, Up: 2, Left: 3 });
    assert.equal(Object.isFrozen(f.CellShift), true);
  });
});

test('insertCells Down moves the band and its references', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('B3', [
      [1, 2],
      [3, 4],
    ]);
    ws.setValue('B6', 9);
    ws.setFormula('D1', '=B3');
    ws.setFormula('E1', '=SUM(B3:C4)');
    ws.insertCells('B3:C4', f.CellShift.Down);
    assert.equal(ws.getValue('B3'), null);
    assert.equal(ws.getValue('C4'), null);
    assert.equal(ws.getValue('B5'), 1);
    assert.equal(ws.getValue('C6'), 4);
    assert.equal(ws.getValue('B8'), 9);
    assert.equal(ws.getFormula('D1'), '=B5');
    assert.equal(ws.getFormula('E1'), '=SUM(B5:C6)');
    assert.equal(ws.getValue('E1'), 10);
  });
});

test('insertCells Right moves the band to the right and an A1 or rect range is the same call', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('B3', [[1, 2, 3]]);
    ws.insertCells(rect(2, 1, 2, 1), f.CellShift.Right);
    assert.equal(ws.getValue('B3'), null);
    assert.equal(ws.getValue('C3'), 1);
    assert.equal(ws.getValue('E3'), 3);
  });
});

test('deleteCells Up closes the gap and moves references with it', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('B3', [[1], [2], [3], [4]]);
    ws.setFormula('D1', '=B6');
    ws.deleteCells('B3:B4', f.CellShift.Up);
    assert.equal(ws.getValue('B3'), 3);
    assert.equal(ws.getValue('B4'), 4);
    assert.equal(ws.getValue('B5'), null);
    assert.equal(ws.getFormula('D1'), '=B4');
  });
});

test('deleteCells Left turns references to deleted cells into #REF!', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('B3', [[1, 2, 3]]);
    ws.setFormula('A1', '=B3');
    ws.setFormula('A2', '=D3');
    ws.deleteCells('B3:C3', f.CellShift.Left);
    assert.equal(ws.getFormula('A1'), '=#REF!');
    assert.equal(ws.getFormula('A2'), '=B3');
    assert.equal(ws.getValue('B3'), 3);
  });
});

test('a range straddling the band is unchanged', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('A3', [[1, 2, 3, 4]]);
    ws.setFormula('F1', '=SUM(A3:D3)');
    ws.insertCells('B3:C4', f.CellShift.Down);
    assert.equal(ws.getFormula('F1'), '=SUM(A3:D3)');
    assert.equal(ws.getValue('A3'), 1);
    assert.equal(ws.getValue('D3'), 4);
  });
});

test('a CF range crossing the band splits into the moved part and the parts left behind', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.addCfRule({ type: f.CfRuleType.Expression, ranges: 'A3:D3', formula1: 'TRUE' });
    ws.insertCells('B3:C4', f.CellShift.Down);
    const ranges = ws.listCfRules().flatMap((r) => r.ranges);
    assert.deepEqual(sorted(ranges), sorted([rect(2, 0, 2, 0), rect(4, 1, 4, 2), rect(2, 3, 2, 3)]));
  });
});

test('a merge only partly inside the band is unmerged and one inside moves', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.addMerge('B3:C3');
    ws.addMerge('E3:F4');
    ws.insertCells('B3:C4', f.CellShift.Down);
    ws.insertCells('E4:E4', f.CellShift.Down);
    assert.deepEqual(sorted(ws.listMerges()), sorted([rect(4, 1, 4, 2)]));
  });
});

test('a refusal is a FormulonError PreconditionFailed and leaves the workbook unchanged', async () => {
  await withWorkbook((wb, f) => {
    const { StatusCode, CellShift } = f;
    const ws = wb.sheet(0);
    ws.setValues('A1', [
      ['Item', 'Qty'],
      ['a', 1],
      ['b', 2],
    ]);
    ws.addTable({ name: 'Sales', range: 'A1:B3' });
    rejectsWith(() => ws.insertCells('B2:B2', CellShift.Down), StatusCode.PreconditionFailed, 'table straddle');
    assert.equal(ws.getValue('B2'), 1);

    ws.setValue('H1048576', 5);
    rejectsWith(() => ws.insertCells('H1:H1', CellShift.Down), StatusCode.PreconditionFailed, 'pushed off the grid');
    assert.equal(ws.getValue('H1048576'), 5);
  });
});

test('a note, a fill or a merge pushed off the grid refuses the insert; a bold font does not', async () => {
  await withWorkbook((wb, f) => {
    const { StatusCode, CellShift } = f;
    const ws = wb.sheet(0);
    ws.setComment('XFD3', { text: 'last column' });
    rejectsWith(() => ws.insertCells('B3', CellShift.Right), StatusCode.PreconditionFailed, 'note, last column');
    assert.equal(ws.getComment('XFD3').text, 'last column');
    ws.setComment('XFD3', null);
    ws.setComment('C1048576', { text: 'last row' });
    rejectsWith(() => ws.insertRows(2, 2), StatusCode.PreconditionFailed, 'note, last row');
    ws.setComment('C1048576', null);
    ws.updateCellFormat('XFD5', { fill: { foreground: '#FFFF00' } });
    rejectsWith(() => ws.insertCells('B5', CellShift.Right), StatusCode.PreconditionFailed, 'fill, last column');
    ws.updateCellFormat('XFD7', { font: { bold: true } });
    ws.insertCells('B7', CellShift.Right);
    ws.addMerge('B1048575:C1048576');
    rejectsWith(() => ws.insertCells('B3', CellShift.Down), StatusCode.PreconditionFailed, 'merge, last row');
  });
});

test('wrong direction, zero shift on delete, whole-column Down and unknown ordinals are InvalidArgument', async () => {
  await withWorkbook((wb, f) => {
    const { StatusCode, CellShift } = f;
    const ws = wb.sheet(0);
    rejectsWith(() => ws.insertCells('B3:C4', CellShift.Up), StatusCode.InvalidArgument, 'insert Up');
    rejectsWith(() => ws.insertCells('B3:C4', CellShift.Left), StatusCode.InvalidArgument, 'insert Left');
    rejectsWith(() => ws.deleteCells('B3:C4', CellShift.Down), StatusCode.InvalidArgument, 'delete Down');
    rejectsWith(() => ws.deleteCells('B3:C4', CellShift.Right), StatusCode.InvalidArgument, 'delete Right');
    rejectsWith(() => ws.insertCells('B:C', CellShift.Down), StatusCode.InvalidArgument, 'whole column Down');
    rejectsWith(() => ws.deleteCells('B:C', CellShift.Up), StatusCode.InvalidArgument, 'whole column Up');
    rejectsWith(() => ws.insertCells('3:4', CellShift.Right), StatusCode.InvalidArgument, 'whole row Right');
    rejectsWith(() => ws.insertCells('B3:C4', 9), StatusCode.InvalidArgument, 'unknown ordinal');
    rejectsWith(
      () => ws.insertCells(rect(1048576, 0, 1048576, 0), CellShift.Down),
      StatusCode.InvalidArgument,
      'past the grid',
    );
  });
});

test('a whole-column range with Right or Left acts as insertColumns and deleteColumns', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('A1', [[1, 2, 3]]);
    ws.insertCells('B:C', f.CellShift.Right);
    assert.equal(ws.getValue('B1'), null);
    assert.equal(ws.getValue('D1'), 2);
    ws.deleteCells('B:C', f.CellShift.Left);
    assert.equal(ws.getValue('B1'), 2);
  });
});

test('a whole-row range with Down acts as insertRows', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('A1', [[1], [2], [3]]);
    ws.insertCells('2:3', f.CellShift.Down);
    assert.equal(ws.getValue('A2'), null);
    assert.equal(ws.getValue('A4'), 2);
  });
});

test('totals row on shifts an occupied row below the table', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setValues('A1', [
      ['Item', 'Qty'],
      ['a', 1],
      ['b', 2],
    ]);
    const table = ws.addTable({ name: 'Sales', range: 'A1:B3' });
    ws.setValue('B4', 5);
    table.update({ totalsRow: true });
    assert.equal(table.getProperties().totalsRow, true);
    assert.equal(ws.getValue('B5'), 5);
  });
});

test('totals row off is refused when a formula references the totals row', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('A1', [
      ['Item', 'Qty'],
      ['a', 1],
      ['b', 2],
    ]);
    const table = ws.addTable({ name: 'Sales', range: 'A1:B3', totalsRow: true });
    ws.setFormula('D1', '=B3');
    rejectsWith(() => table.update({ totalsRow: false }), f.StatusCode.PreconditionFailed, 'referenced totals row');
    assert.equal(table.getProperties().totalsRow, true);
  });
});

test('argument types: fractional coordinate is RangeError, boolean or missing shift and missing range are TypeError', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.throws(() => ws.insertCells(rect(1.5, 0, 2, 0), f.CellShift.Down), RangeError);
    assert.throws(() => ws.insertCells(rect(Number.NaN, 0, 2, 0), f.CellShift.Down), RangeError);
    assert.throws(() => ws.insertCells(rect(0, 0, 2, 0), true), TypeError);
    assert.throws(() => ws.insertCells('B3:C4'), TypeError);
    assert.throws(() => ws.deleteCells('B3:C4'), TypeError);
    assert.throws(() => ws.insertCells(null, f.CellShift.Down), TypeError);
    assert.throws(() => ws.insertCells(undefined, f.CellShift.Down), TypeError);
  });
});
