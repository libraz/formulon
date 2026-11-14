// Sheet.listCircularReferences: the unresolved cycle members on a sheet, row-major.

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

test('listCircularReferences names the cycle members and empties once the cycle is broken', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    assert.deepEqual(ws.listCircularReferences(), []);
    ws.setFormula('A1', '=B1');
    ws.setFormula('B1', '=A1');
    wb.recalc();
    assert.deepEqual(
      ws.listCircularReferences().map((c) => ({ row: c.row, col: c.col })),
      [
        { row: 0, col: 0 },
        { row: 0, col: 1 },
      ],
    );
    ws.setValue('B1', 1);
    wb.recalc();
    assert.deepEqual(ws.listCircularReferences(), []);
  });
});

test('listCircularReferences orders members row-major and recalculates first under automatic calculation', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setFormula('C2', '=A1');
    ws.setFormula('A1', '=B3');
    ws.setFormula('B3', '=C2');
    ws.setFormula('D1', '=A1+1');
    assert.deepEqual(
      ws.listCircularReferences().map((c) => ({ row: c.row, col: c.col })),
      [
        { row: 0, col: 0 },
        { row: 1, col: 2 },
        { row: 2, col: 1 },
      ],
    );
  });
});

test('listCircularReferences lists nothing under iterative calculation', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setFormula('A1', '=B1+1');
    ws.setFormula('B1', '=A1/2');
    wb.recalc();
    assert.equal(ws.listCircularReferences().length, 2);
    wb.updateIterative({ enabled: true });
    wb.recalc();
    assert.deepEqual(ws.listCircularReferences(), []);
  });
});
