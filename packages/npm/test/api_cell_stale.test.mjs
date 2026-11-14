// Sheet.isCellStale: whether a formula cell holds a value its inputs have outdated, without recalculating.

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

test('isCellStale flags a dependent formula after an edit under manual calculation and clears on recalc', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    wb.setCalcMode(f.CalcMode.Manual);
    ws.setValue('A1', 1);
    ws.setFormula('B1', '=A1');
    wb.recalc();
    assert.equal(ws.isCellStale('B1'), false);
    ws.setValue('A1', 2);
    assert.equal(ws.isCellStale('B1'), true);
    assert.equal(ws.isCellStale('A1'), false);
    assert.equal(ws.isCellStale(0, 1), true);
    assert.equal(ws.isCellStale(0, 0), false);
    wb.recalc();
    assert.equal(ws.isCellStale('B1'), false);
    assert.equal(ws.isCellStale(0, 1), false);
  });
});

test('isCellStale flags every volatile and its dependents after any edit, even one nothing reads', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    wb.setCalcMode(f.CalcMode.Manual);
    ws.setFormula('A1', '=RAND()');
    ws.setFormula('B1', '=A1+1');
    ws.setFormula('C1', '=1+1');
    wb.recalc();
    assert.equal(ws.isCellStale('A1'), false);
    ws.setValue('F9', 1);
    assert.equal(ws.isCellStale('A1'), true);
    assert.equal(ws.isCellStale('B1'), true);
    assert.equal(ws.isCellStale('C1'), false);
    wb.recalc();
    assert.equal(ws.isCellStale('A1'), false);
    assert.equal(ws.isCellStale('B1'), false);
  });
});

test('isCellStale is false under automatic calculation after an edit', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    wb.setCalcMode(f.CalcMode.Auto);
    ws.setValue('A1', 1);
    ws.setFormula('B1', '=A1');
    wb.recalc();
    ws.setValue('A1', 2);
    assert.equal(ws.isCellStale('B1'), false);
    assert.equal(ws.isCellStale(0, 1), false);
  });
});

test('isCellStale is false for an empty cell', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    wb.setCalcMode(f.CalcMode.Manual);
    assert.equal(ws.isCellStale('D4'), false);
    assert.equal(ws.isCellStale(3, 3), false);
  });
});
