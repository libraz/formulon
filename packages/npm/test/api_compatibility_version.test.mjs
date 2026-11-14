// Workbook.getCompatibilityVersion / setCompatibilityVersion: the stored version, its effect on
// text-position functions, and the refusals.

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

function expectCode(fn, code) {
  assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, code, `code ${e.code} (${e.message})`);
    return true;
  });
}

test('CompatibilityVersion holds the ordinals stored in the file', async () => {
  const f = await getModule();
  assert.deepEqual({ ...f.CompatibilityVersion }, { Version1: 1, Version2: 2, Version3: 3 });
  assert.equal(Object.isFrozen(f.CompatibilityVersion), true);
});

test('a new workbook reads version 2', async () => {
  await withWorkbook((wb, f) => {
    assert.equal(wb.getCompatibilityVersion(), f.CompatibilityVersion.Version2);
  });
});

test('LEN counts UTF-16 units under version 1 and code points under version 2', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setFormula('A1', '=LEN("a\u{1F600}")');
    wb.setCompatibilityVersion(f.CompatibilityVersion.Version1);
    assert.equal(wb.getCompatibilityVersion(), 1);
    assert.equal(ws.getValue('A1'), 3);
    wb.setCompatibilityVersion(f.CompatibilityVersion.Version2);
    assert.equal(wb.getCompatibilityVersion(), 2);
    assert.equal(ws.getValue('A1'), 2);
  });
});

test('version 3 is stored, evaluates as version 2 and marks formulas only across version 1', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    wb.setCalcMode(f.CalcMode.Manual);
    ws.setFormula('A1', '=LEN("a\u{1F600}")');
    wb.recalc();
    wb.setCompatibilityVersion(f.CompatibilityVersion.Version3);
    assert.equal(wb.getCompatibilityVersion(), 3);
    assert.equal(ws.isCellStale('A1'), false);
    wb.setCompatibilityVersion(f.CompatibilityVersion.Version1);
    assert.equal(ws.isCellStale('A1'), true);
    wb.recalc();
    assert.equal(ws.getValue('A1'), 3);
    wb.setCompatibilityVersion(f.CompatibilityVersion.Version3);
    assert.equal(wb.getCompatibilityVersion(), 3);
    assert.equal(ws.isCellStale('A1'), true);
    wb.recalc();
    assert.equal(ws.getValue('A1'), 2);
  });
});

test('a value outside 1..3 is InvalidArgument', async () => {
  await withWorkbook((wb, f) => {
    expectCode(() => wb.setCompatibilityVersion(0), f.StatusCode.InvalidArgument);
    expectCode(() => wb.setCompatibilityVersion(4), f.StatusCode.InvalidArgument);
    expectCode(() => wb.setCompatibilityVersion(-1), f.StatusCode.InvalidArgument);
    assert.equal(wb.getCompatibilityVersion(), 2);
  });
});

test('a non-number argument is a TypeError and a non-integer number a RangeError', async () => {
  await withWorkbook((wb) => {
    assert.throws(() => wb.setCompatibilityVersion('2'), TypeError);
    assert.throws(() => wb.setCompatibilityVersion(null), TypeError);
    assert.throws(() => wb.setCompatibilityVersion(undefined), TypeError);
    assert.throws(() => wb.setCompatibilityVersion(Number.NaN), RangeError);
    assert.throws(() => wb.setCompatibilityVersion(1.5), RangeError);
    assert.equal(wb.getCompatibilityVersion(), 2);
  });
});
