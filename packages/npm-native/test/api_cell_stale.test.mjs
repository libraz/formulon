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

test('a manual-mode save computes stale formulas first unless calcOnSave is off', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    wb.setCalcMode(f.CalcMode.Manual);
    ws.setValue('A1', 1);
    ws.setFormula('B1', '=A1*2');
    wb.recalc();
    ws.setValue('A1', 5);
    assert.equal(wb.getCalcOnSave(), true);
    wb.save();
    assert.equal(ws.isCellStale('B1'), false);
    assert.equal(ws.getValue('B1'), 10);

    ws.setValue('A1', 7);
    wb.setCalcOnSave(false);
    const bytes = wb.save();
    assert.equal(ws.isCellStale('B1'), true);
    assert.equal(ws.getValue('B1'), 10);
    const loaded = f.Workbook.load(bytes);
    try {
      assert.equal(loaded.sheet(0).isCellStale('B1'), true);
      assert.equal(loaded.sheet(0).getValue('B1'), 10);
    } finally {
      loaded.dispose();
    }
  });
});

const isFormulonError = (code) => (e) => e.name === 'FormulonError' && e.code === code;

// A manual-calculation book saved with B1 and C1 stale (A1 edited after the recalc) and D1 clean.
async function staleBytes() {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    const ws = wb.sheet(0);
    wb.setCalcMode(f.CalcMode.Manual);
    ws.setValue('A1', 1);
    ws.setFormula('B1', '=A1*2');
    ws.setFormula('C1', '=B1+1');
    ws.setFormula('D1', '=1+1');
    wb.recalc();
    ws.setValue('A1', 2);
    wb.setCalcOnSave(false);
    return { f, bytes: wb.save() };
  } finally {
    wb.dispose();
  }
}

function staleSet(wb) {
  const ws = wb.sheet(0);
  return ['A1', 'B1', 'C1', 'D1'].map((ref) => [ref, ws.isCellStale(ref), ws.getValue(ref)]);
}

test('a profile given at load keeps the stale set and the cached values of the default load', async () => {
  const { f, bytes } = await staleBytes();
  const plain = f.Workbook.load(bytes);
  const profiled = f.Workbook.load(bytes, { profileId: 'mac-365-ja_JP' });
  try {
    assert.equal(plain.getProfileId(), 'win-365-en_US');
    assert.equal(profiled.getProfileId(), 'mac-365-ja_JP');
    assert.deepEqual(staleSet(profiled), staleSet(plain));
    assert.equal(profiled.sheet(0).isCellStale('B1'), true);
    assert.equal(profiled.sheet(0).isCellStale('D1'), false);
  } finally {
    plain.dispose();
    profiled.dispose();
  }
});

test('absent, undefined and null load options, and a null profileId, keep the load profile', async () => {
  const { f, bytes } = await staleBytes();
  for (const options of [undefined, null, {}, { profileId: undefined }, { profileId: null }]) {
    const wb = f.Workbook.load(bytes, options);
    try {
      assert.equal(wb.getProfileId(), 'win-365-en_US');
    } finally {
      wb.dispose();
    }
  }
});

test('load options refuse an unknown key, a wrong type and an unknown profile id', async () => {
  const { f, bytes } = await staleBytes();
  assert.throws(() => f.Workbook.load(bytes, { profile: 'mac-365-ja_JP' }), TypeError);
  assert.throws(() => f.Workbook.load(bytes, 'mac-365-ja_JP'), TypeError);
  assert.throws(() => f.Workbook.load(bytes, 5), TypeError);
  assert.throws(() => f.Workbook.load(bytes, []), TypeError);
  assert.throws(() => f.Workbook.load(bytes, { profileId: 5 }), TypeError);
  assert.throws(() => f.Workbook.load(bytes, { profileId: true }), TypeError);
  const invalid = isFormulonError(f.StatusCode.InvalidArgument);
  assert.throws(() => f.Workbook.load(bytes, { profileId: 'not-a-profile' }), invalid);
  assert.throws(() => f.Workbook.load(bytes, { profileId: '' }), invalid);
});

test('setProfileId after a load still makes every formula stale', async () => {
  const { f, bytes } = await staleBytes();
  const wb = f.Workbook.load(bytes);
  try {
    assert.equal(wb.sheet(0).isCellStale('D1'), false);
    wb.setProfileId('mac-365-ja_JP');
    for (const ref of ['B1', 'C1', 'D1']) assert.equal(wb.sheet(0).isCellStale(ref), true, ref);
  } finally {
    wb.dispose();
  }
});
