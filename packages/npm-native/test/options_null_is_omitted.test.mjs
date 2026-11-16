// An options argument given as `null` is the same as omitting it; no call refuses it.

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

test('convertCivilToSerial and convertSerialToCivil take null options as omitted', async () => {
  const f = await getModule();
  const civil = { year: 2020, month: 1, day: 15, hour: 0, minute: 0, second: 0 };
  assert.equal(f.convertCivilToSerial(civil, null), f.convertCivilToSerial(civil));
  assert.deepEqual(f.convertSerialToCivil(43845, null), f.convertSerialToCivil(43845));
});

test('Sheet.evaluateCfRange takes null options as omitted', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    assert.deepEqual(ws.evaluateCfRange('A1', null), ws.evaluateCfRange('A1'));
  });
});

test('Workbook.addSheet and Workbook.copySheet take null options as omitted', async () => {
  await withWorkbook((wb) => {
    const added = wb.addSheet('Added', null);
    assert.equal(added.index, 1);
    const copy = wb.copySheet(0, null);
    assert.equal(copy.index, 2);
  });
});

test('Workbook.formatReference takes null options as omitted', async () => {
  await withWorkbook((wb) => {
    assert.equal(wb.formatReference('A1:B2', null), wb.formatReference('A1:B2'));
  });
});
