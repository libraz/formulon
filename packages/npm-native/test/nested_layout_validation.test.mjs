import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const addonPath =
  process.env.FORMULON_NODE_ADDON ??
  [
    path.resolve(testDir, '..', 'dist', 'prebuilds', `${process.platform}-${process.arch}`, 'formulon.node'),
    path.resolve(testDir, '..', 'dist', 'formulon.node'),
    path.resolve(testDir, '..', 'formulon.node'),
  ].find((candidate) => existsSync(candidate));
assert.ok(addonPath, 'native addon not found; set FORMULON_NODE_ADDON or stage dist/formulon.node');
const addon = createRequire(import.meta.url)(addonPath);

const RANGE = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };

function protectionSnapshot(wb) {
  const result = wb.getSheetProtection(0);
  assert.equal(result.status.ok, true, JSON.stringify(result.status));
  return structuredClone(result.protection);
}

test('sheet protection rejects every narrow nested integer before mutation', () => {
  const wb = addon.Workbook.createDefault();
  try {
    const keys = [
      'enabled',
      'sheet',
      'objects',
      'scenarios',
      'formatCells',
      'formatColumns',
      'formatRows',
      'insertColumns',
      'insertRows',
      'insertHyperlinks',
      'deleteColumns',
      'deleteRows',
      'selectLockedCells',
      'selectUnlockedCells',
      'sort',
      'autoFilter',
      'pivotTables',
    ];
    for (const key of keys) {
      const before = protectionSnapshot(wb);
      assert.throws(() => wb.setSheetProtection(0, { [key]: 1.5 }), RangeError, key);
      assert.deepEqual(protectionSnapshot(wb), before, `${key} changed protection`);
    }
    const beforeSpin = protectionSnapshot(wb);
    assert.throws(() => wb.setSheetProtection(0, { spinCount: 2 ** 32 }), RangeError);
    assert.deepEqual(protectionSnapshot(wb), beforeSpin);
  } finally {
    wb.dispose();
  }
});

test('sheet protection stops at a throwing nested getter before native mutation', () => {
  const wb = addon.Workbook.createDefault();
  try {
    const before = protectionSnapshot(wb);
    const spec = { enabled: 1, algorithmName: 'SHA-512' };
    Object.defineProperty(spec, 'formatRows', {
      enumerable: true,
      get() {
        throw new Error('protection getter');
      },
    });
    assert.throws(() => wb.setSheetProtection(0, spec), /protection getter/);
    assert.deepEqual(protectionSnapshot(wb), before);
  } finally {
    wb.dispose();
  }
});

test('sheet format defaults stops at a throwing nested getter before native mutation', () => {
  const wb = addon.Workbook.createDefault();
  try {
    const before = wb.getSheetFormatDefaults(0);
    const spec = { defaultColWidth: 12.5 };
    Object.defineProperty(spec, 'defaultRowHeight', {
      enumerable: true,
      get() {
        throw new Error('format getter');
      },
    });
    assert.throws(() => wb.setSheetFormatDefaults(0, spec), /format getter/);
    assert.deepEqual(wb.getSheetFormatDefaults(0), before);
  } finally {
    wb.dispose();
  }
});

test('cell geometry rejects fractional nested range coordinates before the query', () => {
  const wb = addon.Workbook.createDefault();
  try {
    assert.throws(() => wb.getCellRectPt(0, { ...RANGE, firstRow: 1.5 }, 0), RangeError);
    assert.throws(() => wb.getCellRectPt(0, { ...RANGE, lastCol: Number.NaN }, 0), RangeError);
    const range = {};
    Object.defineProperty(range, 'firstRow', {
      enumerable: true,
      get() {
        throw new Error('geometry getter');
      },
    });
    assert.throws(() => wb.getCellRectPt(0, range, 0), /geometry getter/);
  } finally {
    wb.dispose();
  }
});

test('partial recalc rejects fractional nested viewport coordinates before recalc', () => {
  const wb = addon.Workbook.createDefault();
  try {
    assert.ok(wb.setNumber(0, 0, 0, 1).ok);
    assert.ok(wb.setFormula(0, 0, 1, '=A1+1').ok);
    assert.throws(() => wb.partialRecalc({ ...RANGE, sheet: 0, firstRow: 1.5 }), RangeError);
    const viewport = { sheet: 0, firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 0 };
    Object.defineProperty(viewport, 'lastCol', {
      enumerable: true,
      get() {
        throw new Error('viewport getter');
      },
    });
    assert.throws(() => wb.partialRecalc(viewport), /viewport getter/);
    const value = wb.getValue(0, 0, 1);
    assert.equal(value.status.ok, true, JSON.stringify(value.status));
    assert.equal(value.value.kind, 0);
    assert.equal(value.value.number, 0);
  } finally {
    wb.dispose();
  }
});
