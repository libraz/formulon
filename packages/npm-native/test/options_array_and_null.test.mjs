// The options argument of restoreDrawing, addImage, parseReference, localizeFormula and canonicalizeFormula: an
// array is a TypeError, and null or undefined is the same as omitting it.

import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { FormulonError } from '../dist/common.mjs';

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
addon.__setErrorClass(FormulonError);

const PNG_1X1 = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='),
  (c) => c.charCodeAt(0),
);

function withWorkbook(fn) {
  const wb = addon.Workbook.createDefault();
  try {
    return fn(wb);
  } finally {
    wb.dispose();
  }
}

const NOT_OBJECT = /options must be an object/;

function refused(fn) {
  assert.throws(fn, (e) => {
    assert.ok(e instanceof TypeError, `expected a TypeError, got ${e}`);
    assert.match(e.message, NOT_OBJECT);
    return true;
  });
}

test('an array options argument is a TypeError on every options-taking call', () => {
  withWorkbook((wb) => {
    refused(() => wb.addImage(0, Uint8Array.from(PNG_1X1), []));
    refused(() => wb.restoreDrawing(0, Uint8Array.from(PNG_1X1), []));
    refused(() => wb.parseReference('A1', []));
    refused(() => addon.localizeFormula('=SUM(A1)', 'win-365-en_US', []));
    refused(() => addon.canonicalizeFormula('=SUM(A1)', 'win-365-en_US', []));
  });
});

for (const [label, absent] of [
  ['null', null],
  ['undefined', undefined],
]) {
  test(`addImage and restoreDrawing take ${label} options as none`, () => {
    withWorkbook((wb) => {
      const id = wb.addImage(0, Uint8Array.from(PNG_1X1), absent);
      assert.deepEqual(wb.getPicture(0, id).bytes, PNG_1X1);
      const restored = wb.restoreDrawing(0, wb.snapshotDrawing(0, id), absent);
      assert.equal(restored, id);
    });
  });
}
