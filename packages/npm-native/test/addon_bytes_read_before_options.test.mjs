// A byte argument is copied when it is read, before any option getter the caller supplied can run: a getter that
// overwrites the buffer must not change what Workbook.load, addImage or restoreDrawing consume.

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

test('Workbook.load reads the original bytes when a profileId getter overwrites them', () => {
  const bytes = withWorkbook((wb) => wb.save());
  const original = Uint8Array.from(bytes);
  const options = {
    get profileId() {
      bytes.fill(0);
      return undefined;
    },
  };
  const loaded = addon.Workbook.load(bytes, options);
  try {
    assert.deepEqual(loaded.save().length > 0, true);
  } finally {
    loaded.dispose();
  }
  assert.notDeepEqual(bytes, original, 'the getter ran');
});

test('addImage stores the original picture when a name getter overwrites the bytes', () => {
  withWorkbook((wb) => {
    const png = Uint8Array.from(PNG_1X1);
    const id = wb.addImage(0, png, {
      get name() {
        png.fill(0);
        return 'Logo';
      },
    });
    assert.deepEqual(png, new Uint8Array(png.length), 'the getter ran');
    assert.deepEqual(wb.getPicture(0, id).bytes, PNG_1X1);
  });
});

test('restoreDrawing restores the original snapshot when an assignNewId getter overwrites the bytes', () => {
  withWorkbook((wb) => {
    const id = wb.addImage(0, Uint8Array.from(PNG_1X1), {});
    const snapshot = wb.snapshotDrawing(0, id);
    const restored = wb.restoreDrawing(0, snapshot, {
      get assignNewId() {
        snapshot.fill(0);
        return true;
      },
    });
    assert.notEqual(restored, id);
    assert.deepEqual(wb.getPicture(0, restored).bytes, PNG_1X1);
  });
});
