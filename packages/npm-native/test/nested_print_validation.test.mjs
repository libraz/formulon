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

function readSetup(wb) {
  const result = wb.getSheetPageSetup(0);
  assert.equal(result.status.ok, true, JSON.stringify(result.status));
  return result;
}

function readMargins(wb) {
  const result = wb.getSheetPageMargins(0);
  assert.equal(result.status.ok, true, JSON.stringify(result.status));
  return result;
}

test('page setup narrow fields reject lossy values before native mutation', () => {
  const wb = addon.Workbook.createDefault();
  try {
    assert.equal(wb.setSheetPageSetup(0, { scale: 125 }).ok, true);
    for (const value of [2 ** 32 + 100, 100.5, Number.NaN, Number.POSITIVE_INFINITY, -1]) {
      assert.throws(() => wb.setSheetPageSetup(0, { scale: value }), RangeError, `scale=${String(value)}`);
      assert.equal(readSetup(wb).scale, 125, `invalid scale changed the stored value: ${String(value)}`);
    }
    assert.throws(() => wb.setSheetPageSetup(0, { scale: Symbol('scale') }), TypeError);
    assert.equal(readSetup(wb).scale, 125);
  } finally {
    wb.dispose();
  }
});

test('all page setup integer fields share the checked reader', () => {
  const wb = addon.Workbook.createDefault();
  try {
    assert.equal(wb.setSheetPageSetup(0, { scale: 125 }).ok, true);
    const before = readSetup(wb);
    for (const key of ['orientation', 'paperSize', 'fitToWidth', 'fitToHeight']) {
      assert.throws(() => wb.setSheetPageSetup(0, { [key]: 2 ** 32 + 1 }), RangeError, key);
      assert.equal(readSetup(wb)[key], before[key], `${key} failure changed the stored value`);
    }
  } finally {
    wb.dispose();
  }
});

test('reader preserves fractional doubles and rejects pending getters before C mutation', () => {
  const wb = addon.Workbook.createDefault();
  try {
    assert.ok(wb.setSheetPageMargins(0, { left: 0.25, top: 1.5 }).ok);
    const margins = readMargins(wb);
    assert.equal(margins.left, 0.25);
    assert.equal(margins.top, 1.5);

    const setup = {};
    Object.defineProperty(setup, 'scale', {
      enumerable: true,
      get() {
        throw new Error('scale getter');
      },
    });
    assert.throws(() => wb.setSheetPageSetup(0, setup), /scale getter/);
    assert.equal(readSetup(wb).scale, 100);

    const marginsWithGetter = {};
    Object.defineProperty(marginsWithGetter, 'left', {
      enumerable: true,
      get() {
        throw new Error('left getter');
      },
    });
    assert.throws(() => wb.setSheetPageMargins(0, marginsWithGetter), /left getter/);
    const unchanged = readMargins(wb);
    assert.equal(unchanged.left, 0.25);
    assert.equal(unchanged.top, 1.5);
  } finally {
    wb.dispose();
  }
});

test('print option and header/footer readers stop before native mutation on getter failure', () => {
  const wb = addon.Workbook.createDefault();
  try {
    const options = {};
    Object.defineProperty(options, 'gridLines', {
      enumerable: true,
      get() {
        throw new Error('gridLines getter');
      },
    });
    assert.throws(() => wb.setSheetPrintOptions(0, options), /gridLines getter/);

    const headerFooter = {};
    Object.defineProperty(headerFooter, 'oddHeader', {
      enumerable: true,
      get() {
        throw new Error('oddHeader getter');
      },
    });
    assert.throws(() => wb.setSheetHeaderFooter(0, headerFooter), /oddHeader getter/);
    assert.equal(wb.getSheetHeaderFooterXml(0).xml, '');
  } finally {
    wb.dispose();
  }
});

test('nested print readers snapshot each stateful getter once', () => {
  const wb = addon.Workbook.createDefault();
  try {
    assert.equal(wb.setSheetPageSetup(0, { paperSize: 42 }).ok, true);
    let paperSizeReads = 0;
    const firstPaperSize = {};
    Object.defineProperty(firstPaperSize, 'paperSize', {
      get() {
        paperSizeReads += 1;
        return paperSizeReads === 1 ? 9 : undefined;
      },
    });
    assert.equal(wb.setSheetPageSetup(0, firstPaperSize).ok, true);
    assert.equal(paperSizeReads, 1);
    assert.equal(readSetup(wb).paperSize, 9);
    assert.equal(readSetup(wb).paperSizeStated, true);

    let omittedPaperSizeReads = 0;
    const omittedPaperSize = {};
    Object.defineProperty(omittedPaperSize, 'paperSize', {
      get() {
        omittedPaperSizeReads += 1;
        return omittedPaperSizeReads === 1 ? undefined : 9;
      },
    });
    assert.equal(wb.setSheetPageSetup(0, omittedPaperSize).ok, true);
    assert.equal(omittedPaperSizeReads, 1);
    assert.equal(readSetup(wb).paperSize, 9);
    assert.equal(readSetup(wb).paperSizeStated, true);

    let leftReads = 0;
    const firstMargin = {};
    Object.defineProperty(firstMargin, 'left', {
      get() {
        leftReads += 1;
        return leftReads === 1 ? 0.75 : undefined;
      },
    });
    assert.equal(wb.setSheetPageMargins(0, firstMargin).ok, true);
    assert.equal(leftReads, 1);
    assert.equal(readMargins(wb).left, 0.75);

    let gridReads = 0;
    const firstGrid = {};
    Object.defineProperty(firstGrid, 'gridLines', {
      get() {
        gridReads += 1;
        return gridReads === 1;
      },
    });
    assert.equal(wb.setSheetPrintOptions(0, firstGrid).ok, true);
    assert.equal(gridReads, 1);
    assert.match(wb.getSheetPrintOptionsXml(0).xml, /gridLines="true"/);

    let omittedGridReads = 0;
    const omittedGrid = {};
    Object.defineProperty(omittedGrid, 'gridLines', {
      get() {
        omittedGridReads += 1;
        return omittedGridReads === 1 ? undefined : false;
      },
    });
    assert.equal(wb.setSheetPrintOptions(0, omittedGrid).ok, true);
    assert.equal(omittedGridReads, 1);
    assert.match(wb.getSheetPrintOptionsXml(0).xml, /gridLines="true"/);

    let headerFlagReads = 0;
    const headerFlag = {};
    Object.defineProperty(headerFlag, 'differentFirst', {
      get() {
        headerFlagReads += 1;
        return headerFlagReads === 1;
      },
    });
    assert.equal(wb.setSheetHeaderFooter(0, headerFlag).ok, true);
    assert.equal(headerFlagReads, 1);
    assert.match(wb.getSheetHeaderFooterXml(0).xml, /differentFirst="true"/);
  } finally {
    wb.dispose();
  }
});
