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

function readSetup(wb) {
  const result = wb.getPageSetup(0);
  return result;
}

function readMargins(wb) {
  const result = wb.getPageMargins(0);
  return result;
}

test('page setup narrow fields reject lossy values before native mutation', () => {
  const wb = addon.Workbook.createDefault();
  try {
    wb.updatePageSetup(0, { scalePct: 125 });
    for (const value of [2 ** 32 + 100, 100.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      assert.throws(() => wb.updatePageSetup(0, { scalePct: value }), RangeError, `scalePct=${String(value)}`);
      assert.equal(readSetup(wb).scalePct, 125, `invalid scale changed the stored value: ${String(value)}`);
    }
    assert.throws(
      () => wb.updatePageSetup(0, { scalePct: -1 }),
      (e) => e.name === 'FormulonError' && e.code === 2,
    );
    assert.equal(readSetup(wb).scalePct, 125, 'a negative scale changed the stored value');
    assert.throws(() => wb.updatePageSetup(0, { scalePct: Symbol('scalePct') }), TypeError);
    assert.equal(readSetup(wb).scalePct, 125);
  } finally {
    wb.dispose();
  }
});

test('all page setup integer fields share the checked reader', () => {
  const wb = addon.Workbook.createDefault();
  try {
    wb.updatePageSetup(0, { scalePct: 125 });
    const before = readSetup(wb);
    for (const key of ['orientation', 'paperSize', 'fitToWidth', 'fitToHeight']) {
      assert.throws(() => wb.updatePageSetup(0, { [key]: 2 ** 32 + 1 }), RangeError, key);
      assert.equal(readSetup(wb)[key], before[key], `${key} failure changed the stored value`);
    }
  } finally {
    wb.dispose();
  }
});

test('reader preserves fractional doubles and rejects pending getters before C mutation', () => {
  const wb = addon.Workbook.createDefault();
  try {
    wb.updatePageMargins(0, { leftIn: 0.25, topIn: 1.5 });
    const margins = readMargins(wb);
    assert.equal(margins.leftIn, 0.25);
    assert.equal(margins.topIn, 1.5);

    const setup = {};
    Object.defineProperty(setup, 'scalePct', {
      enumerable: true,
      get() {
        throw new Error('scale getter');
      },
    });
    assert.throws(() => wb.updatePageSetup(0, setup), /scale getter/);
    assert.equal(readSetup(wb).scalePct, 100);

    const marginsWithGetter = {};
    Object.defineProperty(marginsWithGetter, 'leftIn', {
      enumerable: true,
      get() {
        throw new Error('left getter');
      },
    });
    assert.throws(() => wb.updatePageMargins(0, marginsWithGetter), /left getter/);
    const unchanged = readMargins(wb);
    assert.equal(unchanged.leftIn, 0.25);
    assert.equal(unchanged.topIn, 1.5);
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
    assert.throws(() => wb.updatePrintOptions(0, options), /gridLines getter/);

    const headerFooter = {};
    Object.defineProperty(headerFooter, 'oddHeader', {
      enumerable: true,
      get() {
        throw new Error('oddHeader getter');
      },
    });
    assert.throws(() => wb.updateHeaderFooter(0, headerFooter), /oddHeader getter/);
    assert.equal(wb.getHeaderFooterXml(0), null);
  } finally {
    wb.dispose();
  }
});

test('nested print readers snapshot each stateful getter once', () => {
  const wb = addon.Workbook.createDefault();
  try {
    wb.updatePageSetup(0, { paperSize: 42 });
    let paperSizeReads = 0;
    const firstPaperSize = {};
    Object.defineProperty(firstPaperSize, 'paperSize', {
      get() {
        paperSizeReads += 1;
        return paperSizeReads === 1 ? 9 : undefined;
      },
    });
    wb.updatePageSetup(0, firstPaperSize);
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
    wb.updatePageSetup(0, omittedPaperSize);
    assert.equal(omittedPaperSizeReads, 1);
    assert.equal(readSetup(wb).paperSize, 9);
    assert.equal(readSetup(wb).paperSizeStated, true);

    let leftReads = 0;
    const firstMargin = {};
    Object.defineProperty(firstMargin, 'leftIn', {
      get() {
        leftReads += 1;
        return leftReads === 1 ? 0.75 : undefined;
      },
    });
    wb.updatePageMargins(0, firstMargin);
    assert.equal(leftReads, 1);
    assert.equal(readMargins(wb).leftIn, 0.75);

    let gridReads = 0;
    const firstGrid = {};
    Object.defineProperty(firstGrid, 'gridLines', {
      get() {
        gridReads += 1;
        return gridReads === 1;
      },
    });
    wb.updatePrintOptions(0, firstGrid);
    assert.equal(gridReads, 1);
    assert.match(wb.getPrintOptionsXml(0), /gridLines="true"/);

    let omittedGridReads = 0;
    const omittedGrid = {};
    Object.defineProperty(omittedGrid, 'gridLines', {
      get() {
        omittedGridReads += 1;
        return omittedGridReads === 1 ? undefined : false;
      },
    });
    wb.updatePrintOptions(0, omittedGrid);
    assert.equal(omittedGridReads, 1);
    assert.match(wb.getPrintOptionsXml(0), /gridLines="true"/);

    let headerFlagReads = 0;
    const headerFlag = {};
    Object.defineProperty(headerFlag, 'differentFirst', {
      get() {
        headerFlagReads += 1;
        return headerFlagReads === 1;
      },
    });
    wb.updateHeaderFooter(0, headerFlag);
    assert.equal(headerFlagReads, 1);
    assert.match(wb.getHeaderFooterXml(0), /differentFirst="true"/);
  } finally {
    wb.dispose();
  }
});
