import test from 'node:test';
import { pathToFileURL } from 'node:url';
import { assert, getModule } from './smoke_support.mjs';

const RANGE = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };

function dataBarRule(overrides = {}) {
  return {
    ranges: [RANGE],
    type: 3,
    dataBar: {
      min: { type: 3 },
      max: { type: 4 },
      ...overrides,
    },
  };
}

async function getValidationModule() {
  if (process.env.FORMULON_NODE_ADDON) {
    const addon = await import(pathToFileURL(process.env.FORMULON_NODE_ADDON).href);
    return addon.default ?? addon;
  }
  return getModule();
}

test('nested style integers reject narrowing before addXf mutates the workbook', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const before = wb.getCellXfCount().value;
    assert.throws(() => wb.addCellXf({ horizontalAlign: 256 }), RangeError);
    assert.equal(wb.getCellXfCount().value, before);

    assert.throws(() => wb.addCellXf({ numFmtId: 65536 }), RangeError);
    assert.equal(wb.getCellXfCount().value, before);

    assert.throws(() => wb.addCellXf({ fontIndex: 2 ** 32 }), RangeError);
    assert.equal(wb.getCellXfCount().value, before);

    const beforeFonts = wb.getFontCount().value;
    assert.throws(() => wb.addFont({ underline: 256 }), RangeError);
    assert.equal(wb.getFontCount().value, beforeFonts);

    const beforeFills = wb.getFillCount().value;
    assert.throws(() => wb.addFill({ pattern: 256 }), RangeError);
    assert.equal(wb.getFillCount().value, beforeFills);

    const beforeBorders = wb.getBorderCount().value;
    assert.throws(() => wb.addBorder({ left: { style: 256 } }), RangeError);
    assert.equal(wb.getBorderCount().value, beforeBorders);

    for (const value of [Number.NaN, Number.POSITIVE_INFINITY, 1.5, -1]) {
      assert.throws(() => wb.addCellXf({ horizontalAlign: value }), RangeError, `horizontalAlign=${String(value)}`);
      assert.equal(wb.getCellXfCount().value, before);
    }

    const omitted = wb.addCellXf({ horizontalAlign: null, verticalAlign: undefined });
    assert.ok(omitted.status.ok, JSON.stringify(omitted.status));
    assert.equal(wb.getCellXfCount().value, before);
  } finally {
    wb.dispose();
  }
});

test('nested conditional-format integers reject narrowing before addConditionalFormat mutates', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const beforeDxfs = wb.getDxfCount().value;
    assert.throws(() => wb.addDxf({ numFmt: { numFmtId: 65536 } }), RangeError);
    assert.equal(wb.getDxfCount().value, beforeDxfs);

    assert.equal(wb.getConditionalFormats(0).length, 0);
    assert.throws(() => wb.addCfRule(0, dataBarRule({ minLengthPct: 256 })), RangeError);
    assert.equal(wb.getConditionalFormats(0).length, 0);

    for (const value of [Number.NaN, Number.POSITIVE_INFINITY, 1.5, -1]) {
      assert.throws(
        () => wb.addCfRule(0, dataBarRule({ minLengthPct: value })),
        RangeError,
        `minLengthPct=${String(value)}`,
      );
      assert.equal(wb.getConditionalFormats(0).length, 0);
    }

    const omitted = wb.addCfRule(0, dataBarRule({ minLengthPct: null, maxLengthPct: undefined }));
    assert.ok(omitted.status.ok, JSON.stringify(omitted));
    const stored = wb.getConditionalFormats(0);
    assert.equal(stored.length, 1);
    assert.equal(stored[0].dataBar.minLengthPct, 10);
    assert.equal(stored[0].dataBar.maxLengthPct, 90);
  } finally {
    wb.dispose();
  }
});

test('nested style and conditional-format fields require primitive numeric values', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const beforeXfs = wb.getCellXfCount().value;
    assert.throws(() => wb.addCellXf({ horizontalAlign: '1' }), TypeError);
    assert.equal(wb.getCellXfCount().value, beforeXfs);

    assert.throws(() => wb.addCellXf({ horizontalAlign: Symbol('numeric') }), TypeError);
    assert.equal(wb.getCellXfCount().value, beforeXfs);

    assert.throws(() => wb.addCfRule(0, dataBarRule({ minLengthPct: '10' })), TypeError);
    assert.equal(wb.getConditionalFormats(0).length, 0);

    assert.throws(() => wb.addCfRule(0, dataBarRule({ minLengthPct: Symbol('numeric') })), TypeError);
    assert.equal(wb.getConditionalFormats(0).length, 0);
  } finally {
    wb.dispose();
  }
});

test('nested readers preserve the original getter error and abort the whole mutation', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const beforeXfs = wb.getCellXfCount().value;
    const style = {};
    Object.defineProperty(style, 'horizontalAlign', {
      enumerable: true,
      get() {
        throw new Error('style getter');
      },
    });
    assert.throws(() => wb.addCellXf(style), /style getter/);
    assert.equal(wb.getCellXfCount().value, beforeXfs);

    const rule = dataBarRule();
    Object.defineProperty(rule.dataBar, 'minLengthPct', {
      enumerable: true,
      get() {
        throw new Error('conditional-format getter');
      },
    });
    assert.throws(() => wb.addCfRule(0, rule), /conditional-format getter/);
    assert.equal(wb.getConditionalFormats(0).length, 0);
  } finally {
    wb.dispose();
  }
});

test('style and conditional-format readers snapshot stateful fields once', async () => {
  const mod = await getValidationModule();
  const wb = mod.Workbook.createDefault();
  try {
    let horizontalReads = 0;
    const statefulXf = {};
    Object.defineProperty(statefulXf, 'horizontalAlign', {
      get() {
        horizontalReads += 1;
        return horizontalReads === 1 ? 1 : undefined;
      },
    });
    const xf = wb.addCellXf(statefulXf);
    assert.ok(xf.status.ok, JSON.stringify(xf.status));
    assert.equal(horizontalReads, 1);
    const storedXf = wb.getCellXf(xf.index);
    assert.equal(storedXf.horizontalAlign, 1);
    assert.equal(storedXf.hasHorizontalAlign, true);

    const firstDxf = wb.addDxf({ font: { name: 'first' } });
    const secondDxf = wb.addDxf({ font: { name: 'second' } });
    assert.ok(firstDxf.status.ok, JSON.stringify(firstDxf.status));
    assert.ok(secondDxf.status.ok, JSON.stringify(secondDxf.status));
    let dxfReads = 0;
    const rule = dataBarRule();
    Object.defineProperty(rule, 'dxfId', {
      get() {
        dxfReads += 1;
        return dxfReads === 1 ? secondDxf.index : undefined;
      },
    });
    const cf = wb.addCfRule(0, rule);
    assert.ok(cf.status.ok, JSON.stringify(cf.status));
    assert.equal(dxfReads, 1);
    assert.equal(wb.getConditionalFormats(0)[0].dxfId, secondDxf.index);

    const beforeMemory = wb.getMemoryUsage();
    const lateFont = { name: 'late font '.repeat(1024), bold: true };
    Object.defineProperty(lateFont, 'color', {
      get() {
        throw new Error('late font getter');
      },
    });
    const beforeFonts = wb.getFontCount().value;
    for (let i = 0; i < 8; i += 1) {
      assert.throws(() => wb.addFont(lateFont), /late font getter/);
    }
    assert.equal(wb.getFontCount().value, beforeFonts);
    assert.equal(wb.getMemoryUsage(), beforeMemory);
  } finally {
    wb.dispose();
  }
});

test('conditional-format arrays skip nullish entries and keep later records', async () => {
  const mod = await getValidationModule();
  const wb = mod.Workbook.createDefault();
  try {
    const added = wb.addCfRule(0, {
      // Include null, undefined, and a hole before the valid range. The old
      // reader skipped those entries and still passed the later range to C.
      ranges: Object.assign(new Array(4), { 0: null, 1: undefined, 3: RANGE }),
      type: 3,
      dataBar: { min: { type: 3 }, max: { type: 4 } },
    });
    assert.ok(added.status.ok, JSON.stringify(added.status));
    const stored = wb.getConditionalFormats(0);
    assert.equal(stored.length, 1);
    assert.equal(stored[0].ranges.length, 1);
    assert.deepEqual(stored[0].ranges[0], RANGE);
  } finally {
    wb.dispose();
  }
});

test('named-style and theme integer fields reject before their C mutations', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const styleCount = wb.listCellStyles().length;
    for (const value of [2 ** 32, Number.NaN, 1.5, -1]) {
      assert.throws(() => wb.setCellStyle({ name: 'invalid', xfId: value }), RangeError);
      assert.equal(wb.listCellStyles().length, styleCount);
    }

    const before = wb.getThemeFonts().colors;
    for (const value of [Symbol('theme'), Number.NaN, 1.5, -1, 2 ** 32]) {
      const values = Array.from({ length: 12 }, (_, i) => (i === 0 ? value : before[i]));
      assert.throws(
        () => wb.setThemeColors(values),
        typeof value === 'symbol' ? TypeError : /must be a number|outside uint32/,
      );
      assert.deepEqual(wb.getThemeFonts().colors, before);
    }
  } finally {
    wb.dispose();
  }
});

test('style and CF reader defaults survive nullish fields and valid u8 boundaries', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const font = wb.addFont({ name: 'Reader defaults', size: null, colorArgb: undefined, underline: null });
    assert.ok(font.status.ok, JSON.stringify(font.status));
    const storedFont = wb.getFont(font.index);
    assert.equal(storedFont.size, 11);
    assert.equal(storedFont.colorArgb >>> 0, 0xff000000);

    const cf = wb.addCfRule(
      0,
      dataBarRule({
        fill: { r: 255, g: 0, b: 1 },
        minLengthPct: 0,
        maxLengthPct: 100,
      }),
    );
    assert.ok(cf.status.ok, JSON.stringify(cf));
    const stored = wb.getConditionalFormats(0)[0].dataBar;
    assert.equal(stored.fill.r, 255);
    assert.equal(stored.fill.a, 255);
    assert.equal(stored.minLengthPct, 0);
    assert.equal(stored.maxLengthPct, 100);
  } finally {
    wb.dispose();
  }
});
