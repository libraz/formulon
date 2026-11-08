import test from 'node:test';
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

test('nested style integers reject narrowing before addXf mutates the workbook', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const before = wb.ooxml.listCellXfs().length;
    assert.throws(() => wb.ooxml.addCellXf({ horizontalAlign: 256 }), RangeError);
    assert.equal(wb.ooxml.listCellXfs().length, before);

    assert.throws(() => wb.ooxml.addCellXf({ numFmtId: 65536 }), RangeError);
    assert.equal(wb.ooxml.listCellXfs().length, before);

    assert.throws(() => wb.ooxml.addCellXf({ fontIndex: 2 ** 32 }), RangeError);
    assert.equal(wb.ooxml.listCellXfs().length, before);

    const beforeFonts = wb.ooxml.listFonts().length;
    assert.throws(() => wb.ooxml.addFont({ underline: 256 }), RangeError);
    assert.equal(wb.ooxml.listFonts().length, beforeFonts);

    const beforeFills = wb.ooxml.listFills().length;
    assert.throws(() => wb.ooxml.addFill({ pattern: 256 }), RangeError);
    assert.equal(wb.ooxml.listFills().length, beforeFills);

    const beforeBorders = wb.ooxml.listBorders().length;
    assert.throws(() => wb.ooxml.addBorder({ left: { style: 256 } }), RangeError);
    assert.equal(wb.ooxml.listBorders().length, beforeBorders);

    for (const value of [Number.NaN, Number.POSITIVE_INFINITY, 1.5, -1]) {
      assert.throws(
        () => wb.ooxml.addCellXf({ horizontalAlign: value }),
        RangeError,
        `horizontalAlign=${String(value)}`,
      );
      assert.equal(wb.ooxml.listCellXfs().length, before);
    }

    wb.ooxml.addCellXf({ horizontalAlign: null, verticalAlign: undefined });
    assert.equal(wb.ooxml.listCellXfs().length, before);
  } finally {
    wb.dispose();
  }
});

test('nested conditional-format integers reject narrowing before addConditionalFormat mutates', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const beforeDxfs = wb.ooxml.listDxfs().length;
    assert.throws(() => wb.ooxml.addDxf({ numFmt: { numFmtId: 65536 } }), RangeError);
    assert.equal(wb.ooxml.listDxfs().length, beforeDxfs);

    assert.equal(wb.sheet(0).listCfRules().length, 0);
    assert.throws(() => wb.sheet(0).addCfRule(dataBarRule({ minLengthPct: 256 })), RangeError);
    assert.equal(wb.sheet(0).listCfRules().length, 0);

    for (const value of [Number.NaN, Number.POSITIVE_INFINITY, 1.5, -1]) {
      assert.throws(
        () => wb.sheet(0).addCfRule(dataBarRule({ minLengthPct: value })),
        RangeError,
        `minLengthPct=${String(value)}`,
      );
      assert.equal(wb.sheet(0).listCfRules().length, 0);
    }

    wb.sheet(0).addCfRule(dataBarRule({ minLengthPct: null, maxLengthPct: undefined }));
    const stored = wb.sheet(0).listCfRules();
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
    const beforeXfs = wb.ooxml.listCellXfs().length;
    assert.throws(() => wb.ooxml.addCellXf({ horizontalAlign: '1' }), TypeError);
    assert.equal(wb.ooxml.listCellXfs().length, beforeXfs);

    assert.throws(() => wb.ooxml.addCellXf({ horizontalAlign: Symbol('numeric') }), TypeError);
    assert.equal(wb.ooxml.listCellXfs().length, beforeXfs);

    assert.throws(() => wb.sheet(0).addCfRule(dataBarRule({ minLengthPct: '10' })), TypeError);
    assert.equal(wb.sheet(0).listCfRules().length, 0);

    assert.throws(() => wb.sheet(0).addCfRule(dataBarRule({ minLengthPct: Symbol('numeric') })), TypeError);
    assert.equal(wb.sheet(0).listCfRules().length, 0);
  } finally {
    wb.dispose();
  }
});

test('nested readers preserve the original getter error and abort the whole mutation', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const beforeXfs = wb.ooxml.listCellXfs().length;
    const style = {};
    Object.defineProperty(style, 'horizontalAlign', {
      enumerable: true,
      get() {
        throw new Error('style getter');
      },
    });
    assert.throws(() => wb.ooxml.addCellXf(style), /style getter/);
    assert.equal(wb.ooxml.listCellXfs().length, beforeXfs);

    const rule = dataBarRule();
    Object.defineProperty(rule.dataBar, 'minLengthPct', {
      enumerable: true,
      get() {
        throw new Error('conditional-format getter');
      },
    });
    assert.throws(() => wb.sheet(0).addCfRule(rule), /conditional-format getter/);
    assert.equal(wb.sheet(0).listCfRules().length, 0);
  } finally {
    wb.dispose();
  }
});

test('style and conditional-format readers snapshot stateful fields once', async () => {
  const mod = await getModule();
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
    const xf = wb.ooxml.addCellXf(statefulXf);
    assert.equal(horizontalReads, 1);
    const storedXf = wb.ooxml.getCellXf(xf);
    assert.equal(storedXf.horizontalAlign, 1);

    wb.ooxml.addDxf({ font: { name: 'first' } });
    const secondDxf = wb.ooxml.addDxf({ font: { name: 'second' } });
    let dxfReads = 0;
    const rule = { ranges: [RANGE], type: mod.CfRuleType.Expression, formula1: '=TRUE' };
    Object.defineProperty(rule, 'dxfId', {
      enumerable: true,
      get() {
        dxfReads += 1;
        return dxfReads === 1 ? secondDxf : undefined;
      },
    });
    wb.sheet(0).addCfRule(rule);
    assert.equal(dxfReads, 1);
    assert.equal(wb.sheet(0).listCfRules()[0].dxfId, secondDxf);

    const beforeMemory = wb.getMemoryUsage();
    const lateFont = { name: 'late font '.repeat(1024), bold: true };
    Object.defineProperty(lateFont, 'color', {
      get() {
        throw new Error('late font getter');
      },
    });
    const beforeFonts = wb.ooxml.listFonts().length;
    for (let i = 0; i < 8; i += 1) {
      assert.throws(() => wb.ooxml.addFont(lateFont), /late font getter/);
    }
    assert.equal(wb.ooxml.listFonts().length, beforeFonts);
    assert.equal(wb.getMemoryUsage(), beforeMemory);
  } finally {
    wb.dispose();
  }
});

test('conditional-format arrays skip nullish entries and keep later records', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    wb.sheet(0).addCfRule({
      // Include null, undefined, and a hole before the valid range. The old
      // reader skipped those entries and still passed the later range to C.
      ranges: Object.assign(new Array(4), { 0: null, 1: undefined, 3: RANGE }),
      type: 3,
      dataBar: { min: { type: 3 }, max: { type: 4 } },
    });
    const stored = wb.sheet(0).listCfRules();
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
    const styleCount = wb.ooxml.listCellStyles().length;
    for (const value of [2 ** 32, Number.NaN, 1.5, -1]) {
      assert.throws(() => wb.ooxml.setCellStyle({ name: 'invalid', xfId: value }), RangeError);
      assert.equal(wb.ooxml.listCellStyles().length, styleCount);
    }

    const before = wb.getThemeColors().colors;
    for (const value of [Symbol('theme'), Number.NaN, 1.5, -1, 2 ** 32]) {
      const values = Array.from({ length: 12 }, (_, i) => (i === 0 ? value : before[i]));
      assert.throws(
        () => wb.setThemeColors({ colors: values }),
        typeof value === 'symbol' ? TypeError : /must be a number|outside uint32/,
      );
      assert.deepEqual(wb.getThemeColors().colors, before);
    }
  } finally {
    wb.dispose();
  }
});

test('style and CF reader defaults survive nullish fields and valid u8 boundaries', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const font = wb.ooxml.addFont({ name: 'Reader defaults', size: null, color: undefined, underline: null });
    const storedFont = wb.ooxml.getFont(font);
    assert.equal(storedFont.size, 11);
    assert.equal(storedFont.color.kind, mod.ColorKind.None);

    wb.sheet(0).addCfRule(
      dataBarRule({
        fill: '#FF0001',
        minLengthPct: 0,
        maxLengthPct: 100,
      }),
    );
    const stored = wb.sheet(0).listCfRules()[0].dataBar;
    assert.equal(stored.fill.kind, mod.ColorKind.Rgb);
    assert.equal(stored.fill.rgb, 0xffff0001);
    assert.equal(stored.minLengthPct, 0);
    assert.equal(stored.maxLengthPct, 100);
  } finally {
    wb.dispose();
  }
});
