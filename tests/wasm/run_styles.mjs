import assert from 'node:assert/strict';
import { StatusCode } from '../../packages/npm/common.mjs';
import { expectFailure, readAll } from './smoke_support.mjs';

const isNestedRangeValue = (value) => typeof value === 'number';
const rgb = (value) => ({ kind: 1, argb: value, theme: 0, tint: 0, indexed: 0 });
const NO_COLOR = { kind: 0, argb: 0, theme: 0, tint: 0, indexed: 0 };

export function registerStyles(Module, test) {
  test('xf index round-trips through setCellXfIndex / getCellXfIndex', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.setNumber(0, 2, 3, 5);
      // An index past the xf table is refused and leaves the cell alone.
      expectFailure(() => wb.setCellXfIndex(0, 2, 3, 9), StatusCode.IndexOutOfRange);
      assert.equal(wb.getCellXfIndex(0, 2, 3), 0);
      const font = wb.addFont({ name: 'Arial', sizePt: 12, bold: true });
      const xf = wb.addCellXf({ fontIndex: font });
      wb.setCellXfIndex(0, 2, 3, xf);
      const r = wb.getCellXfIndex(0, 2, 3);
      assert.equal(r, xf);
    } finally {
      wb.delete();
    }
  });

  test('getCellXfIndex reads null for an absent cell and the stored index for a present one', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.getCellXfIndex(0, 98, 25), null);
      wb.setNumber(0, 2, 3, 5);
      assert.equal(wb.getCellXfIndex(0, 2, 3), 0);
      wb.setCellXfIndex(0, 98, 25, 0);
      assert.equal(wb.getCellXfIndex(0, 98, 25), 0);
    } finally {
      wb.delete();
    }
  });

  test('style building blocks: addFont -> addXf -> setCellXfIndex -> getXf', () => {
    const wb = Module.Workbook.createDefault();
    try {
      // A new workbook carries the minimum style table Excel writes, so
      // a caller-appended record never lands on one of the slots Excel
      // reserves: one font, the `none` and `gray125` fills, one border,
      // one xf.
      assert.equal(wb.getFontCount(), 1);
      assert.equal(wb.getFillCount(), 2);
      assert.equal(wb.getBorderCount(), 1);
      assert.equal(wb.getCellXfCount(), 1);

      const f1 = wb.addFont({
        name: 'Arial',
        sizePt: 12,
        bold: true,
        italic: false,
        strike: false,
        underline: 0,
        color: rgb(0xff112233),
      });
      assert.equal(typeof f1, 'number');

      // Adding the same font again returns the same index (linear-search dedup).
      const f1b = wb.addFont({
        name: 'Arial',
        sizePt: 12,
        bold: true,
        italic: false,
        strike: false,
        underline: 0,
        color: rgb(0xff112233),
      });
      assert.equal(f1b, f1);
      assert.equal(wb.getFontCount(), 2);

      const fill = wb.addFill({ pattern: 1, foreground: rgb(0xffff0000), background: rgb(0xff000000) });
      const border = wb.addBorder({
        left: { style: 1, color: rgb(0xff000000) },
        right: { style: 1, color: rgb(0xff000000) },
        top: { style: 1, color: rgb(0xff000000) },
        bottom: { style: 1, color: rgb(0xff000000) },
        diagonal: { style: 0, color: NO_COLOR },
        diagonalUp: false,
        diagonalDown: false,
      });

      // Built-in num_fmt resolves without growing the table.
      const builtin = wb.addNumberFormat('General');
      assert.equal(builtin, 0);

      // Custom num_fmt yields >= 164.
      const custom = wb.addNumberFormat('"USD" #,##0');
      assert.ok(custom >= 164, `expected custom id >= 164, got ${custom}`);

      const xf = wb.addCellXf({
        fontIndex: f1,
        fillIndex: fill,
        borderIndex: border,
        numFmtId: custom,
        horizontal: 1,
        vertical: 2,
        wrapText: true,
      });
      assert.ok(wb.getCellXfCount() >= 1);

      // Adding the same xf is a no-op.
      const xfDup = wb.addCellXf({
        fontIndex: f1,
        fillIndex: fill,
        borderIndex: border,
        numFmtId: custom,
        horizontal: 1,
        vertical: 2,
        wrapText: true,
      });
      assert.equal(xfDup, xf);

      // Stamp the xf on a cell.
      wb.setNumber(0, 0, 0, 1);
      wb.setCellXfIndex(0, 0, 0, xf);

      // Read back through the getters.
      const reread = wb.getCellXf(xf);
      assert.equal(reread.fontIndex, f1);
      assert.equal(reread.numFmtId, custom);
      assert.equal(reread.wrapText, true);
      assert.equal(reread.horizontal, 1);
      assert.equal(reread.vertical, 2);
      assert.equal(reread.justifyLastLine, null);

      const optional = wb.addCellXf({
        fontIndex: f1,
        fillIndex: fill,
        borderIndex: border,
        numFmtId: custom,
        textRotation: 255,
        indent: 0,
        relativeIndent: -3,
        shrinkToFit: false,
        readingOrder: 0,
        justifyLastLine: true,
      });
      const optionalRead = wb.getCellXf(optional);
      assert.equal(optionalRead.textRotation, 255);
      assert.equal(optionalRead.indent, 0);
      assert.equal(optionalRead.relativeIndent, -3);
      assert.equal(optionalRead.shrinkToFit, false);
      assert.equal(optionalRead.readingOrder, 0);
      assert.equal(optionalRead.justifyLastLine, true);
      assert.equal(optionalRead.horizontal, null);
      assert.equal(wb.addCellXf(optionalRead), optional);

      const omitted = wb.addCellXf({
        fontIndex: f1,
        fillIndex: fill,
        borderIndex: border,
        numFmtId: custom,
      });
      const omittedRead = wb.getCellXf(omitted);
      assert.deepEqual(
        [omittedRead.horizontal, omittedRead.vertical, omittedRead.wrapText, omittedRead.justifyLastLine],
        [null, null, null, null],
      );
      assert.equal(omittedRead.textRotation, null);

      // The presence flags are gone: an alignment is present when one of its keys is, so a stray flag adds none.
      const flagged = wb.getCellXf(
        wb.addCellXf({ fontIndex: f1, fillIndex: fill, borderIndex: border, numFmtId: custom, hasAlignment: true }),
      );
      assert.equal(flagged.horizontal, null);

      const rfont = wb.getFont(f1);
      assert.equal(rfont.name, 'Arial');
      assert.equal(rfont.sizePt, 12);
      assert.equal(rfont.bold, true);

      const rfill = wb.getFill(fill);
      assert.equal(rfill.pattern, 1);
      assert.equal(rfill.foreground.kind, 1);
      assert.equal(rfill.foreground.argb, 0xffff0000);

      const rborder = wb.getBorder(border);
      assert.equal(rborder.left.style, 1);
      assert.equal(rborder.diagonalUp, false);

      const rfmt = wb.getNumberFormat(custom);
      assert.equal(rfmt, '"USD" #,##0');
    } finally {
      wb.delete();
    }
  });

  test('selector colours survive get/add identity and save/load', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const theme = { kind: 2, argb: 0, theme: 3, tint: 0.5, indexed: 0 };
      const indexed = { kind: 3, argb: 0, theme: 0, tint: 0, indexed: 9 };
      const automatic = { kind: 4, argb: 0, theme: 0, tint: 0, indexed: 0 };
      const font = wb.addFont({
        name: 'SelectorFont',
        sizePt: 11,
        color: theme,
      });
      const gotFont = wb.getFont(font);
      assert.equal(gotFont.color.kind, 2);
      assert.equal(gotFont.color.theme, 3);
      assert.equal(gotFont.color.tint, 0.5);
      const fontAgain = wb.addFont(gotFont);
      assert.equal(fontAgain, font);

      const fill = wb.addFill({
        pattern: 1,
        foreground: indexed,
        background: automatic,
      });
      const gotFill = wb.getFill(fill);
      assert.equal(gotFill.foreground.kind, 3);
      assert.equal(gotFill.foreground.indexed, 9);
      assert.equal(gotFill.background.kind, 4);
      const fillAgain = wb.addFill(gotFill);
      assert.equal(fillAgain, fill);

      const border = wb.addBorder({
        left: { style: 1, color: theme },
        right: { style: 1, color: indexed },
        top: { style: 1, color: automatic },
        bottom: { style: 0, color: NO_COLOR },
        diagonal: { style: 0, color: NO_COLOR },
      });
      const gotBorder = wb.getBorder(border);
      assert.equal(gotBorder.left.color.kind, 2);
      assert.equal(gotBorder.left.color.theme, 3);
      assert.equal(gotBorder.right.color.kind, 3);
      assert.equal(gotBorder.right.color.indexed, 9);
      assert.equal(gotBorder.top.color.kind, 4);
      const borderAgain = wb.addBorder(gotBorder);
      assert.equal(borderAgain, border);

      const dxf = wb.addDxf({
        font: { name: 'DxfSelector', sizePt: 9, color: automatic },
        fill: { pattern: 1, foreground: indexed },
        border: { left: { style: 1, color: theme } },
      });
      const gotDxf = wb.getDxf(dxf);
      assert.equal(gotDxf.font.color.kind, 4);
      assert.equal(gotDxf.fill.foreground.kind, 3);
      assert.equal(gotDxf.border.left.color.kind, 2);
      const dxfAgain = wb.addDxf(gotDxf);
      assert.equal(gotDxf.font.sizePt, 9);
      const bareDxf = wb.addDxf({ font: { bold: true } });
      assert.equal(wb.getDxf(bareDxf).font.sizePt, null);
      assert.equal(wb.getDxf(bareDxf).font.bold, true);
      assert.equal(wb.getDxf(bareDxf).font.italic, null);
      assert.equal(dxfAgain, dxf);

      const saved = wb.save();
      const loaded = Module.Workbook.load(saved);
      try {
        assert.equal(loaded.disposed, false);
        const loadedFont = loaded.getFont(font);
        const loadedFill = loaded.getFill(fill);
        const loadedBorder = loaded.getBorder(border);
        const loadedDxf = loaded.getDxf(dxf);
        assert.equal(loadedFont.color.kind, 2);
        assert.equal(loadedFill.foreground.kind, 3);
        assert.equal(loadedBorder.left.color.kind, 2);
        assert.equal(loadedDxf.font.color.kind, 4);
      } finally {
        loaded.delete();
      }
    } finally {
      wb.delete();
    }
  });

  test('dxf alignment, protection and number format survive get/add identity and save/load', () => {
    const wb = Module.Workbook.createDefault();
    const alignmentRecord = {
      horizontal: 2,
      vertical: 2,
      wrapText: true,
      shrinkToFit: false,
      indent: 0,
      textRotation: 0,
      readingOrder: 0,
      justifyLastLine: false,
      relativeIndent: 0,
    };
    const protectionRecord = { locked: false, hidden: true };
    try {
      const alignment = wb.addDxf({ alignment: alignmentRecord });
      const protection = wb.addDxf({ protection: protectionRecord });
      assert.notEqual(alignment, protection);

      const gotAlignment = wb.getDxf(alignment);
      assert.deepEqual(gotAlignment.alignment, alignmentRecord);
      assert.equal(gotAlignment.protection, null);
      const gotProtection = wb.getDxf(protection);
      assert.equal(gotProtection.alignment, null);
      assert.deepEqual(gotProtection.protection, protectionRecord);

      const alignmentAgain = wb.addDxf(gotAlignment);
      assert.equal(alignmentAgain, alignment);
      const protectionAgain = wb.addDxf(gotProtection);
      assert.equal(protectionAgain, protection);

      const saved = wb.save();
      const loaded = Module.Workbook.load(saved);
      try {
        assert.equal(loaded.disposed, false);
        const loadedAlignment = loaded.getDxf(alignment);
        assert.deepEqual(loadedAlignment.alignment, alignmentRecord);
        assert.equal(loadedAlignment.protection, null);
        const loadedProtection = loaded.getDxf(protection);
        assert.equal(loadedProtection.alignment, null);
        assert.deepEqual(loadedProtection.protection, protectionRecord);
        const loadedAlignmentAgain = loaded.addDxf(loadedAlignment);
        assert.equal(loadedAlignmentAgain, alignment);
        const loadedProtectionAgain = loaded.addDxf(loadedProtection);
        assert.equal(loadedProtectionAgain, protection);
      } finally {
        loaded.delete();
      }
    } finally {
      wb.delete();
    }
  });

  test('addFont / getFont preserve superscript and subscript', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const superscript = wb.addFont({ name: 'Arial', sizePt: 12, vertAlign: 1 });
      assert.equal(wb.getFont(superscript).vertAlign, 1);

      const subscript = wb.addFont({ name: 'Arial', sizePt: 12, vertAlign: 2 });
      assert.notEqual(subscript, superscript);
      assert.equal(wb.getFont(subscript).vertAlign, 2);
    } finally {
      wb.delete();
    }
  });

  test('addFont(getFont(i)) is the identity and does not grow the table', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const added = wb.addFont({ name: 'Arial', sizePt: 12, vertAlign: 1, color: rgb(0xff112233) });
      const before = wb.getFontCount();
      const readBack = wb.getFont(added);
      const again = wb.addFont(readBack);
      assert.equal(again, added);
      assert.equal(wb.getFontCount(), before);
    } finally {
      wb.delete();
    }
  });

  test('a one-field font rewrite keeps the superscript', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const added = wb.addFont({ name: 'Arial', sizePt: 12, vertAlign: 1, color: rgb(0xff112233) });
      const edited = wb.getFont(added);
      edited.color = rgb(0xff00ff00);
      const recolored = wb.addFont(edited);
      assert.notEqual(recolored, added);
      const reread = wb.getFont(recolored);
      assert.equal(reread.vertAlign, 1);
      assert.equal(reread.color.argb, 0xff00ff00);
    } finally {
      wb.delete();
    }
  });

  test('addDxf / getDxf round-trip a superscript differential font', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const added = wb.addDxf({ font: { name: 'Calibri', sizePt: 9, vertAlign: 1 } });
      const readBack = wb.getDxf(added);
      assert.equal(readBack.font.vertAlign, 1);
      const before = wb.getDxfCount();
      const again = wb.addDxf(readBack);
      assert.equal(again, added);
      assert.equal(wb.getDxfCount(), before);
    } finally {
      wb.delete();
    }
  });

  test('addXf rejects out-of-range font_index', () => {
    const wb = Module.Workbook.createDefault();
    try {
      expectFailure(() =>
        wb.addCellXf({
          fontIndex: 99,
          fillIndex: 0,
          borderIndex: 0,
          numFmtId: 0,
          horizontal: 0,
          vertical: 0,
          wrapText: false,
        }),
      );
    } finally {
      wb.delete();
    }
  });

  test('style narrow numeric fields throw on masked values without mutating tables', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const assertRejected = (label, invoke, count) => {
        assert.throws(invoke, RangeError, `${label} accepted`);
        assert.equal(count(), count.before, `${label}: mutation changed table count`);
      };

      const fontCount = () => wb.getFontCount();
      fontCount.before = fontCount();
      assertRejected('font.underline=256', () => wb.addFont({ name: 'bad-u8', underline: 256 }), fontCount);

      const fillCount = () => wb.getFillCount();
      fillCount.before = fillCount();
      assertRejected('fill.pattern=256', () => wb.addFill({ pattern: 256 }), fillCount);

      const borderCount = () => wb.getBorderCount();
      borderCount.before = borderCount();
      assertRejected('border.left.style=256', () => wb.addBorder({ left: { style: 256 } }), borderCount);

      const xfCount = () => wb.getCellXfCount();
      xfCount.before = xfCount();
      assertRejected(
        'xf.numFmtId=65536',
        () => wb.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 65536 }),
        xfCount,
      );

      const dxfCount = () => wb.getDxfCount();
      dxfCount.before = dxfCount();
      assertRejected('dxf.alignment.horizontal=256', () => wb.addDxf({ alignment: { horizontal: 256 } }), dxfCount);
    } finally {
      wb.delete();
    }
  });

  test('style nested numeric fields throw on fractions, non-finite values, and wrong primitives', () => {
    const invalid = [1.5, Number.NaN, Number.POSITIVE_INFINITY, '1', true, {}, Symbol('numeric')];
    for (const value of invalid) {
      const wb = Module.Workbook.createDefault();
      try {
        const before = wb.getFontCount();
        assert.throws(
          () => wb.addFont({ name: 'bad-value', underline: value }),
          isNestedRangeValue(value) ? RangeError : TypeError,
          `font.underline=${String(value)} accepted`,
        );
        assert.equal(wb.getFontCount(), before, `font.underline=${String(value)} mutated table`);
      } finally {
        wb.delete();
      }
    }
  });

  test('style and conditional-format u32/i32 fields throw on wraparound values', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const xfBefore = wb.getCellXfCount();
      assert.throws(() => wb.addCellXf({ fontIndex: 0x100000000, fillIndex: 0, borderIndex: 0 }), RangeError);
      assert.equal(wb.getCellXfCount(), xfBefore);

      const fontBefore = wb.getFontCount();
      assert.throws(() => wb.addFont({ name: 'bad-u32', color: rgb(0x100000000) }), RangeError);
      assert.equal(wb.getFontCount(), fontBefore);

      const styleXfBefore = wb.getCellStyleXfCount();
      assert.throws(
        () => wb.addCellStyleXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, styleXfIndex: 0x100000000 }),
        RangeError,
      );
      assert.equal(wb.getCellStyleXfCount(), styleXfBefore);

      const themeBefore = wb.getThemeColors().colors;
      assert.throws(
        () => wb.setThemeColors([rgb(0x100000000), ...themeBefore.slice(1).map((c) => rgb(c))]),
        RangeError,
      );
      assert.deepEqual(wb.getThemeColors().colors, themeBefore);

      const dxf = wb.addDxf({ font: { name: 'cf-dxf' } });
      const base = {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 3,
        priority: 0x100000000,
        dxfIndex: dxf,
        dataBar: { min: { type: 3 }, max: { type: 4 }, fill: rgb(0xff010203) },
      };
      const cfBefore = readAll(
        () => wb.getCfRuleCount(0),
        (i) => wb.getCfRule(0, i),
      ).length;
      assert.throws(() => wb.addCfRule(0, base), RangeError);
      assert.equal(
        readAll(
          () => wb.getCfRuleCount(0),
          (i) => wb.getCfRule(0, i),
        ).length,
        cfBefore,
      );

      const rangeRule = structuredClone(base);
      rangeRule.priority = 0;
      rangeRule.ranges[0].firstRow = 0x100000000;
      assert.throws(() => wb.addCfRule(0, rangeRule), RangeError);
      assert.equal(
        readAll(
          () => wb.getCfRuleCount(0),
          (i) => wb.getCfRule(0, i),
        ).length,
        cfBefore,
      );

      const notArray = structuredClone(base);
      notArray.priority = 0;
      notArray.ranges = 'A1';
      assert.throws(() => wb.addCfRule(0, notArray), TypeError);
      assert.equal(
        readAll(
          () => wb.getCfRuleCount(0),
          (i) => wb.getCfRule(0, i),
        ).length,
        cfBefore,
      );
    } finally {
      wb.delete();
    }
  });

  test('style nullish narrow fields retain their existing defaults', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const absentFont = wb.addFont({ name: 'nullish-font' });
      assert.equal(wb.getFont(absentFont).underline, 0);
      const nullFont = wb.addFont({ name: 'null-font', underline: null });
      assert.equal(wb.getFont(nullFont).underline, 0);
      const undefinedFont = wb.addFont({ name: 'undefined-font', underline: undefined });
      assert.equal(wb.getFont(undefinedFont).underline, 0);

      const fill = wb.addFill({ pattern: null });
      assert.equal(wb.getFill(fill).pattern, 0);

      const border = wb.addBorder({ left: { style: undefined } });
      assert.equal(wb.getBorder(border).left.style, 0);

      const xf = wb.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: null });
      assert.equal(wb.getCellXf(xf).numFmtId, 0);

      const dxf = wb.addDxf({ numberFormat: undefined });
      assert.equal(wb.getDxf(dxf).numberFormat, null);
      const formatted = wb.addDxf({ numberFormat: '0.00%' });
      assert.equal(wb.getDxf(formatted).numberFormat, '0.00%');
      assert.equal(wb.addDxf(wb.getDxf(formatted)), formatted);
    } finally {
      wb.delete();
    }
  });

  test('conditional-format nested narrow numeric fields throw on masked values without mutation', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const base = {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 3,
        dataBar: {
          min: { type: 3 },
          max: { type: 4 },
          fill: rgb(0xff010203),
        },
      };
      const invalid = [
        ['dataBar.fill.argb=2^32', { dataBar: { fill: rgb(0x100000000) } }],
        ['dataBar.fill.kind=256', { dataBar: { fill: { ...rgb(0), kind: 256 } } }],
        ['dataBar.minLengthPct=256', { dataBar: { minLengthPct: 256 } }],
        ['dataBar.maxLengthPct=256', { dataBar: { maxLengthPct: 256 } }],
        ['cfvo.type=256', { dataBar: { min: { type: 256 } } }],
      ];
      for (const [label, override] of invalid) {
        const before = readAll(
          () => wb.getCfRuleCount(0),
          (i) => wb.getCfRule(0, i),
        ).length;
        const rule = structuredClone(base);
        Object.assign(rule.dataBar, override.dataBar);
        if (override.dataBar.min) {
          rule.dataBar.min = override.dataBar.min;
        }
        assert.throws(() => wb.addCfRule(0, rule), RangeError, `${label} accepted`);
        assert.equal(
          readAll(
            () => wb.getCfRuleCount(0),
            (i) => wb.getCfRule(0, i),
          ).length,
          before,
          `${label}: mutation changed count`,
        );
      }
    } finally {
      wb.delete();
    }
  });

  test('conditional-format nested numeric fields throw on wrong primitives and preserve settings', () => {
    const values = [1.5, Number.NaN, Number.POSITIVE_INFINITY, '1', true, {}, Symbol('cf-numeric')];
    for (const value of values) {
      const wb = Module.Workbook.createDefault();
      try {
        const rule = {
          ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          type: 3,
          dataBar: {
            min: { type: 3 },
            max: { type: 4 },
            fill: { ...rgb(3), argb: value },
          },
        };
        assert.throws(
          () => wb.addCfRule(0, rule),
          isNestedRangeValue(value) ? RangeError : TypeError,
          `dataBar.fill.argb=${String(value)} accepted`,
        );
        assert.equal(
          readAll(
            () => wb.getCfRuleCount(0),
            (i) => wb.getCfRule(0, i),
          ).length,
          0,
          `dataBar.fill.argb=${String(value)} mutated rules`,
        );
      } finally {
        wb.delete();
      }
    }
  });

  test('conditional-format nullish nested fields retain existing defaults', () => {
    const wb = Module.Workbook.createDefault();
    try {
      wb.addCfRule(0, {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 3,
        dataBar: {
          min: { type: 3, gte: null },
          max: { type: 4, gte: undefined },
          fill: { kind: 2, argb: null, theme: 3, tint: undefined, indexed: null },
          minLengthPct: null,
          maxLengthPct: undefined,
          axisPosition: null,
          direction: undefined,
        },
      });
      const bar = readAll(
        () => wb.getCfRuleCount(0),
        (i) => wb.getCfRule(0, i),
      )[0].dataBar;
      assert.equal(bar.fill.kind, 2);
      assert.equal(bar.fill.argb, 0);
      assert.equal(bar.fill.theme, 3);
      assert.equal(bar.fill.tint, 0);
      assert.equal(bar.fill.indexed, 0);
      assert.equal(bar.minLengthPct, 10);
      assert.equal(bar.maxLengthPct, 90);
      assert.equal(bar.axisPosition, 0);
      assert.equal(bar.direction, 0);
      assert.equal(
        readAll(
          () => wb.getCfRuleCount(0),
          (i) => wb.getCfRule(0, i),
        )[0].dataBar.min.gte,
        true,
      );
    } finally {
      wb.delete();
    }
  });

  test('style and conditional-format readers snapshot stateful fields once', () => {
    const wb = Module.Workbook.createDefault();
    try {
      let horizontalReads = 0;
      const xfSpec = {};
      Object.defineProperty(xfSpec, 'horizontal', {
        get() {
          horizontalReads += 1;
          return horizontalReads === 1 ? 1 : undefined;
        },
      });
      const xf = wb.addCellXf(xfSpec);
      assert.equal(horizontalReads, 1);
      const storedXf = wb.getCellXf(xf);
      assert.equal(storedXf.horizontal, 1);

      wb.addDxf({ font: { name: 'first' } });
      const secondDxf = wb.addDxf({ font: { name: 'second' } });
      let dxfReads = 0;
      const rule = {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 3,
        dataBar: {
          min: { type: 3 },
          max: { type: 4 },
          fill: rgb(0xff010203),
        },
      };
      Object.defineProperty(rule, 'dxfIndex', {
        get() {
          dxfReads += 1;
          return dxfReads === 1 ? secondDxf : undefined;
        },
      });
      wb.addCfRule(0, rule);
      assert.equal(dxfReads, 1);
      assert.equal(
        readAll(
          () => wb.getCfRuleCount(0),
          (i) => wb.getCfRule(0, i),
        )[0].dxfIndex,
        secondDxf,
      );

      const lateFont = { name: 'late font '.repeat(1024) };
      Object.defineProperty(lateFont, 'sizePt', {
        get() {
          throw new Error('late font getter');
        },
      });
      const beforeFonts = wb.getFontCount();
      for (let i = 0; i < 8; i += 1) {
        assert.throws(() => wb.addFont(lateFont), /late font getter/);
      }
      assert.equal(wb.getFontCount(), beforeFonts);
    } finally {
      wb.delete();
    }
  });
}
