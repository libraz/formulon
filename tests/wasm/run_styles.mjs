import assert from 'node:assert/strict';

const isNestedRangeValue = (value) => typeof value === 'number';

export function registerStyles(Module, test) {
  test('xf index round-trips through setCellXfIndex / getCellXfIndex', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setNumber(0, 2, 3, 5).ok);
      assert.ok(wb.setCellXfIndex(0, 2, 3, 9).ok);
      const r = wb.getCellXfIndex(0, 2, 3);
      assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
      assert.equal(r.xfIndex, 9);
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
      assert.equal(wb.fontCount().value, 1);
      assert.equal(wb.fillCount().value, 2);
      assert.equal(wb.borderCount().value, 1);
      assert.equal(wb.xfCount().value, 1);

      const f1 = wb.addFont({
        name: 'Arial',
        size: 12,
        bold: true,
        italic: false,
        strike: false,
        underline: 0,
        colorArgb: 0xff112233,
      });
      assert.ok(f1.status.ok, `addFont: ${JSON.stringify(f1.status)}`);
      assert.equal(typeof f1.index, 'number');

      // Adding the same font again returns the same index (linear-search dedup).
      const f1b = wb.addFont({
        name: 'Arial',
        size: 12,
        bold: true,
        italic: false,
        strike: false,
        underline: 0,
        colorArgb: 0xff112233,
      });
      assert.ok(f1b.status.ok);
      assert.equal(f1b.index, f1.index);
      assert.equal(wb.fontCount().value, 2);

      const fill = wb.addFill({ pattern: 1, fgArgb: 0xffff0000, bgArgb: 0xff000000 });
      assert.ok(fill.status.ok);
      const border = wb.addBorder({
        left: { style: 1, colorArgb: 0xff000000 },
        right: { style: 1, colorArgb: 0xff000000 },
        top: { style: 1, colorArgb: 0xff000000 },
        bottom: { style: 1, colorArgb: 0xff000000 },
        diagonal: { style: 0, colorArgb: 0 },
        diagonalUp: false,
        diagonalDown: false,
      });
      assert.ok(border.status.ok);

      // Built-in num_fmt resolves without growing the table.
      const builtin = wb.addNumFmt('General');
      assert.ok(builtin.status.ok);
      assert.equal(builtin.numFmtId, 0);

      // Custom num_fmt yields >= 164.
      const custom = wb.addNumFmt('"USD" #,##0');
      assert.ok(custom.status.ok);
      assert.ok(custom.numFmtId >= 164, `expected custom id >= 164, got ${custom.numFmtId}`);

      const xf = wb.addXf({
        fontIndex: f1.index,
        fillIndex: fill.index,
        borderIndex: border.index,
        numFmtId: custom.numFmtId,
        horizontalAlign: 1,
        verticalAlign: 2,
        wrapText: true,
      });
      assert.ok(xf.status.ok, `addXf: ${JSON.stringify(xf.status)}`);
      assert.ok(wb.xfCount().value >= 1);

      // Adding the same xf is a no-op.
      const xfDup = wb.addXf({
        fontIndex: f1.index,
        fillIndex: fill.index,
        borderIndex: border.index,
        numFmtId: custom.numFmtId,
        horizontalAlign: 1,
        verticalAlign: 2,
        wrapText: true,
      });
      assert.ok(xfDup.status.ok);
      assert.equal(xfDup.index, xf.index);

      // Stamp the xf on a cell.
      assert.ok(wb.setNumber(0, 0, 0, 1).ok);
      assert.ok(wb.setCellXfIndex(0, 0, 0, xf.index).ok);

      // Read back through the getters.
      const reread = wb.getCellXf(xf.index);
      assert.ok(reread.status.ok);
      assert.equal(reread.fontIndex, f1.index);
      assert.equal(reread.numFmtId, custom.numFmtId);
      assert.equal(reread.wrapText, true);
      assert.equal(reread.hasAlignment, true);
      assert.equal(reread.hasHorizontalAlign, true);
      assert.equal(reread.hasVerticalAlign, true);
      assert.equal(reread.hasWrapText, true);
      assert.equal(reread.hasJustifyLastLine, false);

      const optional = wb.addXf({
        fontIndex: f1.index,
        fillIndex: fill.index,
        borderIndex: border.index,
        numFmtId: custom.numFmtId,
        textRotation: 255,
        indent: 0,
        relativeIndent: -3,
        shrinkToFit: false,
        readingOrder: 0,
        justifyLastLine: true,
      });
      assert.ok(optional.status.ok, `optional alignment: ${JSON.stringify(optional.status)}`);
      const optionalRead = wb.getCellXf(optional.index);
      assert.ok(optionalRead.status.ok);
      assert.equal(optionalRead.textRotation, 255);
      assert.equal(optionalRead.indent, 0);
      assert.equal(optionalRead.relativeIndent, -3);
      assert.equal(optionalRead.shrinkToFit, false);
      assert.equal(optionalRead.readingOrder, 0);
      assert.equal(optionalRead.justifyLastLine, true);
      assert.equal(optionalRead.hasAlignment, true);
      assert.equal(optionalRead.hasJustifyLastLine, true);
      assert.equal(wb.addXf(optionalRead).index, optional.index);

      const omitted = wb.addXf({
        fontIndex: f1.index,
        fillIndex: fill.index,
        borderIndex: border.index,
        numFmtId: custom.numFmtId,
      });
      assert.ok(omitted.status.ok);
      const omittedRead = wb.getCellXf(omitted.index);
      assert.ok(omittedRead.status.ok);
      assert.equal(omittedRead.hasAlignment, false);

      const explicitEmpty = wb.addXf({
        fontIndex: f1.index,
        fillIndex: fill.index,
        borderIndex: border.index,
        numFmtId: custom.numFmtId,
        hasAlignment: true,
      });
      assert.ok(explicitEmpty.status.ok);
      assert.notEqual(explicitEmpty.index, omitted.index);
      const explicitEmptyRead = wb.getCellXf(explicitEmpty.index);
      assert.ok(explicitEmptyRead.status.ok);
      assert.equal(explicitEmptyRead.hasAlignment, true);

      const rfont = wb.getFont(f1.index);
      assert.ok(rfont.status.ok);
      assert.equal(rfont.name, 'Arial');
      assert.equal(rfont.size, 12);
      assert.equal(rfont.bold, true);

      const rfill = wb.getFill(fill.index);
      assert.ok(rfill.status.ok);
      assert.equal(rfill.pattern, 1);
      assert.equal(rfill.fgArgb, 0xffff0000);

      const rborder = wb.getBorder(border.index);
      assert.ok(rborder.status.ok);
      assert.equal(rborder.left.style, 1);
      assert.equal(rborder.diagonalUp, false);

      const rfmt = wb.getNumFmt(custom.numFmtId);
      assert.ok(rfmt.status.ok);
      assert.equal(rfmt.formatCode, '"USD" #,##0');
    } finally {
      wb.delete();
    }
  });

  test('selector colours survive get/add identity and save/load', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const theme = { kind: 2, rgb: 0, theme: 3, tint: 0.5, indexed: 0 };
      const indexed = { kind: 3, rgb: 0, theme: 0, tint: 0, indexed: 9 };
      const automatic = { kind: 4, rgb: 0, theme: 0, tint: 0, indexed: 0 };
      const font = wb.addFont({
        name: 'SelectorFont',
        size: 11,
        colorArgb: 0x01020304,
        color: theme,
      });
      assert.ok(font.status.ok);
      const gotFont = wb.getFont(font.index);
      assert.ok(gotFont.status.ok);
      assert.equal(gotFont.color.kind, 2);
      assert.equal(gotFont.color.theme, 3);
      assert.equal(gotFont.color.tint, 0.5);
      assert.equal(gotFont.colorArgb, 0x01020304);
      const fontAgain = wb.addFont(gotFont);
      assert.ok(fontAgain.status.ok);
      assert.equal(fontAgain.index, font.index);

      const fill = wb.addFill({
        pattern: 1,
        fgArgb: 0x05060708,
        bgArgb: 0x090a0b0c,
        fg: indexed,
        bg: automatic,
      });
      assert.ok(fill.status.ok);
      const gotFill = wb.getFill(fill.index);
      assert.ok(gotFill.status.ok);
      assert.equal(gotFill.fg.kind, 3);
      assert.equal(gotFill.fg.indexed, 9);
      assert.equal(gotFill.bg.kind, 4);
      const fillAgain = wb.addFill(gotFill);
      assert.ok(fillAgain.status.ok);
      assert.equal(fillAgain.index, fill.index);

      const border = wb.addBorder({
        left: { style: 1, colorArgb: 0x01020304, color: theme },
        right: { style: 1, colorArgb: 0x05060708, color: indexed },
        top: { style: 1, colorArgb: 0x090a0b0c, color: automatic },
        bottom: { style: 0, colorArgb: 0 },
        diagonal: { style: 0, colorArgb: 0 },
      });
      assert.ok(border.status.ok);
      const gotBorder = wb.getBorder(border.index);
      assert.ok(gotBorder.status.ok);
      assert.equal(gotBorder.left.color.kind, 2);
      assert.equal(gotBorder.left.color.theme, 3);
      assert.equal(gotBorder.right.color.kind, 3);
      assert.equal(gotBorder.right.color.indexed, 9);
      assert.equal(gotBorder.top.color.kind, 4);
      const borderAgain = wb.addBorder(gotBorder);
      assert.ok(borderAgain.status.ok);
      assert.equal(borderAgain.index, border.index);

      const dxf = wb.addDxf({
        font: { name: 'DxfSelector', size: 9, colorArgb: 0x11121314, color: automatic },
        fill: { pattern: 1, fgArgb: 0x15161718, fg: indexed },
        border: { left: { style: 1, colorArgb: 0x191a1b1c, color: theme } },
      });
      assert.ok(dxf.status.ok);
      const gotDxf = wb.getDxf(dxf.index);
      assert.ok(gotDxf.status.ok);
      assert.equal(gotDxf.font.color.kind, 4);
      assert.equal(gotDxf.fill.fg.kind, 3);
      assert.equal(gotDxf.border.left.color.kind, 2);
      const dxfAgain = wb.addDxf(gotDxf);
      assert.ok(dxfAgain.status.ok);
      assert.equal(dxfAgain.index, dxf.index);

      const saved = wb.save();
      assert.ok(saved.status.ok);
      const loaded = Module.Workbook.loadBytes(saved.bytes);
      try {
        assert.ok(loaded.isValid());
        const loadedFont = loaded.getFont(font.index);
        const loadedFill = loaded.getFill(fill.index);
        const loadedBorder = loaded.getBorder(border.index);
        const loadedDxf = loaded.getDxf(dxf.index);
        assert.ok(loadedFont.status.ok);
        assert.ok(loadedFill.status.ok);
        assert.ok(loadedBorder.status.ok);
        assert.ok(loadedDxf.status.ok);
        assert.equal(loadedFont.color.kind, 2);
        assert.equal(loadedFill.fg.kind, 3);
        assert.equal(loadedBorder.left.color.kind, 2);
        assert.equal(loadedDxf.font.color.kind, 4);
      } finally {
        loaded.delete();
      }
    } finally {
      wb.delete();
    }
  });

  test('dxf alignment and protection XML survive get/add identity and save/load', () => {
    const wb = Module.Workbook.createDefault();
    const alignmentXml = '<alignment horizontal="center" wrapText="1"/>';
    const protectionXml = '<protection locked="0" hidden="1"/>';
    try {
      const alignment = wb.addDxf({ alignmentXml });
      assert.ok(alignment.status.ok);
      const protection = wb.addDxf({ protectionXml });
      assert.ok(protection.status.ok);
      assert.notEqual(alignment.index, protection.index);

      const gotAlignment = wb.getDxf(alignment.index);
      assert.ok(gotAlignment.status.ok);
      assert.equal(gotAlignment.alignmentXml, alignmentXml);
      assert.equal(gotAlignment.protectionXml, undefined);
      const gotProtection = wb.getDxf(protection.index);
      assert.ok(gotProtection.status.ok);
      assert.equal(gotProtection.alignmentXml, undefined);
      assert.equal(gotProtection.protectionXml, protectionXml);

      const alignmentAgain = wb.addDxf(gotAlignment);
      assert.ok(alignmentAgain.status.ok);
      assert.equal(alignmentAgain.index, alignment.index);
      const protectionAgain = wb.addDxf(gotProtection);
      assert.ok(protectionAgain.status.ok);
      assert.equal(protectionAgain.index, protection.index);

      const saved = wb.save();
      assert.ok(saved.status.ok);
      const loaded = Module.Workbook.loadBytes(saved.bytes);
      try {
        assert.ok(loaded.isValid());
        const loadedAlignment = loaded.getDxf(alignment.index);
        assert.ok(loadedAlignment.status.ok);
        assert.equal(loadedAlignment.alignmentXml, alignmentXml);
        assert.equal(loadedAlignment.protectionXml, undefined);
        const loadedProtection = loaded.getDxf(protection.index);
        assert.ok(loadedProtection.status.ok);
        assert.equal(loadedProtection.alignmentXml, undefined);
        assert.equal(loadedProtection.protectionXml, protectionXml);
        const loadedAlignmentAgain = loaded.addDxf(loadedAlignment);
        assert.ok(loadedAlignmentAgain.status.ok);
        assert.equal(loadedAlignmentAgain.index, alignment.index);
        const loadedProtectionAgain = loaded.addDxf(loadedProtection);
        assert.ok(loadedProtectionAgain.status.ok);
        assert.equal(loadedProtectionAgain.index, protection.index);
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
      const superscript = wb.addFont({ name: 'Arial', size: 12, vertAlign: 1 });
      assert.ok(superscript.status.ok);
      assert.equal(wb.getFont(superscript.index).vertAlign, 1);

      const subscript = wb.addFont({ name: 'Arial', size: 12, vertAlign: 2 });
      assert.ok(subscript.status.ok);
      assert.notEqual(subscript.index, superscript.index);
      assert.equal(wb.getFont(subscript.index).vertAlign, 2);
    } finally {
      wb.delete();
    }
  });

  test('addFont(getFont(i)) is the identity and does not grow the table', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const added = wb.addFont({ name: 'Arial', size: 12, vertAlign: 1, colorArgb: 0xff112233 });
      assert.ok(added.status.ok);
      const before = wb.fontCount().value;
      const readBack = wb.getFont(added.index);
      const again = wb.addFont(readBack);
      assert.ok(again.status.ok);
      assert.equal(again.index, added.index);
      assert.equal(wb.fontCount().value, before);
    } finally {
      wb.delete();
    }
  });

  test('a one-field font rewrite keeps the superscript', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const added = wb.addFont({ name: 'Arial', size: 12, vertAlign: 1, colorArgb: 0xff112233 });
      assert.ok(added.status.ok);
      const edited = wb.getFont(added.index);
      edited.colorArgb = 0xff00ff00;
      const recolored = wb.addFont(edited);
      assert.ok(recolored.status.ok);
      assert.notEqual(recolored.index, added.index);
      const reread = wb.getFont(recolored.index);
      assert.equal(reread.vertAlign, 1);
      assert.equal(reread.colorArgb, 0xff00ff00);
    } finally {
      wb.delete();
    }
  });

  test('addDxf / getDxf round-trip a superscript differential font', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const added = wb.addDxf({ font: { name: 'Calibri', size: 9, vertAlign: 1 } });
      assert.ok(added.status.ok);
      const readBack = wb.getDxf(added.index);
      assert.equal(readBack.font.vertAlign, 1);
      const before = wb.dxfCount().value;
      const again = wb.addDxf(readBack);
      assert.ok(again.status.ok);
      assert.equal(again.index, added.index);
      assert.equal(wb.dxfCount().value, before);
    } finally {
      wb.delete();
    }
  });

  test('addXf rejects out-of-range font_index', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const r = wb.addXf({
        fontIndex: 99,
        fillIndex: 0,
        borderIndex: 0,
        numFmtId: 0,
        horizontalAlign: 0,
        verticalAlign: 0,
        wrapText: false,
      });
      assert.ok(!r.status.ok, 'expected addXf to reject out-of-range font_index');
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

      const fontCount = () => wb.fontCount().value;
      fontCount.before = fontCount();
      assertRejected('font.underline=256', () => wb.addFont({ name: 'bad-u8', underline: 256 }), fontCount);

      const fillCount = () => wb.fillCount().value;
      fillCount.before = fillCount();
      assertRejected('fill.pattern=256', () => wb.addFill({ pattern: 256 }), fillCount);

      const borderCount = () => wb.borderCount().value;
      borderCount.before = borderCount();
      assertRejected('border.left.style=256', () => wb.addBorder({ left: { style: 256 } }), borderCount);

      const xfCount = () => wb.xfCount().value;
      xfCount.before = xfCount();
      assertRejected(
        'xf.numFmtId=65536',
        () => wb.addXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 65536 }),
        xfCount,
      );

      const dxfCount = () => wb.dxfCount().value;
      dxfCount.before = dxfCount();
      assertRejected('dxf.numFmt.numFmtId=65536', () => wb.addDxf({ numFmt: { numFmtId: 65536 } }), dxfCount);
    } finally {
      wb.delete();
    }
  });

  test('style nested numeric fields throw on fractions, non-finite values, and wrong primitives', () => {
    const invalid = [1.5, Number.NaN, Number.POSITIVE_INFINITY, '1', true, {}, Symbol('numeric')];
    for (const value of invalid) {
      const wb = Module.Workbook.createDefault();
      try {
        const before = wb.fontCount().value;
        assert.throws(
          () => wb.addFont({ name: 'bad-value', underline: value }),
          isNestedRangeValue(value) ? RangeError : TypeError,
          `font.underline=${String(value)} accepted`,
        );
        assert.equal(wb.fontCount().value, before, `font.underline=${String(value)} mutated table`);
      } finally {
        wb.delete();
      }
    }
  });

  test('style and conditional-format u32/i32 fields throw on wraparound values', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const xfBefore = wb.xfCount().value;
      assert.throws(() => wb.addXf({ fontIndex: 0x100000000, fillIndex: 0, borderIndex: 0 }), RangeError);
      assert.equal(wb.xfCount().value, xfBefore);

      const fontBefore = wb.fontCount().value;
      assert.throws(() => wb.addFont({ name: 'bad-u32', colorArgb: 0x100000000 }), RangeError);
      assert.equal(wb.fontCount().value, fontBefore);

      const styleXfBefore = wb.cellStyleXfCount().value;
      assert.throws(
        () => wb.addCellStyleXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, xfId: 0x100000000 }),
        RangeError,
      );
      assert.equal(wb.cellStyleXfCount().value, styleXfBefore);

      const themeBefore = wb.getTheme().colors;
      assert.throws(() => wb.setThemeColors([0x100000000, ...themeBefore.slice(1)]), RangeError);
      assert.deepEqual(wb.getTheme().colors, themeBefore);

      const dxf = wb.addDxf({ font: { name: 'cf-dxf' } });
      assert.ok(dxf.status.ok);
      const base = {
        sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 3,
        priority: 0x100000000,
        dxfId: dxf.index,
        dataBar: { min: { type: 3 }, max: { type: 4 }, fill: { r: 1, g: 2, b: 3 } },
      };
      const cfBefore = wb.getConditionalFormats(0).length;
      assert.throws(() => wb.addConditionalFormat(0, base), RangeError);
      assert.equal(wb.getConditionalFormats(0).length, cfBefore);

      const rangeRule = structuredClone(base);
      rangeRule.priority = 0;
      rangeRule.sqref[0].firstRow = 0x100000000;
      assert.throws(() => wb.addConditionalFormat(0, rangeRule), RangeError);
      assert.equal(wb.getConditionalFormats(0).length, cfBefore);
    } finally {
      wb.delete();
    }
  });

  test('style nullish narrow fields retain their existing defaults', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const absentFont = wb.addFont({ name: 'nullish-font' });
      assert.ok(absentFont.status.ok);
      assert.equal(wb.getFont(absentFont.index).underline, 0);
      const nullFont = wb.addFont({ name: 'null-font', underline: null });
      assert.ok(nullFont.status.ok);
      assert.equal(wb.getFont(nullFont.index).underline, 0);
      const undefinedFont = wb.addFont({ name: 'undefined-font', underline: undefined });
      assert.ok(undefinedFont.status.ok);
      assert.equal(wb.getFont(undefinedFont.index).underline, 0);

      const fill = wb.addFill({ pattern: null });
      assert.ok(fill.status.ok);
      assert.equal(wb.getFill(fill.index).pattern, 0);

      const border = wb.addBorder({ left: { style: undefined } });
      assert.ok(border.status.ok);
      assert.equal(wb.getBorder(border.index).left.style, 0);

      const xf = wb.addXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: null });
      assert.ok(xf.status.ok);
      assert.equal(wb.getCellXf(xf.index).numFmtId, 0);

      const dxf = wb.addDxf({ numFmt: { numFmtId: undefined } });
      assert.ok(dxf.status.ok);
      assert.equal(wb.getDxf(dxf.index).numFmt.numFmtId, 0);
    } finally {
      wb.delete();
    }
  });

  test('conditional-format nested narrow numeric fields throw on masked values without mutation', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const base = {
        sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 3,
        dataBar: {
          min: { type: 3 },
          max: { type: 4 },
          fill: { r: 1, g: 2, b: 3 },
        },
      };
      const invalid = [
        ['dataBar.fill.r=256', { dataBar: { fill: { r: 256 } } }],
        ['dataBar.minLengthPct=256', { dataBar: { minLengthPct: 256 } }],
        ['dataBar.maxLengthPct=256', { dataBar: { maxLengthPct: 256 } }],
        ['cfvo.type=256', { dataBar: { min: { type: 256 } } }],
      ];
      for (const [label, override] of invalid) {
        const before = wb.getConditionalFormats(0).length;
        const rule = structuredClone(base);
        Object.assign(rule.dataBar, override.dataBar);
        if (override.dataBar.min) {
          rule.dataBar.min = override.dataBar.min;
        }
        assert.throws(() => wb.addConditionalFormat(0, rule), RangeError, `${label} accepted`);
        assert.equal(wb.getConditionalFormats(0).length, before, `${label}: mutation changed count`);
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
          sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
          type: 3,
          dataBar: {
            min: { type: 3 },
            max: { type: 4 },
            fill: { r: value, g: 2, b: 3 },
          },
        };
        assert.throws(
          () => wb.addConditionalFormat(0, rule),
          isNestedRangeValue(value) ? RangeError : TypeError,
          `dataBar.fill.r=${String(value)} accepted`,
        );
        assert.equal(wb.getConditionalFormats(0).length, 0, `dataBar.fill.r=${String(value)} mutated rules`);
      } finally {
        wb.delete();
      }
    }
  });

  test('conditional-format nullish nested fields retain existing defaults', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const result = wb.addConditionalFormat(0, {
        sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 3,
        dataBar: {
          min: { type: 3, gte: null },
          max: { type: 4, gte: undefined },
          fill: { r: null, g: undefined, b: 3, a: null },
          minLengthPct: null,
          maxLengthPct: undefined,
          axisPosition: null,
          direction: undefined,
        },
      });
      assert.ok(result.status.ok, JSON.stringify(result.status));
      const bar = wb.getConditionalFormats(0)[0].dataBar;
      assert.equal(bar.fill.r, 0);
      assert.equal(bar.fill.g, 0);
      assert.equal(bar.fill.b, 3);
      assert.equal(bar.fill.a, 255);
      assert.equal(bar.minLengthPct, 10);
      assert.equal(bar.maxLengthPct, 90);
      assert.equal(bar.axisPosition, 0);
      assert.equal(bar.direction, 0);
      assert.equal(wb.getConditionalFormats(0)[0].dataBar.min.gte, true);
    } finally {
      wb.delete();
    }
  });

  test('style and conditional-format readers snapshot stateful fields once', () => {
    const wb = Module.Workbook.createDefault();
    try {
      let horizontalReads = 0;
      const xfSpec = {};
      Object.defineProperty(xfSpec, 'horizontalAlign', {
        get() {
          horizontalReads += 1;
          return horizontalReads === 1 ? 1 : undefined;
        },
      });
      const xf = wb.addXf(xfSpec);
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
      const rule = {
        sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 3,
        dataBar: {
          min: { type: 3 },
          max: { type: 4 },
          fill: { r: 1, g: 2, b: 3 },
        },
      };
      Object.defineProperty(rule, 'dxfId', {
        get() {
          dxfReads += 1;
          return dxfReads === 1 ? secondDxf.index : undefined;
        },
      });
      const cf = wb.addConditionalFormat(0, rule);
      assert.ok(cf.status.ok, JSON.stringify(cf.status));
      assert.equal(dxfReads, 1);
      assert.equal(wb.getConditionalFormats(0)[0].dxfId, secondDxf.index);

      const lateFont = { name: 'late font '.repeat(1024) };
      Object.defineProperty(lateFont, 'size', {
        get() {
          throw new Error('late font getter');
        },
      });
      const beforeFonts = wb.fontCount().value;
      for (let i = 0; i < 8; i += 1) {
        assert.throws(() => wb.addFont(lateFont), /late font getter/);
      }
      assert.equal(wb.fontCount().value, beforeFonts);
    } finally {
      wb.delete();
    }
  });
}
