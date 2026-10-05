import assert from 'node:assert/strict';

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
}
