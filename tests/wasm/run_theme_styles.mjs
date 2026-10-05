import assert from 'node:assert/strict';

const SOURCE_PART = 0;
const SOURCE_DEFAULT = 1;
const CONTEXT_FONT = 0;
const CONTEXT_FILL_FG = 1;
const RESOLUTION_EXACT = 0;
const RESOLUTION_DEFAULT_THEME = 1;
const RESOLUTION_INDEX_OUT_OF_RANGE = 2;
const RESOLUTION_AUTO_CONTEXT = 4;
const ACCENT1_DEFAULT = 0xff4472c4;

function withWorkbook(Module, fn) {
  const wb = Module.Workbook.createDefault();
  try {
    fn(wb);
  } finally {
    wb.delete();
  }
}

const themeSpec = (theme, tint) => ({ kind: 2, rgb: 0, theme, tint, indexed: 0 });

export function registerThemeStyles(Module, test) {
  test('getTheme reports the default theme until a part exists', () => {
    withWorkbook(Module, (wb) => {
      const t = wb.getTheme();
      assert.ok(t.status.ok);
      assert.equal(t.source, SOURCE_DEFAULT);
      assert.equal(t.colors.length, 12);
      assert.equal(t.colors[4], ACCENT1_DEFAULT);
      assert.equal(t.fonts.majorLatin, 'Calibri Light');
      assert.equal(t.fonts.minorLatin, 'Calibri');
      assert.equal(typeof t.fonts.minorEastAsian, 'string');
      assert.ok(t.fonts.minorEastAsian.length > 0);
    });
  });

  test('setThemeColors / setThemeFonts persist through save and reload', () => {
    withWorkbook(Module, (wb) => {
      const colors = Array.from({ length: 12 }, (_, i) => 0xff101010 + i);
      assert.ok(wb.setThemeColors(colors).ok);
      assert.ok(
        wb.setThemeFonts({
          majorLatin: 'Arial',
          majorEastAsian: 'MS Gothic',
          minorLatin: 'Verdana',
          minorEastAsian: 'MS Mincho',
        }).ok,
      );
      const live = wb.getTheme();
      assert.equal(live.source, SOURCE_PART);
      assert.deepEqual(live.colors, colors);
      assert.equal(live.fonts.majorLatin, 'Arial');

      const saved = wb.save();
      assert.ok(saved.status.ok);
      const reloaded = Module.Workbook.loadBytes(saved.bytes);
      try {
        const t = reloaded.getTheme();
        assert.equal(t.source, SOURCE_PART);
        assert.deepEqual(t.colors, colors);
        assert.equal(t.fonts.minorLatin, 'Verdana');
        assert.equal(t.fonts.minorEastAsian, 'MS Mincho');
      } finally {
        reloaded.delete();
      }
    });
  });

  test('setThemeColors rejects a malformed colour list', () => {
    withWorkbook(Module, (wb) => {
      assert.equal(wb.setThemeColors([1, 2, 3]).ok, false);
      assert.equal(wb.setThemeColors(null).ok, false);
    });
  });

  test('resolveColor reports how each colour was resolved', () => {
    withWorkbook(Module, (wb) => {
      const rgb = wb.resolveColor({ kind: 1, rgb: 0xff123456, theme: 0, tint: 0, indexed: 0 }, CONTEXT_FILL_FG);
      assert.ok(rgb.status.ok);
      assert.equal(rgb.argb, 0xff123456);
      assert.equal(rgb.resolution, RESOLUTION_EXACT);

      const fromDefault = wb.resolveColor(themeSpec(4, 0), CONTEXT_FONT);
      assert.equal(fromDefault.argb, ACCENT1_DEFAULT);
      assert.equal(fromDefault.resolution, RESOLUTION_DEFAULT_THEME);

      assert.ok(wb.setThemeColors(wb.getTheme().colors).ok);
      const lighter = wb.resolveColor(themeSpec(4, 0.4), CONTEXT_FILL_FG);
      assert.equal(lighter.argb, 0xff8ea9db);
      assert.equal(lighter.resolution, RESOLUTION_EXACT);
      assert.equal(wb.resolveColor(themeSpec(0, 0), CONTEXT_FONT).argb, 0xffffffff);

      const outOfRange = wb.resolveColor(themeSpec(12, 0), CONTEXT_FONT);
      assert.equal(outOfRange.argb, 0xff000000);
      assert.equal(outOfRange.resolution, RESOLUTION_INDEX_OUT_OF_RANGE);

      const auto = { kind: 4, rgb: 0, theme: 0, tint: 0, indexed: 0 };
      assert.equal(wb.resolveColor(auto, CONTEXT_FONT).argb, 0xff000000);
      const fill = wb.resolveColor(auto, CONTEXT_FILL_FG);
      assert.equal(fill.argb, 0xffffffff);
      assert.equal(fill.resolution, RESOLUTION_AUTO_CONTEXT);

      assert.equal(wb.resolveColor(themeSpec(4, Number.NaN), CONTEXT_FONT).status.ok, false);
      assert.equal(wb.resolveColor(auto, 99).status.ok, false);
    });
  });

  test('getEffectiveStyle resolves the xf a cell shows', () => {
    withWorkbook(Module, (wb) => {
      const font = wb.addFont({
        name: 'Arial',
        size: 11,
        bold: false,
        italic: false,
        strike: false,
        underline: 0,
        colorArgb: 0,
        color: themeSpec(4, 0),
      });
      const fill = wb.addFill({
        pattern: 1,
        fgArgb: 0xffff0000,
        bgArgb: 0,
        fg: { kind: 1, rgb: 0xffff0000, theme: 0, tint: 0, indexed: 0 },
        bg: { kind: 4, rgb: 0, theme: 0, tint: 0, indexed: 0 },
      });
      const xf = wb.addXf({
        fontIndex: font.index,
        fillIndex: fill.index,
        borderIndex: 0,
        numFmtId: 14,
        locked: false,
        hidden: true,
      });
      assert.ok(xf.status.ok, JSON.stringify(xf.status));
      assert.ok(wb.setCellXfIndex(0, 0, 0, xf.index).ok);

      const e = wb.getEffectiveStyle(0, 0, 0);
      assert.ok(e.status.ok);
      assert.equal(e.xfIndex, xf.index);
      assert.equal(e.source, 0);
      assert.equal(e.fontIndex, font.index);
      assert.equal(e.fillIndex, fill.index);
      assert.equal(e.font.argb, ACCENT1_DEFAULT);
      assert.equal(e.font.resolution, RESOLUTION_DEFAULT_THEME);
      assert.equal(e.fillForeground.argb, 0xffff0000);
      assert.equal(e.fillBackground.resolution, RESOLUTION_AUTO_CONTEXT);
      assert.deepEqual(Object.keys(e.borders), ['left', 'right', 'top', 'bottom', 'diagonal']);
      assert.equal(e.borders.left.resolution, RESOLUTION_AUTO_CONTEXT);
      assert.equal(e.locked, false);
      assert.equal(e.hidden, true);
      assert.equal(e.numFmtCode, 'mm-dd-yy');

      const plain = wb.getEffectiveStyle(0, 5, 5);
      assert.ok(plain.status.ok);
      assert.equal(plain.source, 3);
      assert.equal(plain.xfIndex, 0);
      assert.equal(plain.locked, true);
      assert.equal(plain.numFmtCode, 'General');

      assert.equal(wb.getEffectiveStyle(99, 0, 0).status.ok, false);
    });
  });

  test('setCellStyle takes a record and removeCellStyle repoints formats', () => {
    withWorkbook(Module, (wb) => {
      const before = wb.cellStyleCount().value;
      const styleXf = wb.addCellStyleXf({ numFmtId: 14 });
      assert.ok(styleXf.status.ok);
      assert.ok(wb.setCellStyle({ name: 'Dated', xfId: styleXf.index }).ok);
      assert.equal(wb.cellStyleCount().value, before + 1);
      let found = null;
      for (let i = 0; i < wb.cellStyleCount().value; i += 1) {
        const cs = wb.getCellStyle(i);
        if (cs.name === 'Dated') found = cs;
      }
      assert.ok(found, 'Dated is listed');
      assert.equal(found.xfId, styleXf.index);
      assert.equal(found.builtinId, 0xffffffff);
      assert.equal(found.hidden, false);

      // Replacing by name keeps one entry and takes the new flags.
      assert.ok(wb.setCellStyle({ name: 'Dated', xfId: styleXf.index, hidden: true, customBuiltin: true }).ok);
      assert.equal(wb.cellStyleCount().value, before + 1);

      // A built-in ordinal beyond the old 0..47 range round-trips.
      assert.ok(wb.setCellStyle({ name: 'Normal 2', xfId: 0, builtinId: 53 }).ok);
      let builtin = null;
      for (let i = 0; i < wb.cellStyleCount().value; i += 1) {
        const cs = wb.getCellStyle(i);
        if (cs.name === 'Normal 2') builtin = cs;
      }
      assert.equal(builtin.builtinId, 53);

      assert.ok(wb.removeCellStyle('Dated').ok);
      assert.equal(wb.cellStyleCount().value, before + 1);
      assert.equal(wb.removeCellStyle('Dated').ok, false);
    });
  });
}
