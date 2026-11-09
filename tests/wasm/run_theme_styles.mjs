import assert from 'node:assert/strict';
import { expectFailure } from './smoke_support.mjs';

// FontScheme ordinals (None is 0).
const MAJOR = 1;
const MINOR = 2;
const SOURCE_PART = 0;
const SOURCE_DEFAULT = 1;
const CONTEXT_FONT = 0;
const CONTEXT_FILL_FG = 1;
const RESOLUTION_EXACT = 0;
const RESOLUTION_DEFAULT_THEME = 1;
const RESOLUTION_INDEX_OUT_OF_RANGE = 2;
const RESOLUTION_AUTO_CONTEXT = 4;
const ACCENT1_DEFAULT = 0xff156082;

function withWorkbook(Module, fn) {
  const wb = Module.Workbook.createDefault();
  try {
    fn(wb);
  } finally {
    wb.delete();
  }
}

const themeColor = (argb) => ({ kind: 1, argb, theme: 0, tint: 0, indexed: 0 });
const themeSpec = (theme, tint) => ({ kind: 2, argb: 0, theme, tint, indexed: 0 });

export function registerThemeStyles(Module, test) {
  test('getThemeColors / getThemeFonts report the default theme until a part exists', () => {
    withWorkbook(Module, (wb) => {
      const c = wb.getThemeColors();
      assert.equal(c.source, SOURCE_DEFAULT);
      assert.equal(c.colors.length, 12);
      assert.equal(c.colors[4], ACCENT1_DEFAULT);
      const f = wb.getThemeFonts();
      assert.deepEqual(Object.keys(f).sort(), ['majorEastAsian', 'majorLatin', 'minorEastAsian', 'minorLatin']);
      assert.equal(f.majorLatin, 'Aptos Display');
      assert.equal(f.minorLatin, 'Aptos Narrow');
      assert.equal(typeof f.minorEastAsian, 'string');
      assert.ok(f.minorEastAsian.length > 0);
    });
  });

  test('setThemeColors / setThemeFonts persist through save and reload', () => {
    withWorkbook(Module, (wb) => {
      const colors = Array.from({ length: 12 }, (_, i) => 0xff101010 + i);
      wb.setThemeColors(colors.map(themeColor));
      wb.setThemeFonts({
        majorLatin: 'Arial',
        majorEastAsian: 'MS Gothic',
        minorLatin: 'Verdana',
        minorEastAsian: 'MS Mincho',
      });
      const live = wb.getThemeColors();
      assert.equal(live.source, SOURCE_PART);
      assert.deepEqual(live.colors, colors);
      assert.equal(wb.getThemeFonts().majorLatin, 'Arial');

      const saved = wb.save();
      const reloaded = Module.Workbook.load(saved);
      try {
        const c = reloaded.getThemeColors();
        assert.equal(c.source, SOURCE_PART);
        assert.deepEqual(c.colors, colors);
        const f = reloaded.getThemeFonts();
        assert.equal(f.minorLatin, 'Verdana');
        assert.equal(f.minorEastAsian, 'MS Mincho');
      } finally {
        reloaded.delete();
      }
    });
  });

  test('resetTheme returns an edited theme to the default and drops the part', () => {
    withWorkbook(Module, (wb) => {
      wb.resetTheme();
      assert.equal(wb.getThemeColors().source, SOURCE_DEFAULT);
      const colors = Array.from({ length: 12 }, (_, i) => 0xff101010 + i);
      wb.setThemeColors(colors.map(themeColor));
      assert.equal(wb.getThemeColors().source, SOURCE_PART);
      wb.resetTheme();
      const t = wb.getThemeColors();
      assert.equal(t.source, SOURCE_DEFAULT);
      assert.equal(t.colors[4], ACCENT1_DEFAULT);

      const saved = wb.save();
      assert.equal(Buffer.from(saved).includes('xl/theme/theme1.xml'), false);
      const reloaded = Module.Workbook.load(saved);
      try {
        assert.equal(reloaded.getThemeColors().source, SOURCE_DEFAULT);
      } finally {
        reloaded.delete();
      }
    });
  });

  test('setThemeColors rejects a malformed colour list', () => {
    withWorkbook(Module, (wb) => {
      assert.throws(() => wb.setThemeColors([1, 2, 3]), TypeError);
      assert.throws(() => wb.setThemeColors(Array.from({ length: 12 }, () => themeSpec(4, 0))), TypeError);
      assert.throws(() => wb.setThemeColors(null), TypeError);
    });
  });

  test('resolveColor reports how each colour was resolved', () => {
    withWorkbook(Module, (wb) => {
      const rgb = wb.resolveColor({ kind: 1, argb: 0xff123456, theme: 0, tint: 0, indexed: 0 }, CONTEXT_FILL_FG);
      assert.equal(rgb.argb, 0xff123456);
      assert.equal(rgb.resolution, RESOLUTION_EXACT);

      const fromDefault = wb.resolveColor(themeSpec(4, 0), CONTEXT_FONT);
      assert.equal(fromDefault.argb, ACCENT1_DEFAULT);
      assert.equal(fromDefault.resolution, RESOLUTION_DEFAULT_THEME);

      wb.setThemeColors(wb.getThemeColors().colors.map(themeColor));
      const lighter = wb.resolveColor(themeSpec(4, 0.4), CONTEXT_FILL_FG);
      assert.equal(lighter.argb, 0xff44b3e1);
      assert.equal(lighter.resolution, RESOLUTION_EXACT);
      assert.equal(wb.resolveColor(themeSpec(0, 0), CONTEXT_FONT).argb, 0xffffffff);

      const outOfRange = wb.resolveColor(themeSpec(12, 0), CONTEXT_FONT);
      assert.equal(outOfRange.argb, 0xff000000);
      assert.equal(outOfRange.resolution, RESOLUTION_INDEX_OUT_OF_RANGE);

      const auto = { kind: 4, argb: 0, theme: 0, tint: 0, indexed: 0 };
      assert.equal(wb.resolveColor(auto, CONTEXT_FONT).argb, 0xff000000);
      const fill = wb.resolveColor(auto, CONTEXT_FILL_FG);
      assert.equal(fill.argb, 0xffffffff);
      assert.equal(fill.resolution, RESOLUTION_AUTO_CONTEXT);

      expectFailure(() => wb.resolveColor(themeSpec(4, Number.NaN), CONTEXT_FONT));
      expectFailure(() => wb.resolveColor(auto, 99));
    });
  });

  test('getEffectiveStyle resolves the xf a cell shows', () => {
    withWorkbook(Module, (wb) => {
      const font = wb.addFont({
        name: 'Arial',
        sizePt: 11,
        bold: false,
        italic: false,
        strike: false,
        underline: 0,
        color: themeSpec(4, 0),
      });
      const fill = wb.addFill({
        pattern: 1,
        foreground: { kind: 1, argb: 0xffff0000, theme: 0, tint: 0, indexed: 0 },
        background: { kind: 4, argb: 0, theme: 0, tint: 0, indexed: 0 },
      });
      const xf = wb.addCellXf({
        fontIndex: font,
        fillIndex: fill,
        borderIndex: 0,
        numFmtId: 14,
        locked: false,
        hidden: true,
      });
      wb.setCellXfIndex(0, 0, 0, xf);

      const e = wb.getEffectiveStyle(0, 0, 0);
      assert.equal(e.xfIndex, xf);
      assert.equal(e.source, 0);
      assert.equal(e.fontIndex, font);
      assert.equal(e.fillIndex, fill);
      assert.equal(e.fontColor.argb, ACCENT1_DEFAULT);
      assert.equal(e.fontColor.resolution, RESOLUTION_DEFAULT_THEME);
      assert.equal(e.fillForeground.argb, 0xffff0000);
      assert.equal(e.fillBackground.resolution, RESOLUTION_AUTO_CONTEXT);
      assert.deepEqual(Object.keys(e.borders), ['left', 'right', 'top', 'bottom', 'diagonal']);
      assert.equal(e.borders.left.resolution, RESOLUTION_AUTO_CONTEXT);
      assert.deepEqual(e.protection, { locked: false, hidden: true });
      assert.equal(e.numberFormat, 'm/d/yy');

      const plain = wb.getEffectiveStyle(0, 5, 5);
      assert.equal(plain.source, 3);
      assert.equal(plain.xfIndex, 0);
      assert.equal(plain.protection.locked, true);
      assert.equal(plain.numberFormat, 'General');

      expectFailure(() => wb.getEffectiveStyle(99, 0, 0));
    });
  });

  test('resetCellXfIndex restores row / column inheritance', () => {
    withWorkbook(Module, (wb) => {
      const xfCol = wb.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 14 });
      const xfCell = wb.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 10 });
      wb.setColumnXfIndex(0, 0, 0, xfCol);
      wb.setCellXfIndex(0, 0, 0, xfCell);
      let e = wb.getEffectiveStyle(0, 0, 0);
      assert.equal(e.source, 0);
      assert.equal(e.xfIndex, xfCell);

      wb.resetCellXfIndex(0, 0, 0);
      e = wb.getEffectiveStyle(0, 0, 0);
      assert.equal(e.source, 2);
      assert.equal(e.xfIndex, xfCol);

      wb.setCellXfIndex(0, 0, 0, 0);
      e = wb.getEffectiveStyle(0, 0, 0);
      assert.equal(e.source, 0);
      assert.equal(e.xfIndex, 0);

      const before = wb.getCellCount(0);
      wb.resetCellXfIndex(0, 7, 7);
      assert.equal(wb.getCellCount(0), before);
      expectFailure(() => wb.resetCellXfIndex(99, 0, 0));
    });
  });

  test('attribute patch over a column keeps the row style (bases / remap / snapshot / restore)', () => {
    withWorkbook(Module, (wb) => {
      const rowXf = wb.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 14 });
      const xfA = wb.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 9 });
      const xfB = wb.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 10 });
      wb.setRowXfIndex(0, 2, 2, rowXf);
      assert.equal(wb.getEffectiveStyle(0, 2, 1).source, 1);
      expectFailure(() => wb.setRowXfIndex(0, 5, 4, rowXf));

      const bases = wb.getRangeXfBases(0, { firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 1 });
      assert.ok(bases.includes(0) && bases.includes(rowXf));
      // Row 2 of a one-column selection is a single cell; excluding it hides the row style.
      const excluded = wb.getRangeXfBasesExcluding(
        0,
        { firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 1 },
        [2, 2, 9],
        [1, 1, 9],
      );
      assert.deepEqual(excluded, [0]);
      assert.throws(
        () => wb.getRangeXfBasesExcluding(0, { firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 1 }, [2], []),
        TypeError,
      );

      const snap = wb.snapshotRangeXfs(0, { firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 1 });
      assert.ok(snap.length > 0);
      assert.throws(
        () => wb.remapRangeXfs(0, { firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 1 }, [0, rowXf], [xfA]),
        TypeError,
      );
      wb.remapRangeXfs(0, { firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 1 }, [0, rowXf], [xfA, xfB]);

      let e = wb.getEffectiveStyle(0, 2, 1);
      assert.equal(e.source, 0);
      assert.equal(e.xfIndex, xfB);
      e = wb.getEffectiveStyle(0, 2, 2);
      assert.equal(e.source, 1);
      assert.equal(e.xfIndex, rowXf);
      e = wb.getEffectiveStyle(0, 499, 1);
      assert.equal(e.source, 2);
      assert.equal(e.xfIndex, xfA);

      wb.restoreRangeXfs(0, snap);
      e = wb.getEffectiveStyle(0, 2, 1);
      assert.equal(e.source, 1);
      assert.equal(e.xfIndex, rowXf);
      assert.equal(wb.getEffectiveStyle(0, 499, 1).source, 3);
      expectFailure(() => wb.restoreRangeXfs(0, new Uint8Array([1, 2, 3])));

      wb.resetRowXfIndex(0, 2, 2);
      assert.equal(wb.getEffectiveStyle(0, 2, 2).source, 3);
    });
  });

  test('setCellStyle takes a record and removeCellStyle repoints formats', () => {
    withWorkbook(Module, (wb) => {
      const before = wb.listCellStyles().length;
      const styleXf = wb.addCellStyleXf({ numFmtId: 14 });
      wb.setCellStyle({ name: 'Dated', styleXfIndex: styleXf });
      assert.equal(wb.listCellStyles().length, before + 1);
      let found = null;
      for (let i = 0; i < wb.listCellStyles().length; i += 1) {
        const cs = wb.listCellStyles()[i];
        if (cs.name === 'Dated') found = cs;
      }
      assert.ok(found, 'Dated is listed');
      assert.equal(found.styleXfIndex, styleXf);
      assert.equal(found.builtinId, null);
      assert.equal(found.hidden, false);

      // Replacing by name keeps one entry and takes the new flags.
      wb.setCellStyle({ name: 'Dated', styleXfIndex: styleXf, hidden: true, customBuiltin: true });
      assert.equal(wb.listCellStyles().length, before + 1);

      // A built-in ordinal beyond the old 0..47 range round-trips.
      wb.setCellStyle({ name: 'Normal 2', styleXfIndex: 0, builtinId: 53 });
      let builtin = null;
      for (let i = 0; i < wb.listCellStyles().length; i += 1) {
        const cs = wb.listCellStyles()[i];
        if (cs.name === 'Normal 2') builtin = cs;
      }
      assert.equal(builtin.builtinId, 53);

      wb.removeCellStyle('Dated');
      assert.equal(wb.listCellStyles().length, before + 1);
      expectFailure(() => wb.removeCellStyle('Dated'));
    });
  });

  test('createForProfile builds the locale defaults and throws for a bad id', () => {
    const wb = Module.Workbook.createForProfile('mac-365-ja_JP');
    try {
      assert.equal(wb.disposed, false);
      assert.equal(wb.getProfileId(), 'mac-365-ja_JP');
      const font = wb.getFont(0);
      assert.equal(font.name, '游ゴシック');
      assert.equal(font.sizePt, 12);
    } finally {
      wb.delete();
    }
    expectFailure(() => Module.Workbook.createForProfile('no-such-profile'));
  });

  test('getThemeScriptFont reads the default script faces and setThemeScriptFont round-trips', () => {
    withWorkbook(Module, (wb) => {
      const face = (book, scheme, script) => book.getThemeScriptFont(scheme, script);
      assert.deepEqual(face(wb, MINOR, 'Jpan'), { scheme: MINOR, script: 'Jpan', typeface: '游ゴシック' });
      assert.equal(face(wb, MINOR, 'Hang').typeface, '맑은 고딕');
      const majorHang = face(wb, MAJOR, 'Hang');
      assert.equal(majorHang.scheme, MAJOR);

      wb.setThemeScriptFont(MINOR, 'Hang', 'Test Face');
      const saved = wb.save();
      const reloaded = Module.Workbook.load(saved);
      try {
        assert.equal(face(reloaded, MINOR, 'Hang').typeface, 'Test Face');
        assert.equal(face(reloaded, MINOR, 'Jpan').typeface, '游ゴシック');
        assert.deepEqual(face(reloaded, MAJOR, 'Hang'), majorHang);
      } finally {
        reloaded.delete();
      }
    });
  });

  test('listThemeScriptFonts lists the scheme entries and reflects a set face', () => {
    withWorkbook(Module, (wb) => {
      const major = wb.listThemeScriptFonts(MAJOR);
      assert.ok(major.length > 0);
      assert.ok(major.every((entry) => entry.scheme === MAJOR && typeof entry.script === 'string'));
      assert.deepEqual(
        major.find((entry) => entry.script === 'Jpan'),
        wb.getThemeScriptFont(MAJOR, 'Jpan'),
      );

      wb.setThemeScriptFont(MAJOR, 'Jpan', 'Test Face');
      const after = wb.listThemeScriptFonts(MAJOR);
      assert.equal(after.find((entry) => entry.script === 'Jpan').typeface, 'Test Face');
      assert.equal(after.length, major.length);
      assert.equal(wb.listThemeScriptFonts(MINOR).find((entry) => entry.script === 'Jpan').typeface, '游ゴシック');

      expectFailure(() => wb.listThemeScriptFonts(0));
    });
  });

  test('listProfileIds lists every profile id and each one builds a workbook', () => {
    const ids = Module.listProfileIds();
    assert.equal(ids.length, 28);
    assert.equal(new Set(ids).size, ids.length);
    assert.ok(ids.includes('win-365-en_US'));
    assert.ok(ids.includes('mac-365-ja_JP'));
    for (const id of ids) {
      const wb = Module.Workbook.createForProfile(id);
      try {
        assert.equal(wb.getProfileId(), id);
      } finally {
        wb.delete();
      }
    }
  });

  test('setThemeScriptFont rejects an invalid script or empty typeface', () => {
    withWorkbook(Module, (wb) => {
      expectFailure(() => wb.setThemeScriptFont(MAJOR, 'Jp', 'X'));
      expectFailure(() => wb.setThemeScriptFont(MAJOR, 'Jpan', ''));
      expectFailure(() => wb.setThemeScriptFont(0, 'Jpan', 'X'));
    });
  });
}
