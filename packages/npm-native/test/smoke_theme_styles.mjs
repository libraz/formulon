import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

const DEFAULT_ACCENT1 = 0xff156082;
const argbHex = (value) => `#${(value >>> 0).toString(16).toUpperCase().padStart(8, '0')}`;

test('getThemeColors / getThemeFonts read the default theme of a fresh workbook', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const theme = wb.getThemeColors();
    assert.ok([0, 1].includes(theme.source), `source ${theme.source}`);
    assert.equal(theme.colors.length, 12);
    const fonts = wb.getThemeFonts();
    assert.equal(typeof fonts.majorLatin, 'string');
    assert.equal(typeof fonts.minorEastAsian, 'string');
    assert.ok(fonts.minorLatin.length > 0);
    assert.equal(theme.colors[4] >>> 0, DEFAULT_ACCENT1 >>> 0);
  } finally {
    wb.dispose();
  }
});

test('setThemeColors and setThemeFonts round-trip through their getters', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const colors = Array.from({ length: 12 }, (_, i) => (0xff000000 | (0x112233 + i * 0x010101)) >>> 0);
    wb.setThemeColors({ colors: colors.map(argbHex) });
    wb.setThemeFonts({
      majorLatin: 'Aptos Display',
      majorEastAsian: 'Yu Gothic',
      minorLatin: 'Aptos',
      minorEastAsian: 'Yu Mincho',
    });
    const theme = wb.getThemeColors();
    assert.equal(theme.source, 0);
    assert.deepEqual(theme.colors, colors);
    assert.deepEqual(wb.getThemeFonts(), {
      majorLatin: 'Aptos Display',
      majorEastAsian: 'Yu Gothic',
      minorLatin: 'Aptos',
      minorEastAsian: 'Yu Mincho',
    });

    assert.throws(() => wb.setThemeColors({ colors: [1, 2, 3] }), TypeError);
  } finally {
    wb.dispose();
  }
});

test('resetTheme returns an edited theme to the default', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    wb.resetTheme();
    assert.equal(wb.getThemeColors().source, 1);
    const colors = Array.from({ length: 12 }, (_, i) => (0xff000000 | (0x112233 + i * 0x010101)) >>> 0);
    wb.setThemeColors({ colors: colors.map(argbHex) });
    assert.equal(wb.getThemeColors().source, 0);
    wb.resetTheme();
    const theme = wb.getThemeColors();
    assert.equal(theme.source, 1);
    assert.equal(theme.colors[4] >>> 0, DEFAULT_ACCENT1 >>> 0);

    const saved = wb.save();
    assert.equal(Buffer.from(saved).includes('xl/theme/theme1.xml'), false);
  } finally {
    wb.dispose();
  }
});

test('resolveColor resolves literal, theme and automatic colours by context', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const rgb = wb.resolveColor({ kind: 1, argb: 0xff112233 }, mod.ColorContext.Font);
    assert.equal(rgb.argb >>> 0, 0xff112233);
    assert.equal(rgb.resolution, mod.ColorResolution.Exact);

    const theme = wb.resolveColor({ kind: 2, theme: 4, tint: 0 }, mod.ColorContext.FillForeground);
    assert.ok([mod.ColorResolution.Exact, mod.ColorResolution.DefaultTheme].includes(theme.resolution));

    const fontAuto = wb.resolveColor({ kind: 4 }, mod.ColorContext.Font);
    assert.equal(fontAuto.argb >>> 0, 0xff000000);
    assert.equal(fontAuto.resolution, mod.ColorResolution.AutoContext);
    const fillAuto = wb.resolveColor({ kind: 4 }, mod.ColorContext.FillBackground);
    assert.equal(fillAuto.argb >>> 0, 0xffffffff);

    assert.throws(() => wb.resolveColor({ kind: 1, argb: 1 }, 99), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('getEffectiveStyle reports the cell xf with resolved colours', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const font = wb.ooxml.addFont({ name: 'Arial', sizePt: 11, color: '#FFFF0000' });
    const xf = wb.ooxml.addCellXf({ fontIndex: font, fillIndex: 0, borderIndex: 0, numFmtId: 0, applyFont: true });
    wb.sheet(0).ooxml.setCellXfIndex(2, 3, xf);

    const eff = wb.sheet(0).ooxml.getEffectiveStyle(2, 3);
    assert.equal(eff.xfIndex, xf);
    assert.equal(eff.source, 0);
    assert.equal(eff.fontIndex, font);
    assert.equal(eff.fontColor.argb >>> 0, 0xffff0000);
    assert.equal(typeof eff.borders.diagonal.argb, 'number');
    assert.equal(eff.protection.locked, true);
    assert.equal(eff.protection.hidden, false);
    assert.equal('locked' in eff, false);
    assert.equal(typeof eff.numberFormat, 'string');

    const plain = wb.sheet(0).ooxml.getEffectiveStyle(50, 50);
    assert.equal(plain.xfIndex, 0);
    assert.equal(plain.source, 3);

    assert.throws(() => wb.sheet(99).ooxml.getEffectiveStyle(0, 0), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('resetCellXfIndex restores row / column inheritance', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const xfCol = wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 14 });
    const xfCell = wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 10 });
    wb.sheet(0).ooxml.setColumnXfIndex(0, 0, xfCol);
    wb.sheet(0).ooxml.setCellXfIndex(0, 0, xfCell);
    let e = wb.sheet(0).ooxml.getEffectiveStyle(0, 0);
    assert.equal(e.source, 0);
    assert.equal(e.xfIndex, xfCell);

    wb.sheet(0).ooxml.resetCellXfIndex(0, 0);
    e = wb.sheet(0).ooxml.getEffectiveStyle(0, 0);
    assert.equal(e.source, 2);
    assert.equal(e.xfIndex, xfCol);

    wb.sheet(0).ooxml.setCellXfIndex(0, 0, 0);
    e = wb.sheet(0).ooxml.getEffectiveStyle(0, 0);
    assert.equal(e.source, 0);
    assert.equal(e.xfIndex, 0);

    const before = wb.sheet(0).listCells().length;
    wb.sheet(0).ooxml.resetCellXfIndex(7, 7);
    assert.equal(wb.sheet(0).listCells().length, before);
    assert.throws(() => wb.sheet(99).ooxml.resetCellXfIndex(0, 0), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('attribute patch over a column keeps the row style (bases / remap / snapshot / restore)', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const rowXf = wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 14 });
    const xfA = wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 9 });
    const xfB = wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 10 });
    wb.sheet(0).ooxml.setRowXfIndex(2, 2, rowXf);
    assert.equal(wb.sheet(0).ooxml.getEffectiveStyle(2, 1).source, 1);
    assert.throws(() => wb.sheet(0).ooxml.setRowXfIndex(5, 4, rowXf), isFormulonError);

    const bases = wb.sheet(0).ooxml.getRangeXfBases({ firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 1 });
    assert.ok(bases.includes(0) && bases.includes(rowXf));
    // Row 2 of a one-column selection is a single cell; excluding it hides the row style.
    const excluded = wb
      .sheet(0)
      .ooxml.getRangeXfBasesExcluding({ firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 1 }, [2, 2, 9], [1, 1, 9]);
    assert.deepEqual(excluded, [0]);
    assert.throws(
      () =>
        wb.sheet(0).ooxml.getRangeXfBasesExcluding({ firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 1 }, [2], []),
      TypeError,
    );

    const snap = wb.sheet(0).ooxml.snapshotRangeXfs({ firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 1 });
    assert.ok(snap.length > 0);
    assert.throws(
      () =>
        wb.sheet(0).ooxml.remapRangeXfs({ firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 1 }, [0, rowXf], [xfA]),
      TypeError,
    );
    wb.sheet(0).ooxml.remapRangeXfs({ firstRow: 0, firstCol: 1, lastRow: 1048575, lastCol: 1 }, [0, rowXf], [xfA, xfB]);

    let e = wb.sheet(0).ooxml.getEffectiveStyle(2, 1);
    assert.equal(e.source, 0);
    assert.equal(e.xfIndex, xfB);
    e = wb.sheet(0).ooxml.getEffectiveStyle(2, 2);
    assert.equal(e.source, 1);
    assert.equal(e.xfIndex, rowXf);
    e = wb.sheet(0).ooxml.getEffectiveStyle(499, 1);
    assert.equal(e.source, 2);
    assert.equal(e.xfIndex, xfA);

    wb.sheet(0).ooxml.restoreRangeXfs(snap);
    e = wb.sheet(0).ooxml.getEffectiveStyle(2, 1);
    assert.equal(e.source, 1);
    assert.equal(e.xfIndex, rowXf);
    assert.equal(wb.sheet(0).ooxml.getEffectiveStyle(499, 1).source, 3);
    assert.throws(() => wb.sheet(0).ooxml.restoreRangeXfs(new Uint8Array([1, 2, 3])), isFormulonError);

    wb.sheet(0).ooxml.resetRowXfIndex(2, 2);
    assert.equal(wb.sheet(0).ooxml.getEffectiveStyle(2, 2).source, 3);
  } finally {
    wb.dispose();
  }
});

test('setCellStyle round-trips a record and validates its fields', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const before = wb.listCellStyles().length;
    const rec = { name: 'Probe Style', styleXfIndex: 0, builtinId: 53, iLevel: 0, hidden: true, customBuiltin: true };
    wb.setCellStyle(rec);
    assert.equal(wb.listCellStyles().length, before + 1);
    const read = wb.listCellStyles()[before];
    assert.equal(read.name, 'Probe Style');
    assert.equal(read.builtinId, 53);
    assert.equal(read.hidden, true);
    assert.equal(read.customBuiltin, true);

    wb.setCellStyle({ name: 'Probe Style' });
    assert.equal(wb.listCellStyles().length, before + 1);
    assert.equal(wb.listCellStyles()[before].builtinId, 0xffffffff);

    assert.throws(() => wb.setCellStyle({ name: 'Bad Level', iLevel: 3 }), isFormulonError);
    assert.throws(() => wb.setCellStyle({ name: '' }), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('removeCellStyle rejects the Normal style and an unknown name', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const styles = wb.listCellStyles();
    const normal = styles[0];
    assert.throws(() => wb.removeCellStyle(normal.name), isFormulonError);
    assert.throws(() => wb.removeCellStyle('No Such Style'), isFormulonError);
    if (wb.listCellStyles().length > 1) {
      const other = wb.listCellStyles()[1];
      const before = wb.listCellStyles().length;
      wb.removeCellStyle(other.name);
      assert.equal(wb.listCellStyles().length, before - 1);
    }
  } finally {
    wb.dispose();
  }
});

test('theme and colour enums carry the C ABI ordinals', async () => {
  const mod = await getModule();
  assert.deepEqual({ ...mod.ColorContext }, { Font: 0, FillForeground: 1, FillBackground: 2, Border: 3 });
  assert.deepEqual(
    { ...mod.ColorResolution },
    { Exact: 0, DefaultTheme: 1, IndexOutOfRange: 2, ThemeUnparseable: 3, AutoContext: 4 },
  );
  assert.deepEqual({ ...mod.ThemeSource }, { Part: 0, Default: 1, Unparseable: 2 });
  assert.deepEqual({ ...mod.EffectiveStyleSource }, { Cell: 0, Row: 1, Column: 2, Default: 3 });
});

test('createForProfile builds the locale defaults and reports a bad id as an invalid wrapper', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createForProfile('mac-365-ja_JP');
  try {
    assert.equal(wb.disposed, false);
    assert.equal(wb.getProfileId(), 'mac-365-ja_JP');
    const font = wb.ooxml.getFont(0);
    assert.equal(font.name, '游ゴシック');
    assert.equal(font.sizePt, 12);
  } finally {
    wb.dispose();
  }
  assert.throws(() => mod.Workbook.createForProfile('no-such-profile'), isFormulonError);
});

test('getThemeScriptFont reads the default script faces and setThemeScriptFont round-trips', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const face = (book, major, script) => book.getThemeScriptFont({ major, script })?.typeface;
    assert.equal(face(wb, false, 'Jpan'), '游ゴシック');
    assert.equal(face(wb, false, 'Hang'), '맑은 고딕');
    const majorJpan = face(wb, true, 'Jpan');

    wb.setThemeScriptFont({ major: false, script: 'Hang', typeface: 'Test Face' });
    const saved = wb.save();
    const reloaded = mod.Workbook.load(saved);
    try {
      assert.equal(face(reloaded, false, 'Hang'), 'Test Face');
      assert.equal(face(reloaded, false, 'Jpan'), '游ゴシック');
      assert.equal(face(reloaded, true, 'Jpan'), majorJpan);
    } finally {
      reloaded.dispose();
    }
  } finally {
    wb.dispose();
  }
});

test('setThemeScriptFont rejects an invalid script or empty typeface', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.throws(() => wb.setThemeScriptFont({ major: true, script: 'Jp', typeface: 'X' }), isFormulonError);
    assert.throws(() => wb.setThemeScriptFont({ major: true, script: 'Jpan', typeface: '' }), isFormulonError);
  } finally {
    wb.dispose();
  }
});
