import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

const DEFAULT_ACCENT1 = 0xff4472c4;

test('getTheme reads the default theme of a fresh workbook', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const theme = wb.getTheme();
    assert.ok(theme.status.ok, JSON.stringify(theme.status));
    assert.ok([0, 1].includes(theme.source), `source ${theme.source}`);
    assert.equal(theme.colors.length, 12);
    assert.equal(typeof theme.fonts.majorLatin, 'string');
    assert.equal(typeof theme.fonts.minorEastAsian, 'string');
    assert.ok(theme.fonts.minorLatin.length > 0);
    assert.equal(theme.colors[4] >>> 0, DEFAULT_ACCENT1 >>> 0);
  } finally {
    wb.dispose();
  }
});

test('setThemeColors and setThemeFonts round-trip through getTheme', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const colors = Array.from({ length: 12 }, (_, i) => (0xff000000 | (0x112233 + i * 0x010101)) >>> 0);
    assert.ok(wb.setThemeColors(colors).ok);
    assert.ok(
      wb.setThemeFonts({
        majorLatin: 'Aptos Display',
        majorEastAsian: 'Yu Gothic',
        minorLatin: 'Aptos',
        minorEastAsian: 'Yu Mincho',
      }).ok,
    );
    const theme = wb.getTheme();
    assert.ok(theme.status.ok);
    assert.equal(theme.source, 0);
    assert.deepEqual(theme.colors, colors);
    assert.equal(theme.fonts.majorLatin, 'Aptos Display');
    assert.equal(theme.fonts.majorEastAsian, 'Yu Gothic');
    assert.equal(theme.fonts.minorLatin, 'Aptos');
    assert.equal(theme.fonts.minorEastAsian, 'Yu Mincho');

    const bad = wb.setThemeColors([1, 2, 3]);
    assert.equal(bad.ok, false);
  } finally {
    wb.dispose();
  }
});

test('resetTheme returns an edited theme to the default', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(wb.resetTheme().ok);
    assert.equal(wb.getTheme().source, 1);
    const colors = Array.from({ length: 12 }, (_, i) => (0xff000000 | (0x112233 + i * 0x010101)) >>> 0);
    assert.ok(wb.setThemeColors(colors).ok);
    assert.equal(wb.getTheme().source, 0);
    assert.ok(wb.resetTheme().ok);
    const theme = wb.getTheme();
    assert.equal(theme.source, 1);
    assert.equal(theme.colors[4] >>> 0, DEFAULT_ACCENT1 >>> 0);

    const saved = wb.save();
    assert.ok(saved.status.ok);
    assert.equal(Buffer.from(saved.bytes).includes('xl/theme/theme1.xml'), false);
  } finally {
    wb.dispose();
  }
});

test('resolveColor resolves literal, theme and automatic colours by context', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const rgb = wb.resolveColor({ kind: 1, rgb: 0xff112233 }, mod.ColorContext.Font);
    assert.ok(rgb.status.ok);
    assert.equal(rgb.argb >>> 0, 0xff112233);
    assert.equal(rgb.resolution, mod.ColorResolution.Exact);

    const theme = wb.resolveColor({ kind: 2, theme: 4, tint: 0 }, mod.ColorContext.FillForeground);
    assert.ok(theme.status.ok);
    assert.ok([mod.ColorResolution.Exact, mod.ColorResolution.DefaultTheme].includes(theme.resolution));

    const fontAuto = wb.resolveColor({ kind: 4 }, mod.ColorContext.Font);
    assert.equal(fontAuto.argb >>> 0, 0xff000000);
    assert.equal(fontAuto.resolution, mod.ColorResolution.AutoContext);
    const fillAuto = wb.resolveColor({ kind: 4 }, mod.ColorContext.FillBackground);
    assert.equal(fillAuto.argb >>> 0, 0xffffffff);

    const bad = wb.resolveColor({ kind: 1, rgb: 1 }, 99);
    assert.equal(bad.status.ok, false);
  } finally {
    wb.dispose();
  }
});

test('getEffectiveStyle reports the cell xf with resolved colours', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const font = wb.addFont({ name: 'Arial', size: 11, colorArgb: 0xffff0000 });
    assert.ok(font.status.ok);
    const xf = wb.addXf({ fontIndex: font.index, fillIndex: 0, borderIndex: 0, numFmtId: 0, applyFont: true });
    assert.ok(xf.status.ok);
    assert.ok(wb.setCellXfIndex(0, 2, 3, xf.index).ok);

    const eff = wb.getEffectiveStyle(0, 2, 3);
    assert.ok(eff.status.ok, JSON.stringify(eff.status));
    assert.equal(eff.xfIndex, xf.index);
    assert.equal(eff.source, 0);
    assert.equal(eff.fontIndex, font.index);
    assert.equal(eff.font.argb >>> 0, 0xffff0000);
    assert.equal(typeof eff.borders.diagonal.argb, 'number');
    assert.equal(eff.locked, true);
    assert.equal(typeof eff.numFmtCode, 'string');

    const plain = wb.getEffectiveStyle(0, 50, 50);
    assert.ok(plain.status.ok);
    assert.equal(plain.xfIndex, 0);
    assert.equal(plain.source, 3);

    const failed = wb.getEffectiveStyle(99, 0, 0);
    assert.equal(failed.status.ok, false);
    assert.equal(failed.numFmtCode, '');
  } finally {
    wb.dispose();
  }
});

test('setCellStyle round-trips a record and validates its fields', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const before = wb.cellStyleCount().value;
    const rec = { name: 'Probe Style', xfId: 0, builtinId: 53, iLevel: 0, hidden: true, customBuiltin: true };
    assert.ok(wb.setCellStyle(rec).ok);
    assert.equal(wb.cellStyleCount().value, before + 1);
    const read = wb.getCellStyle(before);
    assert.equal(read.name, 'Probe Style');
    assert.equal(read.builtinId, 53);
    assert.equal(read.hidden, true);
    assert.equal(read.customBuiltin, true);

    assert.ok(wb.setCellStyle({ name: 'Probe Style' }).ok);
    assert.equal(wb.cellStyleCount().value, before + 1);
    assert.equal(wb.getCellStyle(before).builtinId, 0xffffffff);

    assert.equal(wb.setCellStyle({ name: 'Bad Level', iLevel: 3 }).ok, false);
    assert.equal(wb.setCellStyle({ name: '' }).ok, false);
  } finally {
    wb.dispose();
  }
});

test('removeCellStyle rejects the Normal style and an unknown name', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const normal = wb.getCellStyle(0);
    assert.ok(normal.status.ok);
    assert.equal(wb.removeCellStyle(normal.name).ok, false);
    assert.equal(wb.removeCellStyle('No Such Style').ok, false);
    if (wb.cellStyleCount().value > 1) {
      const other = wb.getCellStyle(1);
      const before = wb.cellStyleCount().value;
      assert.ok(wb.removeCellStyle(other.name).ok);
      assert.equal(wb.cellStyleCount().value, before - 1);
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
