import test from 'node:test';
import { assert, expectFailure, getModule, readAll } from './smoke_support.mjs';

// ---- Outline, view flags, font toggles, gradient fills, custom icons, ------
// ---- calcPr accessors and defined-name attributes ---------------------------

const release = (wb) => wb.delete();

// A rejected argument surfaces as a thrown TypeError or as a failed status, per surface.
const rejected = (call) => {
  try {
    const result = call();
    return result?.status?.ok === false;
  } catch (error) {
    return error instanceof TypeError || error instanceof RangeError;
  }
};

const roundTrip = (mod, wb) => {
  const saved = wb.save();
  return mod.Workbook.load(saved);
};

test('collapsed reads back on rows and columns and survives save/load', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setRowOutline(0, 1, 1, 2);
  wb.setRowCollapsed(0, 3, 3, true);
  wb.setColumnOutline(0, 2, 3, 3);
  wb.setColumnCollapsed(0, 4, 4, true);
  const collapsedRows = (b) =>
    b
      .listRowOverrides(0)
      .filter((r) => r.collapsed)
      .map((r) => r.row);
  assert.deepEqual(collapsedRows(wb), [3]);
  const loaded = roundTrip(mod, wb);
  assert.deepEqual(collapsedRows(loaded), [3]);
  assert.equal(loaded.listColumnOverrides(0).filter((c) => c.collapsed).length, 1);
  // Clearing the flag drops an override that carries nothing else.
  loaded.setRowCollapsed(0, 3, 3, false);
  assert.deepEqual(collapsedRows(loaded), []);
  expectFailure(() => wb.setRowCollapsed(9, 0, 0, true));
  expectFailure(() => wb.setColumnCollapsed(0, 3, 2, true));
  release(loaded);
  release(wb);
});

test('showFormulas and showOutlineSymbols read back and round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const initial = wb.getView(0);
  assert.equal(initial.showFormulas, 0);
  assert.equal(initial.showOutlineSymbols, 1);
  wb.setShowFormulas(0, true);
  wb.setShowOutlineSymbols(0, false);
  const loaded = roundTrip(mod, wb);
  const view = loaded.getView(0);
  assert.equal(view.showFormulas, 1);
  assert.equal(view.showOutlineSymbols, 0);
  release(loaded);
  release(wb);
});

test('font toggles round-trip and dedup', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const record = {
    name: 'Arial',
    size: 12,
    hasOutline: true,
    outline: true,
    hasShadow: true,
    shadow: false,
    hasCondense: true,
    condense: true,
    hasExtend: true,
    extend: true,
  };
  const added = wb.addFont(record);
  assert.equal(wb.addFont(record), added);
  const font = wb.getFont(added);
  assert.deepEqual(
    [font.hasOutline, font.outline, font.hasShadow, font.shadow, font.condense, font.extend],
    [true, true, true, false, true, true],
  );
  const loaded = roundTrip(mod, wb);
  assert.equal(loaded.getFont(added).condense, true);
  release(loaded);
  release(wb);
});

test('gradient fills add, read, dedup and round-trip in cells and dxfs', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const rgb = (argb) => ({ kind: 1, rgb: argb, theme: 0, tint: 0, indexed: 0 });
  const gradient = {
    type: 1,
    degree: 0,
    left: 0.5,
    right: 0.5,
    top: 0.25,
    bottom: 0.75,
    stops: [
      { position: 0, colorArgb: 0xffff0000, color: rgb(0xffff0000) },
      { position: 1, colorArgb: 0, color: { kind: 2, rgb: 0, theme: 4, tint: 0.5, indexed: 0 } },
    ],
  };
  const added = wb.addFill({ gradient });
  assert.equal(wb.addFill({ gradient }), added);
  const fill = wb.getFill(added);
  assert.equal(fill.gradient.type, 1);
  assert.equal(fill.gradient.top, 0.25);
  assert.equal(fill.gradient.stops.length, 2);
  assert.equal(fill.gradient.stops[0].colorArgb, 0xffff0000);
  assert.equal(fill.gradient.stops[1].color.theme, 4);
  assert.equal(wb.getFill(0).gradient, undefined);

  const dxf = wb.addDxf({ fill: { gradient: { ...gradient, type: 0, degree: 45 } } });
  const loaded = roundTrip(mod, wb);
  assert.equal(loaded.getFill(added).gradient.stops.length, 2);
  assert.equal(loaded.getDxf(dxf).fill.gradient.degree, 45);
  assert.ok(rejected(() => wb.addFill({ gradient: { stops: [{ position: 'x' }] } })));
  release(loaded);
  release(wb);
});

test('custom icon sets round-trip through the conditional-format API', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const icons = [
    { iconSet: '3Triangles', iconId: 0 },
    { iconSet: 'NoIcons', iconId: 0 },
    { iconSet: '5Boxes', iconId: 4 },
  ];
  const rule = (list) => ({
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
    type: 4,
    iconSet: {
      name: 0,
      thresholds: [
        { type: 1, value: '33', gte: true },
        { type: 1, value: '67', gte: true },
      ],
      icons: list,
    },
  });
  wb.addCfRule(0, rule(icons));
  const loaded = roundTrip(mod, wb);
  const back = readAll(
    () => loaded.getCfRuleCount(0),
    (i) => loaded.getCfRule(0, i),
  ).find((r) => r.iconSet);
  assert.deepEqual(back.iconSet.icons, icons);
  expectFailure(() => wb.addCfRule(0, rule(icons.slice(0, 2))));
  expectFailure(() => wb.addCfRule(0, rule([icons[0], { iconSet: '7Dwarfs', iconId: 0 }, icons[2]])));
  release(loaded);
  release(wb);
});

test('calcPr accessors and the date system', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getFullPrecision(), true);
  assert.equal(wb.getCalcOnSave(), true);
  assert.equal(wb.getRefMode(), 0);
  assert.equal(wb.getDate1904(), false);
  wb.setFullPrecision(false);
  wb.setCalcOnSave(false);
  wb.setRefMode(1);
  wb.setDate1904(true);
  expectFailure(() => wb.setRefMode(9));
  const loaded = roundTrip(mod, wb);
  assert.equal(loaded.getFullPrecision(), false);
  assert.equal(loaded.getCalcOnSave(), false);
  assert.equal(loaded.getRefMode(), 1);
  assert.equal(loaded.getDate1904(), true);
  release(loaded);
  release(wb);
});

test("the file name spells and binds the workbook's own names", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getFileName(), '');
  wb.setDefinedName('LN', '10');
  wb.setFileName('M5.xlsx');
  assert.equal(wb.getFileName(), 'M5.xlsx');
  wb.setFormula(0, 0, 0, '=M5.xlsx!LN+1');
  assert.equal(wb.getFormula(0, 0, 0), '=M5.xlsx!LN+1');
  assert.equal(wb.getCellAt(0, 0).formula, '=M5.xlsx!LN+1');
  wb.setFileName('M5 b-1.xlsx');
  assert.equal(wb.getFormula(0, 0, 0), "='M5 b-1.xlsx'!LN+1");
  expectFailure(() => wb.setFileName('Book1'));
  assert.equal(wb.getFileName(), 'M5 b-1.xlsx');
  wb.setFileName('');
  assert.equal(wb.getFormula(0, 0, 0), '=[0]!LN+1');
  release(wb);
});

test('defined-name comment and hidden read back, set and survive a replacement', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setDefinedNameScoped('Rate', 'Sheet1!$A$1', -1);
  assert.equal(wb.listDefinedNames()[0].comment, '');
  assert.equal(wb.listDefinedNames()[0].hidden, false);
  wb.setDefinedNameAttrs('rate', -1, { comment: 'Tax rate', hidden: true });
  wb.setDefinedNameScoped('Rate', 'Sheet1!$B$2', -1);
  wb.setDefinedNameAttrs('Rate', -1, {});
  const entry = wb.listDefinedNames()[0];
  assert.deepEqual([entry.formula, entry.comment, entry.hidden], ['Sheet1!$B$2', 'Tax rate', true]);
  wb.setDefinedNameAttrs('Rate', -1, { hidden: false });
  expectFailure(() => wb.setDefinedNameAttrs('Nope', -1, { comment: 'x' }));
  const loaded = roundTrip(mod, wb);
  assert.deepEqual([loaded.listDefinedNames()[0].comment, loaded.listDefinedNames()[0].hidden], ['Tax rate', false]);
  release(loaded);
  release(wb);
});

test('sheet format defaults expose customHeight and zeroHeight and keep omitted flags', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const initial = wb.getFormatDefaults(0);
  assert.equal(initial.customHeight, false);
  assert.equal(initial.zeroHeight, false);
  wb.setFormatDefaults(0, { defaultRowHeight: 18, customHeight: true });
  assert.deepEqual([wb.getFormatDefaults(0).customHeight, wb.getFormatDefaults(0).zeroHeight], [true, false]);
  // An omitted flag keeps the stored one.
  wb.setFormatDefaults(0, { defaultRowHeight: 18, zeroHeight: true });
  const loaded = roundTrip(mod, wb);
  const defaults = loaded.getFormatDefaults(0);
  assert.deepEqual([defaults.customHeight, defaults.zeroHeight, defaults.defaultRowHeight], [true, true, 18]);
  loaded.setFormatDefaults(0, { defaultRowHeight: 18, customHeight: false, zeroHeight: false });
  assert.deepEqual([loaded.getFormatDefaults(0).customHeight, loaded.getFormatDefaults(0).zeroHeight], [false, false]);
  release(loaded);
  release(wb);
});
