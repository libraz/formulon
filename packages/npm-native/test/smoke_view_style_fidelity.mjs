import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

// ---- Outline, view flags, font toggles, gradient fills, custom icons, ------
// ---- calcPr accessors and defined-name attributes ---------------------------

const release = (wb) => wb.dispose();

// A rejected argument surfaces as a binding TypeError / RangeError or as a host FormulonError.
const rejected = (call) => {
  try {
    call();
    return false;
  } catch (error) {
    return error instanceof TypeError || error instanceof RangeError || isFormulonError(error);
  }
};

const roundTrip = (mod, wb) => {
  const saved = wb.save();
  return mod.Workbook.load(saved);
};

test('collapsed reads back on rows and columns and survives save/load', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setRowOutline(1, 1, 2);
  wb.sheet(0).setRowCollapsed(3, 3, true);
  wb.sheet(0).setColumnOutline(2, 3, 3);
  wb.sheet(0).setColumnCollapsed(4, 4, true);
  const collapsedRows = (b) =>
    b
      .sheet(0)
      .listRowOverrides()
      .filter((r) => r.collapsed)
      .map((r) => r.row);
  assert.deepEqual(collapsedRows(wb), [3]);
  const loaded = roundTrip(mod, wb);
  assert.deepEqual(collapsedRows(loaded), [3]);
  assert.equal(
    loaded
      .sheet(0)
      .listColumnOverrides()
      .filter((c) => c.collapsed).length,
    1,
  );
  // Clearing the flag drops an override that carries nothing else.
  loaded.sheet(0).setRowCollapsed(3, 3, false);
  assert.deepEqual(collapsedRows(loaded), []);
  assert.throws(() => wb.sheet(9).setRowCollapsed(0, 0, true), isFormulonError);
  assert.throws(() => wb.sheet(0).setColumnCollapsed(3, 2, true), isFormulonError);
  release(loaded);
  release(wb);
});

test('showFormulas and showOutlineSymbols read back and round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const initial = wb.sheet(0).getView();
  assert.equal(initial.showFormulas, false);
  assert.equal(initial.showOutlineSymbols, true);
  wb.sheet(0).setView({ showFormulas: true });
  wb.sheet(0).setView({ showOutlineSymbols: false });
  const loaded = roundTrip(mod, wb);
  const view = loaded.sheet(0).getView();
  assert.equal(view.showFormulas, true);
  assert.equal(view.showOutlineSymbols, false);
  release(loaded);
  release(wb);
});

test('font toggles round-trip and dedup', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const record = {
    name: 'Arial',
    size: 12,
    outline: true,
    shadow: false,
    condense: true,
    extend: true,
  };
  const added = wb.addFont(record);
  assert.equal(wb.addFont(record), added);
  const font = wb.getFont(added);
  assert.deepEqual([font.outline, font.shadow, font.condense, font.extend], [true, false, true, true]);
  // An element the record leaves out reads back as null.
  assert.equal(font.bold, null);
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
  wb.sheet(0).addCfRule(rule(icons));
  const loaded = roundTrip(mod, wb);
  const back = loaded
    .sheet(0)
    .listCfRules()
    .find((r) => r.iconSet);
  assert.deepEqual(back.iconSet.icons, icons);
  assert.throws(() => wb.sheet(0).addCfRule(rule(icons.slice(0, 2))), isFormulonError);
  assert.throws(
    () => wb.sheet(0).addCfRule(rule([icons[0], { iconSet: '7Dwarfs', iconId: 0 }, icons[2]])),
    isFormulonError,
  );
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
  assert.throws(() => wb.setRefMode(9), isFormulonError);
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
  assert.equal(wb.getFileName(), null);
  wb.setDefinedName({ name: 'LN', formula: '10' });
  wb.setFileName('M5.xlsx');
  assert.equal(wb.getFileName(), 'M5.xlsx');
  wb.sheet(0).setFormula(0, 0, '=M5.xlsx!LN+1');
  assert.equal(wb.sheet(0).getFormula(0, 0), '=M5.xlsx!LN+1');
  assert.equal(wb.sheet(0).listCells()[0].formula, '=M5.xlsx!LN+1');
  wb.setFileName('M5 b-1.xlsx');
  assert.equal(wb.sheet(0).getFormula(0, 0), "='M5 b-1.xlsx'!LN+1");
  assert.throws(() => wb.setFileName('Book1'), isFormulonError);
  assert.equal(wb.getFileName(), 'M5 b-1.xlsx');
  wb.setFileName(null);
  assert.equal(wb.getFileName(), null);
  assert.equal(wb.sheet(0).getFormula(0, 0), '=[0]!LN+1');
  release(wb);
});

test('defined-name comment and hidden read back, and a set replaces the whole name', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$A$1' });
  assert.equal(wb.listDefinedNames()[0].comment, null);
  assert.equal(wb.listDefinedNames()[0].hidden, false);
  wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$A$1', comment: 'Tax rate', hidden: true });
  let entry = wb.listDefinedNames()[0];
  assert.deepEqual([entry.formula, entry.comment, entry.hidden], ['=Sheet1!$A$1', 'Tax rate', true]);
  // Omitted comment / hidden return to their defaults.
  wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$B$2' });
  entry = wb.listDefinedNames()[0];
  assert.deepEqual([entry.formula, entry.comment, entry.hidden], ['=Sheet1!$B$2', null, false]);
  wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$B$2', comment: 'Tax rate' });
  const loaded = roundTrip(mod, wb);
  assert.deepEqual([loaded.listDefinedNames()[0].comment, loaded.listDefinedNames()[0].hidden], ['Tax rate', false]);
  release(loaded);
  release(wb);
});

test('sheet format defaults expose customHeight and zeroHeight and keep omitted flags', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const initial = wb.sheet(0).getFormatDefaults();
  assert.equal(initial.customHeight, false);
  assert.equal(initial.zeroHeight, false);
  wb.sheet(0).setFormatDefaults({ defaultRowHeight: 18, customHeight: true });
  assert.deepEqual(
    [wb.sheet(0).getFormatDefaults().customHeight, wb.sheet(0).getFormatDefaults().zeroHeight],
    [true, false],
  );
  // An omitted flag keeps the stored one.
  wb.sheet(0).setFormatDefaults({ defaultRowHeight: 18, zeroHeight: true });
  const loaded = roundTrip(mod, wb);
  const defaults = loaded.sheet(0).getFormatDefaults();
  assert.deepEqual([defaults.customHeight, defaults.zeroHeight, defaults.defaultRowHeight], [true, true, 18]);
  loaded.sheet(0).setFormatDefaults({ defaultRowHeight: 18, customHeight: false, zeroHeight: false });
  assert.deepEqual(
    [loaded.sheet(0).getFormatDefaults().customHeight, loaded.sheet(0).getFormatDefaults().zeroHeight],
    [false, false],
  );
  release(loaded);
  release(wb);
});
