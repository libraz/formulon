import test from 'node:test';
import { assert, expectFailure, getModule } from './smoke_support.mjs';

// ---- Outline, view flags, font toggles, gradient fills, custom icons, ------
// ---- calcPr accessors and defined-name attributes ---------------------------

const release = (wb) => wb.dispose();

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
  expectFailure(() => wb.sheet(9).setRowCollapsed(0, 0, true));
  expectFailure(() => wb.sheet(0).setColumnCollapsed(3, 2, true));
  release(loaded);
  release(wb);
});

test('showFormulas and showOutlineSymbols read back and round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const initial = wb.sheet(0).getView();
  assert.equal(initial.showFormulas, false);
  assert.equal(initial.showOutlineSymbols, true);
  wb.sheet(0).updateView({ showFormulas: true, showOutlineSymbols: false });
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
    sizePt: 12,
    outline: true,
    shadow: false,
    condense: true,
    extend: true,
  };
  const added = wb.ooxml.addFont(record);
  assert.equal(wb.ooxml.addFont(record), added);
  const font = wb.ooxml.getFont(added);
  assert.deepEqual([font.outline, font.shadow, font.condense, font.extend], [true, false, true, true]);
  // An element the record leaves out reads back as null.
  assert.equal(font.bold, null);
  const loaded = roundTrip(mod, wb);
  assert.equal(loaded.ooxml.getFont(added).condense, true);
  release(loaded);
  release(wb);
});

test('gradient fills add, read, dedup and round-trip in cells and dxfs', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const rgb = (argb) => ({ kind: 1, argb: argb, theme: 0, tint: 0, indexed: 0 });
  const gradient = {
    type: 1,
    angleDeg: 0,
    left: 0.5,
    right: 0.5,
    top: 0.25,
    bottom: 0.75,
    stops: [
      { position: 0, color: rgb(0xffff0000) },
      { position: 1, color: { kind: 2, argb: 0, theme: 4, tint: 0.5, indexed: 0 } },
    ],
  };
  const added = wb.ooxml.addFill({ gradient });
  assert.equal(wb.ooxml.addFill({ gradient }), added);
  const fill = wb.ooxml.getFill(added);
  assert.equal(fill.gradient.type, 1);
  assert.equal(fill.gradient.top, 0.25);
  assert.equal(fill.gradient.stops.length, 2);
  assert.equal(fill.gradient.stops[0].color.argb, 0xffff0000);
  assert.equal(fill.gradient.stops[1].color.theme, 4);
  assert.equal(wb.ooxml.getFill(0).gradient, null);

  const dxf = wb.ooxml.addDxf({ fill: { gradient: { ...gradient, type: 0, angleDeg: 45 } } });
  wb.sheet(0).addCfRule({ type: mod.CfRuleType.Expression, ranges: 'A1', formula1: 'TRUE', dxfIndex: dxf });
  const loaded = roundTrip(mod, wb);
  assert.equal(loaded.ooxml.getFill(added).gradient.stops.length, 2);
  assert.equal(loaded.ooxml.getDxf(dxf).fill.gradient.angleDeg, 45);
  assert.ok(rejected(() => wb.ooxml.addFill({ gradient: { stops: [{ position: 'x' }] } })));
  release(loaded);
  release(wb);
});

test('a dxf carries its number format as a format code and reads null without one', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const dxf = wb.ooxml.addDxf({ numberFormat: '0.00%' });
  const read = wb.ooxml.getDxf(dxf);
  assert.equal(read.numberFormat, '0.00%');
  for (const key of ['numFmt', 'numFmtId', 'numFmtCode']) assert.equal(key in read, false, key);
  assert.equal(wb.ooxml.addDxf(read), dxf);
  const bare = wb.ooxml.addDxf({ font: { bold: true } });
  assert.equal(wb.ooxml.getDxf(bare).numberFormat, null);
  wb.sheet(0).addCfRule({ type: mod.CfRuleType.Expression, ranges: 'A1', formula1: 'TRUE', dxfIndex: dxf });
  wb.sheet(0).addCfRule({ type: mod.CfRuleType.Expression, ranges: 'A2', formula1: 'TRUE', dxfIndex: bare });
  const loaded = mod.Workbook.load(wb.save());
  assert.equal(loaded.ooxml.getDxf(dxf).numberFormat, '0.00%');
  loaded.dispose();
  wb.dispose();
});

test('getDefaultFont reads the workbook default font and follows setDefaultFont', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getDefaultFont().name, 'Calibri');
  assert.deepEqual(wb.getDefaultFont(), wb.ooxml.getFont(0));
  wb.ooxml.addFont({ name: 'Meiryo', sizePt: 11 });
  assert.equal(wb.getDefaultFont().name, 'Calibri');
  wb.setDefaultFont({ name: 'Meiryo', sizePt: 11, charset: 128 });
  assert.equal(wb.getDefaultFont().name, 'Meiryo');
  assert.equal(wb.getDefaultFont().charset, 128);
  assert.deepEqual(wb.getDefaultFont(), wb.ooxml.getFont(0));
  wb.dispose();
});

test('custom icon sets round-trip through the conditional-format API', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const icons = [
    { iconSet: mod.IconSetName.ThreeTriangles, iconIndex: 0 },
    { iconSet: mod.IconSetName.NoIcons, iconIndex: 0 },
    { iconSet: mod.IconSetName.FiveBoxes, iconIndex: 4 },
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
  expectFailure(() => wb.sheet(0).addCfRule(rule(icons.slice(0, 2))));
  expectFailure(() => wb.sheet(0).addCfRule(rule([icons[0], { iconSet: 99, iconIndex: 0 }, icons[2]])));
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

test("the location's file name spells and binds the workbook's own names", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getLocation(), null);
  wb.setDefinedName({ name: 'LN', formula: '10' });
  wb.setLocation({ fileName: 'M5.xlsx' });
  assert.deepEqual(wb.getLocation(), { directory: null, fileName: 'M5.xlsx' });
  wb.sheet(0).setFormula(0, 0, '=M5.xlsx!LN+1');
  assert.equal(wb.sheet(0).getFormula(0, 0), '=M5.xlsx!LN+1');
  assert.equal(wb.sheet(0).listCells()[0].formula, '=M5.xlsx!LN+1');
  wb.setLocation({ fileName: 'M5 b-1.xlsx' });
  assert.equal(wb.sheet(0).getFormula(0, 0), "='M5 b-1.xlsx'!LN+1");
  expectFailure(() => wb.setLocation({ fileName: 'Book1' }));
  assert.deepEqual(wb.getLocation(), { directory: null, fileName: 'M5 b-1.xlsx' });
  wb.setLocation(null);
  assert.equal(wb.getLocation(), null);
  assert.equal(wb.sheet(0).getFormula(0, 0), '=[0]!LN+1');
  release(wb);
});

test('setLocation refuses undefined and a record without fileName, and reads back the canonical form', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.throws(() => wb.setLocation(undefined), TypeError);
  assert.throws(() => wb.setLocation({ directory: '/d' }), TypeError);
  assert.throws(() => wb.setLocation('Book.xlsx'), TypeError);
  assert.equal(wb.getLocation(), null);
  for (const [given, canonical] of [
    ['/a/b/', '/a/b'],
    ['/', '/'],
    ['c:\\a\\', 'C:/a'],
    ['\\\\server\\share\\a', '//server/share/a'],
  ]) {
    wb.setLocation({ directory: given, fileName: 'b.xlsx' });
    const read = wb.getLocation();
    assert.deepEqual(read, { directory: canonical, fileName: 'b.xlsx' });
    wb.setLocation(read);
    assert.deepEqual(wb.getLocation(), read);
  }
  const refused = [{ directory: '', fileName: 'b.xlsx' }, { directory: 'rel', fileName: 'b.xlsx' }, { fileName: '' }];
  for (const bad of refused) {
    expectFailure(() => wb.setLocation(bad));
  }
  assert.deepEqual(wb.getLocation(), { directory: '//server/share/a', fileName: 'b.xlsx' });
  release(wb);
});

test('CELL and saved hyperlink targets follow the location', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setProfileId('mac-365-ja_JP');
  wb.addSheet('Sheet2');
  const ws = wb.sheet(0);
  ws.setFormula(0, 0, '=CELL("filename",A1)');
  ws.setFormula(1, 0, '=CELL("address",Sheet2!A1)');
  ws.setHyperlink({ range: 'B1', target: '/Users/libraz/Documents/a b/f.xlsx' });
  const folder = '/Users/libraz/Library/Containers/com.microsoft.Excel/Data/loc_probe3/p q';
  wb.setLocation({ directory: folder, fileName: 'Book 1.xlsx' });
  wb.recalc();
  assert.equal(ws.getValue(0, 0), `${folder}/[Book 1.xlsx]Sheet1`);
  assert.equal(ws.getValue(1, 0), "'[Book 1.xlsx]Sheet2'!$A$1");
  assert.equal(ws.listHyperlinks()[0].target, 'file:///Users/libraz/Documents/a b/f.xlsx');
  const loaded = roundTrip(mod, wb);
  const link = loaded.sheet(0).listHyperlinks()[0];
  assert.equal(link.target, '../../../../../../Documents/a%20b/f.xlsx');
  assert.equal(link.display, '../../../../../../Documents/a b/f.xlsx');
  assert.equal(loaded.getLocation(), null);
  release(loaded);
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
  wb.sheet(0).updateFormatDefaults({ defaultRowHeightPt: 18, customHeight: true });
  assert.deepEqual(
    [wb.sheet(0).getFormatDefaults().customHeight, wb.sheet(0).getFormatDefaults().zeroHeight],
    [true, false],
  );
  // An omitted flag keeps the stored one.
  wb.sheet(0).updateFormatDefaults({ defaultRowHeightPt: 18, zeroHeight: true });
  const loaded = roundTrip(mod, wb);
  const defaults = loaded.sheet(0).getFormatDefaults();
  assert.deepEqual([defaults.customHeight, defaults.zeroHeight, defaults.defaultRowHeightPt], [true, true, 18]);
  loaded.sheet(0).updateFormatDefaults({ defaultRowHeightPt: 18, customHeight: false, zeroHeight: false });
  assert.deepEqual(
    [loaded.sheet(0).getFormatDefaults().customHeight, loaded.sheet(0).getFormatDefaults().zeroHeight],
    [false, false],
  );
  release(loaded);
  release(wb);
});
