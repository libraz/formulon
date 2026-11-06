import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

const A1_B2 = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };

test('getCellXf carries apply flags, quotePrefix and protection', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const added = wb.addCellXf({
    fontIndex: 0,
    fillIndex: 0,
    borderIndex: 0,
    numFmtId: 0,
    horizontalAlign: 0,
    verticalAlign: 2,
    wrapText: false,
    applyNumberFormat: true,
    applyFont: true,
    applyProtection: true,
    quotePrefix: true,
    locked: false,
    hidden: true,
  });
  const xf = wb.getCellXf(added);
  assert.equal(xf.applyNumberFormat, true);
  assert.equal(xf.applyFont, true);
  assert.equal(xf.applyFill, false);
  assert.equal(xf.applyBorder, false);
  assert.equal(xf.applyAlignment, false);
  assert.equal(xf.applyProtection, true);
  assert.equal(xf.quotePrefix, true);
  assert.equal(xf.hasProtection, true);
  assert.equal(xf.locked, false);
  assert.equal(xf.hidden, true);

  const plain = wb.getCellXf(0);
  assert.equal(plain.hasProtection, false);
  assert.equal(plain.locked, true);
  assert.equal(plain.hidden, false);
  assert.equal(plain.quotePrefix, false);

  assert.throws(() => wb.getCellXf(99999), isFormulonError);
});

test('row overrides report hasHeight / customHeight and clearRowHeight drops the height', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setRowHeight(0, 3, 3, 30);
  let rows = wb.listRowOverrides(0);
  const row = rows.find((r) => r.row === 3);
  assert.ok(row);
  assert.equal(row.height, 30);
  assert.equal(row.hasHeight, 1);
  assert.equal(row.customHeight, 1);
  wb.clearRowHeight(0, 3, 3);
  rows = wb.listRowOverrides(0);
  assert.equal(
    rows.find((r) => r.row === 3),
    undefined,
  );
  assert.throws(() => wb.clearRowHeight(0, 2000000, 2000000), isFormulonError);
});

test('setRowAutoHeight caches the height and keeps the row automatic', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setRowAutoHeight(0, 4, 4, 31.5);
  const row = wb.listRowOverrides(0).find((r) => r.row === 4);
  assert.ok(row);
  assert.equal(row.height, 31.5);
  assert.equal(row.hasHeight, 1);
  assert.equal(row.customHeight, 0);
});

test('clearColumnWidth follows later default changes and keeps hidden / outline state', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setFormatDefaults(0, { defaultColWidth: 10 });
  wb.setColumnWidth(0, 0, 0, 10);
  wb.clearColumnWidth(0, 0, 0);
  wb.setFormatDefaults(0, { defaultColWidth: 20 });
  assert.equal(wb.getColumnWidthPt(0, 0, 0), wb.getColumnWidthPt(0, 1, 0));
  assert.equal(wb.listColumnOverrides(0).length, 0);

  wb.setColumnWidth(0, 0, 4, 20);
  wb.setColumnHidden(0, 2, 2, true);
  wb.setColumnOutline(0, 2, 2, 3);
  wb.clearColumnWidth(0, 1, 3);
  assert.deepEqual(
    wb.listColumnOverrides(0).map((c) => [c.first, c.last, c.hasWidth, c.hidden, c.outlineLevel]),
    [
      [0, 0, 1, 0, 0],
      [2, 2, 0, 1, 3],
      [4, 4, 1, 0, 0],
    ],
  );
  assert.throws(() => wb.clearColumnWidth(0, 5, 3), isFormulonError);
  assert.throws(() => wb.clearColumnWidth(0, 0, 16384), isFormulonError);
});

test('getSheetFormatDefaults reports the effective default width and height', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  let d = wb.getFormatDefaults(0);
  // Stored width of the base-8 Calibri 11 default (64 px, shown as 8.43).
  assert.equal(d.effectiveDefaultColWidth, 2340 / 256);
  assert.equal(d.effectiveDefaultRowHeight, 102 / 7);
  wb.setFormatDefaults(0, { defaultColWidth: 12.5, defaultRowHeight: 20.25 });
  d = wb.getFormatDefaults(0);
  assert.equal(d.effectiveDefaultColWidth, 12.5);
  assert.equal(d.effectiveDefaultRowHeight, 20.25);
});

test('sheet format defaults round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setFormatDefaults(0, { defaultColWidth: 12.5, defaultRowHeight: 18, baseColWidth: 10 });
  const d = wb.getFormatDefaults(0);
  assert.equal(d.defaultColWidth, 12.5);
  assert.equal(d.defaultRowHeight, 18);
  assert.equal(d.baseColWidth, 10);
  assert.equal(d.hasDefaultColWidth, true);
  assert.equal(d.hasDefaultRowHeight, true);
  assert.throws(() => wb.setFormatDefaults(0, { defaultColWidth: -1 }), isFormulonError);
  assert.throws(() => wb.getFormatDefaults(99), isFormulonError);
});

test('geometry accessors agree with each other', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setRowHeight(0, 0, 0, 20);
  wb.setRowHeight(0, 1, 1, 30);
  const model = wb.getWidthModel(0, 0);
  assert.equal(typeof model.pointsPerChar, 'number');
  assert.equal(typeof model.calibrated, 'boolean');
  assert.equal(typeof model.normalFontName, 'string');
  assert.equal(model.platform, 'win');

  const c0 = wb.getColumnWidthPt(0, 0, 0);
  const c1 = wb.getColumnWidthPt(0, 1, 0);
  assert.ok(c0 > 0);
  assert.equal(wb.getRowHeightPt(0, 1), 30);

  const rect = wb.getCellRectPt(0, A1_B2, 0);
  assert.equal(rect.x, 0);
  assert.equal(rect.y, 0);
  assert.ok(Math.abs(rect.width - (c0 + c1)) < 1e-9);
  assert.equal(rect.height, 50);
  assert.throws(() => wb.getCellRectPt(0, { firstRow: 1, firstCol: 0, lastRow: 0, lastCol: 0 }, 0), isFormulonError);
  assert.throws(() => wb.getCellRectPt(0, A1_B2, 7), isFormulonError);
  wb.getColumnWidthPt(0, 0, 1);

  const pt = wb.convertColumnCharsToPt(0, 0, 2340 / 256);
  assert.equal(pt, 48);
  const chars = wb.convertColumnPtToChars(0, 0, pt);
  assert.equal(chars, 2340 / 256);
  assert.equal(wb.convertColumnCharsToPt(0, 0, 0), 0);
  assert.throws(() => wb.convertColumnCharsToPt(0, 0, -1), isFormulonError);
  assert.throws(() => wb.getRowHeightPt(99, 0), isFormulonError);
});

test('getFormula / getFormulaR1C1 return null for non-formula cells', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setNumber(0, 0, 0, 5);
  wb.setFormula(0, 1, 1, '=A1+1');
  const a1 = wb.getFormula(0, 1, 1);
  assert.equal(a1, '=A1+1');
  assert.equal(wb.getFormulaR1C1(0, 1, 1), 'R[-1]C[-1]+1');
  assert.equal(wb.getFormula(0, 0, 0), null);
  assert.equal(wb.getFormulaR1C1(0, 5, 5), null);
  assert.throws(() => wb.getFormula(99, 0, 0), isFormulonError);
});

test('getCellsInRange pages through populated cells with a cursor', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setNumber(0, 0, 0, 1);
  wb.setText(0, 0, 2, 'x');
  wb.setFormula(0, 1, 1, '=A1+1');
  wb.setNumber(0, 5, 5, 9);
  const range = { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 3 };

  const all = wb.getCellsInRange(0, range);
  assert.equal(all.nextCursor, null);
  assert.deepEqual(
    all.cells.map((c) => [c.row, c.col]),
    [
      [0, 0],
      [0, 2],
      [1, 1],
    ],
  );
  assert.equal(all.cells[0].formula, null);
  assert.equal(all.cells[0].value.number, 1);
  assert.equal(all.cells[1].value.text, 'x');
  assert.equal(all.cells[2].formula, '=A1+1');

  const first = wb.getCellsInRange(0, range, 0, 2);
  assert.equal(first.cells.length, 2);
  assert.equal(typeof first.nextCursor, 'number');
  const rest = wb.getCellsInRange(0, range, first.nextCursor, 2);
  assert.deepEqual(
    [...first.cells, ...rest.cells].map((c) => [c.row, c.col]),
    all.cells.map((c) => [c.row, c.col]),
  );
  assert.equal(rest.nextCursor, null);

  assert.throws(() => wb.getCellsInRange(0, { firstRow: 3, firstCol: 0, lastRow: 0, lastCol: 0 }), isFormulonError);
});

test('getMergesInRange lists only intersecting merges', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.addMerge(0, { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 });
  wb.addMerge(0, { firstRow: 10, firstCol: 10, lastRow: 11, lastCol: 11 });
  const hit = wb.getMergesInRange(0, { firstRow: 1, firstCol: 1, lastRow: 5, lastCol: 5 });
  assert.equal(hit.length, 1);
  assert.equal(hit[0].firstRow, 0);
  assert.equal(hit[0].lastCol, 1);
  assert.equal(wb.getMergesInRange(0, { firstRow: 4, firstCol: 4, lastRow: 5, lastCol: 5 }).length, 0);
  assert.throws(() => wb.getMergesInRange(99, A1_B2), isFormulonError);
});

test('getDisplayText and formatValue render through the number format', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const fmt = wb.addNumFmt('0.00');
  const xf = wb.addCellXf({
    fontIndex: 0,
    fillIndex: 0,
    borderIndex: 0,
    numFmtId: fmt,
    horizontalAlign: 0,
    verticalAlign: 2,
    wrapText: false,
  });
  wb.setNumber(0, 0, 0, 3.14159);
  wb.setCellXfIndex(0, 0, 0, xf);
  const shown = wb.getDisplayText(0, 0, 0);
  assert.equal(shown.text, '3.14');
  assert.equal(shown.displayStatus, 0);
  assert.equal(wb.getDisplayText(0, 7, 7).text, '');
  assert.throws(() => wb.getDisplayText(99, 0, 0), isFormulonError);

  const num = (n) => ({ kind: mod.ValueKind.Number, number: n });
  const fixed = wb.formatValue(num(1234.5), '#,##0.0');
  assert.equal(fixed.text, '1,234.5');
  assert.equal(fixed.displayStatus, 0);
  assert.equal(wb.formatValue(num(7), '').text, '7');
  const overflow = wb.formatValue(num(-1), 'yyyy-mm-dd');
  assert.equal(overflow.displayStatus, 1);
  assert.equal(overflow.text, '########');
  const text = wb.formatValue({ kind: mod.ValueKind.Text, text: 'abc' }, '@');
  assert.equal(text.text, 'abc');
  assert.throws(() => wb.formatValue({ kind: 99 }, ''), TypeError);
  assert.throws(
    () => wb.formatValue({ number: 0.5 }, '0.0%'),
    (e) => e instanceof TypeError && /requires a `kind`|kind/.test(e.message),
    'a value without `kind` is rejected, not read as blank',
  );
});

test('paginate reports paper, margins, printable area, pages and break provenance', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setNumber(0, 0, 0, 1);
  wb.setNumber(0, 99, 5, 1);
  const p = wb.getPagination(0);
  assert.ok(p.paper.widthPt > 0 && p.paper.heightPt > 0);
  assert.equal(typeof p.paper.landscape, 'boolean');
  assert.equal(typeof p.paper.known, 'boolean');
  assert.ok(p.margins.left > 0);
  assert.ok(p.printable.width > 0 && p.printable.height > 0);
  assert.ok(p.scale > 0);
  assert.ok(p.pageOrder === 0 || p.pageOrder === 1);
  assert.equal(p.printTitles.hasRows, false);
  assert.equal(p.printTitles.hasCols, false);
  assert.equal(p.pages.length, p.pageCount);
  assert.ok(p.pages.length >= 1);
  assert.equal(p.pages[0].rect.firstRow, 0);
  assert.equal(p.pages[0].rect.firstCol, 0);
  assert.ok(p.pages[0].widthPt > 0 && p.pages[0].heightPt > 0);
  assert.equal(p.horizontalBreakManual.length, p.horizontalBreaks.length);
  assert.equal(p.verticalBreakManual.length, p.verticalBreaks.length);
  assert.ok(p.horizontalBreakManual.every((m) => typeof m === 'boolean'));

  assert.throws(() => wb.getPagination(99), isFormulonError);
});
