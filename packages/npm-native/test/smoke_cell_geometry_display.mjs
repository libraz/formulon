import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

const A1_B2 = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };

test('getCellXf carries apply flags, quotePrefix and protection', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const added = wb.ooxml.addCellXf({
    fontIndex: 0,
    fillIndex: 0,
    borderIndex: 0,
    numFmtId: 0,
    horizontal: 0,
    vertical: 2,
    wrapText: false,
    applyNumberFormat: true,
    applyFont: true,
    applyProtection: true,
    quotePrefix: true,
    locked: false,
    hidden: true,
  });
  const xf = wb.ooxml.getCellXf(added);
  assert.equal(xf.applyNumberFormat, true);
  assert.equal(xf.applyFont, true);
  assert.equal(xf.applyFill, false);
  assert.equal(xf.applyBorder, false);
  assert.equal(xf.applyAlignment, false);
  assert.equal(xf.applyProtection, true);
  assert.equal(xf.quotePrefix, true);
  assert.equal(xf.locked, false);
  assert.equal(xf.hidden, true);

  const plain = wb.ooxml.getCellXf(0);
  // Without a protection element both read null: the default (locked) applies.
  assert.equal(plain.locked, null);
  assert.equal(plain.hidden, null);
  assert.equal(plain.quotePrefix, false);

  assert.throws(() => wb.ooxml.getCellXf(99999), isFormulonError);
});

test('row overrides report height / customHeight and resetRowHeight drops the height', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setRowHeight(3, 3, 30);
  let rows = wb.sheet(0).listRowOverrides();
  const row = rows.find((r) => r.row === 3);
  assert.ok(row);
  assert.equal(row.heightPt, 30);
  assert.equal(row.customHeight, true);
  wb.sheet(0).resetRowHeight(3, 3);
  rows = wb.sheet(0).listRowOverrides();
  assert.equal(
    rows.find((r) => r.row === 3),
    undefined,
  );
  assert.throws(() => wb.sheet(0).resetRowHeight(2000000, 2000000), isFormulonError);
});

test('setRowAutoHeight caches the height and keeps the row automatic', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setRowAutoHeight(4, 4, 31.5);
  const row = wb
    .sheet(0)
    .listRowOverrides()
    .find((r) => r.row === 4);
  assert.ok(row);
  assert.equal(row.heightPt, 31.5);
  assert.equal(row.customHeight, false);
});

test('resetColumnWidth follows later default changes and keeps hidden / outline state', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).updateFormatDefaults({ defaultColWidthChars: 10 });
  wb.sheet(0).setColumnWidth(0, 0, 10);
  wb.sheet(0).resetColumnWidth(0, 0);
  wb.sheet(0).updateFormatDefaults({ defaultColWidthChars: 20 });
  assert.equal(wb.sheet(0).getColumnWidthPt(0, 0), wb.sheet(0).getColumnWidthPt(1, 0));
  assert.equal(wb.sheet(0).listColumnOverrides().length, 0);

  wb.sheet(0).setColumnWidth(0, 4, 20);
  wb.sheet(0).setColumnHidden(2, 2, true);
  wb.sheet(0).setColumnOutline(2, 2, 3);
  wb.sheet(0).resetColumnWidth(1, 3);
  assert.deepEqual(
    wb
      .sheet(0)
      .listColumnOverrides()
      .map((c) => [c.first, c.last, c.widthChars !== null, c.hidden, c.outlineLevel]),
    [
      [0, 0, true, false, 0],
      [2, 2, false, true, 3],
      [4, 4, true, false, 0],
    ],
  );
  assert.throws(() => wb.sheet(0).resetColumnWidth(5, 3), isFormulonError);
  assert.throws(() => wb.sheet(0).resetColumnWidth(0, 16384), isFormulonError);
});

test('getSheetFormatDefaults reports the effective default width and height', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  let d = wb.sheet(0).getFormatDefaults();
  // Stored width of the base-8 Calibri 11 default (64 px, shown as 8.43).
  assert.equal(d.effectiveDefaultColWidthChars, 2340 / 256);
  assert.equal(d.effectiveDefaultRowHeightPt, 102 / 7);
  wb.sheet(0).updateFormatDefaults({ defaultColWidthChars: 12.5, defaultRowHeightPt: 20.25 });
  d = wb.sheet(0).getFormatDefaults();
  assert.equal(d.effectiveDefaultColWidthChars, 12.5);
  assert.equal(d.effectiveDefaultRowHeightPt, 20.25);
});

test('sheet format defaults round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).updateFormatDefaults({ defaultColWidthChars: 12.5, defaultRowHeightPt: 18, baseColWidthChars: 10 });
  const d = wb.sheet(0).getFormatDefaults();
  assert.equal(d.defaultColWidthChars, 12.5);
  assert.equal(d.defaultRowHeightPt, 18);
  assert.equal(d.baseColWidthChars, 10);
  assert.throws(() => wb.sheet(0).updateFormatDefaults({ defaultColWidthChars: -1 }), isFormulonError);
  assert.throws(() => wb.sheet(99).getFormatDefaults(), isFormulonError);
});

test('geometry accessors agree with each other', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setRowHeight(0, 0, 20);
  wb.sheet(0).setRowHeight(1, 1, 30);
  const model = wb.sheet(0).getWidthModel(0);
  assert.equal(typeof model.pointsPerChar, 'number');
  assert.equal(typeof model.calibrated, 'boolean');
  assert.equal(typeof model.normalFontName, 'string');
  assert.equal(model.platform, mod.WidthPlatform.Win);

  const c0 = wb.sheet(0).getColumnWidthPt(0, 0);
  const c1 = wb.sheet(0).getColumnWidthPt(1, 0);
  assert.ok(c0 > 0);
  assert.equal(wb.sheet(0).getRowHeightPt(1), 30);

  const rect = wb.sheet(0).getCellRectPt(A1_B2, 0);
  assert.equal(rect.x, 0);
  assert.equal(rect.y, 0);
  assert.ok(Math.abs(rect.widthPt - (c0 + c1)) < 1e-9);
  assert.equal(rect.heightPt, 50);
  // The mode is optional and defaults to Display.
  assert.deepEqual(wb.sheet(0).getCellRectPt(A1_B2), rect);
  assert.deepEqual(wb.sheet(0).getWidthModel(), model);
  assert.throws(
    () => wb.sheet(0).getCellRectPt({ firstRow: 1, firstCol: 0, lastRow: 0, lastCol: 0 }, 0),
    isFormulonError,
  );
  assert.throws(() => wb.sheet(0).getCellRectPt(A1_B2, 7), RangeError);
  wb.sheet(0).getColumnWidthPt(0, 1);

  const pt = wb.sheet(0).convertColumnCharsToPt(2340 / 256, 0);
  assert.equal(pt, 48);
  const chars = wb.sheet(0).convertColumnPtToChars(pt, 0);
  assert.equal(chars, 2340 / 256);
  // The mode is last and defaults to Display.
  assert.equal(wb.sheet(0).convertColumnCharsToPt(2340 / 256), pt);
  assert.equal(wb.sheet(0).convertColumnPtToChars(pt), chars);
  assert.equal(wb.sheet(0).getColumnWidthPt(0), c0);
  assert.equal(wb.sheet(0).convertColumnCharsToPt(0, 0), 0);
  assert.throws(() => wb.sheet(0).convertColumnCharsToPt(-1, 0), isFormulonError);
  // A mode that is not a GeometryMode value (an argument left in its old place) is refused.
  assert.throws(() => wb.sheet(0).convertColumnCharsToPt(10, 7), RangeError);
  assert.throws(() => wb.sheet(0).convertColumnPtToChars(10, 7), RangeError);
  assert.throws(() => wb.sheet(0).getColumnWidthPt(0, 7), RangeError);
  assert.throws(() => wb.sheet(99).getRowHeightPt(0), isFormulonError);
});

test('getFormula / getFormulaR1C1 return null for non-formula cells', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setNumber(0, 0, 5);
  wb.sheet(0).setFormula(1, 1, '=A1+1');
  const a1 = wb.sheet(0).getFormula(1, 1);
  assert.equal(a1, '=A1+1');
  assert.equal(wb.sheet(0).getFormulaR1C1(1, 1), '=R[-1]C[-1]+1');
  assert.equal(wb.sheet(0).getFormula(0, 0), null);
  assert.equal(wb.sheet(0).getFormulaR1C1(5, 5), null);
  assert.throws(() => wb.sheet(99).getFormula(0, 0), isFormulonError);
});

test('getCellsInRange pages through populated cells with a cursor', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setNumber(0, 0, 1);
  wb.sheet(0).setText(0, 2, 'x');
  wb.sheet(0).setFormula(1, 1, '=A1+1');
  wb.sheet(0).setNumber(5, 5, 9);
  const range = { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 3 };

  const all = wb.sheet(0).getCellsInRange(range);
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
  assert.equal(all.cells[0].value, 1);
  assert.equal(all.cells[1].value, 'x');
  assert.equal(all.cells[2].formula, '=A1+1');

  const first = wb.sheet(0).getCellsInRange(range, 0, 2);
  assert.equal(first.cells.length, 2);
  assert.equal(typeof first.nextCursor, 'number');
  const rest = wb.sheet(0).getCellsInRange(range, first.nextCursor, 2);
  assert.deepEqual(
    [...first.cells, ...rest.cells].map((c) => [c.row, c.col]),
    all.cells.map((c) => [c.row, c.col]),
  );
  assert.equal(rest.nextCursor, null);

  assert.throws(
    () => wb.sheet(0).getCellsInRange({ firstRow: 3, firstCol: 0, lastRow: 0, lastCol: 0 }),
    isFormulonError,
  );
});

test('getMergesInRange lists only intersecting merges', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).addMerge({ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 });
  wb.sheet(0).addMerge({ firstRow: 10, firstCol: 10, lastRow: 11, lastCol: 11 });
  const hit = wb.sheet(0).getMergesInRange({ firstRow: 1, firstCol: 1, lastRow: 5, lastCol: 5 });
  assert.equal(hit.length, 1);
  assert.equal(hit[0].firstRow, 0);
  assert.equal(hit[0].lastCol, 1);
  assert.equal(wb.sheet(0).getMergesInRange({ firstRow: 4, firstCol: 4, lastRow: 5, lastCol: 5 }).length, 0);
  assert.throws(() => wb.sheet(99).getMergesInRange(A1_B2), isFormulonError);
});

test('getDisplayText and formatValue render through the number format', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const fmt = wb.ooxml.addNumberFormat('0.00');
  const xf = wb.ooxml.addCellXf({
    fontIndex: 0,
    fillIndex: 0,
    borderIndex: 0,
    numFmtId: fmt,
    horizontal: 0,
    vertical: 2,
    wrapText: false,
  });
  wb.sheet(0).setNumber(0, 0, 3.14159);
  wb.sheet(0).ooxml.setCellXfIndex(0, 0, xf);
  const shown = wb.sheet(0).getDisplayText(0, 0);
  assert.equal(shown.text, '3.14');
  assert.equal(shown.displayStatus, 0);
  assert.equal(wb.sheet(0).getDisplayText(7, 7).text, '');
  assert.throws(() => wb.sheet(99).getDisplayText(0, 0), isFormulonError);

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
  wb.sheet(0).setNumber(0, 0, 1);
  wb.sheet(0).setNumber(99, 5, 1);
  const p = wb.sheet(0).getPagination();
  assert.ok(p.paper.widthPt > 0 && p.paper.heightPt > 0);
  assert.equal(typeof p.paper.landscape, 'boolean');
  assert.equal(typeof p.paper.known, 'boolean');
  assert.ok(p.margins.leftIn > 0);
  assert.ok(p.printable.widthPt > 0 && p.printable.heightPt > 0);
  assert.ok(p.scalePct > 0);
  assert.ok(p.pageOrder === 0 || p.pageOrder === 1);
  assert.equal(p.printTitles, null);
  assert.equal(p.pages.length, p.pageCount);
  assert.ok(p.pages.length >= 1);
  assert.equal(p.pages[0].rect.firstRow, 0);
  assert.equal(p.pages[0].rect.firstCol, 0);
  assert.ok(p.pages[0].widthPt > 0 && p.pages[0].heightPt > 0);
  // Break provenance travels on each break; the parallel manual arrays are gone.
  assert.ok(p.rowBreaks.every((b) => typeof b.manual === 'boolean'));
  assert.ok(p.columnBreaks.every((b) => typeof b.manual === 'boolean'));
  assert.equal('rowBreakManual' in p, false);
  assert.equal('columnBreakManual' in p, false);

  assert.throws(() => wb.sheet(99).getPagination(), isFormulonError);
});
