import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

const A1_B2 = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };

test('getCellXf carries apply flags, quotePrefix and protection', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const added = wb.addXf({
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
  assert.ok(added.status.ok, JSON.stringify(added.status));
  const xf = wb.getCellXf(added.index);
  assert.ok(xf.status.ok);
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

  const failed = wb.getCellXf(99999);
  assert.equal(failed.status.ok, false);
  assert.equal(failed.quotePrefix, false);
  assert.equal(failed.locked, true);
});

test('row overrides report hasHeight / customHeight and clearRowHeight drops the height', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setRowHeight(0, 3, 30).ok);
  let rows = wb.getSheetRowOverrides(0).rows;
  const row = rows.find((r) => r.row === 3);
  assert.ok(row);
  assert.equal(row.height, 30);
  assert.equal(row.hasHeight, 1);
  assert.equal(row.customHeight, 1);
  assert.ok(wb.clearRowHeight(0, 3).ok);
  rows = wb.getSheetRowOverrides(0).rows;
  assert.equal(
    rows.find((r) => r.row === 3),
    undefined,
  );
  assert.equal(wb.clearRowHeight(0, 2000000).ok, false);
});

test('sheet format defaults round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const set = wb.setSheetFormatDefaults(0, { defaultColWidth: 12.5, defaultRowHeight: 18, baseColWidth: 10 });
  assert.ok(set.ok, JSON.stringify(set));
  const d = wb.getSheetFormatDefaults(0);
  assert.ok(d.status.ok);
  assert.equal(d.defaultColWidth, 12.5);
  assert.equal(d.defaultRowHeight, 18);
  assert.equal(d.baseColWidth, 10);
  assert.equal(d.hasDefaultColWidth, true);
  assert.equal(d.hasDefaultRowHeight, true);
  assert.equal(wb.setSheetFormatDefaults(0, { defaultColWidth: -1 }).ok, false);
  assert.equal(wb.getSheetFormatDefaults(99).status.ok, false);
  assert.equal(wb.getSheetFormatDefaults(99).hasDefaultColWidth, false);
});

test('geometry accessors agree with each other', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setRowHeight(0, 0, 20).ok);
  assert.ok(wb.setRowHeight(0, 1, 30).ok);
  const model = wb.getWidthModel(0, 0);
  assert.ok(model.status.ok, JSON.stringify(model.status));
  assert.equal(typeof model.pointsPerChar, 'number');
  assert.equal(typeof model.calibrated, 'boolean');
  assert.equal(typeof model.normalFontName, 'string');
  assert.equal(model.platform, 'win');

  const c0 = wb.getColumnWidthPt(0, 0, 0);
  const c1 = wb.getColumnWidthPt(0, 1, 0);
  assert.ok(c0.status.ok);
  assert.ok(c0.value > 0);
  assert.equal(wb.getRowHeightPt(0, 1).value, 30);

  const rect = wb.getCellRectPt(0, A1_B2, 0);
  assert.ok(rect.status.ok);
  assert.equal(rect.x, 0);
  assert.equal(rect.y, 0);
  assert.ok(Math.abs(rect.width - (c0.value + c1.value)) < 1e-9);
  assert.equal(rect.height, 50);
  assert.equal(wb.getCellRectPt(0, { firstRow: 1, firstCol: 0, lastRow: 0, lastCol: 0 }, 0).status.ok, false);
  assert.equal(wb.getCellRectPt(0, A1_B2, 7).status.ok, false);
  assert.ok(wb.getColumnWidthPt(0, 0, 1).status.ok);

  const pt = wb.columnCharsToPt(0, 0, 8.43);
  assert.ok(pt.status.ok);
  assert.ok(pt.value > 0);
  const chars = wb.columnPtToChars(0, 0, pt.value);
  assert.ok(chars.status.ok);
  assert.ok(Math.abs(chars.value - 8.43) < 1e-6);
  assert.equal(wb.columnCharsToPt(0, 0, 0).value, 0);
  assert.equal(wb.columnCharsToPt(0, 0, -1).status.ok, false);
  assert.equal(wb.getRowHeightPt(99, 0).status.ok, false);
});

test('getFormula / getFormulaR1C1 return null for non-formula cells', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setNumber(0, 0, 0, 5).ok);
  assert.ok(wb.setFormula(0, 1, 1, '=A1+1').ok);
  const a1 = wb.getFormula(0, 1, 1);
  assert.ok(a1.status.ok);
  assert.equal(a1.formula, '=A1+1');
  assert.equal(wb.getFormulaR1C1(0, 1, 1).formula, 'R[-1]C[-1]+1');
  assert.equal(wb.getFormula(0, 0, 0).formula, null);
  assert.equal(wb.getFormulaR1C1(0, 5, 5).formula, null);
  const bad = wb.getFormula(99, 0, 0);
  assert.equal(bad.status.ok, false);
  assert.equal(bad.formula, null);
});

test('getCellsInRange pages through populated cells with a cursor', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setNumber(0, 0, 0, 1).ok);
  assert.ok(wb.setText(0, 0, 2, 'x').ok);
  assert.ok(wb.setFormula(0, 1, 1, '=A1+1').ok);
  assert.ok(wb.setNumber(0, 5, 5, 9).ok);
  const range = { firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 3 };

  const all = wb.getCellsInRange(0, range);
  assert.ok(all.status.ok, JSON.stringify(all.status));
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
  assert.ok(rest.status.ok);
  assert.deepEqual(
    [...first.cells, ...rest.cells].map((c) => [c.row, c.col]),
    all.cells.map((c) => [c.row, c.col]),
  );
  assert.equal(rest.nextCursor, null);

  const bad = wb.getCellsInRange(0, { firstRow: 3, firstCol: 0, lastRow: 0, lastCol: 0 });
  assert.equal(bad.status.ok, false);
  assert.deepEqual(bad.cells, []);
  assert.equal(bad.nextCursor, null);
});

test('getMergesInRange lists only intersecting merges', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.addMerge(0, { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }).ok);
  assert.ok(wb.addMerge(0, { firstRow: 10, firstCol: 10, lastRow: 11, lastCol: 11 }).ok);
  const hit = wb.getMergesInRange(0, { firstRow: 1, firstCol: 1, lastRow: 5, lastCol: 5 });
  assert.ok(hit.status.ok);
  assert.equal(hit.length, 1);
  assert.equal(hit[0].firstRow, 0);
  assert.equal(hit[0].lastCol, 1);
  assert.equal(wb.getMergesInRange(0, { firstRow: 4, firstCol: 4, lastRow: 5, lastCol: 5 }).length, 0);
  assert.equal(wb.getMergesInRange(99, A1_B2).status.ok, false);
});

test('getDisplayText and formatValue render through the number format', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const fmt = wb.addNumFmt('0.00');
  assert.ok(fmt.status.ok);
  const xf = wb.addXf({
    fontIndex: 0,
    fillIndex: 0,
    borderIndex: 0,
    numFmtId: fmt.numFmtId,
    horizontalAlign: 0,
    verticalAlign: 2,
    wrapText: false,
  });
  assert.ok(wb.setNumber(0, 0, 0, 3.14159).ok);
  assert.ok(wb.setCellXfIndex(0, 0, 0, xf.index).ok);
  const shown = wb.getDisplayText(0, 0, 0);
  assert.ok(shown.status.ok, JSON.stringify(shown.status));
  assert.equal(shown.text, '3.14');
  assert.equal(shown.displayStatus, 0);
  assert.equal(wb.getDisplayText(0, 7, 7).text, '');
  assert.equal(wb.getDisplayText(99, 0, 0).status.ok, false);

  const num = (n) => ({ kind: mod.ValueKind.Number, number: n, boolean: 0, text: '', errorCode: 0 });
  const fixed = wb.formatValue(num(1234.5), '#,##0.0');
  assert.ok(fixed.status.ok);
  assert.equal(fixed.text, '1,234.5');
  assert.equal(fixed.displayStatus, 0);
  assert.equal(wb.formatValue(num(7), '').text, '7');
  const overflow = wb.formatValue(num(-1), 'yyyy-mm-dd');
  assert.ok(overflow.status.ok);
  assert.equal(overflow.displayStatus, 1);
  assert.equal(overflow.text, '########');
  const text = wb.formatValue({ kind: mod.ValueKind.Text, number: 0, boolean: 0, text: 'abc', errorCode: 0 }, '@');
  assert.equal(text.text, 'abc');
  assert.equal(wb.formatValue({ kind: 99 }, '').status.ok, false);
});

test('paginate reports paper, margins, printable area, pages and break provenance', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setNumber(0, 0, 0, 1).ok);
  assert.ok(wb.setNumber(0, 99, 5, 1).ok);
  const p = wb.paginate(0);
  assert.ok(p.status.ok, JSON.stringify(p.status));
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
  assert.equal(p.pages[0].firstRow, 0);
  assert.equal(p.pages[0].firstCol, 0);
  assert.ok(p.pages[0].widthPt > 0 && p.pages[0].heightPt > 0);
  assert.equal(p.horizontalBreakManual.length, p.horizontalBreaks.length);
  assert.equal(p.verticalBreakManual.length, p.verticalBreaks.length);
  assert.ok(p.horizontalBreakManual.every((m) => typeof m === 'boolean'));

  const failed = wb.paginate(99);
  assert.equal(failed.status.ok, false);
  assert.deepEqual(failed.pages, []);
  assert.equal(failed.paper.known, false);
  assert.equal(failed.printTitles.hasRows, false);
  assert.deepEqual(failed.horizontalBreakManual, []);
});
