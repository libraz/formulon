import assert from 'node:assert/strict';
import { StatusCode } from '../../packages/npm/common.mjs';
import { expectFailure } from './smoke_support.mjs';

const RANGE_A1_F6 = { firstRow: 0, firstCol: 0, lastRow: 5, lastCol: 5 };

function withWorkbook(Module, fn) {
  const wb = Module.Workbook.createDefault();
  try {
    fn(wb);
  } finally {
    wb.delete();
  }
}

export function registerCellGeometryDisplay(Module, test) {
  test('clearColumnWidth returns a column to the sheet default and follows later default changes', () => {
    withWorkbook(Module, (wb) => {
      wb.setFormatDefaults(0, { defaultColWidth: 10 });
      wb.setColumnWidth(0, 0, 0, 10);
      wb.clearColumnWidth(0, 0, 0);
      wb.setFormatDefaults(0, { defaultColWidth: 20 });
      assert.equal(wb.getColumnWidthPt(0, 0, 0), wb.getColumnWidthPt(0, 1, 0));
      assert.equal(wb.listColumnOverrides(0).length, 0);
    });
  });

  test('clearColumnWidth splits a span, keeps hidden / outline, and rejects bad arguments', () => {
    withWorkbook(Module, (wb) => {
      wb.setColumnWidth(0, 0, 4, 20);
      wb.setColumnHidden(0, 2, 2, true);
      wb.setColumnOutline(0, 2, 2, 3);
      wb.clearColumnWidth(0, 1, 3);
      const cols = wb.listColumnOverrides(0);
      assert.deepEqual(
        cols.map((c) => [c.first, c.last, c.hasWidth, c.hidden, c.outlineLevel]),
        [
          [0, 0, 1, 0, 0],
          [2, 2, 0, 1, 3],
          [4, 4, 1, 0, 0],
        ],
      );
      wb.clearColumnWidth(0, 8, 9);
      assert.equal(wb.listColumnOverrides(0).length, 3);
      expectFailure(() => wb.clearColumnWidth(0, 5, 3));
      expectFailure(() => wb.clearColumnWidth(0, 0, 16384));
      expectFailure(() => wb.clearColumnWidth(9, 0, 0));
      assert.equal(wb.listColumnOverrides(0).length, 3);
    });
  });

  test('getSheetFormatDefaults reports the effective default width and height', () => {
    withWorkbook(Module, (wb) => {
      let d = wb.getFormatDefaults(0);
      assert.equal(d.hasDefaultColWidth, false);
      assert.equal(d.effectiveDefaultColWidth, 2340 / 256); // Base-8 Calibri 11 default, shown as 8.43.
      assert.equal(d.effectiveDefaultRowHeight, 102 / 7);
      wb.setFormatDefaults(0, { defaultColWidth: 12.5, defaultRowHeight: 20.25 });
      d = wb.getFormatDefaults(0);
      assert.equal(d.effectiveDefaultColWidth, 12.5);
      assert.equal(d.effectiveDefaultRowHeight, 20.25);
    });
  });

  test('getFormula / getFormulaR1C1 report the stored formula or null', () => {
    withWorkbook(Module, (wb) => {
      wb.setNumber(0, 0, 0, 5);
      wb.setFormula(0, 0, 1, '=A1*2');
      const a1 = wb.getFormula(0, 0, 1);
      assert.equal(a1, '=A1*2');
      assert.equal(wb.getFormula(0, 0, 0), null);
      const r1c1 = wb.getFormulaR1C1(0, 0, 1);
      assert.equal(r1c1, 'RC[-1]*2');
      assert.equal(wb.getFormulaR1C1(0, 0, 0), null);
      expectFailure(() => wb.getFormula(9, 0, 0), StatusCode.SheetIndexOutOfRange);
    });
  });

  test('getCellsInRange pages through populated cells with a cursor', () => {
    withWorkbook(Module, (wb) => {
      wb.setNumber(0, 0, 0, 1);
      wb.setFormula(0, 0, 1, '=A1+1');
      wb.setText(0, 1, 0, 'hi');
      wb.recalc();

      const all = wb.getCellsInRange(0, RANGE_A1_F6);
      assert.equal(all.cells.length, 3);
      assert.equal(all.nextCursor, null);
      assert.equal(all.cells[1].formula, '=A1+1');
      assert.equal(all.cells[1].value.number, 2);
      assert.equal(all.cells[2].value.text, 'hi');

      const seen = [];
      let cursor = null;
      for (let guard = 0; guard < 10; guard += 1) {
        const page = wb.getCellsInRange(0, RANGE_A1_F6, cursor, 1);
        assert.ok(page.cells.length <= 1);
        seen.push(...page.cells.map((c) => `${c.row},${c.col}`));
        if (page.nextCursor === null) break;
        cursor = page.nextCursor;
      }
      assert.deepEqual(seen, ['0,0', '0,1', '1,0']);

      assert.throws(() => wb.getCellsInRange(0, RANGE_A1_F6, -1), RangeError);
      const unchanged = wb.getCellsInRange(0, RANGE_A1_F6);
      assert.deepEqual(unchanged.cells, all.cells);
      assert.equal(unchanged.nextCursor, null);
    });
  });

  test('getCellsInRange throws on every invalid range integer before the C call', () => {
    const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Symbol('range integer')];
    withWorkbook(Module, (wb) => {
      wb.setNumber(0, 0, 0, 7);
      for (const value of invalid) {
        for (const key of ['firstRow', 'firstCol', 'lastRow', 'lastCol']) {
          assert.throws(
            () =>
              wb.getCellsInRange(0, {
                firstRow: 0,
                firstCol: 0,
                lastRow: 4,
                lastCol: 4,
                [key]: value,
              }),
            value === 2 ** 32 ? RangeError : TypeError,
            `getCellsInRange ${key}=${String(value)} accepted`,
          );
        }
      }

      const range = { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 4 };
      Object.defineProperty(range, 'firstRow', {
        enumerable: true,
        get() {
          throw new Error('cell range getter');
        },
      });
      assert.throws(() => wb.getCellsInRange(0, range), /cell range getter/);
    });
  });

  test('getMergesInRange lists only intersecting merges', () => {
    withWorkbook(Module, (wb) => {
      const merge = { firstRow: 0, lastRow: 1, firstCol: 3, lastCol: 4 };
      wb.addMerge(0, merge);
      const hit = wb.getMergesInRange(0, { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 3 });
      assert.deepEqual([...hit], [merge]);
      const miss = wb.getMergesInRange(0, { firstRow: 5, lastRow: 6, firstCol: 0, lastCol: 0 });
      assert.equal(miss.length, 0);
      expectFailure(() => wb.getMergesInRange(9, merge));
    });
  });

  test('getDisplayText / formatValue render through the number format', () => {
    withWorkbook(Module, (wb) => {
      wb.setNumber(0, 0, 0, 1234.5);
      const plain = wb.getDisplayText(0, 0, 0);
      assert.equal(plain.text, '1234.5');
      assert.equal(plain.displayStatus, 0);
      assert.equal(wb.getDisplayText(0, 7, 7).text, '');

      const pct = wb.formatValue({ kind: 1, number: 0.5 }, '0.0%');
      assert.equal(pct.text, '50.0%');
      assert.equal(pct.displayStatus, 0);
      const txt = wb.formatValue({ kind: 3, text: 'abc' }, '@');
      assert.equal(txt.text, 'abc');
      const bad = wb.formatValue({ kind: 1, number: 1 }, '[[');
      assert.equal(bad.displayStatus, 2);
      assert.throws(() => wb.formatValue({ number: 0.5 }, '0.0%'), TypeError);
      expectFailure(() => wb.getDisplayText(9, 0, 0));
    });
  });

  test('formatValue shares native fraction rounding and convergent selection', () => {
    withWorkbook(Module, (wb) => {
      // Fixed denominators round binary64 products; variable denominators keep the last bounded convergent.
      const cases = [
        [0.015, '?/100', '2/100'],
        [0.0045, '?/1000', '5/1000'],
        [0.6125, '?/?', '3/5'],
      ];
      for (const [number, format, expected] of cases) {
        const out = wb.formatValue({ kind: 1, number }, format);
        assert.equal(out.text, expected, `${number} "${format}"`);
      }
    });
  });

  test('nested SUBTOTAL skips inner subtotal cells in value, formula and display text', () => {
    withWorkbook(Module, (wb) => {
      wb.setNumber(0, 0, 0, 1);
      wb.setNumber(0, 1, 0, 2);
      wb.setNumber(0, 2, 0, 3);
      wb.setFormula(0, 3, 0, '=SUBTOTAL(9,A1:A3)');
      wb.setFormula(0, 4, 0, '=SUBTOTAL(9,A1:A4)');
      wb.recalc();
      assert.equal(wb.getValue(0, 3, 0).number, 6);
      assert.equal(wb.getValue(0, 4, 0).number, 6);
      assert.equal(wb.getFormula(0, 4, 0), '=SUBTOTAL(9,A1:A4)');
      assert.equal(wb.getFormulaR1C1(0, 4, 0), 'SUBTOTAL(9,R[-4]C:R[-1]C)');
      assert.equal(wb.getDisplayText(0, 4, 0).text, '6');
      const page = wb.getCellsInRange(0, { firstRow: 3, firstCol: 0, lastRow: 4, lastCol: 0 });
      assert.deepEqual(
        page.cells.map((c) => c.value.number),
        [6, 6],
      );
    });
  });

  test('row heights: clearRowHeight and the has/custom flags on row overrides', () => {
    withWorkbook(Module, (wb) => {
      wb.setRowHeight(0, 3, 3, 30);
      const rows = wb.listRowOverrides(0);
      assert.equal(rows.length, 1);
      assert.equal(rows[0].height, 30);
      assert.equal(rows[0].hasHeight, 1);
      assert.equal(rows[0].customHeight, 1);
      wb.clearRowHeight(0, 3, 3);
      assert.equal(wb.listRowOverrides(0).length, 0);
      wb.clearRowHeight(0, 3, 3);
      expectFailure(() => wb.clearRowHeight(9, 0, 0));
    });
  });

  test('row heights: setRowAutoHeight caches the height and keeps the row automatic', () => {
    withWorkbook(Module, (wb) => {
      wb.setRowHidden(0, 4, 4, true);
      wb.setRowAutoHeight(0, 4, 4, 31.5);
      const rows = wb.listRowOverrides(0);
      assert.equal(rows.length, 1);
      assert.equal(rows[0].height, 31.5);
      assert.equal(rows[0].hasHeight, 1);
      assert.equal(rows[0].customHeight, 0);
      assert.equal(rows[0].hidden, 1);
      expectFailure(() => wb.setRowAutoHeight(0, 4, 4, -1));
    });
  });

  test('sheet format defaults round-trip with presence flags', () => {
    withWorkbook(Module, (wb) => {
      const initial = wb.getFormatDefaults(0);
      assert.equal(initial.hasDefaultColWidth, false);
      assert.equal(initial.hasDefaultRowHeight, false);
      assert.equal(initial.baseColWidth, 8);

      wb.setFormatDefaults(0, { defaultColWidth: 10, defaultRowHeight: 18, baseColWidth: 9 });
      const set = wb.getFormatDefaults(0);
      assert.equal(set.defaultColWidth, 10);
      assert.equal(set.defaultRowHeight, 18);
      assert.equal(set.baseColWidth, 9);
      assert.equal(set.hasDefaultColWidth, true);
      assert.equal(set.hasDefaultRowHeight, true);

      expectFailure(() => wb.setFormatDefaults(0, { defaultColWidth: -1 }));
      assert.equal(wb.getFormatDefaults(0).defaultColWidth, 10);
      expectFailure(() => wb.getFormatDefaults(9));
    });
  });

  test('point geometry: rect, widths, heights, width model and conversions', () => {
    withWorkbook(Module, (wb) => {
      const range = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };
      const rect = wb.getCellRectPt(0, range, 0);
      assert.equal(rect.x, 0);
      assert.equal(rect.y, 0);
      assert.ok(rect.width > 0 && rect.height > 0);

      const col = wb.getColumnWidthPt(0, 0, 1);
      const row = wb.getRowHeightPt(0, 0);
      assert.ok(col > 0);
      assert.ok(row > 0);

      const wide = wb.getCellRectPt(0, { ...range, lastRow: 0, lastCol: 0 }, 0);
      assert.ok(Math.abs(wide.height - row) < 1e-9);

      wb.setColumnHidden(0, 0, 0, true);
      assert.equal(wb.getColumnWidthPt(0, 0, 0), 0);
      wb.setColumnHidden(0, 0, 0, false);

      const model = wb.getWidthModel(0, 0);
      assert.equal(typeof model.calibrated, 'boolean');
      assert.equal(typeof model.normalFontName, 'string');
      assert.ok(model.pointsPerChar > 0);
      assert.ok(model.platform.length > 0);

      const pt = wb.convertColumnCharsToPt(0, 0, 10);
      // Stored width 10 is Truncate((2560 + 18) / 256 * 7) = 70 px with Calibri 11.
      assert.equal(pt, 70 * 0.75);
      const chars = wb.convertColumnPtToChars(0, 0, pt);
      assert.ok(Math.abs(chars - 10) < 1e-6);
      assert.equal(wb.convertColumnCharsToPt(0, 0, 0), 0);
      expectFailure(() => wb.convertColumnCharsToPt(0, 0, -1));
      expectFailure(() => wb.getWidthModel(0, 7));
      expectFailure(() => wb.getCellRectPt(0, range, 7));
    });
  });

  test('cell xf carries apply flags, quotePrefix and protection', () => {
    withWorkbook(Module, (wb) => {
      const base = wb.getCellXf(0);
      assert.equal(base.hasProtection, false);
      assert.equal(base.locked, true);
      assert.equal(base.hidden, false);
      assert.equal(base.quotePrefix, false);
      assert.equal(base.applyFont, false);

      const added = wb.addCellXf({
        fontIndex: 0,
        fillIndex: 0,
        borderIndex: 0,
        numFmtId: 0,
        horizontalAlign: 0,
        verticalAlign: 2,
        wrapText: false,
        applyFont: true,
        applyProtection: true,
        quotePrefix: true,
        locked: false,
        hidden: true,
      });
      const xf = wb.getCellXf(added);
      assert.equal(xf.applyFont, true);
      assert.equal(xf.applyProtection, true);
      assert.equal(xf.applyNumberFormat, false);
      assert.equal(xf.quotePrefix, true);
      assert.equal(xf.hasProtection, true);
      assert.equal(xf.locked, false);
      assert.equal(xf.hidden, true);
    });
  });

  test('paginate carries paper, margins, scale, titles, pages and manual-break flags', () => {
    withWorkbook(Module, (wb) => {
      wb.setNumber(0, 0, 0, 1);
      wb.setNumber(0, 120, 0, 2);
      wb.addRowBreak(0, 10, true);
      wb.setPrintTitles(0, '1:2', '');
      const r = wb.getPagination(0);
      assert.ok(r.paper.widthPt > 0 && r.paper.heightPt > 0);
      assert.equal(typeof r.paper.landscape, 'boolean');
      assert.equal(typeof r.paper.known, 'boolean');
      assert.ok(r.margins.left > 0 && r.margins.top > 0);
      assert.ok(r.printable.width > 0 && r.printable.height > 0);
      assert.equal(r.scale, 1);
      assert.equal(r.pageOrder, 0);
      assert.equal(r.printTitles.hasRows, true);
      assert.equal(r.printTitles.firstRow, 0);
      assert.equal(r.printTitles.lastRow, 1);
      assert.equal(r.printTitles.hasCols, false);
      assert.equal(r.pages.length, r.pageCount);
      assert.ok(r.pages.length >= 2);
      assert.equal(r.pages[0].rect.firstRow, 0);
      assert.ok(r.pages[0].widthPt > 0);
      assert.equal(r.horizontalBreakManual.length, r.horizontalBreaks.length);
      assert.equal(r.verticalBreakManual.length, r.verticalBreaks.length);
      const manualAt = r.horizontalBreaks.indexOf(10);
      assert.ok(manualAt >= 0, `breaks=${JSON.stringify(r.horizontalBreaks)}`);
      assert.equal(r.horizontalBreakManual[manualAt], true);
      assert.ok(r.horizontalBreakManual.some((m) => m === false));

      expectFailure(() => wb.getPagination(9), StatusCode.SheetIndexOutOfRange);
    });
  });
}
