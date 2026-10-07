import assert from 'node:assert/strict';

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
      assert.ok(wb.setSheetFormatDefaults(0, { defaultColWidth: 10 }).ok);
      assert.ok(wb.setColumnWidth(0, 0, 0, 10).ok);
      assert.ok(wb.clearColumnWidth(0, 0, 0).ok);
      assert.ok(wb.setSheetFormatDefaults(0, { defaultColWidth: 20 }).ok);
      assert.equal(wb.getColumnWidthPt(0, 0, 0).value, wb.getColumnWidthPt(0, 1, 0).value);
      assert.equal(wb.getSheetColumns(0).columns.length, 0);
    });
  });

  test('clearColumnWidth splits a span, keeps hidden / outline, and rejects bad arguments', () => {
    withWorkbook(Module, (wb) => {
      assert.ok(wb.setColumnWidth(0, 0, 4, 20).ok);
      assert.ok(wb.setColumnHidden(0, 2, 2, true).ok);
      assert.ok(wb.setColumnOutline(0, 2, 2, 3).ok);
      assert.ok(wb.clearColumnWidth(0, 1, 3).ok);
      const cols = wb.getSheetColumns(0).columns;
      assert.deepEqual(
        cols.map((c) => [c.first, c.last, c.hasWidth, c.hidden, c.outlineLevel]),
        [
          [0, 0, 1, 0, 0],
          [2, 2, 0, 1, 3],
          [4, 4, 1, 0, 0],
        ],
      );
      assert.ok(wb.clearColumnWidth(0, 8, 9).ok);
      assert.equal(wb.getSheetColumns(0).columns.length, 3);
      assert.equal(wb.clearColumnWidth(0, 5, 3).ok, false);
      assert.equal(wb.clearColumnWidth(0, 0, 16384).ok, false);
      assert.equal(wb.clearColumnWidth(9, 0, 0).ok, false);
      assert.equal(wb.getSheetColumns(0).columns.length, 3);
    });
  });

  test('getSheetFormatDefaults reports the effective default width and height', () => {
    withWorkbook(Module, (wb) => {
      let d = wb.getSheetFormatDefaults(0);
      assert.equal(d.hasDefaultColWidth, false);
      assert.equal(d.effectiveDefaultColWidth, 8.43);
      assert.equal(d.effectiveDefaultRowHeight, 102 / 7);
      assert.ok(wb.setSheetFormatDefaults(0, { defaultColWidth: 12.5, defaultRowHeight: 20.25 }).ok);
      d = wb.getSheetFormatDefaults(0);
      assert.equal(d.effectiveDefaultColWidth, 12.5);
      assert.equal(d.effectiveDefaultRowHeight, 20.25);
    });
  });

  test('getFormula / getFormulaR1C1 report the stored formula or null', () => {
    withWorkbook(Module, (wb) => {
      assert.ok(wb.setNumber(0, 0, 0, 5).ok);
      assert.ok(wb.setFormula(0, 0, 1, '=A1*2').ok);
      const a1 = wb.getFormula(0, 0, 1);
      assert.ok(a1.status.ok);
      assert.equal(a1.formula, '=A1*2');
      assert.equal(wb.getFormula(0, 0, 0).formula, null);
      const r1c1 = wb.getFormulaR1C1(0, 0, 1);
      assert.ok(r1c1.status.ok);
      assert.equal(r1c1.formula, 'RC[-1]*2');
      assert.equal(wb.getFormulaR1C1(0, 0, 0).formula, null);
      assert.equal(wb.getFormula(9, 0, 0).status.ok, false);
      assert.equal(wb.getFormula(9, 0, 0).formula, null);
    });
  });

  test('getCellsInRange pages through populated cells with a cursor', () => {
    withWorkbook(Module, (wb) => {
      assert.ok(wb.setNumber(0, 0, 0, 1).ok);
      assert.ok(wb.setFormula(0, 0, 1, '=A1+1').ok);
      assert.ok(wb.setText(0, 1, 0, 'hi').ok);
      assert.ok(wb.recalc().ok);

      const all = wb.getCellsInRange(0, RANGE_A1_F6);
      assert.ok(all.status.ok);
      assert.equal(all.cells.length, 3);
      assert.equal(all.nextCursor, null);
      assert.equal(all.cells[1].formula, '=A1+1');
      assert.equal(all.cells[1].value.number, 2);
      assert.equal(all.cells[2].value.text, 'hi');

      const seen = [];
      let cursor = null;
      for (let guard = 0; guard < 10; guard += 1) {
        const page = wb.getCellsInRange(0, RANGE_A1_F6, cursor, 1);
        assert.ok(page.status.ok);
        assert.ok(page.cells.length <= 1);
        seen.push(...page.cells.map((c) => `${c.row},${c.col}`));
        if (page.nextCursor === null) break;
        cursor = page.nextCursor;
      }
      assert.deepEqual(seen, ['0,0', '0,1', '1,0']);

      assert.throws(() => wb.getCellsInRange(0, RANGE_A1_F6, -1), RangeError);
      const unchanged = wb.getCellsInRange(0, RANGE_A1_F6);
      assert.ok(unchanged.status.ok);
      assert.deepEqual(unchanged.cells, all.cells);
      assert.equal(unchanged.nextCursor, null);
    });
  });

  test('getCellsInRange throws on every invalid range integer before the C call', () => {
    const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Symbol('range integer')];
    withWorkbook(Module, (wb) => {
      assert.ok(wb.setNumber(0, 0, 0, 7).ok);
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
            typeof value === 'symbol' ? TypeError : RangeError,
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
      assert.ok(wb.addMerge(0, merge).ok);
      const hit = wb.getMergesInRange(0, { firstRow: 0, lastRow: 0, firstCol: 0, lastCol: 3 });
      assert.ok(hit.status.ok);
      assert.deepEqual([...hit], [merge]);
      const miss = wb.getMergesInRange(0, { firstRow: 5, lastRow: 6, firstCol: 0, lastCol: 0 });
      assert.ok(miss.status.ok);
      assert.equal(miss.length, 0);
      assert.equal(wb.getMergesInRange(9, merge).status.ok, false);
    });
  });

  test('getDisplayText / formatValue render through the number format', () => {
    withWorkbook(Module, (wb) => {
      assert.ok(wb.setNumber(0, 0, 0, 1234.5).ok);
      const plain = wb.getDisplayText(0, 0, 0);
      assert.ok(plain.status.ok);
      assert.equal(plain.text, '1234.5');
      assert.equal(plain.displayStatus, 0);
      assert.equal(wb.getDisplayText(0, 7, 7).text, '');

      const pct = wb.formatValue({ kind: 1, number: 0.5, boolean: 0, text: '', errorCode: 0 }, '0.0%');
      assert.ok(pct.status.ok);
      assert.equal(pct.text, '50.0%');
      assert.equal(pct.displayStatus, 0);
      const txt = wb.formatValue({ kind: 3, number: 0, boolean: 0, text: 'abc', errorCode: 0 }, '@');
      assert.equal(txt.text, 'abc');
      const bad = wb.formatValue({ kind: 1, number: 1, boolean: 0, text: '', errorCode: 0 }, '[[');
      assert.ok(bad.status.ok);
      assert.equal(bad.displayStatus, 2);
      assert.equal(wb.getDisplayText(9, 0, 0).status.ok, false);
    });
  });

  test('formatValue picks fraction candidates in binary64, as native does', () => {
    withWorkbook(Module, (wb) => {
      // A wider long double keeps 0.015 * 100 below 1.5 and picks 1/100, 4/1000 and 3/5.
      const cases = [
        [0.015, '?/100', '2/100'],
        [0.0045, '?/1000', '5/1000'],
        [0.6125, '?/?', '5/8'],
      ];
      for (const [number, format, expected] of cases) {
        const out = wb.formatValue({ kind: 1, number, boolean: 0, text: '', errorCode: 0 }, format);
        assert.ok(out.status.ok);
        assert.equal(out.text, expected, `${number} "${format}"`);
      }
    });
  });

  test('nested SUBTOTAL skips inner subtotal cells in value, formula and display text', () => {
    withWorkbook(Module, (wb) => {
      assert.ok(wb.setNumber(0, 0, 0, 1).ok);
      assert.ok(wb.setNumber(0, 1, 0, 2).ok);
      assert.ok(wb.setNumber(0, 2, 0, 3).ok);
      assert.ok(wb.setFormula(0, 3, 0, '=SUBTOTAL(9,A1:A3)').ok);
      assert.ok(wb.setFormula(0, 4, 0, '=SUBTOTAL(9,A1:A4)').ok);
      assert.ok(wb.recalc().ok);
      assert.equal(wb.getValue(0, 3, 0).value.number, 6);
      assert.equal(wb.getValue(0, 4, 0).value.number, 6);
      assert.equal(wb.getFormula(0, 4, 0).formula, '=SUBTOTAL(9,A1:A4)');
      assert.equal(wb.getFormulaR1C1(0, 4, 0).formula, 'SUBTOTAL(9,R[-4]C:R[-1]C)');
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
      assert.ok(wb.setRowHeight(0, 3, 30).ok);
      const rows = wb.getSheetRowOverrides(0).rows;
      assert.equal(rows.length, 1);
      assert.equal(rows[0].height, 30);
      assert.equal(rows[0].hasHeight, 1);
      assert.equal(rows[0].customHeight, 1);
      assert.ok(wb.clearRowHeight(0, 3).ok);
      assert.equal(wb.getSheetRowOverrides(0).rows.length, 0);
      assert.ok(wb.clearRowHeight(0, 3).ok);
      assert.equal(wb.clearRowHeight(9, 0).ok, false);
    });
  });

  test('sheet format defaults round-trip with presence flags', () => {
    withWorkbook(Module, (wb) => {
      const initial = wb.getSheetFormatDefaults(0);
      assert.ok(initial.status.ok);
      assert.equal(initial.hasDefaultColWidth, false);
      assert.equal(initial.hasDefaultRowHeight, false);
      assert.equal(initial.baseColWidth, 8);

      assert.ok(wb.setSheetFormatDefaults(0, { defaultColWidth: 10, defaultRowHeight: 18, baseColWidth: 9 }).ok);
      const set = wb.getSheetFormatDefaults(0);
      assert.equal(set.defaultColWidth, 10);
      assert.equal(set.defaultRowHeight, 18);
      assert.equal(set.baseColWidth, 9);
      assert.equal(set.hasDefaultColWidth, true);
      assert.equal(set.hasDefaultRowHeight, true);

      assert.equal(wb.setSheetFormatDefaults(0, { defaultColWidth: -1 }).ok, false);
      assert.equal(wb.getSheetFormatDefaults(0).defaultColWidth, 10);
      assert.equal(wb.getSheetFormatDefaults(9).status.ok, false);
    });
  });

  test('point geometry: rect, widths, heights, width model and conversions', () => {
    withWorkbook(Module, (wb) => {
      const range = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };
      const rect = wb.getCellRectPt(0, range, 0);
      assert.ok(rect.status.ok);
      assert.equal(rect.x, 0);
      assert.equal(rect.y, 0);
      assert.ok(rect.width > 0 && rect.height > 0);

      const col = wb.getColumnWidthPt(0, 0, 1);
      const row = wb.getRowHeightPt(0, 0);
      assert.ok(col.status.ok && col.value > 0);
      assert.ok(row.status.ok && row.value > 0);

      const wide = wb.getCellRectPt(0, { ...range, lastRow: 0, lastCol: 0 }, 0);
      assert.ok(Math.abs(wide.height - row.value) < 1e-9);

      assert.ok(wb.setColumnHidden(0, 0, 0, true).ok);
      assert.equal(wb.getColumnWidthPt(0, 0, 0).value, 0);
      assert.ok(wb.setColumnHidden(0, 0, 0, false).ok);

      const model = wb.getWidthModel(0, 0);
      assert.ok(model.status.ok);
      assert.equal(typeof model.calibrated, 'boolean');
      assert.equal(typeof model.normalFontName, 'string');
      assert.ok(model.pointsPerChar > 0);
      assert.ok(model.platform.length > 0);

      const pt = wb.columnCharsToPt(0, 0, 10);
      assert.ok(pt.status.ok);
      assert.ok(Math.abs(pt.value - (10 * model.pointsPerChar + model.paddingPt)) < 1e-6);
      const chars = wb.columnPtToChars(0, 0, pt.value);
      assert.ok(Math.abs(chars.value - 10) < 1e-6);
      assert.equal(wb.columnCharsToPt(0, 0, 0).value, 0);
      assert.equal(wb.columnCharsToPt(0, 0, -1).status.ok, false);
      assert.equal(wb.getWidthModel(0, 7).status.ok, false);
      assert.equal(wb.getCellRectPt(0, range, 7).status.ok, false);
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

      const added = wb.addXf({
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
      assert.ok(added.status.ok);
      const xf = wb.getCellXf(added.index);
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
      assert.ok(wb.setNumber(0, 0, 0, 1).ok);
      assert.ok(wb.setNumber(0, 120, 0, 2).ok);
      assert.ok(wb.addSheetRowBreak(0, 10, true).ok);
      assert.ok(wb.setSheetPrintTitles(0, '1:2', '').ok);
      const r = wb.paginate(0);
      assert.ok(r.status.ok);
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
      assert.equal(r.pages[0].firstRow, 0);
      assert.ok(r.pages[0].widthPt > 0);
      assert.equal(r.horizontalBreakManual.length, r.horizontalBreaks.length);
      assert.equal(r.verticalBreakManual.length, r.verticalBreaks.length);
      const manualAt = r.horizontalBreaks.indexOf(10);
      assert.ok(manualAt >= 0, `breaks=${JSON.stringify(r.horizontalBreaks)}`);
      assert.equal(r.horizontalBreakManual[manualAt], true);
      assert.ok(r.horizontalBreakManual.some((m) => m === false));

      const failed = wb.paginate(9);
      assert.equal(failed.status.ok, false);
      assert.deepEqual(failed.pages, []);
      assert.equal(failed.paper.known, false);
    });
  });
}
