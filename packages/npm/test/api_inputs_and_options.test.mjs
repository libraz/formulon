// Inputs and options: native values and Dates where a value goes, option records for paging, A1 text for single
// rows and columns, read records that pass back, and a validated createFormulon options bag.
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

const mod = await loadPackage();
const f = await mod.createFormulon();
const { LogLevel, PivotAxis, PivotFilterKind, StatusCode, TableTotalsRowFunction, WorkbookFormat } = f;

/** Runs `body(ws, wb)` on a default workbook, then disposes it. */
function withSheet(body) {
  const wb = f.Workbook.createDefault();
  try {
    return body(wb.sheet(0), wb);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

function expectCode(fn, code, messagePattern) {
  assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, code, `code ${e.code}`);
    if (messagePattern) assert.match(e.message, messagePattern);
    return true;
  });
}

function expectTypeError(fn, message) {
  assert.throws(fn, (e) => {
    assert.ok(e instanceof TypeError, `expected a TypeError, got ${e}`);
    if (message !== undefined) assert.equal(e.message, message);
    return true;
  });
}

async function expectRejectedTypeError(promiseFn, message) {
  await assert.rejects(promiseFn, (e) => {
    assert.ok(e instanceof TypeError, `expected a TypeError, got ${e}`);
    assert.equal(e.message, message);
    return true;
  });
}

const plain = (record) => ({ ...record });

describe('a native value goes where a cell value goes', () => {
  it('formatValue takes a native number', () => {
    withSheet((_ws, wb) => {
      const shown = wb.formatValue(1234.5, '#,##0.00');
      assert.equal(shown.text, '1,234.50');
      assert.equal(shown.displayStatus, f.DisplayStatus.Ok);
      assert.equal(wb.formatValue('abc', '@').text, 'abc');
      assert.equal(wb.formatValue(new Date(2024, 0, 1), 'yyyy-mm-dd').text, '2024-01-01');
    });
  });

  it('formatValue refuses undefined and an array', () => {
    withSheet((_ws, wb) => {
      expectTypeError(() => wb.formatValue(undefined, 'General'), 'Workbook.formatValue value must be a cell value');
      expectTypeError(() => wb.formatValue([1, 2], 'General'), 'Workbook.formatValue value must be a cell value');
    });
  });
});

describe('a date input takes a Date or a civil record', () => {
  it('setPinnedNow with a date-only civil pins midnight', () => {
    withSheet((_ws, wb) => {
      wb.setPinnedNow({ year: 2026, month: 1, day: 2 });
      assert.deepEqual(plain(wb.getPinnedNow()), { year: 2026, month: 1, day: 2, hour: 0, minute: 0, second: 0 });
      expectTypeError(() => wb.setPinnedNow({ year: 2026, month: 1 }));
    });
  });

  it('convertCivilToSerial takes a Date', () => {
    assert.equal(f.convertCivilToSerial(new Date(2024, 2, 1, 12)), 45352.5);
    assert.equal(f.convertCivilToSerial(new Date(2024, 2, 1)), 45352);
    assert.equal(f.convertCivilToSerial(new Date(1904, 0, 2), { date1904: true }), 1);
  });

  it('threaded comment created converts a Date with its UTC getters', () => {
    const previous = process.env.TZ;
    process.env.TZ = 'Asia/Tokyo';
    try {
      const instants = [new Date(Date.UTC(2026, 0, 2, 3, 4, 5)), new Date(Date.UTC(2026, 0, 1, 23, 30, 0))];
      const personId = '{00000000-0000-4000-8000-0000000000aa}';
      withSheet((ws, wb) => {
        wb.addPerson({ id: personId, displayName: 'Alice', userId: 'alice', providerId: 'None' });
        instants.forEach((instant, i) => {
          const id = `{00000000-0000-4000-8000-00000000010${i}}`;
          ws.addThreadedComment({ id, cell: { row: i, col: 0 }, personId, text: 'Hi', created: instant });
          assert.deepEqual(plain(ws.getThreadedComment(id).created), {
            year: instant.getUTCFullYear(),
            month: instant.getUTCMonth() + 1,
            day: instant.getUTCDate(),
            hour: instant.getUTCHours(),
            minute: instant.getUTCMinutes(),
            second: instant.getUTCSeconds(),
          });
        });
        const civilId = '{00000000-0000-4000-8000-000000000200}';
        ws.addThreadedComment({
          id: civilId,
          cell: { row: 5, col: 0 },
          personId,
          text: 'Civil',
          created: { year: 2026, month: 3, day: 4, hour: 5, minute: 6, second: 7 },
        });
        assert.deepEqual(plain(ws.getThreadedComment(civilId).created), {
          year: 2026,
          month: 3,
          day: 4,
          hour: 5,
          minute: 6,
          second: 7,
        });
      });
    } finally {
      if (previous === undefined) delete process.env.TZ;
      else process.env.TZ = previous;
    }
  });
});

describe('the status spelling is enumName', () => {
  it('getStatusName is gone', () => {
    assert.equal(f.getStatusName, undefined);
    assert.equal('getStatusName' in f, false);
    assert.equal(f.enumName(StatusCode, StatusCode.InvalidArgument), 'InvalidArgument');
  });
});

describe('formatReference takes one range and an options record', () => {
  it('formatReference takes an A1 range, a CellRect or a ParsedReference', () => {
    withSheet((_ws, wb) => {
      const rect = { firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 2 };
      assert.equal(wb.formatReference('B2:C3'), 'B2:C3');
      assert.equal(wb.formatReference(rect), 'B2:C3');
      assert.equal(wb.formatReference('B2:C3', { absolute: true }), '$B$2:$C$3');
      assert.equal(wb.formatReference(rect, { sheetName: 'Sheet1' }), 'Sheet1!B2:C3');
      assert.equal(wb.formatReference('B2:C3', { sheetName: null }), 'B2:C3');
      assert.equal(wb.formatReference('B2:C3', { sheetName: '' }), 'B2:C3');

      const parsed = wb.parseReference('Sheet1!B2:C3');
      assert.equal(wb.formatReference(parsed), 'Sheet1!B2:C3');
      assert.equal(wb.formatReference(parsed, { absolute: true }), 'Sheet1!$B$2:$C$3');
      assert.equal(wb.formatReference(parsed, { sheetName: 'Other' }), 'Other!B2:C3');
    });
  });
});

describe('the log sink receives records', () => {
  it('setLogSink delivers a parsed LogEvent', async () => {
    const records = [];
    f.setLogSink((record) => records.push(record));
    f.setLogMinLevel(LogLevel.Warn);
    try {
      withSheet((ws, wb) => {
        // A structured reference the XLSB encoder cannot lower makes the writer emit a warn record.
        ws.setFormula(0, 0, '=SUM(T[C])');
        wb.recalc();
        wb.save(WorkbookFormat.Xlsb);
      });
      const deadline = Date.now() + 1000;
      while (records.length === 0 && Date.now() < deadline) {
        await new Promise((resolve) => setTimeout(resolve, 5));
      }
      assert.ok(records.length > 0, 'expected at least one record');
      for (const record of records) {
        assert.equal(ArrayBuffer.isView(record), false, 'a record is parsed data, not bytes');
        assert.equal(typeof record, 'object');
        assert.equal(typeof record.event, 'string');
        assert.equal(record.level, LogLevel.Warn);
      }
      assert.ok(records.some((r) => r.event === 'xlsb.writer.formula_downgraded'));
      expectTypeError(() => f.setLogSink(5), 'Formulon.setLogSink sink must be a function');
    } finally {
      f.setLogMinLevel(LogLevel.Off);
      f.setLogSink(null);
    }
  });
});

describe('createFormulon validates its options', () => {
  it('createFormulon rejects an unknown key', async () => {
    await expectRejectedTypeError(() => mod.createFormulon({ bogus: 1 }), 'createFormulon options has no field bogus');
    const again = await mod.createFormulon({ print: null, printErr: null, locateFile: null, wasm: null });
    assert.equal(typeof again.Workbook, 'function');
  });

  it('createFormulon rejects a non-object', async () => {
    await expectRejectedTypeError(() => mod.createFormulon('str'), 'createFormulon options must be an object');
    await expectRejectedTypeError(() => mod.createFormulon(42), 'createFormulon options must be an object');
  });

  it('createFormulon rejects a non-function print and locateFile', async () => {
    await expectRejectedTypeError(
      () => mod.createFormulon({ print: 5 }),
      'createFormulon options.print must be a function',
    );
    await expectRejectedTypeError(
      () => mod.createFormulon({ printErr: 'x' }),
      'createFormulon options.printErr must be a function',
    );
    await expectRejectedTypeError(
      () => mod.createFormulon({ locateFile: 'x' }),
      'createFormulon options.locateFile must be a function',
    );
    await expectRejectedTypeError(
      () => mod.createFormulon({ wasm: 42 }),
      'createFormulon options.wasm must be a BufferSource or a WebAssembly.Module',
    );
    await expectRejectedTypeError(
      () => mod.createFormulon({ wasm: new Uint8Array(0) }),
      'createFormulon options.wasm must not be empty',
    );
  });
});

describe('view and format defaults spell column out', () => {
  it('view and format defaults spell column out', () => {
    withSheet((ws) => {
      const view = ws.getView();
      assert.equal(view.showRowColumnHeaders, true);
      assert.equal('showRowColHeaders' in view, false);
      ws.updateView({ showRowColumnHeaders: false });
      assert.equal(ws.getView().showRowColumnHeaders, false);
      expectTypeError(() => ws.updateView({ showRowColHeaders: true }));

      ws.updateFormatDefaults({ defaultColumnWidthChars: 12, baseColumnWidthChars: 9 });
      const defaults = ws.getFormatDefaults();
      assert.equal(defaults.defaultColumnWidthChars, 12);
      assert.equal(defaults.baseColumnWidthChars, 9);
      assert.equal(defaults.effectiveDefaultColumnWidthChars, 12);
      for (const old of ['defaultColWidthChars', 'baseColWidthChars', 'effectiveDefaultColWidthChars']) {
        assert.equal(old in defaults, false, `${old} is gone`);
      }
      expectTypeError(() => ws.updateFormatDefaults({ defaultColWidthChars: 5 }));
      ws.updateFormatDefaults({ defaultColumnWidthChars: null });
      assert.equal(ws.getFormatDefaults().defaultColumnWidthChars, null);
    });
  });
});

describe('paging takes an options record', () => {
  const fill = (ws) => {
    for (let row = 0; row < 5; row += 1) ws.setNumber(row, 0, row + 1);
  };

  it('getCellsInRange pages with options', () => {
    withSheet((ws) => {
      fill(ws);
      const first = ws.getCellsInRange('A1:A5', { limit: 2 });
      assert.deepEqual(
        first.cells.map((c) => c.row),
        [0, 1],
      );
      assert.notEqual(first.nextCursor, null);
      const second = ws.getCellsInRange('A1:A5', { cursor: first.nextCursor, limit: 2 });
      assert.deepEqual(
        second.cells.map((c) => c.row),
        [2, 3],
      );
      const third = ws.getCellsInRange('A1:A5', { cursor: second.nextCursor, limit: 2 });
      assert.deepEqual(
        third.cells.map((c) => c.row),
        [4],
      );
      assert.equal(third.nextCursor, null);

      for (const options of [{}, { limit: null }, { limit: 0 }, { cursor: null }]) {
        assert.equal(ws.getCellsInRange('A1:A5', options).cells.length, 5);
      }
      assert.equal(ws.getCellsInRange(null, { limit: 3 }).cells.length, 3);
      assert.equal(ws.getCellsInRange().cells.length, 5);
      assert.equal(Array.isArray(ws.getInvalidCells({ limit: 1 }).cells), true);
      assert.equal(ws.getInvalidCells().nextCursor, null);

      expectTypeError(
        () => ws.getCellsInRange('A1:A5', { bogus: 1 }),
        'Sheet.getCellsInRange options has no field bogus',
      );
      expectCode(
        () => ws.getCellsInRange('A1:A5', { limit: -1 }),
        StatusCode.InvalidArgument,
        /^Sheet\.getCellsInRange options\.limit must not be negative$/,
      );
    });
  });

  it('a positional cursor is a TypeError', () => {
    withSheet((ws) => {
      fill(ws);
      expectTypeError(() => ws.getCellsInRange(null, null, 100), 'Sheet.getCellsInRange takes at most 2 arguments');
      expectTypeError(() => ws.getCellsInRange('A1:A5', 2));
      expectTypeError(() => ws.getInvalidCells(null, 5), 'Sheet.getInvalidCells takes at most 1 argument');
      expectTypeError(() => ws.getInvalidCells(3));
    });
  });
});

describe('a single row or column takes a number or an A1 string', () => {
  it('single-axis getters take A1', () => {
    withSheet((ws) => {
      ws.setColumnWidth(1, 1, 20);
      ws.setRowHeight(2, 2, 30);
      assert.equal(ws.getColumnWidthPt('B'), ws.getColumnWidthPt(1));
      assert.notEqual(ws.getColumnWidthPt('B'), ws.getColumnWidthPt('A'));
      assert.equal(ws.getColumnWidthPt('B', f.GeometryMode.Print), ws.getColumnWidthPt(1, f.GeometryMode.Print));
      assert.equal(ws.getRowHeightPt('3'), ws.getRowHeightPt(2));
      assert.notEqual(ws.getRowHeightPt('3'), ws.getRowHeightPt('1'));
      assert.equal(ws.ooxml.getColumnXfIndex('B'), ws.ooxml.getColumnXfIndex(1));
      assert.equal(ws.ooxml.getRowXfIndex('3'), ws.ooxml.getRowXfIndex(2));
    });
  });

  it('a multi-column span is InvalidArgument', () => {
    withSheet((ws) => {
      expectCode(
        () => ws.getColumnWidthPt('B:D'),
        StatusCode.InvalidArgument,
        /^Sheet\.getColumnWidthPt a1 must name one column$/,
      );
      expectCode(
        () => ws.getRowHeightPt('2:4'),
        StatusCode.InvalidArgument,
        /^Sheet\.getRowHeightPt a1 must name one row$/,
      );
      expectCode(
        () => ws.ooxml.getColumnXfIndex('B:D'),
        StatusCode.InvalidArgument,
        /^SheetOoxml\.getColumnXfIndex .* must name one column$/,
      );
      expectCode(
        () => ws.ooxml.getRowXfIndex('2:4'),
        StatusCode.InvalidArgument,
        /^SheetOoxml\.getRowXfIndex .* must name one row$/,
      );
    });
  });
});

describe('a read record passes back to its setter', () => {
  it('a read pivot filter passes back to addFilter', () => {
    withSheet((ws, wb) => {
      ws.setText('A1', 'Region');
      ws.setText('B1', 'Amount');
      [
        ['East', 10],
        ['West', 30],
      ].forEach(([region, amount], i) => {
        ws.setText(i + 1, 0, region);
        ws.setNumber(i + 1, 1, amount);
      });
      const cache = wb.addPivotCache();
      cache.setWorksheetSource({ range: 'A1:B3', sheet: 'Sheet1' });
      cache.addField('Region');
      cache.addField('Amount');
      for (const item of ['East', 'West']) cache.addFieldSharedItem(0, item);
      [
        [0, 10],
        [1, 30],
      ].forEach(([item, amount]) => {
        const record = cache.addRecord();
        cache.setRecordValue(record, 0, item);
        cache.setRecordValue(record, 1, amount);
      });
      const pivot = ws.addPivot({ name: 'Pivot1', cacheId: cache.cacheId, anchor: 'F1' });
      pivot.setFieldAxis('Region', PivotAxis.Row);
      pivot.addFilter({ fieldIndex: 'Region', kind: PivotFilterKind.LabelContains, value: 'East' });
      pivot.addFilter({ fieldIndex: 'Region', kind: PivotFilterKind.LabelBetween, value: 'A', valueHigh: 'M' });

      const read = pivot.listFilters().map(plain);
      assert.equal(read.length, 2);
      assert.equal(read[0].valueHigh, null);
      for (const filter of read) pivot.addFilter(filter);
      const after = pivot.listFilters().map(plain);
      assert.equal(after.length, 4);
      assert.deepEqual(after.slice(2), read);
    });
  });

  it('a column patch takes a read column and ignores its name', () => {
    withSheet((ws) => {
      ws.setText('A1', 'Id');
      ws.setText('B1', 'Amount');
      ws.setNumber(1, 0, 1);
      ws.setNumber(1, 1, 5);
      const table = ws.addTable({ range: 'A1:B2', name: 'Sales' });
      const columns = table.getProperties().columns;
      table.update({ columns: [columns[0]] });
      table.update({
        columns: [{}, { ...columns[1], totalsRowFunction: TableTotalsRowFunction.Sum }],
      });
      const after = table.getProperties().columns;
      assert.deepEqual(
        after.map((c) => c.name),
        ['Id', 'Amount'],
      );
      assert.equal(after[0].totalsRowFunction, columns[0].totalsRowFunction);
      assert.equal(after[1].totalsRowFunction, TableTotalsRowFunction.Sum);
      table.update({ columns: [{}, { ...after[1], name: 'Renamed' }] });
      assert.equal(table.getProperties().columns[1].name, 'Amount');
      assert.equal(table.getProperties().columns[1].totalsRowFunction, TableTotalsRowFunction.Sum);
    });
  });

  it('setRowBreak takes a read RowBreak back and ignores its row', () => {
    withSheet((ws) => {
      ws.setRowBreak(5, { first: 2, last: 9, manual: true });
      const read = ws.listRowBreaks()[0];
      ws.setRowBreak(7, read);
      const byRow = Object.fromEntries(ws.listRowBreaks().map((b) => [b.row, plain(b)]));
      assert.deepEqual(Object.keys(byRow).sort(), ['5', '7']);
      assert.deepEqual(byRow[7], { row: 7, first: 2, last: 9, manual: true });
      assert.deepEqual(byRow[5], { row: 5, first: 2, last: 9, manual: true });

      ws.setColumnBreak(3, { first: 1, last: 4, manual: false });
      const column = ws.listColumnBreaks()[0];
      ws.setColumnBreak(6, column);
      const columns = Object.fromEntries(ws.listColumnBreaks().map((b) => [b.col, plain(b)]));
      assert.deepEqual(columns[6], { col: 6, first: 1, last: 4, manual: false });
      assert.deepEqual(columns[3], { col: 3, first: 1, last: 4, manual: false });
    });
  });
});

describe('phonetic runs are checked field by field', () => {
  it('setCellPhoneticRuns refuses unknown and missing run fields', () => {
    withSheet((ws) => {
      ws.setText('A1', '漢字');
      expectTypeError(
        () => ws.setCellPhoneticRuns('A1', [{ start: 0, end: 2, text: 'かんじ', sb: 1 }]),
        'Sheet.setCellPhoneticRuns runs[0] has no field sb',
      );
      expectTypeError(
        () => ws.setCellPhoneticRuns('A1', [{ start: 0, text: 'かんじ' }]),
        'Sheet.setCellPhoneticRuns runs[0].end is required',
      );
      expectTypeError(
        () =>
          ws.setCellPhoneticRuns('A1', [
            { start: 0, end: 2, text: 'かんじ' },
            { end: 2, text: 'x' },
          ]),
        'Sheet.setCellPhoneticRuns runs[1].start is required',
      );
      expectTypeError(() => ws.setCellPhoneticRuns('A1', { start: 0, end: 2, text: 'x' }));
      ws.setCellPhoneticRuns('A1', [{ start: 0, end: 2, text: 'かんじ' }]);
      ws.setCellPhoneticRuns('A1', null);
    });
  });
});
