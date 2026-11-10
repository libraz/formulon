// The argument error model: the class and status code of a refused argument by what the argument is.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

async function withWorkbook(fn) {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

// Asserts that `fn` fails with a FormulonError carrying `code`.
function rejectsWith(fn, code, label) {
  assert.throws(
    fn,
    (e) => {
      assert.equal(e.name, 'FormulonError', `${label}: expected a FormulonError, got ${e}`);
      assert.equal(e.code, code, `${label}: code ${e.code}`);
      return true;
    },
    label,
  );
}

function withTable(wb, fn) {
  const ws = wb.sheet(0);
  ws.setValues('A1', [
    ['Item', 'Qty'],
    ['a', 1],
    ['b', 2],
  ]);
  return fn(ws, ws.addTable({ name: 'Sales', range: 'A1:B3' }));
}

test('a sheet key: negative is SheetIndexOutOfRange, fractional or beyond 32 bits is RangeError, a name is NotFound', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    rejectsWith(() => wb.sheet(-1), StatusCode.SheetIndexOutOfRange, 'negative index');
    rejectsWith(() => wb.sheet(5), StatusCode.SheetIndexOutOfRange, 'past the end');
    rejectsWith(() => wb.sheet('x'), StatusCode.NotFound, 'unknown name');
    assert.throws(() => wb.sheet(1.5), RangeError);
    assert.throws(() => wb.sheet(Number.NaN), RangeError);
    assert.throws(() => wb.sheet(Number.POSITIVE_INFINITY), RangeError);
    assert.throws(() => wb.sheet(2 ** 32), RangeError);
    assert.throws(() => wb.sheet(true), TypeError);
    assert.throws(() => wb.sheet(null), TypeError);
    assert.equal(wb.sheet(0).index, 0);
  });
});

test('a cell coordinate: negative or beyond the grid is InvalidArgument, fractional or beyond 32 bits is RangeError', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    rejectsWith(() => ws.getValue(-1, 0), StatusCode.InvalidArgument, 'negative row');
    rejectsWith(() => ws.getValue(0, -1), StatusCode.InvalidArgument, 'negative column');
    rejectsWith(() => ws.getValue(0, 16384), StatusCode.InvalidArgument, 'column past the grid');
    rejectsWith(() => ws.getValue(2 ** 31, 0), StatusCode.InvalidArgument, 'row at 2^31');
    rejectsWith(() => ws.setNumber(-1, 0, 1), StatusCode.InvalidArgument, 'negative row on a write');
    assert.throws(() => ws.getValue(1.5, 0), RangeError);
    assert.throws(() => ws.getValue(0, Number.NaN), RangeError);
    assert.throws(() => ws.getValue(2 ** 32, 0), RangeError);
    assert.throws(() => ws.getValue('1', 0), TypeError);
  });
});

test('the fields of a rectangle follow the coordinate rule', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    for (const key of ['firstRow', 'firstCol', 'lastRow', 'lastCol']) {
      const rect = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };
      rejectsWith(() => ws.addMerge({ ...rect, [key]: -1 }), StatusCode.InvalidArgument, `${key} negative`);
      assert.throws(() => ws.addMerge({ ...rect, [key]: 1.5 }), RangeError, `${key} fractional`);
    }
    rejectsWith(
      () => ws.getCellRectPt({ firstRow: -1, firstCol: 0, lastRow: 1, lastCol: 1 }),
      StatusCode.InvalidArgument,
      'negative rectangle field on a read',
    );
  });
});

test('an element position: negative or past the end is IndexOutOfRange, fractional is RangeError', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    withTable(wb, (ws, table) => {
      rejectsWith(() => ws.table(-1), StatusCode.IndexOutOfRange, 'negative table index');
      rejectsWith(() => ws.table(5), StatusCode.IndexOutOfRange, 'table index past the end');
      rejectsWith(() => table.renameColumn(-1, 'x'), StatusCode.IndexOutOfRange, 'negative column index');
      rejectsWith(() => table.renameColumn(99, 'x'), StatusCode.IndexOutOfRange, 'column index past the end');
      rejectsWith(() => ws.removeValidation(-1), StatusCode.IndexOutOfRange, 'negative validation index');
      assert.throws(() => ws.table(1.5), RangeError);
      assert.throws(() => table.renameColumn(1.5, 'x'), RangeError);
      assert.throws(() => table.renameColumn(2 ** 32, 'x'), RangeError);
      assert.throws(() => table.renameColumn('a', 'x'), TypeError);
    });
  });
});

test('a font index follows the position rule', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    rejectsWith(() => wb.ooxml.getFont(-1), StatusCode.IndexOutOfRange, 'negative font index');
    rejectsWith(() => wb.ooxml.getFont(999), StatusCode.IndexOutOfRange, 'font index past the end');
    assert.throws(() => wb.ooxml.getFont(1.5), RangeError);
  });
});

test('a quantity: negative is InvalidArgument, fractional is RangeError', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    rejectsWith(() => ws.insertRows(0, -1), StatusCode.InvalidArgument, 'negative row count');
    assert.throws(() => ws.insertRows(0, 1.5), RangeError);
  });
});

test('an integer that is not a GeometryMode is InvalidArgument, a non-number is TypeError', async () => {
  await withWorkbook((wb, { StatusCode, GeometryMode }) => {
    const ws = wb.sheet(0);
    rejectsWith(() => ws.getColumnWidthPt(0, 5), StatusCode.InvalidArgument, 'getColumnWidthPt');
    rejectsWith(() => ws.getCellRectPt('A1', 5), StatusCode.InvalidArgument, 'getCellRectPt');
    rejectsWith(() => ws.getWidthModel(5), StatusCode.InvalidArgument, 'getWidthModel');
    assert.throws(() => ws.getColumnWidthPt(0, '1'), TypeError);
    assert.equal(typeof ws.getColumnWidthPt(0, GeometryMode.Print), 'number');
  });
});

test('a table column total function that is not a member is InvalidArgument, a wrong type is TypeError', async () => {
  await withWorkbook((wb, { StatusCode, TableTotalsRowFunction }) => {
    withTable(wb, (_ws, table) => {
      rejectsWith(
        () => table.update({ columns: [{}, { totalsRowFunction: 42 }] }),
        StatusCode.InvalidArgument,
        'ordinal 42',
      );
      assert.throws(() => table.update({ columns: [{}, { totalsRowFunction: 'sum' }] }), TypeError);
      table.update({ columns: [{}, { totalsRowFunction: TableTotalsRowFunction.Sum }] });
    });
  });
});

test('byte inputs accept Uint8Array, ArrayBuffer and views, and refuse anything else with TypeError', async () => {
  await withWorkbook((wb, f) => {
    wb.sheet(0).setValue('A1', 'kept');
    const bytes = wb.save();
    const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
    for (const [label, input] of [
      ['Uint8Array', bytes],
      ['ArrayBuffer', buffer],
      ['DataView', new DataView(buffer)],
    ]) {
      const loaded = f.Workbook.load(input);
      try {
        assert.equal(loaded.sheet(0).getValue('A1'), 'kept', label);
      } finally {
        loaded.dispose();
      }
    }
    assert.throws(() => f.Workbook.load('in.xlsx'), TypeError);
    assert.throws(() => f.Workbook.load(null), TypeError);
    assert.throws(() => f.Workbook.load(undefined), TypeError);
    assert.throws(() => f.Workbook.load({}), TypeError);
    assert.throws(() => f.Workbook.load([1, 2, 3]), TypeError);
    rejectsWith(() => f.Workbook.load(new ArrayBuffer(0)), f.StatusCode.InvalidArgument, 'empty input');
    rejectsWith(() => f.Workbook.load(new Uint8Array([1, 2, 3])), f.StatusCode.IoZipCorrupt, 'not a ZIP');
    assert.throws(() => wb.formatValue(undefined, '0.0'), TypeError);
  });
});

test('recalcParallel: threadCount is optional, a non-number is TypeError, a fraction or non-finite RangeError, a negative or above 8 InvalidArgument', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    assert.equal(typeof wb.recalcParallel(), 'object');
    assert.equal(typeof wb.recalcParallel(undefined), 'object');
    assert.equal(typeof wb.recalcParallel(0), 'object');
    assert.equal(typeof wb.recalcParallel(8), 'object');
    for (const bad of [1.5, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
      assert.throws(() => wb.recalcParallel(bad), RangeError, String(bad));
    }
    for (const bad of ['2', null, true, {}]) {
      assert.throws(() => wb.recalcParallel(bad), TypeError, String(bad));
    }
    rejectsWith(() => wb.recalcParallel(-1), StatusCode.InvalidArgument, 'negative');
    rejectsWith(() => wb.recalcParallel(9), StatusCode.InvalidArgument, 'above the cap');
  });
});

test('ragged rows of setValues are InvalidArgument, a non-sequence grid or row is TypeError', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    rejectsWith(() => ws.setValues('A1', [[1, 2], [3]]), StatusCode.InvalidArgument, 'ragged rows');
    rejectsWith(
      () =>
        ws.setValues(1048575, 16383, [
          [1, 2],
          [3, 4],
        ]),
      StatusCode.InvalidArgument,
      'past the grid',
    );
    assert.throws(() => ws.setValues('A1', 'x'), TypeError);
    assert.throws(() => ws.setValues('A1', [1]), TypeError);
  });
});

test('localizeFormula and canonicalizeFormula report a syntax error with offsets into the text passed', async () => {
  const f = await getModule();
  assert.equal(f.canonicalizeFormula('=SUMME(1;2)', 'mac-365-de_DE'), '=SUM(1,2)');
  assert.equal(f.localizeFormula('=SUM(1,2)', 'mac-365-de_DE'), '=SUMME(1;2)');

  const syntaxError = (fn) => {
    try {
      fn();
    } catch (e) {
      assert.equal(e.name, 'FormulonError');
      assert.ok(e.code >= 1000 && e.code < 2000, `parser status expected, got ${e.code}`);
      assert.ok(e.parseError, 'parseError is attached');
      return e;
    }
    assert.fail('a syntax error is expected');
  };

  const unclosed = syntaxError(() => f.canonicalizeFormula('=SUMME(1;', 'mac-365-de_DE'));
  assert.equal(unclosed.parseError.end, 9);
  assert.equal(unclosed.parseError.line, 1);
  assert.equal(unclosed.parseError.column, 10);

  const text = '=WENN(A1>0;SUMME(B1:B3;"x);0)';
  const unterminated = syntaxError(() => f.canonicalizeFormula(text, 'mac-365-de_DE'));
  assert.equal(unterminated.parseError.start, 23);
  assert.equal(text[unterminated.parseError.start], '"');

  const canonical = syntaxError(() => f.localizeFormula('=SUM(1,', 'mac-365-de_DE'));
  assert.equal(canonical.parseError.start, 7);
});
