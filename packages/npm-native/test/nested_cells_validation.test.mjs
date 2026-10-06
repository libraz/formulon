import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

function withWorkbook(Workbook, fn) {
  const wb = Workbook.createDefault();
  try {
    return fn(wb);
  } finally {
    wb.dispose();
  }
}

function validRuns() {
  return [
    { sb: 0, eb: 2, text: 'トウキョウ' },
    { sb: 2, eb: 3, text: 'ト' },
  ];
}

function validRange() {
  return { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 4 };
}

test('cell phonetic run integers reject narrowing before replacing existing runs', async () => {
  const mod = await getModule();
  const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY];
  for (const value of invalid) {
    withWorkbook(mod.Workbook, (wb) => {
      assert.ok(wb.setText(0, 0, 0, '東京都').ok);
      assert.ok(wb.setCellPhoneticRuns(0, 0, 0, validRuns()).ok);
      const before = wb.getCellPhoneticRuns(0, 0, 0).runs;
      for (const key of ['sb', 'eb']) {
        assert.throws(
          () => wb.setCellPhoneticRuns(0, 0, 0, [{ ...validRuns()[0], [key]: value }, validRuns()[1]]),
          RangeError,
          `setCellPhoneticRuns ${key}=${String(value)}`,
        );
        assert.deepEqual(wb.getCellPhoneticRuns(0, 0, 0).runs, before);
      }
    });
  }
});

test('cell phonetic properties reject narrowing before changing the stored properties', async () => {
  const mod = await getModule();
  const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY];
  for (const value of invalid) {
    withWorkbook(mod.Workbook, (wb) => {
      assert.ok(wb.setText(0, 0, 0, '大阪').ok);
      assert.ok(wb.setCellPhoneticProperties(0, 0, 0, { fontId: 3, type: 2, alignment: 2 }).ok);
      const before = wb.getCellPhoneticProperties(0, 0, 0);
      for (const key of ['fontId', 'type', 'alignment']) {
        assert.throws(
          () => wb.setCellPhoneticProperties(0, 0, 0, { fontId: 3, type: 2, alignment: 2, [key]: value }),
          RangeError,
          `setCellPhoneticProperties ${key}=${String(value)}`,
        );
        const after = wb.getCellPhoneticProperties(0, 0, 0);
        assert.deepEqual(
          { fontId: after.fontId, type: after.type, alignment: after.alignment },
          { fontId: before.fontId, type: before.type, alignment: before.alignment },
        );
      }
    });
  }
});

test('cell nested value integers reject narrowing before formatValue reads them', async () => {
  const mod = await getModule();
  const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY];
  for (const value of invalid) {
    withWorkbook(mod.Workbook, (wb) => {
      assert.throws(
        () => wb.formatValue({ kind: value, number: 1, boolean: 0, text: '', errorCode: 0 }, 'General'),
        RangeError,
        `formatValue kind=${String(value)}`,
      );
      assert.throws(
        () => wb.formatValue({ kind: 2, number: 0, boolean: value, text: '', errorCode: 0 }, 'General'),
        RangeError,
        `formatValue boolean=${String(value)}`,
      );
      assert.throws(
        () => wb.formatValue({ kind: 4, number: 0, boolean: 0, text: '', errorCode: value }, 'General'),
        RangeError,
        `formatValue errorCode=${String(value)}`,
      );
    });
  }
});

test('cell nested value symbols are rejected before formatValue', async () => {
  const mod = await getModule();
  withWorkbook(mod.Workbook, (wb) => {
    const symbol = Symbol('cell value integer');
    assert.throws(
      () => wb.formatValue({ kind: symbol, number: 1, boolean: 0, text: '', errorCode: 0 }, 'General'),
      TypeError,
    );
    assert.throws(
      () => wb.formatValue({ kind: 2, number: 0, boolean: symbol, text: '', errorCode: 0 }, 'General'),
      TypeError,
    );
    assert.throws(
      () => wb.formatValue({ kind: 4, number: 0, boolean: 0, text: '', errorCode: symbol }, 'General'),
      TypeError,
    );
  });
});

test('cell range integers reject narrowing before getCellsInRange touches the workbook', async () => {
  const mod = await getModule();
  const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY];
  for (const value of invalid) {
    withWorkbook(mod.Workbook, (wb) => {
      assert.ok(wb.setNumber(0, 0, 0, 7).ok);
      for (const key of ['firstRow', 'firstCol', 'lastRow', 'lastCol']) {
        assert.throws(
          () => wb.getCellsInRange(0, { ...validRange(), [key]: value }),
          RangeError,
          `getCellsInRange ${key}=${String(value)}`,
        );
      }
    });
  }
});

test('cell range getters stop before getCellsInRange calls the C API', async () => {
  const mod = await getModule();
  withWorkbook(mod.Workbook, (wb) => {
    assert.ok(wb.setNumber(0, 0, 0, 7).ok);
    const range = validRange();
    Object.defineProperty(range, 'firstRow', {
      enumerable: true,
      get() {
        throw new Error('cell range getter');
      },
    });
    assert.throws(() => wb.getCellsInRange(0, range), /cell range getter/);
  });
});
