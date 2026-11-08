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
    { start: 0, end: 2, text: 'トウキョウ' },
    { start: 2, end: 3, text: 'ト' },
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
      wb.sheet(0).setText(0, 0, '東京都');
      wb.sheet(0).setCellPhoneticRuns(0, 0, validRuns());
      const before = wb.sheet(0).getCellPhoneticRuns(0, 0);
      for (const key of ['start', 'end']) {
        assert.throws(
          () => wb.sheet(0).setCellPhoneticRuns(0, 0, [{ ...validRuns()[0], [key]: value }, validRuns()[1]]),
          RangeError,
          `setCellPhoneticRuns ${key}=${String(value)}`,
        );
        assert.deepEqual(wb.sheet(0).getCellPhoneticRuns(0, 0), before);
      }
    });
  }
});

test('cell phonetic run text rejects coercible non-strings like the WASM surface', async () => {
  const mod = await getModule();
  withWorkbook(mod.Workbook, (wb) => {
    wb.sheet(0).setText(0, 0, '東京都');
    wb.sheet(0).setCellPhoneticRuns(0, 0, validRuns());
    const before = wb.sheet(0).getCellPhoneticRuns(0, 0);
    for (const value of [42, true, { toString: () => 'ト' }]) {
      assert.throws(
        () => wb.sheet(0).setCellPhoneticRuns(0, 0, [{ ...validRuns()[0], text: value }, validRuns()[1]]),
        TypeError,
        `setCellPhoneticRuns text=${String(value)}`,
      );
      assert.deepEqual(wb.sheet(0).getCellPhoneticRuns(0, 0), before);
    }
    const bytes = new TextEncoder().encode('トウキョウ');
    wb.sheet(0).setCellPhoneticRuns(0, 0, [{ ...validRuns()[0], text: bytes }, validRuns()[1]]);
    assert.deepEqual(wb.sheet(0).getCellPhoneticRuns(0, 0), before);
  });
});

test('cell phonetic properties reject narrowing before changing the stored properties', async () => {
  const mod = await getModule();
  const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY];
  for (const value of invalid) {
    withWorkbook(mod.Workbook, (wb) => {
      wb.sheet(0).setText(0, 0, '大阪');
      wb.sheet(0).setCellPhoneticProperties(0, 0, { fontIndex: 3, type: 2, alignment: 2 });
      const before = wb.sheet(0).getCellPhoneticProperties(0, 0);
      for (const key of ['fontIndex', 'type', 'alignment']) {
        assert.throws(
          () => wb.sheet(0).setCellPhoneticProperties(0, 0, { fontIndex: 3, type: 2, alignment: 2, [key]: value }),
          RangeError,
          `setCellPhoneticProperties ${key}=${String(value)}`,
        );
        const after = wb.sheet(0).getCellPhoneticProperties(0, 0);
        assert.deepEqual(
          { fontIndex: after.fontIndex, type: after.type, alignment: after.alignment },
          { fontIndex: before.fontIndex, type: before.type, alignment: before.alignment },
        );
      }
    });
  }
});

test('cell nested value fields of the wrong shape are rejected before formatValue reads them', async () => {
  const mod = await getModule();
  const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY];
  for (const value of invalid) {
    withWorkbook(mod.Workbook, (wb) => {
      assert.throws(
        () => wb.formatValue({ kind: value, number: 1 }, 'General'),
        TypeError,
        `formatValue kind=${String(value)}`,
      );
      assert.throws(
        () => wb.formatValue({ kind: 2, boolean: value }, 'General'),
        TypeError,
        `formatValue boolean=${String(value)}`,
      );
      assert.throws(
        () => wb.formatValue({ kind: 4, errorCode: value }, 'General'),
        TypeError,
        `formatValue errorCode=${String(value)}`,
      );
    });
  }
});

test('cell nested value symbols are rejected before formatValue', async () => {
  const mod = await getModule();
  withWorkbook(mod.Workbook, (wb) => {
    const symbol = Symbol('cell value integer');
    assert.throws(() => wb.formatValue({ kind: symbol, number: 1 }, 'General'), TypeError);
    assert.throws(() => wb.formatValue({ kind: 2, boolean: symbol }, 'General'), TypeError);
    assert.throws(() => wb.formatValue({ kind: 4, errorCode: symbol }, 'General'), TypeError);
  });
});

test('cell range integers reject narrowing before getCellsInRange touches the workbook', async () => {
  const mod = await getModule();
  const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY];
  for (const value of invalid) {
    withWorkbook(mod.Workbook, (wb) => {
      wb.sheet(0).setNumber(0, 0, 7);
      for (const key of ['firstRow', 'firstCol', 'lastRow', 'lastCol']) {
        assert.throws(
          () => wb.sheet(0).getCellsInRange({ ...validRange(), [key]: value }),
          value === 2 ** 32 ? RangeError : TypeError,
          `getCellsInRange ${key}=${String(value)}`,
        );
      }
    });
  }
});

test('cell range getters stop before getCellsInRange calls the C API', async () => {
  const mod = await getModule();
  withWorkbook(mod.Workbook, (wb) => {
    wb.sheet(0).setNumber(0, 0, 7);
    const range = validRange();
    Object.defineProperty(range, 'firstRow', {
      enumerable: true,
      get() {
        throw new Error('cell range getter');
      },
    });
    assert.throws(() => wb.sheet(0).getCellsInRange(range), /cell range getter/);
  });
});
