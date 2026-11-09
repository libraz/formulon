import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

function makePivot(Workbook) {
  const wb = Workbook.createDefault();
  wb.setProfileId('win-365-ja_JP');
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ range: 'A1:B3', sheet: 'Sheet1' });
  cache.addField('Region');
  cache.addField('Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const record = cache.addRecord();
    cache.setRecordValue(record, 0, region);
    cache.setRecordValue(record, 1, amount);
  }
  const pivot = wb.sheet(0).addPivot({ name: 'Pivot1', cacheId: cache.cacheId, anchor: { row: 0, col: 4 } });
  const region = pivot.addField({ sourceName: 'Region', axis: 0 });
  const amount = pivot.addField({ sourceName: 'Amount', axis: 2 });
  pivot.setRowFieldOrder([region]);
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 });
  return { wb, pivot, region, amount };
}

function withPivot(Workbook, fn) {
  const fixture = makePivot(Workbook);
  try {
    return fn(fixture);
  } finally {
    fixture.wb.dispose();
  }
}

function pivotGrid(wb, pivot) {
  const layout = pivot.getLayout();
  const rows = new Map();
  for (const cell of layout.cells) {
    const value = cell.value;
    const text = typeof value === 'number' ? String(value) : typeof value === 'string' ? value : '';
    if (!rows.has(cell.row)) rows.set(cell.row, new Array(layout.range.lastCol - layout.range.firstCol + 1).fill(''));
    rows.get(cell.row)[cell.col - layout.range.firstCol] = text;
  }
  return [...rows.keys()].sort((a, b) => a - b).map((row) => rows.get(row).join('|'));
}

const BASE_GRID = ['行ラベル|Sum of Amount', 'East|10', 'West|30', '総計|40'];

test('pivot nested integer fields reject narrowing before their C mutations', async () => {
  const mod = await getModule();
  const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY];

  for (const value of invalid) {
    withPivot(mod.Workbook, ({ wb, pivot }) => {
      const before = pivot.listFields().length;
      // `axis` is required: a value that is no whole number is a TypeError, a whole one past uint32 a RangeError.
      assert.throws(
        () => pivot.addField({ sourceName: 'Region', axis: value }),
        Number.isInteger(value) ? RangeError : TypeError,
        `pivotFieldAdd axis=${String(value)}`,
      );
      assert.equal(pivot.listFields().length, before);
    });

    withPivot(mod.Workbook, ({ wb, pivot, amount }) => {
      const before = pivot.listDataFields().length;
      assert.throws(
        () => pivot.addDataField({ name: 'Bad', fieldIndex: value, aggregation: 0 }),
        Number.isInteger(value) ? RangeError : TypeError,
        `pivotDataFieldAdd fieldIndex=${String(value)}`,
      );
      assert.equal(pivot.listDataFields().length, before);

      assert.throws(
        () => pivot.setDataField(0, { name: 'Bad', fieldIndex: amount, aggregation: value }),
        RangeError,
        `pivotDataFieldSet aggregation=${String(value)}`,
      );
      assert.equal(pivot.listDataFields().length, before);
      assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);

      for (const key of ['showAs', 'showAsBaseField', 'showAsBaseItem']) {
        assert.throws(
          () =>
            pivot.setDataField(0, {
              name: 'Bad',
              fieldIndex: amount,
              aggregation: 0,
              [key]: value,
            }),
          RangeError,
          `pivotDataFieldSet ${key}=${String(value)}`,
        );
        assert.equal(pivot.listDataFields().length, before);
        assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);
      }
    });
  }
});

test('pivot nested integer fields reject wrong primitives without changing the pivot', async () => {
  const mod = await getModule();
  const symbol = Symbol('pivot numeric');
  withPivot(mod.Workbook, ({ wb, pivot }) => {
    const fieldCount = pivot.listFields().length;
    assert.throws(() => pivot.addField({ sourceName: 'Region', axis: symbol }), TypeError);
    assert.equal(pivot.listFields().length, fieldCount);

    const dataFieldCount = pivot.listDataFields().length;
    assert.throws(() => pivot.addDataField({ name: 'Bad', fieldIndex: symbol, aggregation: 0 }), TypeError);
    assert.equal(pivot.listDataFields().length, dataFieldCount);

    const filterCount = pivot.listFilters().length;
    assert.throws(
      () =>
        pivot.addFilter({
          axis: 0,
          fieldName: 'Region',
          type: 1,
          value: symbol,
        }),
      TypeError,
    );
    assert.equal(pivot.listFilters().length, filterCount);
  });
});

test('pivot filters validate axis, type, and data-field index before mutation', async () => {
  const mod = await getModule();
  const invalidSpecs = [{ axis: 2 ** 32 }, { type: 2 ** 32 }, { dataFieldIndex: 2 ** 32 }];
  for (const override of invalidSpecs) {
    withPivot(mod.Workbook, ({ wb, pivot }) => {
      const before = pivot.listFilters().length;
      assert.throws(
        () =>
          pivot.addFilter({
            axis: 0,
            fieldName: 'Region',
            type: 1,
            value: 15,
            ...override,
          }),
        RangeError,
        `pivotFilterAdd ${JSON.stringify(override)}`,
      );
      assert.equal(pivot.listFilters().length, before);
      assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);
    });
  }
});

test('pivot order arrays validate every element before replacing the order', async () => {
  const mod = await getModule();
  const invalid = [2 ** 32 + 1, 1.5, Number.NaN, Number.POSITIVE_INFINITY, Symbol('pivot order')];
  for (const value of invalid) {
    withPivot(mod.Workbook, ({ wb, pivot }) => {
      assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);
      assert.throws(
        () => pivot.setRowFieldOrder([value]),
        typeof value === 'symbol' ? TypeError : RangeError,
        `pivotSetRowFieldOrder element=${String(value)}`,
      );
      assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);
    });
  }

  withPivot(mod.Workbook, ({ wb, pivot, region }) => {
    pivot.setRowFieldOrder([region]);
    assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);
  });
});

test('pivot order stops at a throwing array getter before changing the report', async () => {
  const mod = await getModule();
  withPivot(mod.Workbook, ({ wb, pivot }) => {
    const indices = [];
    Object.defineProperty(indices, 0, {
      enumerable: true,
      get() {
        throw new Error('pivot order getter');
      },
    });
    assert.throws(() => pivot.setRowFieldOrder(indices), /pivot order getter/);
    assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);
  });
});

test('pivot nested string getters stop before changing the report', async () => {
  const mod = await getModule();
  withPivot(mod.Workbook, ({ wb, pivot }) => {
    const fieldCount = pivot.listFields().length;
    const spec = { axis: 0 };
    Object.defineProperty(spec, 'sourceName', {
      enumerable: true,
      get() {
        throw new Error('pivot field getter');
      },
    });
    assert.throws(() => pivot.addField(spec), /pivot field getter/);
    assert.equal(pivot.listFields().length, fieldCount);
    assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);
  });
});

test('pivot cache source snapshots each nested getter once before mutation', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const cache = wb.addPivotCache();
    let reads = 0;
    const source = { sheet: 'Sheet1' };
    Object.defineProperty(source, 'range', {
      enumerable: true,
      get() {
        reads += 1;
        if (reads > 1) throw new Error('pivot cache source getter read twice');
        return 'A1:B3';
      },
    });
    cache.setWorksheetSource(source);
    assert.equal(reads, 1);
    assert.deepEqual(JSON.parse(JSON.stringify(cache.getWorksheetSource().range.rect)), {
      firstRow: 0,
      firstCol: 0,
      lastRow: 2,
      lastCol: 1,
    });
  } finally {
    wb.dispose();
  }
});

test('pivot optional strings and nullish numeric defaults keep the valid contract', async () => {
  const mod = await getModule();
  withPivot(mod.Workbook, ({ wb, pivot, amount }) => {
    pivot.addDataField({
      name: 'Count of Amount',
      fieldIndex: amount,
      aggregation: 1,
      numberFormat: null,
      showAs: undefined,
      showAsBaseField: null,
      showAsBaseItem: undefined,
    });
    assert.equal(pivot.listDataFields().length, 2);
  });
});

test('setWorksheetSource(null) clears the source and an empty source is still a present one', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const cache = wb.addPivotCache();
    cache.setWorksheetSource({ range: 'A1:B3', sheet: 'Sheet1' });
    assert.deepEqual(JSON.parse(JSON.stringify(cache.getWorksheetSource())), {
      range: { sheet: 0, rect: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 } },
      name: null,
    });
    cache.setWorksheetSource(null);
    assert.equal(cache.getWorksheetSource(), null);
    cache.setWorksheetSource({});
    assert.deepEqual(JSON.parse(JSON.stringify(cache.getWorksheetSource())), { range: null, name: null });
  } finally {
    wb.dispose();
  }
});
