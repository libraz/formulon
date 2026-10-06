import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

function statusOf(result) {
  return typeof result?.ok === 'boolean' ? result : result?.status;
}

function must(result, what) {
  const status = statusOf(result);
  assert.ok(status?.ok, `${what}: ${JSON.stringify(status)}`);
  return result;
}

function makePivot(Workbook) {
  const wb = Workbook.createDefault();
  const cache = must(wb.pivotCacheCreate(0), 'pivotCacheCreate').index;
  must(
    wb.pivotCacheSetWorksheetSource(cache, { present: true, ref: 'A1:B3', sheet: 'Sheet1' }),
    'pivotCacheSetWorksheetSource',
  );
  must(wb.pivotCacheFieldAdd(cache, 'Region'), 'pivotCacheFieldAdd Region');
  must(wb.pivotCacheFieldAdd(cache, 'Amount'), 'pivotCacheFieldAdd Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const record = must(wb.pivotCacheRecordAdd(cache), 'pivotCacheRecordAdd').index;
    must(wb.pivotCacheRecordSetText(cache, record, 0, region), 'pivotCacheRecordSetText');
    must(wb.pivotCacheRecordSetNumber(cache, record, 1, amount), 'pivotCacheRecordSetNumber');
  }
  const pivot = must(wb.pivotCreate(0, 'Pivot1', cache, 0, 4), 'pivotCreate').index;
  const region = must(wb.pivotFieldAdd(0, pivot, { sourceName: 'Region', axis: 0 }), 'region field').index;
  const amount = must(wb.pivotFieldAdd(0, pivot, { sourceName: 'Amount', axis: 2 }), 'amount field').index;
  must(wb.pivotSetRowFieldOrder(0, pivot, [region]), 'pivotSetRowFieldOrder');
  must(
    wb.pivotDataFieldAdd(0, pivot, { name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 }),
    'pivotDataFieldAdd',
  );
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
  const layout = wb.pivotLayout(0, pivot);
  assert.ok(layout.status.ok, JSON.stringify(layout.status));
  const rows = new Map();
  for (const cell of layout.cells) {
    const value = cell.value;
    const text = value.kind === 1 ? String(value.number) : value.kind === 3 ? value.text : '';
    if (!rows.has(cell.row)) rows.set(cell.row, new Array(layout.cols).fill(''));
    rows.get(cell.row)[cell.col - layout.left] = text;
  }
  return [...rows.keys()].sort((a, b) => a - b).map((row) => rows.get(row).join('|'));
}

const BASE_GRID = ['行ラベル|Sum of Amount', 'East|10', 'West|30', '総計|40'];

test('pivot nested integer fields reject narrowing before their C mutations', async () => {
  const mod = await getModule();
  const invalid = [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY];

  for (const value of invalid) {
    withPivot(mod.Workbook, ({ wb, pivot }) => {
      const before = wb.pivotFieldCount(0, pivot).value;
      assert.throws(
        () => wb.pivotFieldAdd(0, pivot, { sourceName: 'Region', axis: value }),
        RangeError,
        `pivotFieldAdd axis=${String(value)}`,
      );
      assert.equal(wb.pivotFieldCount(0, pivot).value, before);
    });

    withPivot(mod.Workbook, ({ wb, pivot, amount }) => {
      const before = wb.pivotDataFieldCount(0, pivot).value;
      assert.throws(
        () => wb.pivotDataFieldAdd(0, pivot, { name: 'Bad', fieldIndex: value, aggregation: 0 }),
        RangeError,
        `pivotDataFieldAdd fieldIndex=${String(value)}`,
      );
      assert.equal(wb.pivotDataFieldCount(0, pivot).value, before);

      assert.throws(
        () => wb.pivotDataFieldSet(0, pivot, 0, { name: 'Bad', fieldIndex: amount, aggregation: value }),
        RangeError,
        `pivotDataFieldSet aggregation=${String(value)}`,
      );
      assert.equal(wb.pivotDataFieldCount(0, pivot).value, before);
      assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);

      for (const key of ['showAs', 'showAsBaseField', 'showAsBaseItem']) {
        assert.throws(
          () =>
            wb.pivotDataFieldSet(0, pivot, 0, {
              name: 'Bad',
              fieldIndex: amount,
              aggregation: 0,
              [key]: value,
            }),
          RangeError,
          `pivotDataFieldSet ${key}=${String(value)}`,
        );
        assert.equal(wb.pivotDataFieldCount(0, pivot).value, before);
        assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);
      }
    });
  }
});

test('pivot nested integer fields reject wrong primitives without changing the pivot', async () => {
  const mod = await getModule();
  const symbol = Symbol('pivot numeric');
  withPivot(mod.Workbook, ({ wb, pivot }) => {
    const fieldCount = wb.pivotFieldCount(0, pivot).value;
    assert.throws(() => wb.pivotFieldAdd(0, pivot, { sourceName: 'Region', axis: symbol }), TypeError);
    assert.equal(wb.pivotFieldCount(0, pivot).value, fieldCount);

    const dataFieldCount = wb.pivotDataFieldCount(0, pivot).value;
    assert.throws(() => wb.pivotDataFieldAdd(0, pivot, { name: 'Bad', fieldIndex: symbol, aggregation: 0 }), TypeError);
    assert.equal(wb.pivotDataFieldCount(0, pivot).value, dataFieldCount);

    const filterCount = wb.pivotFilterCount(0, pivot).value;
    assert.throws(
      () =>
        wb.pivotFilterAdd(0, pivot, {
          axis: 0,
          fieldName: 'Region',
          type: 1,
          valueKind: symbol,
          valueDouble: 15,
        }),
      TypeError,
    );
    assert.equal(wb.pivotFilterCount(0, pivot).value, filterCount);
  });
});

test('pivot filters validate axis, value kind, and data-field index before mutation', async () => {
  const mod = await getModule();
  const invalidSpecs = [
    { axis: 2 ** 32 },
    { type: 2 ** 32 },
    { valueKind: 1.5 },
    { valueInt: 1.5 },
    { valueHighKind: Number.NaN },
    { valueHighInt: Number.POSITIVE_INFINITY },
    { dataFieldIndex: 2 ** 32 },
  ];
  for (const override of invalidSpecs) {
    withPivot(mod.Workbook, ({ wb, pivot }) => {
      const before = wb.pivotFilterCount(0, pivot).value;
      assert.throws(
        () =>
          wb.pivotFilterAdd(0, pivot, {
            axis: 0,
            fieldName: 'Region',
            type: 1,
            valueKind: 1,
            valueDouble: 15,
            ...override,
          }),
        RangeError,
        `pivotFilterAdd ${JSON.stringify(override)}`,
      );
      assert.equal(wb.pivotFilterCount(0, pivot).value, before);
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
        () => wb.pivotSetRowFieldOrder(0, pivot, [value]),
        typeof value === 'symbol' ? TypeError : RangeError,
        `pivotSetRowFieldOrder element=${String(value)}`,
      );
      assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);
    });
  }

  withPivot(mod.Workbook, ({ wb, pivot, region }) => {
    assert.ok(wb.pivotSetRowFieldOrder(0, pivot, [region]).ok);
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
    assert.throws(() => wb.pivotSetRowFieldOrder(0, pivot, indices), /pivot order getter/);
    assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);
  });
});

test('pivot nested string getters stop before changing the report', async () => {
  const mod = await getModule();
  withPivot(mod.Workbook, ({ wb, pivot }) => {
    const fieldCount = wb.pivotFieldCount(0, pivot).value;
    const spec = { axis: 0 };
    Object.defineProperty(spec, 'sourceName', {
      enumerable: true,
      get() {
        throw new Error('pivot field getter');
      },
    });
    assert.throws(() => wb.pivotFieldAdd(0, pivot, spec), /pivot field getter/);
    assert.equal(wb.pivotFieldCount(0, pivot).value, fieldCount);
    assert.deepEqual(pivotGrid(wb, pivot), BASE_GRID);
  });
});

test('pivot cache source snapshots each nested getter once before mutation', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const cache = wb.pivotCacheCreate(0).index;
    let reads = 0;
    const source = { present: true, sheet: 'Sheet1' };
    Object.defineProperty(source, 'ref', {
      enumerable: true,
      get() {
        reads += 1;
        if (reads > 1) throw new Error('pivot cache source getter read twice');
        return 'A1:B3';
      },
    });
    const result = wb.pivotCacheSetWorksheetSource(cache, source);
    assert.ok(result.ok, JSON.stringify(result));
    assert.equal(reads, 1);
    assert.equal(wb.pivotCacheGetWorksheetSource(cache).ref, 'A1:B3');
  } finally {
    wb.dispose();
  }
});

test('pivot optional strings and nullish numeric defaults keep the valid contract', async () => {
  const mod = await getModule();
  withPivot(mod.Workbook, ({ wb, pivot, amount }) => {
    const result = wb.pivotDataFieldAdd(0, pivot, {
      name: 'Count of Amount',
      fieldIndex: amount,
      aggregation: 1,
      numberFormat: null,
      showAs: undefined,
      showAsBaseField: null,
      showAsBaseItem: undefined,
    });
    assert.ok(result.status.ok, JSON.stringify(result.status));
    assert.equal(wb.pivotDataFieldCount(0, pivot).value, 2);
  });
});
