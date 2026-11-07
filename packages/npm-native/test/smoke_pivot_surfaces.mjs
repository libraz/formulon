import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

// Builds a workbook with one pivot, ready for the filter-contract cases.
// `dispose` is the caller's job; the WASM package uses `delete` instead.
function makePivotWorkbook(Workbook) {
  const wb = Workbook.createDefault();
  const cache = wb.addPivotCache(0);
  const { cacheId } = cache;
  // A cache with no declared source cannot be saved: Excel offers to repair
  // any package containing one, so the writer refuses rather than emit it.
  cache.setWorksheetSource({ ref: 'A1:B3', sheet: 'Sheet1' });
  cache.addField('Region');
  cache.addField('Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const rec = cache.addRecord();
    cache.setRecordValue(rec, 0, region);
    cache.setRecordValue(rec, 1, amount);
  }
  const pivot = wb.sheet(0).addPivot({ name: 'Pivot1', cacheId, anchorRow: 0, anchorCol: 4 });
  pivot.addField({ sourceName: 'Region', axis: 0 });
  const amountField = pivot.addField({ sourceName: 'Amount', axis: 2 });
  pivot.addDataField({
    name: 'Sum of Amount',
    fieldIndex: amountField,
    aggregation: 0,
  });
  return { wb, pivot: pivot };
}

// greaterThan on a double payload: the shape pivotFilterAt must read back.
const PIVOT_FILTER = Object.freeze({
  axis: 0,
  fieldName: 'Region',
  type: 1,
  valueKind: 1,
  valueDouble: 15,
});

test('pivotFilterAt reads back an added filter field for field', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    pivot.addFilter(PIVOT_FILTER);
    assert.equal(pivot.listFilters().length, 1);
    const got = pivot.listFilters()[0];
    assert.equal(got.axis, PIVOT_FILTER.axis);
    assert.equal(got.fieldName, PIVOT_FILTER.fieldName);
    assert.equal(got.type, PIVOT_FILTER.type);
    assert.equal(got.valueKind, PIVOT_FILTER.valueKind);
    assert.equal(got.valueDouble, PIVOT_FILTER.valueDouble);
  } finally {
    wb.dispose();
  }
});

test('pivotFilterCount counts the filters on the pivot', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    assert.equal(pivot.listFilters().length, 0);
    pivot.addFilter(PIVOT_FILTER);
    assert.equal(pivot.listFilters().length, 1);
  } finally {
    wb.dispose();
  }
});

test('pivot filters survive save/load', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  let bytes;
  try {
    pivot.addFilter(PIVOT_FILTER);
    assert.equal(pivot.listFilters().length, 1);
    const saved = wb.save();
    bytes = saved;
  } finally {
    wb.dispose();
  }
  const reloaded = mod.Workbook.load(bytes);
  try {
    assert.equal(reloaded.sheet(0).listPivots().length, 1);
    // The filter is written as `<filters>` and decodes back into the list.
    assert.equal(reloaded.sheet(0).pivot(0).listFilters().length, 1);
    const got = reloaded.sheet(0).pivot(0).listFilters()[0];
    assert.equal(got.fieldName, PIVOT_FILTER.fieldName);
    assert.equal(got.type, PIVOT_FILTER.type);
    assert.equal(got.valueKind, PIVOT_FILTER.valueKind);
    assert.equal(got.valueDouble, PIVOT_FILTER.valueDouble);
  } finally {
    reloaded.dispose();
  }
});

test('listPivotFilters rejects an out-of-range pivot', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    assert.throws(() => wb.sheet(0).pivot(99).listFilters(), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('a required pivot spec string field left out of the object is rejected, not coerced to "undefined"', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    const fieldCountBefore = pivot.listFields().length;
    // no `sourceName`
    assert.throws(
      () => pivot.addField({ axis: 0 }),
      (e) => isFormulonError(e, 7001),
    );
    assert.equal(pivot.listFields().length, fieldCountBefore, 'no field must have been added');

    // `name` is optional: an omitted one is derived from the workbook profile.
    const dataFieldCountBefore = pivot.listDataFields().length;
    const dataField = pivot.addDataField({ fieldIndex: 0, aggregation: 0 }); // no `name`
    assert.equal(pivot.listDataFields().length, dataFieldCountBefore + 1);
    assert.equal(pivot.listDataFields()[dataField].name, 'Sum of Region');

    const filterCountBefore = pivot.listFilters().length;
    assert.throws(
      () => pivot.addFilter({ axis: 0, type: 1, valueKind: 1, valueDouble: 15 }),
      (e) => isFormulonError(e, 7001),
    ); // no `fieldName`
    assert.equal(pivot.listFilters().length, filterCountBefore, 'no filter must have been added');
  } finally {
    wb.dispose();
  }
});

test('pivot definition read-backs report what the setters wrote', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    assert.equal(pivot.getCache(), wb.listPivotCaches()[0]);
    assert.equal(pivot.name, 'Pivot1');
    assert.deepEqual((({ rows, columns }) => ({ rows, columns }))(pivot.getGrandTotals()), {
      rows: true,
      columns: true,
    });
    pivot.setGrandTotals({ rows: false, columns: true });
    const totals = pivot.getGrandTotals();
    assert.equal(totals.rows, false);
    assert.equal(totals.columns, true);

    assert.deepEqual(pivot.getRowFieldOrder(), [0]);
    assert.deepEqual(pivot.getColumnFieldOrder(), []);

    pivot.setFieldSort(0, 0, { byField: 'Sum of Amount' });
    const region = pivot.listFields()[0];
    assert.equal(region.sourceName, 'Region');
    assert.equal(region.axis, 0);
    assert.equal(region.cacheFieldIndex, 0);
    assert.equal(region.sortOrder, 0);
    assert.equal(region.sortByField, 'Sum of Amount');
    assert.equal(region.datePeriods, 0);
    assert.equal(region.groupBaseIndex, null);
    assert.equal(pivot.listFields()[1].axis, 2);

    pivot.addFieldItem(0, 'East');
    pivot.addFieldItem(0, 'West');
    pivot.updateItem(0, 1, { visible: false });
    assert.equal(pivot.listFieldItems(0).length, 2);
    const east = pivot.getFieldItem(0, 0);
    assert.equal(east.name, 'East');
    // The scratch cache holds no shared items, so a label binds to nothing.
    assert.equal(east.cacheIndex, null);
    assert.equal(east.visible, true);
    assert.equal(pivot.getFieldItem(0, 1).visible, false);
    pivot.addFieldSharedItem(0, 5, { visible: true });
    const bound = pivot.getFieldItem(0, 2);
    assert.equal(bound.cacheIndex, 5);
    assert.equal(bound.name, '');

    const dataField = pivot.listDataFields()[0];
    assert.equal(dataField.name, 'Sum of Amount');
    assert.equal(dataField.fieldIndex, 1);
    assert.equal(dataField.aggregation, 0);
    assert.equal(dataField.showAsBaseField, null);
    assert.equal(dataField.showAsBaseItem, null);
  } finally {
    wb.dispose();
  }
});

test('pivot style info defaults on create, reads back what the setter wrote and rejects bad arguments', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    const initial = pivot.getStyleInfo();
    assert.notEqual(initial, null);
    assert.equal(initial.name, 'PivotStyleLight16');
    assert.deepEqual(
      [initial.showRowHeaders, initial.showColHeaders, initial.showRowStripes, initial.showColStripes],
      [true, true, false, false],
    );
    assert.equal(initial.showLastColumn, true);

    // A patch: the omitted flags keep their values.
    pivot.setStyleInfo({ name: 'PivotStyleDark3', showRowStripes: true });
    const set = pivot.getStyleInfo();
    assert.equal(set.name, 'PivotStyleDark3');
    assert.deepEqual(
      [set.showRowHeaders, set.showColHeaders, set.showRowStripes, set.showColStripes, set.showLastColumn],
      [true, true, true, false, true],
    );

    // Survives a save/load.
    const saved = wb.save();
    const reloaded = mod.Workbook.load(saved);
    try {
      assert.equal(reloaded.sheet(0).pivot(0).getStyleInfo().name, 'PivotStyleDark3');
      assert.equal(reloaded.sheet(0).pivot(0).getStyleInfo().showRowStripes, true);
    } finally {
      reloaded.dispose();
    }

    // An empty patch changes nothing; an empty name leaves the element without a style name.
    pivot.setStyleInfo({});
    assert.equal(pivot.getStyleInfo().name, 'PivotStyleDark3');
    pivot.setStyleInfo({ name: '' });
    const bare = pivot.getStyleInfo();
    assert.notEqual(bare, null);
    assert.equal(bare.name, '');

    assert.throws(() => wb.sheet(0).pivot(9).getStyleInfo(), isFormulonError);
    assert.throws(() => wb.sheet(0).pivot(9).setStyleInfo({ name: 'X' }), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('pivot definition read-backs reject out-of-range indices', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    for (const read of [
      () => wb.sheet(0).pivot(9).getCache(),
      () => wb.sheet(0).pivot(9).name,
      () => wb.sheet(9).pivot(pivot).getGrandTotals(),
      () => wb.sheet(0).pivot(9).getRowFieldOrder(),
      () => wb.sheet(0).pivot(9).getColumnFieldOrder(),
      () => wb.sheet(0).pivot(99).listFields(),
      () => pivot.listFieldItems(99).length,
      () => pivot.getFieldItem(0, 99),
      () => wb.sheet(0).pivot(99).listDataFields(),
    ]) {
      assert.throws(read, isFormulonError);
    }
  } finally {
    wb.dispose();
  }
});

test('an empty or omitted data field name is derived from the workbook profile', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    // The second data field over the same source takes a numbered name.
    for (const [profile, expected] of [
      ['win-365-ja_JP', ['合計 / Amount', '合計 / Amount2']],
      ['win-365-en_US', ['Sum of Amount', 'Sum of Amount2']],
    ]) {
      wb.setExcelProfileId(profile);
      for (const [i, spec] of [
        { fieldIndex: 1, aggregation: 0 },
        { name: '', fieldIndex: 1, aggregation: 0 },
      ].entries()) {
        const added = pivot.addDataField(spec);
        assert.equal(pivot.listDataFields()[added].name, expected[i]);
      }
      pivot.setDataField(0, { fieldIndex: 1, aggregation: 2 });
      assert.ok(pivot.listDataFields()[0].name.endsWith(' Amount'));
    }
    assert.throws(() => pivot.addDataField({ fieldIndex: 99, aggregation: 0 }), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('pivotFieldAdd binds by cache column name and a null axis takes a field off the report', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const cache = wb.addPivotCache(0);
    const { cacheId } = cache;
    cache.setWorksheetSource({ ref: 'A1:D3', sheet: 'Sheet1' });
    for (const name of ['Region', 'Product', 'Sales', 'Qty']) cache.addField(name);
    for (const sales of [100, 200]) {
      const rec = cache.addRecord();
      for (const [col, value] of [0, 0, sales, 1].entries()) cache.setRecordValue(rec, col, value);
    }
    const pivot = wb.sheet(0).addPivot({ name: 'PT', cacheId, anchorRow: 0, anchorCol: 6 });
    const region = pivot.addField({ sourceName: 'Region', axis: 0 });
    const sales = pivot.addField({ sourceName: 'Sales', axis: 2 });
    assert.equal(region, 0);
    assert.equal(sales, 2);
    assert.equal(pivot.listFields().length, 4);
    assert.equal(pivot.listFields()[1].axis, null);
    pivot.addDataField({ fieldIndex: 2, aggregation: 0 });
    assert.equal(pivot.listDataFields().length, 1);

    pivot.setFieldAxis(2, null);
    assert.equal(pivot.listFields()[2].axis, null);
    assert.equal(pivot.listDataFields().length, 0);
    pivot.setFieldAxis(0, null);
    assert.deepEqual(pivot.getRowFieldOrder(), []);
    assert.throws(() => pivot.setFieldAxis(0, 4), isFormulonError);
    assert.throws(() => pivot.setFieldAxis(0, -1), RangeError);
  } finally {
    wb.dispose();
  }
});
