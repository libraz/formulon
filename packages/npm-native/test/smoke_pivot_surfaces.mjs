import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

// Builds a workbook with one pivot, ready for the filter-contract cases.
// `dispose` is the caller's job; the WASM package uses `delete` instead.
function makePivotWorkbook(Workbook) {
  const wb = Workbook.createDefault();
  const cacheId = wb.addPivotCache(0);
  // A cache with no declared source cannot be saved: Excel offers to repair
  // any package containing one, so the writer refuses rather than emit it.
  wb.setPivotCacheWorksheetSource(cacheId, { present: true, ref: 'A1:B3', sheet: 'Sheet1' });
  wb.addPivotCacheField(cacheId, 'Region');
  wb.addPivotCacheField(cacheId, 'Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const rec = wb.addPivotCacheRecord(cacheId);
    wb.setPivotCacheRecordText(cacheId, rec, 0, region);
    wb.setPivotCacheRecordNumber(cacheId, rec, 1, amount);
  }
  const pivot = wb.sheet(0).addPivot('Pivot1', cacheId, 0, 4);
  wb.sheet(0).addPivotField(pivot, { sourceName: 'Region', axis: 0 });
  const amountField = wb.sheet(0).addPivotField(pivot, { sourceName: 'Amount', axis: 2 });
  wb.sheet(0).addPivotDataField(pivot, {
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
    wb.sheet(0).addPivotFilter(pivot, PIVOT_FILTER);
    assert.equal(wb.sheet(0).listPivotFilters(pivot).length, 1);
    const got = wb.sheet(0).listPivotFilters(pivot)[0];
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
    assert.equal(wb.sheet(0).listPivotFilters(pivot).length, 0);
    wb.sheet(0).addPivotFilter(pivot, PIVOT_FILTER);
    assert.equal(wb.sheet(0).listPivotFilters(pivot).length, 1);
  } finally {
    wb.dispose();
  }
});

test('pivot filters survive save/load', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  let bytes;
  try {
    wb.sheet(0).addPivotFilter(pivot, PIVOT_FILTER);
    assert.equal(wb.sheet(0).listPivotFilters(pivot).length, 1);
    const saved = wb.save();
    bytes = saved;
  } finally {
    wb.dispose();
  }
  const reloaded = mod.Workbook.load(bytes);
  try {
    assert.equal(reloaded.sheet(0).getPivotCount(), 1);
    // The filter is written as `<filters>` and decodes back into the list.
    assert.equal(reloaded.sheet(0).listPivotFilters(0).length, 1);
    const got = reloaded.sheet(0).listPivotFilters(0)[0];
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
    assert.throws(() => wb.sheet(0).listPivotFilters(pivot + 99), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('a required pivot spec string field left out of the object is rejected, not coerced to "undefined"', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    const fieldCountBefore = wb.sheet(0).listPivotFields(pivot).length;
    // no `sourceName`
    assert.throws(
      () => wb.sheet(0).addPivotField(pivot, { axis: 0 }),
      (e) => isFormulonError(e, 7001),
    );
    assert.equal(wb.sheet(0).listPivotFields(pivot).length, fieldCountBefore, 'no field must have been added');

    // `name` is optional: an omitted one is derived from the workbook profile.
    const dataFieldCountBefore = wb.sheet(0).listPivotDataFields(pivot).length;
    const dataField = wb.sheet(0).addPivotDataField(pivot, { fieldIndex: 0, aggregation: 0 }); // no `name`
    assert.equal(wb.sheet(0).listPivotDataFields(pivot).length, dataFieldCountBefore + 1);
    assert.equal(wb.sheet(0).listPivotDataFields(pivot)[dataField].name, 'Sum of Region');

    const filterCountBefore = wb.sheet(0).listPivotFilters(pivot).length;
    assert.throws(
      () => wb.sheet(0).addPivotFilter(pivot, { axis: 0, type: 1, valueKind: 1, valueDouble: 15 }),
      (e) => isFormulonError(e, 7001),
    ); // no `fieldName`
    assert.equal(wb.sheet(0).listPivotFilters(pivot).length, filterCountBefore, 'no filter must have been added');
  } finally {
    wb.dispose();
  }
});

test('pivot definition read-backs report what the setters wrote', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    assert.equal(wb.sheet(0).getPivotCacheId(pivot), wb.listPivotCaches()[0]);
    assert.equal(wb.sheet(0).getPivotName(pivot), 'Pivot1');
    assert.deepEqual((({ rows, columns }) => ({ rows, columns }))(wb.sheet(0).getPivotGrandTotals(pivot)), {
      rows: true,
      columns: true,
    });
    wb.sheet(0).setPivotGrandTotals(pivot, { rows: false, columns: true });
    const totals = wb.sheet(0).getPivotGrandTotals(pivot);
    assert.equal(totals.rows, false);
    assert.equal(totals.columns, true);

    assert.deepEqual(wb.sheet(0).getPivotRowFieldOrder(pivot), [0]);
    assert.deepEqual(wb.sheet(0).getPivotColumnFieldOrder(pivot), []);

    wb.sheet(0).setPivotFieldSort(pivot, 0, 0, 'Sum of Amount');
    const region = wb.sheet(0).listPivotFields(pivot)[0];
    assert.equal(region.sourceName, 'Region');
    assert.equal(region.axis, 0);
    assert.equal(region.cacheFieldIndex, 0);
    assert.equal(region.sortOrder, 0);
    assert.equal(region.sortByField, 'Sum of Amount');
    assert.equal(region.datePeriods, 0);
    assert.equal(region.groupBaseIndex, null);
    assert.equal(wb.sheet(0).listPivotFields(pivot)[1].axis, 2);

    wb.sheet(0).addPivotFieldItem(pivot, 0, 'East', true);
    wb.sheet(0).addPivotFieldItem(pivot, 0, 'West', true);
    wb.sheet(0).setPivotFieldItemVisible(pivot, 0, 1, false);
    assert.equal(wb.sheet(0).getPivotFieldItemCount(pivot, 0), 2);
    const east = wb.sheet(0).getPivotFieldItem(pivot, 0, 0);
    assert.equal(east.name, 'East');
    // The scratch cache holds no shared items, so a label binds to nothing.
    assert.equal(east.cacheIndex, null);
    assert.equal(east.visible, true);
    assert.equal(wb.sheet(0).getPivotFieldItem(pivot, 0, 1).visible, false);
    wb.sheet(0).addPivotFieldSharedItem(pivot, 0, 5, true);
    const bound = wb.sheet(0).getPivotFieldItem(pivot, 0, 2);
    assert.equal(bound.cacheIndex, 5);
    assert.equal(bound.name, '');

    const dataField = wb.sheet(0).listPivotDataFields(pivot)[0];
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
    const initial = wb.sheet(0).getPivotStyleInfo(pivot);
    assert.notEqual(initial, null);
    assert.equal(initial.name, 'PivotStyleLight16');
    assert.deepEqual(
      [initial.showRowHeaders, initial.showColHeaders, initial.showRowStripes, initial.showColStripes],
      [true, true, false, false],
    );
    assert.equal(initial.showLastColumn, true);

    wb.sheet(0).setPivotStyleInfo(pivot, { name: 'PivotStyleDark3', showRowStripes: true });
    const set = wb.sheet(0).getPivotStyleInfo(pivot);
    assert.equal(set.name, 'PivotStyleDark3');
    assert.deepEqual(
      [set.showRowHeaders, set.showColHeaders, set.showRowStripes, set.showColStripes, set.showLastColumn],
      [false, false, true, false, false],
    );

    // Survives a save/load.
    const saved = wb.save();
    const reloaded = mod.Workbook.load(saved);
    try {
      assert.equal(reloaded.sheet(0).getPivotStyleInfo(0).name, 'PivotStyleDark3');
      assert.equal(reloaded.sheet(0).getPivotStyleInfo(0).showRowStripes, true);
    } finally {
      reloaded.dispose();
    }

    // An omitted name keeps the element without a style name.
    wb.sheet(0).setPivotStyleInfo(pivot, {});
    const bare = wb.sheet(0).getPivotStyleInfo(pivot);
    assert.notEqual(bare, null);
    assert.equal(bare.name, '');

    assert.throws(() => wb.sheet(0).getPivotStyleInfo(9), isFormulonError);
    assert.throws(() => wb.sheet(0).setPivotStyleInfo(9, { name: 'X' }), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('pivot definition read-backs reject out-of-range indices', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    for (const read of [
      () => wb.sheet(0).getPivotCacheId(9),
      () => wb.sheet(0).getPivotName(9),
      () => wb.sheet(9).getPivotGrandTotals(pivot),
      () => wb.sheet(0).getPivotRowFieldOrder(9),
      () => wb.sheet(0).getPivotColumnFieldOrder(9),
      () => wb.sheet(0).listPivotFields(99),
      () => wb.sheet(0).getPivotFieldItemCount(pivot, 99),
      () => wb.sheet(0).getPivotFieldItem(pivot, 0, 99),
      () => wb.sheet(0).listPivotDataFields(99),
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
        const added = wb.sheet(0).addPivotDataField(pivot, spec);
        assert.equal(wb.sheet(0).listPivotDataFields(pivot)[added].name, expected[i]);
      }
      wb.sheet(0).setPivotDataField(pivot, 0, { fieldIndex: 1, aggregation: 2 });
      assert.ok(wb.sheet(0).listPivotDataFields(pivot)[0].name.endsWith(' Amount'));
    }
    assert.throws(() => wb.sheet(0).addPivotDataField(pivot, { fieldIndex: 99, aggregation: 0 }), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('pivotFieldAdd binds by cache column name and a null axis takes a field off the report', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const cacheId = wb.addPivotCache(0);
    wb.setPivotCacheWorksheetSource(cacheId, { present: true, ref: 'A1:D3', sheet: 'Sheet1' });
    for (const name of ['Region', 'Product', 'Sales', 'Qty']) wb.addPivotCacheField(cacheId, name);
    for (const sales of [100, 200]) {
      const rec = wb.addPivotCacheRecord(cacheId);
      for (const [col, value] of [0, 0, sales, 1].entries()) wb.setPivotCacheRecordNumber(cacheId, rec, col, value);
    }
    const pivot = wb.sheet(0).addPivot('PT', cacheId, 0, 6);
    const region = wb.sheet(0).addPivotField(pivot, { sourceName: 'Region', axis: 0 });
    const sales = wb.sheet(0).addPivotField(pivot, { sourceName: 'Sales', axis: 2 });
    assert.equal(region, 0);
    assert.equal(sales, 2);
    assert.equal(wb.sheet(0).listPivotFields(pivot).length, 4);
    assert.equal(wb.sheet(0).listPivotFields(pivot)[1].axis, null);
    wb.sheet(0).addPivotDataField(pivot, { fieldIndex: 2, aggregation: 0 });
    assert.equal(wb.sheet(0).listPivotDataFields(pivot).length, 1);

    wb.sheet(0).setPivotFieldAxis(pivot, 2, null);
    assert.equal(wb.sheet(0).listPivotFields(pivot)[2].axis, null);
    assert.equal(wb.sheet(0).listPivotDataFields(pivot).length, 0);
    wb.sheet(0).setPivotFieldAxis(pivot, 0, null);
    assert.deepEqual(wb.sheet(0).getPivotRowFieldOrder(pivot), []);
    assert.throws(() => wb.sheet(0).setPivotFieldAxis(pivot, 0, 4), isFormulonError);
    assert.throws(() => wb.sheet(0).setPivotFieldAxis(pivot, 0, -1), RangeError);
  } finally {
    wb.dispose();
  }
});
