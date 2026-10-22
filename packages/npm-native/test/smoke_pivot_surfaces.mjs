import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

// Builds a workbook with one pivot, ready for the filter-contract cases.
// `dispose` is the caller's job; the WASM package uses `delete` instead.
function makePivotWorkbook(Workbook) {
  const wb = Workbook.createDefault();
  const cacheId = wb.pivotCacheCreate(0).index;
  // A cache with no declared source cannot be saved: Excel offers to repair
  // any package containing one, so the writer refuses rather than emit it.
  assert.ok(wb.pivotCacheSetWorksheetSource(cacheId, { present: true, ref: 'A1:B3', sheet: 'Sheet1' }).ok);
  assert.ok(wb.pivotCacheFieldAdd(cacheId, 'Region').status.ok);
  assert.ok(wb.pivotCacheFieldAdd(cacheId, 'Amount').status.ok);
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const rec = wb.pivotCacheRecordAdd(cacheId).index;
    assert.ok(wb.pivotCacheRecordSetText(cacheId, rec, 0, region).ok);
    assert.ok(wb.pivotCacheRecordSetNumber(cacheId, rec, 1, amount).ok);
  }
  const pivot = wb.pivotCreate(0, 'Pivot1', cacheId, 0, 4);
  assert.ok(pivot.status.ok, JSON.stringify(pivot.status));
  assert.ok(wb.pivotFieldAdd(0, pivot.index, { sourceName: 'Region', axis: 0 }).status.ok);
  const amountField = wb.pivotFieldAdd(0, pivot.index, { sourceName: 'Amount', axis: 2 });
  assert.ok(amountField.status.ok);
  assert.ok(
    wb.pivotDataFieldAdd(0, pivot.index, {
      name: 'Sum of Amount',
      fieldIndex: amountField.index,
      aggregation: 0,
    }).status.ok,
  );
  return { wb, pivot: pivot.index };
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
    assert.ok(wb.pivotFilterAdd(0, pivot, PIVOT_FILTER).ok);
    assert.equal(wb.pivotFilterCount(0, pivot).value, 1);
    const got = wb.pivotFilterAt(0, pivot, 0);
    assert.ok(got.status.ok, JSON.stringify(got.status));
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
    assert.equal(wb.pivotFilterCount(0, pivot).value, 0);
    assert.ok(wb.pivotFilterAdd(0, pivot, PIVOT_FILTER).ok);
    assert.equal(wb.pivotFilterCount(0, pivot).value, 1);
  } finally {
    wb.dispose();
  }
});

test('pivot filters survive save/load', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  let bytes;
  try {
    assert.ok(wb.pivotFilterAdd(0, pivot, PIVOT_FILTER).ok);
    assert.equal(wb.pivotFilterCount(0, pivot).value, 1);
    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    bytes = saved.bytes;
  } finally {
    wb.dispose();
  }
  const reloaded = mod.Workbook.loadBytes(bytes);
  try {
    assert.equal(reloaded.pivotCount(0).value, 1);
    // The filter is written as `<filters>` and decodes back into the list.
    assert.equal(reloaded.pivotFilterCount(0, 0).value, 1);
    const got = reloaded.pivotFilterAt(0, 0, 0);
    assert.ok(got.status.ok, JSON.stringify(got.status));
    assert.equal(got.fieldName, PIVOT_FILTER.fieldName);
    assert.equal(got.type, PIVOT_FILTER.type);
    assert.equal(got.valueKind, PIVOT_FILTER.valueKind);
    assert.equal(got.valueDouble, PIVOT_FILTER.valueDouble);
  } finally {
    reloaded.dispose();
  }
});

test('pivotFilterAt rejects an out-of-range index', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    const got = wb.pivotFilterAt(0, pivot, 99);
    assert.equal(got.status.ok, false);
  } finally {
    wb.dispose();
  }
});

test('a required pivot spec string field left out of the object is rejected, not coerced to "undefined"', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    const fieldCountBefore = wb.pivotFieldCount(0, pivot).value;
    const field = wb.pivotFieldAdd(0, pivot, { axis: 0 }); // no `sourceName`
    assert.equal(field.status.ok, false);
    assert.equal(field.status.status, 7001);
    assert.equal(wb.pivotFieldCount(0, pivot).value, fieldCountBefore, 'no field must have been added');

    // `name` is optional: an omitted one is derived from the workbook profile.
    const dataFieldCountBefore = wb.pivotDataFieldCount(0, pivot).value;
    const dataField = wb.pivotDataFieldAdd(0, pivot, { fieldIndex: 0, aggregation: 0 }); // no `name`
    assert.ok(dataField.status.ok, JSON.stringify(dataField.status));
    assert.equal(wb.pivotDataFieldCount(0, pivot).value, dataFieldCountBefore + 1);
    assert.equal(wb.pivotDataFieldAt(0, pivot, dataField.index).name, 'Sum of Region');

    const filterCountBefore = wb.pivotFilterCount(0, pivot).value;
    const filter = wb.pivotFilterAdd(0, pivot, { axis: 0, type: 1, valueKind: 1, valueDouble: 15 }); // no `fieldName`
    assert.equal(filter.ok, false);
    assert.equal(filter.status, 7001);
    assert.equal(wb.pivotFilterCount(0, pivot).value, filterCountBefore, 'no filter must have been added');
  } finally {
    wb.dispose();
  }
});

test('pivot definition read-backs report what the setters wrote', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    assert.equal(wb.pivotCacheId(0, pivot).index, wb.pivotCacheIdAt(0).index);
    assert.equal(wb.pivotName(0, pivot).value, 'Pivot1');
    assert.deepEqual((({ rows, cols }) => ({ rows, cols }))(wb.pivotGetGrandTotals(0, pivot)), {
      rows: true,
      cols: true,
    });
    assert.ok(wb.pivotSetGrandTotals(0, pivot, false, true).ok);
    const totals = wb.pivotGetGrandTotals(0, pivot);
    assert.ok(totals.status.ok);
    assert.equal(totals.rows, false);
    assert.equal(totals.cols, true);

    assert.deepEqual(wb.pivotRowFieldOrder(0, pivot).indices, [0]);
    assert.deepEqual(wb.pivotColFieldOrder(0, pivot).indices, []);

    assert.ok(wb.pivotFieldSetSort(0, pivot, 0, 0, 'Sum of Amount').ok);
    const region = wb.pivotFieldAt(0, pivot, 0);
    assert.ok(region.status.ok, JSON.stringify(region.status));
    assert.equal(region.sourceName, 'Region');
    assert.equal(region.axis, 0);
    assert.equal(region.cacheFieldIndex, 0);
    assert.equal(region.sortOrder, 0);
    assert.equal(region.sortByField, 'Sum of Amount');
    assert.equal(region.hasDateGroup, false);
    assert.equal(region.dateStartYear, -1);
    assert.equal(wb.pivotFieldAt(0, pivot, 1).axis, 2);

    assert.ok(wb.pivotFieldAddItem(0, pivot, 0, 'East', true).ok);
    assert.ok(wb.pivotFieldAddItem(0, pivot, 0, 'West', true).ok);
    assert.ok(wb.pivotFieldSetItemVisible(0, pivot, 0, 1, false).ok);
    assert.equal(wb.pivotFieldItemCount(0, pivot, 0).value, 2);
    const east = wb.pivotFieldItemAt(0, pivot, 0, 0);
    assert.equal(east.name, 'East');
    // The scratch cache holds no shared items, so a label binds to nothing.
    assert.equal(east.cacheIndex, -1);
    assert.equal(east.visible, true);
    assert.equal(wb.pivotFieldItemAt(0, pivot, 0, 1).visible, false);
    assert.ok(wb.pivotFieldAddItemAt(0, pivot, 0, 5, true).ok);
    const bound = wb.pivotFieldItemAt(0, pivot, 0, 2);
    assert.equal(bound.cacheIndex, 5);
    assert.equal(bound.name, '');

    const dataField = wb.pivotDataFieldAt(0, pivot, 0);
    assert.ok(dataField.status.ok, JSON.stringify(dataField.status));
    assert.equal(dataField.name, 'Sum of Amount');
    assert.equal(dataField.fieldIndex, 1);
    assert.equal(dataField.aggregation, 0);
    assert.equal(dataField.showAsBaseField, -1);
    assert.equal(dataField.showAsBaseItem, -1);
  } finally {
    wb.dispose();
  }
});

test('pivot style info defaults on create, reads back what the setter wrote and rejects bad arguments', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    const initial = wb.pivotGetStyleInfo(0, pivot);
    assert.ok(initial.status.ok, JSON.stringify(initial.status));
    assert.equal(initial.present, true);
    assert.equal(initial.name, 'PivotStyleLight16');
    assert.deepEqual(
      [initial.showRowHeaders, initial.showColHeaders, initial.showRowStripes, initial.showColStripes],
      [true, true, false, false],
    );
    assert.equal(initial.showLastColumn, true);

    assert.ok(wb.pivotSetStyleInfo(0, pivot, { name: 'PivotStyleDark3', showRowStripes: true }).ok);
    const set = wb.pivotGetStyleInfo(0, pivot);
    assert.equal(set.name, 'PivotStyleDark3');
    assert.deepEqual(
      [set.showRowHeaders, set.showColHeaders, set.showRowStripes, set.showColStripes, set.showLastColumn],
      [false, false, true, false, false],
    );

    // Survives a save/load.
    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    const reloaded = mod.Workbook.loadBytes(saved.bytes);
    try {
      assert.equal(reloaded.pivotGetStyleInfo(0, 0).name, 'PivotStyleDark3');
      assert.equal(reloaded.pivotGetStyleInfo(0, 0).showRowStripes, true);
    } finally {
      reloaded.dispose();
    }

    // An omitted name keeps the element without a style name.
    assert.ok(wb.pivotSetStyleInfo(0, pivot, {}).ok);
    const bare = wb.pivotGetStyleInfo(0, pivot);
    assert.equal(bare.present, true);
    assert.equal(bare.name, '');

    assert.equal(wb.pivotGetStyleInfo(0, 9).status.ok, false);
    assert.equal(wb.pivotSetStyleInfo(0, 9, { name: 'X' }).ok, false);
  } finally {
    wb.dispose();
  }
});

test('pivot definition read-backs reject out-of-range indices', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    for (const r of [
      wb.pivotCacheId(0, 9),
      wb.pivotName(0, 9),
      wb.pivotGetGrandTotals(9, pivot),
      wb.pivotRowFieldOrder(0, 9),
      wb.pivotColFieldOrder(0, 9),
      wb.pivotFieldAt(0, pivot, 99),
      wb.pivotFieldItemCount(0, pivot, 99),
      wb.pivotFieldItemAt(0, pivot, 0, 99),
      wb.pivotDataFieldAt(0, pivot, 99),
    ]) {
      const status = r.status;
      assert.equal(status.ok, false, JSON.stringify(status));
    }
    assert.equal(wb.pivotFieldAt(0, pivot, 99).axis, -1);
  } finally {
    wb.dispose();
  }
});

test('an empty or omitted data field name is derived from the workbook profile', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    for (const [profile, expected] of [
      ['win-365-ja_JP', '合計 / Amount'],
      ['win-365-en_US', 'Sum of Amount'],
    ]) {
      assert.ok(wb.setExcelProfileId(profile).ok);
      for (const spec of [
        { fieldIndex: 1, aggregation: 0 },
        { name: '', fieldIndex: 1, aggregation: 0 },
      ]) {
        const added = wb.pivotDataFieldAdd(0, pivot, spec);
        assert.ok(added.status.ok, JSON.stringify(added.status));
        assert.equal(wb.pivotDataFieldAt(0, pivot, added.index).name, expected);
      }
      assert.ok(wb.pivotDataFieldSet(0, pivot, 0, { fieldIndex: 1, aggregation: 2 }).ok);
      assert.ok(wb.pivotDataFieldAt(0, pivot, 0).name.endsWith(' Amount'));
    }
    assert.equal(wb.pivotDataFieldAdd(0, pivot, { fieldIndex: 99, aggregation: 0 }).status.ok, false);
  } finally {
    wb.dispose();
  }
});

test('pivotFieldAdd binds by cache column name and axis -1 takes a field off the report', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const cacheId = wb.pivotCacheCreate(0).index;
    assert.ok(wb.pivotCacheSetWorksheetSource(cacheId, { present: true, ref: 'A1:D3', sheet: 'Sheet1' }).ok);
    for (const name of ['Region', 'Product', 'Sales', 'Qty']) assert.ok(wb.pivotCacheFieldAdd(cacheId, name).status.ok);
    for (const sales of [100, 200]) {
      const rec = wb.pivotCacheRecordAdd(cacheId).index;
      for (const [col, value] of [0, 0, sales, 1].entries())
        assert.ok(wb.pivotCacheRecordSetNumber(cacheId, rec, col, value).ok);
    }
    const pivot = wb.pivotCreate(0, 'PT', cacheId, 0, 6).index;
    const region = wb.pivotFieldAdd(0, pivot, { sourceName: 'Region', axis: 0 });
    const sales = wb.pivotFieldAdd(0, pivot, { sourceName: 'Sales', axis: 2 });
    assert.equal(region.index, 0);
    assert.equal(sales.index, 2);
    assert.equal(wb.pivotFieldCount(0, pivot).value, 4);
    assert.equal(wb.pivotFieldAt(0, pivot, 1).axis, -1);
    assert.ok(wb.pivotDataFieldAdd(0, pivot, { fieldIndex: 2, aggregation: 0 }).status.ok);
    assert.equal(wb.pivotDataFieldCount(0, pivot).value, 1);

    assert.ok(wb.pivotFieldSetAxis(0, pivot, 2, -1).ok);
    assert.equal(wb.pivotFieldAt(0, pivot, 2).axis, -1);
    assert.equal(wb.pivotDataFieldCount(0, pivot).value, 0);
    assert.ok(wb.pivotFieldSetAxis(0, pivot, 0, -1).ok);
    assert.deepEqual(wb.pivotRowFieldOrder(0, pivot).indices, []);
    assert.equal(wb.pivotFieldSetAxis(0, pivot, 0, 4).ok, false);
    assert.equal(wb.pivotFieldSetAxis(0, pivot, 0, -2).ok, false);
  } finally {
    wb.dispose();
  }
});
