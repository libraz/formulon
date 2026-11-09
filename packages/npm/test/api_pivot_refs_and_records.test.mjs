// Pivot references and records on the public entry: a cache passed by id or handle, a field passed by index or name,
// the subtotal-function set, cache record reads, one-at-a-time readers and the pivot definition record.
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

const f = await (await loadPackage()).createFormulon();

const { PivotAxis, PivotAggregation, PivotFilterKind, PivotReportLayout, StatusCode, ValueKind } = f;

/** Runs `body(ws, wb)` on a default workbook, then disposes it. */
function withSheet(body) {
  const wb = f.Workbook.createDefault();
  try {
    return body(wb.sheet(0), wb);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

function expectCode(fn, code) {
  assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, code, `code ${e.code}`);
    return true;
  });
}

const expectTypeError = (fn) => assert.throws(fn, (e) => e instanceof TypeError, 'expected a TypeError');

const ROWS = [
  ['East', 10, 45292],
  ['West', 30, 45323],
  ['East', 5, 45658],
];

/** A worksheet source, a cache over Region / Amount / Date with records, and nothing placed yet. */
function addCache(ws, wb) {
  ws.setText('A1', 'Region');
  ws.setText('B1', 'Amount');
  ws.setText('C1', 'Date');
  ROWS.forEach(([region, amount, date], i) => {
    ws.setText(i + 1, 0, region);
    ws.setNumber(i + 1, 1, amount);
    ws.setNumber(i + 1, 2, date);
  });
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ range: 'A1:C4', sheet: 'Sheet1' });
  for (const name of ['Region', 'Amount', 'Date']) cache.addField(name);
  for (const item of ['East', 'West']) cache.addFieldSharedItem(0, item);
  ROWS.forEach(([region, amount, date]) => {
    const record = cache.addRecord();
    cache.setRecordValue(record, 0, region === 'East' ? 0 : 1);
    cache.setRecordValue(record, 1, amount);
    cache.setRecordValue(record, 2, date);
  });
  return cache;
}

/** A cache and a pivot with Region on rows, Amount summed as a data field and Date on no axis. */
function buildPivot(ws, wb) {
  const cache = addCache(ws, wb);
  const pivot = ws.addPivot({ name: 'Pivot1', cacheId: cache.cacheId, anchor: 'F1' });
  pivot.addField({ fieldIndex: 'Region', axis: PivotAxis.Row });
  pivot.addField({ fieldIndex: 'Amount', axis: PivotAxis.Value });
  pivot.addField({ fieldIndex: 'Date', axis: PivotAxis.Row });
  pivot.setFieldAxis('Date', null);
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: 'Amount', aggregation: PivotAggregation.Sum });
  return { cache, pivot };
}

describe('a pivot cache is passed by id or by handle', () => {
  it('addPivot, pivotCache and removePivotCache take a PivotCache handle as they take its id', () => {
    withSheet((ws, wb) => {
      const cache = addCache(ws, wb);
      const byHandle = ws.addPivot({ name: 'ByHandle', cacheId: cache, anchor: 'F1' });
      const byId = ws.addPivot({ name: 'ById', cacheId: cache.cacheId, anchor: 'K1' });
      assert.equal(byHandle.getProperties().cacheId, cache.cacheId);
      assert.equal(byId.getProperties().cacheId, cache.cacheId);
      assert.equal(wb.pivotCache(cache), cache);
      assert.equal(wb.pivotCache(cache.cacheId), cache);
      const spare = wb.addPivotCache();
      assert.equal(wb.listPivotCaches().length, 2);
      wb.removePivotCache(spare);
      assert.deepEqual(
        wb.listPivotCaches().map((c) => c.cacheId),
        [cache.cacheId],
      );
    });
  });

  it('a cache handle of another workbook is InvalidArgument', () => {
    withSheet((ws, wb) => {
      const other = f.Workbook.createDefault();
      try {
        const foreign = other.addPivotCache();
        expectCode(() => wb.pivotCache(foreign), StatusCode.InvalidArgument);
        expectCode(() => wb.removePivotCache(foreign), StatusCode.InvalidArgument);
        expectCode(() => ws.addPivot({ name: 'P', cacheId: foreign, anchor: 'F1' }), StatusCode.InvalidArgument);
      } finally {
        other.dispose();
      }
    });
  });

  it('a value that is neither an id nor a handle is a TypeError', () => {
    withSheet((ws, wb) => {
      expectTypeError(() => wb.pivotCache('1'));
      expectTypeError(() => wb.pivotCache({}));
      expectTypeError(() => ws.addPivot({ name: 'P', cacheId: 'x', anchor: 'F1' }));
    });
  });
});

describe('a pivot field is passed by index or by source name', () => {
  it('a field name and its index filter alike, and the filter reads back the field index', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      pivot.addFilter({ fieldIndex: 'Region', type: PivotFilterKind.LabelContains, value: 'East' });
      pivot.addFilter({ fieldIndex: 0, type: PivotFilterKind.LabelContains, value: 'East' });
      const [byName, byIndex] = pivot.listFilters();
      assert.deepEqual(byName, byIndex);
      assert.equal(byName.fieldIndex, 0);
      assert.equal(byName.type, PivotFilterKind.LabelContains);
      assert.equal(byName.value, 'East');
      assert.equal('axis' in byName, false);
      assert.equal('fieldName' in byName, false);
    });
  });

  it('a name matches case-insensitively', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      pivot.addFilter({ fieldIndex: 'rEGION', type: PivotFilterKind.LabelContains, value: 'East' });
      assert.equal(pivot.listFilters()[0].fieldIndex, 0);
    });
  });

  it('the old axis and fieldName keys are refused on a filter', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      const base = { fieldIndex: 'Region', type: PivotFilterKind.LabelContains, value: 'East' };
      expectTypeError(() => pivot.addFilter({ ...base, axis: PivotAxis.Row }));
      expectTypeError(() => pivot.addFilter({ ...base, fieldName: 'Region' }));
      assert.equal(pivot.listFilters().length, 0);
    });
  });

  it('a filter without a field is a TypeError', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      expectTypeError(() => pivot.addFilter({ type: PivotFilterKind.LabelContains, value: 'East' }));
    });
  });

  it('a name no field has is NotFound and an index past the last field is IndexOutOfRange', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      const filter = { type: PivotFilterKind.LabelContains, value: 'East' };
      expectCode(() => pivot.addFilter({ ...filter, fieldIndex: 'Nope' }), StatusCode.NotFound);
      expectCode(() => pivot.addFilter({ ...filter, fieldIndex: 99 }), StatusCode.IndexOutOfRange);
      expectCode(() => pivot.addFilter({ ...filter, fieldIndex: -1 }), StatusCode.IndexOutOfRange);
      assert.equal(pivot.listFilters().length, 0);
    });
  });

  it('a filter on a field that sits on no axis is InvalidArgument and adds nothing', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      expectCode(
        () => pivot.addFilter({ fieldIndex: 'Date', type: PivotFilterKind.LabelContains, value: 'x' }),
        StatusCode.InvalidArgument,
      );
      assert.equal(pivot.listFilters().length, 0);
    });
  });

  it('addField places a cache field named by its name or its index', () => {
    withSheet((ws, wb) => {
      const cache = addCache(ws, wb);
      const pivot = ws.addPivot({ name: 'P', cacheId: cache.cacheId, anchor: 'F1' });
      assert.equal(pivot.addField({ fieldIndex: 'Region', axis: PivotAxis.Row }), 0);
      assert.equal(pivot.addField({ fieldIndex: 1, axis: PivotAxis.Value }), 1);
      assert.deepEqual(
        pivot.listFields().map((x) => [x.sourceName, x.axis]),
        [
          ['Region', PivotAxis.Row],
          ['Amount', PivotAxis.Value],
        ],
      );
      expectCode(() => pivot.addField({ fieldIndex: 99, axis: PivotAxis.Row }), StatusCode.IndexOutOfRange);
    });
  });

  it('addField refuses the sourceName key', () => {
    withSheet((ws, wb) => {
      const cache = addCache(ws, wb);
      const pivot = ws.addPivot({ name: 'P', cacheId: cache.cacheId, anchor: 'F1' });
      expectTypeError(() => pivot.addField({ sourceName: 'Date', axis: PivotAxis.Row }));
      assert.equal(pivot.listFields().length, 0);
    });
  });

  it('a data field names the field it aggregates by name or by index', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      const second = pivot.addDataField({
        name: 'Count of Amount',
        fieldIndex: 1,
        aggregation: PivotAggregation.Count,
      });
      assert.equal(pivot.getDataField(second).fieldIndex, 1);
      const third = pivot.addDataField({
        name: 'Max of Amount',
        fieldIndex: 'amount',
        aggregation: PivotAggregation.Max,
      });
      assert.equal(pivot.getDataField(third).fieldIndex, 1);
      expectCode(
        () => pivot.addDataField({ name: 'X', fieldIndex: 'Nope', aggregation: PivotAggregation.Sum }),
        StatusCode.NotFound,
      );
    });
  });

  it('every Pivot method that takes a field accepts a name', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      assert.deepEqual(pivot.getField('Region'), pivot.getField(0));
      pivot.updateField('Region', { customName: 'Area' });
      assert.equal(pivot.getField(0).customName, 'Area');
      assert.deepEqual(pivot.listFieldItems('Region'), pivot.listFieldItems(0));
      expectCode(() => pivot.getField('Nope'), StatusCode.NotFound);
    });
  });
});

describe('a pivot field carries its subtotal functions as a set', () => {
  it('setFieldSubtotalFunctions replaces the set and reads back in ascending function order', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      pivot.setFieldSubtotalFunctions(0, [PivotAggregation.Sum, PivotAggregation.Count]);
      assert.deepEqual([...pivot.getField(0).subtotalFunctions], [PivotAggregation.Sum, PivotAggregation.Count]);
      assert.equal(pivot.getField(0).defaultSubtotal, false);
      pivot.setFieldSubtotalFunctions('Region', [PivotAggregation.Max]);
      assert.deepEqual([...pivot.getField(0).subtotalFunctions], [PivotAggregation.Max]);
    });
  });

  it('an empty list clears the set', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      pivot.setFieldSubtotalFunctions(0, [PivotAggregation.Sum]);
      pivot.setFieldSubtotalFunctions(0, []);
      assert.deepEqual([...pivot.getField(0).subtotalFunctions], []);
      assert.equal(pivot.getField(0).defaultSubtotal, true);
    });
  });

  it('a repeated function is InvalidArgument and leaves the set unchanged', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      pivot.setFieldSubtotalFunctions(0, [PivotAggregation.Sum]);
      expectCode(
        () => pivot.setFieldSubtotalFunctions(0, [PivotAggregation.Sum, PivotAggregation.Sum]),
        StatusCode.InvalidArgument,
      );
      assert.deepEqual([...pivot.getField(0).subtotalFunctions], [PivotAggregation.Sum]);
    });
  });

  it('the add and clear methods are gone and the list must be an array', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      assert.equal(pivot.addFieldSubtotalFunction, undefined);
      assert.equal(pivot.clearFieldSubtotalFunctions, undefined);
      expectTypeError(() => pivot.setFieldSubtotalFunctions(0, PivotAggregation.Sum));
    });
  });

  it('the set survives save and load', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      pivot.setFieldSubtotalFunctions(0, [PivotAggregation.Sum, PivotAggregation.Count]);
      const back = f.Workbook.load(wb.save());
      try {
        assert.deepEqual(
          [...back.sheet(0).pivot(0).getField(0).subtotalFunctions],
          [PivotAggregation.Sum, PivotAggregation.Count],
        );
      } finally {
        back.dispose();
      }
    });
  });
});

describe('pivot cache records are read back', () => {
  it('getRecordValue returns the native value setRecordValue stored', () => {
    withSheet((ws, wb) => {
      const cache = addCache(ws, wb);
      cache.setRecordValue(0, 1, 42);
      assert.equal(cache.getRecordValue(0, 1), 42);
      cache.setRecordValue(0, 1, 'text');
      assert.equal(cache.getRecordValue(0, 1), 'text');
      cache.setRecordValue(0, 1, true);
      assert.equal(cache.getRecordValue(0, 1), true);
      cache.setRecordValue(0, 1, null);
      assert.equal(cache.getRecordValue(0, 1), null);
      cache.setRecordValue(0, 1, { kind: ValueKind.Error, errorCode: f.ErrorCode.NA });
      const error = cache.getRecordValue(0, 1);
      assert.ok(error instanceof f.ExcelError, 'an error value reads as an ExcelError');
      assert.equal(error.code, f.ErrorCode.NA);
    });
  });

  it('a field backed by shared items reads the shared item index', () => {
    withSheet((ws, wb) => {
      const cache = addCache(ws, wb);
      assert.equal(cache.getRecordValue(0, 0), 0);
      assert.equal(cache.getRecordValue(1, 0), 1);
      assert.equal(cache.getRecordValue(2, 2), 45658);
    });
  });

  it('a record or field past the end is IndexOutOfRange', () => {
    withSheet((ws, wb) => {
      const cache = addCache(ws, wb);
      expectCode(() => cache.getRecordValue(3, 0), StatusCode.IndexOutOfRange);
      expectCode(() => cache.getRecordValue(0, 3), StatusCode.IndexOutOfRange);
    });
  });

  it('a negative record position is IndexOutOfRange', () => {
    withSheet((ws, wb) => {
      const cache = addCache(ws, wb);
      expectCode(() => cache.getRecordValue(-1, 0), StatusCode.IndexOutOfRange);
      expectCode(() => cache.getRecordValue(0, -1), StatusCode.IndexOutOfRange);
    });
  });
});

describe('pivot fields, data fields and filters are read one at a time', () => {
  it('getField, getDataField and getFilter equal the list element at that position', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      pivot.addFilter({ fieldIndex: 'Region', type: PivotFilterKind.LabelContains, value: 'East' });
      pivot.addFilter({ fieldIndex: 'Region', type: PivotFilterKind.LabelContains, value: 'a' });
      for (const [i, field] of pivot.listFields().entries()) assert.deepEqual(pivot.getField(i), field);
      for (const [i, field] of pivot.listDataFields().entries()) assert.deepEqual(pivot.getDataField(i), field);
      for (const [i, filter] of pivot.listFilters().entries()) assert.deepEqual(pivot.getFilter(i), filter);
      assert.equal(pivot.listFilters().length, 2);
    });
  });

  it('a position past the end or below zero is IndexOutOfRange', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      expectCode(() => pivot.getField(3), StatusCode.IndexOutOfRange);
      expectCode(() => pivot.getDataField(1), StatusCode.IndexOutOfRange);
      expectCode(() => pivot.getDataField(-1), StatusCode.IndexOutOfRange);
      expectCode(() => pivot.getFilter(0), StatusCode.IndexOutOfRange);
      expectCode(() => pivot.getFilter(-1), StatusCode.IndexOutOfRange);
    });
  });
});

describe('a negative pivot field position', () => {
  it('is IndexOutOfRange on getField', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      expectCode(() => pivot.getField(-1), StatusCode.IndexOutOfRange);
    });
  });
});

describe('Pivot.getProperties returns the definition', () => {
  it('carries the name, cache id, report layout, grand totals and style options', () => {
    withSheet((ws, wb) => {
      const { cache, pivot } = buildPivot(ws, wb);
      const properties = pivot.getProperties();
      assert.deepEqual(Object.keys(properties).sort(), [
        'cacheId',
        'grandTotals',
        'name',
        'reportLayout',
        'styleOptions',
      ]);
      assert.equal(properties.name, 'Pivot1');
      assert.equal(properties.cacheId, cache.cacheId);
      assert.equal(properties.reportLayout, pivot.getReportLayout());
      assert.deepEqual(properties.grandTotals, pivot.getGrandTotals());
      assert.deepEqual(properties.styleOptions, pivot.getStyleOptions());
    });
  });

  it('follows setReportLayout, updateGrandTotals and updateStyleOptions', () => {
    withSheet((ws, wb) => {
      const { pivot } = buildPivot(ws, wb);
      pivot.setReportLayout(PivotReportLayout.Tabular);
      pivot.updateGrandTotals({ rows: false });
      pivot.updateStyleOptions({ styleName: 'PivotStyleMedium9', showRowStripes: true });
      const properties = pivot.getProperties();
      assert.equal(properties.reportLayout, PivotReportLayout.Tabular);
      assert.equal(properties.grandTotals.rows, false);
      assert.equal(properties.styleOptions.styleName, 'PivotStyleMedium9');
      assert.equal(properties.styleOptions.showRowStripes, true);
      assert.deepEqual(properties.grandTotals, pivot.getGrandTotals());
      assert.deepEqual(properties.styleOptions, pivot.getStyleOptions());
    });
  });
});

describe('the pivot read and set members are present', () => {
  it('Pivot and PivotCache carry the one-at-a-time readers, the record read and the subtotal set', () => {
    withSheet((ws, wb) => {
      const { cache, pivot } = buildPivot(ws, wb);
      for (const name of ['getField', 'getDataField', 'getFilter', 'getProperties', 'setFieldSubtotalFunctions']) {
        assert.equal(typeof pivot[name], 'function', name);
      }
      assert.equal(typeof cache.getRecordValue, 'function');
    });
  });
});
