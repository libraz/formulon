// A pivot carries one field per cache field, cache values state their date-ness, fields are passed by name, filters
// name their kind, and a pivot hands out its cache.
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
const { PivotAxis, PivotAggregation, PivotFilterKind, PivotShowValuesAs, StatusCode } = f;

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

/** A worksheet source and a cache over Region / Amount / Date with records; no pivot yet. */
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

const newPivot = (ws, wb, name = 'Pivot1') => {
  const cache = addCache(ws, wb);
  return { cache, pivot: ws.addPivot({ name, cacheId: cache.cacheId, anchor: 'F1' }) };
};

describe('a pivot carries one field per cache field', () => {
  it('new pivot lists one field per cache field', () => {
    withSheet((ws, wb) => {
      const { pivot } = newPivot(ws, wb);
      assert.equal(pivot.listFields().length, 3);
      assert.equal(typeof pivot.addField, 'undefined');
    });
  });

  it('setFieldAxis by name works on a new pivot', () => {
    withSheet((ws, wb) => {
      const { pivot } = newPivot(ws, wb);
      pivot.setFieldAxis('Region', PivotAxis.Row);
      assert.deepEqual(pivot.getRowFieldOrder(), [0]);
    });
  });

  it('a cache field added later appears on every pivot of that cache', () => {
    withSheet((ws, wb) => {
      const { cache, pivot } = newPivot(ws, wb);
      cache.addCalculatedField('Double', '=Amount*2');
      assert.equal(pivot.listFields().length, 4);
    });
  });
});

describe('a pivot cache record value states whether it is a date', () => {
  it('getRecordValue returns value and isDate', () => {
    withSheet((ws, wb) => {
      const cache = addCache(ws, wb);
      const plain = cache.getRecordValue(0, 'Amount');
      assert.equal(plain.value, 10);
      assert.equal(plain.isDate, false);
      const record = cache.addRecord();
      cache.setRecordValue(record, 2, new Date(2024, 0, 1));
      const date = cache.getRecordValue(record, 'Date');
      assert.equal(date.value, 45292);
      assert.equal(date.isDate, true);
    });
  });

  it('writing a read record value back keeps the date', () => {
    withSheet((ws, wb) => {
      const cache = addCache(ws, wb);
      const source = cache.addRecord();
      cache.setRecordValue(source, 2, new Date(2024, 0, 1));
      const read = cache.getRecordValue(source, 'Date');
      const target = cache.addRecord();
      cache.setRecordValue(target, 'Date', read);
      assert.deepEqual(cache.getRecordValue(target, 'Date'), read);
    });
  });
});

describe('a pivot field is passed by cache field name everywhere', () => {
  it('cache methods take a field name', () => {
    withSheet((ws, wb) => {
      const cache = addCache(ws, wb);
      cache.addCalculatedField('Double', '=Amount*2');
      cache.setCalculatedFieldFormula('Double', '=Amount*3');
      assert.match(cache.listFields()[3].formula, /Amount\*3/);
      assert.equal(cache.listFieldSharedItems('Region').length, 2);
      cache.addFieldSharedItem('Region', 'North');
      assert.equal(cache.listFieldSharedItems(0).length, 3);
      assert.equal(cache.getRecordValue(1, 'Amount').value, 30);
      cache.setRecordValue(1, 'Amount', 31);
      assert.equal(cache.getRecordValue(1, 1).value, 31);
      cache.clearFieldSharedItems('Region');
      assert.equal(cache.listFieldSharedItems(0).length, 0);
      cache.removeCalculatedField('Double');
      assert.equal(cache.listFields().length, 3);
    });
  });

  it('setRowFieldOrder takes names and PIVOT_VALUES_FIELD', () => {
    withSheet((ws, wb) => {
      const { pivot } = newPivot(ws, wb);
      pivot.setFieldAxis('Region', PivotAxis.Row);
      pivot.setFieldAxis('Date', PivotAxis.Row);
      pivot.setFieldAxis('Amount', PivotAxis.Value);
      pivot.addDataField({ name: 'Sum of Amount', fieldIndex: 'Amount', aggregation: PivotAggregation.Sum });
      pivot.addDataField({ name: 'Count of Amount', fieldIndex: 'Amount', aggregation: PivotAggregation.Count });
      pivot.setRowFieldOrder(['Date', 'Region']);
      assert.deepEqual(pivot.getRowFieldOrder(), [2, 0]);
      pivot.setRowFieldOrder(['Region', f.PIVOT_VALUES_FIELD, 'Date']);
      assert.deepEqual(pivot.getRowFieldOrder(), [0, f.PIVOT_VALUES_FIELD, 2]);
    });
  });

  it('showAsBaseField takes a field name', () => {
    withSheet((ws, wb) => {
      const { pivot } = newPivot(ws, wb);
      pivot.setFieldAxis('Region', PivotAxis.Row);
      pivot.setFieldAxis('Amount', PivotAxis.Value);
      const index = pivot.addDataField({
        name: 'Difference',
        fieldIndex: 'Amount',
        showAs: PivotShowValuesAs.DifferenceFrom,
        showAsBaseField: 'Region',
      });
      assert.equal(pivot.getDataField(index).showAsBaseField, 0);
    });
  });

  it('unknown field name is NotFound', () => {
    withSheet((ws, wb) => {
      const { cache, pivot } = newPivot(ws, wb);
      expectCode(() => cache.listFieldSharedItems('Nope'), StatusCode.NotFound);
      expectCode(() => cache.getRecordValue(0, 'Nope'), StatusCode.NotFound);
      expectCode(() => pivot.setRowFieldOrder(['Nope']), StatusCode.NotFound);
    });
  });
});

describe('a pivot filter names its kind', () => {
  it('addFilter takes kind', () => {
    withSheet((ws, wb) => {
      const { pivot } = newPivot(ws, wb);
      pivot.setFieldAxis('Region', PivotAxis.Row);
      pivot.addFilter({ fieldIndex: 'Region', kind: PivotFilterKind.LabelContains, value: 'East' });
      const [filter] = pivot.listFilters();
      assert.equal(filter.kind, PivotFilterKind.LabelContains);
      assert.equal('type' in filter, false);
      expectTypeError(() =>
        pivot.addFilter({ fieldIndex: 'Region', type: PivotFilterKind.LabelContains, value: 'East' }),
      );
    });
  });
});

describe('a pivot hands out its cache and refuses bad names', () => {
  it('pv.cache() returns the workbook PivotCache handle', () => {
    withSheet((ws, wb) => {
      const { cache, pivot } = newPivot(ws, wb);
      assert.equal(pivot.cache(), cache);
      assert.equal(pivot.cache(), wb.pivotCache(cache.cacheId));
      assert.equal(typeof pivot.getCache, 'undefined');
    });
  });

  it('rename empty is InvalidArgument', () => {
    withSheet((ws, wb) => {
      const { pivot } = newPivot(ws, wb);
      expectCode(() => pivot.rename(''), StatusCode.InvalidArgument);
      assert.equal(pivot.name, 'Pivot1');
    });
  });
});
