// PivotCache.refresh rebuilds a cache from its worksheet source, and listFields states the sharedItems type flags.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

// The staged entry named by the package's own "main" field.
async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  if (!pkg.main) throw new Error('package.json is missing "main"');
  return import(pathToFileURL(path.resolve(pkgRoot, pkg.main)).href);
}

const f = await (await loadPackage()).createFormulon();
const { PivotAxis, PivotAggregation, StatusCode } = f;

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

const ROWS = [
  ['East', 10],
  ['West', 30],
  ['East', 5],
];

/** Source cells Region / Amount, a cache over them with records, and a pivot summing Amount by Region at E1. */
function build(ws, wb, rows = ROWS) {
  ws.setText('A1', 'Region');
  ws.setText('B1', 'Amount');
  rows.forEach(([region, amount], i) => {
    ws.setText(i + 1, 0, region);
    ws.setNumber(i + 1, 1, amount);
  });
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ range: `A1:B${rows.length + 1}`, sheet: 'Sheet1' });
  for (const name of ['Region', 'Amount']) cache.addField(name);
  for (const item of ['East', 'West']) cache.addFieldSharedItem(0, item);
  rows.forEach(([region, amount]) => {
    const record = cache.addRecord();
    cache.setRecordValue(record, 0, region === 'East' ? 0 : 1);
    cache.setRecordValue(record, 1, amount);
  });
  const pivot = ws.addPivot({ name: 'Pivot1', cacheId: cache.cacheId, anchor: 'E1' });
  pivot.setFieldAxis('Region', PivotAxis.Row);
  pivot.addDataField({ fieldIndex: 'Amount', aggregation: PivotAggregation.Sum });
  pivot.project();
  return { cache, pivot };
}

const sharedValues = (cache, field) => cache.listFieldSharedItems(field).map((item) => item.value);

describe('PivotCache.refresh', () => {
  it('re-reads the source into the cache and the report', () => {
    withSheet((ws, wb) => {
      const { cache } = build(ws, wb);
      assert.deepEqual(sharedValues(cache, 'Region'), ['East', 'West']);
      ws.setText('A4', 'North');
      ws.setNumber('B2', 1000);
      assert.equal(cache.refresh(), undefined);
      assert.deepEqual(sharedValues(cache, 'Region'), ['East', 'West', 'North']);
      assert.equal(cache.getRecordValue(0, 'Amount').value, 1000);
      let seen = false;
      for (let row = 0; row < 12; row++) for (let col = 4; col < 7; col++) seen ||= ws.getValue(row, col) === 'North';
      assert.ok(seen, 'the report shows the new item');
    });
  });

  it('fails with PreconditionFailed and changes nothing when the source does not resolve', () => {
    withSheet((ws, wb) => {
      const { cache } = build(ws, wb);
      cache.setWorksheetSource({ name: 'NoSuchName' });
      ws.setText('A4', 'North');
      expectCode(() => cache.refresh(), StatusCode.PreconditionFailed);
      assert.deepEqual(sharedValues(cache, 'Region'), ['East', 'West']);
      cache.setWorksheetSource(null);
      expectCode(() => cache.refresh(), StatusCode.PreconditionFailed);
    });
  });

  it('blocks a report that cannot grow and renders it once its range is free', () => {
    withSheet((ws, wb) => {
      const { cache, pivot } = build(ws, wb);
      assert.equal(pivot.getSpill(), null);
      // The report fills E1:F4; North makes it one row longer, into E5.
      for (let row = 4; row < 9; row++) ws.setText(row, 4, 'keep');
      ws.setText('A4', 'North');
      assert.equal(cache.refresh(), undefined);
      assert.deepEqual(sharedValues(cache, 'Region'), ['East', 'West', 'North']);
      assert.equal(pivot.isProjectionPending(), false);
      const needed = { firstRow: 0, firstCol: 4, lastRow: 4, lastCol: 5 };
      assert.deepEqual(pivot.getSpill(), { blocked: true, range: needed });
      assert.deepEqual(ws.getSpill('E1'), { blocked: true, range: needed });
      assert.equal(ws.getSpill('F2'), null);
      assert.equal(ws.getValue(4, 4), 'keep');
      for (let row = 4; row < 9; row++) ws.setBlank(row, 4);
      assert.equal(pivot.getSpill(), null);
      assert.equal(ws.getSpill('E1'), null);
      let seen = false;
      for (let row = 0; row < 6; row++) seen ||= ws.getValue(row, 4) === 'North';
      assert.ok(seen, 'the report shows the new item');
    });
  });

  it('getSpill fails with IndexOutOfRange for a pivot index past the last', () => {
    withSheet((ws, wb) => {
      build(ws, wb);
      expectCode(() => ws.pivot(9).getSpill(), StatusCode.IndexOutOfRange);
    });
  });

  it('fails with SecurityResourceLimit past the record cell limit', () => {
    withSheet((ws, wb) => {
      const { cache } = build(ws, wb);
      cache.setWorksheetSource({ range: 'A1:AG1048576', sheet: 'Sheet1' });
      expectCode(() => cache.refresh(), StatusCode.SecurityResourceLimit);
    });
  });

  it('fails with NotFound once the cache is removed', () => {
    withSheet((ws, wb) => {
      const cache = wb.addPivotCache();
      wb.removePivotCache(cache);
      expectCode(() => cache.refresh(), StatusCode.NotFound);
    });
  });
});

describe('PivotCache.listFields type flags', () => {
  it('states the flags of a text field and a number field', () => {
    withSheet((ws, wb) => {
      const { cache } = build(ws, wb);
      const [region, amount] = cache.listFields();
      assert.equal(region.containsString, true);
      assert.equal(region.containsNumber, false);
      assert.equal(region.containsSemiMixedTypes, true);
      assert.equal(region.containsNonDate, true);
      assert.equal(amount.containsNumber, true);
      assert.equal(amount.containsInteger, true);
      assert.equal(amount.containsString, false);
      assert.equal(amount.containsBlank, false);
      assert.equal(amount.containsDate, false);
      assert.equal(amount.containsMixedTypes, false);
    });
  });

  it('follows a refresh', () => {
    withSheet((ws, wb) => {
      const { cache } = build(ws, wb);
      ws.setNumber('B2', 1.5);
      ws.setBlank('B3');
      cache.refresh();
      const amount = cache.listFields()[1];
      assert.equal(amount.containsNumber, true);
      assert.equal(amount.containsInteger, false);
      assert.equal(amount.containsBlank, true);
      assert.equal(amount.containsSemiMixedTypes, true);
    });
  });
});

const NUMBERS = [
  ['East', 20],
  ['West', 150],
  ['East', 250],
  ['West', 300],
];

describe('Pivot numeric grouping', () => {
  it('groups by an interval, reads the grouping back and removes it', () => {
    withSheet((ws, wb) => {
      const { pivot } = build(ws, wb, NUMBERS);
      assert.equal(pivot.getFieldNumericGroup('Amount'), null);
      assert.equal(pivot.getFieldDateGroup('Amount'), null);

      pivot.setFieldNumericGroup('Amount', { interval: 100 });
      assert.deepEqual(pivot.getFieldNumericGroup('Amount'), {
        start: 20,
        end: 300,
        interval: 100,
        autoStart: true,
        autoEnd: true,
      });
      assert.equal(pivot.listFields().length, 3);
      // The group field's items: <20, 20-119, 120-219, 220-319, >320.
      pivot.setFieldAxis(2, PivotAxis.Row);
      pivot.project();
      assert.equal(pivot.listFieldItems(2).length, 5);
      const shown = [];
      for (let row = 0; row < 12; row++) for (let col = 4; col < 7; col++) shown.push(ws.getValue(row, col));
      for (const label of ['20-119', '120-219', '220-319']) assert.ok(shown.includes(label), label);
      assert.deepEqual(pivot.getFieldNumericGroup(2), pivot.getFieldNumericGroup('Amount'));

      pivot.setFieldNumericGroup('Amount', { interval: 100, start: 0, end: 400 });
      assert.deepEqual(pivot.getFieldNumericGroup('Amount'), {
        start: 0,
        end: 400,
        interval: 100,
        autoStart: false,
        autoEnd: false,
      });

      pivot.setFieldNumericGroup('Amount', null);
      assert.equal(pivot.getFieldNumericGroup('Amount'), null);
      assert.equal(pivot.listFields().length, 2);
    });
  });

  it('refuses a non-integer interval with NotImplemented and leaves the pivot alone', () => {
    withSheet((ws, wb) => {
      const { pivot } = build(ws, wb, NUMBERS);
      expectCode(() => pivot.setFieldNumericGroup('Amount', { interval: 1.5 }), StatusCode.NotImplemented);
      expectCode(() => pivot.setFieldNumericGroup('Amount', { interval: 0 }), StatusCode.InvalidArgument);
      assert.equal(pivot.getFieldNumericGroup('Amount'), null);
      assert.equal(pivot.listFields().length, 2);
    });
  });

  it('rejects malformed groups', () => {
    withSheet((ws, wb) => {
      const { pivot } = build(ws, wb, NUMBERS);
      assert.throws(() => pivot.setFieldNumericGroup('Amount', { interval: 'x' }), TypeError);
      assert.throws(() => pivot.setFieldNumericGroup('Amount', { interval: 10, start: 'x' }), TypeError);
      assert.throws(() => pivot.setFieldNumericGroup('Amount', 5), TypeError);
    });
  });
});
