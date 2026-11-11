// PivotCache.refresh rebuilds a cache from its worksheet source, and listFields states the sharedItems type flags.
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
const { PivotAxis, PivotAggregation, SpillCause, StatusCode } = f;

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
      // The report fills E1:F4; North makes it one row longer, into E5.
      const drawn = { firstRow: 0, firstCol: 4, lastRow: 3, lastCol: 5 };
      assert.deepEqual(pivot.getSpill(), { anchor: { row: 0, col: 4 }, blocked: false, cause: null, range: drawn });
      assert.deepEqual(Object.keys(pivot.getSpill()), ['anchor', 'blocked', 'cause', 'range']);
      assert.deepEqual(Object.keys(ws.getSpill('E1')), ['anchor', 'blocked', 'cause', 'range']);
      for (let row = 4; row < 9; row++) ws.setText(row, 4, 'keep');
      ws.setText('A4', 'North');
      assert.equal(cache.refresh(), undefined);
      assert.deepEqual(sharedValues(cache, 'Region'), ['East', 'West', 'North']);
      assert.equal(pivot.isProjectionPending(), false);
      const needed = { firstRow: 0, firstCol: 4, lastRow: 4, lastCol: 5 };
      assert.deepEqual(pivot.getSpill(), {
        anchor: { row: 0, col: 4 },
        blocked: true,
        cause: SpillCause.Occupied,
        range: needed,
      });
      assert.deepEqual(ws.getSpill('E1'), {
        anchor: { row: 0, col: 4 },
        blocked: true,
        cause: SpillCause.Occupied,
        range: needed,
      });
      assert.equal(ws.getSpill('F2'), null);
      assert.equal(ws.getValue(4, 4), 'keep');
      for (let row = 4; row < 9; row++) ws.setBlank(row, 4);
      assert.deepEqual(pivot.getSpill(), { anchor: { row: 0, col: 4 }, blocked: false, cause: null, range: needed });
      assert.deepEqual(ws.getSpill('E1'), { anchor: { row: 0, col: 4 }, blocked: false, cause: null, range: needed });
      assert.equal(ws.getSpill('F2'), null);
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

  it('groups by a fractional interval and labels the ranges with decimals', () => {
    withSheet((ws, wb) => {
      const { pivot } = build(
        ws,
        wb,
        Array.from({ length: 12 }, (_, i) => [i % 2 ? 'East' : 'West', Number((0.3 * (i + 1)).toFixed(1))]),
      );
      pivot.setFieldNumericGroup('Amount', { interval: 0.5 });
      assert.deepEqual(pivot.getFieldNumericGroup('Amount'), {
        interval: 0.5,
        start: 0.3,
        end: 3.6,
        autoStart: true,
        autoEnd: true,
      });
      pivot.setFieldAxis(2, PivotAxis.Row);
      pivot.project();
      assert.equal(pivot.listFieldItems(2).length, 9);
      const shown = [];
      for (let row = 0; row < 14; row++) for (let col = 4; col < 7; col++) shown.push(ws.getValue(row, col));
      for (const label of ['0.3-0.8', '0.8-1.3', '1.3-1.8', '1.8-2.3', '2.3-2.8', '2.8-3.3', '3.3-3.8']) {
        assert.ok(shown.includes(label), label);
      }
    });
  });

  it('labels fractional ranges with the workbook locale decimal separator', () => {
    withSheet((ws, wb) => {
      wb.setProfileId('win-365-de_DE');
      const { pivot } = build(
        ws,
        wb,
        Array.from({ length: 12 }, (_, i) => [i % 2 ? 'East' : 'West', Number((0.3 * (i + 1)).toFixed(1))]),
      );
      pivot.setFieldNumericGroup('Amount', { interval: 0.5 });
      pivot.setFieldAxis(2, PivotAxis.Row);
      pivot.project();
      const shown = [];
      for (let row = 0; row < 14; row++) for (let col = 4; col < 7; col++) shown.push(ws.getValue(row, col));
      assert.ok(shown.includes('0,3-0,8'));
      assert.ok(!shown.includes('0.3-0.8'));
    });
  });

  it('reads the numeric record in a fixed key order and takes it back unchanged', () => {
    withSheet((ws, wb) => {
      const { pivot } = build(ws, wb, NUMBERS);
      const groups = [{ interval: 100 }, { interval: 100, start: 0, end: 400 }, { interval: 25.5, start: 10.25 }];
      for (const group of groups) {
        pivot.setFieldNumericGroup('Amount', group);
        const read = pivot.getFieldNumericGroup('Amount');
        assert.deepEqual(Object.keys(read), ['interval', 'start', 'end', 'autoStart', 'autoEnd']);
        pivot.setFieldNumericGroup('Amount', read);
        assert.deepEqual(pivot.getFieldNumericGroup('Amount'), read);
        assert.equal(pivot.listFields().length, 3);
      }
    });
  });

  it('ignores a bound whose auto flag is true', () => {
    withSheet((ws, wb) => {
      const { pivot } = build(ws, wb, NUMBERS);
      pivot.setFieldNumericGroup('Amount', { interval: 10, start: 5, autoStart: true });
      assert.deepEqual(pivot.getFieldNumericGroup('Amount'), {
        interval: 10,
        start: 20,
        end: 300,
        autoStart: true,
        autoEnd: true,
      });
      const stale = { interval: 10, start: Number.NaN, end: 1e9, autoStart: true, autoEnd: true };
      pivot.setFieldNumericGroup('Amount', stale);
      assert.equal(pivot.getFieldNumericGroup('Amount').end, 300);
    });
  });

  it('refuses a non-finite interval or bound with InvalidArgument and leaves the pivot alone', () => {
    withSheet((ws, wb) => {
      const { pivot } = build(ws, wb, NUMBERS);
      for (const key of ['interval', 'start', 'end']) {
        for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
          const group = { interval: 100, start: 0, end: 400, [key]: bad };
          assert.throws(
            () => pivot.setFieldNumericGroup('Amount', group),
            (e) => {
              assert.equal(e.name, 'FormulonError');
              assert.equal(e.code, StatusCode.InvalidArgument);
              assert.ok(e.message.includes(`Pivot.setFieldNumericGroup group.${key} must be finite`), e.message);
              return true;
            },
          );
        }
      }
      expectCode(() => pivot.setFieldNumericGroup('Amount', { interval: 0 }), StatusCode.InvalidArgument);
      const inverted = { interval: 10, start: 50, end: 20 };
      expectCode(() => pivot.setFieldNumericGroup('Amount', inverted), StatusCode.InvalidArgument);
      assert.equal(pivot.getFieldNumericGroup('Amount'), null);
      assert.equal(pivot.listFields().length, 2);
    });
  });

  it('refuses a field of the wrong type with PreconditionFailed', () => {
    withSheet((ws, wb) => {
      const { pivot } = build(ws, wb, NUMBERS);
      expectCode(() => pivot.setFieldNumericGroup('Region', { interval: 10 }), StatusCode.PreconditionFailed);
      const years = { periods: f.PivotDatePeriod.Years };
      expectCode(() => pivot.setFieldDateGroup('Region', years), StatusCode.PreconditionFailed);
      assert.equal(pivot.getFieldNumericGroup('Region'), null);
      assert.equal(pivot.listFields().length, 2);
    });
  });

  it('requires the bound an explicit false auto flag asks for', () => {
    withSheet((ws, wb) => {
      const { pivot } = build(ws, wb, NUMBERS);
      for (const key of ['start', 'end']) {
        const auto = `auto${key[0].toUpperCase()}${key.slice(1)}`;
        assert.throws(
          () => pivot.setFieldNumericGroup('Amount', { interval: 10, [auto]: false }),
          (e) => {
            assert.ok(e instanceof TypeError);
            assert.equal(e.message, `Pivot.setFieldNumericGroup group.${key} is required`);
            return true;
          },
        );
      }
      assert.equal(pivot.getFieldNumericGroup('Amount'), null);
    });
  });

  it('rejects malformed groups', () => {
    withSheet((ws, wb) => {
      const { pivot } = build(ws, wb, NUMBERS);
      assert.throws(() => pivot.setFieldNumericGroup('Amount', { interval: 'x' }), TypeError);
      assert.throws(() => pivot.setFieldNumericGroup('Amount', { interval: 10, start: 'x' }), TypeError);
      assert.throws(() => pivot.setFieldNumericGroup('Amount', { interval: 10, autoStart: 1 }), TypeError);
      assert.throws(() => pivot.setFieldNumericGroup('Amount', { interval: 10, bogus: 1 }), TypeError);
      assert.throws(() => pivot.setFieldNumericGroup('Amount', 5), TypeError);
    });
  });
});
