// Empty strings in the PIVOT setters: where an empty caption is a stored state (item captions, data field names) it
// is a value and `null` clears or derives; where the file holds one state for both (field names and number formats,
// option captions) they mean the same; a worksheet source attribute may not be empty.

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
const { PivotAxis, PivotAggregation, StatusCode, WorkbookFormat } = f;

const ROWS = [
  ['East', 10],
  ['West', 30],
  ['East', 5],
];

/** Runs `body(wb, pivot, cache)` on a default workbook holding a refreshed, projected pivot; disposes it after. */
function withPivot(body) {
  const wb = f.Workbook.createDefault();
  try {
    const ws = wb.sheet(0);
    ws.setText('A1', 'Region');
    ws.setText('B1', 'Amount');
    ROWS.forEach(([region, amount], i) => {
      ws.setText(i + 1, 0, region);
      ws.setNumber(i + 1, 1, amount);
    });
    const cache = wb.addPivotCache();
    cache.setWorksheetSource({ range: 'A1:B4', sheet: 'Sheet1' });
    for (const name of ['Region', 'Amount']) cache.addField(name);
    for (const item of ['East', 'West']) cache.addFieldSharedItem(0, item);
    ROWS.forEach(([region, amount]) => {
      const record = cache.addRecord();
      cache.setRecordValue(record, 0, region === 'East' ? 0 : 1);
      cache.setRecordValue(record, 1, amount);
    });
    const pivot = ws.addPivot({ name: 'PT', cacheId: cache.cacheId, anchor: 'F1' });
    pivot.setFieldAxis('Region', PivotAxis.Row);
    pivot.setFieldAxis('Amount', PivotAxis.Value);
    pivot.addDataField({ name: 'Total', fieldIndex: 'Amount', aggregation: PivotAggregation.Sum });
    cache.refresh();
    pivot.project();
    return body(wb, pivot, cache);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

/** Saves `wb` as `format`, reloads it and passes the reloaded first pivot to `body`. */
function reloaded(wb, format, body) {
  const back = f.Workbook.load(wb.save(format));
  try {
    return body(back.sheet(0).pivot(0));
  } finally {
    back.dispose();
  }
}

function expectCode(fn, code) {
  assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, code, `code ${e.code}`);
    return true;
  });
}

describe('field custom name and number format: "" and null are one state', () => {
  for (const cleared of ['', null]) {
    it(`customName ${JSON.stringify(cleared)} restores the source name`, () => {
      withPivot((_wb, pivot) => {
        pivot.updateField('Region', { customName: 'Area' });
        assert.equal(pivot.listFields()[0].customName, 'Area');
        pivot.updateField('Region', { customName: cleared });
        assert.equal(pivot.listFields()[0].customName, null);
        assert.equal(pivot.listFields()[0].sourceName, 'Region');
      });
    });

    it(`numberFormat ${JSON.stringify(cleared)} means none`, () => {
      withPivot((_wb, pivot) => {
        pivot.updateField('Amount', { numberFormat: '0.00' });
        assert.equal(pivot.listFields()[1].numberFormat, '0.00');
        pivot.updateField('Amount', { numberFormat: cleared });
        assert.equal(pivot.listFields()[1].numberFormat, null);
      });
    });
  }
});

describe('item caption: "" is an empty label, null un-renames', () => {
  it('reads back both states', () => {
    withPivot((_wb, pivot) => {
      assert.equal(pivot.getFieldItem('Region', 0).caption, null);
      pivot.updateFieldItem('Region', 0, { caption: '' });
      assert.equal(pivot.getFieldItem('Region', 0).caption, '');
      pivot.updateFieldItem('Region', 0, { caption: null });
      assert.equal(pivot.getFieldItem('Region', 0).caption, null);
      pivot.updateFieldItem('Region', 1, { caption: 'Ouest' });
      pivot.updateFieldItem('Region', 1, { caption: '' });
      assert.equal(pivot.getFieldItem('Region', 1).caption, '');
    });
  });

  for (const [label, format] of [
    ['xlsx', WorkbookFormat.Xlsx],
  ]) {
    it(`the empty label survives an ${label} save and reload`, () => {
      withPivot((wb, pivot) => {
        pivot.updateFieldItem('Region', 0, { caption: '' });
        pivot.updateFieldItem('Region', 1, { caption: 'Ouest' });
        reloaded(wb, format, (back) => {
          assert.equal(back.getFieldItem('Region', 0).caption, '');
          assert.equal(back.getFieldItem('Region', 1).caption, 'Ouest');
        });
        pivot.updateFieldItem('Region', 0, { caption: null });
        reloaded(wb, format, (back) => assert.equal(back.getFieldItem('Region', 0).caption, null));
      });
    });
  }
});

describe('data field name: "" is an empty caption, null derives', () => {
  it('reads back both states', () => {
    withPivot((_wb, pivot) => {
      const derived = pivot.getDataFieldDefaultName('Amount', PivotAggregation.Sum, 0);
      assert.notEqual(derived, '');
      pivot.setDataField(0, { name: '', fieldIndex: 'Amount' });
      assert.equal(pivot.getDataField(0).name, '');
      pivot.setDataField(0, { name: null, fieldIndex: 'Amount' });
      assert.equal(pivot.getDataField(0).name, derived);
      pivot.setDataField(0, { fieldIndex: 'Amount' });
      assert.equal(pivot.getDataField(0).name, derived);
      const added = pivot.addDataField({ name: '', fieldIndex: 'Amount', aggregation: PivotAggregation.Count });
      assert.equal(pivot.getDataField(added).name, '');
    });
  });

  for (const [label, format] of [
    ['xlsx', WorkbookFormat.Xlsx],
  ]) {
    it(`the empty caption survives an ${label} save and reload`, () => {
      withPivot((wb, pivot) => {
        pivot.setDataField(0, { name: '', fieldIndex: 'Amount' });
        reloaded(wb, format, (back) => {
          assert.equal(back.listDataFields().length, 1);
          assert.equal(back.getDataField(0).name, '');
        });
      });
    });
  }

  it('numberFormat "" and null both mean none', () => {
    withPivot((_wb, pivot) => {
      for (const none of ['', null]) {
        pivot.setDataField(0, { name: 'Total', fieldIndex: 'Amount', numberFormat: '0.00' });
        assert.equal(pivot.getDataField(0).numberFormat, '0.00');
        pivot.setDataField(0, { name: 'Total', fieldIndex: 'Amount', numberFormat: none });
        assert.equal(pivot.getDataField(0).numberFormat, null);
      }
    });
  });
});

describe('option captions: "" and null both unset', () => {
  for (const unset of ['', null]) {
    it(`${JSON.stringify(unset)} unsets every caption and alt text`, () => {
      withPivot((_wb, pivot) => {
        const keys = ['errorCaption', 'missingCaption', 'altText', 'altTextSummary', 'grandTotalCaption'];
        pivot.updateOptions(Object.fromEntries(keys.map((key) => [key, 'x'])));
        for (const key of keys) assert.equal(pivot.getOptions()[key], 'x', key);
        pivot.updateOptions(Object.fromEntries(keys.map((key) => [key, unset])));
        for (const key of keys) assert.equal(pivot.getOptions()[key], null, key);
      });
    });
  }
});

describe('worksheet source: "" is refused, null is absent', () => {
  it('fails each empty attribute with InvalidArgument and keeps the source', () => {
    withPivot((_wb, _pivot, cache) => {
      const refused = [{ range: 'A1:B4', sheet: '' }, { range: 'A1:B4', sheet: 'Sheet1', name: '' }, { name: '' }];
      for (const source of refused) {
        expectCode(() => cache.setWorksheetSource(source), StatusCode.InvalidArgument);
      }
      assert.equal(cache.getWorksheetSource().range.sheet, 0);
      cache.setWorksheetSource({ range: 'A1:B4', sheet: 'Sheet1', name: null });
      assert.equal(cache.getWorksheetSource().name, null);
    });
  });
});
