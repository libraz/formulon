// Returned keys, refused arguments and name folding on the public surface.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

async function withWorkbook(fn) {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

// Asserts that `fn` fails with a FormulonError carrying `code`.
function rejectsWith(fn, code, label) {
  assert.throws(
    fn,
    (e) => {
      assert.equal(e.name, 'FormulonError', `${label}: expected a FormulonError, got ${e}`);
      assert.equal(e.code, code, `${label}: code ${e.code}`);
      return true;
    },
    label,
  );
}

const GUID = /^\{[0-9A-F]{8}(-[0-9A-F]{4}){3}-[0-9A-F]{12}\}$/;
const PERSON = '{0A1B2C3D-0000-4000-8000-00000000ABCD}';

// A pivot over four records: Region and Product carry shared items, Amount is plain.
function pivotOf(wb, f) {
  const ws = wb.sheet(0);
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ range: 'A1:C5', sheet: 'Sheet1' });
  for (const name of ['Region', 'Product', 'Amount']) cache.addField(name);
  for (const item of ['East', 'West']) cache.addFieldSharedItem(0, item);
  for (const item of ['Apple', 'Pear']) cache.addFieldSharedItem(1, item);
  for (const [region, product, amount] of [
    [0, 0, 10],
    [1, 0, 30],
    [0, 1, 5],
    [1, 1, 7],
  ]) {
    const rec = cache.addRecord();
    cache.setRecordValue(rec, 0, region);
    cache.setRecordValue(rec, 1, product);
    cache.setRecordValue(rec, 2, amount);
  }
  const pv = ws.addPivot({ name: 'Pivot1', cacheId: cache.cacheId, anchor: { row: 0, col: 6 } });
  return { cache, pv, f };
}

const rect = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });
const rectOf = (table) => ({ ...table.getProperties().range.rect });

function salesTable(ws, range = 'A1:B3') {
  ws.setValues('A1', [
    ['Item', 'Qty'],
    ['East', 10],
    ['West', 30],
  ]);
  return ws.addTable({ name: 'Sales', range });
}

// ---- Enum and exception names ----

test('PivotAxis spells its column member out and an added field lands on the column axis', async () => {
  await withWorkbook((wb, f) => {
    assert.equal(f.PivotAxis.Column, 1);
    assert.equal('Col' in f.PivotAxis, false);
    const { cache, pv } = pivotOf(wb, f);
    pv.addField({ fieldIndex: 'Product', axis: f.PivotAxis.Column });
    assert.deepEqual(pv.getColumnFieldOrder(), [1]);
    assert.deepEqual(pv.getRowFieldOrder(), []);
    assert.ok(cache.cacheId >= 0);
  });
});

test('a formula parse failure carries a parser status and not InvalidArgument', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    let caught = null;
    try {
      wb.sheet(0).setFormula('A1', '=1+');
    } catch (e) {
      caught = e;
    }
    assert.ok(caught, 'the formula is refused');
    assert.equal(caught.name, 'FormulonError');
    assert.notEqual(caught.code, StatusCode.InvalidArgument);
    assert.ok(caught.code >= 1000 && caught.code <= 1014, `parser status, got ${caught.code}`);
  });
});

// ---- The locale error-name list ----

test('locale error names are read through getLocaleFacts only', async () => {
  const f = await getModule();
  assert.equal('listLocaleErrorNames' in f, false);
  const names = f.getLocaleFacts('mac-365-de_DE').errorNames;
  assert.equal(names.find((e) => e.canonical === '#N/A').localized, '#NV');
});

// ---- Unknown keys at every depth ----

test('unknown keys are refused as TypeError in option objects, civil times, date systems and colours', async () => {
  await withWorkbook((wb, f) => {
    const { pv } = pivotOf(wb, f);
    const index = pv.addField({ fieldIndex: 'Region', axis: f.PivotAxis.Row });
    assert.throws(() => pv.addFieldItem(index, 'East', { visble: false }), TypeError);

    const civil = { year: 2024, month: 1, day: 1, hour: 0, minute: 0, second: 0 };
    assert.equal(typeof f.convertCivilToSerial(civil), 'number');
    assert.throws(() => f.convertCivilToSerial(civil, { date1904: true, bogus: 1 }), TypeError);
    assert.throws(() => wb.setPinnedNow({ ...civil, bogus: 1 }), TypeError);

    const rgb = { kind: f.ColorKind.Rgb, argb: 0xffff0000, theme: 0, tint: 0, indexed: 0 };
    wb.sheet(0).setTabColor(rgb);
    assert.throws(() => wb.sheet(0).setTabColor({ ...rgb, bogus: 1 }), TypeError);
  });
});

test('convertCivilToSerial refuses a field a civil time does not have', async () => {
  const f = await getModule();
  const civil = { year: 2024, month: 1, day: 1, hour: 0, minute: 0, second: 0 };
  assert.throws(() => f.convertCivilToSerial({ ...civil, hourz: 12 }), TypeError);
});

test('a nested chart patch refuses an unknown key', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('A1', [
      ['x', 'y'],
      [1, 2],
      [2, 4],
    ]);
    const chart = ws.addChart({ type: f.ChartPlotType.Line, range: 'A1:B3' });
    const id = chart.getProperties().plots[0].series[0].id;
    assert.throws(() => chart.updateSeries(0, id, { line: { width: 2 } }), TypeError);
  });
});

// ---- Chart title at insert ----

test('addChart treats an omitted title as automatic, null as none and an empty string as InvalidArgument', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('A1', [
      ['x', 'y'],
      [1, 2],
      [2, 4],
    ]);
    const auto = ws.addChart({ type: f.ChartPlotType.Line, range: 'A1:B3' });
    assert.equal(auto.getProperties().title.kind, f.ChartTitleKind.Auto);
    const none = ws.addChart({ type: f.ChartPlotType.Line, range: 'A1:B3', title: null });
    assert.equal(none.getProperties().title, null);
    const text = ws.addChart({ type: f.ChartPlotType.Line, range: 'A1:B3', title: 'Sales' });
    assert.equal(text.getProperties().title.text, 'Sales');
    const before = ws.listDrawings().length;
    rejectsWith(
      () => ws.addChart({ type: f.ChartPlotType.Line, range: 'A1:B3', title: '' }),
      f.StatusCode.InvalidArgument,
      'empty insert title',
    );
    assert.equal(ws.listDrawings().length, before);
    rejectsWith(() => none.setTitle(''), f.StatusCode.InvalidArgument, 'empty set title');
  });
});

// ---- The table totals row grows and shrinks the table ----

test('a totals row toggle grows the table by one row below and shrinks it back', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    table.setColumnTotals(1, { totalsRowFunction: 1 });
    table.update({ totalsRow: true });
    const props = table.getProperties();
    assert.equal(props.totalsRow, true);
    assert.deepEqual({ ...props.range.rect }, rect(0, 0, 3, 1));
    assert.equal(ws.getValue('B3'), 30);
    assert.equal(ws.getValue('B4'), 40);
    table.update({ totalsRow: false });
    assert.deepEqual(rectOf(table), rect(0, 0, 2, 1));
    assert.equal(table.getProperties().totalsRow, false);
  });
});

test('a totals row over an occupied row or on the last grid row is PreconditionFailed and changes nothing', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    ws.setValue('B4', 'keep');
    rejectsWith(() => table.update({ totalsRow: true }), StatusCode.PreconditionFailed, 'occupied row below');
    assert.deepEqual(rectOf(table), rect(0, 0, 2, 1));
    assert.equal(table.getProperties().totalsRow, false);
    assert.equal(ws.getValue('B4'), 'keep');

    const edge = wb.addSheet('Edge');
    edge.setValues('A1048575', [
      ['Item', 'Qty'],
      ['East', 1],
    ]);
    const last = edge.addTable({ name: 'Edge1', range: 'A1048575:B1048576' });
    rejectsWith(() => last.update({ totalsRow: true }), StatusCode.PreconditionFailed, 'last grid row');
    assert.deepEqual(rectOf(last), rect(1048574, 0, 1048575, 1));
    assert.equal(last.getProperties().totalsRow, false);
  });
});

test('a totals flag with an explicit range only flips the flag', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    table.update({ range: 'A1:B3', totalsRow: true });
    const props = table.getProperties();
    assert.equal(props.totalsRow, true);
    assert.deepEqual({ ...props.range.rect }, rect(0, 0, 2, 1));
  });
});

test('a refused table patch leaves the range and flags as they were', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    rejectsWith(
      () => table.update({ range: 'A1:C9', totalsRow: true }),
      StatusCode.InvalidArgument,
      'range wider than the columns',
    );
    assert.deepEqual(rectOf(table), rect(0, 0, 2, 1));
    assert.equal(table.getProperties().totalsRow, false);
  });
});

// ---- Merges ----

test('addMerge returns the stored rectangle and refuses an overlapping merge as AlreadyExists', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    assert.deepEqual({ ...ws.addMerge('A1:B2') }, { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 });
    rejectsWith(() => ws.addMerge('B2:C3'), StatusCode.AlreadyExists, 'overlap');
    rejectsWith(() => ws.addMerge('A1:B2'), StatusCode.AlreadyExists, 'identical');
    assert.deepEqual({ ...ws.addMerge('C1:D2') }, { firstRow: 0, firstCol: 2, lastRow: 1, lastCol: 3 });
    assert.equal(ws.listMerges().length, 2);
  });
});

test('a merge over a table or a data table is PreconditionFailed', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    salesTable(ws);
    rejectsWith(() => ws.addMerge('A1:B2'), StatusCode.PreconditionFailed, 'table');
    ws.addDataTable({ range: 'E2:E3', columnInput: 'D1' });
    rejectsWith(() => ws.addMerge('E2:F3'), StatusCode.PreconditionFailed, 'data table');
    assert.equal(ws.listMerges().length, 0);
  });
});

// ---- Returned keys ----

test('the adds that used to return nothing return their key or index', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.deepEqual({ ...ws.addDataTable({ range: 'B2:D6', columnInput: 'A1' }) }, { row: 1, col: 1 });
    assert.equal(wb.addPerson({ id: PERSON.toLowerCase(), displayName: 'A' }), PERSON);
    const { cache } = pivotOf(wb, f);
    const count = cache.listFieldSharedItems(0).length;
    assert.equal(cache.addFieldSharedItem(0, 'Q'), count);
    assert.equal(cache.addFieldSharedItem(0, 'R'), count + 1);
    assert.equal(cache.listFieldSharedItems(0).length, count + 2);
  });
});

// ---- Conditional formatting rules and protected ranges by key ----

function cellIs(wb, extra = {}) {
  return {
    ranges: 'A1:A5',
    type: 1,
    operator: 5,
    formula1: '10',
    dxfIndex: wb.ooxml.addDxf({ font: { bold: true } }),
    ...extra,
  };
}

test('addCfRule returns the rule id and the id addresses the rule', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    const id = ws.addCfRule(cellIs(wb));
    assert.equal(typeof id, 'string');
    assert.match(id, GUID);
    assert.equal(ws.listCfRules()[0].id, id);
    ws.updateCfRule(id, { stopIfTrue: true });
    assert.equal(ws.listCfRules()[0].stopIfTrue, true);
    rejectsWith(() => ws.removeCfRule('{00000000-0000-0000-0000-000000000000}'), StatusCode.NotFound, 'unknown id');
    rejectsWith(() => ws.removeCfRule(5), StatusCode.IndexOutOfRange, 'position past the end');
    ws.removeCfRule(id);
    assert.equal(ws.listCfRules().length, 0);
  });
});

test('a position still addresses a CF rule and the rule keeps its id', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const first = ws.addCfRule(cellIs(wb));
    const second = ws.addCfRule(cellIs(wb, { formula1: '20' }));
    assert.notEqual(first, second);
    ws.removeCfRule(0);
    assert.deepEqual(
      ws.listCfRules().map((r) => r.id),
      [second],
    );
  });
});

test('a CF rule id is a brace-wrapped GUID or it is refused', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    rejectsWith(() => ws.addCfRule(cellIs(wb, { id: 'abc' })), StatusCode.InvalidArgument, 'not a GUID');
    assert.equal(ws.listCfRules().length, 0);
    const own = '{abcdef01-2345-4678-89ab-cdef01234567}';
    const kept = ws.addCfRule(cellIs(wb, { id: own }));
    assert.equal(kept.toUpperCase(), own.toUpperCase());
    const again = ws.addCfRule(cellIs(wb, { id: own }));
    assert.notEqual(again.toUpperCase(), own.toUpperCase());
    assert.match(again, GUID);
    const omitted = ws.addCfRule(cellIs(wb));
    assert.match(omitted, GUID);
  });
});

test('addProtectedRange returns the name and the name addresses the range', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    assert.equal(ws.addProtectedRange({ name: 'P', ranges: 'A1' }), 'P');
    rejectsWith(() => ws.addProtectedRange({ name: 'P', ranges: 'B1' }), StatusCode.AlreadyExists, 'duplicate');
    ws.updateProtectedRange('P', { ranges: 'A1:B2' });
    assert.equal(ws.listProtectedRanges()[0].name, 'P');
    rejectsWith(() => ws.removeProtectedRange('Nope'), StatusCode.NotFound, 'unknown name');
    rejectsWith(() => ws.removeProtectedRange(3), StatusCode.IndexOutOfRange, 'position past the end');
    ws.removeProtectedRange('P');
    assert.equal(ws.listProtectedRanges().length, 0);
  });
});

// ---- Page-axis field order ----

test('the page field order reads back, moves fields onto the page axis and refuses the Values level', async () => {
  await withWorkbook((wb, f) => {
    const { pv } = pivotOf(wb, f);
    const region = pv.addField({ fieldIndex: 'Region', axis: f.PivotAxis.Row });
    const product = pv.addField({ fieldIndex: 'Product', axis: f.PivotAxis.Page });
    const amount = pv.addField({ fieldIndex: 'Amount', axis: f.PivotAxis.Page });
    assert.deepEqual(pv.getPageFieldOrder(), [product, amount]);
    pv.setPageFieldOrder([amount, product]);
    assert.deepEqual(pv.getPageFieldOrder(), [amount, product]);
    pv.setPageFieldOrder([region]);
    assert.deepEqual(pv.getPageFieldOrder(), [region]);
    assert.deepEqual(pv.getRowFieldOrder(), []);
    rejectsWith(() => pv.setPageFieldOrder([f.PIVOT_VALUES_FIELD]), f.StatusCode.InvalidArgument, 'values level');
    rejectsWith(() => pv.setPageFieldOrder([product, product]), f.StatusCode.InvalidArgument, 'duplicate');
    rejectsWith(() => pv.setPageFieldOrder([99]), f.StatusCode.IndexOutOfRange, 'past the end');
    assert.deepEqual(pv.getPageFieldOrder(), [region]);
    pv.setPageFieldOrder([]);
    assert.deepEqual(pv.getPageFieldOrder(), []);
  });
});

// ---- Name folding ----

test('custom table styles, defined names and table columns match by the one case fold', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    wb.setTableStyle({ name: 'Ä1', table: true });
    assert.equal(wb.getTableStyle('ä1').name, 'Ä1');
    wb.setTableStyle({ name: 'ä1', table: true });
    assert.equal(wb.listTableStyles().filter((s) => s.name.toLowerCase() === 'ä1').length, 1);

    wb.setDefinedName({ name: 'Ä', formula: '=1' });
    assert.equal(wb.getDefinedName('ä').name, 'Ä');

    const ws = wb.sheet(0);
    ws.setValues('A1', [
      ['Ä', 'B'],
      [1, 2],
    ]);
    const table = ws.addTable({ name: 'Tbl', range: 'A1:B2' });
    const renamed = table.renameColumn(1, 'ä');
    assert.notEqual(renamed.toLowerCase(), 'ä');
    rejectsWith(() => ws.addTable({ name: 'tbl', range: 'D1:E2' }), StatusCode.AlreadyExists, 'table name');
  });
});

// ---- Documented reads ----

test('evaluateFormula reads an empty cell as zero while getValue reads it as null', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    assert.equal(ws.evaluateFormula('A2', '=B9'), 0);
    assert.equal(ws.getValue('B9'), null);
  });
});

test('getDate refuses a serial outside the date system', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    ws.setNumber('A1', 0.5);
    rejectsWith(() => ws.getDate('A1'), StatusCode.InvalidArgument, 'serial below 1');
  });
});

test('a view patch mixing a freeze and a split is InvalidArgument', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    rejectsWith(
      () => wb.sheet(0).updateView({ freezeRows: 1, splitXTwips: 100 }),
      StatusCode.InvalidArgument,
      'freeze and split',
    );
  });
});
