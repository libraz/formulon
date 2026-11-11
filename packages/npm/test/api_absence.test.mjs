// Absence on the public surface: a read reports `null` for a value that is not there, a read record goes back
// to its writer unchanged, and removed spellings stay removed.

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

const isFormulonError = (code) => (e) => e.name === 'FormulonError' && e.code === code;

/** A pivot over two records (Region, Amount) with Region on rows; returns the cache and the pivot. */
function pivotOver(ws, wb, f) {
  const cache = wb.addPivotCache();
  cache.addField('Region');
  cache.addField('Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const record = cache.addRecord();
    cache.setRecordValue(record, 0, region);
    cache.setRecordValue(record, 1, amount);
  }
  const pivot = ws.addPivot({ name: 'P', cacheId: cache.cacheId, anchor: 'E1' });
  pivot.setFieldAxis('Region', f.PivotAxis.Row);
  pivot.setFieldAxis('Amount', f.PivotAxis.Value);
  const amount = 1;
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: f.PivotAggregation.Sum });
  return { cache, pivot, amount };
}

test('a cache id is assigned when omitted or null and an explicit 0 is refused', async () => {
  await withWorkbook((wb, f) => {
    const omitted = wb.addPivotCache();
    const viaNull = wb.addPivotCache(null);
    assert.ok(omitted.cacheId >= 1);
    assert.ok(viaNull.cacheId >= 1);
    assert.notEqual(omitted.cacheId, viaNull.cacheId);
    assert.throws(() => wb.addPivotCache(0), isFormulonError(f.StatusCode.InvalidArgument));
    assert.equal(wb.addPivotCache(41).cacheId, 41);
  });
});

test('a cache worksheet source has no present key and an input carrying one is refused', async () => {
  await withWorkbook((wb) => {
    const cache = wb.addPivotCache();
    cache.setWorksheetSource({ range: 'A1:B3', sheet: 'Sheet1' });
    const source = cache.getWorksheetSource();
    assert.equal('present' in source, false);
    assert.equal(source.name, null);
    assert.throws(() => cache.setWorksheetSource({ range: 'A1:B3', sheet: 'Sheet1', present: true }), TypeError);
    cache.setWorksheetSource(source);
    cache.setWorksheetSource(null);
    assert.equal(cache.getWorksheetSource(), null);
  });
});

test('the date-group clear method is gone and a null group ungroups', async () => {
  await withWorkbook((wb, f) => {
    const { pivot } = pivotOver(wb.sheet(0), wb, f);
    assert.equal('clearFieldDateGroup' in pivot, false);
    assert.equal(typeof pivot.clearFieldDateGroup, 'undefined');
    // Ungrouping a field that is not grouped is accepted.
    pivot.setFieldDateGroup(0, null);
    assert.equal(pivot.listFields().length, 2);
  });
});

test('the show-as base members are named and the module constants are gone', async () => {
  const f = await getModule();
  assert.equal(f.PivotShowAsBase.Previous, 1048828);
  assert.equal(f.PivotShowAsBase.Next, 1048829);
  assert.ok(Object.isFrozen(f.PivotShowAsBase));
  assert.equal('PIVOT_SHOW_AS_BASE_PREVIOUS' in f, false);
  assert.equal('PIVOT_SHOW_AS_BASE_NEXT' in f, false);
});

test('a data field reads the show-as base it was given and null where it has none', async () => {
  await withWorkbook((wb, f) => {
    const { pivot, amount } = pivotOver(wb.sheet(0), wb, f);
    const plain = pivot.listDataFields()[0];
    assert.equal(plain.showAsBaseField, null);
    assert.equal(plain.showAsBaseItem, null);
    pivot.addDataField({
      name: 'Diff',
      fieldIndex: amount,
      aggregation: f.PivotAggregation.Sum,
      showAs: f.PivotShowValuesAs.DifferenceFrom,
      showAsBaseField: 0,
      showAsBaseItem: f.PivotShowAsBase.Next,
    });
    const read = pivot.listDataFields().find((d) => d.name === 'Diff');
    assert.equal(read.showAsBaseField, 0);
    assert.equal(read.showAsBaseItem, f.PivotShowAsBase.Next);
  });
});

test('a pivot field with no axis reads null and the -1 sentinel is refused', async () => {
  await withWorkbook((wb, f) => {
    const { pivot } = pivotOver(wb.sheet(0), wb, f);
    pivot.setFieldAxis(0, null);
    assert.equal(pivot.listFields()[0].axis, null);
    assert.throws(() => pivot.setFieldAxis(0, -1), isFormulonError(f.StatusCode.InvalidArgument));
  });
});

test('the pivot layout names a field null where it has none', async () => {
  await withWorkbook((wb, f) => {
    const { pivot } = pivotOver(wb.sheet(0), wb, f);
    const cells = pivot.getLayout().cells;
    assert.ok(cells.length > 0);
    for (const cell of cells) {
      assert.ok(cell.fieldName === null || typeof cell.fieldName === 'string');
      assert.notEqual(cell.fieldName, '');
    }
  });
});

test('an unstyled cell reads a null xf index; an explicit zero reads 0 only over an axis format', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    assert.equal(ws.ooxml.getCellXfIndex('Z99'), null);
    assert.equal(ws.ooxml.getCellXfIndex(98, 25), null);
    // With no row or column format an explicit default overrides nothing.
    ws.ooxml.setCellXfIndex('Z99', 0);
    assert.equal(ws.ooxml.getCellXfIndex('Z99'), null);
    const xf = wb.ooxml.addCellXf({ ...wb.ooxml.getCellXf(0), numFmtId: 10 });
    ws.ooxml.setColumnXfIndex('Z', xf);
    ws.ooxml.setCellXfIndex('Z99', 0);
    assert.equal(ws.ooxml.getCellXfIndex('Z99'), 0);
    assert.equal(ws.ooxml.getCellXfIndex('Z98'), null);
  });
});

test('getSpill and getArrayFormula read null outside a region and the spill record has no engaged key', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    assert.equal(ws.getSpill('Z1'), null);
    assert.equal(ws.getSpill(0, 25), null);
    assert.equal(ws.getArrayFormula('Z1'), null);
    ws.setFormula('A1', '=SEQUENCE(3,1)');
    wb.recalc();
    const spill = ws.getSpill('A2');
    assert.equal(spill.blocked, false);
    assert.deepEqual({ ...spill.range }, { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 });
    assert.equal('engaged' in spill, false);
    assert.equal('getSpillInfo' in ws, false);
  });
});

test('a conditional-format rule reads every key and null for the ones its type does not use', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.addCfRule({ type: f.CfRuleType.Expression, ranges: 'A1', formula1: 'A1>1' });
    const rule = ws.listCfRules()[0];
    for (const key of [
      'rank',
      'percent',
      'bottom',
      'text',
      'timePeriod',
      'colorScale',
      'dataBar',
      'iconSet',
      'dxfIndex',
    ]) {
      assert.equal(key in rule, true, key);
      assert.equal(rule[key], null, key);
    }
    assert.equal(rule.formula1, '=A1>1');
    assert.equal(rule.formula2, null);
    // The read record goes back to the writer.
    ws.addCfRule(rule);
    assert.equal(ws.listCfRules().length, 2);
  });
});

test('a differential format reads null for the groups it does not carry and passes back to addDxf', async () => {
  await withWorkbook((wb) => {
    const fontOnly = wb.ooxml.addDxf({ font: { bold: true } });
    const read = wb.ooxml.getDxf(fontOnly);
    assert.equal(read.fill, null);
    assert.equal(read.alignment, null);
    assert.equal(read.protection, null);
    assert.equal(wb.ooxml.addDxf(read), fontOnly);
  });
});

test('a cell format reads null for attributes it leaves unstated and passes back to addCellXf', async () => {
  await withWorkbook((wb) => {
    const xf = wb.ooxml.getCellXf(0);
    for (const key of ['horizontal', 'vertical', 'wrapText', 'textRotation']) {
      assert.equal(key in xf, true, key);
    }
    assert.equal(xf.textRotation, null);
    assert.equal(wb.ooxml.addCellXf(xf), 0);
  });
});

test('page setup has no orientationStated key and a stated orientation reads back', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.equal('orientationStated' in ws.getPageSetup(), false);
    assert.equal(ws.getPageSetup().orientation, null);
    ws.updatePageSetup({ orientation: f.PageOrientation.Landscape });
    assert.equal(ws.getPageSetup().orientation, f.PageOrientation.Landscape);
  });
});

test('phonetic properties report stated and carry no present key', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setText('A1', '漢字');
    ws.setCellPhoneticRuns('A1', [{ start: 0, end: 2, text: 'かんじ' }]);
    const before = ws.getCellPhoneticProperties('A1');
    assert.equal(before.stated, false);
    assert.equal('present' in before, false);
    ws.setCellPhoneticProperties('A1', { fontIndex: 0, type: 0, alignment: 0 });
    assert.equal(ws.getCellPhoneticProperties('A1').stated, true);
    // The read record, stated key included, goes back unchanged.
    ws.setCellPhoneticProperties('A1', ws.getCellPhoneticProperties('A1'));
    assert.equal(ws.getCellPhoneticProperties('A1').stated, true);
  });
});

test('an auto-filter column reads its differential format index and null for a values column', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('A1', [['h'], ['a'], ['b']]);
    ws.setAutoFilter({
      range: 'A1:A3',
      columns: [{ columnOffset: 0, kind: f.FilterKind.Values, values: ['a'] }],
    });
    const read = ws.getAutoFilter();
    assert.equal(read.columns[0].dxfIndex, null);
    ws.setAutoFilter(read);
    assert.equal(ws.getAutoFilter().columns[0].dxfIndex, null);
  });
});

test('function metadata reads null for an open upper arity and for absent display text', async () => {
  const f = await getModule();
  const sum = f.getFunctionMetadata('SUM');
  assert.equal(sum.maxArity, null);
  assert.equal(sum.signature, null);
  assert.equal(sum.description, null);
  assert.equal(f.getFunctionMetadata('ABS').maxArity, 1);
});

test('evaluateCfRange without time-period rules matches no time-period rule', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setNumber('A1', 43845);
    ws.addCfRule({ type: f.CfRuleType.TimePeriod, ranges: 'A1', timePeriod: f.DatePeriod.Today });
    const today = { year: 2020, month: 1, day: 15, hour: 12, minute: 0, second: 0 };
    assert.equal(ws.evaluateCfRange('A1', { today }).length, 1);
    assert.equal(ws.evaluateCfRange('A1', { today, timePeriodRules: true }).length, 1);
    assert.equal(ws.evaluateCfRange('A1', { timePeriodRules: false }).length, 0);
    assert.throws(
      () => ws.evaluateCfRange('A1', { today, timePeriodRules: false }),
      isFormulonError(f.StatusCode.InvalidArgument),
    );
    assert.throws(() => ws.evaluateCfRange('A1', { timePeriodRule: false }), TypeError);
  });
});

test('a non-XY chart series reads a null x number format', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setText(0, 1, 'S1');
    for (let r = 1; r <= 3; r++) {
      ws.setText(r, 0, `c${r}`);
      ws.setNumber(r, 1, r);
    }
    const chart = ws.addChart({ type: f.ChartPlotType.Bar, range: 'Sheet1!$A$1:$B$4' });
    const series = ws.chart(chart.objectId).getData().series[0];
    assert.equal(series.xNumberFormat, null);
    assert.equal(series.numberFormat, 'General');
  });
});
