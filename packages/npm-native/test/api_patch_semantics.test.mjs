// Patch and record semantics of the chart and pivot handles: an empty patch changes nothing, only `null` clears,
// a patch is validated as a whole before anything applies, and every writable field a read record shows has a patch
// key.

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

/** Runs `body(ws, wb)` on a fresh default workbook and disposes it. */
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

const expectInvalid = (fn) => expectCode(fn, f.StatusCode.InvalidArgument);
const expectTypeError = (fn) => assert.throws(fn, (e) => e instanceof TypeError, 'expected a TypeError');

const plain = (value) => JSON.parse(JSON.stringify(value) ?? 'null');
const pick = (record, keys) => Object.fromEntries(keys.map((k) => [k, record[k]]));

/** Sheet1: B1:D1 series names, A2:A5 categories, B2:D5 numbers. */
function fillGrid(ws) {
  ['S1', 'S2', 'S3'].forEach((name, c) => {
    ws.setText(0, c + 1, name);
  });
  ['a', 'b', 'c', 'd'].forEach((category, r) => {
    ws.setText(r + 1, 0, category);
    for (let c = 0; c < 3; c += 1) ws.setNumber(r + 1, c + 1, (r + 1) * (c + 2));
  });
}

/** A column chart over A1:D5 with a Right legend; returns the chart handle. */
function addColumnChart(ws, extra = {}) {
  fillGrid(ws);
  return ws.addChart({
    type: f.ChartPlotType.Bar,
    range: 'Sheet1!A1:D5',
    legend: f.ChartLegendPosition.Right,
    ...extra,
  });
}

describe('a chart point override', () => {
  const fillRed = { fill: { color: '#FF0000' } };

  it('survives an empty patch', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      chart.updatePoint(0, 0, 1, fillRed);
      assert.equal(chart.getProperties().plots[0].series[0].points.length, 1);
      chart.updatePoint(0, 0, 1, {});
      const points = chart.getProperties().plots[0].series[0].points;
      assert.equal(points.length, 1);
      assert.equal(points[0].index, 1);
      assert.equal(points[0].fill.color.argb >>> 0, 0xffff0000);
    });
  });

  it('is removed by setPoint with null, and removing an absent override does nothing', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      chart.updatePoint(0, 0, 1, fillRed);
      chart.setPoint(0, 0, 1, null);
      assert.equal(chart.getProperties().plots[0].series[0].points.length, 0);
      chart.setPoint(0, 0, 2, null);
      assert.equal(chart.getProperties().plots[0].series[0].points.length, 0);
    });
  });

  it('accepts the point record a read returned', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      chart.updatePoint(0, 0, 1, fillRed);
      const read = chart.getProperties().plots[0].series[0].points[0];
      chart.setPoint(0, 0, 3, read);
      const points = chart.getProperties().plots[0].series[0].points;
      assert.deepEqual(
        points.map((p) => p.index),
        [1, 3],
      );
      assert.equal(points[1].fill.color.argb >>> 0, 0xffff0000);
    });
  });
});

describe('a chart type is one vocabulary: the plot type plus a bar direction', () => {
  it('the retired insert-kind table is gone', () => {
    assert.equal(f.ChartInsertKind, undefined);
  });

  it('addChart reads back a bar chart drawn horizontally', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws, { type: f.ChartPlotType.Bar, barDirection: f.ChartBarDirection.Bar });
      const plot = chart.getProperties().plots[0];
      assert.equal(plot.type, f.ChartPlotType.Bar);
      assert.equal(plot.barDirection, f.ChartBarDirection.Bar);
    });
  });

  it('a bar chart without a direction is drawn as columns', () => {
    withSheet((ws) => {
      assert.equal(addColumnChart(ws).getProperties().plots[0].barDirection, f.ChartBarDirection.Col);
    });
  });

  for (const type of ['Bar', 'Line', 'Area', 'Pie', 'Doughnut', 'Scatter']) {
    it(`setType round-trips the ${type} plot it reads back`, () => {
      withSheet((ws) => {
        const chart = addColumnChart(ws, { type: f.ChartPlotType[type] });
        const plot = chart.getProperties().plots[0];
        assert.equal(plot.type, f.ChartPlotType[type]);
        chart.setType(0, plot.type, {
          ...(plot.barDirection === null ? {} : { barDirection: plot.barDirection }),
          ...(plot.grouping === null ? {} : { grouping: plot.grouping }),
        });
        const after = chart.getProperties().plots[0];
        assert.equal(after.type, plot.type);
        assert.equal(after.barDirection, plot.barDirection);
        assert.equal(after.grouping, plot.grouping);
      });
    });
  }

  it('setType switches a column plot to a horizontal stacked one', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      chart.setType(0, f.ChartPlotType.Bar, {
        barDirection: f.ChartBarDirection.Bar,
        grouping: f.ChartGrouping.Stacked,
      });
      const plot = chart.getProperties().plots[0];
      assert.equal(plot.barDirection, f.ChartBarDirection.Bar);
      assert.equal(plot.grouping, f.ChartGrouping.Stacked);
    });
  });

  it('a 3-D type is InvalidArgument and a bar direction on a non-bar type is a TypeError', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      expectInvalid(() => chart.setType(0, f.ChartPlotType.Bar3D));
      expectTypeError(() => chart.setType(0, f.ChartPlotType.Line, { barDirection: f.ChartBarDirection.Bar }));
      expectTypeError(() => chart.setType(0, f.ChartPlotType.Line, { barDirection: f.ChartBarDirection.Col }));
    });
  });
});

describe('a chart source and filter are records', () => {
  it('the source reads as a SheetRect and passes back to setSource', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      const source = chart.getProperties().source;
      assert.deepEqual(plain(source.range), { sheet: 0, rect: { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 3 } });
      chart.setSource(source);
      assert.deepEqual(plain(chart.getProperties().source), plain(source));
    });
  });

  it('a SheetRect range and the qualified A1 text insert the same chart', () => {
    withSheet((ws) => {
      fillGrid(ws);
      const rectRange = { sheet: 0, rect: { firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 1 } };
      const a = ws.addChart({ type: f.ChartPlotType.Line, range: rectRange });
      const b = ws.addChart({ type: f.ChartPlotType.Line, range: 'Sheet1!A1:B5' });
      assert.deepEqual(plain(a.getProperties().source.range), plain(b.getProperties().source.range));
    });
  });

  it('setSource takes range and refuses the old source key', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      chart.setSource({ range: 'Sheet1!A1:B5', direction: f.ChartSeriesDirection.Columns });
      assert.equal(chart.getProperties().source.range.rect.lastCol, 1);
      expectTypeError(() => chart.setSource({ source: 'Sheet1!A1:B5' }));
    });
  });

  it('legend entries and the series filter are independent records', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      assert.deepEqual(plain(chart.getProperties().filter), { hiddenSeries: [], hiddenCategories: [] });
      chart.updateLegend({ hiddenEntries: [0] });
      chart.setFilter({ hiddenSeries: [1], hiddenCategories: [2] });
      const props = chart.getProperties();
      assert.deepEqual(plain(props.legend.hiddenEntries), [0]);
      assert.deepEqual(plain(props.filter), { hiddenSeries: [1], hiddenCategories: [2] });
      assert.equal(props.legend.hiddenSeries, undefined);
      assert.equal(props.hiddenCategories, undefined);
    });
  });

  it('the filter a read returned passes back to setFilter', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      chart.setFilter({ hiddenSeries: [2], hiddenCategories: [0, 3] });
      const read = chart.getProperties().filter;
      chart.setFilter({ hiddenSeries: [], hiddenCategories: [] });
      chart.setFilter(read);
      assert.deepEqual(plain(chart.getProperties().filter), plain(read));
    });
  });

  it('a legend patch refuses the old hiddenSeries key', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      expectTypeError(() => chart.updateLegend({ hiddenSeries: [0] }));
    });
  });

  it('the legend a read returned passes back to setLegend', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      chart.updateLegend({ hiddenEntries: [1] });
      const legend = chart.getProperties().legend;
      chart.setLegend(null);
      chart.setLegend(legend);
      assert.deepEqual(plain(chart.getProperties().legend), plain(legend));
    });
  });
});

describe('every writable field of a chart axis and of a series label has a patch key', () => {
  const valueAxis = (chart) => chart.getProperties().axes.find((a) => a.kind === f.ChartAxisKind.Value);

  it('feeding an axis its own read fields back changes nothing', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      const before = plain(chart.getProperties());
      for (const axis of chart.getProperties().axes) {
        chart.updateAxis(
          axis.id,
          pick(axis, ['minorUnit', 'minorGridlines', 'position', 'sourceLinked', 'majorGridlines', 'visible']),
        );
      }
      assert.deepEqual(plain(chart.getProperties()), before);
    });
  });

  it('minorUnit, minorGridlines, position and sourceLinked set and read back', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      const id = valueAxis(chart).id;
      chart.updateAxis(id, {
        minorUnit: 0.5,
        minorGridlines: true,
        position: f.ChartAxisPosition.Right,
        sourceLinked: false,
      });
      const axis = valueAxis(chart);
      assert.equal(axis.minorUnit, 0.5);
      assert.equal(axis.minorGridlines, true);
      assert.equal(axis.position, f.ChartAxisPosition.Right);
      assert.equal(axis.sourceLinked, false);
    });
  });

  it('null clears minorUnit but is a TypeError on the flags and the position', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      const id = valueAxis(chart).id;
      chart.updateAxis(id, { minorUnit: 2 });
      chart.updateAxis(id, { minorUnit: null });
      assert.equal(valueAxis(chart).minorUnit, null);
      for (const key of ['minorGridlines', 'position', 'sourceLinked']) {
        expectTypeError(() => chart.updateAxis(id, { [key]: null }));
      }
    });
  });

  it('series labels take sourceLinked and deleted, and null on them is a TypeError', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      chart.updateSeries(0, 0, { labels: { showValue: true, sourceLinked: false, deleted: false } });
      let labels = chart.getProperties().plots[0].series[0].labels;
      assert.equal(labels.sourceLinked, false);
      assert.equal(labels.deleted, false);
      chart.updateSeries(0, 0, { labels: { deleted: true } });
      labels = chart.getProperties().plots[0].series[0].labels;
      assert.equal(labels.deleted, true);
      expectTypeError(() => chart.updateSeries(0, 0, { labels: { deleted: null } }));
      expectTypeError(() => chart.updateSeries(0, 0, { labels: { sourceLinked: null } }));
    });
  });

  it('the display patch sets the chart style and null removes it', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      chart.updateDisplay({ style: 26 });
      assert.equal(chart.getProperties().style, 26);
      chart.updateDisplay({ style: null });
      assert.equal(chart.getProperties().style, null);
      expectInvalid(() => chart.updateDisplay({ style: 0 }));
      expectInvalid(() => chart.updateDisplay({ style: 49 }));
    });
  });
});

describe('a series patch accepts the data source a read returned, literals included', () => {
  it('a reference series round-trips', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      const series = chart.getProperties().plots[0].series[0];
      chart.updateSeries(0, 0, { values: series.values, name: series.name, categories: series.categories });
      const after = chart.getProperties().plots[0].series[0];
      assert.deepEqual(plain(after.values), plain(series.values));
      assert.deepEqual(plain(after.name), plain(series.name));
    });
  });

  it('literal values and a literal name read back as a Literal source', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      chart.updateSeries(0, 0, {
        values: { kind: f.ChartSourceKind.Literal, formula: '', literal: [1, 2, 3, 4] },
        name: { kind: f.ChartSourceKind.Literal, formula: '', literal: ['Fixed'] },
      });
      const series = chart.getProperties().plots[0].series[0];
      assert.equal(series.values.kind, f.ChartSourceKind.Literal);
      assert.deepEqual(plain(series.values.literal), [1, 2, 3, 4]);
      assert.equal(series.name.kind, f.ChartSourceKind.Literal);
      assert.deepEqual(plain(series.name.literal), ['Fixed']);
      chart.updateSeries(0, 0, { values: series.values });
      assert.deepEqual(plain(chart.getProperties().plots[0].series[0].values.literal), [1, 2, 3, 4]);
    });
  });

  it('a literal source with no points is InvalidArgument', () => {
    withSheet((ws) => {
      const chart = addColumnChart(ws);
      expectInvalid(() =>
        chart.updateSeries(0, 0, { values: { kind: f.ChartSourceKind.Literal, formula: '', literal: [] } }),
      );
    });
  });
});

/** Sheet1 A1:B3 Region / Amount with a cache that shares the Region items, and a pivot: Region on rows, a Sum. */
function pivotOver(ws, wb) {
  ws.setText('A1', 'Region');
  ws.setText('B1', 'Amount');
  ws.setText('A2', 'East');
  ws.setNumber('B2', 10);
  ws.setText('A3', 'West');
  ws.setNumber('B3', 30);
  const cache = wb.addPivotCache();
  cache.addField('Region');
  cache.addField('Amount');
  cache.addFieldSharedItem(0, 'East');
  cache.addFieldSharedItem(0, 'West');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const record = cache.addRecord();
    cache.setRecordValue(record, 0, region);
    cache.setRecordValue(record, 1, amount);
  }
  const pivot = ws.addPivot({ name: 'P', cacheId: cache.cacheId, anchor: 'E1' });
  const region = pivot.addField({ sourceName: 'Region', axis: f.PivotAxis.Row });
  pivot.addFieldItem(region, 'East');
  pivot.addFieldItem(region, 'West');
  const amount = pivot.addField({ sourceName: 'Amount', axis: f.PivotAxis.Value });
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: f.PivotAggregation.Sum });
  return { cache, pivot };
}

describe('a pivot field or item patch is validated whole and applied atomically', () => {
  it('a field patch with a rejected key changes none of its keys', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      assert.equal(pivot.listFields()[0].insertBlankRow, false);
      expectInvalid(() => pivot.updateField(0, { insertBlankRow: true, customName: 'Amount' }));
      assert.equal(pivot.listFields()[0].insertBlankRow, false);
      assert.equal(pivot.listFields()[0].customName, null);
    });
  });

  it('a field patch whose sort keys disagree changes none of its keys', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      expectInvalid(() =>
        pivot.updateField(0, {
          insertBlankRow: true,
          showDetail: false,
          sortOrder: f.PivotSortOrder.Manual,
          sortByField: 'Sum of Amount',
        }),
      );
      const field = pivot.listFields()[0];
      assert.equal(field.insertBlankRow, false);
      assert.equal(field.showDetail, true);
    });
  });

  it('a field patch with a wrongly typed key changes nothing', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      expectTypeError(() => pivot.updateField(0, { showAll: true, customName: 5 }));
      assert.equal(pivot.listFields()[0].showAll, false);
    });
  });

  it('an item patch with a wrongly typed key changes nothing', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      assert.equal(pivot.getFieldItem(0, 0).visible, true);
      expectTypeError(() => pivot.updateFieldItem(0, 0, { visible: false, caption: 5 }));
      assert.equal(pivot.getFieldItem(0, 0).visible, true);
      assert.equal(pivot.getFieldItem(0, 0).caption, null);
    });
  });

  it('a valid field patch applies every key', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      pivot.updateField(0, { insertBlankRow: true, customName: 'Area', showAll: true });
      const field = pivot.listFields()[0];
      assert.deepEqual([field.insertBlankRow, field.customName, field.showAll], [true, 'Area', true]);
    });
  });

  it('an empty field patch changes nothing', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      const before = plain(pivot.listFields());
      pivot.updateField(0, {});
      assert.deepEqual(plain(pivot.listFields()), before);
    });
  });
});

describe('sort order, sort field and show-detail are field patch keys', () => {
  const KEYS = [
    'customName',
    'numberFormat',
    'defaultSubtotal',
    'insertBlankRow',
    'fillDownLabels',
    'showAll',
    'subtotalTop',
    'sortOrder',
    'sortByField',
    'showDetail',
  ];

  it('the writable fields of a field read pass back unchanged', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      const before = plain(pivot.listFields());
      for (const key of KEYS) assert.ok(key in before[0], `a field read carries ${key}`);
      for (let i = 0; i < before.length; i += 1) pivot.updateField(i, pick(pivot.listFields()[i], KEYS));
      assert.deepEqual(plain(pivot.listFields()), before);
    });
  });

  it('showDetail false collapses every item and reads false', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      assert.equal(pivot.listFields()[0].showDetail, true);
      pivot.updateField(0, { showDetail: false });
      assert.equal(pivot.listFields()[0].showDetail, false);
      assert.ok(pivot.listFieldItems(0).every((item) => item.showDetail === false));
      pivot.updateField(0, { showDetail: true });
      assert.equal(pivot.listFields()[0].showDetail, true);
    });
  });

  it('sortOrder with a sort field applies, and null on sortByField returns to label order', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      pivot.updateField(0, { sortOrder: f.PivotSortOrder.Descending, sortByField: 'Sum of Amount' });
      let field = pivot.listFields()[0];
      assert.equal(field.sortOrder, f.PivotSortOrder.Descending);
      assert.equal(field.sortByField, 'Sum of Amount');
      pivot.updateField(0, { sortByField: null });
      field = pivot.listFields()[0];
      assert.equal(field.sortByField, null);
    });
  });

  it('a manual order with a sort field is InvalidArgument', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      expectInvalid(() => pivot.updateField(0, { sortOrder: f.PivotSortOrder.Manual, sortByField: 'Sum of Amount' }));
    });
  });

  it('the per-property writers are gone', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      assert.equal(pivot.setFieldSort, undefined);
      assert.equal(pivot.setFieldShowDetail, undefined);
    });
  });

  it('a flag patch refuses null', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      expectTypeError(() => pivot.updateField(0, { showDetail: null }));
      expectTypeError(() => pivot.updateField(0, { sortOrder: null }));
    });
  });
});

describe('pivot style options update as a merge', () => {
  it('a single flag changes that flag and keeps the style name and the others', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      const before = pivot.getStyleOptions();
      pivot.updateStyleOptions({ showRowStripes: !before.showRowStripes });
      const after = pivot.getStyleOptions();
      assert.equal(after.showRowStripes, !before.showRowStripes);
      assert.deepEqual(plain({ ...after, showRowStripes: 0 }), plain({ ...before, showRowStripes: 0 }));
    });
  });

  it('an empty patch changes nothing and null clears only the style name', () => {
    withSheet((ws, wb) => {
      const { pivot } = pivotOver(ws, wb);
      const before = plain(pivot.getStyleOptions());
      pivot.updateStyleOptions({});
      assert.deepEqual(plain(pivot.getStyleOptions()), before);
      pivot.updateStyleOptions({ styleName: null });
      assert.equal(pivot.getStyleOptions().styleName, null);
      assert.equal(pivot.getStyleOptions().showRowHeaders, before.showRowHeaders);
      expectTypeError(() => pivot.updateStyleOptions({ showRowStripes: null }));
    });
  });
});
