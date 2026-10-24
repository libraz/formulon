import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

function must(r, what) {
  const status = typeof r.ok === 'boolean' ? r : r.status;
  assert.ok(status.ok, `${what}: ${JSON.stringify(status)}`);
  return r;
}

// Sheet1: B1:D1 names, A2:A5 text categories, B2:D5 numbers, F1:G5 scatter pairs.
function fillGrid(wb) {
  ['S1', 'S2', 'S3'].forEach((n, c) => {
    wb.setText(0, 0, c + 1, n);
  });
  ['a', 'b', 'c', 'd'].forEach((cat, r) => {
    wb.setText(0, r + 1, 0, cat);
    for (let c = 0; c < 3; c += 1) wb.setNumber(0, r + 1, c + 1, (r + 1) * (c + 2));
    wb.setNumber(0, r + 1, 5, r * 3);
    wb.setNumber(0, r + 1, 6, r * r);
  });
  wb.setText(0, 0, 5, 'X');
  wb.setText(0, 0, 6, 'Y');
}

function newBook(mod) {
  const wb = mod.Workbook.createDefault();
  fillGrid(wb);
  return wb;
}

function insert(wb, mod, extra = {}) {
  const r = must(
    wb.insertChart(0, {
      kind: mod.ChartInsertKind.Column,
      source: 'Sheet1!$A$1:$D$5',
      legend: mod.ChartLegendPosition.Right,
      ...extra,
    }),
    'insertChart',
  );
  assert.ok(r.objectId > 0);
  return r.objectId;
}

const chartOf = (wb, id) => must(wb.getChart(0, id), 'getChart').chart;
const dataOf = (wb, id) => must(wb.getChartData(0, id), 'getChartData').data;

test('chart enums keep their C ordinals', async () => {
  const mod = await getModule();
  assert.equal(mod.ChartPlotType.Bar, 0);
  assert.equal(mod.ChartPlotType.Chartex, 16);
  assert.equal(mod.ChartInsertKind.Scatter, 5);
  assert.equal(mod.ChartInsertKind.Doughnut, 6);
  assert.equal(mod.ChartGrouping.Standard, 3);
  assert.equal(mod.ChartBlanksAs.Span, 3);
  assert.equal(mod.ChartPointState.Connect, 2);
  assert.equal(mod.ChartAxisKind.Series, 3);
  assert.equal(mod.ChartAxisCrosses.At, 4);
  assert.equal(mod.ChartTitleKind.Auto, 3);
  assert.equal(mod.ChartSeriesDirection.Rows, 2);
  assert.equal(mod.ChartLegendPosition.TopRight, 5);
  assert.equal(mod.ChartLineDash.Other, 4);
  assert.equal(mod.ChartMarkerSymbol.Picture, 11);
  assert.equal(mod.ChartLabelPosition.Right, 8);
});

test('insertChart creates a two-cell chart that getChart and getChartData read back', async () => {
  const mod = await getModule();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod, { title: 'Sales', name: 'Revenue', description: 'alt', placement: { row: 7, col: 2 } });
    const listed = wb.listDrawingObjects(0);
    assert.equal(listed.length, 1);
    assert.equal(listed[0].objectId, id);
    assert.equal(listed[0].kind, mod.DrawingObjectKind.Chart);
    assert.equal(listed[0].anchorKind, mod.AnchorKind.TwoCell);
    assert.equal(listed[0].fromRow, 7);
    assert.equal(listed[0].fromCol, 2);

    const chart = chartOf(wb, id);
    assert.equal(chart.objectId, id);
    assert.equal(chart.name, 'Revenue');
    assert.equal(chart.description, 'alt');
    assert.equal(chart.isChartex, false);
    assert.equal(chart.isPivotChart, false);
    assert.equal(chart.title.kind, mod.ChartTitleKind.Text);
    assert.equal(chart.title.text, 'Sales');
    assert.equal(chart.legend.position, mod.ChartLegendPosition.Right);
    assert.deepEqual(chart.legend.hiddenSeries, []);
    assert.equal(chart.plots.length, 1);
    const plot = chart.plots[0];
    assert.equal(plot.type, mod.ChartPlotType.Bar);
    assert.equal(plot.barDirection, 'col');
    assert.equal(plot.axisIds.length, 2);
    assert.equal(plot.series.length, 3);
    const s0 = plot.series[0];
    assert.equal(s0.name.kind, 'ref');
    assert.equal(s0.name.formula, 'Sheet1!$B$1');
    assert.equal(s0.categories.formula, 'Sheet1!$A$2:$A$5');
    assert.equal(s0.values.formula, 'Sheet1!$B$2:$B$5');
    assert.equal(s0.x.kind, 'none');
    assert.equal(typeof s0.fill.color, 'number');
    assert.equal(s0.marker, null);
    assert.equal(s0.labels, null);
    assert.deepEqual(s0.points, []);
    assert.equal(chart.axes.length, 2);
    assert.deepEqual(chart.axes.map((a) => a.id).sort(), [...plot.axisIds].sort());
    assert.ok(['b', 't', 'l', 'r'].includes(chart.axes[0].position));
    assert.equal(chart.source.range, 'Sheet1!$A$1:$D$5');
    assert.equal(chart.source.direction, mod.ChartSeriesDirection.Columns);
    assert.equal(chart.source.headerRow, true);
    assert.equal(chart.source.headerColumn, true);

    const data = dataOf(wb, id);
    assert.equal(data.domain, 'category');
    assert.equal(data.pointCount, 4);
    assert.deepEqual(data.categories.labels, ['a', 'b', 'c', 'd']);
    assert.equal(data.categories.values, null);
    assert.equal(data.categories.generated, false);
    assert.equal(data.series.length, 3);
    assert.equal(data.series[0].name, 'S1');
    assert.ok(data.series[0].values instanceof Float64Array);
    assert.deepEqual([...data.series[0].values], [2, 4, 6, 8]);
    assert.ok(data.series[0].states instanceof Uint8Array);
    assert.deepEqual([...data.series[0].states], [0, 0, 0, 0]);
    assert.equal(data.series[0].x, null);
    assert.equal(data.series[0].sizes, null);
    assert.equal(data.series[0].ranges.values.length, 1);
    assert.deepEqual(data.series[0].ranges.values[0], { sheet: 0, top: 1, left: 1, bottom: 4, right: 1 });
    assert.deepEqual(data.unresolved, []);

    // Resolved data follows the cells.
    must(wb.setNumber(0, 1, 1, 100), 'setNumber');
    assert.equal(dataOf(wb, id).series[0].values[0], 100);
  } finally {
    wb.dispose();
  }
});

test('insertChart covers the creation kinds and omitted options', async () => {
  const mod = await getModule();
  const wb = newBook(mod);
  try {
    const kinds = [
      [mod.ChartInsertKind.Line, mod.ChartPlotType.Line],
      [mod.ChartInsertKind.Area, mod.ChartPlotType.Area],
      [mod.ChartInsertKind.Pie, mod.ChartPlotType.Pie],
      [mod.ChartInsertKind.Doughnut, mod.ChartPlotType.Doughnut],
    ];
    for (const [kind, plotType] of kinds) {
      const id = insert(wb, mod, { kind, source: 'Sheet1!$A$1:$B$5' });
      assert.equal(chartOf(wb, id).plots[0].type, plotType);
    }
    const bar = insert(wb, mod, {
      kind: mod.ChartInsertKind.Bar,
      grouping: mod.ChartGrouping.Stacked,
      source: 'A1:D5',
      direction: mod.ChartSeriesDirection.Columns,
      headerRow: true,
      headerColumn: true,
      title: '',
    });
    const barPlot = chartOf(wb, bar).plots[0];
    assert.equal(barPlot.barDirection, 'bar');
    assert.equal(barPlot.grouping, mod.ChartGrouping.Stacked);
    assert.equal(chartOf(wb, bar).title.kind, mod.ChartTitleKind.None);

    const xXf = wb.addXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 2 });
    assert.ok(xXf.status.ok, JSON.stringify(xXf.status));
    assert.ok(wb.setRangeXfIndex(0, 1, 5, 4, 5, xXf.index).ok);
    const scatter = insert(wb, mod, { kind: mod.ChartInsertKind.Scatter, source: 'Sheet1!$F$1:$G$5' });
    const sdata = dataOf(wb, scatter);
    assert.equal(sdata.domain, 'xy');
    assert.ok(sdata.series[0].x instanceof Float64Array);
    assert.deepEqual([...sdata.series[0].x], [0, 3, 6, 9]);
    assert.equal(sdata.series[0].xFormatCode, '0.00');
    assert.equal(sdata.series[0].formatCode, 'General');

    // An omitted title is Excel's automatic one.
    const auto = insert(wb, mod);
    assert.equal(chartOf(wb, auto).title.kind, mod.ChartTitleKind.Auto);
    assert.equal(wb.listDrawingObjects(0).at(-1).anchorKind, mod.AnchorKind.TwoCell);
  } finally {
    wb.dispose();
  }
});

test('typed chart edits are visible through getChart', async () => {
  const mod = await getModule();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod);

    must(wb.setChartTitle(0, id, 'Quarterly'), 'setChartTitle');
    assert.equal(chartOf(wb, id).title.text, 'Quarterly');
    must(wb.setChartTitle(0, id, null), 'setChartTitle null');
    assert.equal(chartOf(wb, id).title.kind, mod.ChartTitleKind.None);
    must(wb.setChartTitle(0, id, { kind: mod.ChartTitleKind.Auto }), 'setChartTitle auto');
    assert.equal(chartOf(wb, id).title.kind, mod.ChartTitleKind.Auto);

    must(wb.setChartLegend(0, id, mod.ChartLegendPosition.Bottom, true), 'setChartLegend');
    let legend = chartOf(wb, id).legend;
    assert.equal(legend.position, mod.ChartLegendPosition.Bottom);
    assert.equal(legend.overlay, true);
    must(wb.setChartLegend(0, id, mod.ChartLegendPosition.Bottom, false, [1]), 'setChartLegend hidden');
    legend = chartOf(wb, id).legend;
    assert.equal(legend.overlay, false);
    assert.deepEqual(legend.hiddenSeries, [1]);
    must(wb.setChartLegend(0, id, mod.ChartLegendPosition.None, false), 'setChartLegend none');
    assert.equal(chartOf(wb, id).legend.position, mod.ChartLegendPosition.None);

    must(wb.setChartDisplay(0, id, mod.ChartBlanksAs.Zero, false), 'setChartDisplay');
    let info = chartOf(wb, id);
    assert.equal(info.blanksAs, mod.ChartBlanksAs.Zero);
    assert.equal(info.plotVisibleOnly, false);
    must(wb.setChartDisplay(0, id, mod.ChartBlanksAs.Absent, null), 'setChartDisplay absent');
    info = chartOf(wb, id);
    assert.equal(info.blanksAs, mod.ChartBlanksAs.Absent);
    assert.equal(info.plotVisibleOnly, null);

    must(wb.setChartPlot(0, id, 0, { gapWidth: 40, overlap: 10, varyColors: true }), 'setChartPlot');
    const plot = chartOf(wb, id).plots[0];
    assert.equal(plot.gapWidth, 40);
    assert.equal(plot.overlap, 10);
    assert.equal(plot.varyColors, true);

    const valueAxis = chartOf(wb, id).axes.find((a) => a.kind === mod.ChartAxisKind.Value);
    must(
      wb.setChartAxis(0, id, valueAxis.id, {
        min: 0,
        max: 20,
        majorUnit: 5,
        title: 'Units',
        numberFormat: '0.0',
        majorGridlines: false,
        reversed: true,
      }),
      'setChartAxis',
    );
    const axis = chartOf(wb, id).axes.find((a) => a.id === valueAxis.id);
    assert.equal(axis.min, 0);
    assert.equal(axis.max, 20);
    assert.equal(axis.majorUnit, 5);
    assert.equal(axis.title.kind, mod.ChartTitleKind.Text);
    assert.equal(axis.title.text, 'Units');
    assert.equal(axis.numberFormat, '0.0');
    assert.equal(axis.majorGridlines, false);
    assert.equal(axis.reversed, true);
    // The axis title takes the setChartTitle forms; Auto leaves the text to the viewer.
    must(wb.setChartAxis(0, id, valueAxis.id, { title: { kind: mod.ChartTitleKind.Auto } }), 'setChartAxis auto');
    assert.equal(chartOf(wb, id).axes.find((a) => a.id === valueAxis.id).title.kind, mod.ChartTitleKind.Auto);
    must(
      wb.setChartAxis(0, id, valueAxis.id, { title: { kind: mod.ChartTitleKind.Ref, formula: 'Sheet1!$A$1' } }),
      'setChartAxis ref',
    );
    assert.equal(chartOf(wb, id).axes.find((a) => a.id === valueAxis.id).title.formula, 'Sheet1!$A$1');
    // null clears a bound.
    must(wb.setChartAxis(0, id, valueAxis.id, { max: null }), 'setChartAxis clear');
    assert.equal(chartOf(wb, id).axes.find((a) => a.id === valueAxis.id).max, null);

    // Markers and smoothing exist on line plots only.
    must(wb.setChartType(0, id, 0, mod.ChartInsertKind.Line), 'setChartType');
    assert.equal(chartOf(wb, id).plots[0].type, mod.ChartPlotType.Line);
    must(
      wb.setChartSeries(0, id, 0, 1, {
        fill: { color: 0xffff0000, transparency: 25 },
        line: { color: 0xff0000ff, widthPt: 2, dash: mod.ChartLineDash.Dash },
        marker: { symbol: mod.ChartMarkerSymbol.Circle, size: 7 },
        smooth: true,
        labels: { showValue: true, position: mod.ChartLabelPosition.OutsideEnd, numberFormat: '0.00' },
      }),
      'setChartSeries format',
    );
    let series = chartOf(wb, id).plots[0].series[1];
    // A line series takes no fill of its own; its colour is the line's.
    assert.equal(series.fill.color, null);
    assert.equal(series.fill.transparency, null);
    assert.equal(series.line.color, 0xff0000ff);
    assert.equal(series.line.widthPt, 2);
    assert.equal(series.line.dash, mod.ChartLineDash.Dash);
    assert.equal(series.marker.symbol, mod.ChartMarkerSymbol.Circle);
    assert.equal(series.marker.size, 7);
    assert.equal(series.smooth, true);
    assert.equal(series.labels.showValue, true);
    assert.equal(series.labels.position, mod.ChartLabelPosition.OutsideEnd);
    assert.equal(series.labels.numberFormat, '0.00');
    // null removes the marker and the labels; an omitted field is kept.
    must(wb.setChartSeries(0, id, 0, 1, { marker: null, labels: null }), 'setChartSeries clear');
    series = chartOf(wb, id).plots[0].series[1];
    assert.equal(series.marker, null);
    assert.equal(series.labels, null);
    assert.equal(series.line.color, 0xff0000ff);

    // Reference edits refill the cache and resolved data follows.
    must(wb.setChartSeries(0, id, 0, 0, { values: 'Sheet1!$D$2:$D$5', name: '' }), 'setChartSeries refs');
    series = chartOf(wb, id).plots[0].series[0];
    assert.equal(series.values.formula, 'Sheet1!$D$2:$D$5');
    assert.equal(series.name.kind, 'none');
    assert.deepEqual([...dataOf(wb, id).series[0].values], [4, 8, 12, 16]);

    // A line point takes no fill; its line carries the colour.
    const otherPoints = chartOf(wb, id).plots[0].series[1].points.filter((point) => point.idx !== 2);
    must(wb.setChartPoint(0, id, 0, 1, 2, { line: { color: 0xff00ff00 } }), 'setChartPoint');
    const points = chartOf(wb, id).plots[0].series[1].points;
    assert.equal(points.length, otherPoints.length + 1);
    assert.equal(points.find((point) => point.idx === 2).line.color, 0xff00ff00);
    assert.deepEqual(
      points.filter((point) => point.idx !== 2),
      otherPoints,
    );
    must(wb.setChartPoint(0, id, 0, 1, 2, null), 'setChartPoint remove');
    assert.deepEqual(chartOf(wb, id).plots[0].series[1].points, otherPoints);

    must(wb.setChartFilter(0, id, [2], [1]), 'setChartFilter');
    info = chartOf(wb, id);
    assert.equal(info.plots[0].series.find((s) => s.idx === 2).filtered, true);
    assert.deepEqual(info.hiddenCategories, [1]);
    const filtered = dataOf(wb, id);
    assert.ok(filtered.series.find((s) => s.idx === 2).filtered);
    must(wb.setChartFilter(0, id, [], []), 'setChartFilter clear');
    assert.deepEqual(chartOf(wb, id).hiddenCategories, []);

    must(wb.setChartType(0, id, 0, mod.ChartInsertKind.Column, mod.ChartGrouping.Stacked), 'setChartType column');
    assert.equal(chartOf(wb, id).plots[0].type, mod.ChartPlotType.Bar);
    assert.equal(chartOf(wb, id).plots[0].grouping, mod.ChartGrouping.Stacked);

    const pie = insert(wb, mod, { kind: mod.ChartInsertKind.Doughnut, source: 'Sheet1!$A$1:$B$5' });
    must(wb.setChartPlot(0, pie, 0, { holeSize: 60, firstSliceAngle: 90 }), 'setChartPlot hole');
    assert.equal(chartOf(wb, pie).plots[0].holeSize, 60);
    assert.equal(chartOf(wb, pie).plots[0].firstSliceAngle, 90);
  } finally {
    wb.dispose();
  }
});

test('series can be added, removed and rebuilt from a new source', async () => {
  const mod = await getModule();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod, { source: 'Sheet1!$A$1:$C$5' });
    assert.equal(chartOf(wb, id).plots[0].series.length, 2);

    const added = must(
      wb.addChartSeries(0, id, 0, { values: 'Sheet1!$D$2:$D$5', name: 'Sheet1!$D$1', categories: 'Sheet1!$A$2:$A$5' }),
      'addChartSeries',
    );
    assert.equal(chartOf(wb, id).plots[0].series.length, 3);
    assert.ok(chartOf(wb, id).plots[0].series.some((s) => s.idx === added.seriesIdx));
    assert.equal(dataOf(wb, id).series.length, 3);

    must(wb.removeChartSeries(0, id, 0, added.seriesIdx), 'removeChartSeries');
    assert.equal(chartOf(wb, id).plots[0].series.length, 2);

    must(wb.setChartSeries(0, id, 0, 0, { fill: { color: 0xff112233 } }), 'setChartSeries');
    must(wb.setChartSource(0, id, 'Sheet1!$A$1:$D$5', mod.ChartSeriesDirection.Columns, true, true), 'setChartSource');
    let chart = chartOf(wb, id);
    assert.equal(chart.plots[0].series.length, 3);
    assert.equal(chart.plots[0].series[0].fill.color, 0xff112233);
    assert.equal(chart.source.range, 'Sheet1!$A$1:$D$5');

    // Rows direction swaps the split: four series of three points.
    must(wb.setChartSource(0, id, 'Sheet1!$A$1:$D$5', mod.ChartSeriesDirection.Rows), 'setChartSource rows');
    chart = chartOf(wb, id);
    assert.equal(chart.plots[0].series.length, 4);
    assert.equal(chart.source.direction, mod.ChartSeriesDirection.Rows);
  } finally {
    wb.dispose();
  }
});

test('previewChartData resolves a draft without changing the chart', async () => {
  const mod = await getModule();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod);
    const before = JSON.stringify(chartOf(wb, id));
    const preview = must(
      wb.previewChartData(0, id, { source: 'Sheet1!$A$1:$B$5', blanksAs: mod.ChartBlanksAs.Gap }),
      'previewChartData',
    ).data;
    assert.equal(preview.series.length, 1);
    assert.deepEqual([...preview.series[0].values], [2, 4, 6, 8]);
    const rows = must(
      wb.previewChartData(0, id, { direction: mod.ChartSeriesDirection.Rows, headerRow: true, headerColumn: true }),
      'previewChartData rows',
    ).data;
    assert.equal(rows.series.length, 4);
    assert.equal(JSON.stringify(chartOf(wb, id)), before);
    // No overrides is the chart as it stands.
    assert.equal(must(wb.previewChartData(0, id, {}), 'previewChartData {}').data.series.length, 3);
    assert.equal(wb.previewChartData(0, id, { source: 'NoSuchSheet!A1:B2' }).status.ok, false);
    assert.equal(wb.previewChartData(0, 9999, {}).data, null);
  } finally {
    wb.dispose();
  }
});

test('a chart survives save and load and goes through the drawing operations', async () => {
  const mod = await getModule();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod, { title: 'Kept', placement: { row: 3, col: 3 } });
    const other = insert(wb, mod, { kind: mod.ChartInsertKind.Line, source: 'Sheet1!$A$1:$B$5' });

    const saved = must(wb.save(), 'save');
    const wb2 = mod.Workbook.loadBytes(saved.bytes ?? saved);
    try {
      assert.equal(chartOf(wb2, id).title.text, 'Kept');
      assert.deepEqual([...dataOf(wb2, id).series[0].values], [2, 4, 6, 8]);
      assert.equal(wb2.listDrawingObjects(0).filter((o) => o.kind === mod.DrawingObjectKind.Chart).length, 2);
    } finally {
      wb2.dispose();
    }

    // Move, reorder, snapshot, remove and restore keep the chart's id.
    must(wb.setDrawingAnchor(0, id, { row: 10, col: 1, widthEmu: 3000000, heightEmu: 2000000 }), 'setDrawingAnchor');
    const moved = wb.listDrawingObjects(0).find((o) => o.objectId === id);
    assert.equal(moved.fromRow, 10);
    assert.equal(moved.fromCol, 1);
    assert.equal(moved.cx, 3000000);
    assert.equal(moved.cy, 2000000);
    must(wb.setDrawingZOrder(0, id, 0), 'setDrawingZOrder');
    assert.deepEqual(
      Array.from(wb.listDrawingObjects(0), (o) => o.objectId),
      [id, other],
    );

    const snap = must(wb.snapshotDrawing(0, id), 'snapshotDrawing');
    assert.ok(snap.bytes instanceof Uint8Array);
    assert.ok(snap.bytes.length > 0);
    must(wb.removeDrawing(0, id), 'removeDrawing');
    assert.equal(wb.getChart(0, id).status.ok, false);
    assert.deepEqual(
      Array.from(wb.listDrawingObjects(0), (o) => o.objectId),
      [other],
    );

    const restored = must(wb.restoreDrawing(0, snap.bytes), 'restoreDrawing');
    assert.equal(restored.objectId, id);
    assert.equal(chartOf(wb, id).title.text, 'Kept');
    assert.deepEqual(
      Array.from(wb.listDrawingObjects(0), (o) => o.objectId),
      [id, other],
    );

    const copy = must(wb.restoreDrawing(0, snap.bytes, { newId: true }), 'restoreDrawing newId');
    assert.notEqual(copy.objectId, id);
    assert.equal(chartOf(wb, copy.objectId).title.text, 'Kept');
    assert.equal(wb.listDrawingObjects(0).length, 3);
  } finally {
    wb.dispose();
  }
});

test('chart calls report failures in their status and reject malformed arguments', async () => {
  const mod = await getModule();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod);

    const missing = wb.getChart(0, 9999);
    assert.equal(missing.status.ok, false);
    assert.equal(missing.chart, null);
    assert.equal(wb.getChart(99, id).status.ok, false);
    const noData = wb.getChartData(0, 9999);
    assert.equal(noData.status.ok, false);
    assert.equal(noData.data, null);

    // An unresolvable source writes nothing.
    const bad = wb.insertChart(0, { kind: mod.ChartInsertKind.Column, source: 'NoSuchSheet!A1:B2' });
    assert.equal(bad.status.ok, false);
    assert.equal(bad.status.status, 5026);
    assert.equal(bad.objectId, 0);
    assert.equal(wb.listDrawingObjects(0).length, 1);
    assert.equal(wb.insertChart(0, { kind: 99, source: 'A1:B2' }).status.ok, false);
    assert.throws(() => wb.insertChart(0, 'A1:B2'), TypeError);

    assert.equal(wb.setChartSeries(0, id, 0, 0, { values: '' }).ok, false);
    assert.equal(wb.setChartSeries(0, id, 0, 0, { values: 'NoSuchSheet!A1:A2' }).status, 5026);
    assert.equal(wb.setChartSeries(0, id, 5, 0, {}).ok, false);
    assert.equal(wb.setChartSeries(0, id, 0, 99, {}).ok, false);
    assert.equal(wb.setChartSeries(0, id, 0, 0, { fill: { color: 1 }, marker: { symbol: null } }).ok, false);
    assert.equal(wb.setChartAxis(0, id, 12345, {}).ok, false);
    assert.equal(wb.setChartPlot(0, id, 0, { holeSize: 50 }).ok, false);
    assert.equal(wb.setChartType(0, id, 0, 99).ok, false);
    assert.equal(wb.setChartSource(0, id, 'NoSuchSheet!A1:B2', mod.ChartSeriesDirection.Auto).status, 5026);
    assert.equal(wb.removeChartSeries(0, id, 0, 99).ok, false);
    assert.equal(wb.addChartSeries(0, id, 0, { name: 'Sheet1!$B$1' }).status.ok, false);
    assert.equal(wb.setChartTitle(0, 9999, 'x').ok, false);
    assert.equal(wb.removeDrawing(0, 9999).ok, false);
    assert.equal(wb.snapshotDrawing(0, 9999).status.ok, false);

    // Wrong types raise before the engine is called.
    assert.throws(() => wb.getChart('0', id), TypeError);
    assert.throws(() => wb.getChart(0, -1), RangeError);
    assert.throws(() => wb.setChartType(0, id, 0, 1.5), RangeError);
    assert.throws(() => wb.setChartSeries(0, id, 0, 0, { fill: { color: 'red' } }), TypeError);
    assert.throws(() => wb.setChartSeries(0, id, 0, 0, { values: 5 }), TypeError);
    assert.throws(() => wb.setChartSeries(0, id, 0, 0, 'nope'), TypeError);
    assert.throws(() => wb.setChartAxis(0, id, 1, { min: 'low' }), TypeError);
    assert.throws(() => wb.setChartFilter(0, id, [1, 'x'], []), TypeError);
    assert.throws(() => wb.setChartFilter(0, id, 'x', []), TypeError);
    assert.throws(() => wb.setChartTitle(0, id, 5), TypeError);
  } finally {
    wb.dispose();
  }
});

test('chart calls on a disposed workbook report an invalid handle', async () => {
  const mod = await getModule();
  const wb = newBook(mod);
  const id = insert(wb, mod);
  wb.dispose();
  assert.equal(wb.getChart(0, id).status.ok, false);
  assert.equal(wb.getChartData(0, id).status.ok, false);
  assert.equal(wb.insertChart(0, { kind: 0, source: 'A1:B2' }).status.ok, false);
  assert.equal(wb.setChartTitle(0, id, 'x').ok, false);
  assert.equal(wb.addChartSeries(0, id, 0, { values: 'A1:A2' }).status.ok, false);
});
