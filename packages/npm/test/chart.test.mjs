import test from 'node:test';
import { assert, expectFailure, getModule, loadStagedModule } from './smoke_support.mjs';

// Sheet1: B1:D1 names, A2:A5 text categories, B2:D5 numbers, F1:G5 scatter pairs.
function fillGrid(wb) {
  ['S1', 'S2', 'S3'].forEach((n, c) => {
    wb.sheet(0).setText(0, c + 1, n);
  });
  ['a', 'b', 'c', 'd'].forEach((cat, r) => {
    wb.sheet(0).setText(r + 1, 0, cat);
    for (let c = 0; c < 3; c += 1) wb.sheet(0).setNumber(r + 1, c + 1, (r + 1) * (c + 2));
    wb.sheet(0).setNumber(r + 1, 5, r * 3);
    wb.sheet(0).setNumber(r + 1, 6, r * r);
  });
  wb.sheet(0).setText(0, 5, 'X');
  wb.sheet(0).setText(0, 6, 'Y');
}

// The chart enums are named exports of the package; the workbook factory is the module.
async function load() {
  const [mod, enums] = await Promise.all([getModule(), loadStagedModule()]);
  return { Workbook: mod.Workbook, ...enums };
}

function newBook(mod) {
  const wb = mod.Workbook.createDefault();
  fillGrid(wb);
  return wb;
}

function insert(wb, mod, extra = {}) {
  const r = wb.sheet(0).addChart({
    kind: mod.ChartInsertKind.Column,
    source: 'Sheet1!$A$1:$D$5',
    legend: mod.ChartLegendPosition.Right,
    ...extra,
  });
  assert.ok(r > 0);
  return r;
}

const chartOf = (wb, id) => wb.sheet(0).getChart(id);
const dataOf = (wb, id) => wb.sheet(0).getChartData(id);

test('chart enums keep their C ordinals', async () => {
  const mod = await load();
  assert.equal(mod.ChartPlotType.Bar, 0);
  assert.equal(mod.ChartPlotType.Chartex, 16);
  assert.equal(mod.ChartInsertKind.Scatter, 5);
  assert.equal(mod.ChartInsertKind.Doughnut, 6);
  assert.equal(mod.ChartGrouping.Standard, 3);
  assert.equal(mod.ChartBlanksAs.Span, 2);
  assert.equal(mod.ChartPointState.Connect, 2);
  assert.equal(mod.ChartAxisKind.Series, 3);
  assert.equal(mod.ChartAxisCrosses.At, 3);
  assert.equal(mod.ChartTitleKind.Auto, 3);
  assert.equal(mod.ChartSeriesDirection.Rows, 2);
  assert.equal(mod.ChartLegendPosition.TopRight, 4);
  assert.equal(mod.ChartAxisPosition.Right, 3);
  assert.equal(mod.ChartBarDirection.Bar, 1);
  assert.equal(mod.ChartLineDash.Other, 4);
  assert.equal(mod.ChartMarkerSymbol.Picture, 11);
  assert.equal(mod.ChartLabelPosition.Right, 8);
});

test('insertChart creates a two-cell chart that getChart and getChartData read back', async () => {
  const mod = await load();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod, {
      title: 'Sales',
      placement: { name: 'Revenue', description: 'alt', anchor: { row: 7, col: 2 } },
    });
    const listed = wb.sheet(0).listDrawingObjects();
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
    assert.equal(plot.barDirection, mod.ChartBarDirection.Col);
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
    assert.ok(Object.values(mod.ChartAxisPosition).includes(chart.axes[0].position));
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
    assert.deepEqual(data.series[0].ranges.values[0], {
      sheet: 0,
      rect: { firstRow: 1, firstCol: 1, lastRow: 4, lastCol: 1 },
    });
    assert.deepEqual(data.unresolved, []);

    // Resolved data follows the cells.
    wb.sheet(0).setNumber(1, 1, 100);
    assert.equal(dataOf(wb, id).series[0].values[0], 100);
  } finally {
    wb.dispose();
  }
});

test('insertChart covers the creation kinds and omitted options', async () => {
  const mod = await load();
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
    assert.equal(barPlot.barDirection, mod.ChartBarDirection.Bar);
    assert.equal(barPlot.grouping, mod.ChartGrouping.Stacked);
    assert.equal(chartOf(wb, bar).title, null);

    const xXf = wb.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 2 });
    wb.sheet(0).setRangeXfIndex({ firstRow: 1, firstCol: 5, lastRow: 4, lastCol: 5 }, xXf);
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
    assert.equal(wb.sheet(0).listDrawingObjects().at(-1).anchorKind, mod.AnchorKind.TwoCell);
  } finally {
    wb.dispose();
  }
});

test('typed chart edits are visible through getChart', async () => {
  const mod = await load();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod);

    wb.sheet(0).setChartTitle(id, 'Quarterly');
    assert.equal(chartOf(wb, id).title.text, 'Quarterly');
    wb.sheet(0).setChartTitle(id, null);
    assert.equal(chartOf(wb, id).title, null);
    wb.sheet(0).setChartTitle(id, { kind: mod.ChartTitleKind.Auto });
    assert.equal(chartOf(wb, id).title.kind, mod.ChartTitleKind.Auto);

    wb.sheet(0).setChartLegend(id, mod.ChartLegendPosition.Bottom, true);
    let legend = chartOf(wb, id).legend;
    assert.equal(legend.position, mod.ChartLegendPosition.Bottom);
    assert.equal(legend.overlay, true);
    wb.sheet(0).setChartLegend(id, mod.ChartLegendPosition.Bottom, false, [1]);
    legend = chartOf(wb, id).legend;
    assert.equal(legend.overlay, false);
    assert.deepEqual(legend.hiddenSeries, [1]);
    wb.sheet(0).setChartLegend(id, null, false);
    assert.equal(chartOf(wb, id).legend, null);

    // A new chart carries `dispNaAsBlank`.
    assert.equal(chartOf(wb, id).naAsBlank, true);
    wb.sheet(0).setChartDisplay(id, { blanksAs: mod.ChartBlanksAs.Zero, plotVisibleOnly: false });
    let info = chartOf(wb, id);
    assert.equal(info.blanksAs, mod.ChartBlanksAs.Zero);
    assert.equal(info.plotVisibleOnly, false);
    // An omitted key keeps, including `naAsBlank`.
    assert.equal(info.naAsBlank, true);
    wb.sheet(0).setChartDisplay(id, { blanksAs: null, plotVisibleOnly: null });
    info = chartOf(wb, id);
    assert.equal(info.blanksAs, null);
    assert.equal(info.plotVisibleOnly, null);
    assert.equal(info.naAsBlank, true);
    wb.sheet(0).setChartDisplay(id, { plotVisibleOnly: true });
    assert.equal(chartOf(wb, id).blanksAs, null);
    assert.equal(chartOf(wb, id).plotVisibleOnly, true);
    // `false` and `null` both remove the extension, which reads back as `null`.
    wb.sheet(0).setChartDisplay(id, { naAsBlank: false });
    assert.equal(chartOf(wb, id).naAsBlank, null);
    wb.sheet(0).setChartDisplay(id, { naAsBlank: true });
    assert.equal(chartOf(wb, id).naAsBlank, true);
    wb.sheet(0).setChartDisplay(id, { naAsBlank: null });
    assert.equal(chartOf(wb, id).naAsBlank, null);
    assert.equal(chartOf(wb, id).plotVisibleOnly, true);
    expectFailure(() => wb.sheet(0).setChartDisplay(id, { blanksAs: -1 }));
    assert.throws(() => wb.sheet(0).setChartDisplay(id, 5), TypeError);
    wb.sheet(0).getChartPreview(id, { naAsBlank: true });
    wb.sheet(0).getChartPreview(id, { naAsBlank: false });
    assert.equal(chartOf(wb, id).naAsBlank, null);

    wb.sheet(0).setChartPlot(id, 0, { gapWidth: 40, overlap: 10, varyColors: true });
    const plot = chartOf(wb, id).plots[0];
    assert.equal(plot.gapWidth, 40);
    assert.equal(plot.overlap, 10);
    assert.equal(plot.varyColors, true);

    const valueAxis = chartOf(wb, id).axes.find((a) => a.kind === mod.ChartAxisKind.Value);
    wb.sheet(0).setChartAxis(id, valueAxis.id, {
      min: 0,
      max: 20,
      majorUnit: 5,
      title: 'Units',
      numberFormat: '0.0',
      majorGridlines: false,
      reversed: true,
    });
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
    wb.sheet(0).setChartAxis(id, valueAxis.id, { title: { kind: mod.ChartTitleKind.Auto } });
    assert.equal(chartOf(wb, id).axes.find((a) => a.id === valueAxis.id).title.kind, mod.ChartTitleKind.Auto);
    wb.sheet(0).setChartAxis(id, valueAxis.id, { title: { kind: mod.ChartTitleKind.Ref, formula: 'Sheet1!$A$1' } });
    assert.equal(chartOf(wb, id).axes.find((a) => a.id === valueAxis.id).title.formula, 'Sheet1!$A$1');
    // null clears a bound.
    wb.sheet(0).setChartAxis(id, valueAxis.id, { max: null });
    assert.equal(chartOf(wb, id).axes.find((a) => a.id === valueAxis.id).max, null);

    // Markers and smoothing exist on line plots only.
    wb.sheet(0).setChartType(id, 0, mod.ChartInsertKind.Line);
    assert.equal(chartOf(wb, id).plots[0].type, mod.ChartPlotType.Line);
    wb.sheet(0).setChartSeries(id, 0, 1, {
      fill: { color: 0xffff0000, transparency: 25 },
      line: { color: 0xff0000ff, widthPt: 2, dash: mod.ChartLineDash.Dash },
      marker: { symbol: mod.ChartMarkerSymbol.Circle, size: 7 },
      smooth: true,
      labels: { showValue: true, position: mod.ChartLabelPosition.OutsideEnd, numberFormat: '0.00' },
    });
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
    wb.sheet(0).setChartSeries(id, 0, 1, { marker: null, labels: null });
    series = chartOf(wb, id).plots[0].series[1];
    assert.equal(series.marker, null);
    assert.equal(series.labels, null);
    assert.equal(series.line.color, 0xff0000ff);

    // Reference edits refill the cache and resolved data follows.
    wb.sheet(0).setChartSeries(id, 0, 0, { values: 'Sheet1!$D$2:$D$5', name: '' });
    series = chartOf(wb, id).plots[0].series[0];
    assert.equal(series.values.formula, 'Sheet1!$D$2:$D$5');
    assert.equal(series.name.kind, 'none');
    assert.deepEqual([...dataOf(wb, id).series[0].values], [4, 8, 12, 16]);

    // A line point takes no fill; its line carries the colour.
    const otherPoints = chartOf(wb, id).plots[0].series[1].points.filter((point) => point.index !== 2);
    wb.sheet(0).setChartPoint(id, 0, 1, 2, { line: { color: 0xff00ff00 } });
    const points = chartOf(wb, id).plots[0].series[1].points;
    assert.equal(points.length, otherPoints.length + 1);
    assert.equal(points.find((point) => point.index === 2).line.color, 0xff00ff00);
    assert.deepEqual(
      points.filter((point) => point.index !== 2),
      otherPoints,
    );
    wb.sheet(0).setChartPoint(id, 0, 1, 2, null);
    assert.deepEqual(chartOf(wb, id).plots[0].series[1].points, otherPoints);

    wb.sheet(0).setChartFilter(id, [2], [1]);
    info = chartOf(wb, id);
    assert.equal(info.plots[0].series.find((s) => s.index === 2).filtered, true);
    assert.deepEqual(info.hiddenCategories, [1]);
    const filtered = dataOf(wb, id);
    assert.ok(filtered.series.find((s) => s.index === 2).filtered);
    wb.sheet(0).setChartFilter(id, [], []);
    assert.deepEqual(chartOf(wb, id).hiddenCategories, []);

    wb.sheet(0).setChartType(id, 0, mod.ChartInsertKind.Column, mod.ChartGrouping.Stacked);
    assert.equal(chartOf(wb, id).plots[0].type, mod.ChartPlotType.Bar);
    assert.equal(chartOf(wb, id).plots[0].grouping, mod.ChartGrouping.Stacked);

    const pie = insert(wb, mod, { kind: mod.ChartInsertKind.Doughnut, source: 'Sheet1!$A$1:$B$5' });
    wb.sheet(0).setChartPlot(pie, 0, { holeSize: 60, firstSliceAngle: 90 });
    assert.equal(chartOf(wb, pie).plots[0].holeSize, 60);
    assert.equal(chartOf(wb, pie).plots[0].firstSliceAngle, 90);
  } finally {
    wb.dispose();
  }
});

test('series can be added, removed and rebuilt from a new source', async () => {
  const mod = await load();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod, { source: 'Sheet1!$A$1:$C$5' });
    assert.equal(chartOf(wb, id).plots[0].series.length, 2);

    const added = wb.sheet(0).addChartSeries(id, 0, {
      values: 'Sheet1!$D$2:$D$5',
      name: 'Sheet1!$D$1',
      categories: 'Sheet1!$A$2:$A$5',
    });
    assert.equal(chartOf(wb, id).plots[0].series.length, 3);
    assert.ok(chartOf(wb, id).plots[0].series.some((s) => s.index === added));
    assert.equal(dataOf(wb, id).series.length, 3);

    wb.sheet(0).removeChartSeries(id, 0, added);
    assert.equal(chartOf(wb, id).plots[0].series.length, 2);

    wb.sheet(0).setChartSeries(id, 0, 0, { fill: { color: 0xff112233 } });
    wb.sheet(0).setChartSource(id, {
      source: 'Sheet1!$A$1:$D$5',
      direction: mod.ChartSeriesDirection.Columns,
      headerRow: true,
      headerColumn: true,
    });
    let chart = chartOf(wb, id);
    assert.equal(chart.plots[0].series.length, 3);
    assert.equal(chart.plots[0].series[0].fill.color, 0xff112233);
    assert.equal(chart.source.range, 'Sheet1!$A$1:$D$5');

    // Rows direction swaps the split: four series of three points.
    wb.sheet(0).setChartSource(id, { source: 'Sheet1!$A$1:$D$5', direction: mod.ChartSeriesDirection.Rows });
    chart = chartOf(wb, id);
    assert.equal(chart.plots[0].series.length, 4);
    assert.equal(chart.source.direction, mod.ChartSeriesDirection.Rows);
  } finally {
    wb.dispose();
  }
});

test('previewChartData resolves a draft without changing the chart', async () => {
  const mod = await load();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod);
    const before = JSON.stringify(chartOf(wb, id));
    const preview = wb.sheet(0).getChartPreview(id, { source: 'Sheet1!$A$1:$B$5', blanksAs: mod.ChartBlanksAs.Gap });
    assert.equal(preview.series.length, 1);
    assert.deepEqual([...preview.series[0].values], [2, 4, 6, 8]);
    const rows = wb.sheet(0).getChartPreview(id, {
      direction: mod.ChartSeriesDirection.Rows,
      headerRow: true,
      headerColumn: true,
    });
    assert.equal(rows.series.length, 4);
    assert.equal(JSON.stringify(chartOf(wb, id)), before);
    // No overrides is the chart as it stands.
    assert.equal(wb.sheet(0).getChartPreview(id, {}).series.length, 3);
    expectFailure(() => wb.sheet(0).getChartPreview(id, { source: 'NoSuchSheet!A1:B2' }));
    expectFailure(() => wb.sheet(0).getChartPreview(9999, {}));
  } finally {
    wb.dispose();
  }
});

test('a chart survives save and load and goes through the drawing operations', async () => {
  const mod = await load();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod, { title: 'Kept', placement: { anchor: { row: 3, col: 3 } } });
    const other = insert(wb, mod, { kind: mod.ChartInsertKind.Line, source: 'Sheet1!$A$1:$B$5' });

    const saved = wb.save();
    const wb2 = mod.Workbook.load(saved);
    try {
      assert.equal(chartOf(wb2, id).title.text, 'Kept');
      assert.deepEqual([...dataOf(wb2, id).series[0].values], [2, 4, 6, 8]);
      assert.equal(
        wb2
          .sheet(0)
          .listDrawingObjects()
          .filter((o) => o.kind === mod.DrawingObjectKind.Chart).length,
        2,
      );
    } finally {
      wb2.dispose();
    }

    // Move, reorder, snapshot, remove and restore keep the chart's id.
    wb.sheet(0).setDrawingAnchor(id, { row: 10, col: 1, widthEmu: 3000000, heightEmu: 2000000 });
    const moved = wb
      .sheet(0)
      .listDrawingObjects()
      .find((o) => o.objectId === id);
    assert.equal(moved.fromRow, 10);
    assert.equal(moved.fromCol, 1);
    assert.equal(moved.widthEmu, 3000000);
    assert.equal(moved.heightEmu, 2000000);
    wb.sheet(0).setDrawingZOrder(id, 0);
    assert.deepEqual(
      Array.from(wb.sheet(0).listDrawingObjects(), (o) => o.objectId),
      [id, other],
    );

    const snap = wb.sheet(0).snapshotDrawing(id);
    assert.ok(snap instanceof Uint8Array);
    assert.ok(snap.length > 0);
    wb.sheet(0).removeDrawingObject(id);
    assert.equal(wb.sheet(0).getChart(id), null);
    assert.deepEqual(
      Array.from(wb.sheet(0).listDrawingObjects(), (o) => o.objectId),
      [other],
    );

    const restored = wb.sheet(0).restoreDrawing(snap);
    assert.equal(restored, id);
    assert.equal(chartOf(wb, id).title.text, 'Kept');
    assert.deepEqual(
      Array.from(wb.sheet(0).listDrawingObjects(), (o) => o.objectId),
      [id, other],
    );

    const copy = wb.sheet(0).restoreDrawing(snap, { newId: true });
    assert.notEqual(copy, id);
    assert.equal(chartOf(wb, copy).title.text, 'Kept');
    assert.equal(wb.sheet(0).listDrawingObjects().length, 3);
  } finally {
    wb.dispose();
  }
});

test('chart calls throw on failure and reject malformed arguments', async () => {
  const mod = await load();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod);

    const missing = wb.sheet(0).getChart(9999);
    assert.equal(missing, null);
    expectFailure(() => wb.sheet(99).getChart(id));
    expectFailure(() => wb.sheet(0).getChartData(9999));

    // An unresolvable source writes nothing.
    expectFailure(
      () => wb.sheet(0).addChart({ kind: mod.ChartInsertKind.Column, source: 'NoSuchSheet!A1:B2' }),
      mod.StatusCode.IoChartInvalidReference,
    );
    assert.equal(wb.sheet(0).listDrawingObjects().length, 1);
    expectFailure(() => wb.sheet(0).addChart({ kind: 99, source: 'A1:B2' }));
    assert.throws(() => wb.sheet(0).addChart('A1:B2'), TypeError);

    expectFailure(() => wb.sheet(0).setChartSeries(id, 0, 0, { values: '' }));
    expectFailure(
      () => wb.sheet(0).setChartSeries(id, 0, 0, { values: 'NoSuchSheet!A1:A2' }),
      mod.StatusCode.IoChartInvalidReference,
    );
    expectFailure(() => wb.sheet(0).setChartSeries(id, 5, 0, {}));
    expectFailure(() => wb.sheet(0).setChartSeries(id, 0, 99, {}));
    expectFailure(() => wb.sheet(0).setChartSeries(id, 0, 0, { fill: { color: 1 }, marker: { symbol: null } }));
    expectFailure(() => wb.sheet(0).setChartAxis(id, 12345, {}));
    expectFailure(() => wb.sheet(0).setChartPlot(id, 0, { holeSize: 50 }));
    expectFailure(() => wb.sheet(0).setChartType(id, 0, 99));
    expectFailure(
      () => wb.sheet(0).setChartSource(id, { source: 'NoSuchSheet!A1:B2', direction: mod.ChartSeriesDirection.Auto }),
      mod.StatusCode.IoChartInvalidReference,
    );
    expectFailure(() => wb.sheet(0).removeChartSeries(id, 0, 99));
    expectFailure(() => wb.sheet(0).addChartSeries(id, 0, { name: 'Sheet1!$B$1' }));
    expectFailure(() => wb.sheet(0).setChartTitle(9999, 'x'));
    expectFailure(() => wb.sheet(0).removeDrawingObject(9999));
    expectFailure(() => wb.sheet(0).snapshotDrawing(9999));

    // Wrong types raise before the engine is called.
    assert.throws(() => wb.sheet(0).getChart(String(id)), TypeError);
    assert.throws(() => wb.sheet(0).getChart(-1), RangeError);
    assert.throws(() => wb.sheet(0).setChartType(id, 0, 1.5), RangeError);
    assert.throws(() => wb.sheet(0).setChartSeries(id, 0, 0, { fill: { color: 'red' } }), TypeError);
    assert.throws(() => wb.sheet(0).setChartSeries(id, 0, 0, { values: 5 }), TypeError);
    assert.throws(() => wb.sheet(0).setChartSeries(id, 0, 0, 'nope'), TypeError);
    assert.throws(() => wb.sheet(0).setChartAxis(id, 1, { min: 'low' }), TypeError);
    assert.throws(() => wb.sheet(0).setChartFilter(id, [1, 'x'], []), TypeError);
    assert.throws(() => wb.sheet(0).setChartFilter(id, 'x', []), TypeError);
    assert.throws(() => wb.sheet(0).setChartTitle(id, 5), TypeError);
  } finally {
    wb.dispose();
  }
});
