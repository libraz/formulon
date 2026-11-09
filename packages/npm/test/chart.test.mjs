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
    type: mod.ChartPlotType.Bar,
    range: 'Sheet1!$A$1:$D$5',
    legend: mod.ChartLegendPosition.Right,
    ...extra,
  });
  assert.ok(r.objectId > 0);
  assert.equal(wb.sheet(0).chart(r.objectId), r);
  return r.objectId;
}

const drawingsOf = (wb) =>
  wb
    .sheet(0)
    .listDrawings()
    .map((d) => d.getDrawingProperties());
const chartOf = (wb, id) => wb.sheet(0).chart(id).getProperties();
const dataOf = (wb, id) => wb.sheet(0).chart(id).getData();

test('chart enums keep their C ordinals', async () => {
  const mod = await load();
  assert.equal(mod.ChartPlotType.Bar, 0);
  assert.equal(mod.ChartPlotType.Extended, 16);
  assert.equal(mod.ChartPlotType.Scatter, 10);
  assert.equal(mod.ChartPlotType.Doughnut, 8);
  assert.equal(mod.ChartBarDirection.Col, 0);
  assert.equal(mod.ChartBarDirection.Bar, 1);
  assert.equal(mod.ChartInsertKind, undefined);
  assert.equal(mod.ChartGrouping.Standard, 3);
  assert.equal(mod.BlanksAs.Span, 2);
  assert.equal(mod.ChartPointState.Connect, 2);
  assert.equal(mod.ChartAxisKind.Series, 3);
  assert.equal(mod.ChartAxisCrosses.At, 3);
  assert.equal(mod.ChartTitleKind.Auto, 3);
  assert.equal(mod.ChartSeriesDirection.Rows, 2);
  assert.deepEqual({ ...mod.ChartSourceKind }, { None: 0, Ref: 1, Literal: 2 });
  assert.equal(mod.ChartLegendPosition.TopRight, 2);
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
      placement: { name: 'Revenue', description: 'alt', anchor: { fromRow: 7, fromCol: 2 } },
    });
    const listed = drawingsOf(wb);
    assert.equal(listed.length, 1);
    assert.equal(listed[0].objectId, id);
    assert.equal(listed[0].kind, mod.DrawingObjectKind.Chart);
    assert.equal(listed[0].anchor.kind, mod.AnchorKind.TwoCell);
    assert.equal(listed[0].anchor.fromRow, 7);
    assert.equal(listed[0].anchor.fromCol, 2);
    assert.equal(listed[0].name, 'Revenue');
    assert.equal(listed[0].description, 'alt');
    assert.equal(wb.sheet(0).chart(id).name, 'Revenue');

    const chart = chartOf(wb, id);
    assert.equal(chart.objectId, id);
    assert.equal('name' in chart, false);
    assert.equal('description' in chart, false);
    assert.equal(chart.isExtended, false);
    assert.equal(chart.isPivotChart, false);
    assert.equal(chart.title.kind, mod.ChartTitleKind.Text);
    assert.equal(chart.title.text, 'Sales');
    assert.equal(chart.legend.position, mod.ChartLegendPosition.Right);
    assert.deepEqual(chart.legend.hiddenEntries, []);
    assert.equal(chart.plots.length, 1);
    const plot = chart.plots[0];
    assert.equal(plot.type, mod.ChartPlotType.Bar);
    assert.equal(plot.barDirection, mod.ChartBarDirection.Col);
    assert.equal(plot.axisIds.length, 2);
    assert.equal(plot.series.length, 3);
    const s0 = plot.series[0];
    assert.equal(s0.name.kind, mod.ChartSourceKind.Ref);
    assert.equal(s0.name.formula, 'Sheet1!$B$1');
    assert.equal(s0.categories.formula, 'Sheet1!$A$2:$A$5');
    assert.equal(s0.values.formula, 'Sheet1!$B$2:$B$5');
    assert.equal(s0.x.kind, mod.ChartSourceKind.None);
    // A new chart's series take the palette colour as a theme reference (accent1 first).
    assert.equal(s0.fill.color.kind, mod.ColorKind.Theme);
    assert.equal(s0.fill.color.theme, 4);
    assert.equal(s0.marker, null);
    assert.equal(s0.labels, null);
    assert.deepEqual(s0.points, []);
    assert.equal(chart.axes.length, 2);
    assert.deepEqual(chart.axes.map((a) => a.id).sort(), [...plot.axisIds].sort());
    assert.ok(Object.values(mod.ChartAxisPosition).includes(chart.axes[0].position));
    assert.deepEqual(
      { sheet: chart.source.range.sheet, ...chart.source.range.rect },
      { sheet: 0, firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 3 },
    );
    assert.equal(chart.source.direction, mod.ChartSeriesDirection.Columns);
    assert.equal(chart.source.headerRow, true);
    assert.equal(chart.source.headerColumn, true);

    const data = dataOf(wb, id);
    assert.equal(data.domain, mod.ChartDomain.Category);
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
    const types = [mod.ChartPlotType.Line, mod.ChartPlotType.Area, mod.ChartPlotType.Pie, mod.ChartPlotType.Doughnut];
    for (const type of types) {
      const id = insert(wb, mod, { type, range: 'Sheet1!$A$1:$B$5' });
      assert.equal(chartOf(wb, id).plots[0].type, type);
    }
    const column = insert(wb, mod, { range: 'Sheet1!$A$1:$B$5' });
    assert.equal(chartOf(wb, column).plots[0].barDirection, mod.ChartBarDirection.Col);
    const bar = insert(wb, mod, {
      type: mod.ChartPlotType.Bar,
      barDirection: mod.ChartBarDirection.Bar,
      grouping: mod.ChartGrouping.Stacked,
      range: 'A1:D5',
      direction: mod.ChartSeriesDirection.Columns,
      headerRow: true,
      headerColumn: true,
      title: '',
    });
    const barPlot = chartOf(wb, bar).plots[0];
    assert.equal(barPlot.barDirection, mod.ChartBarDirection.Bar);
    assert.equal(barPlot.grouping, mod.ChartGrouping.Stacked);
    assert.equal(chartOf(wb, bar).title, null);

    const xXf = wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 2 });
    wb.sheet(0).ooxml.setRangeXfIndex({ firstRow: 1, firstCol: 5, lastRow: 4, lastCol: 5 }, xXf);
    const scatter = insert(wb, mod, { type: mod.ChartPlotType.Scatter, range: 'Sheet1!$F$1:$G$5' });
    const sdata = dataOf(wb, scatter);
    assert.equal(sdata.domain, mod.ChartDomain.Xy);
    assert.ok(sdata.series[0].x instanceof Float64Array);
    assert.deepEqual([...sdata.series[0].x], [0, 3, 6, 9]);
    assert.equal(sdata.series[0].xNumberFormat, '0.00');
    assert.equal(sdata.series[0].numberFormat, 'General');

    // An omitted title is Excel's automatic one.
    const auto = insert(wb, mod);
    assert.equal(chartOf(wb, auto).title.kind, mod.ChartTitleKind.Auto);
    assert.equal(drawingsOf(wb).at(-1).anchor.kind, mod.AnchorKind.TwoCell);
  } finally {
    wb.dispose();
  }
});

test('typed chart edits are visible through getChart', async () => {
  const mod = await load();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod);

    wb.sheet(0).chart(id).setTitle('Quarterly');
    assert.equal(chartOf(wb, id).title.text, 'Quarterly');
    wb.sheet(0).chart(id).setTitle(null);
    assert.equal(chartOf(wb, id).title, null);
    wb.sheet(0).chart(id).setTitle({ kind: mod.ChartTitleKind.Auto });
    assert.equal(chartOf(wb, id).title.kind, mod.ChartTitleKind.Auto);

    wb.sheet(0).chart(id).setLegend({ position: mod.ChartLegendPosition.Bottom, overlay: true, hiddenEntries: [] });
    let legend = chartOf(wb, id).legend;
    assert.equal(legend.position, mod.ChartLegendPosition.Bottom);
    assert.equal(legend.overlay, true);
    wb.sheet(0)
      .chart(id)
      .setLegend({
        position: mod.ChartLegendPosition.Bottom,
        overlay: false,
        hiddenEntries: [1],
      });
    legend = chartOf(wb, id).legend;
    assert.equal(legend.overlay, false);
    assert.deepEqual(legend.hiddenEntries, [1]);
    wb.sheet(0).chart(id).setLegend(null);
    assert.equal(chartOf(wb, id).legend, null);

    // A new chart carries `dispNaAsBlank`.
    assert.equal(chartOf(wb, id).naAsBlank, true);
    wb.sheet(0).chart(id).updateDisplay({ blanksAs: mod.BlanksAs.Zero, plotVisibleOnly: false });
    let info = chartOf(wb, id);
    assert.equal(info.blanksAs, mod.BlanksAs.Zero);
    assert.equal(info.plotVisibleOnly, false);
    // An omitted key keeps, including `naAsBlank`.
    assert.equal(info.naAsBlank, true);
    wb.sheet(0).chart(id).updateDisplay({ blanksAs: null, plotVisibleOnly: null });
    info = chartOf(wb, id);
    assert.equal(info.blanksAs, null);
    assert.equal(info.plotVisibleOnly, null);
    assert.equal(info.naAsBlank, true);
    wb.sheet(0).chart(id).updateDisplay({ plotVisibleOnly: true });
    assert.equal(chartOf(wb, id).blanksAs, null);
    assert.equal(chartOf(wb, id).plotVisibleOnly, true);
    // `false` and `null` both remove the extension, which reads back as `null`.
    wb.sheet(0).chart(id).updateDisplay({ naAsBlank: false });
    assert.equal(chartOf(wb, id).naAsBlank, null);
    wb.sheet(0).chart(id).updateDisplay({ naAsBlank: true });
    assert.equal(chartOf(wb, id).naAsBlank, true);
    wb.sheet(0).chart(id).updateDisplay({ naAsBlank: null });
    assert.equal(chartOf(wb, id).naAsBlank, null);
    assert.equal(chartOf(wb, id).plotVisibleOnly, true);
    expectFailure(() => wb.sheet(0).chart(id).updateDisplay({ blanksAs: -1 }));
    assert.throws(() => wb.sheet(0).chart(id).updateDisplay(5), TypeError);
    wb.sheet(0).chart(id).getPreview({ naAsBlank: true });
    wb.sheet(0).chart(id).getPreview({ naAsBlank: false });
    assert.equal(chartOf(wb, id).naAsBlank, null);

    wb.sheet(0).chart(id).updatePlot(0, { gapWidthPct: 40, overlap: 10, varyColors: true });
    const plot = chartOf(wb, id).plots[0];
    assert.equal(plot.gapWidthPct, 40);
    assert.equal(plot.overlap, 10);
    assert.equal(plot.varyColors, true);

    const valueAxis = chartOf(wb, id).axes.find((a) => a.kind === mod.ChartAxisKind.Value);
    wb.sheet(0).chart(id).updateAxis(valueAxis.id, {
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
    wb.sheet(0)
      .chart(id)
      .updateAxis(valueAxis.id, { title: { kind: mod.ChartTitleKind.Auto } });
    assert.equal(chartOf(wb, id).axes.find((a) => a.id === valueAxis.id).title.kind, mod.ChartTitleKind.Auto);
    wb.sheet(0)
      .chart(id)
      .updateAxis(valueAxis.id, { title: { kind: mod.ChartTitleKind.Ref, formula: 'Sheet1!$A$1' } });
    assert.equal(chartOf(wb, id).axes.find((a) => a.id === valueAxis.id).title.formula, 'Sheet1!$A$1');
    // null clears a bound.
    wb.sheet(0).chart(id).updateAxis(valueAxis.id, { max: null });
    assert.equal(chartOf(wb, id).axes.find((a) => a.id === valueAxis.id).max, null);

    // Markers and smoothing exist on line plots only.
    wb.sheet(0).chart(id).setType(0, mod.ChartPlotType.Line);
    assert.equal(chartOf(wb, id).plots[0].type, mod.ChartPlotType.Line);
    wb.sheet(0)
      .chart(id)
      .updateSeries(0, 1, {
        fill: { color: '#FF0000', transparencyPct: 25 },
        line: { color: '#0000FF', widthPt: 2, dash: mod.ChartLineDash.Dash },
        marker: { symbol: mod.ChartMarkerSymbol.Circle, sizePt: 7 },
        smooth: true,
        labels: { showValue: true, position: mod.ChartLabelPosition.OutsideEnd, numberFormat: '0.00' },
      });
    let series = chartOf(wb, id).plots[0].series[1];
    // A line series takes no fill of its own; its colour is the line's.
    assert.equal(series.fill.color, null);
    assert.equal(series.fill.transparencyPct, null);
    assert.equal(series.line.color.argb, 0xff0000ff);
    assert.equal(series.line.widthPt, 2);
    assert.equal(series.line.dash, mod.ChartLineDash.Dash);
    assert.equal(series.marker.symbol, mod.ChartMarkerSymbol.Circle);
    assert.equal(series.marker.sizePt, 7);
    assert.equal(series.smooth, true);
    assert.equal(series.labels.showValue, true);
    assert.equal(series.labels.position, mod.ChartLabelPosition.OutsideEnd);
    assert.equal(series.labels.numberFormat, '0.00');
    // null removes the marker and the labels; an omitted field is kept.
    wb.sheet(0).chart(id).updateSeries(0, 1, { marker: null, labels: null });
    series = chartOf(wb, id).plots[0].series[1];
    assert.equal(series.marker, null);
    assert.equal(series.labels, null);
    assert.equal(series.line.color.argb, 0xff0000ff);

    // Reference edits refill the cache and resolved data follows.
    wb.sheet(0).chart(id).updateSeries(0, 0, { values: 'Sheet1!$D$2:$D$5', name: '' });
    series = chartOf(wb, id).plots[0].series[0];
    assert.equal(series.values.formula, 'Sheet1!$D$2:$D$5');
    assert.equal(series.name.kind, mod.ChartSourceKind.None);
    assert.deepEqual([...dataOf(wb, id).series[0].values], [4, 8, 12, 16]);

    // A line point takes no fill; its line carries the colour.
    const otherPoints = chartOf(wb, id).plots[0].series[1].points.filter((point) => point.index !== 2);
    wb.sheet(0)
      .chart(id)
      .updatePoint(0, 1, 2, { line: { color: '#00FF00' } });
    const points = chartOf(wb, id).plots[0].series[1].points;
    assert.equal(points.length, otherPoints.length + 1);
    assert.equal(points.find((point) => point.index === 2).line.color.argb, 0xff00ff00);
    assert.deepEqual(
      points.filter((point) => point.index !== 2),
      otherPoints,
    );
    wb.sheet(0).chart(id).setPoint(0, 1, 2, null);
    assert.deepEqual(chartOf(wb, id).plots[0].series[1].points, otherPoints);

    wb.sheet(0)
      .chart(id)
      .setFilter({ hiddenSeries: [2], hiddenCategories: [1] });
    info = chartOf(wb, id);
    assert.equal(info.plots[0].series.find((s) => s.seriesId === 2).filtered, true);
    assert.deepEqual(info.filter, { hiddenSeries: [2], hiddenCategories: [1] });
    const filtered = dataOf(wb, id);
    assert.ok(filtered.series.find((s) => s.seriesId === 2).filtered);
    wb.sheet(0).chart(id).setFilter({});
    assert.deepEqual(chartOf(wb, id).filter, { hiddenSeries: [], hiddenCategories: [] });

    wb.sheet(0).chart(id).setType(0, mod.ChartPlotType.Bar, { grouping: mod.ChartGrouping.Stacked });
    assert.equal(chartOf(wb, id).plots[0].type, mod.ChartPlotType.Bar);
    assert.equal(chartOf(wb, id).plots[0].grouping, mod.ChartGrouping.Stacked);

    const pie = insert(wb, mod, { type: mod.ChartPlotType.Doughnut, range: 'Sheet1!$A$1:$B$5' });
    wb.sheet(0).chart(pie).updatePlot(0, { holeSize: 60, firstSliceAngleDeg: 90 });
    assert.equal(chartOf(wb, pie).plots[0].holeSize, 60);
    assert.equal(chartOf(wb, pie).plots[0].firstSliceAngleDeg, 90);
  } finally {
    wb.dispose();
  }
});

test('series can be added, removed and rebuilt from a new source', async () => {
  const mod = await load();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod, { range: 'Sheet1!$A$1:$C$5' });
    assert.equal(chartOf(wb, id).plots[0].series.length, 2);

    const added = wb.sheet(0).chart(id).addSeries(0, {
      values: 'Sheet1!$D$2:$D$5',
      name: 'Sheet1!$D$1',
      categories: 'Sheet1!$A$2:$A$5',
    });
    assert.equal(chartOf(wb, id).plots[0].series.length, 3);
    assert.ok(chartOf(wb, id).plots[0].series.some((s) => s.seriesId === added));
    assert.equal(dataOf(wb, id).series.length, 3);

    wb.sheet(0).chart(id).removeSeries(0, added);
    assert.equal(chartOf(wb, id).plots[0].series.length, 2);

    wb.sheet(0)
      .chart(id)
      .updateSeries(0, 0, { fill: { color: '#112233' } });
    wb.sheet(0).chart(id).setSource({
      range: 'Sheet1!$A$1:$D$5',
      direction: mod.ChartSeriesDirection.Columns,
      headerRow: true,
      headerColumn: true,
    });
    let chart = chartOf(wb, id);
    assert.equal(chart.plots[0].series.length, 3);
    assert.equal(chart.plots[0].series[0].fill.color.argb, 0xff112233);
    assert.deepEqual(
      { sheet: chart.source.range.sheet, ...chart.source.range.rect },
      { sheet: 0, firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 3 },
    );

    // Rows direction swaps the split: four series of three points.
    wb.sheet(0).chart(id).setSource({ range: 'Sheet1!$A$1:$D$5', direction: mod.ChartSeriesDirection.Rows });
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
    const preview = wb.sheet(0).chart(id).getPreview({ range: 'Sheet1!$A$1:$B$5', blanksAs: mod.BlanksAs.Gap });
    assert.equal(preview.series.length, 1);
    assert.deepEqual([...preview.series[0].values], [2, 4, 6, 8]);
    const rows = wb.sheet(0).chart(id).getPreview({
      direction: mod.ChartSeriesDirection.Rows,
      headerRow: true,
      headerColumn: true,
    });
    assert.equal(rows.series.length, 4);
    assert.equal(JSON.stringify(chartOf(wb, id)), before);
    // No overrides is the chart as it stands.
    assert.equal(wb.sheet(0).chart(id).getPreview({}).series.length, 3);
    expectFailure(() => wb.sheet(0).chart(id).getPreview({ range: 'NoSuchSheet!A1:B2' }));
    expectFailure(() => wb.sheet(0).chart(9999).getPreview({}));
  } finally {
    wb.dispose();
  }
});

test('a chart survives save and load and goes through the drawing operations', async () => {
  const mod = await load();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod, { title: 'Kept', placement: { anchor: { fromRow: 3, fromCol: 3 } } });
    const other = insert(wb, mod, { type: mod.ChartPlotType.Line, range: 'Sheet1!$A$1:$B$5' });

    const saved = wb.save();
    const wb2 = mod.Workbook.load(saved);
    try {
      assert.equal(chartOf(wb2, id).title.text, 'Kept');
      assert.deepEqual([...dataOf(wb2, id).series[0].values], [2, 4, 6, 8]);
      assert.equal(drawingsOf(wb2).filter((o) => o.kind === mod.DrawingObjectKind.Chart).length, 2);
    } finally {
      wb2.dispose();
    }

    // Move, reorder, snapshot, remove and restore keep the chart's id.
    wb.sheet(0)
      .chart(id)
      .update({
        anchor: { kind: mod.AnchorKind.OneCell, fromRow: 10, fromCol: 1, widthEmu: 3000000, heightEmu: 2000000 },
      });
    const moved = drawingsOf(wb).find((o) => o.objectId === id).anchor;
    assert.equal(moved.kind, mod.AnchorKind.OneCell);
    assert.equal(moved.fromRow, 10);
    assert.equal(moved.fromCol, 1);
    assert.equal(moved.widthEmu, 3000000);
    assert.equal(moved.heightEmu, 2000000);
    wb.sheet(0).chart(id).setZOrder(0);
    assert.deepEqual(
      Array.from(drawingsOf(wb), (o) => o.objectId),
      [id, other],
    );

    const snap = wb.sheet(0).chart(id).snapshot();
    assert.ok(snap instanceof Uint8Array);
    assert.ok(snap.length > 0);
    const handle = wb.sheet(0).chart(id);
    wb.sheet(0).removeDrawing(id);
    expectFailure(() => handle.getProperties(), mod.StatusCode.NotFound);
    expectFailure(() => wb.sheet(0).chart(id), mod.StatusCode.NotFound);
    assert.deepEqual(
      Array.from(drawingsOf(wb), (o) => o.objectId),
      [other],
    );

    const restored = wb.sheet(0).restoreDrawing(snap);
    assert.equal(restored, handle);
    assert.equal(restored.objectId, id);
    assert.equal(chartOf(wb, id).title.text, 'Kept');
    assert.deepEqual(
      Array.from(drawingsOf(wb), (o) => o.objectId),
      [id, other],
    );

    const copy = wb.sheet(0).restoreDrawing(snap, { assignNewId: true });
    assert.notEqual(copy, id);
    assert.equal(chartOf(wb, copy).title.text, 'Kept');
    assert.equal(drawingsOf(wb).length, 3);
  } finally {
    wb.dispose();
  }
});

test('chart calls throw on failure and reject malformed arguments', async () => {
  const mod = await load();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod);

    expectFailure(() => wb.sheet(0).chart(9999), mod.StatusCode.NotFound);
    expectFailure(() => wb.sheet(99).chart(id));

    // A source on a sheet the workbook does not have writes nothing.
    expectFailure(
      () => wb.sheet(0).addChart({ type: mod.ChartPlotType.Bar, range: 'NoSuchSheet!A1:B2' }),
      mod.StatusCode.NotFound,
    );
    assert.equal(drawingsOf(wb).length, 1);
    expectFailure(() => wb.sheet(0).addChart({ type: 99, range: 'A1:B2' }));
    assert.throws(() => wb.sheet(0).addChart('A1:B2'), TypeError);

    expectFailure(() => wb.sheet(0).chart(id).updateSeries(0, 0, { values: '' }));
    expectFailure(
      () => wb.sheet(0).chart(id).updateSeries(0, 0, { values: 'NoSuchSheet!A1:A2' }),
      mod.StatusCode.IoChartInvalidReference,
    );
    expectFailure(() => wb.sheet(0).chart(id).updateSeries(5, 0, {}));
    expectFailure(() => wb.sheet(0).chart(id).updateSeries(0, 99, {}));
    expectFailure(() =>
      wb
        .sheet(0)
        .chart(id)
        .updateSeries(0, 0, { fill: { color: '#010203' }, marker: { symbol: null } }),
    );
    expectFailure(() => wb.sheet(0).chart(id).updateAxis(12345, {}));
    expectFailure(() => wb.sheet(0).chart(id).updatePlot(0, { holeSize: 50 }));
    expectFailure(() => wb.sheet(0).chart(id).setType(0, 99));
    expectFailure(
      () => wb.sheet(0).chart(id).setSource({ range: 'NoSuchSheet!A1:B2', direction: mod.ChartSeriesDirection.Auto }),
      mod.StatusCode.NotFound,
    );
    expectFailure(() => wb.sheet(0).chart(id).removeSeries(0, 99));
    expectFailure(() => wb.sheet(0).chart(id).addSeries(0, { name: 'Sheet1!$B$1' }));
    expectFailure(() => wb.sheet(0).chart(9999).setTitle('x'));
    expectFailure(() => wb.sheet(0).removeDrawing(9999), mod.StatusCode.NotFound);
    expectFailure(() => wb.sheet(0).chart(9999).snapshot());

    // Wrong types raise before the engine is called.
    assert.throws(() => wb.sheet(0).chart(String(id)).getProperties(), TypeError);
    expectFailure(() => wb.sheet(0).chart(-1).getProperties(), mod.StatusCode.InvalidArgument);
    assert.throws(() => wb.sheet(0).chart(1.5).getProperties(), RangeError);
    assert.throws(() => wb.sheet(0).chart(id).setType(0, 1.5), RangeError);
    assert.throws(
      () =>
        wb
          .sheet(0)
          .chart(id)
          .updateSeries(0, 0, { fill: { color: 'red' } }),
      TypeError,
    );
    assert.throws(() => wb.sheet(0).chart(id).updateSeries(0, 0, { values: 5 }), TypeError);
    assert.throws(() => wb.sheet(0).chart(id).updateSeries(0, 0, 'nope'), TypeError);
    assert.throws(() => wb.sheet(0).chart(id).updateAxis(1, { min: 'low' }), TypeError);
    assert.throws(
      () =>
        wb
          .sheet(0)
          .chart(id)
          .setFilter({ hiddenSeries: [1, 'x'] }),
      TypeError,
    );
    assert.throws(() => wb.sheet(0).chart(id).setFilter({ hiddenSeries: 'x' }), TypeError);
    assert.throws(() => wb.sheet(0).chart(id).setTitle(5), TypeError);
  } finally {
    wb.dispose();
  }
});

test('a theme colour round-trips through a series fill and is refused when indexed', async () => {
  const mod = await load();
  const wb = newBook(mod);
  try {
    const id = insert(wb, mod);
    const theme = { kind: mod.ColorKind.Theme, argb: 0, theme: 4, tint: 0.4, indexed: 0 };
    wb.sheet(0)
      .chart(id)
      .updateSeries(0, 0, { fill: { color: theme } });
    const got = chartOf(wb, id).plots[0].series[0].fill.color;
    assert.equal(got.kind, mod.ColorKind.Theme);
    assert.equal(got.theme, 4);
    assert.ok(Math.abs(got.tint - 0.4) < 1e-5);

    const indexed = { kind: mod.ColorKind.Indexed, argb: 0, theme: 0, tint: 0, indexed: 5 };
    expectFailure(
      () =>
        wb
          .sheet(0)
          .chart(id)
          .updateSeries(0, 0, { fill: { color: indexed } }),
      mod.StatusCode.InvalidArgument,
    );

    const shape = wb.sheet(0).addShape({
      presetGeometry: 'rect',
      anchor: { kind: mod.AnchorKind.OneCell, fromRow: 2, widthEmu: 9144, heightEmu: 9144 },
      fillKind: mod.ShapePaint.Solid,
      fillColor: theme,
    });
    const paint = shape.getProperties().fillColor;
    assert.equal(paint.kind, mod.ColorKind.Theme);
    assert.equal(paint.theme, 4);
    assert.ok(Math.abs(paint.tint - 0.4) < 1e-5);
    expectFailure(
      () => shape.update({ fillKind: mod.ShapePaint.Solid, fillColor: indexed }),
      mod.StatusCode.InvalidArgument,
    );
    // A colour is never dropped: without its kind it is a TypeError, with None a status error.
    assert.throws(() => shape.update({ fillColor: theme }), TypeError);
    assert.throws(() => shape.update({ lineKind: null, lineColor: theme }), TypeError);
    assert.throws(
      () =>
        wb
          .sheet(0)
          .addShape({ anchor: { kind: mod.AnchorKind.OneCell, widthEmu: 9144, heightEmu: 9144 }, lineColor: theme }),
      TypeError,
    );
    expectFailure(
      () => shape.update({ fillKind: mod.ShapePaint.None, fillColor: theme }),
      mod.StatusCode.InvalidArgument,
    );
    assert.equal(shape.getProperties().fillColor.theme, 4);
  } finally {
    wb.dispose();
  }
});
