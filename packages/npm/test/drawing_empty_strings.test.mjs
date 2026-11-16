// Empty strings in the DRAWING / SHAPE / SLICER / CHART / SPARKLINE setters (api-conventions §5 rule 30): where the
// file holds one state for '' and the absent value they mean the same; where '' is not a legal value it is refused
// and the absent value clears or derives.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

const isFormulonError = (e) => e.name === 'FormulonError';

async function withWorkbook(fn) {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

function png(width, height) {
  const bytes = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52];
  for (const v of [width, height]) bytes.push((v >>> 24) & 0xff, (v >>> 16) & 0xff, (v >>> 8) & 0xff, v & 0xff);
  bytes.push(8, 6, 0, 0, 0, 0x1f, 0x15, 0xc4, 0x89);
  return new Uint8Array(bytes);
}

const isStatus = (code) => (e) => e.name === 'FormulonError' && e.code === code;
const box = { fromRow: 1, fromCol: 1, toRow: 5, toCol: 4 };

function fillGrid(ws) {
  for (let c = 1; c <= 3; c += 1) {
    ws.setText(0, c, 'S');
    for (let r = 1; r <= 4; r += 1) ws.setNumber(r, c, r * (c + 1));
  }
  for (let r = 1; r <= 4; r += 1) ws.setText(r, 0, 'a');
}

function salesTable(ws) {
  ws.setText(8, 0, 'Product');
  ws.setText(9, 0, 'Widget');
  return ws.addTable({
    range: { firstRow: 8, firstCol: 0, lastRow: 9, lastCol: 0 },
    name: 'Sales',
    columns: ['Product'],
  });
}

test("a picture description of '' and the absent value are one state", async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const empty = ws.addPicture(png(4, 4), { description: '' });
    const absent = ws.addPicture(png(4, 4));
    assert.ok(!empty.getDrawingProperties().description);
    assert.ok(!absent.getDrawingProperties().description);
  });
});

test("a placement name of '' is InvalidArgument on every kind and the absent value derives one", async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    fillGrid(ws);
    const table = salesTable(ws);
    const invalid = isStatus(f.StatusCode.InvalidArgument);
    const chart = { type: f.ChartPlotType.Line, range: 'Sheet1!$A$1:$D$5', anchor: box };
    assert.throws(() => ws.addPicture(png(4, 4), { name: '' }), invalid);
    assert.throws(() => ws.addShape({ name: '', anchor: box }), invalid);
    assert.throws(() => ws.addChart({ ...chart, name: '' }), invalid);
    assert.throws(() => ws.addSlicer(table, 0, { name: '', anchor: { ...box, fromCol: 6, toCol: 9 } }), invalid);
    assert.equal(ws.listDrawings().length, 0);

    const made = [
      ws.addPicture(png(4, 4)),
      ws.addShape({ anchor: box }),
      ws.addChart(chart),
      ws.addSlicer(table, 0, { anchor: { ...box, fromCol: 6, toCol: 9 } }),
    ];
    for (const drawing of made) assert.notEqual(drawing.getDrawingProperties().name, '');
  });
});

test("a drawing patch description of '' removes the description, as null does", async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const shape = ws.addShape({ anchor: box, description: 'alt' });
    assert.equal(shape.getDrawingProperties().description, 'alt');
    shape.update({ description: '' });
    assert.ok(!shape.getDrawingProperties().description);
    shape.update({ description: 'again' });
    shape.update({ description: null });
    assert.ok(!shape.getDrawingProperties().description);
  });
});

test("a cell picture altText of '' means none", async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setCellPicture('A1', { bytes: png(2, 2), altText: 'old' });
    ws.setCellPicture('A1', { bytes: png(2, 2), altText: '' });
    assert.equal(ws.getCellPicture('A1').altText, null);
  });
});

test("a shape presetGeometry of '' is InvalidArgument and the absent value is rect; text '' is none", async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.throws(() => ws.addShape({ anchor: box, presetGeometry: '' }), isStatus(f.StatusCode.InvalidArgument));
    assert.equal(ws.listDrawings().length, 0);
    assert.equal(ws.addShape({ anchor: box }).getProperties().presetGeometry, 'rect');
    const emptyText = ws.addShape({ anchor: box, text: '' });
    const absentText = ws.addShape({ anchor: box });
    assert.equal(emptyText.getProperties().text, absentText.getProperties().text);
  });
});

test("a slicer caption of '' and the absent value both take the column name", async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const table = salesTable(ws);
    const empty = ws.addSlicer(table, 0, { anchor: box, caption: '' });
    const absent = ws.addSlicer(table, 0, { anchor: { ...box, fromCol: 6, toCol: 9 } });
    assert.equal(empty.getProperties().caption, 'Product');
    assert.equal(absent.getProperties().caption, 'Product');
  });
});

test("a sparkline source of '' is InvalidArgument and null is a deleted source", async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const invalid = isStatus(f.StatusCode.InvalidArgument);
    assert.throws(() => ws.addSparklineGroup({ sparklines: [{ source: '', location: 'F1' }] }), invalid);
    assert.equal(ws.listSparklineGroups().length, 0);
    ws.addSparklineGroup({ sparklines: [{ source: null, location: 'F1' }] });
    assert.equal(ws.listSparklineGroups()[0].sparklines[0].source, null);
    assert.throws(() => ws.updateSparklineGroup(0, { sparklines: [{ source: '', location: 'F2' }] }), invalid);
  });
});

test("a sparkline dateAxisRange of '' is InvalidArgument and null is no date axis", async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const invalid = isStatus(f.StatusCode.InvalidArgument);
    const sparklines = [{ source: 'Sheet1!A1:D1', location: 'F1' }];
    assert.throws(() => ws.addSparklineGroup({ sparklines, dateAxisRange: '' }), invalid);
    assert.equal(ws.listSparklineGroups().length, 0);
    const id = ws.addSparklineGroup({ sparklines, dateAxisRange: null });
    assert.equal(ws.listSparklineGroups()[0].dateAxisRange, null);
    assert.throws(() => ws.updateSparklineGroup(id, { dateAxisRange: '' }), invalid);
  });
});

test("a sparkline location of '' fails as an A1 text that does not parse", async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.throws(
      () => ws.addSparklineGroup({ sparklines: [{ source: 'Sheet1!A1:D1', location: '' }] }),
      (e) => isFormulonError(e) && e.code === f.StatusCode.ParserInvalidReference,
    );
    assert.equal(ws.listSparklineGroups().length, 0);
  });
});

test("a chart series reference of '' is InvalidArgument for every input and null removes name, categories and x", async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    fillGrid(ws);
    const invalid = isStatus(f.StatusCode.InvalidArgument);
    const chart = ws.addChart({ type: f.ChartPlotType.Line, range: 'Sheet1!$A$1:$D$5', anchor: box });
    for (const key of ['name', 'categories', 'x', 'values']) {
      assert.throws(() => chart.updateSeries(0, 0, { [key]: '' }), invalid, key);
    }
    chart.updateSeries(0, 0, { name: 'Sheet1!$B$1' });
    assert.equal(chart.getProperties().plots[0].series[0].name.kind, f.ChartSourceKind.Ref);
    chart.updateSeries(0, 0, { name: null });
    assert.equal(chart.getProperties().plots[0].series[0].name.kind, f.ChartSourceKind.None);
  });
});

test("a chart title text of '' and a missing text are one state", async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    fillGrid(ws);
    const invalid = isStatus(f.StatusCode.InvalidArgument);
    const chart = { type: f.ChartPlotType.Line, range: 'Sheet1!$A$1:$D$5', anchor: box };
    assert.throws(() => ws.addChart({ ...chart, title: '' }), invalid);
    assert.throws(
      () => ws.addChart({ ...chart, title: { kind: f.ChartTitleKind.Text, text: null, formula: null } }),
      invalid,
    );
    assert.equal(ws.listDrawings().length, 0);
  });
});
