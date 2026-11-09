// Drawing paint and units on the public entry: atomic updates, shape paint records, preset geometry, point-based
// slicer and plot units, a flat chart input and the role of an unresolved chart input.

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
const { ChartLineDash, ChartPlotType, ColorKind, StatusCode } = f;

/** Runs `body(ws, wb)` on a default workbook holding a small grid, then disposes it. */
function withSheet(body) {
  const wb = f.Workbook.createDefault();
  try {
    const ws = wb.sheet(0);
    ws.setText('A1', 'Region');
    ws.setText('B1', 'Amount');
    ws.setText('A2', 'East');
    ws.setNumber('B2', 10);
    ws.setText('A3', 'West');
    ws.setNumber('B3', 30);
    return body(ws, wb);
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

const expectTypeError = (fn) => assert.throws(fn, (e) => e instanceof TypeError, 'expected a TypeError');
const box = { fromRow: 1, fromCol: 1, toRow: 5, toCol: 4 };
const chartOf = (ws, extra = {}) =>
  ws.addChart({ type: ChartPlotType.Bar, range: 'Sheet1!$A$1:$B$3', legend: null, ...extra });
const nameOf = (object) => object.getDrawingProperties().name;

describe('a shape update is checked as a whole', () => {
  it('Shape.update with an invalid line width keeps the new name unapplied', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ anchor: box, name: 'Original' });
      expectCode(() => shape.update({ name: 'Renamed', line: { widthPt: -1 } }), StatusCode.InvalidArgument);
      assert.equal(nameOf(shape), 'Original');
    });
  });

  it('connector update with text keeps the name', () => {
    withSheet((ws) => {
      const connector = ws.addShape({ anchor: box, name: 'Wire', presetGeometry: 'line' });
      expectCode(() => connector.update({ name: 'Renamed', text: 'hello' }), StatusCode.InvalidArgument);
      assert.equal(nameOf(connector), 'Wire');
    });
  });
});

describe('a shape preset geometry is an ECMA-376 shape name', () => {
  it('unknown presetGeometry is InvalidArgument', () => {
    withSheet((ws) => {
      expectCode(() => ws.addShape({ anchor: box, presetGeometry: 'notAShape' }), StatusCode.InvalidArgument);
      assert.equal(ws.listDrawings().length, 0);
      const shape = ws.addShape({ anchor: box, presetGeometry: 'roundRect' });
      assert.equal(shape.getProperties().presetGeometry, 'roundRect');
      expectCode(() => shape.update({ presetGeometry: 'notAShape' }), StatusCode.InvalidArgument);
      assert.equal(shape.getProperties().presetGeometry, 'roundRect');
    });
  });
});

describe('a shape paints with the chart fill and line records', () => {
  it('shape fill and line use the chart records', () => {
    withSheet((ws) => {
      const shape = ws.addShape({
        anchor: box,
        fill: { color: '#FF0000', transparencyPct: 25 },
        line: { color: '#00FF00', widthPt: 2, dash: ChartLineDash.Dash },
      });
      const { fill, line } = shape.getProperties();
      assert.equal(fill.color.kind, ColorKind.Rgb);
      assert.equal(fill.color.argb & 0xffffff, 0xff0000);
      assert.equal(fill.transparencyPct, 25);
      assert.equal(fill.none, false);
      assert.equal(fill.other, false);
      assert.equal(line.widthPt, 2);
      assert.equal(line.dash, ChartLineDash.Dash);
      assert.equal(line.other, false);
      for (const key of ['fillKind', 'fillColor', 'lineKind', 'lineColor', 'lineWidthPt']) {
        assert.equal(key in shape.getProperties(), false, key);
      }
      assert.equal(f.ShapePaint, undefined);
      shape.update({ fill: { none: true } });
      assert.equal(shape.getProperties().fill.none, true);
    });
  });

  it('fill null restores the style paint', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ anchor: box, fill: { color: '#FF0000' }, line: { color: '#00FF00', widthPt: 2 } });
      shape.update({ fill: null, line: null });
      const props = shape.getProperties();
      assert.equal(props.fill, null);
      assert.equal(props.line, null);
      assert.equal(ws.addShape({ anchor: box }).getProperties().fill, null);
    });
  });

  it('a shape paint record rejects unknown nested keys', () => {
    withSheet((ws) => {
      expectTypeError(() => ws.addShape({ anchor: box, fill: { colour: '#FF0000' } }));
      expectTypeError(() => ws.addShape({ anchor: box, line: { width: 2 } }));
    });
  });

  it('series fill null when unstated', () => {
    withSheet((ws) => {
      const chart = chartOf(ws);
      const [series] = chart.getProperties().plots[0].series;
      // A new chart states its series paint, as Excel's does; null removes it.
      chart.updateSeries(0, series.seriesId, { fill: null, line: null });
      const cleared = chart.getProperties().plots[0].series[0];
      assert.equal(cleared.fill, null);
      assert.equal(cleared.line, null);
      chart.updateSeries(0, series.seriesId, { fill: { color: '#336699' } });
      const painted = chart.getProperties().plots[0].series[0];
      assert.equal(painted.fill.color.argb & 0xffffff, 0x336699);
      assert.equal(painted.line, null);
    });
  });
});

describe('a slicer row height is in points', () => {
  it('slicer rowHeightPt round trip', () => {
    withSheet((ws) => {
      const table = ws.addTable({ range: 'A1:B3', name: 'Tbl' });
      const slicer = ws.addSlicer(table, 0, {
        anchor: { fromRow: 8, fromCol: 3, toRow: 16, toCol: 5 },
        rowHeightPt: 24,
      });
      assert.equal(slicer.getProperties().rowHeightPt, 24);
      slicer.update({ rowHeightPt: 30.5 });
      assert.equal(slicer.getProperties().rowHeightPt, 30.5);
      assert.equal('rowHeightEmu' in slicer.getProperties(), false);
      expectTypeError(() => slicer.update({ rowHeightEmu: 241300 }));
      expectCode(() => slicer.update({ rowHeightPt: -1 }), StatusCode.InvalidArgument);
    });
  });
});

describe('a chart is inserted from a flat input', () => {
  it('addChart takes anchor and name at the top level', () => {
    withSheet((ws) => {
      const chart = chartOf(ws, { name: 'Sales', anchor: { fromRow: 2, fromCol: 2, toRow: 12, toCol: 8 } });
      const drawing = chart.getDrawingProperties();
      assert.equal(drawing.name, 'Sales');
      assert.equal(drawing.anchor.fromRow, 2);
      assert.equal(drawing.anchor.toCol, 8);
    });
  });

  it('placement key is a TypeError', () => {
    withSheet((ws) => {
      expectTypeError(() => chartOf(ws, { placement: { name: 'Sales' } }));
      assert.equal(ws.listDrawings().length, 0);
    });
  });
});

describe('plot, sparkline and gradient keys carry their unit', () => {
  it('unit-suffixed plot, sparkline and gradient keys', () => {
    withSheet((ws, wb) => {
      const bar = chartOf(ws);
      bar.updatePlot(0, { overlapPct: 25 });
      const barPlot = bar.getProperties().plots[0];
      assert.equal(barPlot.overlapPct, 25);
      assert.equal('overlap' in barPlot, false);
      expectTypeError(() => bar.updatePlot(0, { overlap: 25 }));

      const ring = chartOf(ws, { type: ChartPlotType.Doughnut });
      ring.updatePlot(0, { holeSizePct: 40 });
      const ringPlot = ring.getProperties().plots[0];
      assert.equal(ringPlot.holeSizePct, 40);
      assert.equal('holeSize' in ringPlot, false);

      const id = ws.addSparklineGroup({ sparklines: [{ source: 'Sheet1!B2:B3', location: 'D1' }], lineWidthPt: 2.25 });
      assert.equal(ws.listSparklineGroups()[0].lineWidthPt, 2.25);
      ws.updateSparklineGroup(id, { lineWidthPt: 3 });
      const group = ws.listSparklineGroups()[0];
      assert.equal(group.lineWidthPt, 3);
      assert.equal('lineWeight' in group, false);
      expectTypeError(() => ws.updateSparklineGroup(id, { lineWeight: 3 }));

      const red = { kind: ColorKind.Rgb, argb: 0xffff0000, theme: 0, tint: 0, indexed: 0 };
      const index = wb.ooxml.addFill({ gradient: { type: 0, angleDeg: 45, stops: [{ position: 0, color: red }] } });
      const gradient = wb.ooxml.getFill(index).gradient;
      assert.equal(gradient.angleDeg, 45);
      assert.equal('degree' in gradient, false);
    });
  });
});

describe('chart data names the role of an unresolved input', () => {
  it('unresolved entries carry role', () => {
    withSheet((ws, wb) => {
      const other = wb.addSheet('Other');
      other.setNumber('A1', 1);
      other.setNumber('A2', 2);
      const chart = chartOf(ws);
      chart.addSeries(0, { values: 'Other!$A$1:$A$2', name: 'Sheet1!$B$1' });
      wb.removeSheet(other);
      const { unresolved } = chart.getData();
      assert.ok(unresolved.length > 0);
      for (const entry of unresolved) {
        assert.ok(Object.values(f.ChartRole).includes(entry.role));
        assert.equal('source' in entry, false);
      }
    });
  });
});
