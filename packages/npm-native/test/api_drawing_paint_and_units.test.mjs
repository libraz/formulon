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
      assert.equal(fill.color.base.kind, ColorKind.Rgb);
      assert.equal(fill.color.base.argb & 0xffffff, 0xff0000);
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
      assert.equal(painted.fill.color.base.argb & 0xffffff, 0x336699);
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

describe('a drawing colour carries its transforms and the colour Excel renders', () => {
  const base = { kind: ColorKind.Rgb, argb: 0xff336699 };
  const fillOf = (shape) => shape.getProperties().fill;

  it('reads base, transforms and resolved colour back', () => {
    withSheet((ws) => {
      const shape = ws.addShape({
        anchor: box,
        fill: {
          color: {
            base,
            transforms: [
              { name: 'lumMod', value: 50000 },
              { name: 'alpha', value: 40000 },
            ],
          },
        },
      });
      const { color } = fillOf(shape);
      assert.equal(color.base.argb >>> 0, 0xff336699);
      assert.deepEqual(color.transforms, [
        { name: 'lumMod', value: 50000 },
        { name: 'alpha', value: 40000 },
      ]);
      assert.equal(color.resolved.resolution, f.ColorResolution.Exact);
      // An alpha transform leaves the rendered opacity unknown.
      assert.equal(color.resolvedTransparencyPct, null);
    });
  });

  it('a plain colour has no transforms and an opaque resolved colour', () => {
    withSheet((ws) => {
      const { color } = fillOf(ws.addShape({ anchor: box, fill: { color: base } }));
      assert.deepEqual(color.transforms, []);
      assert.equal(color.resolved.argb >>> 0, 0xff336699);
      assert.equal(color.resolvedTransparencyPct, 0);
    });
  });

  it('a transform without val reads value null', () => {
    withSheet((ws) => {
      const { color } = fillOf(ws.addShape({ anchor: box, fill: { color: { base, transforms: [{ name: 'comp' }] } } }));
      assert.deepEqual(color.transforms, [{ name: 'comp', value: null }]);
    });
  });

  it('accepts the colour a read returned and ignores its resolved members', () => {
    withSheet((ws) => {
      const shape = ws.addShape({
        anchor: box,
        fill: { color: { base, transforms: [{ name: 'lumMod', value: 75000 }] } },
      });
      const read = fillOf(shape).color;
      const other = ws.addShape({ anchor: box, fill: { color: { ...read, resolvedTransparencyPct: 99 } } });
      assert.deepEqual(fillOf(other).color.transforms, read.transforms);
      other.update({ line: { color: read } });
      assert.deepEqual(other.getProperties().line.color.transforms, read.transforms);
    });
  });

  it('a patch replaces the colour whole and null clears it', () => {
    withSheet((ws) => {
      const shape = ws.addShape({
        anchor: box,
        fill: { color: { base, transforms: [{ name: 'lumMod', value: 50000 }] } },
      });
      shape.update({ fill: { color: { base, transforms: [{ name: 'lumOff', value: 20000 }] } } });
      assert.deepEqual(fillOf(shape).color.transforms, [{ name: 'lumOff', value: 20000 }]);
      shape.update({ fill: { color: null } });
      assert.equal(fillOf(shape), null);
    });
  });

  it('refuses a bad transform name, too many transforms and a mismatched opacity', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ anchor: box });
      const patch =
        (transforms, extra = {}) =>
        () =>
          shape.update({ fill: { color: { base, transforms }, ...extra } });
      expectCode(patch([{ name: '1bad', value: 1 }]), StatusCode.InvalidArgument);
      expectCode(patch([{ name: '', value: 1 }]), StatusCode.InvalidArgument);
      const many = Array.from({ length: 65 }, () => ({ name: 'lumMod', value: 100000 }));
      expectCode(patch(many), StatusCode.SecurityResourceLimit);
      expectCode(patch([{ name: 'alpha', value: 40000 }], { transparencyPct: 10 }), StatusCode.InvalidArgument);
      patch([{ name: 'alpha', value: 40000 }], { transparencyPct: 60 })();
      assert.equal(fillOf(shape).transparencyPct, 60);
    });
  });

  it('a tint beside transforms is refused', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ anchor: box });
      const theme = { kind: ColorKind.Theme, theme: 4, tint: 0.5 };
      expectCode(
        () => shape.update({ fill: { color: { base: theme, transforms: [{ name: 'lumMod', value: 1 }] } } }),
        StatusCode.InvalidArgument,
      );
    });
  });
});

describe('a line carries its end decorations and the system dashes', () => {
  const lineOf = (shape) => shape.getProperties().line;

  it('reads headEnd and tailEnd, a stated size or null', () => {
    withSheet((ws) => {
      const shape = ws.addShape({
        anchor: box,
        line: {
          color: '#00FF00',
          headEnd: {
            type: f.ChartLineEndType.Stealth,
            width: f.ChartLineEndSize.Small,
            length: f.ChartLineEndSize.Large,
          },
          tailEnd: { type: f.ChartLineEndType.Arrow },
        },
      });
      assert.deepEqual(lineOf(shape).headEnd, { type: 2, width: 1, length: 2 });
      assert.deepEqual(lineOf(shape).tailEnd, { type: 5, width: null, length: null });
    });
  });

  it('a shape line without ends reads null for both', () => {
    withSheet((ws) => {
      const line = lineOf(ws.addShape({ anchor: box, line: { color: '#00FF00' } }));
      assert.equal(line.headEnd, null);
      assert.equal(line.tailEnd, null);
    });
  });

  it('a patch replaces an end whole and null removes it', () => {
    withSheet((ws) => {
      const shape = ws.addShape({
        anchor: box,
        line: { color: '#00FF00', headEnd: { type: 1, width: 0, length: 0 }, tailEnd: { type: 4 } },
      });
      shape.update({ line: { headEnd: { type: 3 }, tailEnd: null } });
      assert.deepEqual(lineOf(shape).headEnd, { type: 3, width: null, length: null });
      assert.equal(lineOf(shape).tailEnd, null);
      shape.update({ line: null });
      assert.equal(lineOf(shape), null);
    });
  });

  it('refuses an end type outside the table', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ anchor: box, line: { color: '#00FF00' } });
      expectCode(() => shape.update({ line: { headEnd: { type: 6 } } }), StatusCode.InvalidArgument);
      expectCode(() => shape.update({ line: { tailEnd: { type: 1, width: 3 } } }), StatusCode.InvalidArgument);
      expectTypeError(() => shape.update({ line: { headEnd: {} } }));
    });
  });

  it('the system dashes round trip', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ anchor: box, line: { color: '#00FF00' } });
      for (let dash = 5; dash <= 11; dash++) {
        shape.update({ line: { dash } });
        assert.equal(lineOf(shape).dash, dash);
      }
      expectCode(() => shape.update({ line: { dash: 12 } }), StatusCode.InvalidArgument);
    });
  });
});

describe('chart and plot area paint', () => {
  it('reads null members for a chart that states none', () => {
    withSheet((ws) => {
      const { chartArea, plotArea } = chartOf(ws).getProperties();
      for (const area of [chartArea, plotArea]) {
        assert.deepEqual(Object.keys(area).sort(), ['fill', 'line']);
      }
    });
  });

  it('updateDisplay sets and clears the area paint, and keeps what it does not name', () => {
    withSheet((ws) => {
      const chart = chartOf(ws);
      chart.updateDisplay({
        chartArea: { fill: { color: '#112233' }, line: { color: '#445566', widthPt: 2 } },
        plotArea: {
          fill: {
            color: { base: { kind: ColorKind.Rgb, argb: 0xff778899 }, transforms: [{ name: 'alpha', value: 50000 }] },
          },
        },
      });
      let props = chart.getProperties();
      assert.equal(props.chartArea.fill.color.base.argb >>> 0, 0xff112233);
      assert.equal(props.chartArea.line.widthPt, 2);
      assert.equal(props.plotArea.fill.color.transforms[0].name, 'alpha');
      chart.updateDisplay({ chartArea: { line: null } });
      props = chart.getProperties();
      assert.equal(props.chartArea.line, null);
      assert.equal(props.chartArea.fill.color.base.argb >>> 0, 0xff112233);
      chart.updateDisplay({ chartArea: { fill: null } });
      assert.equal(chart.getProperties().chartArea.fill, null);
      chart.updateDisplay({ blanksAs: f.BlanksAs.Gap });
      assert.equal(chart.getProperties().plotArea.fill.color.transforms[0].name, 'alpha');
    });
  });

  it('the area keys themselves cannot be cleared', () => {
    withSheet((ws) => {
      const chart = chartOf(ws);
      expectTypeError(() => chart.updateDisplay({ chartArea: null }));
      expectTypeError(() => chart.updateDisplay({ plotArea: null }));
    });
  });
});

describe('a shape reads and writes its adjust values and text direction', () => {
  it('adjustValues round trips and a patch merges by name', () => {
    withSheet((ws) => {
      const shape = ws.addShape({
        anchor: box,
        presetGeometry: 'wedgeRectCallout',
        adjustValues: { adj1: 10000, adj2: 20000 },
      });
      assert.deepEqual({ ...shape.getProperties().adjustValues }, { adj1: 10000, adj2: 20000 });
      shape.update({ adjustValues: { adj1: 5 } });
      assert.deepEqual({ ...shape.getProperties().adjustValues }, { adj1: 5, adj2: 20000 });
      shape.update({ adjustValues: { adj2: null } });
      assert.deepEqual({ ...shape.getProperties().adjustValues }, { adj1: 5 });
      shape.update({ adjustValues: {} });
      assert.deepEqual({ ...shape.getProperties().adjustValues }, { adj1: 5 });
      shape.update({ presetGeometry: 'rect' });
      assert.deepEqual({ ...shape.getProperties().adjustValues }, {});
    });
  });

  it('a shape without adjust values reads an empty record', () => {
    withSheet((ws) => {
      assert.deepEqual({ ...ws.addShape({ anchor: box }).getProperties().adjustValues }, {});
    });
  });

  it('refuses null on input, an empty name and too many entries', () => {
    withSheet((ws) => {
      expectTypeError(() => ws.addShape({ anchor: box, adjustValues: null }));
      expectTypeError(() =>
        ws.addShape({ anchor: box, presetGeometry: 'wedgeRectCallout', adjustValues: { adj1: null } }),
      );
      expectTypeError(() => ws.addShape({ anchor: box, presetGeometry: 'wedgeRectCallout', adjustValues: 3 }));
      const shape = ws.addShape({ anchor: box, presetGeometry: 'wedgeRectCallout' });
      expectTypeError(() => shape.update({ adjustValues: null }));
      expectCode(() => shape.update({ adjustValues: { '': 1 } }), StatusCode.InvalidArgument);
      const many = Object.fromEntries(Array.from({ length: 65 }, (_, i) => [`adj${i + 1}`, i]));
      expectCode(() => shape.update({ adjustValues: many }), StatusCode.SecurityResourceLimit);
    });
  });

  it('textVert reads null until stated, round trips and clears with null', () => {
    withSheet((ws) => {
      assert.equal(ws.addShape({ anchor: box }).getProperties().textVert, null);
      const shape = ws.addShape({ anchor: box, textVert: f.ShapeTextVert.Vert270 });
      assert.equal(shape.getProperties().textVert, 2);
      shape.update({ textVert: f.ShapeTextVert.MongolianVert });
      assert.equal(shape.getProperties().textVert, 6);
      shape.update({ textVert: null });
      assert.equal(shape.getProperties().textVert, null);
      expectTypeError(() => ws.addShape({ anchor: box, textVert: null }));
    });
  });

  it('a shape that is no connector reads null connections', () => {
    withSheet((ws) => {
      const props = ws.addShape({ anchor: box }).getProperties();
      assert.equal(props.startConnection, null);
      assert.equal(props.endConnection, null);
    });
  });
});
