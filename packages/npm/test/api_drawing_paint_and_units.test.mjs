// Drawing paint and units on the public entry: atomic updates, shape paint records, preset geometry, point-based
// slicer and plot units, a flat chart input and the role of an unresolved chart input.
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
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
      assert.equal(fill.color.resolved.argb >>> 0, 0xffff0000);
      assert.equal(fill.color.resolvedTransparencyPct, 25);
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

describe('a paint colour carries its transforms', () => {
  const shapeWith = (ws, color, extra = {}) => ws.addShape({ anchor: box, fill: { color, ...extra } });

  it('reads the base, the transforms in order and the rendered colour', () => {
    withSheet((ws) => {
      const shape = shapeWith(ws, {
        base: '#336699',
        transforms: [{ name: 'lumMod', value: 75000 }, { name: 'gray' }],
      });
      const { color } = shape.getProperties().fill;
      assert.equal(color.base.kind, ColorKind.Rgb);
      assert.equal(color.base.argb >>> 0, 0xff336699);
      assert.deepEqual(color.transforms, [
        { name: 'lumMod', value: 75000 },
        { name: 'gray', value: null },
      ]);
      assert.equal(typeof color.resolved.argb, 'number');
      assert.equal(color.resolved.resolution, f.ColorResolution.Exact);
      assert.equal(color.resolvedTransparencyPct, 0);
    });
  });

  it('a plain colour reads with no transforms', () => {
    withSheet((ws) => {
      const { color } = shapeWith(ws, '#FF0000').getProperties().fill;
      assert.deepEqual(color.transforms, []);
      assert.equal(color.resolved.argb >>> 0, 0xffff0000);
    });
  });

  it('an alpha transform reads as the rendered transparency, and a read record passes back as input', () => {
    withSheet((ws) => {
      const shape = shapeWith(ws, { base: '#FF0000', transforms: [{ name: 'alpha', value: 60000 }] });
      const { fill } = shape.getProperties();
      assert.equal(fill.color.resolvedTransparencyPct, 40);
      const copy = ws.addShape({ anchor: box });
      copy.update({ fill });
      assert.deepEqual(copy.getProperties().fill.color, fill.color);
    });
  });

  it('an alpha transform beside transparencyPct must state the same opacity', () => {
    withSheet((ws) => {
      const color = { base: '#FF0000', transforms: [{ name: 'alpha', value: 60000 }] };
      expectCode(() => shapeWith(ws, color, { transparencyPct: 10 }), StatusCode.InvalidArgument);
      shapeWith(ws, color, { transparencyPct: 40 });
      assert.equal(ws.listDrawings().length, 1);
    });
  });

  it('refuses a malformed transform name, a tint beside transforms and more than 64 transforms', () => {
    withSheet((ws) => {
      const withTransforms = (transforms, base = '#FF0000') => shapeWith(ws, { base, transforms });
      expectCode(() => withTransforms([{ name: '1bad', value: 1 }]), StatusCode.InvalidArgument);
      const themed = { kind: ColorKind.Theme, argb: 0, theme: 4, tint: 0.4, indexed: 0 };
      expectCode(() => withTransforms([{ name: 'alpha', value: 1 }], themed), StatusCode.InvalidArgument);
      const many = Array.from({ length: 65 }, () => ({ name: 'lumMod', value: 1 }));
      expectCode(() => withTransforms(many), StatusCode.SecurityResourceLimit);
      expectTypeError(() => withTransforms([{ value: 1 }]));
      expectTypeError(() => withTransforms([{ name: 'alpha', value: 'half' }]));
      assert.equal(ws.listDrawings().length, 0);
    });
  });

  it('a set colour with no base is refused; null removes one', () => {
    withSheet((ws) => {
      const none = { kind: ColorKind.None, argb: 0, theme: 0, tint: 0, indexed: 0 };
      const shape = ws.addShape({ anchor: box, fill: { color: '#FF0000' }, line: { color: '#FF0000' } });
      expectCode(() => shape.update({ fill: { color: none } }), StatusCode.InvalidArgument);
      expectCode(() => shape.update({ line: { color: none } }), StatusCode.InvalidArgument);
      const { fill, line } = shape.getProperties();
      assert.equal(fill.color.base.argb >>> 0, 0xffff0000);
      assert.equal(line.color.base.argb >>> 0, 0xffff0000);
      const chart = chartOf(ws);
      chart.updateSeries(0, 0, { labels: { showValue: true, fontColor: '#FF0000' } });
      expectCode(() => chart.updateSeries(0, 0, { labels: { fontColor: none } }), StatusCode.InvalidArgument);
      assert.equal(chart.getProperties().plots[0].series[0].labels.fontColor.argb >>> 0, 0xffff0000);
    });
  });

  it('a chart series takes the same colour', () => {
    withSheet((ws) => {
      const chart = chartOf(ws);
      const color = { base: '#336699', transforms: [{ name: 'lumMod', value: 50000 }] };
      chart.updateSeries(0, 0, { fill: { color } });
      const read = chart.getProperties().plots[0].series[0].fill.color;
      assert.equal(read.base.argb >>> 0, 0xff336699);
      assert.deepEqual(read.transforms, [{ name: 'lumMod', value: 50000 }]);
    });
  });
});

describe('a line carries its end decorations and the system dash presets', () => {
  const lineOf = (ws, line) => ws.addShape({ anchor: box, presetGeometry: 'line', line });

  it('headEnd and tailEnd round trip, and null removes one end', () => {
    withSheet((ws) => {
      const shape = lineOf(ws, {
        color: '#000000',
        headEnd: { type: f.ChartLineEndType.Triangle, width: f.ChartLineEndSize.Large },
        tailEnd: { type: f.ChartLineEndType.Arrow, width: null, length: f.ChartLineEndSize.Small },
      });
      let { line } = shape.getProperties();
      assert.deepEqual(line.headEnd, {
        type: f.ChartLineEndType.Triangle,
        width: f.ChartLineEndSize.Large,
        length: null,
      });
      assert.deepEqual(line.tailEnd, {
        type: f.ChartLineEndType.Arrow,
        width: null,
        length: f.ChartLineEndSize.Small,
      });
      shape.update({ line: { headEnd: null } });
      line = shape.getProperties().line;
      assert.equal(line.headEnd, null);
      assert.equal(line.tailEnd.type, f.ChartLineEndType.Arrow);
      // An end replaces the stored one whole: the size it leaves out is dropped.
      shape.update({ line: { tailEnd: { type: f.ChartLineEndType.Oval } } });
      assert.deepEqual(shape.getProperties().line.tailEnd, {
        type: f.ChartLineEndType.Oval,
        width: null,
        length: null,
      });
    });
  });

  it('a line end needs a type in range', () => {
    withSheet((ws) => {
      expectTypeError(() => lineOf(ws, { headEnd: {} }));
      expectTypeError(() => lineOf(ws, { headEnd: { width: f.ChartLineEndSize.Small } }));
      expectCode(() => lineOf(ws, { headEnd: { type: 6 } }), StatusCode.InvalidArgument);
      const wide = { type: f.ChartLineEndType.Arrow, width: 3 };
      expectCode(() => lineOf(ws, { tailEnd: wide }), StatusCode.InvalidArgument);
    });
  });

  it('the system and long dash presets read as their own members', () => {
    withSheet((ws) => {
      for (const name of ['SysDot', 'SysDash', 'LgDash', 'LgDashDot', 'LgDashDotDot', 'SysDashDot', 'SysDashDotDot']) {
        const shape = lineOf(ws, { dash: f.ChartLineDash[name] });
        assert.equal(shape.getProperties().line.dash, f.ChartLineDash[name], name);
      }
      expectCode(() => lineOf(ws, { dash: 12 }), StatusCode.InvalidArgument);
    });
  });
});

describe('the chart area and the plot area carry their own paint', () => {
  it('both read as records, and a display patch merges into them', () => {
    withSheet((ws) => {
      const chart = chartOf(ws);
      for (const key of ['chartArea', 'plotArea']) {
        const area = chart.getProperties()[key];
        assert.deepEqual(Object.keys(area), ['fill', 'line'], key);
      }
      chart.updateDisplay({
        chartArea: { fill: { color: '#FF0000' }, line: { widthPt: 2 } },
        plotArea: { fill: { color: '#00FF00' } },
      });
      const props = chart.getProperties();
      assert.equal(props.chartArea.fill.color.base.argb >>> 0, 0xffff0000);
      assert.equal(props.chartArea.line.widthPt, 2);
      assert.equal(props.plotArea.fill.color.base.argb >>> 0, 0xff00ff00);
      // An omitted key keeps the paint; the other area is untouched.
      chart.updateDisplay({ chartArea: { line: { widthPt: 3 } } });
      const after = chart.getProperties();
      assert.equal(after.chartArea.fill.color.base.argb >>> 0, 0xffff0000);
      assert.equal(after.chartArea.line.widthPt, 3);
      assert.equal(after.plotArea.fill.color.base.argb >>> 0, 0xff00ff00);
    });
  });

  it('null on an area key is a TypeError, and null on a member removes its paint', () => {
    withSheet((ws) => {
      const chart = chartOf(ws);
      expectTypeError(() => chart.updateDisplay({ chartArea: null }));
      expectTypeError(() => chart.updateDisplay({ plotArea: null }));
      chart.updateDisplay({ chartArea: { fill: { color: '#FF0000' } } });
      chart.updateDisplay({ chartArea: { fill: null } });
      assert.equal(chart.getProperties().chartArea.fill, null);
    });
  });
});
