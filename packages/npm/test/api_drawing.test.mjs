// The common drawing surface on the public entry: one nested anchor, placement input, name / description / hidden on
// every kind, one remover and one enumerator, slicer columns by name, and chart series / axis keys.
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
const { AnchorKind, AnchorEditAs, ChartAxisKind, ChartPlotType, DrawingObjectKind, StatusCode } = f;

const PNG_1X1 = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='),
  (c) => c.charCodeAt(0),
);

/** Runs `body(ws, wb)` on a default workbook holding a small grid and a two-column table, then disposes it. */
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
const corners = (a) => [a.fromRow, a.fromCol, a.toRow, a.toCol];
const chartOf = (ws, extra = {}) =>
  ws.addChart({ type: ChartPlotType.Bar, range: 'Sheet1!$A$1:$B$3', legend: null, ...extra });

/** One drawing object of every kind on `ws`, keyed by kind name. */
function everyKind(ws) {
  const table = ws.addTable({ range: 'A1:B3', name: 'Tbl' });
  return {
    picture: ws.addPicture(PNG_1X1),
    chart: chartOf(ws),
    shape: ws.addShape({ anchor: box }),
    slicer: ws.addSlicer(table, 0, { anchor: { fromRow: 8, fromCol: 3, toRow: 16, toCol: 5 } }),
  };
}

describe('the anchor is one nested record', () => {
  it('a shape added with a box reads back a two-cell anchor with those cells', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ anchor: { fromRow: 1, fromCol: 1, toRow: 4, toCol: 4 } });
      const { anchor } = shape.getDrawingProperties();
      assert.equal(anchor.kind, AnchorKind.TwoCell);
      assert.deepEqual(corners(anchor), [1, 1, 4, 4]);
    });
  });

  it('a picture added at a cell reads back a one-cell anchor', () => {
    withSheet((ws) => {
      const picture = ws.addPicture(PNG_1X1, { anchor: { fromRow: 2, fromCol: 3 } });
      const { anchor } = picture.getDrawingProperties();
      assert.equal(anchor.kind, AnchorKind.OneCell);
      assert.equal(anchor.fromRow, 2);
      assert.equal(anchor.fromCol, 3);
      assert.ok(anchor.widthEmu > 0 && anchor.heightEmu > 0);
    });
  });

  it('a flat anchor key on a shape, a slicer, a picture or a chart placement is a TypeError', () => {
    withSheet((ws) => {
      const table = ws.addTable({ range: 'A1:B3', name: 'Tbl' });
      expectTypeError(() => ws.addShape({ ...box }));
      expectTypeError(() => ws.addShape({ anchorKind: AnchorKind.TwoCell, anchor: box }));
      expectTypeError(() => ws.addSlicer(table, 0, { ...box }));
      expectTypeError(() => ws.addPicture(PNG_1X1, { fromRow: 1 }));
      expectTypeError(() => ws.addPicture(PNG_1X1, { anchor: { row: 1, col: 1 } }));
      expectTypeError(() => chartOf(ws, { placement: { anchor: { row: 5, col: 1 } } }));
      assert.equal(ws.listDrawings().length, 0);
    });
  });

  it('an anchor patch keeps every key it omits and an omitted kind keeps the object kind', () => {
    withSheet((ws) => {
      const chart = chartOf(ws, { placement: { anchor: { fromRow: 2, fromCol: 2, toRow: 12, toCol: 8 } } });
      const before = chart.getDrawingProperties().anchor;
      assert.equal(before.kind, AnchorKind.TwoCell);
      chart.update({ anchor: { fromRow: 3 } });
      const after = chart.getDrawingProperties().anchor;
      assert.equal(after.kind, AnchorKind.TwoCell);
      assert.deepEqual(corners(after), [3, 2, 12, 8]);
      assert.equal(after.editAs, before.editAs);

      const shape = ws.addShape({ anchor: box });
      shape.update({ anchor: { toCol: 6 } });
      assert.deepEqual(corners(shape.getDrawingProperties().anchor), [1, 1, 5, 6]);
    });
  });

  it('a given kind converts a chart anchor and a key that does not fit the resulting kind is InvalidArgument', () => {
    withSheet((ws) => {
      const chart = chartOf(ws, { placement: { anchor: { fromRow: 2, fromCol: 2, toRow: 12, toCol: 8 } } });
      chart.update({ anchor: { kind: AnchorKind.OneCell, widthEmu: 914400, heightEmu: 457200 } });
      let anchor = chart.getDrawingProperties().anchor;
      assert.equal(anchor.kind, AnchorKind.OneCell);
      assert.deepEqual([anchor.fromRow, anchor.fromCol, anchor.widthEmu, anchor.heightEmu], [2, 2, 914400, 457200]);
      expectCode(() => chart.update({ anchor: { toRow: 9 } }), StatusCode.InvalidArgument);
      chart.update({ anchor: { kind: AnchorKind.TwoCell, toRow: 9, toCol: 9 } });
      anchor = chart.getDrawingProperties().anchor;
      assert.equal(anchor.kind, AnchorKind.TwoCell);
      assert.deepEqual(corners(anchor), [2, 2, 9, 9]);
      expectCode(() => chart.update({ anchor: { widthEmu: 100 } }), StatusCode.InvalidArgument);
      expectTypeError(() => chart.update({ anchor: { fromRow: null } }));
    });
  });

  it('a given kind converts a shape anchor and a key that does not fit the resulting kind is InvalidArgument', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ anchor: box });
      shape.update({ anchor: { kind: AnchorKind.OneCell, widthEmu: 914400, heightEmu: 457200 } });
      let anchor = shape.getDrawingProperties().anchor;
      assert.equal(anchor.kind, AnchorKind.OneCell);
      assert.equal(anchor.widthEmu, 914400);
      assert.equal(anchor.heightEmu, 457200);
      assert.equal(anchor.fromRow, 1);
      expectCode(() => shape.update({ anchor: { toRow: 9 } }), StatusCode.InvalidArgument);
      shape.update({ anchor: { kind: AnchorKind.TwoCell, toRow: 9, toCol: 9 } });
      anchor = shape.getDrawingProperties().anchor;
      assert.equal(anchor.kind, AnchorKind.TwoCell);
      assert.deepEqual(corners(anchor), [1, 1, 9, 9]);
      expectCode(() => shape.update({ anchor: { widthEmu: 100 } }), StatusCode.InvalidArgument);
      expectTypeError(() => shape.update({ anchor: { fromRow: null } }));
    });
  });

  it('an editAs of Default means the anchor kind and a read never reports Default', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ anchor: { ...box, editAs: AnchorEditAs.Default } });
      const { anchor } = shape.getDrawingProperties();
      assert.notEqual(anchor.editAs, AnchorEditAs.Default);
      assert.equal(anchor.editAs, AnchorEditAs.TwoCell);
    });
  });
});

describe('the default anchor kind comes from the operation', () => {
  it('a picture is one-cell and a chart, a shape and a slicer are two-cell when no kind is given', () => {
    withSheet((ws) => {
      const kinds = everyKind(ws);
      const kindOf = (h) => h.getDrawingProperties().anchor.kind;
      assert.equal(kindOf(kinds.picture), AnchorKind.OneCell);
      assert.equal(kindOf(kinds.chart), AnchorKind.TwoCell);
      assert.equal(kindOf(kinds.shape), AnchorKind.TwoCell);
      assert.equal(kindOf(kinds.slicer), AnchorKind.TwoCell);
    });
  });

  it('a chart placed by one cell reads two-cell, and a kind-less anchor with a box needs no kind', () => {
    withSheet((ws) => {
      const chart = chartOf(ws, { placement: { anchor: { fromRow: 5 } } });
      const { anchor } = chart.getDrawingProperties();
      assert.equal(anchor.kind, AnchorKind.TwoCell);
      assert.equal(anchor.fromRow, 5);
      assert.ok(anchor.toRow > 5);
    });
  });
});

describe('placement input is shared by pictures and charts', () => {
  it('a chart placement carries name and description and a picture carries hidden too', () => {
    withSheet((ws) => {
      const chart = chartOf(ws, { placement: { name: 'Sales', description: 'Quarterly sales' } });
      const read = chart.getDrawingProperties();
      assert.equal(read.name, 'Sales');
      assert.equal(read.description, 'Quarterly sales');
      assert.equal(read.hidden, false);
      const picture = ws.addPicture(PNG_1X1, { name: 'Logo', description: 'd', hidden: true });
      const pictureRead = picture.getDrawingProperties();
      assert.equal(pictureRead.name, 'Logo');
      assert.equal(pictureRead.description, 'd');
      assert.equal(pictureRead.hidden, true);
    });
  });

  it('a shape and a slicer take the placement keys too', () => {
    withSheet((ws) => {
      const table = ws.addTable({ range: 'A1:B3', name: 'Tbl' });
      const shape = ws.addShape({ name: 'Box', description: 'a box', hidden: true, anchor: box });
      const read = shape.getDrawingProperties();
      assert.deepEqual([read.name, read.description, read.hidden], ['Box', 'a box', true]);
      const slicer = ws.addSlicer(table, 0, {
        description: 'pick one',
        hidden: true,
        anchor: { fromRow: 8, fromCol: 3, toRow: 16, toCol: 5 },
      });
      const slicerRead = slicer.getDrawingProperties();
      assert.deepEqual([slicerRead.description, slicerRead.hidden], ['pick one', true]);
    });
  });
});

describe('every kind reads the same common record', () => {
  it('getDrawingProperties has the same keys and a full anchor for every kind', () => {
    withSheet((ws) => {
      for (const [name, handle] of Object.entries(everyKind(ws))) {
        const read = handle.getDrawingProperties();
        assert.deepEqual(
          Object.keys(read).sort(),
          ['anchor', 'description', 'hidden', 'kind', 'name', 'objectId'],
          name,
        );
        assert.deepEqual(
          Object.keys(read.anchor).sort(),
          [
            'editAs',
            'fromCol',
            'fromColOffsetEmu',
            'fromRow',
            'fromRowOffsetEmu',
            'heightEmu',
            'kind',
            'toCol',
            'toColOffsetEmu',
            'toRow',
            'toRowOffsetEmu',
            'widthEmu',
          ],
          name,
        );
        assert.equal(read.objectId, handle.objectId, name);
      }
    });
  });

  it('a slicer reports editAs, widthEmu and heightEmu like every other kind', () => {
    withSheet((ws) => {
      const { slicer } = everyKind(ws);
      const { anchor } = slicer.getDrawingProperties();
      assert.ok(Object.values(AnchorEditAs).includes(anchor.editAs));
      assert.ok(anchor.widthEmu > 0);
      assert.ok(anchor.heightEmu > 0);
    });
  });

  it('the kind-specific records carry no anchor, name or description', () => {
    withSheet((ws) => {
      const { shape, slicer, chart, picture } = everyKind(ws);
      for (const [name, props] of [
        ['shape', shape.getProperties()],
        ['slicer', slicer.getProperties()],
        ['chart', chart.getProperties()],
        ['picture', picture.getProperties()],
      ]) {
        for (const key of ['fromRow', 'fromCol', 'toRow', 'toCol', 'anchor', 'anchorKind', 'name', 'description']) {
          assert.equal(key in props, false, `${name}.${key}`);
        }
      }
      assert.equal('format' in chart.getDrawingProperties(), false);
      assert.equal('format' in picture.getProperties(), true);
    });
  });
});

describe('name, description and hidden are read and set the same way on every kind', () => {
  it('update takes description and hidden on a picture, a chart, a shape and a slicer', () => {
    withSheet((ws) => {
      for (const [name, handle] of Object.entries(everyKind(ws))) {
        handle.update({ description: 'x', hidden: true });
        let read = handle.getDrawingProperties();
        assert.equal(read.description, 'x', name);
        assert.equal(read.hidden, true, name);
        handle.update({ hidden: false });
        read = handle.getDrawingProperties();
        assert.equal(read.hidden, false, name);
        assert.equal(read.description, 'x', name);
        handle.update({ description: null });
        assert.equal(handle.getDrawingProperties().description, null, name);
      }
    });
  });

  it('update renames a picture, a chart and a shape, and the handle follows', () => {
    withSheet((ws) => {
      const { picture, chart, shape } = everyKind(ws);
      for (const [name, handle] of Object.entries({ picture, chart, shape })) {
        handle.update({ name: `Renamed ${name}` });
        assert.equal(handle.getDrawingProperties().name, `Renamed ${name}`);
        assert.equal(handle.name, `Renamed ${name}`);
      }
    });
  });

  it('a slicer name cannot be set, and nothing else in the patch is applied', () => {
    withSheet((ws) => {
      const { slicer } = everyKind(ws);
      const before = slicer.name;
      expectCode(() => slicer.update({ name: 'Other' }), StatusCode.PreconditionFailed);
      assert.equal(slicer.name, before);
      assert.equal(slicer.getDrawingProperties().name, before);
    });
  });

  it('an empty name and a null on a key other than description are refused', () => {
    withSheet((ws) => {
      const { shape } = everyKind(ws);
      const before = plainOf(shape.getDrawingProperties());
      expectCode(() => shape.update({ name: '' }), StatusCode.InvalidArgument);
      for (const key of ['name', 'hidden', 'anchor']) expectTypeError(() => shape.update({ [key]: null }));
      expectTypeError(() => shape.update({ bogus: 1 }));
      assert.deepEqual(plainOf(shape.getDrawingProperties()), before);
    });
  });

  it('a patch made of a read record ignores objectId and kind', () => {
    withSheet((ws) => {
      const { picture } = everyKind(ws);
      assert.doesNotThrow(() => picture.update({ objectId: 99999, kind: DrawingObjectKind.Chart, hidden: true }));
      const read = picture.getDrawingProperties();
      assert.equal(read.hidden, true);
      assert.equal(read.kind, DrawingObjectKind.Picture);
      assert.notEqual(read.objectId, 99999);
    });
  });

  it('the common keys on Shape.update and Slicer.update apply with the kind keys', () => {
    withSheet((ws) => {
      const { shape, slicer } = everyKind(ws);
      shape.update({ name: 'Both', description: 'd', text: 'hello', anchor: { toCol: 7 } });
      assert.equal(shape.name, 'Both');
      assert.equal(shape.getProperties().text, 'hello');
      assert.equal(shape.getDrawingProperties().anchor.toCol, 7);
      slicer.update({ description: 'd', caption: 'Pick one', hidden: true });
      assert.equal(slicer.getProperties().caption, 'Pick one');
      assert.equal(slicer.getDrawingProperties().hidden, true);
    });
  });
});

const plainOf = (value) => JSON.parse(JSON.stringify(value));

describe('a drawing handle carries its name', () => {
  it('name equals the read name, follows an update and fails once the object is removed', () => {
    withSheet((ws) => {
      const shape = ws.addShape({ name: 'First', anchor: box });
      assert.equal(shape.name, 'First');
      assert.equal(shape.name, shape.getDrawingProperties().name);
      shape.update({ name: 'Second' });
      assert.equal(shape.name, 'Second');
      assert.equal(ws.drawing(shape.objectId).name, 'Second');
      ws.removeDrawing(shape);
      expectCode(() => shape.name, StatusCode.NotFound);
    });
  });
});

describe('one remover and one enumerator', () => {
  it('removeDrawing removes a shape, a slicer, a picture and a chart by id', () => {
    withSheet((ws) => {
      const handles = Object.values(everyKind(ws));
      assert.equal(ws.listDrawings().length, 4);
      for (const handle of handles) {
        const id = handle.objectId;
        ws.removeDrawing(id);
        assert.equal(
          ws.listDrawings().some((d) => d.objectId === id),
          false,
        );
        expectCode(() => ws.drawing(id), StatusCode.NotFound);
      }
      assert.equal(ws.listShapes().length, 0);
      assert.equal(ws.listSlicers().length, 0);
    });
  });

  it('removeDrawing on an unknown id is NotFound and changes nothing', () => {
    withSheet((ws) => {
      const { shape } = everyKind(ws);
      expectCode(() => ws.removeDrawing(shape.objectId + 1000), StatusCode.NotFound);
      assert.equal(ws.listDrawings().length, 4);
    });
  });

  it('the per-kind lists are the drawing list filtered by kind', () => {
    withSheet((ws) => {
      const { picture, chart, shape, slicer } = everyKind(ws);
      const second = ws.addShape({ anchor: { fromRow: 20, fromCol: 1, toRow: 24, toCol: 3 } });
      assert.deepEqual(ws.listPictures(), [picture]);
      assert.deepEqual(ws.listCharts(), [chart]);
      assert.deepEqual(ws.listShapes(), [shape, second]);
      assert.deepEqual(ws.listSlicers(), [slicer]);
      assert.deepEqual(
        ws.listDrawings().map((d) => d.objectId),
        [picture, chart, shape, slicer, second].map((h) => h.objectId),
      );
    });
  });
});

describe('a slicer column is a position or a name', () => {
  it('a name resolves like a position and ignores case', () => {
    withSheet((ws) => {
      const table = ws.addTable({ range: 'A1:B3', name: 'Tbl' });
      const spec = { anchor: { fromRow: 8, fromCol: 3, toRow: 16, toCol: 5 } };
      const byNumber = ws.addSlicer(table, 1, spec);
      const byName = ws.addSlicer('Tbl', 'Amount', spec);
      const folded = ws.addSlicer(table, 'amount', spec);
      for (const slicer of [byNumber, byName, folded]) {
        assert.equal(slicer.getProperties().columnName, 'Amount');
        assert.equal(slicer.getProperties().tableName, 'Tbl');
      }
      assert.equal(ws.addSlicer(table, 'Region', spec).getProperties().columnName, 'Region');
    });
  });

  it('a name the table lacks is NotFound and a negative position is IndexOutOfRange', () => {
    withSheet((ws) => {
      const table = ws.addTable({ range: 'A1:B3', name: 'Tbl' });
      const spec = { anchor: { fromRow: 8, fromCol: 3, toRow: 16, toCol: 5 } };
      expectCode(() => ws.addSlicer(table, 'nope', spec), StatusCode.NotFound);
      expectCode(() => ws.addSlicer(table, -1, spec), StatusCode.IndexOutOfRange);
      assert.throws(
        () => ws.addSlicer(table, 5, spec),
        (e) => e.name === 'FormulonError',
      );
      assert.equal(ws.listSlicers().length, 0);
    });
  });
});

describe('chart series are keyed by seriesId and axes by kind', () => {
  it('addSeries returns the id that getProperties reads back as seriesId', () => {
    withSheet((ws) => {
      const chart = chartOf(ws);
      const id = chart.addSeries(0, {
        values: 'Sheet1!$B$2:$B$3',
        name: 'Sheet1!$B$1',
        categories: 'Sheet1!$A$2:$A$3',
      });
      const series = chart.getProperties().plots[0].series;
      assert.ok(series.some((s) => s.seriesId === id));
      assert.equal(
        series.some((s) => 'index' in s),
        false,
      );
      assert.ok(chart.getData().series.some((s) => s.seriesId === id));
      chart.updateSeries(0, id, { fill: { color: '#112233' } });
      chart.removeSeries(0, id);
      assert.equal(
        chart.getProperties().plots[0].series.some((s) => s.seriesId === id),
        false,
      );
    });
  });

  it('a series id the plot lacks is NotFound for update, point edits and removal', () => {
    withSheet((ws) => {
      const chart = chartOf(ws);
      expectCode(() => chart.updateSeries(0, 9999, {}), StatusCode.NotFound);
      expectCode(() => chart.removeSeries(0, 9999), StatusCode.NotFound);
      expectCode(() => chart.setPoint(0, 9999, 0, null), StatusCode.NotFound);
      expectCode(() => chart.updatePoint(0, 9999, 0, {}), StatusCode.NotFound);
    });
  });

  it('updateAxis takes an axis kind and reads back on the axis of that kind', () => {
    withSheet((ws) => {
      const chart = chartOf(ws);
      chart.updateAxis({ kind: ChartAxisKind.Value }, { min: 0, max: 50 });
      const axes = chart.getProperties().axes;
      const value = axes.find((a) => a.kind === ChartAxisKind.Value);
      assert.equal(value.min, 0);
      assert.equal(value.max, 50);
      const category = axes.find((a) => a.kind === ChartAxisKind.Category);
      assert.equal(category.min, null);
      chart.updateAxis({ kind: ChartAxisKind.Value, plot: 0 }, { max: 80 });
      assert.equal(chart.getProperties().axes.find((a) => a.kind === ChartAxisKind.Value).max, 80);
    });
  });

  it('updateAxis still takes an axis id, and an unknown id or kind is NotFound and a plot past the end IndexOutOfRange', () => {
    withSheet((ws) => {
      const chart = chartOf(ws);
      const value = chart.getProperties().axes.find((a) => a.kind === ChartAxisKind.Value);
      chart.updateAxis(value.id, { min: 5 });
      assert.equal(chart.getProperties().axes.find((a) => a.kind === ChartAxisKind.Value).min, 5);
      expectCode(() => chart.updateAxis({ kind: ChartAxisKind.Series }, { min: 1 }), StatusCode.NotFound);
      expectCode(() => chart.updateAxis(987654, { min: 1 }), StatusCode.NotFound);
      expectCode(
        () => chart.updateAxis({ kind: ChartAxisKind.Value, plot: 3 }, { min: 1 }),
        StatusCode.IndexOutOfRange,
      );
    });
  });

  it('the resolved data names its parts and carries the number format on the categories', () => {
    withSheet((ws) => {
      const chart = chartOf(ws);
      const data = chart.getData();
      assert.ok('numberFormat' in data.categories);
      assert.equal('formatCode' in data.categories, false);
      assert.ok(Array.isArray(data.axisTitleTexts));
      assert.ok(Array.isArray(data.unresolved));
      const [first] = data.series;
      assert.ok('ranges' in first);
      assert.ok(Array.isArray(first.ranges.values));
      assert.equal('index' in first, false);
    });
  });
});
