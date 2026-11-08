import test from 'node:test';
import { inflateRawSync } from 'node:zlib';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

// ---- Drawing shapes, sparkline groups and legacy note shapes -----------------

const release = (wb) => wb.dispose();

const roundTrip = (mod, wb) => {
  const saved = wb.save();
  return mod.Workbook.load(saved);
};

// Reads one stored or deflated entry of a ZIP archive by name.
const readZipEntry = (bytes, name) => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = bytes.length - 22;
  while (view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
  let at = view.getUint32(eocd + 16, true);
  for (let n = view.getUint16(eocd + 10, true); n > 0; n -= 1) {
    const method = view.getUint16(at + 10, true);
    const size = view.getUint32(at + 20, true);
    const nameLen = view.getUint16(at + 28, true);
    const extraLen = view.getUint16(at + 30, true);
    const commentLen = view.getUint16(at + 32, true);
    const local = view.getUint32(at + 42, true);
    const entry = Buffer.from(bytes.subarray(at + 46, at + 46 + nameLen)).toString('utf8');
    if (entry === name) {
      const start = local + 30 + view.getUint16(local + 26, true) + view.getUint16(local + 28, true);
      const raw = bytes.subarray(start, start + size);
      return Buffer.from(method === 8 ? inflateRawSync(raw) : raw).toString('utf8');
    }
    at += 46 + nameLen + extraLen + commentLen;
  }
  return null;
};

const box = (mod) => ({
  name: 'Box',
  presetGeometry: 'roundRect',
  text: 'Hello\nWorld',
  anchorKind: mod.AnchorKind.TwoCell,
  fromRow: 1,
  fromCol: 1,
  toRow: 5,
  toCol: 4,
  fillKind: mod.ShapePaint.Solid,
  fillColor: '#FF0000',
  lineKind: mod.ShapePaint.None,
  rotationDeg: 15,
});

test('shapes insert, list, update, remove and round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.sheet(0).listShapes().length, 0);
  const rect = wb.sheet(0).addShape(box(mod));
  const line = wb.sheet(0).addShape({
    presetGeometry: 'line',
    anchorKind: mod.AnchorKind.OneCell,
    fromRow: 10,
    fromColOffsetEmu: 500,
    widthEmu: 914400,
    heightEmu: 0,
    lineKind: mod.ShapePaint.Solid,
    lineColor: '#00FF00',
    lineWidthPt: 1.5,
  });
  assert.notEqual(rect, line);

  const loaded = roundTrip(mod, wb);
  const shapes = loaded.sheet(0).listShapes();
  assert.equal(shapes.length, 2);
  const props = shapes.map((shape) => shape.getProperties());
  assert.equal(shapes[0].objectId, rect.objectId);
  assert.equal(props[0].kind, mod.DrawingObjectKind.Shape);
  assert.equal(props[0].name, 'Box');
  assert.equal(props[0].presetGeometry, 'roundRect');
  assert.equal(props[0].text, 'Hello\nWorld');
  assert.equal(props[0].anchorKind, mod.AnchorKind.TwoCell);
  assert.equal(props[0].toCol, 4);
  assert.equal(props[0].fillKind, mod.ShapePaint.Solid);
  assert.equal(props[0].fillColor.argb, 0xffff0000);
  assert.equal(props[0].lineKind, mod.ShapePaint.None);
  assert.equal(props[0].rotationDeg, 15);
  assert.equal(props[0].hidden, false);
  assert.equal(props[1].kind, mod.DrawingObjectKind.Connector);
  assert.equal(props[1].anchorKind, mod.AnchorKind.OneCell);
  assert.equal(props[1].fromColOffsetEmu, 500);
  assert.equal(props[1].widthEmu, 914400);
  assert.equal(props[1].lineColor.argb, 0xff00ff00);
  assert.equal(props[1].lineWidthPt, 1.5);

  // An update changes the named fields only, and an anchor moves with a row insert.
  loaded.sheet(0).shape(rect.objectId).update({ name: 'Renamed', hidden: true, fillKind: mod.ShapePaint.None });
  loaded.sheet(0).insertRows(0, 3);
  const moved = loaded.sheet(0).listShapes()[0].getProperties();
  assert.equal(moved.name, 'Renamed');
  assert.equal(moved.hidden, true);
  assert.equal(moved.fillKind, mod.ShapePaint.None);
  assert.equal(moved.presetGeometry, 'roundRect');
  assert.equal(moved.rotationDeg, 15);
  assert.equal(moved.fromRow, 4);
  assert.equal(moved.toRow, 8);

  loaded.sheet(0).removeDrawing(rect.objectId);
  assert.deepEqual(
    loaded
      .sheet(0)
      .listShapes()
      .map((s) => s.objectId),
    [line.objectId],
  );
  release(loaded);
  release(wb);
});

test('shape arguments are validated', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.throws(() => wb.sheet(0).addShape({ presetGeometry: 'rect' }), isFormulonError);
  assert.throws(() => wb.sheet(0).addShape({ ...box(mod), anchorKind: 9 }), isFormulonError);
  assert.throws(() => wb.sheet(0).addShape({ ...box(mod), fillKind: mod.ShapePaint.Other }), isFormulonError);
  assert.throws(() => wb.sheet(0).addShape({ ...box(mod), toRow: 0 }), isFormulonError);
  assert.throws(() => wb.sheet(0).addShape({ ...box(mod), presetGeometry: 'a"/>' }), isFormulonError);
  assert.throws(() => wb.sheet(5).addShape(box(mod)), isFormulonError);
  assert.equal(wb.sheet(0).listShapes().length, 0);
  assert.throws(() => wb.sheet(0).shape(99).update({ name: 'x' }), isFormulonError);
  assert.throws(() => wb.sheet(0).removeDrawing(99), isFormulonError);
  assert.throws(() => wb.sheet(0).addShape({ ...box(mod), fromRow: -1 }), RangeError);
  release(wb);
});

const plainOf = (value) => JSON.parse(JSON.stringify(value));
const onSheet1 = (firstRow, firstCol, lastRow, lastCol) => ({
  sheet: 0,
  rect: { firstRow, firstCol, lastRow, lastCol },
});

const sparklineGroup = (mod) => ({
  type: mod.SparklineType.Column,
  colorSeries: { kind: 1, argb: 0xff336699, theme: 0, tint: 0, indexed: 0 },
  colorNegative: { kind: 2, argb: 0, theme: 5, tint: -0.25, indexed: 0 },
  colorLow: { kind: 3, argb: 0, theme: 0, tint: 0, indexed: 12 },
  markers: true,
  negative: true,
  displayXAxis: true,
  minAxisType: mod.SparklineAxisType.Custom,
  manualMin: -1.5,
  lineWeight: 1.25,
  emptyCellsAs: mod.BlanksAs.Zero,
  sparklines: [
    { source: 'Sheet1!A1:D1', location: 'F1' },
    { source: 'Sheet1!A2:D2', location: 'F2' },
  ],
});

test('sparkline groups add, list, update, replace, remove and follow row edits', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.sheet(0).listSparklineGroups().length, 0);
  wb.sheet(0).addSparklineGroup(sparklineGroup(mod));
  wb.sheet(0).addSparklineGroup({ sparklines: [{ source: 'Sheet1!A9:D9', location: 'G9' }] });

  const loaded = roundTrip(mod, wb);
  const groups = loaded.sheet(0).listSparklineGroups();
  assert.equal(groups.length, 2);
  const [first, second] = groups;
  assert.equal(first.type, mod.SparklineType.Column);
  assert.equal(first.colorSeries.argb, 0xff336699);
  assert.equal(first.colorNegative.theme, 5);
  assert.equal(first.colorNegative.tint, -0.25);
  assert.equal(first.colorLow.indexed, 12);
  assert.equal(first.colorHigh.kind, 0);
  assert.equal(first.markers, true);
  assert.equal(first.high, false);
  assert.equal(first.minAxisType, mod.SparklineAxisType.Custom);
  assert.equal(first.manualMin, -1.5);
  assert.equal(first.manualMax, null);
  assert.equal(first.lineWeight, 1.25);
  assert.equal(first.emptyCellsAs, mod.BlanksAs.Zero);
  assert.deepEqual(plainOf(first.sparklines), [
    { source: onSheet1(0, 0, 0, 3), location: { row: 0, col: 5 } },
    { source: onSheet1(1, 0, 1, 3), location: { row: 1, col: 5 } },
  ]);
  assert.equal(second.type, mod.SparklineType.Line);
  assert.equal(second.lineWeight, 0.75);

  loaded.sheet(0).insertRows(0, 2);
  assert.deepEqual(plainOf(loaded.sheet(0).listSparklineGroups()[0].sparklines[1]), {
    source: onSheet1(3, 0, 3, 3),
    location: { row: 3, col: 5 },
  });

  // Update rewrites the modelled fields and keeps the sparklines.
  loaded.sheet(0).updateSparklineGroup(0, {
    ...first,
    type: mod.SparklineType.Stacked,
    high: true,
    dateAxis: true,
    dateAxisRange: 'Sheet1!A4:D4',
  });
  const updated = loaded.sheet(0).listSparklineGroups()[0];
  assert.equal(updated.type, mod.SparklineType.Stacked);
  assert.equal(updated.high, true);
  assert.equal(updated.dateAxis, true);
  assert.deepEqual(plainOf(updated.dateAxisRange), onSheet1(3, 0, 3, 3));
  assert.equal(updated.sparklines.length, 2);

  loaded.sheet(0).updateSparklineGroup(0, { sparklines: [{ source: 'Sheet1!B1:B5', location: 'H1' }] });
  assert.deepEqual(plainOf(loaded.sheet(0).listSparklineGroups()[0].sparklines), [
    { source: onSheet1(0, 1, 4, 1), location: { row: 0, col: 7 } },
  ]);
  // An empty group cannot be saved, and a patch without `sparklines` keeps them.
  assert.throws(() => loaded.sheet(0).updateSparklineGroup(0, { sparklines: [] }), isFormulonError);
  loaded.sheet(0).updateSparklineGroup(0, { high: false });
  assert.equal(loaded.sheet(0).listSparklineGroups()[0].sparklines.length, 1);
  // A patch is partial: omitted keys stay, null clears only the nullable keys, unknown keys are refused.
  loaded.sheet(0).updateSparklineGroup(0, { manualMin: null });
  assert.equal(loaded.sheet(0).listSparklineGroups()[0].manualMin, null);
  assert.equal(loaded.sheet(0).listSparklineGroups()[0].lineWeight, 1.25);
  assert.throws(() => loaded.sheet(0).updateSparklineGroup(0, { lineWeight: null }), TypeError);
  assert.throws(() => loaded.sheet(0).updateSparklineGroup(0, { noSuchKey: 1 }), TypeError);
  assert.throws(
    () =>
      loaded.sheet(0).updateSparklineGroup(0, { sparklines: [{ source: 'Sheet1!A1:D1', location: 'F1', extra: 1 }] }),
    TypeError,
  );
  loaded.sheet(0).removeSparklineGroup(0);
  assert.equal(loaded.sheet(0).listSparklineGroups().length, 1);
  assert.deepEqual(plainOf(loaded.sheet(0).listSparklineGroups()[0].sparklines[0].location), { row: 10, col: 6 });
  release(loaded);
  release(wb);
});

test('sparkline group arguments are validated', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const good = [{ source: 'Sheet1!A1:D1', location: 'F1' }];
  assert.throws(() => wb.sheet(0).addSparklineGroup({}), isFormulonError);
  assert.throws(() => wb.sheet(0).addSparklineGroup({ sparklines: [] }), isFormulonError);
  assert.throws(
    () => wb.sheet(0).addSparklineGroup({ sparklines: [{ source: 'Sheet1!A1:D1', location: 'F1:F2' }] }),
    isFormulonError,
  );
  assert.throws(() => wb.sheet(0).addSparklineGroup({ sparklines: [{ source: '', location: 'F1' }] }), isFormulonError);
  assert.throws(() => wb.sheet(0).addSparklineGroup({ type: 7, sparklines: good }), isFormulonError);
  assert.throws(() => wb.sheet(0).addSparklineGroup({ lineWeight: -1, sparklines: good }), isFormulonError);
  assert.throws(() => wb.sheet(4).addSparklineGroup({ sparklines: good }), isFormulonError);
  assert.equal(wb.sheet(0).listSparklineGroups().length, 0);
  assert.throws(() => wb.sheet(0).updateSparklineGroup(0, {}), isFormulonError);
  assert.throws(() => wb.sheet(0).updateSparklineGroup(0, { sparklines: good }), isFormulonError);
  assert.throws(() => wb.sheet(0).removeSparklineGroup(0), isFormulonError);
  release(wb);
});

test('a surviving note keeps its own shape when notes are added', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setComment(1, 1, { author: 'Alice', text: 'one' });
  wb.sheet(0).setComment(3, 3, { author: 'Alice', text: 'two' });
  wb.sheet(0).setCommentVisible(1, 1, true);
  const first = roundTrip(mod, wb);
  first.sheet(0).setComment(8, 8, { author: 'Alice', text: 'three' });
  first.sheet(0).setCommentVisible(3, 3, true);
  const saved = first.save();
  const vml = readZipEntry(saved, 'xl/drawings/vmlDrawing1.vml');
  assert.ok(vml, 'comment VML part is written');
  assert.equal(vml.match(/<v:shape /g).length, 3);
  assert.equal(new Set(vml.match(/id="_x0000_s\d+"/g)).size, 3);
  assert.equal(vml.match(/visibility:visible/g).length, 2);
  const again = mod.Workbook.load(saved);
  assert.equal(again.sheet(0).getCommentVisible(1, 1), true);
  assert.equal(again.sheet(0).getCommentVisible(3, 3), true);
  assert.equal(again.sheet(0).getCommentVisible(8, 8), false);
  assert.equal(again.sheet(0).getCommentVisible(9, 9), null);
  release(again);
  release(first);
  release(wb);
});
