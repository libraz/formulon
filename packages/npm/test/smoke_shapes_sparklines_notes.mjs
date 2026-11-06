import test from 'node:test';
import { inflateRawSync } from 'node:zlib';
import { assert, expectFailure, getModule, loadStagedModule, readAll } from './smoke_support.mjs';

// ---- Drawing shapes, sparkline groups and legacy note shapes -----------------

const release = (wb) => wb.delete();

// The enums are named exports of the package; the workbook factory is the module.
const load = async () => {
  const [mod, enums] = await Promise.all([getModule(), loadStagedModule()]);
  return { Workbook: mod.Workbook, ...enums };
};

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
  prst: 'roundRect',
  text: 'Hello\nWorld',
  anchorKind: mod.AnchorKind.TwoCell,
  fromRow: 1,
  fromCol: 1,
  toRow: 5,
  toCol: 4,
  fillKind: mod.ShapePaint.Solid,
  fillRgb: 0xff0000,
  lineKind: mod.ShapePaint.None,
  rotation: 900000,
});

test('shapes insert, list, update, remove and round-trip', async () => {
  const mod = await load();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.listShapes(0).length, 0);
  const rect = wb.addShape(0, box(mod));
  const line = wb.addShape(0, {
    prst: 'line',
    anchorKind: mod.AnchorKind.OneCell,
    fromRow: 10,
    fromColOffset: 500,
    cx: 914400,
    cy: 0,
    lineKind: mod.ShapePaint.Solid,
    lineRgb: 0x00ff00,
    lineWidthEmu: 19050,
  });
  assert.notEqual(rect, line);

  const loaded = roundTrip(mod, wb);
  const shapes = loaded.listShapes(0);
  assert.equal(shapes.length, 2);
  assert.equal(shapes[0].objectId, rect);
  assert.equal(shapes[0].kind, mod.DrawingObjectKind.Shape);
  assert.equal(shapes[0].name, 'Box');
  assert.equal(shapes[0].prst, 'roundRect');
  assert.equal(shapes[0].text, 'Hello\nWorld');
  assert.equal(shapes[0].anchorKind, mod.AnchorKind.TwoCell);
  assert.equal(shapes[0].toCol, 4);
  assert.equal(shapes[0].fillKind, mod.ShapePaint.Solid);
  assert.equal(shapes[0].fillRgb, 0xff0000);
  assert.equal(shapes[0].lineKind, mod.ShapePaint.None);
  assert.equal(shapes[0].rotation, 900000);
  assert.equal(shapes[0].hidden, false);
  assert.equal(shapes[1].kind, mod.DrawingObjectKind.Connector);
  assert.equal(shapes[1].anchorKind, mod.AnchorKind.OneCell);
  assert.equal(shapes[1].fromColOffset, 500);
  assert.equal(shapes[1].cx, 914400);
  assert.equal(shapes[1].lineRgb, 0x00ff00);
  assert.equal(shapes[1].lineWidthEmu, 19050);

  // An update changes the named fields only, and an anchor moves with a row insert.
  loaded.updateShape(0, rect, { name: 'Renamed', hidden: true, fillKind: mod.ShapePaint.None });
  loaded.insertRows(0, 0, 3);
  const moved = loaded.listShapes(0)[0];
  assert.equal(moved.name, 'Renamed');
  assert.equal(moved.hidden, true);
  assert.equal(moved.fillKind, mod.ShapePaint.None);
  assert.equal(moved.prst, 'roundRect');
  assert.equal(moved.rotation, 900000);
  assert.equal(moved.fromRow, 4);
  assert.equal(moved.toRow, 8);

  loaded.removeShape(0, rect);
  assert.deepEqual(
    loaded.listShapes(0).map((s) => s.objectId),
    [line],
  );
  release(loaded);
  release(wb);
});

test('shape arguments are validated', async () => {
  const mod = await load();
  const wb = mod.Workbook.createDefault();
  expectFailure(() => wb.addShape(0, { prst: 'rect' }));
  expectFailure(() => wb.addShape(0, { ...box(mod), anchorKind: 9 }));
  expectFailure(() => wb.addShape(0, { ...box(mod), fillKind: mod.ShapePaint.Other }));
  expectFailure(() => wb.addShape(0, { ...box(mod), toRow: 0 }));
  expectFailure(() => wb.addShape(0, { ...box(mod), prst: 'a"/>' }));
  expectFailure(() => wb.addShape(5, box(mod)));
  assert.equal(wb.listShapes(0).length, 0);
  expectFailure(() => wb.updateShape(0, 99, { name: 'x' }));
  expectFailure(() => wb.removeShape(0, 99));
  assert.throws(() => wb.addShape(0, { ...box(mod), fromRow: -1 }), RangeError);
  release(wb);
});

const sparklineGroup = (mod) => ({
  type: mod.SparklineType.Column,
  colorSeries: { kind: 1, rgb: 0xff336699, theme: 0, tint: 0, indexed: 0 },
  colorNegative: { kind: 2, rgb: 0, theme: 5, tint: -0.25, indexed: 0 },
  colorLow: { kind: 3, rgb: 0, theme: 0, tint: 0, indexed: 12 },
  markers: true,
  negative: true,
  displayXAxis: true,
  minAxisType: mod.SparklineAxisType.Custom,
  manualMin: -1.5,
  lineWeight: 1.25,
  emptyCellsAs: mod.SparklineEmptyCells.Zero,
  sparklines: [
    { formula: 'Sheet1!A1:D1', location: 'F1' },
    { formula: 'Sheet1!A2:D2', location: 'F2' },
  ],
});

test('sparkline groups add, list, update, replace, remove and follow row edits', async () => {
  const mod = await load();
  const wb = mod.Workbook.createDefault();
  assert.equal(
    readAll(
      () => wb.getSparklineGroupCount(0),
      (i) => wb.getSparklineGroup(0, i),
    ).length,
    0,
  );
  wb.addSparklineGroup(0, sparklineGroup(mod));
  wb.addSparklineGroup(0, { sparklines: [{ formula: 'Sheet1!A9:D9', location: 'G9' }] });

  const loaded = roundTrip(mod, wb);
  const groups = readAll(
    () => loaded.getSparklineGroupCount(0),
    (i) => loaded.getSparklineGroup(0, i),
  );
  assert.equal(groups.length, 2);
  const [first, second] = groups;
  assert.equal(first.type, mod.SparklineType.Column);
  assert.equal(first.colorSeries.rgb, 0xff336699);
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
  assert.equal(first.emptyCellsAs, mod.SparklineEmptyCells.Zero);
  assert.deepEqual(first.sparklines, [
    { formula: 'Sheet1!A1:D1', location: 'F1' },
    { formula: 'Sheet1!A2:D2', location: 'F2' },
  ]);
  assert.equal(second.type, mod.SparklineType.Line);
  assert.equal(second.lineWeight, 0.75);

  loaded.insertRows(0, 0, 2);
  assert.equal(
    readAll(
      () => loaded.getSparklineGroupCount(0),
      (i) => loaded.getSparklineGroup(0, i),
    )[0].sparklines[1].location,
    'F4',
  );
  assert.equal(
    readAll(
      () => loaded.getSparklineGroupCount(0),
      (i) => loaded.getSparklineGroup(0, i),
    )[0].sparklines[1].formula,
    'Sheet1!A4:D4',
  );

  // Update rewrites the modelled fields and keeps the sparklines.
  loaded.updateSparklineGroup(0, 0, {
    ...first,
    type: mod.SparklineType.Stacked,
    high: true,
    dateAxis: true,
    dateAxisRef: 'Sheet1!A4:D4',
  });
  const updated = readAll(
    () => loaded.getSparklineGroupCount(0),
    (i) => loaded.getSparklineGroup(0, i),
  )[0];
  assert.equal(updated.type, mod.SparklineType.Stacked);
  assert.equal(updated.high, true);
  assert.equal(updated.dateAxis, true);
  assert.equal(updated.dateAxisRef, 'Sheet1!A4:D4');
  assert.equal(updated.sparklines.length, 2);

  loaded.setSparklineGroupSparklines(0, 0, [{ formula: 'Sheet1!B1:B5', location: 'H1' }]);
  assert.deepEqual(
    readAll(
      () => loaded.getSparklineGroupCount(0),
      (i) => loaded.getSparklineGroup(0, i),
    )[0].sparklines,
    [{ formula: 'Sheet1!B1:B5', location: 'H1' }],
  );
  loaded.removeSparklineGroup(0, 0);
  assert.equal(
    readAll(
      () => loaded.getSparklineGroupCount(0),
      (i) => loaded.getSparklineGroup(0, i),
    ).length,
    1,
  );
  assert.equal(
    readAll(
      () => loaded.getSparklineGroupCount(0),
      (i) => loaded.getSparklineGroup(0, i),
    )[0].sparklines[0].location,
    'G11',
  );
  release(loaded);
  release(wb);
});

test('sparkline group arguments are validated', async () => {
  const mod = await load();
  const wb = mod.Workbook.createDefault();
  const good = [{ formula: 'Sheet1!A1:D1', location: 'F1' }];
  expectFailure(() => wb.addSparklineGroup(0, {}));
  expectFailure(() => wb.addSparklineGroup(0, { sparklines: [] }));
  expectFailure(() => wb.addSparklineGroup(0, { sparklines: [{ formula: 'Sheet1!A1:D1', location: 'F1:F2' }] }));
  expectFailure(() => wb.addSparklineGroup(0, { sparklines: [{ formula: '', location: 'F1' }] }));
  expectFailure(() => wb.addSparklineGroup(0, { type: 7, sparklines: good }));
  expectFailure(() => wb.addSparklineGroup(0, { lineWeight: -1, sparklines: good }));
  expectFailure(() => wb.addSparklineGroup(4, { sparklines: good }));
  assert.equal(
    readAll(
      () => wb.getSparklineGroupCount(0),
      (i) => wb.getSparklineGroup(0, i),
    ).length,
    0,
  );
  expectFailure(() => wb.updateSparklineGroup(0, 0, {}));
  expectFailure(() => wb.setSparklineGroupSparklines(0, 0, good));
  expectFailure(() => wb.removeSparklineGroup(0, 0));
  release(wb);
});

test('a surviving note keeps its own shape when notes are added', async () => {
  const mod = await load();
  const wb = mod.Workbook.createDefault();
  wb.setComment(0, 1, 1, 'Alice', 'one');
  wb.setComment(0, 3, 3, 'Alice', 'two');
  wb.setCommentVisible(0, 1, 1, true);
  const first = roundTrip(mod, wb);
  first.setComment(0, 8, 8, 'Alice', 'three');
  first.setCommentVisible(0, 3, 3, true);
  const saved = first.save();
  const vml = readZipEntry(saved, 'xl/drawings/vmlDrawing1.vml');
  assert.ok(vml, 'comment VML part is written');
  assert.equal(vml.match(/<v:shape /g).length, 3);
  assert.equal(new Set(vml.match(/id="_x0000_s\d+"/g)).size, 3);
  assert.equal(vml.match(/visibility:visible/g).length, 2);
  const again = mod.Workbook.load(saved);
  assert.equal(again.getCommentVisible(0, 1, 1), true);
  assert.equal(again.getCommentVisible(0, 3, 3), true);
  assert.equal(again.getCommentVisible(0, 8, 8), false);
  assert.equal(again.getCommentVisible(0, 9, 9), null);
  release(again);
  release(first);
  release(wb);
});
