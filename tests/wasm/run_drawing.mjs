import assert from 'node:assert/strict';
import { deflateSync } from 'node:zlib';
import {
  AnchorEditAs,
  AnchorKind,
  DrawingObjectKind,
  ImageFormat,
  ShapePaint,
  StatusCode,
} from '../../packages/npm/common.mjs';
import { expectFailure } from './smoke_support.mjs';

const EMU_PER_PX = 9525;

function crc32(buf) {
  let c = ~0;
  for (const b of buf) {
    c ^= b;
    for (let k = 0; k < 8; k += 1) c = (c >>> 1) ^ (0xedb88320 & -(c & 1));
  }
  return ~c >>> 0;
}

function chunk(type, data) {
  const out = Buffer.alloc(12 + data.length);
  out.writeUInt32BE(data.length, 0);
  out.write(type, 4, 'latin1');
  data.copy(out, 8);
  out.writeUInt32BE(crc32(out.subarray(4, 8 + data.length)), 8 + data.length);
  return out;
}

// A solid-colour 8-bit RGB PNG of the given pixel size.
function makePng(width, height) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const raw = Buffer.alloc((1 + width * 3) * height, 0x7f);
  for (let y = 0; y < height; y += 1) raw[y * (1 + width * 3)] = 0;
  return new Uint8Array(
    Buffer.concat([
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      chunk('IHDR', ihdr),
      chunk('IDAT', deflateSync(raw)),
      chunk('IEND', Buffer.alloc(0)),
    ]),
  );
}

function withWorkbook(Module, fn) {
  const wb = Module.Workbook.createDefault();
  try {
    fn(wb);
  } finally {
    wb.dispose();
  }
}

export function registerDrawing(Module, test) {
  test('drawing enums keep their C ordinals', () => {
    assert.equal(ImageFormat.Bmp, 4);
    assert.equal(DrawingObjectKind.GraphicFrame, 5);
    assert.equal(DrawingObjectKind.Other, 6);
    assert.equal(AnchorKind.Default, 0);
    assert.equal(AnchorKind.Absolute, 3);
    assert.equal(AnchorEditAs.Default, 0);
    assert.equal(AnchorEditAs.OneCell, 1);
    assert.equal(AnchorEditAs.OneCell, AnchorKind.OneCell);
  });

  test('getImageInfo reports the format and pixel size, and rejects non-images', () => {
    const info = Module.getImageInfo(makePng(3, 2));
    assert.equal(info.format, ImageFormat.Png);
    assert.equal(info.widthPx, 3);
    assert.equal(info.heightPx, 2);
    expectFailure(() => Module.getImageInfo(new Uint8Array([1, 2, 3, 4])));
  });

  test('addImage, listDrawingObjects, getPicture and removeDrawing round-trip through save and load', () => {
    withWorkbook(Module, (wb) => {
      const png = makePng(4, 3);
      assert.equal(wb.listDrawingObjects(0).length, 0);
      wb.listDrawingObjects(0);

      const ins = wb.addImage(0, png, {
        name: 'Logo',
        description: 'alt text',
        anchor: { fromRow: 2, fromCol: 1, fromColOffsetEmu: 100 },
      });
      assert.ok(ins > 0);

      const check = (book) => {
        const list = book.listDrawingObjects(0);
        assert.equal(list.length, 1);
        const o = list[0];
        assert.equal(o.objectId, ins);
        assert.equal(o.kind, DrawingObjectKind.Picture);
        assert.equal(o.anchor.kind, AnchorKind.OneCell);
        assert.equal(o.anchor.fromRow, 2);
        assert.equal(o.anchor.fromCol, 1);
        assert.equal(o.anchor.fromColOffsetEmu, 100);
        assert.equal(o.anchor.widthEmu, 4 * EMU_PER_PX);
        assert.equal(o.anchor.heightEmu, 3 * EMU_PER_PX);
        assert.equal(o.hidden, false);
        assert.equal(o.name, 'Logo');
        assert.equal(o.description, 'alt text');
        const found = book.getPicture(0, o.objectId);
        const img = found;
        assert.equal(img.format, ImageFormat.Png);
        assert.equal(img.widthPx, 4);
        assert.equal(img.heightPx, 3);
        assert.ok(img.bytes instanceof Uint8Array);
        assert.deepEqual(Buffer.from(img.bytes), Buffer.from(png));
      };
      check(wb);

      const saved = wb.save();
      const loaded = Module.Workbook.load(saved);
      try {
        check(loaded);
        loaded.removeDrawingObject(0, ins);
        assert.equal(loaded.listDrawingObjects(0).length, 0);
        assert.equal(loaded.getPicture(0, ins), null);
        expectFailure(() => loaded.removeDrawingObject(0, ins), StatusCode.NotFound);
      } finally {
        loaded.dispose();
      }
    });
  });

  test('insertImage honours a two-cell anchor, an explicit size and the editAs mode', () => {
    withWorkbook(Module, (wb) => {
      const ins = wb.addImage(0, makePng(2, 2), {
        anchor: {
          kind: AnchorKind.TwoCell,
          editAs: AnchorEditAs.OneCell,
          widthEmu: 2 * EMU_PER_PX * 10,
          heightEmu: EMU_PER_PX * 20,
        },
      });
      const o = wb.listDrawingObjects(0)[0];
      assert.equal(o.anchor.kind, AnchorKind.TwoCell);
      assert.equal(o.anchor.editAs, AnchorEditAs.OneCell);
      assert.equal(o.anchor.widthEmu, 2 * EMU_PER_PX * 10);
      assert.equal(o.anchor.heightEmu, EMU_PER_PX * 20);
      // Default names count from 1 (the id is one ahead) with the profile's word.
      assert.equal(o.name, `Picture ${ins - 1}`);
    });
  });

  test('image calls throw on failure and return null for a missing image', () => {
    withWorkbook(Module, (wb) => {
      expectFailure(() => wb.addImage(0, new Uint8Array([1, 2, 3, 4])));
      expectFailure(() => wb.addImage(99, makePng(1, 1)));
      expectFailure(() => wb.addImage(0, makePng(1, 1), { anchor: { kind: AnchorKind.Absolute } }));
      expectFailure(() => wb.listDrawingObjects(99));
      assert.equal(wb.getPicture(0, 12345), null);
    });
  });

  test('updateDrawing merges an anchor patch per key and keeps the object id', () => {
    withWorkbook(Module, (wb) => {
      const ins = wb.addImage(0, makePng(4, 3), { name: 'Logo', anchor: { fromRow: 1, fromCol: 1 } });
      wb.updateDrawing(0, ins, { anchor: { fromRow: 5, fromRowOffsetEmu: 1000 } });
      let o = wb.listDrawingObjects(0)[0];
      assert.equal(o.objectId, ins);
      assert.equal(o.name, 'Logo');
      assert.equal(o.anchor.kind, AnchorKind.OneCell);
      assert.equal(o.anchor.fromRow, 5);
      assert.equal(o.anchor.fromCol, 1);
      assert.equal(o.anchor.fromRowOffsetEmu, 1000);
      assert.equal(o.anchor.widthEmu, 4 * EMU_PER_PX);

      wb.updateDrawing(0, ins, {
        anchor: { kind: AnchorKind.TwoCell, editAs: AnchorEditAs.OneCell, toRow: 9, toCol: 6 },
      });
      o = wb.listDrawingObjects(0)[0];
      assert.equal(o.anchor.kind, AnchorKind.TwoCell);
      assert.equal(o.anchor.editAs, AnchorEditAs.OneCell);
      assert.equal(o.anchor.fromRow, 5);
      assert.equal(o.anchor.toRow, 9);
      assert.equal(o.anchor.toCol, 6);

      expectFailure(() => wb.updateDrawing(0, ins, { anchor: { widthEmu: 1 } }), StatusCode.InvalidArgument);
      expectFailure(() => wb.updateDrawing(0, 9999, {}), StatusCode.NotFound);
      assert.throws(() => wb.updateDrawing(0, ins, { anchor: { fromRow: 'x' } }), TypeError);
      assert.throws(() => wb.updateDrawing(0, ins, { anchor: { fromRow: null } }), TypeError);
    });
  });

  test('updateDrawing sets name, description and hidden on any drawing object', () => {
    withWorkbook(Module, (wb) => {
      const ins = wb.addImage(0, makePng(2, 2), { name: 'A', description: 'alt' });
      wb.updateDrawing(0, ins, { name: 'B', description: 'new', hidden: true });
      let o = wb.getDrawingObject(0, ins);
      assert.equal(o.name, 'B');
      assert.equal(o.description, 'new');
      assert.equal(o.hidden, true);
      wb.updateDrawing(0, ins, { description: null, hidden: false });
      o = wb.getDrawingObject(0, ins);
      assert.equal(o.description, null);
      assert.equal(o.hidden, false);
      assert.throws(() => wb.updateDrawing(0, ins, { name: null }), TypeError);
      assert.throws(() => wb.updateDrawing(0, ins, { hidden: null }), TypeError);
    });
  });

  test('setDrawingZOrder reorders pictures in listDrawingObjects order', () => {
    withWorkbook(Module, (wb) => {
      const a = wb.addImage(0, makePng(2, 2), { anchor: { fromRow: 0 } });
      const b = wb.addImage(0, makePng(2, 2), { anchor: { fromRow: 3 } });
      const c = wb.addImage(0, makePng(2, 2), { anchor: { fromRow: 6 } });
      const ids = () => Array.from(wb.listDrawingObjects(0), (o) => o.objectId);
      assert.deepEqual(ids(), [a, b, c]);
      wb.setDrawingZOrder(0, c, 0);
      assert.deepEqual(ids(), [c, a, b]);
      expectFailure(() => wb.setDrawingZOrder(0, c, 3));
      expectFailure(() => wb.setDrawingZOrder(0, 9999, 0));
      assert.deepEqual(ids(), [c, a, b]);
    });
  });

  test('snapshotDrawing and restoreDrawing bring back a removed, moved or copied picture', () => {
    withWorkbook(Module, (wb) => {
      const png = makePng(4, 3);
      const a = wb.addImage(0, png, { name: 'A', anchor: { fromRow: 1 } });
      const b = wb.addImage(0, makePng(2, 2), { name: 'B', anchor: { fromRow: 4 } });
      const snap = wb.snapshotDrawing(0, a);
      assert.ok(snap instanceof Uint8Array);
      assert.ok(snap.length > 0);
      const ids = () => Array.from(wb.listDrawingObjects(0), (o) => o.objectId);

      wb.removeDrawingObject(0, a);
      const restored = wb.restoreDrawing(0, snap);
      assert.equal(restored, a);
      assert.deepEqual(ids(), [a, b]);
      assert.deepEqual(Buffer.from(wb.getPicture(0, a).bytes), Buffer.from(png));

      wb.updateDrawing(0, a, { anchor: { fromRow: 9 } });
      wb.setDrawingZOrder(0, a, 1);
      wb.restoreDrawing(0, snap, {});
      assert.deepEqual(ids(), [a, b]);
      assert.equal(wb.listDrawingObjects(0)[0].anchor.fromRow, 1);

      const copy = wb.restoreDrawing(0, snap, { assignNewId: true });
      assert.notEqual(copy, a);
      assert.deepEqual(ids(), [a, b, copy]);

      expectFailure(() => wb.restoreDrawing(0, new Uint8Array([1, 2, 3, 4])));
      expectFailure(() => wb.snapshotDrawing(0, 9999));
    });
  });

  test('addShape takes placement keys and reads shape fields apart from the common record', () => {
    withWorkbook(Module, (wb) => {
      const id = wb.addShape(0, {
        name: 'Box',
        description: 'a box',
        hidden: true,
        anchor: { fromRow: 1, fromCol: 1, toRow: 4, toCol: 4 },
        presetGeometry: 'rect',
        text: 'hi',
      });
      const common = wb.getDrawingObject(0, id);
      assert.equal(common.kind, DrawingObjectKind.Shape);
      assert.equal(common.name, 'Box');
      assert.equal(common.description, 'a box');
      assert.equal(common.hidden, true);
      assert.equal(common.anchor.kind, AnchorKind.TwoCell);
      assert.equal(common.anchor.toRow, 4);

      const shape = wb.getShape(0, id);
      assert.equal(shape.presetGeometry, 'rect');
      assert.equal(shape.text, 'hi');
      assert.equal(shape.fillKind, ShapePaint.Default);
      assert.equal(shape.lineKind, ShapePaint.Default);
      assert.equal('name' in shape, false);
      assert.equal('fromRow' in shape, false);

      wb.updateShape(0, id, { fillKind: ShapePaint.Solid, fillColor: { kind: 1, argb: 0xffff0000 } });
      assert.equal(wb.getShape(0, id).fillKind, ShapePaint.Solid);
      wb.updateShape(0, id, { fillKind: ShapePaint.Default });
      assert.equal(wb.getShape(0, id).fillKind, ShapePaint.Default);
      expectFailure(() => wb.updateShape(0, id, { fillKind: null }), StatusCode.InvalidArgument);

      wb.removeDrawingObject(0, id);
      assert.equal(wb.getShape(0, id), null);
    });
  });

  test('sparkline groups are addressed by their id', () => {
    withWorkbook(Module, (wb) => {
      const a = wb.addSparklineGroup(0, { sparklines: [{ formula: 'Sheet1!A1:C1', location: 'D1' }] });
      const b = wb.addSparklineGroup(0, { sparklines: [{ formula: 'Sheet1!A2:C2', location: 'D2' }] });
      assert.equal(typeof a, 'string');
      assert.notEqual(a, b);
      assert.equal(wb.getSparklineGroupCount(0), 2);
      assert.equal(wb.getSparklineGroup(0, 0).id, a);
      assert.equal(wb.getSparklineGroup(0, 1).id, b);

      wb.updateSparklineGroup(0, b, { markers: true });
      assert.equal(wb.getSparklineGroup(0, 1).markers, true);
      wb.removeSparklineGroup(0, a);
      assert.equal(wb.getSparklineGroupCount(0), 1);
      assert.equal(wb.getSparklineGroup(0, 0).id, b);
      expectFailure(() => wb.removeSparklineGroup(0, a), StatusCode.NotFound);
      expectFailure(() => wb.updateSparklineGroup(0, a, {}), StatusCode.NotFound);
    });
  });
}
