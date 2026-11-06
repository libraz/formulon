import assert from 'node:assert/strict';
import { deflateSync } from 'node:zlib';
import { AnchorEditAs, AnchorKind, DrawingObjectKind, ImageFormat, StatusCode } from '../../packages/npm/common.mjs';
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
    assert.equal(AnchorKind.Absolute, 2);
    assert.equal(AnchorEditAs.OneCell, 0);
    assert.equal(AnchorEditAs.OneCell, AnchorKind.OneCell);
  });

  test('probeImage reports the format and pixel size, and rejects non-images', () => {
    const info = Module.getImageInfo(makePng(3, 2));
    assert.equal(info.format, ImageFormat.Png);
    assert.equal(info.pxWidth, 3);
    assert.equal(info.pxHeight, 2);
    expectFailure(() => Module.getImageInfo(new Uint8Array([1, 2, 3, 4])));
  });

  test('insertImage, listDrawingObjects, getImage and removeDrawing round-trip through save and load', () => {
    withWorkbook(Module, (wb) => {
      const png = makePng(4, 3);
      assert.equal(wb.listDrawingObjects(0).length, 0);
      wb.listDrawingObjects(0);

      const ins = wb.addImage(0, png, {
        name: 'Logo',
        description: 'alt text',
        anchor: { row: 2, col: 1, colOffsetEmu: 100 },
      });
      assert.ok(ins > 0);

      const check = (book) => {
        const list = book.listDrawingObjects(0);
        assert.equal(list.length, 1);
        const o = list[0];
        assert.equal(o.objectId, ins);
        assert.equal(o.kind, DrawingObjectKind.Picture);
        assert.equal(o.anchorKind, AnchorKind.OneCell);
        assert.equal(o.fromRow, 2);
        assert.equal(o.fromCol, 1);
        assert.equal(o.fromColOffsetEmu, 100);
        assert.equal(o.widthEmu, 4 * EMU_PER_PX);
        assert.equal(o.heightEmu, 3 * EMU_PER_PX);
        assert.equal(o.imageFormat, ImageFormat.Png);
        assert.equal(o.name, 'Logo');
        assert.equal(o.description, 'alt text');
        assert.ok(o.mediaPath.length > 0);
        const found = book.getImage(0, o.objectId);
        const img = found;
        assert.equal(img.format, ImageFormat.Png);
        assert.equal(img.pxWidth, 4);
        assert.equal(img.pxHeight, 3);
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
        assert.equal(loaded.getImage(0, ins), null);
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
          anchorKind: AnchorKind.TwoCell,
          editAs: AnchorEditAs.OneCell,
          widthEmu: 2 * EMU_PER_PX * 10,
          heightEmu: EMU_PER_PX * 20,
        },
      });
      const o = wb.listDrawingObjects(0)[0];
      assert.equal(o.anchorKind, AnchorKind.TwoCell);
      assert.equal(o.editAs, AnchorEditAs.OneCell);
      assert.equal(o.widthEmu, 2 * EMU_PER_PX * 10);
      assert.equal(o.heightEmu, EMU_PER_PX * 20);
      // Default names count from 1 (the id is one ahead) with the profile's word.
      assert.equal(o.name, `Picture ${ins - 1}`);
    });
  });

  test('image calls throw on failure and return null for a missing image', () => {
    withWorkbook(Module, (wb) => {
      expectFailure(() => wb.addImage(0, new Uint8Array([1, 2, 3, 4])));
      expectFailure(() => wb.addImage(99, makePng(1, 1)));
      expectFailure(() => wb.addImage(0, makePng(1, 1), { anchor: { anchorKind: AnchorKind.Absolute } }));
      expectFailure(() => wb.listDrawingObjects(99));
      assert.equal(wb.getImage(0, 12345), null);
    });
  });

  test('setDrawingAnchor moves and resizes a picture and keeps its id', () => {
    withWorkbook(Module, (wb) => {
      const ins = wb.addImage(0, makePng(4, 3), { name: 'Logo', anchor: { row: 1, col: 1 } });
      wb.setDrawingAnchor(0, ins, {
        anchorKind: AnchorKind.TwoCell,
        editAs: AnchorEditAs.OneCell,
        row: 5,
        col: 3,
        rowOffsetEmu: 1000,
        colOffsetEmu: 2000,
        widthEmu: 200000,
      });
      const o = wb.listDrawingObjects(0)[0];
      assert.equal(o.objectId, ins);
      assert.equal(o.name, 'Logo');
      assert.equal(o.anchorKind, AnchorKind.TwoCell);
      assert.equal(o.editAs, AnchorEditAs.OneCell);
      assert.equal(o.fromRow, 5);
      assert.equal(o.fromCol, 3);
      assert.equal(o.fromRowOffsetEmu, 1000);
      assert.equal(o.fromColOffsetEmu, 2000);
      assert.equal(o.widthEmu, 200000);
      assert.equal(o.heightEmu, 3 * EMU_PER_PX);
      expectFailure(() => wb.setDrawingAnchor(0, ins, { anchorKind: AnchorKind.Absolute }));
      expectFailure(() => wb.setDrawingAnchor(0, 9999, {}));
      assert.throws(() => wb.setDrawingAnchor(0, ins, { row: 'x' }), TypeError);
    });
  });

  test('setDrawingZOrder reorders pictures in listDrawingObjects order', () => {
    withWorkbook(Module, (wb) => {
      const a = wb.addImage(0, makePng(2, 2), { anchor: { row: 0 } });
      const b = wb.addImage(0, makePng(2, 2), { anchor: { row: 3 } });
      const c = wb.addImage(0, makePng(2, 2), { anchor: { row: 6 } });
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
      const a = wb.addImage(0, png, { name: 'A', anchor: { row: 1 } });
      const b = wb.addImage(0, makePng(2, 2), { name: 'B', anchor: { row: 4 } });
      const snap = wb.snapshotDrawing(0, a);
      assert.ok(snap instanceof Uint8Array);
      assert.ok(snap.length > 0);
      const ids = () => Array.from(wb.listDrawingObjects(0), (o) => o.objectId);

      wb.removeDrawingObject(0, a);
      const restored = wb.restoreDrawing(0, snap);
      assert.equal(restored, a);
      assert.deepEqual(ids(), [a, b]);
      assert.deepEqual(Buffer.from(wb.getImage(0, a).bytes), Buffer.from(png));

      wb.setDrawingAnchor(0, a, { row: 9 });
      wb.setDrawingZOrder(0, a, 1);
      wb.restoreDrawing(0, snap, {});
      assert.deepEqual(ids(), [a, b]);
      assert.equal(wb.listDrawingObjects(0)[0].fromRow, 1);

      const copy = wb.restoreDrawing(0, snap, { newId: true });
      assert.notEqual(copy, a);
      assert.deepEqual(ids(), [a, b, copy]);

      expectFailure(() => wb.restoreDrawing(0, new Uint8Array([1, 2, 3, 4])));
      expectFailure(() => wb.snapshotDrawing(0, 9999));
    });
  });
}
