import assert from 'node:assert/strict';
import { deflateSync } from 'node:zlib';
import { AnchorEditAs, AnchorKind, DrawingObjectKind, ImageFormat } from '../../packages/npm/common.mjs';

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
    wb.delete();
  }
}

export function registerDrawing(Module, test) {
  test('drawing enums keep their C ordinals', () => {
    assert.equal(ImageFormat.Bmp, 4);
    assert.equal(DrawingObjectKind.GraphicFrame, 5);
    assert.equal(DrawingObjectKind.Other, 6);
    assert.equal(AnchorKind.Absolute, 2);
    assert.equal(AnchorEditAs.OneCell, 1);
  });

  test('probeImage reports the format and pixel size, and rejects non-images', () => {
    withWorkbook(Module, (wb) => {
      const info = wb.probeImage(makePng(3, 2));
      assert.ok(info.status.ok, JSON.stringify(info.status));
      assert.equal(info.format, ImageFormat.Png);
      assert.equal(info.pxWidth, 3);
      assert.equal(info.pxHeight, 2);
      const bad = wb.probeImage(new Uint8Array([1, 2, 3, 4]));
      assert.equal(bad.status.ok, false);
      assert.equal(bad.pxWidth, 0);
    });
  });

  test('insertImage, listDrawingObjects, getImage and removeImage round-trip through save and load', () => {
    withWorkbook(Module, (wb) => {
      const png = makePng(4, 3);
      assert.equal(wb.listDrawingObjects(0).length, 0);
      assert.ok(wb.listDrawingObjects(0).status.ok);

      const ins = wb.insertImage(0, png, { name: 'Logo', descr: 'alt text', row: 2, col: 1, colOffEmu: 100 });
      assert.ok(ins.status.ok, JSON.stringify(ins.status));
      assert.ok(ins.objectId > 0);

      const check = (book) => {
        const list = book.listDrawingObjects(0);
        assert.ok(list.status.ok, JSON.stringify(list.status));
        assert.equal(list.length, 1);
        const o = list[0];
        assert.equal(o.objectId, ins.objectId);
        assert.equal(o.kind, DrawingObjectKind.Picture);
        assert.equal(o.anchorKind, AnchorKind.OneCell);
        assert.equal(o.fromRow, 2);
        assert.equal(o.fromCol, 1);
        assert.equal(o.fromColOff, 100);
        assert.equal(o.cx, 4 * EMU_PER_PX);
        assert.equal(o.cy, 3 * EMU_PER_PX);
        assert.equal(o.imageFormat, ImageFormat.Png);
        assert.equal(o.name, 'Logo');
        assert.equal(o.descr, 'alt text');
        assert.ok(o.mediaPath.length > 0);
        const img = book.getImage(0, o.objectId);
        assert.ok(img.status.ok, JSON.stringify(img.status));
        assert.equal(img.format, ImageFormat.Png);
        assert.equal(img.pxWidth, 4);
        assert.equal(img.pxHeight, 3);
        assert.ok(img.bytes instanceof Uint8Array);
        assert.deepEqual(Buffer.from(img.bytes), Buffer.from(png));
      };
      check(wb);

      const saved = wb.save();
      assert.ok(saved.status.ok, JSON.stringify(saved.status));
      const loaded = Module.Workbook.loadBytes(saved.bytes);
      try {
        check(loaded);
        const rm = loaded.removeImage(0, ins.objectId);
        assert.ok(rm.ok, JSON.stringify(rm));
        assert.equal(loaded.listDrawingObjects(0).length, 0);
        assert.equal(loaded.getImage(0, ins.objectId).status.ok, false);
        assert.equal(loaded.removeImage(0, ins.objectId).ok, false);
      } finally {
        loaded.delete();
      }
    });
  });

  test('insertImage honours a two-cell anchor, an explicit size and the editAs mode', () => {
    withWorkbook(Module, (wb) => {
      const ins = wb.insertImage(0, makePng(2, 2), {
        anchorKind: AnchorKind.TwoCell,
        editAs: AnchorEditAs.OneCell,
        widthEmu: 2 * EMU_PER_PX * 10,
        heightEmu: EMU_PER_PX * 20,
      });
      assert.ok(ins.status.ok, JSON.stringify(ins.status));
      const o = wb.listDrawingObjects(0)[0];
      assert.equal(o.anchorKind, AnchorKind.TwoCell);
      assert.equal(o.editAs, AnchorEditAs.OneCell);
      assert.equal(o.cx, 2 * EMU_PER_PX * 10);
      assert.equal(o.cy, EMU_PER_PX * 20);
      assert.equal(o.name, `Picture ${ins.objectId}`);
    });
  });

  test('image calls report failures on the envelope without throwing', () => {
    withWorkbook(Module, (wb) => {
      assert.equal(wb.insertImage(0, new Uint8Array([1, 2, 3, 4])).status.ok, false);
      assert.equal(wb.insertImage(0, new Uint8Array([1, 2, 3, 4])).objectId, 0);
      assert.equal(wb.insertImage(99, makePng(1, 1)).status.ok, false);
      assert.equal(wb.insertImage(0, makePng(1, 1), { anchorKind: AnchorKind.Absolute }).status.ok, false);
      assert.equal(wb.listDrawingObjects(99).status.ok, false);
      const img = wb.getImage(0, 12345);
      assert.equal(img.status.ok, false);
      assert.equal(img.bytes.length, 0);
    });
  });
}
