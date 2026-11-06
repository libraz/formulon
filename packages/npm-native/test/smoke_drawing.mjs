import test from 'node:test';
import zlib from 'node:zlib';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

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

// A valid width x 1 RGB PNG.
function makePng(width) {
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(1, 4);
  ihdr[8] = 8;
  ihdr[9] = 2;
  const raw = Buffer.alloc(1 + width * 3, width);
  raw[0] = 0;
  return new Uint8Array(
    Buffer.concat([
      Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]),
      chunk('IHDR', ihdr),
      chunk('IDAT', zlib.deflateSync(raw)),
      chunk('IEND', Buffer.alloc(0)),
    ]),
  );
}

test('drawing enums keep their C ordinals', async () => {
  const mod = await getModule();
  assert.equal(mod.ImageFormat.Png, 1);
  assert.equal(mod.ImageFormat.Bmp, 4);
  assert.equal(mod.DrawingObjectKind.Picture, 0);
  assert.equal(mod.DrawingObjectKind.Other, 6);
  assert.equal(mod.AnchorKind.Absolute, 2);
  assert.equal(mod.AnchorEditAs.OneCell, 1);
});

test('images probe, insert, list, read back, survive save and load, and remove', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const png = makePng(4);
    const probe = wb.getImageInfo(png);
    assert.equal(probe.format, mod.ImageFormat.Png);
    assert.equal(probe.pxWidth, 4);
    assert.equal(probe.pxHeight, 1);
    assert.throws(() => wb.getImageInfo(new Uint8Array([1, 2, 3])), isFormulonError);

    const empty = wb.listDrawingObjects(0);
    assert.equal(empty.length, 0);

    const ins = wb.addImage(0, png, { name: 'logo', descr: 'alt', anchorKind: mod.AnchorKind.OneCell, row: 2, col: 3 });
    assert.ok(ins > 0);

    const check = (book) => {
      const list = book.listDrawingObjects(0);
      assert.equal(list.length, 1);
      const o = list[0];
      assert.equal(o.kind, mod.DrawingObjectKind.Picture);
      assert.equal(o.anchorKind, mod.AnchorKind.OneCell);
      assert.equal(o.fromRow, 2);
      assert.equal(o.fromCol, 3);
      assert.equal(o.cx, 4 * 9525);
      assert.equal(o.cy, 9525);
      assert.equal(o.imageFormat, mod.ImageFormat.Png);
      assert.equal(o.name, 'logo');
      assert.equal(o.descr, 'alt');
      assert.ok(o.mediaPath.length > 0);
      const img = book.getImage(0, o.objectId);
      assert.ok(img.bytes instanceof Uint8Array);
      assert.deepEqual([...img.bytes], [...png]);
      assert.equal(img.format, mod.ImageFormat.Png);
      assert.equal(img.pxWidth, 4);
      assert.equal(img.pxHeight, 1);
      return o.objectId;
    };
    check(wb);

    const saved = wb.save();
    const wb2 = mod.Workbook.load(saved ?? saved);
    try {
      const id = check(wb2);
      wb2.removeDrawing(0, id);
      assert.equal(wb2.listDrawingObjects(0).length, 0);
      assert.equal(wb2.getImage(0, id), null);
      assert.throws(() => wb2.removeDrawing(0, id), isFormulonError);
    } finally {
      wb2.dispose();
    }

    // `opts` is optional: defaults place the picture at A1.
    const plain = wb.addImage(0, png);
    assert.notEqual(plain, ins);
    assert.throws(
      () => wb.getImage(99, 1),
      (e) => isFormulonError(e, mod.StatusCode.SheetIndexOutOfRange),
    );
    assert.throws(() => wb.addImage(0, new Uint8Array([1, 2, 3]), {}), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('setDrawingAnchor moves and resizes a picture and keeps its id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const ins = wb.addImage(0, makePng(4), { name: 'logo', row: 1, col: 1 });
    wb.setDrawingAnchor(0, ins, {
      anchorKind: mod.AnchorKind.TwoCell,
      editAs: mod.AnchorEditAs.OneCell,
      row: 5,
      col: 3,
      rowOffsetEmu: 1000,
      colOffsetEmu: 2000,
      widthEmu: 200000,
    });
    const o = wb.listDrawingObjects(0)[0];
    assert.equal(o.objectId, ins);
    assert.equal(o.name, 'logo');
    assert.equal(o.anchorKind, mod.AnchorKind.TwoCell);
    assert.equal(o.editAs, mod.AnchorEditAs.OneCell);
    assert.equal(o.fromRow, 5);
    assert.equal(o.fromCol, 3);
    assert.equal(o.fromRowOffset, 1000);
    assert.equal(o.fromColOffset, 2000);
    assert.equal(o.cx, 200000);
    assert.equal(o.cy, 9525);
    assert.throws(() => wb.setDrawingAnchor(0, ins, { anchorKind: mod.AnchorKind.Absolute }), isFormulonError);
    assert.throws(() => wb.setDrawingAnchor(0, 9999, {}), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('setDrawingZOrder reorders pictures in listDrawingObjects order', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const a = wb.addImage(0, makePng(2), { row: 0 });
    const b = wb.addImage(0, makePng(2), { row: 3 });
    const c = wb.addImage(0, makePng(2), { row: 6 });
    const ids = () => Array.from(wb.listDrawingObjects(0), (o) => o.objectId);
    assert.deepEqual(ids(), [a, b, c]);
    wb.setDrawingZOrder(0, c, 0);
    assert.deepEqual(ids(), [c, a, b]);
    assert.throws(() => wb.setDrawingZOrder(0, c, 3), isFormulonError);
    assert.throws(() => wb.setDrawingZOrder(0, 9999, 0), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('snapshotDrawing and restoreDrawing bring back a removed, moved or copied picture', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const png = makePng(4);
    const a = wb.addImage(0, png, { name: 'A', row: 1 });
    const b = wb.addImage(0, makePng(2), { name: 'B', row: 4 });
    const snap = wb.snapshotDrawing(0, a);
    assert.ok(snap instanceof Uint8Array);
    assert.ok(snap.length > 0);
    const ids = () => Array.from(wb.listDrawingObjects(0), (o) => o.objectId);

    wb.removeDrawing(0, a);
    const restored = wb.restoreDrawing(0, snap);
    assert.equal(restored, a);
    assert.deepEqual(ids(), [a, b]);
    assert.deepEqual([...wb.getImage(0, a).bytes], [...png]);

    wb.setDrawingAnchor(0, a, { row: 9 });
    wb.setDrawingZOrder(0, a, 1);
    wb.restoreDrawing(0, snap, {});
    assert.deepEqual(ids(), [a, b]);
    assert.equal(wb.listDrawingObjects(0)[0].fromRow, 1);

    const copy = wb.restoreDrawing(0, snap, { newId: true });
    assert.notEqual(copy, a);
    assert.deepEqual(ids(), [a, b, copy]);

    assert.throws(() => wb.restoreDrawing(0, new Uint8Array([1, 2, 3, 4])), isFormulonError);
    assert.throws(() => wb.snapshotDrawing(0, 9999), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('setDrawingName renames a picture and keeps its id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const ins = wb.addImage(0, makePng(4), { name: 'logo', row: 1, col: 1 });
    wb.setDrawingName(0, ins, 'banner');
    const o = wb.listDrawingObjects(0)[0];
    assert.equal(o.objectId, ins);
    assert.equal(o.name, 'banner');
    assert.throws(() => wb.setDrawingName(0, ins, ''), isFormulonError);
    assert.throws(() => wb.setDrawingName(0, 9999, 'x'), isFormulonError);
  } finally {
    wb.dispose();
  }
});
