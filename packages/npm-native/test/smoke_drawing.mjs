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
  assert.equal(mod.AnchorEditAs.OneCell, 0);
  assert.equal(mod.AnchorEditAs.TwoCell, 1);
  assert.equal(mod.AnchorEditAs.Absolute, 2);
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

    const empty = wb.sheet(0).listDrawingObjects();
    assert.equal(empty.length, 0);

    const ins = wb.sheet(0).addImage(png, {
      name: 'logo',
      description: 'alt',
      anchor: { anchorKind: mod.AnchorKind.OneCell, row: 2, col: 3 },
    });
    assert.ok(ins > 0);

    const check = (book) => {
      const list = book.sheet(0).listDrawingObjects();
      assert.equal(list.length, 1);
      const o = list[0];
      assert.equal(o.kind, mod.DrawingObjectKind.Picture);
      assert.equal(o.anchorKind, mod.AnchorKind.OneCell);
      assert.equal(o.fromRow, 2);
      assert.equal(o.fromCol, 3);
      assert.equal(o.widthEmu, 4 * 9525);
      assert.equal(o.heightEmu, 9525);
      assert.equal(o.imageFormat, mod.ImageFormat.Png);
      assert.equal(o.name, 'logo');
      assert.equal(o.description, 'alt');
      assert.ok(o.mediaPath.length > 0);
      const img = book.sheet(0).getImage(o.objectId);
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
      wb2.sheet(0).removeDrawingObject(id);
      assert.equal(wb2.sheet(0).listDrawingObjects().length, 0);
      assert.equal(wb2.sheet(0).getImage(id), null);
      assert.throws(() => wb2.sheet(0).removeDrawingObject(id), isFormulonError);
    } finally {
      wb2.dispose();
    }

    // The input is optional: defaults place the picture at A1.
    const plain = wb.sheet(0).addImage(png);
    assert.notEqual(plain, ins);
    assert.throws(
      () => wb.sheet(99).getImage(1),
      (e) => isFormulonError(e, mod.StatusCode.SheetIndexOutOfRange),
    );
    assert.throws(() => wb.sheet(0).addImage(new Uint8Array([1, 2, 3]), {}), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('setDrawingAnchor moves and resizes a picture and keeps its id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const ins = wb.sheet(0).addImage(makePng(4), { name: 'logo', anchor: { row: 1, col: 1 } });
    wb.sheet(0).setDrawingAnchor(ins, {
      anchorKind: mod.AnchorKind.TwoCell,
      editAs: mod.AnchorEditAs.OneCell,
      row: 5,
      col: 3,
      rowOffsetEmu: 1000,
      colOffsetEmu: 2000,
      widthEmu: 200000,
    });
    const o = wb.sheet(0).listDrawingObjects()[0];
    assert.equal(o.objectId, ins);
    assert.equal(o.name, 'logo');
    assert.equal(o.anchorKind, mod.AnchorKind.TwoCell);
    assert.equal(o.editAs, mod.AnchorEditAs.OneCell);
    assert.equal(o.fromRow, 5);
    assert.equal(o.fromCol, 3);
    assert.equal(o.fromRowOffsetEmu, 1000);
    assert.equal(o.fromColOffsetEmu, 2000);
    assert.equal(o.widthEmu, 200000);
    assert.equal(o.heightEmu, 9525);
    assert.throws(() => wb.sheet(0).setDrawingAnchor(ins, { anchorKind: mod.AnchorKind.Absolute }), isFormulonError);
    assert.throws(() => wb.sheet(0).setDrawingAnchor(9999, {}), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('setDrawingZOrder reorders pictures in listDrawingObjects order', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const a = wb.sheet(0).addImage(makePng(2), { anchor: { row: 0 } });
    const b = wb.sheet(0).addImage(makePng(2), { anchor: { row: 3 } });
    const c = wb.sheet(0).addImage(makePng(2), { anchor: { row: 6 } });
    const ids = () => Array.from(wb.sheet(0).listDrawingObjects(), (o) => o.objectId);
    assert.deepEqual(ids(), [a, b, c]);
    wb.sheet(0).setDrawingZOrder(c, 0);
    assert.deepEqual(ids(), [c, a, b]);
    assert.throws(() => wb.sheet(0).setDrawingZOrder(c, 3), isFormulonError);
    assert.throws(() => wb.sheet(0).setDrawingZOrder(9999, 0), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('snapshotDrawing and restoreDrawing bring back a removed, moved or copied picture', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const png = makePng(4);
    const a = wb.sheet(0).addImage(png, { name: 'A', anchor: { row: 1 } });
    const b = wb.sheet(0).addImage(makePng(2), { name: 'B', anchor: { row: 4 } });
    const snap = wb.sheet(0).snapshotDrawing(a);
    assert.ok(snap instanceof Uint8Array);
    assert.ok(snap.length > 0);
    const ids = () => Array.from(wb.sheet(0).listDrawingObjects(), (o) => o.objectId);

    wb.sheet(0).removeDrawingObject(a);
    const restored = wb.sheet(0).restoreDrawing(snap);
    assert.equal(restored, a);
    assert.deepEqual(ids(), [a, b]);
    assert.deepEqual([...wb.sheet(0).getImage(a).bytes], [...png]);

    wb.sheet(0).setDrawingAnchor(a, { row: 9 });
    wb.sheet(0).setDrawingZOrder(a, 1);
    wb.sheet(0).restoreDrawing(snap, {});
    assert.deepEqual(ids(), [a, b]);
    assert.equal(wb.sheet(0).listDrawingObjects()[0].fromRow, 1);

    const copy = wb.sheet(0).restoreDrawing(snap, { newId: true });
    assert.notEqual(copy, a);
    assert.deepEqual(ids(), [a, b, copy]);

    assert.throws(() => wb.sheet(0).restoreDrawing(new Uint8Array([1, 2, 3, 4])), isFormulonError);
    assert.throws(() => wb.sheet(0).snapshotDrawing(9999), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('setDrawingName renames a picture and keeps its id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const ins = wb.sheet(0).addImage(makePng(4), { name: 'logo', anchor: { row: 1, col: 1 } });
    wb.sheet(0).setDrawingName(ins, 'banner');
    const o = wb.sheet(0).listDrawingObjects()[0];
    assert.equal(o.objectId, ins);
    assert.equal(o.name, 'banner');
    assert.throws(() => wb.sheet(0).setDrawingName(ins, ''), isFormulonError);
    assert.throws(() => wb.sheet(0).setDrawingName(9999, 'x'), isFormulonError);
  } finally {
    wb.dispose();
  }
});
