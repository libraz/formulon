import test from 'node:test';
import zlib from 'node:zlib';
import { assert, getModule } from './smoke_support.mjs';

function must(r, what) {
  const status = typeof r.ok === 'boolean' ? r : r.status;
  assert.ok(status.ok, `${what}: ${JSON.stringify(status)}`);
  return r;
}

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
    const probe = must(wb.probeImage(png), 'probeImage');
    assert.equal(probe.format, mod.ImageFormat.Png);
    assert.equal(probe.pxWidth, 4);
    assert.equal(probe.pxHeight, 1);
    assert.equal(wb.probeImage(new Uint8Array([1, 2, 3])).status.ok, false);

    const empty = wb.listDrawingObjects(0);
    assert.ok(empty.status.ok);
    assert.equal(empty.length, 0);

    const ins = must(
      wb.insertImage(0, png, { name: 'logo', descr: 'alt', anchorKind: mod.AnchorKind.OneCell, row: 2, col: 3 }),
      'insertImage',
    );
    assert.ok(ins.objectId > 0);

    const check = (book) => {
      const list = book.listDrawingObjects(0);
      assert.ok(list.status.ok, JSON.stringify(list.status));
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
      const img = must(book.getImage(0, o.objectId), 'getImage');
      assert.ok(img.bytes instanceof Uint8Array);
      assert.deepEqual([...img.bytes], [...png]);
      assert.equal(img.format, mod.ImageFormat.Png);
      assert.equal(img.pxWidth, 4);
      assert.equal(img.pxHeight, 1);
      return o.objectId;
    };
    check(wb);

    const saved = must(wb.save(), 'save');
    const wb2 = mod.Workbook.loadBytes(saved.bytes ?? saved);
    try {
      const id = check(wb2);
      must(wb2.removeImage(0, id), 'removeImage');
      assert.equal(wb2.listDrawingObjects(0).length, 0);
      assert.equal(wb2.getImage(0, id).status.ok, false);
      assert.equal(wb2.removeImage(0, id).ok, false);
    } finally {
      wb2.dispose();
    }

    // `opts` is optional: defaults place the picture at A1.
    const plain = must(wb.insertImage(0, png), 'insertImage without opts');
    assert.notEqual(plain.objectId, ins.objectId);
    assert.equal(wb.getImage(99, 1).status.ok, false);
    assert.equal(wb.insertImage(0, new Uint8Array([1, 2, 3]), {}).status.ok, false);
  } finally {
    wb.dispose();
  }
});

test('setImageAnchor moves and resizes a picture and keeps its id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const ins = must(wb.insertImage(0, makePng(4), { name: 'logo', row: 1, col: 1 }), 'insertImage');
    must(
      wb.setImageAnchor(0, ins.objectId, {
        anchorKind: mod.AnchorKind.TwoCell,
        editAs: mod.AnchorEditAs.OneCell,
        row: 5,
        col: 3,
        rowOffEmu: 1000,
        colOffEmu: 2000,
        widthEmu: 200000,
      }),
      'setImageAnchor',
    );
    const o = wb.listDrawingObjects(0)[0];
    assert.equal(o.objectId, ins.objectId);
    assert.equal(o.name, 'logo');
    assert.equal(o.anchorKind, mod.AnchorKind.TwoCell);
    assert.equal(o.editAs, mod.AnchorEditAs.OneCell);
    assert.equal(o.fromRow, 5);
    assert.equal(o.fromCol, 3);
    assert.equal(o.fromRowOff, 1000);
    assert.equal(o.fromColOff, 2000);
    assert.equal(o.cx, 200000);
    assert.equal(o.cy, 9525);
    assert.equal(wb.setImageAnchor(0, ins.objectId, { anchorKind: mod.AnchorKind.Absolute }).ok, false);
    assert.equal(wb.setImageAnchor(0, 9999, {}).ok, false);
  } finally {
    wb.dispose();
  }
});

test('setImageZOrder reorders pictures in listDrawingObjects order', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const a = must(wb.insertImage(0, makePng(2), { row: 0 }), 'insertImage a');
    const b = must(wb.insertImage(0, makePng(2), { row: 3 }), 'insertImage b');
    const c = must(wb.insertImage(0, makePng(2), { row: 6 }), 'insertImage c');
    const ids = () => Array.from(wb.listDrawingObjects(0), (o) => o.objectId);
    assert.deepEqual(ids(), [a.objectId, b.objectId, c.objectId]);
    must(wb.setImageZOrder(0, c.objectId, 0), 'setImageZOrder');
    assert.deepEqual(ids(), [c.objectId, a.objectId, b.objectId]);
    assert.equal(wb.setImageZOrder(0, c.objectId, 3).ok, false);
    assert.equal(wb.setImageZOrder(0, 9999, 0).ok, false);
  } finally {
    wb.dispose();
  }
});

test('snapshotImage and restoreImage bring back a removed, moved or copied picture', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const png = makePng(4);
    const a = must(wb.insertImage(0, png, { name: 'A', row: 1 }), 'insertImage a');
    const b = must(wb.insertImage(0, makePng(2), { name: 'B', row: 4 }), 'insertImage b');
    const snap = must(wb.snapshotImage(0, a.objectId), 'snapshotImage');
    assert.ok(snap.bytes instanceof Uint8Array);
    assert.ok(snap.bytes.length > 0);
    const ids = () => Array.from(wb.listDrawingObjects(0), (o) => o.objectId);

    must(wb.removeImage(0, a.objectId), 'removeImage');
    const restored = must(wb.restoreImage(0, snap.bytes), 'restoreImage');
    assert.equal(restored.objectId, a.objectId);
    assert.deepEqual(ids(), [a.objectId, b.objectId]);
    assert.deepEqual([...wb.getImage(0, a.objectId).bytes], [...png]);

    must(wb.setImageAnchor(0, a.objectId, { row: 9 }), 'setImageAnchor');
    must(wb.setImageZOrder(0, a.objectId, 1), 'setImageZOrder');
    must(wb.restoreImage(0, snap.bytes, {}), 'restoreImage over a moved picture');
    assert.deepEqual(ids(), [a.objectId, b.objectId]);
    assert.equal(wb.listDrawingObjects(0)[0].fromRow, 1);

    const copy = must(wb.restoreImage(0, snap.bytes, { newId: true }), 'restoreImage newId');
    assert.notEqual(copy.objectId, a.objectId);
    assert.deepEqual(ids(), [a.objectId, b.objectId, copy.objectId]);

    const bad = wb.restoreImage(0, new Uint8Array([1, 2, 3, 4]));
    assert.equal(bad.status.ok, false);
    assert.equal(bad.objectId, 0);
    const missing = wb.snapshotImage(0, 9999);
    assert.equal(missing.status.ok, false);
    assert.equal(missing.bytes.length, 0);
  } finally {
    wb.dispose();
  }
});
