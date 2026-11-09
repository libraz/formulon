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
  assert.equal(mod.AnchorKind.Default, 0);
  assert.equal(mod.AnchorKind.OneCell, 1);
  assert.equal(mod.AnchorKind.TwoCell, 2);
  assert.equal(mod.AnchorKind.Absolute, 3);
  assert.equal(mod.AnchorEditAs.Default, 0);
  assert.equal(mod.AnchorEditAs.OneCell, 1);
  assert.equal(mod.AnchorEditAs.TwoCell, 2);
  assert.equal(mod.AnchorEditAs.Absolute, 3);
});

test('images probe, insert, list, read back, survive save and load, and remove', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const png = makePng(4);
    const probe = mod.getImageInfo(png);
    assert.equal(probe.format, mod.ImageFormat.Png);
    assert.equal(probe.widthPx, 4);
    assert.equal(probe.heightPx, 1);
    assert.throws(() => mod.getImageInfo(new Uint8Array([1, 2, 3])), isFormulonError);

    const empty = wb.sheet(0).listDrawings();
    assert.equal(empty.length, 0);

    const ins = wb.sheet(0).addPicture(png, {
      name: 'logo',
      description: 'alt',
      anchor: { kind: mod.AnchorKind.OneCell, fromRow: 2, fromCol: 3 },
    });
    assert.ok(ins.objectId > 0);
    assert.equal(ins.kind, mod.DrawingObjectKind.Picture);

    const check = (book) => {
      const list = book.sheet(0).listDrawings();
      assert.equal(list.length, 1);
      const o = list[0].getDrawingProperties();
      assert.equal(o.kind, mod.DrawingObjectKind.Picture);
      assert.equal(o.anchor.kind, mod.AnchorKind.OneCell);
      assert.equal(o.anchor.fromRow, 2);
      assert.equal(o.anchor.fromCol, 3);
      assert.equal(o.anchor.widthEmu, 4 * 9525);
      assert.equal(o.anchor.heightEmu, 9525);
      assert.equal('format' in o, false);
      assert.equal(o.name, 'logo');
      assert.equal(list[0].name, 'logo');
      assert.equal(o.description, 'alt');
      assert.equal(o.hidden, false);
      assert.equal('mediaPath' in o, false);
      assert.equal('imageFormat' in o, false);
      const img = list[0].getProperties();
      assert.ok(img.bytes instanceof Uint8Array);
      assert.deepEqual([...img.bytes], [...png]);
      assert.equal(img.format, mod.ImageFormat.Png);
      assert.equal(img.widthPx, 4);
      assert.equal(img.heightPx, 1);
      return list[0];
    };
    check(wb);

    const saved = wb.save();
    const wb2 = mod.Workbook.load(saved ?? saved);
    try {
      const id = check(wb2);
      wb2.sheet(0).removeDrawing(id);
      assert.equal(wb2.sheet(0).listDrawings().length, 0);
      assert.throws(
        () => id.getProperties(),
        (e) => isFormulonError(e, mod.StatusCode.NotFound),
      );
      assert.throws(() => wb2.sheet(0).removeDrawing(id), isFormulonError);
    } finally {
      wb2.dispose();
    }

    // The input is optional: defaults place the picture at A1.
    const plain = wb.sheet(0).addPicture(png);
    assert.notEqual(plain.objectId, ins.objectId);
    assert.throws(
      () => wb.sheet(99).picture(1).getProperties(),
      (e) => isFormulonError(e, mod.StatusCode.SheetIndexOutOfRange),
    );
    assert.throws(() => wb.sheet(0).addPicture(new Uint8Array([1, 2, 3]), {}), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('Drawing.update moves and resizes a picture and keeps its id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const ins = wb.sheet(0).addPicture(makePng(4), { name: 'logo', anchor: { fromRow: 1, fromCol: 1 } });
    wb.sheet(0)
      .picture(ins)
      .update({
        anchor: {
          kind: mod.AnchorKind.TwoCell,
          editAs: mod.AnchorEditAs.OneCell,
          fromRow: 5,
          fromCol: 3,
          fromRowOffsetEmu: 1000,
          fromColOffsetEmu: 2000,
          toRow: 9,
          toCol: 6,
        },
      });
    const listed = wb.sheet(0).listDrawings()[0];
    assert.equal(listed, ins);
    const o = listed.getDrawingProperties();
    assert.equal(o.name, 'logo');
    assert.equal(o.anchor.kind, mod.AnchorKind.TwoCell);
    assert.equal(o.anchor.editAs, mod.AnchorEditAs.OneCell);
    assert.equal(o.anchor.fromRow, 5);
    assert.equal(o.anchor.fromCol, 3);
    assert.equal(o.anchor.fromRowOffsetEmu, 1000);
    assert.equal(o.anchor.fromColOffsetEmu, 2000);
    assert.equal(o.anchor.toRow, 9);
    assert.equal(o.anchor.toCol, 6);
    // A key omitted from an anchor patch keeps its value.
    wb.sheet(0)
      .picture(ins)
      .update({ anchor: { fromRow: 6 } });
    const kept = wb.sheet(0).picture(ins).getDrawingProperties().anchor;
    assert.equal(kept.kind, mod.AnchorKind.TwoCell);
    assert.deepEqual([kept.fromRow, kept.fromCol, kept.toRow, kept.toCol], [6, 3, 9, 6]);
    assert.throws(
      () =>
        wb
          .sheet(0)
          .picture(ins)
          .update({ anchor: { kind: mod.AnchorKind.Absolute } }),
      isFormulonError,
    );
    assert.throws(() => wb.sheet(0).picture(9999).update({ anchor: {} }), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('setDrawingZOrder reorders pictures in listDrawings order', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const a = wb.sheet(0).addPicture(makePng(2), { anchor: { fromRow: 0 } });
    const b = wb.sheet(0).addPicture(makePng(2), { anchor: { fromRow: 3 } });
    const c = wb.sheet(0).addPicture(makePng(2), { anchor: { fromRow: 6 } });
    const ids = () => Array.from(wb.sheet(0).listDrawings(), (o) => o.objectId);
    assert.deepEqual(ids(), [a.objectId, b.objectId, c.objectId]);
    wb.sheet(0).picture(c).setZOrder(0);
    assert.deepEqual(ids(), [c.objectId, a.objectId, b.objectId]);
    assert.throws(() => wb.sheet(0).picture(c).setZOrder(3), isFormulonError);
    assert.throws(() => wb.sheet(0).picture(9999).setZOrder(0), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('snapshotDrawing and restoreDrawing bring back a removed, moved or copied picture', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const png = makePng(4);
    const a = wb.sheet(0).addPicture(png, { name: 'A', anchor: { fromRow: 1 } });
    const b = wb.sheet(0).addPicture(makePng(2), { name: 'B', anchor: { fromRow: 4 } });
    const snap = wb.sheet(0).picture(a).snapshot();
    assert.ok(snap instanceof Uint8Array);
    assert.ok(snap.length > 0);
    const ids = () => Array.from(wb.sheet(0).listDrawings(), (o) => o.objectId);

    wb.sheet(0).removeDrawing(a);
    const restored = wb.sheet(0).restoreDrawing(snap);
    assert.equal(restored, a);
    assert.deepEqual(ids(), [a.objectId, b.objectId]);
    assert.deepEqual([...wb.sheet(0).picture(a).getProperties().bytes], [...png]);

    wb.sheet(0)
      .picture(a)
      .update({ anchor: { fromRow: 9 } });
    wb.sheet(0).picture(a).setZOrder(1);
    wb.sheet(0).restoreDrawing(snap, {});
    assert.deepEqual(ids(), [a.objectId, b.objectId]);
    assert.equal(wb.sheet(0).listDrawings()[0].getDrawingProperties().anchor.fromRow, 1);

    const copy = wb.sheet(0).restoreDrawing(snap, { assignNewId: true });
    assert.notEqual(copy.objectId, a.objectId);
    assert.deepEqual(ids(), [a.objectId, b.objectId, copy.objectId]);

    assert.throws(() => wb.sheet(0).restoreDrawing(new Uint8Array([1, 2, 3, 4])), isFormulonError);
    assert.throws(() => wb.sheet(0).picture(9999).snapshot(), isFormulonError);
  } finally {
    wb.dispose();
  }
});

test('Drawing.update renames a picture and keeps its id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const ins = wb.sheet(0).addPicture(makePng(4), { name: 'logo', anchor: { fromRow: 1, fromCol: 1 } });
    wb.sheet(0).picture(ins).update({ name: 'banner' });
    const listed = wb.sheet(0).listDrawings()[0];
    assert.equal(listed, ins);
    assert.equal(listed.getDrawingProperties().name, 'banner');
    assert.equal(listed.name, 'banner');
    assert.throws(() => wb.sheet(0).picture(ins).update({ name: '' }), isFormulonError);
    assert.throws(() => wb.sheet(0).picture(9999).update({ name: 'x' }), isFormulonError);
  } finally {
    wb.dispose();
  }
});
