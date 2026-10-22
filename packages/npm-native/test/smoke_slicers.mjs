import test from 'node:test';
import { inflateRawSync } from 'node:zlib';
import { assert, getModule } from './smoke_support.mjs';

// ---- Table slicers -------------------------------------------------------

const release = (wb) => wb.dispose();

// Reads one stored or deflated entry of a ZIP archive by name; null when absent.
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

const saveBytes = (wb) => {
  const saved = wb.save();
  assert.ok(saved.status.ok, JSON.stringify(saved.status));
  return saved.bytes;
};

const salesWorkbook = (mod) => {
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setText(0, 0, 0, 'Product').ok);
  assert.ok(wb.setText(0, 0, 1, 'Amount').ok);
  assert.ok(wb.setNumber(0, 1, 1, 10).ok);
  const created = wb.createTable({ sheetIndex: 0, ref: 'A1:B2', name: 'Sales', columns: ['Product', 'Amount'] });
  assert.ok(created.status.ok, JSON.stringify(created.status));
  return wb;
};

const anchor = { fromRow: 4, fromCol: 3, toRow: 14, toCol: 5 };

test('slicers add, list, update, remove and round-trip', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  assert.equal(wb.listSlicers(0).length, 0);
  const added = wb.addSlicer(0, 0, 0, {
    ...anchor,
    caption: 'Pick a product',
    style: 'SlicerStyleLight3',
    columnCount: 2,
    sortOrder: mod.SlicerSortOrder.Descending,
    crossFilter: mod.SlicerCrossFilter.None,
  });
  assert.ok(added.status.ok, JSON.stringify(added.status));
  assert.equal(added.objectId, 2);

  const bytes = saveBytes(wb);
  const cache = readZipEntry(bytes, 'xl/slicerCaches/slicerCache1.xml');
  assert.match(cache, /sourceName="Product"/);
  assert.match(cache, /sortOrder="descending"/);
  assert.match(readZipEntry(bytes, 'xl/slicers/slicer1.xml'), /caption="Pick a product"/);
  assert.match(readZipEntry(bytes, '[Content_Types].xml'), /application\/vnd\.ms-excel\.slicer\+xml/);
  assert.match(readZipEntry(bytes, 'xl/drawings/drawing1.xml'), /<mc:Fallback>/);

  const loaded = mod.Workbook.loadBytes(bytes);
  const slicers = loaded.listSlicers(0);
  assert.ok(slicers.status.ok);
  assert.equal(slicers.length, 1);
  const s = slicers[0];
  assert.equal(s.objectId, added.objectId);
  assert.equal(s.name, 'Product');
  assert.equal(s.caption, 'Pick a product');
  assert.equal(s.cacheName, 'Slicer_Product');
  assert.equal(s.tableName, 'Sales');
  assert.equal(s.columnName, 'Product');
  assert.equal(s.style, 'SlicerStyleLight3');
  assert.equal(s.columnCount, 2);
  assert.equal(s.rowHeight, 241300);
  assert.equal(s.sortOrder, mod.SlicerSortOrder.Descending);
  assert.equal(s.crossFilter, mod.SlicerCrossFilter.None);
  assert.equal(s.anchorKind, mod.AnchorKind.TwoCell);
  assert.deepEqual([s.fromRow, s.fromCol, s.toRow, s.toCol], [4, 3, 14, 5]);

  // An update changes the named fields only; the frame follows row edits.
  assert.ok(loaded.updateSlicer(0, 0, { caption: 'Renamed', style: '', columnCount: 1, fromRow: 0, toRow: 8 }).ok);
  assert.ok(loaded.insertRows(0, 0, 3).ok);
  const moved = loaded.listSlicers(0)[0];
  assert.equal(moved.caption, 'Renamed');
  assert.equal(moved.style, '');
  assert.equal(moved.columnCount, 1);
  assert.equal(moved.sortOrder, mod.SlicerSortOrder.Descending);
  assert.deepEqual([moved.fromRow, moved.toRow], [3, 11]);

  assert.ok(loaded.removeSlicer(0, 0).ok);
  assert.equal(loaded.listSlicers(0).length, 0);
  const emptied = saveBytes(loaded);
  assert.equal(readZipEntry(emptied, 'xl/slicers/slicer1.xml'), null);
  assert.equal(readZipEntry(emptied, 'xl/slicerCaches/slicerCache1.xml'), null);
  assert.doesNotMatch(readZipEntry(emptied, 'xl/workbook.xml'), /Slicer_Product/);
  release(loaded);
  release(wb);
});

test('removing a table removes its slicers and caches', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  assert.ok(wb.addSlicer(0, 0, 0, anchor).status.ok);
  assert.ok(wb.addSlicer(0, 0, 1, anchor).status.ok);
  assert.equal(wb.listSlicers(0).length, 2);
  assert.ok(wb.removeTable(0).ok);
  assert.equal(wb.listSlicers(0).length, 0);
  const bytes = saveBytes(wb);
  for (const part of [
    'xl/slicers/slicer1.xml',
    'xl/slicerCaches/slicerCache1.xml',
    'xl/slicerCaches/slicerCache2.xml',
  ]) {
    assert.equal(readZipEntry(bytes, part), null, part);
  }
  assert.doesNotMatch(readZipEntry(bytes, 'xl/workbook.xml'), /Slicer_/);
  release(wb);
});

test('slicer arguments are validated', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  assert.equal(wb.addSlicer(0, 0, 0, {}).status.ok, false);
  assert.equal(wb.addSlicer(0, 1, 0, anchor).status.ok, false);
  assert.equal(wb.addSlicer(0, 0, 2, anchor).status.ok, false);
  assert.equal(wb.addSlicer(5, 0, 0, anchor).status.ok, false);
  assert.equal(wb.addSlicer(0, 0, 0, { ...anchor, toRow: 1 }).status.ok, false);
  assert.equal(wb.addSlicer(0, 0, 0, { ...anchor, columnCount: 0 }).status.ok, false);
  assert.equal(wb.addSlicer(0, 0, 0, { ...anchor, sortOrder: 5 }).status.ok, false);
  assert.equal(wb.listSlicers(0).length, 0);
  assert.ok(wb.addSlicer(0, 0, 0, anchor).status.ok);
  assert.equal(wb.updateSlicer(0, 3, { caption: 'x' }).ok, false);
  assert.equal(wb.updateSlicer(0, 0, { name: 'x' }).ok, false);
  assert.equal(wb.removeSlicer(0, 3).ok, false);
  assert.throws(() => wb.addSlicer(0, 0, 0, { ...anchor, fromRow: -1 }), RangeError);
  release(wb);
});
