import test from 'node:test';
import { inflateRawSync } from 'node:zlib';
import { assert, expectFailure, getModule, loadStagedModule } from './smoke_support.mjs';

// ---- Table slicers -------------------------------------------------------

const release = (wb) => wb.dispose();

// The enums are named exports of the package; the workbook factory is the module.
const load = async () => {
  const [mod, enums] = await Promise.all([getModule(), loadStagedModule()]);
  return { Workbook: mod.Workbook, ...enums };
};

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
  return saved;
};

const salesWorkbook = (mod) => {
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setText(0, 0, 'Product');
  wb.sheet(0).setText(0, 1, 'Amount');
  wb.sheet(0).setNumber(1, 1, 10);
  wb.addTable({
    range: { sheet: 0, rect: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 } },
    name: 'Sales',
    columns: ['Product', 'Amount'],
  });
  return wb;
};

const anchor = { fromRow: 4, fromCol: 3, toRow: 14, toCol: 5 };

test('slicers add, list, update, remove and round-trip', async () => {
  const mod = await load();
  const wb = salesWorkbook(mod);
  assert.equal(wb.sheet(0).listSlicers().length, 0);
  const added = wb.sheet(0).addSlicer(0, 0, {
    ...anchor,
    caption: 'Pick a product',
    style: 'SlicerStyleLight3',
    columnCount: 2,
    sortOrder: mod.SlicerSortOrder.Descending,
    crossFilter: mod.SlicerCrossFilter.None,
  });
  assert.equal(added, 2);

  const bytes = saveBytes(wb);
  const cache = readZipEntry(bytes, 'xl/slicerCaches/slicerCache1.xml');
  assert.match(cache, /sourceName="Product"/);
  assert.match(cache, /sortOrder="descending"/);
  assert.match(readZipEntry(bytes, 'xl/slicers/slicer1.xml'), /caption="Pick a product"/);
  assert.match(readZipEntry(bytes, '[Content_Types].xml'), /application\/vnd\.ms-excel\.slicer\+xml/);
  assert.match(readZipEntry(bytes, 'xl/drawings/drawing1.xml'), /<mc:Fallback>/);

  const loaded = mod.Workbook.load(bytes);
  const slicers = loaded.sheet(0).listSlicers();
  assert.equal(slicers.length, 1);
  const s = slicers[0];
  assert.equal(s.objectId, added);
  assert.equal(s.name, 'Product');
  assert.equal(s.caption, 'Pick a product');
  assert.equal(s.cacheName, 'Slicer_Product');
  assert.equal(s.tableName, 'Sales');
  assert.equal(s.columnName, 'Product');
  assert.equal(s.style, 'SlicerStyleLight3');
  assert.equal(s.columnCount, 2);
  assert.equal(s.rowHeightEmu, 241300);
  assert.equal(s.sortOrder, mod.SlicerSortOrder.Descending);
  assert.equal(s.crossFilter, mod.SlicerCrossFilter.None);
  assert.equal(s.anchorKind, mod.AnchorKind.TwoCell);
  assert.deepEqual([s.fromRow, s.fromCol, s.toRow, s.toCol], [4, 3, 14, 5]);

  // An update changes the named fields only; the frame follows row edits.
  loaded.sheet(0).updateSlicer(added, { caption: 'Renamed', style: '', columnCount: 1, fromRow: 0, toRow: 8 });
  const found = loaded.sheet(0).getSlicer(added);
  assert.equal(found.caption, 'Renamed');
  loaded.sheet(0).insertRows(0, 3);
  const moved = loaded.sheet(0).listSlicers()[0];
  assert.equal(moved.caption, 'Renamed');
  assert.equal(moved.style, null);
  assert.equal(moved.columnCount, 1);
  assert.equal(moved.sortOrder, mod.SlicerSortOrder.Descending);
  assert.deepEqual([moved.fromRow, moved.toRow], [3, 11]);

  loaded.sheet(0).removeSlicer(added);
  assert.equal(loaded.sheet(0).listSlicers().length, 0);
  const gone = loaded.sheet(0).getSlicer(added);
  assert.equal(gone, null);
  const emptied = saveBytes(loaded);
  assert.equal(readZipEntry(emptied, 'xl/slicers/slicer1.xml'), null);
  assert.equal(readZipEntry(emptied, 'xl/slicerCaches/slicerCache1.xml'), null);
  assert.doesNotMatch(readZipEntry(emptied, 'xl/workbook.xml'), /Slicer_Product/);
  release(loaded);
  release(wb);
});

test('removing a table removes its slicers and caches', async () => {
  const mod = await load();
  const wb = salesWorkbook(mod);
  wb.sheet(0).addSlicer(0, 0, anchor);
  wb.sheet(0).addSlicer(0, 1, anchor);
  assert.equal(wb.sheet(0).listSlicers().length, 2);
  wb.removeTable(0);
  assert.equal(wb.sheet(0).listSlicers().length, 0);
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
  const mod = await load();
  const wb = salesWorkbook(mod);
  expectFailure(() => wb.sheet(0).addSlicer(0, 0, {}));
  expectFailure(() => wb.sheet(0).addSlicer(1, 0, anchor));
  expectFailure(() => wb.sheet(0).addSlicer(0, 2, anchor));
  expectFailure(() => wb.sheet(5).addSlicer(0, 0, anchor));
  expectFailure(() => wb.sheet(0).addSlicer(0, 0, { ...anchor, toRow: 1 }));
  expectFailure(() => wb.sheet(0).addSlicer(0, 0, { ...anchor, columnCount: 0 }));
  expectFailure(() => wb.sheet(0).addSlicer(0, 0, { ...anchor, sortOrder: 5 }));
  assert.equal(wb.sheet(0).listSlicers().length, 0);
  const added = wb.sheet(0).addSlicer(0, 0, anchor);
  expectFailure(() => wb.sheet(0).updateSlicer(added + 1, { caption: 'x' }));
  expectFailure(() => wb.sheet(0).updateSlicer(added, { name: 'x' }));
  expectFailure(() => wb.sheet(0).removeSlicer(added + 1));
  expectFailure(() => wb.sheet(5).getSlicer(added));
  assert.throws(() => wb.sheet(0).addSlicer(0, 0, { ...anchor, fromRow: -1 }), RangeError);
  release(wb);
});
