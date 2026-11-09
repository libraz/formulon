import test from 'node:test';
import { inflateRawSync } from 'node:zlib';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

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
  return saved;
};

const salesWorkbook = (mod) => {
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setText(0, 0, 'Product');
  wb.sheet(0).setText(0, 1, 'Amount');
  wb.sheet(0).setNumber(1, 1, 10);
  wb.sheet(0).addTable({
    range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 },
    name: 'Sales',
    columns: ['Product', 'Amount'],
  });
  return wb;
};

const box = { fromRow: 4, fromCol: 3, toRow: 14, toCol: 5 };
const anchor = { anchor: box };

test('slicers add, list, update, remove and round-trip', async () => {
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  assert.equal(wb.sheet(0).listSlicers().length, 0);
  const added = wb.sheet(0).addSlicer('Sales', 0, {
    ...anchor,
    caption: 'Pick a product',
    styleName: 'SlicerStyleLight3',
    columnCount: 2,
    sortOrder: mod.SlicerSortOrder.Descending,
    crossFilter: mod.SlicerCrossFilter.None,
  });
  assert.equal(added.objectId, 2);

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
  assert.equal(slicers[0].objectId, added.objectId);
  const s = slicers[0].getProperties();
  assert.equal(s.caption, 'Pick a product');
  assert.equal(s.cacheName, 'Slicer_Product');
  assert.equal(s.tableName, 'Sales');
  assert.equal(s.columnName, 'Product');
  assert.equal(s.styleName, 'SlicerStyleLight3');
  assert.equal(s.columnCount, 2);
  assert.equal(s.rowHeightPt, 19);
  assert.equal(s.sortOrder, mod.SlicerSortOrder.Descending);
  assert.equal(s.crossFilter, mod.SlicerCrossFilter.None);
  const common = slicers[0].getDrawingProperties();
  assert.equal(common.name, 'Product');
  assert.equal(slicers[0].name, 'Product');
  assert.equal(common.anchor.kind, mod.AnchorKind.TwoCell);
  assert.deepEqual(
    [common.anchor.fromRow, common.anchor.fromCol, common.anchor.toRow, common.anchor.toCol],
    [4, 3, 14, 5],
  );
  assert.equal(typeof common.anchor.editAs, 'number');
  assert.ok(common.anchor.widthEmu > 0 && common.anchor.heightEmu > 0);

  // An update changes the named fields only; the frame follows row edits.
  loaded
    .sheet(0)
    .slicer(added.objectId)
    .update({
      caption: 'Renamed',
      styleName: null,
      columnCount: 1,
      anchor: { fromRow: 0, fromCol: 3, toRow: 8, toCol: 5 },
    });
  const found = loaded.sheet(0).slicer(added.objectId).getProperties();
  assert.equal(found.caption, 'Renamed');
  loaded.sheet(0).insertRows(0, 2);
  const movedHandle = loaded.sheet(0).listSlicers()[0];
  const moved = movedHandle.getProperties();
  const movedAnchor = movedHandle.getDrawingProperties().anchor;
  assert.equal(moved.caption, 'Renamed');
  assert.equal(moved.styleName, null);
  assert.equal(moved.columnCount, 1);
  assert.equal(moved.sortOrder, mod.SlicerSortOrder.Descending);
  assert.deepEqual([movedAnchor.fromRow, movedAnchor.fromCol, movedAnchor.toRow, movedAnchor.toCol], [3, 3, 11, 5]);

  loaded.sheet(0).removeDrawing(added.objectId);
  assert.equal(loaded.sheet(0).listSlicers().length, 0);
  assert.throws(
    () => loaded.sheet(0).slicer(added.objectId),
    (e) => isFormulonError(e, mod.StatusCode.NotFound),
  );
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
  wb.sheet(0).addSlicer('Sales', 0, anchor);
  wb.sheet(0).addSlicer('Sales', 1, anchor);
  assert.equal(wb.sheet(0).listSlicers().length, 2);
  wb.sheet(0).removeTable(0);
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
  const mod = await getModule();
  const wb = salesWorkbook(mod);
  assert.throws(() => wb.sheet(0).addSlicer('Sales', 0, {}), isFormulonError);
  assert.throws(() => wb.sheet(0).addSlicer('Missing', 0, anchor), isFormulonError);
  assert.throws(() => wb.sheet(0).addSlicer('Sales', 2, anchor), isFormulonError);
  assert.throws(() => wb.sheet(5).addSlicer('Sales', 0, anchor), isFormulonError);
  assert.throws(() => wb.sheet(0).addSlicer('Sales', 0, { anchor: { ...box, toRow: 1 } }), isFormulonError);
  assert.throws(() => wb.sheet(0).addSlicer('Sales', 0, { ...anchor, columnCount: 20001 }), isFormulonError);
  assert.throws(() => wb.sheet(0).addSlicer('Sales', 0, { ...anchor, sortOrder: 5 }), isFormulonError);
  assert.equal(wb.sheet(0).listSlicers().length, 0);
  // A zero column count on add is the default.
  const defaulted = wb.sheet(0).addSlicer('Sales', 0, { ...anchor, columnCount: 0 });
  assert.equal(defaulted.getProperties().columnCount, 1);
  wb.sheet(0).removeDrawing(defaulted);
  const added = wb.sheet(0).addSlicer('Sales', 0, anchor);
  assert.throws(
    () =>
      wb
        .sheet(0)
        .slicer(added.objectId + 1)
        .update({ caption: 'x' }),
    isFormulonError,
  );
  // A slicer's name keys its cache, so the common patch refuses to set it.
  const nameBefore = wb.sheet(0).slicer(added).name;
  assert.throws(
    () => wb.sheet(0).slicer(added).update({ name: 'x' }),
    (e) => isFormulonError(e, mod.StatusCode.PreconditionFailed),
  );
  assert.equal(wb.sheet(0).slicer(added).name, nameBefore);
  assert.throws(() => wb.sheet(0).slicer(added).update({ unknownKey: 1 }), TypeError);
  assert.throws(() => wb.sheet(0).slicer(added).update({ toRow: 3 }), TypeError);
  assert.throws(() => wb.sheet(0).slicer(added).update({ fromRow: 3 }), TypeError);
  assert.throws(() => wb.sheet(0).slicer(added).update({ caption: null }), TypeError);
  assert.throws(() => wb.sheet(0).removeDrawing(added.objectId + 1), isFormulonError);
  assert.throws(() => wb.sheet(5).slicer(added).getProperties(), isFormulonError);
  assert.throws(
    () => wb.sheet(0).addSlicer('Sales', 0, { anchor: { ...box, fromRow: -1 } }),
    (e) => e.name === 'FormulonError' && e.code === mod.StatusCode.InvalidArgument,
  );
  assert.throws(() => wb.sheet(0).addSlicer('Sales', 0, { anchor: { ...box, fromRow: 1.5 } }), RangeError);
  release(wb);
});
