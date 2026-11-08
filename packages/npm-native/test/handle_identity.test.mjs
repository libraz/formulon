// Handle identity: one object, one handle, across every way of reaching it. The twelve cases are the same
// ones the other binding surfaces run.

import test from 'node:test';
import zlib from 'node:zlib';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

const notFound = (mod) => (e) => isFormulonError(e, mod.StatusCode.NotFound);
const invalidArgument = (mod) => (e) => isFormulonError(e, mod.StatusCode.InvalidArgument);

/** A valid width x 1 RGB PNG. */
function makePng(width) {
  const crcTable = Array.from({ length: 256 }, (_, n) => {
    let c = n;
    for (let k = 0; k < 8; k += 1) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    return c >>> 0;
  });
  const crc32 = (buf) => {
    let c = 0xffffffff;
    for (const b of buf) c = crcTable[(c ^ b) & 0xff] ^ (c >>> 8);
    return (c ^ 0xffffffff) >>> 0;
  };
  const chunk = (type, data) => {
    const out = Buffer.alloc(12 + data.length);
    out.writeUInt32BE(data.length, 0);
    out.write(type, 4, 'latin1');
    data.copy(out, 8);
    out.writeUInt32BE(crc32(out.subarray(4, 8 + data.length)), 8 + data.length);
    return out;
  };
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

/** A workbook whose Sheet1 holds the pivot source (Region / Amount) in A1:B3. */
function newBook(mod) {
  const wb = mod.Workbook.createDefault();
  const ws = wb.sheet(0);
  ws.setText(0, 0, 'Region');
  ws.setText(0, 1, 'Amount');
  ws.setText(1, 0, 'East');
  ws.setNumber(1, 1, 10);
  ws.setText(2, 0, 'West');
  ws.setNumber(2, 1, 30);
  return wb;
}

function addCache(wb, sheetName = 'Sheet1', ref = 'A1:B3') {
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ range: ref, sheet: sheetName });
  cache.addField('Region');
  cache.addField('Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const record = cache.addRecord();
    cache.setRecordValue(record, 0, region);
    cache.setRecordValue(record, 1, amount);
  }
  return cache;
}

function addPivot(sheet, cache, name, anchorRow, anchorCol) {
  const pivot = sheet.addPivot({ name, cacheId: cache.cacheId, anchor: { row: anchorRow, col: anchorCol } });
  pivot.addField({ sourceName: 'Region', axis: 0 });
  const amount = pivot.addField({ sourceName: 'Amount', axis: 2 });
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 });
  return pivot;
}

function addTable(wb, name, firstCol, sheet = 0) {
  return wb.sheet(sheet).addTable({
    range: { firstRow: 10, firstCol, lastRow: 12, lastCol: firstCol + 1 },
    name,
    columns: ['A', 'B'],
  });
}

function withBook(fn) {
  return async () => {
    const mod = await getModule();
    const wb = newBook(mod);
    try {
      await fn(wb, mod);
    } finally {
      wb.dispose();
    }
  };
}

test(
  'identity 1: every way of reaching an object yields the same handle',
  withBook((wb, mod) => {
    const ws = wb.sheet(0);
    const table = addTable(wb, 'Sales', 5);
    assert.equal(ws.table(0), table);
    assert.equal(wb.table('Sales'), table);
    assert.equal(ws.table(table), table);
    assert.equal(ws.listTables()[0], table);

    const cache = addCache(wb);
    assert.equal(wb.pivotCache(cache.cacheId), cache);
    assert.equal(wb.pivotCache(cache), cache);
    assert.equal(wb.listPivotCaches()[0], cache);

    const pivot = addPivot(ws, cache, 'Pivot1', 0, 4);
    assert.equal(ws.pivot(0), pivot);
    assert.equal(ws.pivot('Pivot1'), pivot);
    assert.equal(ws.listPivots()[0], pivot);
    assert.equal(pivot.getCache(), cache);

    const picture = ws.addPicture(makePng(4), { name: 'logo' });
    const chart = ws.addChart({ type: mod.ChartPlotType.Bar, range: 'Sheet1!$A$1:$B$3' });
    assert.equal(ws.drawing(picture.objectId), picture);
    assert.equal(ws.picture(picture.objectId), picture);
    assert.equal(ws.drawing(chart.objectId), chart);
    assert.equal(ws.chart(chart.objectId), chart);
    assert.equal(ws.chart(chart), chart);
    assert.deepEqual([...ws.listDrawings()], [picture, chart]);
    assert.deepEqual([...ws.listPictures()], [picture]);
    assert.deepEqual([...ws.listCharts()], [chart]);
    assert.equal(ws.listCharts()[0], ws.drawing(chart.objectId));

    const sheet = wb.sheet(0);
    assert.equal(wb.sheet('Sheet1'), sheet);
    assert.equal(wb.listSheets()[0], sheet);
  }),
);

test(
  'identity 2: removing an earlier element leaves a table or pivot pointing at the same object',
  withBook((wb) => {
    const ws = wb.sheet(0);
    const first = addTable(wb, 'Sales', 5);
    const second = addTable(wb, 'Costs', 8);
    assert.equal(second.index, 1);
    ws.removeTable(first);
    assert.equal(second.index, 0);
    assert.equal(second.name, 'Costs');
    assert.equal(ws.table(0), second);
    assert.equal(second.getProperties().name, 'Costs');

    const cache = addCache(wb);
    const p1 = addPivot(ws, cache, 'Pivot1', 0, 4);
    const p2 = addPivot(ws, cache, 'Pivot2', 20, 4);
    assert.equal(p2.index, 1);
    ws.removePivot(p1);
    assert.equal(p2.index, 0);
    assert.equal(p2.name, 'Pivot2');
    assert.equal(ws.pivot(0), p2);
    assert.equal(p2.getProperties().name, 'Pivot2');
  }),
);

test(
  'identity 3: the handle of a removed object is NotFound',
  withBook((wb, mod) => {
    const ws = wb.sheet(0);
    const table = addTable(wb, 'Sales', 5);
    ws.removeTable(table);
    assert.throws(() => table.getProperties(), notFound(mod));

    const cache = addCache(wb);
    const pivot = addPivot(ws, cache, 'Pivot1', 0, 4);
    ws.removePivot(pivot);
    assert.throws(() => pivot.getProperties(), notFound(mod));

    const spare = wb.addPivotCache();
    wb.removePivotCache(spare);
    assert.throws(() => spare.getOptions(), notFound(mod));

    const picture = ws.addPicture(makePng(2));
    ws.removeDrawing(picture);
    assert.throws(() => picture.getProperties(), notFound(mod));
    assert.throws(() => picture.getDrawingProperties(), notFound(mod));
    assert.throws(() => ws.drawing(picture.objectId), notFound(mod));
  }),
);

test(
  'identity 4: removing a sheet makes its pivot and drawing handles NotFound',
  withBook((wb, mod) => {
    const cache = addCache(wb);
    const other = wb.addSheet('Other');
    const pivot = addPivot(other, cache, 'Pivot1', 0, 0);
    const picture = other.addPicture(makePng(2));
    assert.equal(wb.listSheets().length, 2);
    wb.removeSheet(other);
    assert.throws(() => pivot.getLayout(), notFound(mod));
    assert.throws(() => picture.getProperties(), notFound(mod));
    assert.throws(() => picture.getDrawingProperties(), notFound(mod));
  }),
);

test(
  "identity 5: a row delete that removes a pivot range leaves the sheet's other pivots on the same objects",
  withBook((wb, mod) => {
    const ws = wb.sheet(0);
    const cache = addCache(wb);
    const doomed = addPivot(ws, cache, 'Pivot1', 0, 4);
    const kept = addPivot(ws, cache, 'Pivot2', 20, 4);
    const rect = doomed.getLayout().range;
    ws.deleteRows(0, rect.lastRow + 1);
    assert.throws(() => doomed.getLayout(), notFound(mod));
    assert.equal(ws.listPivots().length, 1);
    assert.equal(ws.listPivots()[0], kept);
    assert.equal(ws.pivot('Pivot2'), kept);
    assert.equal(kept.name, 'Pivot2');
    assert.equal(kept.index, 0);
    assert.equal(kept.getLayout().range.firstRow, 20 - (rect.lastRow + 1));
  }),
);

test(
  'identity 6: removing the source sheet of a cache removes its pivots and the cache handle',
  withBook((wb, mod) => {
    const ws = wb.sheet(0);
    const data = wb.addSheet('Data');
    data.setText(0, 0, 'Region');
    data.setText(0, 1, 'Amount');
    data.setText(1, 0, 'East');
    data.setNumber(1, 1, 10);
    const bound = addCache(wb, 'Data', 'A1:B2');
    const unrelated = addCache(wb, 'Sheet1', 'A1:B3');
    const boundPivot = addPivot(ws, bound, 'Bound', 0, 4);
    const otherPivot = addPivot(ws, unrelated, 'Other', 20, 4);
    wb.removeSheet(data);
    assert.throws(() => boundPivot.getLayout(), notFound(mod));
    assert.throws(() => bound.getOptions(), notFound(mod));
    assert.equal(ws.listPivots().length, 1);
    assert.equal(ws.listPivots()[0], otherPivot);
    assert.equal(otherPivot.getCache(), unrelated);
    assert.equal(wb.listPivotCaches().length, 1);
    assert.equal(wb.listPivotCaches()[0], unrelated);
  }),
);

test(
  'identity 7: copySheet leaves the existing handles alone and the copy gets handles of its own',
  withBook((wb, mod) => {
    const ws = wb.sheet(0);
    const cache = addCache(wb);
    const pivot = addPivot(ws, cache, 'Pivot1', 0, 4);
    const chart = ws.addChart({ type: mod.ChartPlotType.Bar, range: 'Sheet1!$A$1:$B$3' });
    const copied = wb.copySheet(ws, 1, 'Copy');
    assert.notEqual(copied, ws);
    assert.equal(ws.listPivots()[0], pivot);
    assert.equal(ws.listCharts()[0], chart);
    assert.equal(pivot.index, 0);
    const copyPivots = copied.listPivots();
    assert.equal(copyPivots.length, 1);
    assert.notEqual(copyPivots[0], pivot);
    assert.equal(copyPivots[0].sheet, copied);
    const copyCharts = copied.listCharts();
    assert.equal(copyCharts.length, 1);
    assert.notEqual(copyCharts[0], chart);
    assert.equal(wb.listSheets().length, 2);

    // A sheet with a slicer cannot be copied; nothing changes.
    const sales = addTable(wb, 'Sales', 5);
    const slicer = ws.addSlicer(sales, 0, { fromRow: 30, fromCol: 3, toRow: 40, toCol: 5 });
    assert.throws(
      () => wb.copySheet(ws, 2, 'Copy2'),
      (e) => isFormulonError(e, mod.StatusCode.NotImplemented),
    );
    assert.equal(wb.listSheets().length, 2);
    assert.equal(ws.listPivots()[0], pivot);
    assert.equal(ws.listSlicers()[0], slicer);
    assert.equal(ws.slicer(slicer.objectId), slicer);
  }),
);

test(
  'identity 8: a rename keeps the handle',
  withBook((wb) => {
    const ws = wb.sheet(0);
    const cache = addCache(wb);
    const pivot = addPivot(ws, cache, 'Pivot1', 0, 4);
    pivot.rename('Q');
    assert.equal(pivot.name, 'Q');
    assert.equal(ws.pivot('Q'), pivot);

    const table = addTable(wb, 'Sales', 5);
    table.rename('Revenue');
    assert.equal(table.name, 'Revenue');
    assert.equal(wb.table('Revenue'), table);
    assert.equal('displayName' in table, false);

    ws.rename('Data');
    assert.equal(ws.name, 'Data');
    assert.equal(wb.sheet('Data'), ws);
    assert.equal(wb.sheet(0), ws);
  }),
);

test(
  'identity 9: restoring a drawing under its old id revives the old handle',
  withBook((wb, mod) => {
    const ws = wb.sheet(0);
    const picture = ws.addPicture(makePng(4), { name: 'logo' });
    const id = picture.objectId;
    const snapshot = picture.snapshot();
    ws.removeDrawing(picture);
    assert.throws(() => picture.getProperties(), notFound(mod));
    const restored = ws.restoreDrawing(snapshot);
    assert.equal(restored, picture);
    assert.equal(restored.objectId, id);
    assert.equal(picture.getDrawingProperties().name, 'logo');
    assert.equal(ws.picture(id), picture);

    const chart = ws.addChart({ type: mod.ChartPlotType.Bar, range: 'Sheet1!$A$1:$B$3' });
    const chartSnapshot = chart.snapshot();
    ws.removeDrawing(chart);
    assert.equal(ws.restoreDrawing(chartSnapshot), chart);
    chart.setTitle('Back');
    assert.equal(chart.getProperties().title.text, 'Back');
  }),
);

test(
  'identity 10: an accessor of the wrong kind is InvalidArgument',
  withBook((wb, mod) => {
    const ws = wb.sheet(0);
    const shape = ws.addShape({
      presetGeometry: 'rect',
      anchorKind: mod.AnchorKind.OneCell,
      fromRow: 2,
      widthEmu: 914400,
      heightEmu: 457200,
    });
    const chart = ws.addChart({ type: mod.ChartPlotType.Bar, range: 'Sheet1!$A$1:$B$3' });
    const picture = ws.addPicture(makePng(2));
    assert.throws(() => ws.chart(shape.objectId), invalidArgument(mod));
    assert.throws(() => ws.picture(shape.objectId), invalidArgument(mod));
    assert.throws(() => ws.shape(chart.objectId), invalidArgument(mod));
    assert.throws(() => ws.slicer(picture.objectId), invalidArgument(mod));
    assert.throws(() => ws.chart(picture), invalidArgument(mod));
    // The base accessor accepts every kind and returns the class the kind selects.
    assert.equal(ws.drawing(shape.objectId), shape);
    assert.equal(ws.drawing(chart.objectId), chart);
  }),
);

test(
  'identity 11: a handle of another workbook is InvalidArgument',
  withBook((wb, mod) => {
    const other = newBook(mod);
    try {
      const ws = wb.sheet(0);
      const foreignTable = addTable(other, 'Sales', 5);
      const foreignCache = addCache(other);
      const foreignPivot = addPivot(other.sheet(0), foreignCache, 'Pivot1', 0, 4);
      const foreignPicture = other.sheet(0).addPicture(makePng(2));
      addTable(wb, 'Sales', 5);
      addCache(wb);
      assert.throws(() => ws.table(foreignTable), invalidArgument(mod));
      assert.throws(() => ws.removeTable(foreignTable), invalidArgument(mod));
      assert.throws(() => wb.pivotCache(foreignCache), invalidArgument(mod));
      assert.throws(() => wb.removePivotCache(foreignCache), invalidArgument(mod));
      assert.throws(() => ws.pivot(foreignPivot), invalidArgument(mod));
      assert.throws(() => ws.removePivot(foreignPivot), invalidArgument(mod));
      assert.throws(() => ws.drawing(foreignPicture), invalidArgument(mod));
      assert.throws(() => ws.picture(foreignPicture), invalidArgument(mod));
      assert.throws(() => ws.removeDrawing(foreignPicture), invalidArgument(mod));
      assert.equal(ws.listTables().length, 1);
    } finally {
      other.dispose();
    }
  }),
);

test(
  'identity 12: a pivot edit that fails to place keeps the handle, and a table handle survives every restructuring',
  withBook((wb, mod) => {
    const ws = wb.sheet(0);
    const cache = addCache(wb);
    const p1 = addPivot(ws, cache, 'Pivot1', 0, 4);
    const p2 = addPivot(ws, cache, 'Pivot2', 20, 4);
    assert.throws(
      () => p2.setAnchor('E1'),
      (e) => isFormulonError(e, mod.StatusCode.PreconditionFailed),
    );
    assert.equal(ws.pivot('Pivot2'), p2);
    assert.equal(p2.index, 1);
    assert.equal(ws.listPivots()[1], p2);
    assert.equal(p2.getProperties().name, 'Pivot2');
    assert.equal(ws.listPivots()[0], p1);

    // Tables: more adds, a row delete and a sheet removal elsewhere, then the removal of an earlier table.
    const before = addTable(wb, 'Before', 5);
    const target = addTable(wb, 'Target', 8);
    for (let i = 0; i < 20; i += 1) addTable(wb, `Extra${String.fromCharCode(97 + i)}`, 12 + 3 * i);
    assert.equal(wb.table('Target'), target);
    const side = wb.addSheet('Side');
    side.setText(0, 0, 'x');
    side.deleteRows(0, 1);
    assert.equal(wb.table('Target'), target);
    const doomed = wb.addSheet('Doomed');
    wb.removeSheet(doomed);
    assert.equal(wb.table('Target'), target);
    ws.removeTable(before);
    assert.equal(target.index, 0);
    assert.equal(target.name, 'Target');
    assert.equal(ws.table(0), target);
    assert.equal(ws.listTables()[0], target);
    assert.equal(target.getProperties().range.rect.firstCol, 8);
  }),
);
