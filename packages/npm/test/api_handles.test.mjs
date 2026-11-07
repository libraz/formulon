// Handle identity (twelve cases), handle references, and the edge rows of the lower-object API:
// print getters, pivot number formats, threaded-comment `done`, sparkline replacement.
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

import { packagePart } from './zip_part.mjs';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

let formulonPromise;
function getFormulon() {
  formulonPromise ??= loadPackage().then((mod) => mod.createFormulon());
  return formulonPromise;
}

async function withWorkbook(fn) {
  const formulon = await getFormulon();
  const wb = formulon.Workbook.createDefault();
  try {
    return await fn(wb, formulon);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

function expectCode(fn, code, label) {
  assert.throws(
    fn,
    (e) => {
      assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
      assert.equal(e.code, code, `code ${e.code} (${e.message})`);
      return true;
    },
    label,
  );
}

const plain = (v) => JSON.parse(JSON.stringify(v));

// ---- Fixtures ---------------------------------------------------------------------

// Fills `data` with a Region / Amount table (A1:B3) the caches read.
function fillSource(sheet) {
  sheet.setText(0, 0, 'Region');
  sheet.setText(0, 1, 'Amount');
  sheet.setText(1, 0, 'East');
  sheet.setNumber(1, 1, 10);
  sheet.setText(2, 0, 'West');
  sheet.setNumber(2, 1, 30);
}

// A cache over `sourceSheet`'s A1:B3 with the two records filled in.
function addCache(wb, sourceSheet) {
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ present: true, ref: 'A1:B3', sheet: sourceSheet });
  cache.addField('Region');
  cache.addField('Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const rec = cache.addRecord();
    cache.setRecordValue(rec, 0, region);
    cache.setRecordValue(rec, 1, amount);
  }
  return cache;
}

// A pivot of `cache` on `sheet` anchored at (row, col): Region on rows, Sum of Amount as the data field.
function addPivot(formulon, sheet, cache, name, anchorRow, anchorCol) {
  const pivot = sheet.addPivot({ name, cacheId: cache.cacheId, anchorRow, anchorCol });
  pivot.addField({ sourceName: 'Region', axis: formulon.PivotAxis.Row });
  const amount = pivot.addField({ sourceName: 'Amount', axis: formulon.PivotAxis.Data });
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 });
  return { pivot, amount };
}

// A table of two columns at A1:B2 on `sheet`, named `name`.
function addTable(wb, sheet, name, rect = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }) {
  sheet.setText(rect.firstRow, rect.firstCol, 'Product');
  sheet.setText(rect.firstRow, rect.firstCol + 1, 'Amount');
  sheet.setNumber(rect.firstRow + 1, rect.firstCol + 1, 5);
  return wb.addTable({
    range: { sheet: sheet.index, rect },
    name,
    columns: ['Product', 'Amount'],
  });
}

// A column chart over Sheet1!A1:B3.
function addChart(formulon, sheet) {
  return sheet.addChart({
    kind: formulon.ChartInsertKind.Column,
    source: 'Sheet1!$A$1:$B$3',
    legend: formulon.ChartLegendPosition.Right,
  });
}

const SHAPE = { presetGeometry: 'rect', anchorKind: 0, fromRow: 2, widthEmu: 914400, heightEmu: 457200 };

// ---- Identity: twelve cases ---------------------------------------------------------

test('identity 1: every way to reach one target returns the same handle', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    fillSource(ws);
    const table = addTable(wb, wb.addSheet('T'), 'Tbl');
    assert.equal(wb.table('Tbl'), table);
    assert.equal(wb.table(0), table);
    assert.equal(wb.table(table), table);
    assert.equal(wb.listTables()[0], table);
    assert.equal(table.workbook, wb);

    const cache = addCache(wb, 'Sheet1');
    const { pivot } = addPivot(f, ws, cache, 'P1', 0, 4);
    assert.equal(ws.pivot('P1'), pivot);
    assert.equal(ws.pivot(0), pivot);
    assert.equal(ws.pivot(pivot), pivot);
    assert.equal(ws.listPivots()[0], pivot);
    assert.equal(pivot.getCache(), cache);
    assert.equal(wb.pivotCache(cache.cacheId), cache);
    assert.equal(wb.pivotCache(cache), cache);
    assert.equal(wb.listPivotCaches()[0], cache);

    const chart = addChart(f, ws);
    assert.equal(ws.chart(chart.objectId), chart);
    assert.equal(ws.drawing(chart.objectId), chart);
    assert.equal(ws.drawing(chart), chart);
    assert.equal(ws.listCharts()[0], chart);
    assert.equal(ws.listDrawings()[0], chart);
    const shape = ws.addShape(SHAPE);
    assert.equal(ws.shape(shape.objectId), shape);
    assert.equal(ws.drawing(shape.objectId), shape);
    assert.equal(ws.listShapes()[0], shape);
    assert.equal(ws.listDrawings().at(-1), shape);
    assert.equal(chart.kind, f.DrawingObjectKind.Chart);
    assert.equal(shape.kind, f.DrawingObjectKind.Shape);
  });
});

test('identity 2: removing an earlier table or pivot leaves the handle on its own target', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const data = wb.addSheet('Data');
    fillSource(data);
    const a = addTable(wb, ws, 'TblA', { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 });
    const b = addTable(wb, ws, 'TblB', { firstRow: 5, firstCol: 0, lastRow: 6, lastCol: 1 });
    const c = addTable(wb, ws, 'TblC', { firstRow: 10, firstCol: 0, lastRow: 11, lastCol: 1 });
    assert.deepEqual([a.index, b.index, c.index], [0, 1, 2]);
    wb.removeTable(a);
    assert.equal(b.index, 0);
    assert.equal(c.index, 1);
    assert.equal(c.name, 'TblC');
    assert.equal(wb.table('TblC'), c);
    assert.equal(wb.table(0), b);

    const cache = addCache(wb, 'Data');
    const out = wb.addSheet('Out');
    const { pivot: p1 } = addPivot(f, out, cache, 'P1', 0, 0);
    const { pivot: p2 } = addPivot(f, out, cache, 'P2', 20, 0);
    const { pivot: p3 } = addPivot(f, out, cache, 'P3', 40, 0);
    out.removePivot(p1);
    assert.equal(p2.name, 'P2');
    assert.equal(p2.index, 0);
    assert.equal(p3.index, 1);
    assert.equal(out.pivot(1), p3);
    assert.equal(p3.getProperties().name, 'P3');
  });
});

test('identity 3: the handle of a removed target throws NotFound', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    fillSource(ws);
    const table = addTable(wb, wb.addSheet('T'), 'Tbl');
    wb.removeTable(table);
    expectCode(() => table.name, f.StatusCode.NotFound);
    expectCode(() => table.index, f.StatusCode.NotFound);
    expectCode(() => table.getProperties(), f.StatusCode.NotFound);
    expectCode(() => wb.table('Tbl'), f.StatusCode.NotFound);

    const cache = addCache(wb, 'Sheet1');
    const { pivot } = addPivot(f, ws, cache, 'P1', 0, 4);
    ws.removePivot(pivot);
    expectCode(() => pivot.name, f.StatusCode.NotFound);
    expectCode(() => pivot.getProperties(), f.StatusCode.NotFound);
    expectCode(() => pivot.addField({ sourceName: 'Region', axis: f.PivotAxis.Row }), f.StatusCode.NotFound);
    expectCode(() => ws.pivot('P1'), f.StatusCode.NotFound);

    const chart = addChart(f, ws);
    const shape = ws.addShape(SHAPE);
    ws.removeDrawing(chart);
    ws.removeDrawing(shape.objectId);
    expectCode(() => chart.getDrawingProperties(), f.StatusCode.NotFound);
    expectCode(() => chart.getProperties(), f.StatusCode.NotFound);
    expectCode(() => shape.getProperties(), f.StatusCode.NotFound);
    expectCode(() => ws.drawing(chart.objectId), f.StatusCode.NotFound);

    const spare = wb.addPivotCache();
    wb.removePivotCache(spare);
    expectCode(() => spare.getOptions(), f.StatusCode.NotFound);
    expectCode(() => wb.pivotCache(spare.cacheId), f.StatusCode.NotFound);
  });
});

test('identity 4: removing a sheet orphans its pivot and drawing handles', async () => {
  await withWorkbook((wb, f) => {
    const data = wb.sheet(0);
    fillSource(data);
    const out = wb.addSheet('Out');
    const keep = wb.addSheet('Keep');
    const cache = addCache(wb, 'Sheet1');
    const { pivot } = addPivot(f, out, cache, 'P1', 0, 0);
    const { pivot: keptPivot } = addPivot(f, keep, cache, 'P2', 0, 0);
    const chart = addChart(f, out);
    const keptChart = addChart(f, keep);
    wb.removeSheet(out);
    expectCode(() => pivot.getProperties(), f.StatusCode.NotFound);
    expectCode(() => chart.getDrawingProperties(), f.StatusCode.NotFound);
    assert.equal(keptPivot.name, 'P2');
    assert.equal(keptPivot.sheet, keep);
    assert.equal(keptChart.getDrawingProperties().kind, f.DrawingObjectKind.Chart);
  });
});

test('identity 5: deleting the rows of a pivot report orphans it and no other pivot', async () => {
  await withWorkbook((wb, f) => {
    const data = wb.addSheet('Data');
    fillSource(data);
    const ws = wb.sheet(0);
    const cache = addCache(wb, 'Data');
    const { pivot: p1 } = addPivot(f, ws, cache, 'P1', 0, 0);
    const { pivot: p2 } = addPivot(f, ws, cache, 'P2', 20, 0);
    ws.deleteRows('1:10');
    expectCode(() => p1.getProperties(), f.StatusCode.NotFound);
    assert.equal(p2.getProperties().name, 'P2');
    assert.equal(p2.index, 0);
    assert.equal(ws.listPivots()[0], p2);
    assert.equal(ws.pivot('P2'), p2);
  });
});

test('identity 6: removing a source sheet orphans the pivots and the cache that read it', async () => {
  await withWorkbook((wb, f) => {
    const out = wb.sheet(0);
    const dataA = wb.addSheet('DataA');
    const dataB = wb.addSheet('DataB');
    fillSource(dataA);
    fillSource(dataB);
    const cacheA = addCache(wb, 'DataA');
    const cacheB = addCache(wb, 'DataB');
    const { pivot: pa } = addPivot(f, out, cacheA, 'PA', 0, 0);
    const { pivot: pb } = addPivot(f, out, cacheB, 'PB', 20, 0);
    wb.removeSheet(dataA);
    expectCode(() => pa.getProperties(), f.StatusCode.NotFound);
    expectCode(() => cacheA.getOptions(), f.StatusCode.NotFound);
    assert.equal(pb.getProperties().name, 'PB');
    assert.equal(pb.getCache(), cacheB);
    assert.equal(cacheB.getWorksheetSource().sheet, 'DataB');
    assert.equal(out.listPivots().length, 1);
    assert.equal(out.listPivots()[0], pb);
  });
});

test('identity 7: copySheet keeps the existing handles; a sheet with a slicer is not copied', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    fillSource(ws);
    const cache = addCache(wb, 'Sheet1');
    const { pivot } = addPivot(f, ws, cache, 'P1', 0, 4);
    const chart = addChart(f, ws);
    const copy = wb.copySheet(ws, 1, 'Copy');
    assert.notEqual(copy, ws);
    assert.equal(wb.sheet('Sheet1'), ws);
    // The original handles are undisturbed and the copy has its own.
    assert.equal(ws.pivot('P1'), pivot);
    assert.equal(ws.listCharts()[0], chart);
    assert.equal(pivot.getProperties().name, 'P1');
    const copiedPivot = copy.listPivots()[0];
    assert.notEqual(copiedPivot, pivot);
    assert.equal(copiedPivot.sheet, copy);
    assert.equal(copiedPivot.name, 'P1');
    assert.notEqual(copy.listCharts()[0], chart);
    assert.equal(copy.listCharts()[0].sheet, copy);

    // A slicer on the sheet makes the copy fail; the handles stay as they were.
    const tableSheet = wb.addSheet('Tables');
    const table = addTable(wb, tableSheet, 'Tbl');
    const slicer = tableSheet.addSlicer(table, 0, { fromRow: 4, fromCol: 3, toRow: 14, toCol: 5 });
    const sheetCount = wb.listSheets().length;
    expectCode(() => wb.copySheet(tableSheet, 0, 'TablesCopy'), f.StatusCode.NotImplemented);
    assert.equal(wb.listSheets().length, sheetCount);
    assert.equal(tableSheet.listSlicers()[0], slicer);
    assert.equal(slicer.getProperties().tableName, 'Tbl');
    assert.equal(wb.table('Tbl'), table);
  });
});

test('identity 8: a rename leaves the handle in place', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    fillSource(ws);
    const cache = addCache(wb, 'Sheet1');
    const { pivot } = addPivot(f, ws, cache, 'P1', 0, 4);
    pivot.rename('Q');
    assert.equal(pivot.name, 'Q');
    assert.equal(ws.pivot('Q'), pivot);
    expectCode(() => ws.pivot('P1'), f.StatusCode.NotFound);

    const tables = addTable(wb, wb.addSheet('T'), 'Tbl');
    tables.update({ displayName: 'Renamed' });
    assert.equal(tables.displayName, 'Renamed');
    assert.equal(wb.table('Renamed'), tables);
    tables.update({ name: 'NewName' });
    assert.equal(tables.name, 'NewName');
    assert.equal(wb.table(0), tables);

    const chart = addChart(f, ws);
    chart.rename('Revenue');
    assert.equal(ws.chart(chart.objectId), chart);
    assert.equal(chart.getDrawingProperties().name, 'Revenue');

    ws.rename('Renamed1');
    assert.equal(ws.name, 'Renamed1');
    assert.equal(wb.sheet('Renamed1'), ws);
    assert.equal(wb.sheet(0), ws);
    assert.equal(chart.sheet, ws);
  });
});

test('identity 9: restoring a drawing under its old id revives the old handle', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    fillSource(ws);
    const chart = addChart(f, ws);
    const id = chart.objectId;
    const snapshot = chart.snapshot();
    ws.removeDrawing(chart);
    expectCode(() => chart.getProperties(), f.StatusCode.NotFound);
    const restored = ws.restoreDrawing(snapshot);
    assert.equal(restored, chart);
    assert.equal(restored.objectId, id);
    assert.equal(chart.getProperties().objectId, id);
    assert.equal(ws.chart(id), chart);

    // A cache id comes back the same way.
    const cache = wb.addPivotCache(41);
    wb.removePivotCache(cache);
    expectCode(() => cache.getOptions(), f.StatusCode.NotFound);
    const again = wb.addPivotCache(41);
    assert.equal(again, cache);
    assert.doesNotThrow(() => cache.getOptions());
  });
});

test('identity 10: an accessor of the wrong kind is InvalidArgument', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    fillSource(ws);
    const shape = ws.addShape(SHAPE);
    const chart = addChart(f, ws);
    expectCode(() => ws.chart(shape.objectId), f.StatusCode.InvalidArgument);
    expectCode(() => ws.picture(shape.objectId), f.StatusCode.InvalidArgument);
    expectCode(() => ws.slicer(shape.objectId), f.StatusCode.InvalidArgument);
    expectCode(() => ws.shape(chart.objectId), f.StatusCode.InvalidArgument);
    expectCode(() => ws.picture(chart.objectId), f.StatusCode.InvalidArgument);
    // The right accessor still finds both.
    assert.equal(ws.shape(shape.objectId), shape);
    assert.equal(ws.chart(chart.objectId), chart);
  });
});

test('identity 11: a handle of another workbook is InvalidArgument', async () => {
  const formulon = await getFormulon();
  await withWorkbook((wb, f) => {
    const other = formulon.Workbook.createDefault();
    try {
      const ws = wb.sheet(0);
      const os = other.sheet(0);
      fillSource(ws);
      fillSource(os);
      const foreignTable = addTable(other, other.addSheet('T'), 'Tbl');
      const foreignCache = addCache(other, 'Sheet1');
      const { pivot: foreignPivot } = addPivot(f, os, foreignCache, 'P1', 0, 4);
      const foreignChart = addChart(f, os);
      const localTable = addTable(wb, wb.addSheet('T'), 'Tbl');
      const localCache = addCache(wb, 'Sheet1');
      addPivot(f, ws, localCache, 'P1', 0, 4);
      addChart(f, ws);

      expectCode(() => wb.table(foreignTable), f.StatusCode.InvalidArgument);
      expectCode(() => wb.removeTable(foreignTable), f.StatusCode.InvalidArgument);
      expectCode(() => ws.pivot(foreignPivot), f.StatusCode.InvalidArgument);
      expectCode(() => ws.removePivot(foreignPivot), f.StatusCode.InvalidArgument);
      expectCode(() => wb.pivotCache(foreignCache), f.StatusCode.InvalidArgument);
      expectCode(() => wb.removePivotCache(foreignCache), f.StatusCode.InvalidArgument);
      expectCode(() => ws.drawing(foreignChart), f.StatusCode.InvalidArgument);
      expectCode(() => ws.chart(foreignChart), f.StatusCode.InvalidArgument);
      expectCode(() => ws.removeDrawing(foreignChart), f.StatusCode.InvalidArgument);
      // A drawing of another sheet of the same workbook is refused too.
      const second = wb.addSheet('Second');
      expectCode(() => second.drawing(ws.listCharts()[0]), f.StatusCode.InvalidArgument);
      // The local handles are unaffected.
      assert.equal(wb.table(0), localTable);
    } finally {
      other.dispose();
    }
  });
});

test('identity 12: a failed pivot edit and table churn leave the handles on their targets', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const data = wb.addSheet('Data');
    fillSource(data);
    const cache = addCache(wb, 'Data');
    // Each report is six rows tall: moving the first onto the second cannot be placed.
    const { pivot: first } = addPivot(f, ws, cache, 'P1', 0, 0);
    const { pivot: second } = addPivot(f, ws, cache, 'P2', 7, 0);
    expectCode(() => first.setAnchor('A3:A3'), f.StatusCode.PreconditionFailed, 'moving P1 onto P2');
    assert.equal(first.name, 'P1');
    assert.equal(ws.pivot('P1'), first);
    assert.equal(ws.listPivots()[0], first);
    assert.equal(ws.listPivots()[1], second);
    assert.equal(first.getLayout().rect.firstRow, 0);
    assert.equal(second.getLayout().rect.firstRow, 7);

    // Tables survive growth of the list, row deletion and sheet removal elsewhere, and removal of an earlier one.
    const home = wb.addSheet('Home');
    const target = addTable(wb, home, 'Target', { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 });
    const early = addTable(wb, wb.addSheet('Early'), 'Early');
    for (let i = 0; i < 12; i++) addTable(wb, wb.addSheet(`S${i}`), `Extra${i}`);
    assert.equal(wb.table('Target'), target);
    assert.equal(target.name, 'Target');
    const elsewhere = wb.addSheet('Elsewhere');
    elsewhere.setNumber(0, 0, 1);
    elsewhere.deleteRows('1:1');
    assert.equal(target.name, 'Target');
    wb.removeSheet(elsewhere);
    assert.equal(target.name, 'Target');
    wb.removeTable(early);
    assert.equal(target.name, 'Target');
    assert.equal(target.getProperties().range.rect.lastCol, 1);
    assert.equal(
      wb.listTables().find((t) => t.name === 'Target'),
      target,
    );
    assert.equal(wb.listTables()[target.index], target);
  });
});

// ---- Handle references ----------------------------------------------------------------

test('handle refs: indices, names, ids and their rejected spellings', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    fillSource(ws);
    const table = addTable(wb, wb.addSheet('T'), 'Tbl');
    const cache = addCache(wb, 'Sheet1');
    const { pivot } = addPivot(f, ws, cache, 'P1', 0, 4);
    const NotFound = f.StatusCode.NotFound;
    const OutOfRange = f.StatusCode.IndexOutOfRange;

    // 0 is the first element; the last index is valid and the count is not.
    assert.equal(wb.table(0), table);
    assert.equal(wb.table(-0), table);
    assert.equal(ws.pivot(0), pivot);
    assert.equal(ws.pivot(-0), pivot);
    expectCode(() => wb.table(1), OutOfRange);
    expectCode(() => ws.pivot(1), OutOfRange);
    // 2^31 is out of range, 2^32 is outside the integer type.
    expectCode(() => wb.table(2 ** 31), OutOfRange);
    expectCode(() => ws.pivot(2 ** 31), OutOfRange);
    assert.throws(() => wb.table(2 ** 32), RangeError);
    assert.throws(() => ws.drawing(2 ** 32), RangeError);
    // Non-integers and negatives are range errors, like every other index.
    for (const bad of [NaN, Infinity, -Infinity, -1, 1.5]) {
      assert.throws(() => wb.table(bad), RangeError, `table ${bad}`);
      assert.throws(() => ws.pivot(bad), RangeError, `pivot ${bad}`);
      assert.throws(() => wb.pivotCache(bad), RangeError, `pivotCache ${bad}`);
      assert.throws(() => ws.drawing(bad), RangeError, `drawing ${bad}`);
    }
    // null, undefined and values of another type are type errors.
    for (const bad of [null, undefined, true, {}, []]) {
      assert.throws(() => wb.table(bad), TypeError, `table ${bad}`);
      assert.throws(() => ws.pivot(bad), TypeError, `pivot ${bad}`);
      assert.throws(() => wb.pivotCache(bad), TypeError, `pivotCache ${bad}`);
      assert.throws(() => ws.drawing(bad), TypeError, `drawing ${bad}`);
    }
    assert.throws(() => wb.table(), TypeError);
    assert.throws(() => ws.chart('x'), TypeError);
    // An empty name matches nothing; a name matches case-insensitively.
    expectCode(() => wb.table(''), NotFound);
    expectCode(() => ws.pivot(''), NotFound);
    assert.equal(wb.table('tbl'), table);
    assert.equal(ws.pivot('p1'), pivot);
    // An id nobody holds is NotFound.
    expectCode(() => wb.pivotCache(cache.cacheId + 100), NotFound);
    expectCode(() => ws.drawing(0), NotFound);
    expectCode(() => ws.drawing(987654), NotFound);
    // removeDrawing of the pivot cache's own slicer frame and of an unknown id is refused.
    expectCode(() => ws.removeDrawing(987654), NotFound);
  });
});

// ---- Print getters ----------------------------------------------------------------------

test('print: fitToPage false removes the attribute; the getters report the stated flags', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    assert.equal(ws.getPageSetup().fitToPageStated, false);
    ws.setPageSetup({ fitToPage: true });
    assert.equal(ws.getPageSetup().fitToPage, true);
    assert.equal(ws.getPageSetup().fitToPageStated, true);
    ws.setPageSetup({ fitToPage: false });
    assert.equal(ws.getPageSetup().fitToPage, false);
    assert.equal(ws.getPageSetup().fitToPageStated, false);

    // The print options and header / footer getters return every field, stated or not.
    const options = ws.getPrintOptions();
    assert.deepEqual(plain(options), {
      gridLines: false,
      headings: false,
      horizontalCentered: false,
      verticalCentered: false,
      gridLinesStated: false,
      headingsStated: false,
      horizontalCenteredStated: false,
      verticalCenteredStated: false,
    });
    ws.setPrintOptions({ gridLines: true, horizontalCentered: true });
    const set = ws.getPrintOptions();
    assert.equal(set.gridLines, true);
    assert.equal(set.gridLinesStated, true);
    assert.equal(set.horizontalCentered, true);
    assert.equal(set.headings, false);
    assert.equal(set.headingsStated, false);

    const empty = ws.getHeaderFooter();
    assert.equal(empty.oddHeader, null);
    assert.equal(empty.differentFirst, false);
    assert.equal(empty.differentFirstStated, false);
    ws.setHeaderFooter({ oddHeader: '&CReport', oddFooter: '&P/&N', differentFirst: true });
    const hf = ws.getHeaderFooter();
    assert.equal(hf.oddHeader, '&CReport');
    assert.equal(hf.oddFooter, '&P/&N');
    assert.equal(hf.evenHeader, null);
    assert.equal(hf.differentFirst, true);
    assert.equal(hf.differentFirstStated, true);
    ws.setHeaderFooter({ oddHeader: null });
    assert.equal(ws.getHeaderFooter().oddHeader, null);
    assert.equal(ws.getHeaderFooter().oddFooter, '&P/&N');
  });
});

// ---- Pivot number formats -------------------------------------------------------------------

const pivotPart = (wb) => packagePart(wb.save(), 'xl/pivotTables/pivotTable1.xml');

test('pivot numberFormat is a format code and the builtin id 27 survives a round trip', async () => {
  const formulon = await getFormulon();
  const wb = formulon.Workbook.createDefault();
  try {
    wb.setExcelProfileId('win-365-ja_JP');
    const ws = wb.sheet(0);
    fillSource(ws);
    const cache = addCache(wb, 'Sheet1');
    const { pivot, amount } = addPivot(formulon, ws, cache, 'P1', 0, 4);
    assert.equal(pivot.listFields()[amount].numberFormat, null);

    // The profile's code for id 27 maps back to id 27, and its read-back writes id 27 again.
    const code27 = wb.getNumFmtString(27);
    assert.ok(code27, 'the profile has a code for numFmtId 27');
    pivot.updateField(amount, { numberFormat: code27 });
    assert.match(pivotPart(wb), /<pivotField [^>]*numFmtId="27"/);
    const readBack = pivot.listFields()[amount].numberFormat;
    assert.equal(readBack, code27);
    pivot.updateField(amount, { numberFormat: readBack });
    assert.match(pivotPart(wb), /<pivotField [^>]*numFmtId="27"/);
    assert.equal(pivot.listFields()[amount].numberFormat, code27);

    // A code that is not builtin gets a custom id; a builtin code gets its builtin id.
    pivot.updateField(amount, { numberFormat: '#,##0.00' });
    assert.match(pivotPart(wb), /<pivotField [^>]*numFmtId="4"/);
    assert.equal(pivot.listFields()[amount].numberFormat, '#,##0.00');
    pivot.updateField(amount, { numberFormat: '0.0000"x"' });
    const custom = pivotPart(wb).match(/<pivotField [^>]*numFmtId="(\d+)"/);
    assert.ok(custom && Number(custom[1]) >= 164, 'a custom code gets an id from 164');
    assert.equal(pivot.listFields()[amount].numberFormat, '0.0000"x"');

    // "" and null mean no format.
    for (const none of ['', null]) {
      pivot.updateField(amount, { numberFormat: '#,##0' });
      pivot.updateField(amount, { numberFormat: none });
      assert.equal(pivot.listFields()[amount].numberFormat, null, `numberFormat ${JSON.stringify(none)}`);
      assert.doesNotMatch(pivotPart(wb), /<pivotField [^>]*numFmtId=/);
    }

    // Data fields take the same codes.
    pivot.setDataField(0, { name: 'Sum of Amount', fieldIndex: amount, aggregation: 0, numberFormat: code27 });
    assert.equal(pivot.listDataFields()[0].numberFormat, code27);
    assert.match(pivotPart(wb), /<dataField [^>]*numFmtId="27"/);
    pivot.setDataField(0, { name: 'Sum of Amount', fieldIndex: amount, aggregation: 0, numberFormat: null });
    assert.equal(pivot.listDataFields()[0].numberFormat, null);
  } finally {
    wb.dispose();
  }
});

// ---- Threaded comments --------------------------------------------------------------------------

const PERSON = '{0A1B2C3D-0000-4000-8000-000000000001}';
const THREAD = '{6F1E0D2C-1111-4111-8111-000000000001}';
const REPLY = '{6F1E0D2C-1111-4111-8111-000000000002}';

test('threaded comments: done on a thread resolves it, on a reply is ThreadedCommentInvalid, null is a TypeError', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    wb.addPerson({ id: PERSON, displayName: 'Alice', userId: 'alice', providerId: 'None' });
    ws.addThreadedComment({
      id: THREAD,
      row: 1,
      col: 1,
      personId: PERSON,
      created: '2026-01-02T03:04:05.67',
      text: 'Check',
      parentId: '',
    });
    ws.addThreadedComment({
      id: REPLY,
      row: 1,
      col: 1,
      personId: PERSON,
      created: '2026-01-02T03:05:05.67',
      text: 'Done?',
      parentId: THREAD,
    });
    assert.equal(ws.getThreadedComment(THREAD).done, false);

    ws.updateThreadedComment(THREAD, { done: true });
    assert.equal(ws.getThreadedComment(THREAD).done, true);
    assert.equal(ws.getThreadedComment(THREAD).text, 'Check');
    ws.updateThreadedComment(THREAD, { text: 'Edited' });
    assert.equal(ws.getThreadedComment(THREAD).text, 'Edited');
    assert.equal(ws.getThreadedComment(THREAD).done, true, 'a patch without done keeps it');
    ws.updateThreadedComment(THREAD, { done: false });
    assert.equal(ws.getThreadedComment(THREAD).done, false);

    expectCode(() => ws.updateThreadedComment(REPLY, { done: true }), f.StatusCode.ThreadedCommentInvalid);
    assert.equal(f.StatusCode.ThreadedCommentInvalid, 9006);
    assert.equal(ws.getThreadedComment(REPLY).done, false);
    assert.throws(() => ws.updateThreadedComment(THREAD, { done: null }), TypeError);
    assert.throws(() => ws.updateThreadedComment(THREAD, { done: 'yes' }), TypeError);
    assert.equal(ws.getThreadedComment(THREAD).done, false);
    assert.equal(typeof ws.setThreadResolved, 'undefined');
  });
});

// ---- Sparklines -------------------------------------------------------------------------------------

test('sparklines: update replaces, keeps or refuses to empty the group', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.addSparklineGroup({
      sparklines: [
        { formula: 'Sheet1!A1:D1', location: 'F1' },
        { formula: 'Sheet1!A2:D2', location: 'F2' },
      ],
    });
    const group = ws.listSparklineGroups()[0];
    const { sparklines: original } = group;
    assert.equal(original.length, 2);

    // Without the key the sparklines are kept, whatever else changes.
    const { sparklines: _unused, ...withoutKey } = group;
    void _unused;
    ws.updateSparklineGroup(0, { ...withoutKey, markers: true });
    assert.equal(ws.listSparklineGroups()[0].markers, true);
    assert.deepEqual(plain(ws.listSparklineGroups()[0].sparklines), plain(original));

    // The key replaces them.
    const replacement = [{ formula: 'Sheet1!B1:B5', location: 'H1' }];
    ws.updateSparklineGroup(0, { ...ws.listSparklineGroups()[0], sparklines: replacement });
    assert.deepEqual(plain(ws.listSparklineGroups()[0].sparklines), replacement);
    assert.equal(ws.listSparklineGroups()[0].markers, true);

    // An empty list is refused and the group is untouched.
    expectCode(
      () => ws.updateSparklineGroup(0, { ...ws.listSparklineGroups()[0], sparklines: [] }),
      f.StatusCode.InvalidArgument,
    );
    assert.deepEqual(plain(ws.listSparklineGroups()[0].sparklines), replacement);
    assert.equal(typeof ws.setSparklineGroupSparklines, 'undefined');
  });
});
