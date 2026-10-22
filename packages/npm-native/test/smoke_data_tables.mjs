import test from 'node:test';
import { inflateRawSync } from 'node:zlib';
import { assert, getModule } from './smoke_support.mjs';

// ---- What-If data tables ------------------------------------------------

const release = (wb) => wb.dispose();

// The `ValueKind` ordinals the checks read.
const KIND = Object.freeze({ Blank: 0, Number: 1, Error: 4 });

const load = getModule;

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

const num = (wb, row, col) => {
  const r = wb.getValue(0, row, col);
  assert.ok(r.status.ok, JSON.stringify(r.status));
  assert.equal(r.value.kind, KIND.Number, `${row},${col}`);
  return r.value.number;
};

// B1 = 3 is the column input, B2 = B1 * B1 the formula; A3:A5 = 1, 2, 3 are the values, B3:B5 the results.
const squaresWorkbook = (mod) => {
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setNumber(0, 0, 1, 3).ok);
  assert.ok(wb.setFormula(0, 1, 1, '=B1*B1').ok);
  for (let i = 0; i < 3; i += 1) assert.ok(wb.setNumber(0, 2 + i, 0, i + 1).ok);
  return wb;
};

const squares = { firstRow: 2, lastRow: 4, firstCol: 1, lastCol: 1, colInput: { row: 0, col: 1 } };

test('a one-variable data table computes, reads back and refuses writes', async () => {
  const mod = await load();
  const wb = squaresWorkbook(mod);
  assert.ok(wb.addDataTable(0, squares).ok);
  assert.ok(wb.recalc().ok);
  assert.deepEqual([num(wb, 2, 1), num(wb, 3, 1), num(wb, 4, 1)], [1, 4, 9]);
  assert.equal(num(wb, 0, 1), 3);
  assert.equal(num(wb, 1, 1), 9);

  const found = wb.getDataTable(0, 3, 1);
  assert.ok(found.status.ok);
  assert.deepEqual(found.dataTable, {
    firstRow: 2,
    lastRow: 4,
    firstCol: 1,
    lastCol: 1,
    twoDimensional: false,
    rowOriented: false,
    rowInput: null,
    colInput: { row: 0, col: 1, deleted: false },
  });
  assert.equal(wb.getDataTable(0, 1, 1).dataTable, null);

  // The result range refuses writes; the values column does not.
  assert.equal(wb.setNumber(0, 3, 1, 7).ok, false);
  assert.equal(wb.setText(0, 2, 1, 'x').ok, false);
  assert.equal(wb.setFormula(0, 4, 1, '=1').ok, false);
  assert.equal(num(wb, 3, 1), 4);
  assert.ok(wb.setNumber(0, 3, 0, 10).ok);
  assert.ok(wb.recalc().ok);
  assert.equal(num(wb, 3, 1), 100);
  release(wb);
});

test('a two-variable data table fills a multiplication table', async () => {
  const mod = await load();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setNumber(0, 0, 5, 1).ok);
  assert.ok(wb.setNumber(0, 1, 5, 1).ok);
  assert.ok(wb.setFormula(0, 2, 2, '=F1*F2').ok);
  for (let k = 0; k < 3; k += 1) {
    assert.ok(wb.setNumber(0, 2, 3 + k, k + 1).ok);
    assert.ok(wb.setNumber(0, 3 + k, 2, 10 * (k + 1)).ok);
  }
  assert.ok(
    wb.addDataTable(0, {
      firstRow: 3,
      lastRow: 5,
      firstCol: 3,
      lastCol: 5,
      rowInput: { row: 0, col: 5 },
      colInput: { row: 1, col: 5 },
    }).ok,
  );
  assert.ok(wb.recalc().ok);
  assert.deepEqual([num(wb, 3, 3), num(wb, 4, 4), num(wb, 5, 5)], [10, 40, 90]);
  const table = wb.getDataTable(0, 4, 4).dataTable;
  assert.equal(table.twoDimensional, true);
  assert.deepEqual(table.rowInput, { row: 0, col: 5, deleted: false });
  assert.deepEqual(table.colInput, { row: 1, col: 5, deleted: false });
  // The inputs keep their values.
  assert.equal(num(wb, 0, 5), 1);
  assert.equal(num(wb, 1, 5), 1);
  release(wb);
});

test('a data table survives save and load as a dataTable formula and follows row edits', async () => {
  const mod = await load();
  const wb = squaresWorkbook(mod);
  assert.ok(wb.addDataTable(0, squares).ok);
  assert.ok(wb.recalc().ok);
  const saved = wb.save();
  assert.ok(saved.status.ok);
  const sheet = readZipEntry(saved.bytes, 'xl/worksheets/sheet1.xml');
  assert.match(sheet, /<f t="dataTable" ref="B3:B5"[^>]*r1="B1"[^>]*\/>/);
  assert.equal((sheet.match(/t="dataTable"/g) ?? []).length, 1);

  const loaded = mod.Workbook.loadBytes(saved.bytes);
  assert.deepEqual(loaded.getDataTable(0, 4, 1).dataTable.colInput, { row: 0, col: 1, deleted: false });
  assert.equal(loaded.setNumber(0, 4, 1, 1).ok, false);
  assert.ok(loaded.insertRows(0, 0, 2).ok);
  const moved = loaded.getDataTable(0, 6, 1).dataTable;
  assert.deepEqual([moved.firstRow, moved.lastRow], [4, 6]);
  assert.deepEqual(moved.colInput, { row: 2, col: 1, deleted: false });
  // Deleting the input's row marks it deleted and the table turns into errors.
  assert.ok(loaded.deleteRows(0, 2, 1).ok);
  assert.equal(loaded.getDataTable(0, 5, 1).dataTable.colInput.deleted, true);
  assert.ok(loaded.recalc().ok);
  assert.equal(loaded.getValue(0, 3, 1).value.kind, KIND.Error);
  release(loaded);
  release(wb);
});

test('autoNoTable leaves tables to recalcDataTables and removal blanks them', async () => {
  const mod = await load();
  const wb = squaresWorkbook(mod);
  assert.ok(wb.addDataTable(0, squares).ok);
  assert.ok(wb.setCalcMode(mod.CalcMode.AutoNoTable).ok);
  assert.ok(wb.recalc().ok);
  assert.equal(num(wb, 4, 1), 0);
  assert.equal(num(wb, 1, 1), 9);
  assert.ok(wb.recalcDataTables().ok);
  assert.equal(num(wb, 4, 1), 9);

  assert.ok(wb.removeDataTable(0, 3, 1).ok);
  assert.equal(wb.getDataTable(0, 3, 1).dataTable, null);
  assert.equal(wb.getValue(0, 3, 1).value.kind, KIND.Blank);
  assert.equal(wb.removeDataTable(0, 3, 1).ok, false);
  assert.ok(wb.setNumber(0, 3, 1, 7).ok);
  release(wb);
});

test('data table arguments are validated', async () => {
  const mod = await load();
  const wb = squaresWorkbook(mod);
  assert.equal(wb.addDataTable(0, { firstRow: 2, lastRow: 4, firstCol: 1, lastCol: 1 }).ok, false);
  assert.equal(wb.addDataTable(0, { ...squares, firstRow: 0 }).ok, false);
  assert.equal(wb.addDataTable(0, { ...squares, colInput: { row: 3, col: 1 } }).ok, false);
  assert.equal(wb.addDataTable(5, squares).ok, false);
  assert.equal(wb.getDataTable(5, 0, 0).status.ok, false);
  assert.equal(wb.removeDataTable(5, 0, 0).ok, false);
  assert.throws(() => wb.addDataTable(0, { ...squares, colInput: { row: -1, col: 1 } }), RangeError);
  assert.ok(wb.addDataTable(0, squares).ok);
  assert.equal(wb.addDataTable(0, squares).ok, false);
  release(wb);
});
