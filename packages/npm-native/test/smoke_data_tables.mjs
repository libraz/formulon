import test from 'node:test';
import { inflateRawSync } from 'node:zlib';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

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
  const r = wb.sheet(0).getTypedValue(row, col);
  assert.equal(r.kind, KIND.Number, `${row},${col}`);
  return r.number;
};

// B1 = 3 is the column input, B2 = B1 * B1 the formula; A3:A5 = 1, 2, 3 are the values, B3:B5 the results.
const squaresWorkbook = (mod) => {
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setNumber(0, 1, 3);
  wb.sheet(0).setFormula(1, 1, '=B1*B1');
  for (let i = 0; i < 3; i += 1) wb.sheet(0).setNumber(2 + i, 0, i + 1);
  return wb;
};

const squares = { rect: { firstRow: 2, firstCol: 1, lastRow: 4, lastCol: 1 }, columnInput: { row: 0, col: 1 } };

test('a one-variable data table computes, reads back and refuses writes', async () => {
  const mod = await load();
  const wb = squaresWorkbook(mod);
  wb.sheet(0).addDataTable(squares);
  wb.recalc();
  assert.deepEqual([num(wb, 2, 1), num(wb, 3, 1), num(wb, 4, 1)], [1, 4, 9]);
  assert.equal(num(wb, 0, 1), 3);
  assert.equal(num(wb, 1, 1), 9);

  const found = wb.sheet(0).getDataTable(3, 1);
  assert.deepEqual(found, {
    rect: { firstRow: 2, firstCol: 1, lastRow: 4, lastCol: 1 },
    twoDimensional: false,
    rowOriented: false,
    rowInput: null,
    columnInput: { row: 0, col: 1, deleted: false },
  });
  assert.equal(wb.sheet(0).getDataTable(1, 1), null);

  // The result range refuses writes; the values column does not.
  assert.throws(() => wb.sheet(0).setNumber(3, 1, 7), isFormulonError);
  assert.throws(() => wb.sheet(0).setText(2, 1, 'x'), isFormulonError);
  assert.throws(() => wb.sheet(0).setFormula(4, 1, '=1'), isFormulonError);
  assert.equal(num(wb, 3, 1), 4);
  wb.sheet(0).setNumber(3, 0, 10);
  wb.recalc();
  assert.equal(num(wb, 3, 1), 100);
  release(wb);
});

test('a two-variable data table fills a multiplication table', async () => {
  const mod = await load();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setNumber(0, 5, 1);
  wb.sheet(0).setNumber(1, 5, 1);
  wb.sheet(0).setFormula(2, 2, '=F1*F2');
  for (let k = 0; k < 3; k += 1) {
    wb.sheet(0).setNumber(2, 3 + k, k + 1);
    wb.sheet(0).setNumber(3 + k, 2, 10 * (k + 1));
  }
  wb.sheet(0).addDataTable({
    rect: { firstRow: 3, firstCol: 3, lastRow: 5, lastCol: 5 },
    rowInput: { row: 0, col: 5 },
    columnInput: { row: 1, col: 5 },
  });
  wb.recalc();
  assert.deepEqual([num(wb, 3, 3), num(wb, 4, 4), num(wb, 5, 5)], [10, 40, 90]);
  const table = wb.sheet(0).getDataTable(4, 4);
  assert.equal(table.twoDimensional, true);
  assert.deepEqual(table.rowInput, { row: 0, col: 5, deleted: false });
  assert.deepEqual(table.columnInput, { row: 1, col: 5, deleted: false });
  // The inputs keep their values.
  assert.equal(num(wb, 0, 5), 1);
  assert.equal(num(wb, 1, 5), 1);
  release(wb);
});

test('a data table survives save and load as a dataTable formula and follows row edits', async () => {
  const mod = await load();
  const wb = squaresWorkbook(mod);
  wb.sheet(0).addDataTable(squares);
  wb.recalc();
  const saved = wb.save();
  const sheet = readZipEntry(saved, 'xl/worksheets/sheet1.xml');
  assert.match(sheet, /<f t="dataTable" ref="B3:B5"[^>]*r1="B1"[^>]*\/>/);
  assert.equal((sheet.match(/t="dataTable"/g) ?? []).length, 1);

  const loaded = mod.Workbook.load(saved);
  assert.deepEqual(loaded.sheet(0).getDataTable(4, 1).columnInput, { row: 0, col: 1, deleted: false });
  assert.throws(() => loaded.sheet(0).setNumber(4, 1, 1), isFormulonError);
  loaded.sheet(0).insertRows(0, 2);
  const moved = loaded.sheet(0).getDataTable(6, 1);
  assert.deepEqual([moved.rect.firstRow, moved.rect.lastRow], [4, 6]);
  assert.deepEqual(moved.columnInput, { row: 2, col: 1, deleted: false });
  // Deleting the input's row marks it deleted and the table turns into errors.
  loaded.sheet(0).deleteRows(2, 1);
  assert.equal(loaded.sheet(0).getDataTable(5, 1).columnInput.deleted, true);
  loaded.recalc();
  assert.equal(loaded.sheet(0).getTypedValue(3, 1).kind, KIND.Error);
  release(loaded);
  release(wb);
});

test('autoNoTable leaves tables to recalcDataTables and removal blanks them', async () => {
  const mod = await load();
  const wb = squaresWorkbook(mod);
  wb.sheet(0).addDataTable(squares);
  wb.setCalcMode(mod.CalcMode.AutoNoTable);
  wb.recalc();
  assert.equal(num(wb, 4, 1), 0);
  assert.equal(num(wb, 1, 1), 9);
  wb.recalcDataTables();
  assert.equal(num(wb, 4, 1), 9);

  wb.sheet(0).removeDataTable(3, 1);
  assert.equal(wb.sheet(0).getDataTable(3, 1), null);
  assert.equal(wb.sheet(0).getTypedValue(3, 1).kind, KIND.Blank);
  assert.throws(() => wb.sheet(0).removeDataTable(3, 1), isFormulonError);
  wb.sheet(0).setNumber(3, 1, 7);
  release(wb);
});

test('data table arguments are validated', async () => {
  const mod = await load();
  const wb = squaresWorkbook(mod);
  assert.throws(() => wb.sheet(0).addDataTable({ rect: squares.rect }), isFormulonError);
  assert.throws(
    () => wb.sheet(0).addDataTable({ ...squares, rect: { ...squares.rect, firstRow: 0 } }),
    isFormulonError,
  );
  assert.throws(() => wb.sheet(0).addDataTable({ ...squares, columnInput: { row: 3, col: 1 } }), isFormulonError);
  assert.throws(() => wb.sheet(5).addDataTable(squares), isFormulonError);
  assert.throws(() => wb.sheet(5).getDataTable(0, 0), isFormulonError);
  assert.throws(() => wb.sheet(5).removeDataTable(0, 0), isFormulonError);
  assert.throws(() => wb.sheet(0).addDataTable({ ...squares, columnInput: { row: -1, col: 1 } }), RangeError);
  wb.sheet(0).addDataTable(squares);
  assert.throws(() => wb.sheet(0).addDataTable(squares), isFormulonError);
  release(wb);
});
