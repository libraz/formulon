import test from 'node:test';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

test("Workbook.createDefault + setFormula '=1+2' + recalc -> 3", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setFormula(0, 0, 0, '=1+2');
  wb.recalc();
  const r = wb.getValue(0, 0, 0);
  assert.equal(r.kind, mod.ValueKind.Number);
  assert.equal(r.number, 3);
});

test('recalcParallel evaluates a wide DAG and reports bounded telemetry', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const width = 128;
  for (let row = 0; row < width; row += 1) {
    wb.setNumber(0, row, 0, row + 1);
    wb.setFormula(0, row, 1, `=A${row + 1}*2`);
  }

  const serial = wb.recalcParallel(1);
  assert.equal(serial.workerThreadsStarted, 0);
  assert.equal(serial.workerThreadsUsed, 0);
  assert.equal(serial.cellsEvaluated, width);
  for (let row = 0; row < width; row += 1) {
    const value = wb.getValue(0, row, 1);
    assert.equal(value.kind, mod.ValueKind.Number);
    assert.equal(value.number, (row + 1) * 2);
  }

  for (let row = 0; row < width; row += 1) {
    wb.setNumber(0, row, 0, row + 2);
  }
  const parallel = wb.recalcParallel(8);
  if (parallel.workerThreadsStarted <= 1) {
    assert.equal(parallel.workerThreadsUsed, 0);
    assert.equal(parallel.parallelSteps, 0);
    assert.ok(parallel.serialFallbackSteps > 0, `expected serial fallback: ${JSON.stringify(parallel)}`);
  } else {
    assert.ok(parallel.parallelSteps > 0, `expected parallel steps: ${JSON.stringify(parallel)}`);
    assert.ok(parallel.workerThreadsUsed > 0, `expected an effective worker: ${JSON.stringify(parallel)}`);
    assert.ok(parallel.workerThreadsStarted <= 8);
    assert.ok(parallel.workerThreadsUsed <= parallel.workerThreadsStarted);
  }
  assert.equal(parallel.cellsEvaluated, width);
  for (let row = 0; row < width; row += 1) {
    const value = wb.getValue(0, row, 1);
    assert.equal(value.kind, mod.ValueKind.Number);
    assert.equal(value.number, (row + 2) * 2);
  }

  // Shape errors are raised by the binding, before any C call.
  const assertRejected = (errorClass, input, label) => {
    assert.throws(() => (input === undefined ? wb.recalcParallel() : wb.recalcParallel(input)), errorClass, label);
  };
  assertRejected(TypeError, undefined, 'missing threadCount');
  assertRejected(TypeError, null, 'null threadCount');
  assertRejected(RangeError, NaN, 'NaN threadCount');
  assertRejected(RangeError, Infinity, 'Infinity threadCount');
  assertRejected(RangeError, -Infinity, '-Infinity threadCount');
  assertRejected(RangeError, 1.5, 'fractional threadCount');
  assertRejected(RangeError, -1, 'negative threadCount');
  // An integer above the cap is refused by the engine.
  assert.throws(
    () => wb.recalcParallel(9),
    (e) => isFormulonError(e, mod.StatusCode.InvalidArgument),
    'threadCount above cap',
  );
  wb.dispose();
});

test("evaluateFormulaArray('=SEQUENCE(2,3)') returns a 2x3 grid", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const r = wb.evaluateFormulaArray(0, 0, 0, '=SEQUENCE(2,3)');
  assert.equal(r.rows, 2);
  assert.equal(r.cols, 3);
  assert.equal(r.cells.length, 2);
  assert.equal(r.cells[0].length, 3);
  // Row-major 1..6.
  assert.equal(r.cells[0][0].kind, mod.ValueKind.Number);
  assert.equal(r.cells[0][0].number, 1);
  assert.equal(r.cells[1][2].number, 6);
});

test("evaluateFormulaArray('=1+2') reports a 1x1 scalar array", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const r = wb.evaluateFormulaArray(0, 0, 0, '=1+2');
  assert.equal(r.rows, 1);
  assert.equal(r.cols, 1);
  assert.equal(r.cells[0][0].kind, mod.ValueKind.Number);
  assert.equal(r.cells[0][0].number, 3);
});

test('Workbook.createEmpty + addSheet -> sheetCount/sheetName work', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createEmpty();
  wb.addSheet('Sheet1');
  assert.equal(wb.getSheetCount(), 1);
  const sn = wb.getSheetName(0);
  assert.equal(sn, 'Sheet1');
});

test('Unicode sheet identity rejects folded duplicate and permits casing rename', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createEmpty();
  wb.addSheet('Ä');
  wb.addSheet('Ö');
  assert.throws(() => wb.addSheet('ä'), isFormulonError);
  assert.throws(
    () => wb.renameSheet(1, 'ä'),
    (e) => isFormulonError(e) && e.code !== 0,
  );
  assert.equal(wb.getSheetName(0), 'Ä');
  assert.equal(wb.getSheetName(1), 'Ö');
  wb.renameSheet(0, 'ä');
  const sn = wb.getSheetName(0);
  assert.equal(sn, 'ä');
  assert.equal(wb.getSheetName(1), 'Ö');
  wb.dispose();
});

test('save() returns Uint8Array; loadBytes round-trips the value', async () => {
  const mod = await getModule();
  const wb1 = mod.Workbook.createDefault();
  wb1.setNumber(0, 0, 0, 42);
  const bytes = wb1.save();
  assert.ok(bytes instanceof Uint8Array, 'expected Uint8Array');
  assert.ok(bytes.length > 0, 'expected non-empty save buffer');

  const wb2 = mod.Workbook.load(bytes);
  const r = wb2.getValue(0, 0, 0);
  assert.equal(r.kind, mod.ValueKind.Number);
  assert.equal(r.number, 42);
});

test('loadBytes rejects a non-byte view', async () => {
  const mod = await getModule();
  assert.throws(() => mod.Workbook.load(new Uint16Array([1])), TypeError);
});

// -- Expanded surface coverage -----------------------------------------
// These tests exercise the methods the addon mirrors from the embind
// binding. They are deliberately shallow: each call should round-trip
// some observable state without crashing the Node process.

test('disposed is false for a live workbook', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.disposed, false);
});

test('dispose deterministically releases a workbook and is idempotent', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.disposed, false);
  assert.equal(wb.dispose(), undefined);
  assert.equal(wb.disposed, true);
  assert.throws(() => wb.getValue(0, 0, 0), isFormulonError);
  assert.equal(wb.dispose(), undefined);
});

test('memoryUsage grows with the workbook and reads zero after dispose', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();

  const empty = wb.getMemoryUsage();
  assert.ok(empty > 0, `an empty workbook still owns storage, got ${empty}`);

  // Distinct strings so the shared-string storage cannot fold them into
  // one entry and hide the growth.
  for (let row = 0; row < 2000; row += 1) {
    wb.setText(0, row, 0, `payload-${row}-${'x'.repeat(64)}`);
  }
  const filled = wb.getMemoryUsage();
  assert.ok(filled > empty, `expected growth, got ${empty} -> ${filled}`);

  wb.dispose();
  assert.equal(wb.getMemoryUsage(), 0);
});

test('memoryUsage is stable across repeated calls and survives a recalc', async () => {
  // The external-memory report is computed as a delta against the last
  // reported figure, so a repeated call on an unchanged workbook must be
  // a no-op rather than double-counting, and the operations that trigger
  // their own report must not disturb the estimate either.
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  for (let row = 0; row < 500; row += 1) {
    wb.setFormula(0, row, 0, `=${row}+1`);
  }

  const first = wb.getMemoryUsage();
  assert.equal(wb.getMemoryUsage(), first);

  wb.recalc();
  const afterRecalc = wb.getMemoryUsage();
  assert.ok(afterRecalc >= first, `recalc must not shrink the estimate: ${first} -> ${afterRecalc}`);
  assert.equal(wb.getMemoryUsage(), afterRecalc);

  wb.dispose();
});

test('Workbook factories return instances of the exported Workbook class', async () => {
  const mod = await getModule();
  const defaultBook = mod.Workbook.createDefault();
  const emptyBook = mod.Workbook.createEmpty();
  assert.ok(defaultBook instanceof mod.Workbook);
  assert.ok(emptyBook instanceof mod.Workbook);
});

test('copySheet names the copy automatically or explicitly and reports bad arguments', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setNumber(0, 0, 0, 5);
  wb.setFormula(0, 0, 1, '=A1*2');
  const auto = wb.copySheet(0, 1);
  assert.equal(auto, 1);
  assert.equal(wb.getSheetName(1), 'Sheet1 (2)');
  const named = wb.copySheet(0, 0, 'Archive');
  assert.equal(named, 0);
  assert.equal(wb.getSheetName(0), 'Archive');
  wb.recalc();
  assert.equal(wb.getValue(0, 0, 1).number, 10);
  assert.throws(
    () => wb.copySheet(9, 0),
    (e) => isFormulonError(e, 5050),
  );
  assert.throws(
    () => wb.copySheet(0, 9),
    (e) => isFormulonError(e, 5050),
  );
  assert.throws(
    () => wb.copySheet(0, 0, ''),
    (e) => isFormulonError(e, 5051),
  );
  assert.equal(wb.getSheetCount(), 3);
  wb.dispose();
});
