import test from 'node:test';
import { assert, getModule } from './smoke_support.mjs';

test("Workbook.createDefault + setFormula '=1+2' + recalc -> 3", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setFormula(0, 0, 0, '=1+2').ok);
  assert.ok(wb.recalc().ok);
  const r = wb.getValue(0, 0, 0);
  assert.ok(r.status.ok);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 3);
});

test('recalcParallel evaluates a wide DAG and reports bounded telemetry', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const width = 128;
  for (let row = 0; row < width; row += 1) {
    assert.ok(wb.setNumber(0, row, 0, row + 1).ok);
    assert.ok(wb.setFormula(0, row, 1, `=A${row + 1}*2`).ok);
  }

  const serial = wb.recalcParallel(1);
  assert.ok(serial.status.ok, `recalcParallel(1): ${JSON.stringify(serial.status)}`);
  assert.equal(serial.stats.workerThreadsStarted, 0);
  assert.equal(serial.stats.workerThreadsUsed, 0);
  assert.equal(serial.stats.cellsEvaluated, width);
  for (let row = 0; row < width; row += 1) {
    const value = wb.getValue(0, row, 1);
    assert.ok(value.status.ok, `serial value row ${row}: ${JSON.stringify(value.status)}`);
    assert.equal(value.value.kind, mod.ValueKind.Number);
    assert.equal(value.value.number, (row + 1) * 2);
  }

  for (let row = 0; row < width; row += 1) {
    assert.ok(wb.setNumber(0, row, 0, row + 2).ok);
  }
  const parallel = wb.recalcParallel(8);
  assert.ok(parallel.status.ok, `recalcParallel(8): ${JSON.stringify(parallel.status)}`);
  if (parallel.stats.workerThreadsStarted <= 1) {
    assert.equal(parallel.stats.workerThreadsUsed, 0);
    assert.equal(parallel.stats.parallelSteps, 0);
    assert.ok(parallel.stats.serialFallbackSteps > 0, `expected serial fallback: ${JSON.stringify(parallel)}`);
  } else {
    assert.ok(parallel.stats.parallelSteps > 0, `expected parallel steps: ${JSON.stringify(parallel)}`);
    assert.ok(parallel.stats.workerThreadsUsed > 0, `expected an effective worker: ${JSON.stringify(parallel)}`);
    assert.ok(parallel.stats.workerThreadsStarted <= 8);
    assert.ok(parallel.stats.workerThreadsUsed <= parallel.stats.workerThreadsStarted);
  }
  assert.equal(parallel.stats.cellsEvaluated, width);
  for (let row = 0; row < width; row += 1) {
    const value = wb.getValue(0, row, 1);
    assert.ok(value.status.ok, `parallel value row ${row}: ${JSON.stringify(value.status)}`);
    assert.equal(value.value.kind, mod.ValueKind.Number);
    assert.equal(value.value.number, (row + 2) * 2);
  }

  const assertInvalid = (input, label) => {
    const invalid = input === undefined ? wb.recalcParallel() : wb.recalcParallel(input);
    assert.equal(invalid.status.ok, false, label);
    assert.notEqual(invalid.status.status, 0, label);
    assert.equal(invalid.stats.cellsEvaluated, 0, label);
    assert.equal(invalid.stats.sccsProcessed, 0, label);
    assert.equal(invalid.stats.parallelSteps, 0, label);
    assert.equal(invalid.stats.serialFallbackSteps, 0, label);
    assert.equal(invalid.stats.cycleRecoveries, 0, label);
    assert.equal(invalid.stats.workerThreadsStarted, 0, label);
    assert.equal(invalid.stats.workerThreadsUsed, 0, label);
    return invalid;
  };
  assertInvalid(undefined, 'missing threadCount');
  assertInvalid(null, 'null threadCount');
  assertInvalid(NaN, 'NaN threadCount');
  assertInvalid(Infinity, 'Infinity threadCount');
  assertInvalid(-Infinity, '-Infinity threadCount');
  assertInvalid(1.5, 'fractional threadCount');
  assertInvalid(-1, 'negative threadCount');
  // Rejected by the binding itself, before any C ABI call: the message is
  // binding-authored and the context is empty, never a value built from
  // the threadCount the caller passed (the Status contract's promise for
  // a binding-raised failure).
  const invalid = assertInvalid(9, 'threadCount above cap');
  assert.equal(invalid.status.message, 'recalcParallel: `threadCount` must be an integer in 0..8');
  assert.equal(invalid.status.context, '');
  wb.dispose();
});

test("evaluateFormulaArray('=SEQUENCE(2,3)') returns a 2x3 grid", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const r = wb.evaluateFormulaArray(0, 0, 0, '=SEQUENCE(2,3)');
  assert.ok(r.status.ok);
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
  assert.ok(r.status.ok);
  assert.equal(r.rows, 1);
  assert.equal(r.cols, 1);
  assert.equal(r.cells[0][0].kind, mod.ValueKind.Number);
  assert.equal(r.cells[0][0].number, 3);
});

test('Workbook.createEmpty + addSheet -> sheetCount/sheetName work', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createEmpty();
  assert.ok(wb.addSheet('Sheet1').ok);
  assert.equal(wb.sheetCount().value, 1);
  const sn = wb.sheetName(0);
  assert.ok(sn.status.ok);
  assert.equal(sn.value, 'Sheet1');
});

test('Unicode sheet identity rejects folded duplicate and permits casing rename', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createEmpty();
  assert.ok(wb.addSheet('Ä').ok);
  assert.ok(wb.addSheet('Ö').ok);
  assert.equal(wb.addSheet('ä').ok, false);
  const collision = wb.renameSheet(1, 'ä');
  assert.equal(collision.ok, false);
  assert.notEqual(collision.status, 0);
  assert.equal(wb.sheetName(0).value, 'Ä');
  assert.equal(wb.sheetName(1).value, 'Ö');
  assert.ok(wb.renameSheet(0, 'ä').ok);
  const sn = wb.sheetName(0);
  assert.ok(sn.status.ok);
  assert.equal(sn.value, 'ä');
  assert.equal(wb.sheetName(1).value, 'Ö');
  wb.dispose();
});

test('save() returns Uint8Array; loadBytes round-trips the value', async () => {
  const mod = await getModule();
  const wb1 = mod.Workbook.createDefault();
  assert.ok(wb1.setNumber(0, 0, 0, 42).ok);
  const sr = wb1.save();
  assert.ok(sr.status.ok, `save: ${JSON.stringify(sr.status)}`);
  assert.ok(sr.bytes instanceof Uint8Array, 'expected Uint8Array');
  assert.ok(sr.bytes.length > 0, 'expected non-empty save buffer');

  const wb2 = mod.Workbook.loadBytes(sr.bytes);
  const r = wb2.getValue(0, 0, 0);
  assert.ok(r.status.ok, `loaded getValue: ${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 42);
});

test('loadBytes rejects invalid input with a fresh diagnostic', async () => {
  const mod = await getModule();
  const loaded = mod.Workbook.loadBytes(new Uint16Array([1]));
  assert.equal(loaded.isValid(), false);
  assert.match(mod.lastErrorMessage(), /NULL or empty input/);
});

// -- Expanded surface coverage -----------------------------------------
// These tests exercise the methods the addon mirrors from the embind
// binding. They are deliberately shallow: each call should round-trip
// some observable state without crashing the Node process.

test('isValid returns true for a live workbook', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.isValid(), true);
});

test('dispose deterministically releases a workbook and is idempotent', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.isValid(), true);
  assert.equal(wb.dispose(), undefined);
  assert.equal(wb.isValid(), false);
  assert.equal(wb.getValue(0, 0, 0).status.ok, false);
  assert.equal(wb.dispose(), undefined);
});

test('memoryUsage grows with the workbook and reads zero after dispose', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();

  const empty = wb.memoryUsage();
  assert.ok(empty > 0, `an empty workbook still owns storage, got ${empty}`);

  // Distinct strings so the shared-string storage cannot fold them into
  // one entry and hide the growth.
  for (let row = 0; row < 2000; row += 1) {
    wb.setText(0, row, 0, `payload-${row}-${'x'.repeat(64)}`);
  }
  const filled = wb.memoryUsage();
  assert.ok(filled > empty, `expected growth, got ${empty} -> ${filled}`);

  wb.dispose();
  assert.equal(wb.memoryUsage(), 0);
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

  const first = wb.memoryUsage();
  assert.equal(wb.memoryUsage(), first);

  assert.equal(wb.recalc().ok, true);
  const afterRecalc = wb.memoryUsage();
  assert.ok(afterRecalc >= first, `recalc must not shrink the estimate: ${first} -> ${afterRecalc}`);
  assert.equal(wb.memoryUsage(), afterRecalc);

  wb.dispose();
});

test('Workbook factories return instances of the exported Workbook class', async () => {
  const mod = await getModule();
  const defaultBook = mod.Workbook.createDefault();
  const emptyBook = mod.Workbook.createEmpty();
  assert.ok(defaultBook instanceof mod.Workbook);
  assert.ok(emptyBook instanceof mod.Workbook);
});
