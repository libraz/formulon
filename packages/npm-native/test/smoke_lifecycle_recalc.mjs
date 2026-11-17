import test from 'node:test';
import { assert, getModule, isFormulonError, zipStore } from './smoke_support.mjs';

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PKG_REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';

// A1 = 1, B1 = A1+1 storing 99 and C1 = "a"&" " storing "a", in an automatic workbook with the current calcId;
// `fullCalcOnLoad` asks for a recalculation on open.
function staleCacheBook(mod, fullCalcOnLoad) {
  const calcPr = fullCalcOnLoad ? '<calcPr calcId="191029" fullCalcOnLoad="1"/>' : '<calcPr calcId="191029"/>';
  return mod.Workbook.load(
    zipStore([
      [
        '[Content_Types].xml',
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
          '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
          '</Types>',
      ],
      [
        '_rels/.rels',
        `<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
      ],
      [
        'xl/workbook.xml',
        `<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>${calcPr}</workbook>`,
      ],
      [
        'xl/_rels/workbook.xml.rels',
        `<Relationships xmlns="${PKG_REL_NS}"><Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>`,
      ],
      [
        'xl/worksheets/sheet1.xml',
        `<worksheet xmlns="${MAIN_NS}"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><f>A1+1</f><v>99</v></c>` +
          '<c r="C1" t="str"><f>"a"&amp;" "</f><v>a</v></c></row></sheetData></worksheet>',
      ],
    ]),
  );
}

test("Workbook.createDefault + setFormula '=1+2' + recalc -> 3", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setFormula(0, 0, '=1+2');
  wb.recalc();
  const r = wb.sheet(0).getTypedValue(0, 0);
  assert.equal(r.kind, mod.ValueKind.Number);
  assert.equal(r.number, 3);
});

test('recalcParallel evaluates a wide DAG and reports bounded telemetry', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const width = 128;
  for (let row = 0; row < width; row += 1) {
    wb.sheet(0).setNumber(row, 0, row + 1);
    wb.sheet(0).setFormula(row, 1, `=A${row + 1}*2`);
  }

  const serial = wb.recalcParallel(1);
  assert.equal(serial.workerThreadsStarted, 0);
  assert.equal(serial.workerThreadsUsed, 0);
  assert.equal(serial.cellsEvaluated, width);
  for (let row = 0; row < width; row += 1) {
    const value = wb.sheet(0).getTypedValue(row, 1);
    assert.equal(value.kind, mod.ValueKind.Number);
    assert.equal(value.number, (row + 1) * 2);
  }

  for (let row = 0; row < width; row += 1) {
    wb.sheet(0).setNumber(row, 0, row + 2);
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
    const value = wb.sheet(0).getTypedValue(row, 1);
    assert.equal(value.kind, mod.ValueKind.Number);
    assert.equal(value.number, (row + 2) * 2);
  }

  // Shape errors are raised by the binding, before any C call.
  const assertRejected = (errorClass, input, label) => {
    assert.throws(() => (input === undefined ? wb.recalcParallel() : wb.recalcParallel(input)), errorClass, label);
  };
  const assertInvalidArgument = (input, label) => {
    assert.throws(
      () => wb.recalcParallel(input),
      (e) => isFormulonError(e, mod.StatusCode.InvalidArgument),
      label,
    );
  };
  // An omitted threadCount takes its default (0, the hardware width); null is still a shape error.
  assert.equal(typeof wb.recalcParallel(), 'object');
  assertRejected(TypeError, null, 'null threadCount');
  assertRejected(RangeError, NaN, 'NaN threadCount');
  assertRejected(RangeError, Infinity, 'Infinity threadCount');
  assertRejected(RangeError, -Infinity, '-Infinity threadCount');
  assertRejected(RangeError, 1.5, 'fractional threadCount');
  assertRejected(TypeError, '2', 'string threadCount');
  assertInvalidArgument(-1, 'negative threadCount');
  // An integer above the cap is refused by the engine.
  assert.throws(
    () => wb.recalcParallel(9),
    (e) => isFormulonError(e, mod.StatusCode.SecurityResourceLimit),
    'threadCount above cap',
  );
  wb.dispose();
});

test("evaluateTypedFormulaArray('=SEQUENCE(2,3)') returns a 2x3 grid", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const r = wb.sheet(0).evaluateTypedFormulaArray(0, 0, '=SEQUENCE(2,3)');
  assert.equal(r.length, 2);
  assert.equal(r[0].length, 3);
  assert.equal(r[1].length, 3);
  // Row-major 1..6.
  assert.equal(r[0][0].kind, mod.ValueKind.Number);
  assert.equal(r[0][0].number, 1);
  assert.equal(r[1][2].number, 6);
});

test("evaluateTypedFormulaArray('=1+2') reports a 1x1 scalar array", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const r = wb.sheet(0).evaluateTypedFormulaArray(0, 0, '=1+2');
  assert.equal(r.length, 1);
  assert.equal(r[0].length, 1);
  assert.equal(r[0][0].kind, mod.ValueKind.Number);
  assert.equal(r[0][0].number, 3);
});

test('Workbook.createEmpty + addSheet -> sheetCount/sheetName work', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createEmpty();
  wb.addSheet('Sheet1');
  assert.equal(wb.listSheets().length, 1);
  const sn = wb.sheet(0).name;
  assert.equal(sn, 'Sheet1');
});

test('Unicode sheet identity rejects folded duplicate and permits casing rename', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createEmpty();
  wb.addSheet('Ä');
  wb.addSheet('Ö');
  assert.throws(() => wb.addSheet('ä'), isFormulonError);
  assert.throws(
    () => wb.sheet(1).rename('ä'),
    (e) => isFormulonError(e) && e.code !== 0,
  );
  assert.equal(wb.sheet(0).name, 'Ä');
  assert.equal(wb.sheet(1).name, 'Ö');
  wb.sheet(0).rename('ä');
  const sn = wb.sheet(0).name;
  assert.equal(sn, 'ä');
  assert.equal(wb.sheet(1).name, 'Ö');
  wb.dispose();
});

test('a workbook saved under the current calculation engine opens with its stored values until recalcFull', async () => {
  const mod = await getModule();
  for (const mode of [mod.CalcMode.Auto, mod.CalcMode.Manual]) {
    const wb = staleCacheBook(mod, false);
    wb.setCalcMode(mode);
    assert.equal(wb.hasDirtyCells(), false);
    assert.equal(wb.sheet(0).getTypedValue(0, 1).number, 99);
    assert.equal(wb.sheet(0).getTypedValue(0, 2).text, 'a');
    wb.recalcFull();
    assert.equal(wb.sheet(0).getTypedValue(0, 1).number, 2);
    assert.equal(wb.sheet(0).getTypedValue(0, 2).text, 'a ');
    assert.equal(wb.hasDirtyCells(), false);
    wb.dispose();
  }
  const asked = staleCacheBook(mod, true);
  assert.equal(asked.hasDirtyCells(), true);
  asked.recalc();
  assert.equal(asked.sheet(0).getTypedValue(0, 1).number, 2);
  assert.equal(asked.sheet(0).getTypedValue(0, 2).text, 'a ');
  asked.dispose();
});

test('recalcFullParallel recomputes every formula and refuses before marking anything', async () => {
  const mod = await getModule();
  const wb = staleCacheBook(mod, false);
  assert.throws(
    () => wb.recalcFullParallel(9),
    (e) => isFormulonError(e, mod.StatusCode.SecurityResourceLimit),
  );
  assert.throws(
    () => wb.recalcFullParallel(-1),
    (e) => isFormulonError(e, mod.StatusCode.InvalidArgument),
  );
  assert.throws(() => wb.recalcFullParallel(1.5), RangeError);
  assert.throws(() => wb.recalcFullParallel(null), TypeError);
  assert.equal(wb.hasDirtyCells(), false);
  assert.equal(wb.sheet(0).getTypedValue(0, 1).number, 99);
  const stats = wb.recalcFullParallel(1);
  assert.equal(stats.cellsEvaluated, 2);
  assert.equal(stats.workerThreadsStarted, 0);
  assert.equal(wb.sheet(0).getTypedValue(0, 1).number, 2);
  assert.equal(typeof wb.recalcFullParallel(), 'object');
  wb.dispose();
});

test('save() returns Uint8Array; loadBytes round-trips the value', async () => {
  const mod = await getModule();
  const wb1 = mod.Workbook.createDefault();
  wb1.sheet(0).setNumber(0, 0, 42);
  const bytes = wb1.save();
  assert.ok(bytes instanceof Uint8Array, 'expected Uint8Array');
  assert.ok(bytes.length > 0, 'expected non-empty save buffer');

  const wb2 = mod.Workbook.load(bytes);
  const r = wb2.sheet(0).getTypedValue(0, 0);
  assert.equal(r.kind, mod.ValueKind.Number);
  assert.equal(r.number, 42);
});

test('load reads any view as its bytes, so a non-workbook view is a FormulonError', async () => {
  const mod = await getModule();
  assert.throws(() => mod.Workbook.load(new Uint16Array([1])), isFormulonError);
  assert.throws(() => mod.Workbook.load('bytes'), TypeError);
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
  assert.throws(() => wb.sheet(0).getTypedValue(0, 0), isFormulonError);
  assert.equal(wb.dispose(), undefined);
});

test('memoryUsage grows with the workbook and throws after dispose', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();

  const empty = wb.getMemoryUsage();
  assert.ok(empty > 0, `an empty workbook still owns storage, got ${empty}`);

  // Distinct strings so the shared-string storage cannot fold them into
  // one entry and hide the growth.
  for (let row = 0; row < 2000; row += 1) {
    wb.sheet(0).setText(row, 0, `payload-${row}-${'x'.repeat(64)}`);
  }
  const filled = wb.getMemoryUsage();
  assert.ok(filled > empty, `expected growth, got ${empty} -> ${filled}`);

  wb.dispose();
  assert.throws(() => wb.getMemoryUsage(), isFormulonError);
});

test('memoryUsage is stable across repeated calls and survives a recalc', async () => {
  // The external-memory report is computed as a delta against the last
  // reported figure, so a repeated call on an unchanged workbook must be
  // a no-op rather than double-counting, and the operations that trigger
  // their own report must not disturb the estimate either.
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  for (let row = 0; row < 500; row += 1) {
    wb.sheet(0).setFormula(row, 0, `=${row}+1`);
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
  wb.sheet(0).setNumber(0, 0, 5);
  wb.sheet(0).setFormula(0, 1, '=A1*2');
  const auto = wb.copySheet(0, { insertBefore: 1 });
  assert.equal(auto.index, 1);
  assert.equal(auto.name, 'Sheet1 (2)');
  assert.equal(wb.sheet(1), auto);
  const named = wb.copySheet(0, { insertBefore: 0, newName: 'Archive' });
  assert.equal(named.index, 0);
  assert.equal(named.name, 'Archive');
  assert.equal(wb.sheet(0), named);
  wb.recalc();
  assert.equal(wb.sheet(0).getTypedValue(0, 1).number, 10);
  assert.throws(
    () => wb.copySheet(9, { insertBefore: 0 }),
    (e) => isFormulonError(e, 10),
  );
  assert.throws(
    () => wb.copySheet(0, { insertBefore: 9 }),
    (e) => isFormulonError(e, 10),
  );
  assert.throws(
    () => wb.copySheet(0, { insertBefore: 0, newName: '' }),
    (e) => isFormulonError(e, 11),
  );
  assert.equal(wb.listSheets().length, 3);
  wb.dispose();
});
