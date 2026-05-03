// Copyright 2026 libraz. Licensed under the MIT License.
//
// Node-based smoke tests for the Formulon WASM bundle.
//
// Loads `build-wasm/formulon.js` (the Emscripten ES-module factory),
// exercises every embind export at least once, and exits 1 with a
// descriptive message on any failure.
//
// Run via `make test-wasm`, which checks for the build artefact and a
// Node binary first. The runner is intentionally framework-free: it
// uses only `node:assert/strict` so contributors can run it without an
// `npm install` step.

import assert from 'node:assert/strict';
import { fileURLToPath } from 'node:url';
import path from 'node:path';

// fm_value_kind_t mirror (see src/c_api/formulon_c.h).
const VAL = Object.freeze({
  BLANK: 0,
  NUMBER: 1,
  BOOL: 2,
  TEXT: 3,
  ERROR: 4,
  ARRAY: 5,
  REF: 6,
  LAMBDA: 7,
});

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const moduleUrl = path.resolve(__dirname, '..', '..', 'build-wasm', 'formulon.js');

const cases = [];
function test(name, fn) {
  cases.push({ name, fn });
}

let passed = 0;
let failed = 0;

async function run() {
  // Dynamic import: keeps the file syntactically valid even when the
  // wasm artifact is missing (the harness check is the Makefile's job).
  const factory = (await import(moduleUrl)).default;
  const Module = await factory();

  test('versionString is non-empty', () => {
    const v = Module.versionString();
    assert.equal(typeof v, 'string');
    assert.ok(v.length > 0, `expected non-empty version, got ${JSON.stringify(v)}`);
  });

  test('statusString covers known codes', () => {
    assert.equal(Module.statusString(0), 'kOk');
    // Unknown / out-of-band codes still return a non-null string.
    const unknown = Module.statusString(123456);
    assert.equal(typeof unknown, 'string');
    assert.ok(unknown.length > 0);
  });

  test('evalFormula(=SUM(1,2,3)) returns NUMBER 6', () => {
    const r = Module.evalFormula('=SUM(1,2,3)');
    assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
    assert.equal(r.value.kind, VAL.NUMBER);
    assert.equal(r.value.number, 6);
  });

  test('evalFormula(=IF(TRUE,1,2)) returns NUMBER 1', () => {
    const r = Module.evalFormula('=IF(TRUE,1,2)');
    assert.ok(r.status.ok);
    assert.equal(r.value.kind, VAL.NUMBER);
    assert.equal(r.value.number, 1);
  });

  test('evalFormula(=CONCAT) returns TEXT "hello world"', () => {
    const r = Module.evalFormula('=CONCAT("hello"," ","world")');
    assert.ok(r.status.ok);
    assert.equal(r.value.kind, VAL.TEXT);
    assert.equal(r.value.text, 'hello world');
  });

  test('Workbook construction + setNumber + recalc + getValue', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.isValid());
      assert.equal(wb.sheetCount(), 1);
      const nameRes = wb.sheetName(0);
      assert.ok(nameRes.status.ok);
      assert.equal(nameRes.value, 'Sheet1');

      assert.ok(wb.setNumber(0, 0, 0, 42).ok);
      assert.ok(wb.setFormula(0, 1, 0, '=A1*2').ok);
      assert.ok(wb.recalc().ok);

      const a1 = wb.getValue(0, 0, 0);
      assert.ok(a1.status.ok);
      assert.equal(a1.value.kind, VAL.NUMBER);
      assert.equal(a1.value.number, 42);

      const b1 = wb.getValue(0, 1, 0);
      assert.ok(b1.status.ok);
      assert.equal(b1.value.kind, VAL.NUMBER);
      assert.equal(b1.value.number, 84);

      // addSheet should grow sheetCount.
      assert.ok(wb.addSheet('Second').ok);
      assert.equal(wb.sheetCount(), 2);
    } finally {
      wb.delete();
    }
  });

  test('setText / setBool / setBlank round-trip', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setText(0, 0, 0, 'hello').ok);
      assert.ok(wb.setBool(0, 1, 0, true).ok);
      assert.ok(wb.setBlank(0, 2, 0).ok);
      assert.ok(wb.recalc().ok);

      const a1 = wb.getValue(0, 0, 0);
      assert.equal(a1.value.kind, VAL.TEXT);
      assert.equal(a1.value.text, 'hello');

      const a2 = wb.getValue(0, 1, 0);
      assert.equal(a2.value.kind, VAL.BOOL);
      assert.equal(a2.value.boolean, 1);

      const a3 = wb.getValue(0, 2, 0);
      assert.equal(a3.value.kind, VAL.BLANK);
    } finally {
      wb.delete();
    }
  });

  test('save() + Workbook.loadBytes() round-trip preserves a literal', () => {
    const wb = Module.Workbook.createDefault();
    let saved = null;
    try {
      assert.ok(wb.setNumber(0, 0, 0, 7).ok);
      assert.ok(wb.setFormula(0, 1, 0, '=A1+1').ok);
      assert.ok(wb.recalc().ok);

      const r = wb.save();
      assert.ok(r.status.ok, `save failed: ${JSON.stringify(r.status)}`);
      assert.ok(r.bytes instanceof Uint8Array);
      assert.ok(r.bytes.length > 0);
      saved = r.bytes;
    } finally {
      wb.delete();
    }

    const loaded = Module.Workbook.loadBytes(saved);
    try {
      assert.ok(loaded.isValid(), `load failed: ${Module.lastErrorMessage()}`);
      assert.ok(loaded.sheetCount() >= 1);
      assert.ok(loaded.recalc().ok);
      const a1 = loaded.getValue(0, 0, 0);
      assert.ok(a1.status.ok);
      assert.equal(a1.value.kind, VAL.NUMBER);
      assert.equal(a1.value.number, 7);
    } finally {
      loaded.delete();
    }
  });

  test('formula error surfaces as ERROR-kind value', () => {
    const r = Module.evalFormula('=1/0');
    assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
    assert.equal(r.value.kind, VAL.ERROR);
  });

  test('parser failure populates lastErrorMessage', () => {
    // An obviously unparseable formula. The C ABI either surfaces a
    // non-zero status (parser error) or returns ok with an Error
    // value — accept either, but the diagnostic should be readable
    // when the status is non-ok.
    const r = Module.evalFormula('=#GARBAGE');
    if (!r.status.ok) {
      assert.equal(typeof r.status.message, 'string');
      const msg = Module.lastErrorMessage();
      assert.equal(typeof msg, 'string');
    } else {
      // Acceptable alternative: the formula parses to something that
      // evaluates to an error value.
      assert.equal(r.value.kind, VAL.ERROR);
    }
  });

  test('setIterative accepts knobs without crashing', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setIterative(true, 50, 0.001).ok);
      // Set a fixed-point formula (sqrt(2) iteration). The engine may
      // surface NUMBER (converged) or ERROR depending on the recalc
      // path; we only assert that the call sequence doesn't crash and
      // that getValue returns a well-formed envelope.
      assert.ok(wb.setNumber(0, 0, 0, 1.0).ok);
      assert.ok(wb.setFormula(0, 0, 0, '=0.5*(A1+2/A1)').ok);
      assert.ok(wb.recalc().ok);
      const v = wb.getValue(0, 0, 0);
      assert.ok(v.status.ok);
      // Known well-formed kinds.
      assert.ok([VAL.NUMBER, VAL.ERROR, VAL.BLANK].includes(v.value.kind));
    } finally {
      wb.delete();
    }
  });

  test('renameSheet updates the sheet name', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.sheetCount(), 1);
      assert.ok(wb.renameSheet(0, 'Renamed').ok);
      const r = wb.sheetName(0);
      assert.ok(r.status.ok);
      assert.equal(r.value, 'Renamed');
    } finally {
      wb.delete();
    }
  });

  test('renameSheet rejects forbidden characters', () => {
    const wb = Module.Workbook.createDefault();
    try {
      const r = wb.renameSheet(0, 'Bad/Name');
      assert.equal(r.ok, false);
      assert.notEqual(r.status, 0);
    } finally {
      wb.delete();
    }
  });

  test('removeSheet drops a sheet but rejects the last one', () => {
    const wb = Module.Workbook.createEmpty();
    try {
      assert.ok(wb.addSheet('A').ok);
      assert.ok(wb.addSheet('B').ok);
      assert.ok(wb.addSheet('C').ok);
      assert.equal(wb.sheetCount(), 3);
      assert.ok(wb.removeSheet(1).ok);
      assert.equal(wb.sheetCount(), 2);
      assert.equal(wb.sheetName(0).value, 'A');
      assert.equal(wb.sheetName(1).value, 'C');
    } finally {
      wb.delete();
    }

    // Workbook with a single sheet rejects removeSheet(0).
    const lone = Module.Workbook.createDefault();
    try {
      const r = lone.removeSheet(0);
      assert.equal(r.ok, false);
    } finally {
      lone.delete();
    }
  });

  test('moveSheet rearranges sheets (Excel UI semantics)', () => {
    const wb = Module.Workbook.createEmpty();
    try {
      assert.ok(wb.addSheet('Alpha').ok);
      assert.ok(wb.addSheet('Beta').ok);
      assert.ok(wb.addSheet('Gamma').ok);
      // Move Alpha (0) to the end. Excel semantics: to=2 (post-removal).
      assert.ok(wb.moveSheet(0, 2).ok);
      assert.equal(wb.sheetName(0).value, 'Beta');
      assert.equal(wb.sheetName(1).value, 'Gamma');
      assert.equal(wb.sheetName(2).value, 'Alpha');
    } finally {
      wb.delete();
    }
  });

  test('setDefinedName adds, updates, and removes', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.equal(wb.definedNameCount(), 0);
      assert.ok(wb.setDefinedName('Pi', '=3.14').ok);
      assert.equal(wb.definedNameCount(), 1);
      const a = wb.definedNameAt(0);
      assert.ok(a.status.ok);
      assert.equal(a.name, 'Pi');
      assert.equal(a.formula, '=3.14');

      assert.ok(wb.setDefinedName('PI', '=3.14159').ok);
      assert.equal(wb.definedNameCount(), 1);
      const b = wb.definedNameAt(0);
      assert.equal(b.name, 'Pi');  // authored case preserved
      assert.equal(b.formula, '=3.14159');

      assert.ok(wb.setDefinedName('Pi', '').ok);
      assert.equal(wb.definedNameCount(), 0);
    } finally {
      wb.delete();
    }
  });

  test('cellCount + cellAt iterate populated cells in order', () => {
    const wb = Module.Workbook.createDefault();
    try {
      assert.ok(wb.setNumber(0, 0, 0, 1).ok);
      assert.ok(wb.setNumber(0, 0, 1, 2).ok);
      assert.ok(wb.setFormula(0, 1, 0, '=A1+B1').ok);
      assert.ok(wb.recalc().ok);

      const count = wb.cellCount(0);
      assert.ok(count >= 3, `expected >=3 cells, got ${count}`);

      // The very first iteration entry should be A1 (row=0, col=0).
      const c0 = wb.cellAt(0, 0);
      assert.ok(c0.status.ok);
      assert.equal(c0.row, 0);
      assert.equal(c0.col, 0);
      assert.equal(c0.formula, null);
      assert.equal(c0.value.kind, VAL.NUMBER);
      assert.equal(c0.value.number, 1);
    } finally {
      wb.delete();
    }
  });

  // ---- Run --------------------------------------------------------------
  for (const c of cases) {
    try {
      await c.fn();
      passed += 1;
      console.log(`ok   ${c.name}`);
    } catch (e) {
      failed += 1;
      console.error(`FAIL ${c.name}`);
      console.error(`  ${e && e.stack ? e.stack : e}`);
    }
  }

  console.log('');
  console.log(`Smoke summary: ${passed} passed, ${failed} failed (of ${cases.length})`);
  if (failed > 0) {
    process.exit(1);
  }
}

run().catch((e) => {
  console.error('Fatal harness error:', e);
  process.exit(1);
});
