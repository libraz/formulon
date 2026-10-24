// Argument and lifecycle contracts the WASM package shares: the same
// malformed call fails the same way on both packages and changes nothing.
// `packages/npm/test/smoke_argument_contract.mjs` holds the mirror.

import test from 'node:test';
import v8 from 'node:v8';
import vm from 'node:vm';
import { assert, getModule } from './smoke_support.mjs';

const PERSON_ID = '{11111111-1111-4111-8111-111111111111}';

// A spec object whose every property probe runs `close` first.
function closingSpec(close) {
  const trap = () => {
    close();
    return undefined;
  };
  return new Proxy(
    {},
    {
      has: () => {
        close();
        return false;
      },
      get: trap,
      getOwnPropertyDescriptor: trap,
    },
  );
}

test('dispose() from a spec getter during a call is refused and the call does not crash', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const refused = { name: 'TypeError', message: /cannot dispose Workbook during an active method call/ };
    const spec = () => closingSpec(() => wb.dispose());
    assert.throws(() => wb.addProtectedRange(0, spec()), refused);
    assert.throws(() => wb.updateProtectedRange(0, 0, spec()), refused);
    assert.throws(() => wb.setAutoFilter(0, spec()), refused);
    assert.throws(() => wb.insertShape(0, spec()), refused);
    assert.throws(() => wb.addMerge(0, spec()), refused);
    assert.equal(wb.isValid(), true);
    assert.equal(wb.getProtectedRanges(0).length, 0);
    assert.ok(wb.setNumber(0, 0, 0, 1).ok);
  } finally {
    wb.dispose();
  }
  assert.equal(wb.isValid(), false);
});

test('dispose() from the iterative progress callback is refused the same way', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(wb.setIterative(true, 10, 0.001).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
    let caught;
    assert.ok(
      wb.setIterativeProgress(() => {
        try {
          wb.dispose();
        } catch (err) {
          caught = err;
        }
        return false;
      }).ok,
    );
    assert.ok(wb.recalc().ok);
    assert.ok(caught instanceof TypeError, String(caught));
    assert.match(caught.message, /cannot dispose Workbook during an active method call/);
    assert.equal(wb.isValid(), true);
  } finally {
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

test('a workbook whose progress callback refers to it is collected once unreachable', async () => {
  const mod = await getModule();
  v8.setFlagsFromString('--expose-gc');
  const gc = vm.runInNewContext('gc');
  const count = 4;
  let collected = 0;
  const registry = new FinalizationRegistry(() => {
    collected += 1;
  });
  (() => {
    for (let i = 0; i < count; i += 1) {
      const wb = mod.Workbook.createDefault();
      assert.ok(
        wb.setIterativeProgress(() => {
          void wb;
          return true;
        }).ok,
      );
      registry.register(wb, i);
    }
  })();
  for (let round = 0; round < 100 && collected < count; round += 1) {
    gc();
    await new Promise((resolve) => setImmediate(resolve));
  }
  assert.equal(collected, count, 'every unreachable workbook must be finalized');
});

test('a callback or sink that is not a function is rejected with status 2 and keeps the installed one', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(wb.setIterative(true, 10, 0.001).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
    let calls = 0;
    assert.ok(
      wb.setIterativeProgress(() => {
        calls += 1;
        return true;
      }).ok,
    );
    const bad = wb.setIterativeProgress(42);
    assert.equal(bad.ok, false);
    assert.equal(bad.status, 2);
    assert.equal(bad.context, '');
    assert.ok(wb.recalc().ok);
    assert.ok(calls > 0, 'the rejected install must leave the previous callback in place');

    const badSink = mod.setLogSink(5);
    assert.equal(badSink.ok, false);
    assert.equal(badSink.status, 2);
  } finally {
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

test('a missing required argument throws TypeError and commits nothing', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.throws(() => wb.setNumber(0, 0, 0), TypeError);
    assert.throws(() => wb.setText(0, 0, 0), TypeError);
    assert.throws(() => wb.setBool(0, 0, 0), TypeError);
    assert.equal(wb.getValue(0, 0, 0).value.kind, mod.ValueKind.Blank);

    assert.ok(wb.setDefinedName('X', '=1').ok);
    const nameBefore = wb.definedNameAt(0);
    assert.throws(() => wb.setDefinedName('X'), TypeError);
    assert.equal(wb.definedNameCount().value, 1);
    assert.deepEqual(wb.definedNameAt(0), nameBefore);

    assert.throws(() => wb.setIterative(true), TypeError);
    assert.equal(wb.getIterative().enabled, false);
  } finally {
    wb.dispose();
  }
});

test('a primitive or null where an object is declared throws TypeError and changes nothing', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(wb.setSheetProtection(0, { enabled: 1, sheet: 1 }).ok);
    const protectionBefore = wb.getSheetProtection(0).protection;
    assert.throws(() => wb.setSheetProtection(0, 'x'), TypeError);
    assert.throws(() => wb.setSheetProtection(0, null), TypeError);
    assert.deepEqual(wb.getSheetProtection(0).protection, protectionBefore);

    const fonts = wb.fontCount().value;
    assert.throws(() => wb.addFont(null), TypeError);
    assert.throws(() => wb.addFont(1), TypeError);
    assert.throws(() => wb.setDefaultFont(null), TypeError);
    assert.equal(wb.fontCount().value, fonts);

    assert.throws(() => wb.partialRecalc('x'), TypeError);
    assert.throws(() => wb.getCellRectPt(0, null, 0), TypeError);
    assert.throws(() => wb.setCellPhoneticProperties(0, 0, 0, 'x'), TypeError);
    assert.throws(() => wb.restoreRangeStyle(0, null), TypeError);
  } finally {
    wb.dispose();
  }
});

test('an argument a method rejects with a status reports kInvalidArgument', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    for (const [label, result] of [
      ['setThemeColors([])', wb.setThemeColors([])],
      ['restoreRangeStyle(0, {})', wb.restoreRangeStyle(0, {})],
      ['remapRangeStyle mismatched lengths', wb.remapRangeStyle(0, 0, 0, 0, 0, [1], [])],
      ['probeImage({})', wb.probeImage({})],
    ]) {
      const status = typeof result.ok === 'boolean' ? result : result.status;
      assert.equal(status.ok, false, label);
      assert.equal(status.status, 2, `${label}: ${JSON.stringify(status)}`);
      assert.equal(status.context, '', label);
    }
  } finally {
    wb.dispose();
  }
});

test('a failed read reports the documented default payload', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    for (const xf of [wb.getCellXf(99999), wb.getCellStyleXf(99999)]) {
      assert.equal(xf.status.ok, false);
      assert.equal(xf.locked, true);
      assert.equal(xf.hidden, false);
    }
    const outcome = wb.validateValue(99, 0, 0, { kind: mod.ValueKind.Number, number: 1 });
    assert.equal(outcome.status.ok, false);
    assert.equal(outcome.hasRule, false);
    assert.equal(outcome.valid, false);
  } finally {
    wb.dispose();
  }
});

test('a lone surrogate in a spec string is stored as U+FFFD', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(wb.addPerson({ id: PERSON_ID, displayName: 'a\uD800b' }).ok);
    const person = wb.getPersons().persons.find((p) => p.id === PERSON_ID);
    assert.equal(person.displayName, 'a\uFFFDb');
    assert.ok(wb.setText(0, 0, 0, 'x\uDC00').ok);
    assert.equal(wb.getValue(0, 0, 0).value.text, 'x\uFFFD');
  } finally {
    wb.dispose();
  }
});

test('every method called from the iterative progress callback throws and changes nothing', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(wb.addSheet('Second').ok);
    assert.ok(wb.setNumber(1, 5, 5, 42).ok);
    assert.ok(wb.setIterative(true, 10, 0.001).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
    const attempts = {
      setFormula: () => wb.setFormula(1, 0, 0, '=1'),
      setNumber: () => wb.setNumber(1, 5, 5, 7),
      deleteRows: () => wb.deleteRows(1, 0, 1),
      removeSheet: () => wb.removeSheet(1),
      getValue: () => wb.getValue(1, 5, 5),
    };
    const thrown = {};
    let calls = 0;
    assert.ok(
      wb.setIterativeProgress(() => {
        calls += 1;
        for (const [name, attempt] of Object.entries(attempts)) {
          try {
            attempt();
          } catch (err) {
            thrown[name] = err;
          }
        }
        return false;
      }).ok,
    );
    assert.ok(wb.recalc().ok);
    assert.equal(calls, 1);
    for (const name of Object.keys(attempts)) {
      assert.ok(thrown[name] instanceof TypeError, `${name} must throw`);
      assert.match(thrown[name].message, /kGraphRecalcReentrant, 4005/, name);
    }
    assert.equal(wb.sheetCount().value, 2);
    assert.equal(wb.getValue(1, 5, 5).value.number, 42);
    assert.equal(wb.getValue(1, 0, 0).value.kind, 0, 'the refused setFormula wrote nothing');
  } finally {
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});
