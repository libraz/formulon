// Argument and lifecycle contracts the native package shares: the same
// malformed call fails the same way on both packages and changes nothing.
// `packages/npm-native/test/smoke_argument_contract.mjs` holds the mirror.

import test from 'node:test';
import { assert, getModule, readAll, VAL } from './smoke_support.mjs';

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

test('delete() from a spec getter during a call is refused and the call does not crash', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const refused = { name: 'TypeError', message: /cannot dispose Workbook during an active method call/ };
    const spec = () => closingSpec(() => wb.delete());
    assert.throws(() => wb.addProtectedRange(0, spec()), refused);
    assert.throws(() => wb.updateProtectedRange(0, 0, spec()), refused);
    assert.throws(() => wb.setAutoFilter(0, spec()), refused);
    assert.throws(() => wb.addShape(0, spec()), refused);
    assert.throws(() => wb.addMerge(0, spec()), refused);
    assert.equal(wb.isValid(), true);
    assert.equal(
      readAll(
        () => wb.getProtectedRangeCount(0),
        (i) => wb.getProtectedRange(0, i),
      ).length,
      0,
    );
    assert.ok(wb.setNumber(0, 0, 0, 1).ok);
  } finally {
    wb.delete();
  }
});

test('delete() from the iterative progress callback is refused the same way', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(wb.setIterative(true, 10, 0.001).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
    let caught;
    assert.ok(
      wb.setIterativeProgress(() => {
        try {
          wb.delete();
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
    wb.delete();
  }
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
    wb.delete();
  }
});

test('a missing required argument throws TypeError and commits nothing', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.throws(() => wb.setNumber(0, 0, 0), TypeError);
    assert.throws(() => wb.setText(0, 0, 0), TypeError);
    assert.throws(() => wb.setBool(0, 0, 0), TypeError);
    assert.equal(wb.getValue(0, 0, 0).value.kind, VAL.BLANK);

    assert.ok(wb.setDefinedName('X', '=1').ok);
    const nameBefore = wb.listDefinedNames()[0];
    assert.throws(() => wb.setDefinedName('X'), TypeError);
    assert.equal(wb.listDefinedNames().length, 1);
    assert.deepEqual(wb.listDefinedNames()[0], nameBefore);

    assert.throws(() => wb.setIterative(true), TypeError);
    assert.equal(wb.getIterative().enabled, false);
  } finally {
    wb.delete();
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

    const fonts = wb.getFontCount().value;
    assert.throws(() => wb.addFont(null), TypeError);
    assert.throws(() => wb.addFont(1), TypeError);
    assert.throws(() => wb.setDefaultFont(null), TypeError);
    assert.equal(wb.getFontCount().value, fonts);

    assert.throws(() => wb.recalcViewport(0, 'x'), TypeError);
    assert.throws(() => wb.getCellRectPt(0, null, 0), TypeError);
    assert.throws(() => wb.setCellPhoneticProperties(0, 0, 0, 'x'), TypeError);
    assert.throws(() => wb.restoreRangeStyle(0, null), TypeError);
  } finally {
    wb.delete();
  }
});

test('an argument a method rejects with a status reports kInvalidArgument', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    for (const [label, result] of [
      ['setThemeColors([])', wb.setThemeColors([])],
      ['restoreRangeStyle(0, {})', wb.restoreRangeStyle(0, {})],
      [
        'remapRangeStyle mismatched lengths',
        wb.remapRangeStyle(0, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, [1], []),
      ],
      ['probeImage({})', wb.getImageInfo({})],
    ]) {
      const status = typeof result.ok === 'boolean' ? result : result.status;
      assert.equal(status.ok, false, label);
      assert.equal(status.status, 2, `${label}: ${JSON.stringify(status)}`);
      assert.equal(status.context, '', label);
    }
  } finally {
    wb.delete();
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
    const outcome = wb.validateValue(99, 0, 0, { kind: VAL.NUMBER, number: 1 });
    assert.equal(outcome.status.ok, false);
    assert.equal(outcome.hasRule, false);
    assert.equal(outcome.valid, false);
  } finally {
    wb.delete();
  }
});

test('a lone surrogate in a spec string is stored as U+FFFD', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(wb.addPerson({ id: PERSON_ID, displayName: 'a\uD800b' }).ok);
    const person = wb.listPeople().persons.find((p) => p.id === PERSON_ID);
    assert.equal(person.displayName, 'a\uFFFDb');
    assert.ok(wb.setText(0, 0, 0, 'x\uDC00').ok);
    assert.equal(wb.getValue(0, 0, 0).value.text, 'x\uFFFD');
  } finally {
    wb.delete();
  }
});

test('every method called from the iterative progress callback throws and changes nothing', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(wb.addSheet('Second').status.ok);
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
    assert.equal(wb.getSheetCount().value, 2);
    assert.equal(wb.getValue(1, 5, 5).value.number, 42);
    assert.equal(wb.getValue(1, 0, 0).value.kind, 0, 'the refused setFormula wrote nothing');
  } finally {
    wb.setIterativeProgress(null);
    wb.delete();
  }
});

test('a CellRect argument requires four integer fields', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const rect = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };
    const malformed = [
      { firstRow: 0, firstCol: 0, lastRow: 1 },
      { ...rect, firstRow: '0' },
      { ...rect, lastCol: 1.5 },
      { ...rect, firstCol: Number.NaN },
      { ...rect, lastRow: null },
    ];
    for (const bad of malformed) {
      assert.throws(() => wb.addMerge(0, bad), TypeError, JSON.stringify(bad));
      assert.throws(() => wb.setRangeXfIndex(0, bad, 0), TypeError, JSON.stringify(bad));
      assert.throws(() => wb.getCellRectPt(0, bad, 0), TypeError, JSON.stringify(bad));
    }
    assert.throws(() => wb.addMerge(0, { ...rect, firstRow: -1 }), RangeError);
    assert.equal(wb.listMerges(0).length, 0);
    // A well-typed rectangle the engine refuses is a status, not a throw.
    assert.equal(wb.addMerge(0, { firstRow: 1, firstCol: 0, lastRow: 0, lastCol: 0 }).ok, false);
    assert.equal(wb.setRangeXfIndex(0, { firstRow: 0, firstCol: 0, lastRow: 1048576, lastCol: 0 }, 0).ok, false);
    assert.ok(wb.addMerge(0, rect).ok);
  } finally {
    wb.delete();
  }
});
