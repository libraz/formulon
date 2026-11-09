// Argument and lifecycle contracts the native package shares: the same
// malformed call fails the same way on both packages and changes nothing.
// `packages/npm-native/test/smoke_argument_contract.mjs` holds the mirror.

import test from 'node:test';
import { assert, expectFailure, getModule, VAL } from './smoke_support.mjs';

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
    const refused = {
      name: 'FormulonError',
      code: 4005,
      message: /cannot dispose Workbook during an active method call/,
    };
    const spec = () => closingSpec(() => wb.dispose());
    assert.throws(() => wb.sheet(0).addProtectedRange(spec()), refused);
    assert.throws(() => wb.sheet(0).setAutoFilter(spec()), refused);
    assert.throws(() => wb.sheet(0).addShape(spec()), refused);
    assert.throws(() => wb.sheet(0).addMerge(spec()), refused);
    assert.equal(wb.disposed, false);
    assert.equal(wb.sheet(0).listProtectedRanges().length, 0);
    wb.sheet(0).setNumber(0, 0, 1);
  } finally {
    wb.dispose();
  }
});

test('delete() from the iterative progress callback is refused the same way', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    wb.updateIterative({ enabled: true, maxIterations: 10, maxChange: 0.001 });
    wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
    let caught;
    wb.setIterativeProgress(() => {
      try {
        wb.dispose();
      } catch (err) {
        caught = err;
      }
      return false;
    });
    wb.recalc();
    assert.equal(caught.name, 'FormulonError', String(caught));
    assert.equal(caught.code, 4005);
    assert.match(caught.message, /cannot dispose Workbook during an active method call/);
    assert.equal(wb.disposed, false);
  } finally {
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

test('a callback or sink that is not a function is rejected and keeps the installed one', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    wb.updateIterative({ enabled: true, maxIterations: 10, maxChange: 0.001 });
    wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
    let calls = 0;
    wb.setIterativeProgress(() => {
      calls += 1;
      return true;
    });
    assert.throws(() => wb.setIterativeProgress(42), TypeError);
    wb.recalc();
    assert.ok(calls > 0, 'the rejected install must leave the previous callback in place');

    assert.throws(() => mod.setLogSink(5), TypeError);
  } finally {
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

test('a missing required argument throws TypeError and commits nothing', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.throws(() => wb.sheet(0).setNumber(0, 0), TypeError);
    assert.throws(() => wb.sheet(0).setText(0, 0), TypeError);
    assert.throws(() => wb.sheet(0).setBool(0, 0), TypeError);
    assert.equal(wb.sheet(0).getTypedValue(0, 0).kind, VAL.BLANK);

    wb.setDefinedName({ name: 'X', formula: '=1' });
    const nameBefore = wb.listDefinedNames()[0];
    assert.throws(() => wb.setDefinedName({ name: 'X' }), TypeError);
    assert.throws(() => wb.setDefinedName(), TypeError);
    assert.equal(wb.listDefinedNames().length, 1);
    assert.deepEqual(wb.listDefinedNames()[0], nameBefore);

    assert.throws(() => wb.updateIterative(), TypeError);
    assert.equal(wb.getIterative().enabled, false);
  } finally {
    wb.dispose();
  }
});

test('a primitive or null where an object is declared throws TypeError and changes nothing', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    wb.sheet(0).setSheetProtection({});
    const protectionBefore = wb.sheet(0).getSheetProtection();
    assert.throws(() => wb.sheet(0).setSheetProtection('x'), TypeError);
    assert.throws(() => wb.sheet(0).setSheetProtection(5), TypeError);
    assert.deepEqual(wb.sheet(0).getSheetProtection(), protectionBefore);

    const fonts = wb.ooxml.listFonts().length;
    assert.throws(() => wb.ooxml.addFont(null), TypeError);
    assert.throws(() => wb.ooxml.addFont(1), TypeError);
    assert.throws(() => wb.setDefaultFont(null), TypeError);
    assert.equal(wb.ooxml.listFonts().length, fonts);

    assert.throws(() => wb.sheet(0).recalcViewport(5), TypeError);
    assert.throws(() => wb.sheet(0).getCellRectPt(null, 0), TypeError);
    assert.throws(() => wb.sheet(0).setCellPhoneticProperties(0, 0, 'x'), TypeError);
    assert.throws(() => wb.sheet(0).ooxml.restoreRangeXfs(null), TypeError);
  } finally {
    wb.dispose();
  }
});

test('an argument the binding cannot represent is a TypeError', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    for (const [label, call] of [
      ['setThemeColors([])', () => wb.setThemeColors([])],
      ['restoreRangeXfs(0, {})', () => wb.sheet(0).ooxml.restoreRangeXfs({})],
      [
        'remapRangeXfs mismatched lengths',
        () => wb.sheet(0).ooxml.remapRangeXfs({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, [1], []),
      ],
      ['getImageInfo({})', () => mod.getImageInfo({})],
    ]) {
      assert.throws(call, TypeError, label);
    }
  } finally {
    wb.dispose();
  }
});

test('a failed read throws instead of returning a default payload', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    expectFailure(() => wb.ooxml.getCellXf(99999), mod.StatusCode.IndexOutOfRange);
    expectFailure(() => wb.ooxml.getCellStyleXf(99999), mod.StatusCode.IndexOutOfRange);
    expectFailure(
      () => wb.sheet(99).validateValue(0, 0, { kind: VAL.NUMBER, number: 1 }),
      mod.StatusCode.SheetIndexOutOfRange,
    );
  } finally {
    wb.dispose();
  }
});

test('a lone surrogate in a spec string is stored as U+FFFD', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    wb.addPerson({ id: PERSON_ID, displayName: 'a\uD800b' });
    const person = wb.listPeople().find((p) => p.id === PERSON_ID);
    assert.equal(person.displayName, 'a\uFFFDb');
    wb.sheet(0).setText(0, 0, 'x\uDC00');
    assert.equal(wb.sheet(0).getTypedValue(0, 0).text, 'x\uFFFD');
  } finally {
    wb.dispose();
  }
});

test('every method called from the iterative progress callback throws and changes nothing', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    wb.addSheet('Second');
    wb.sheet(1).setNumber(5, 5, 42);
    wb.updateIterative({ enabled: true, maxIterations: 10, maxChange: 0.001 });
    wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
    const attempts = {
      setFormula: () => wb.sheet(1).setFormula(0, 0, '=1'),
      setNumber: () => wb.sheet(1).setNumber(5, 5, 7),
      deleteRows: () => wb.sheet(1).deleteRows(0, 1),
      removeSheet: () => wb.removeSheet(1),
      getValue: () => wb.sheet(1).getTypedValue(5, 5),
    };
    const thrown = {};
    let calls = 0;
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
    });
    wb.recalc();
    assert.equal(calls, 1);
    for (const name of Object.keys(attempts)) {
      assert.equal(thrown[name]?.name, 'FormulonError', `${name} must throw FormulonError`);
      assert.equal(thrown[name].code, 4005, name);
      assert.match(thrown[name].message, /kGraphRecalcReentrant, 4005/, name);
    }
    assert.equal(wb.listSheets().length, 2);
    assert.equal(wb.sheet(1).getTypedValue(5, 5).number, 42);
    assert.equal(wb.sheet(1).getTypedValue(0, 0).kind, 0, 'the refused setFormula wrote nothing');
  } finally {
    wb.setIterativeProgress(null);
    wb.dispose();
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
      { ...rect, lastRow: null },
    ];
    for (const bad of malformed) {
      assert.throws(() => wb.sheet(0).addMerge(bad), TypeError, JSON.stringify(bad));
      assert.throws(() => wb.sheet(0).ooxml.setRangeXfIndex(bad, 0), TypeError, JSON.stringify(bad));
      assert.throws(() => wb.sheet(0).getCellRectPt(bad, 0), TypeError, JSON.stringify(bad));
    }
    // A non-integer field is a range error; a negative one is a refused coordinate.
    for (const bad of [
      { ...rect, lastCol: 1.5 },
      { ...rect, firstCol: Number.NaN },
    ]) {
      assert.throws(() => wb.sheet(0).addMerge(bad), RangeError, JSON.stringify(bad));
    }
    expectFailure(() => wb.sheet(0).addMerge({ ...rect, firstRow: -1 }));
    assert.equal(wb.sheet(0).listMerges().length, 0);
    // A well-typed rectangle the engine refuses is a status, not a throw.
    expectFailure(() => wb.sheet(0).addMerge({ firstRow: 1, firstCol: 0, lastRow: 0, lastCol: 0 }));
    expectFailure(() =>
      wb.sheet(0).ooxml.setRangeXfIndex({ firstRow: 0, firstCol: 0, lastRow: 1048576, lastCol: 0 }, 0),
    );
    wb.sheet(0).addMerge(rect);
  } finally {
    wb.dispose();
  }
});
