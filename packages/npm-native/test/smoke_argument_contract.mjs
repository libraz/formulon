// Argument and lifecycle contracts the WASM package shares: the same
// malformed call fails the same way on both packages and changes nothing.
// `packages/npm/test/smoke_argument_contract.mjs` holds the mirror.

import test from 'node:test';
import v8 from 'node:v8';
import vm from 'node:vm';
import { assert, getModule, isFormulonError } from './smoke_support.mjs';

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
  assert.equal(wb.disposed, true);
});

test('dispose() from the iterative progress callback is refused the same way', async () => {
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
      wb.setIterativeProgress(() => {
        void wb;
        return true;
      });
      registry.register(wb, i);
    }
  })();
  for (let round = 0; round < 100 && collected < count; round += 1) {
    gc();
    await new Promise((resolve) => setImmediate(resolve));
  }
  assert.equal(collected, count, 'every unreachable workbook must be finalized');
});

test('a callback or sink that is not a function throws TypeError and keeps the installed one', async () => {
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
    assert.equal(wb.sheet(0).getTypedValue(0, 0).kind, mod.ValueKind.Blank);

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

test('an argument the binding cannot represent throws TypeError and changes nothing', async () => {
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
    for (const read of [() => wb.ooxml.getCellXf(99999), () => wb.ooxml.getCellStyleXf(99999)]) {
      assert.throws(read, (e) => isFormulonError(e, mod.StatusCode.IndexOutOfRange));
    }
    assert.throws(
      () => wb.sheet(99).validateValue(0, 0, { kind: mod.ValueKind.Number, number: 1 }),
      (e) => isFormulonError(e, mod.StatusCode.SheetIndexOutOfRange),
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
      { ...rect, lastCol: 1.5 },
      { ...rect, firstCol: Number.NaN },
      { ...rect, lastRow: null },
    ];
    for (const bad of malformed) {
      assert.throws(() => wb.sheet(0).addMerge(bad), TypeError, JSON.stringify(bad));
      assert.throws(() => wb.sheet(0).ooxml.setRangeXfIndex(bad, 0), TypeError, JSON.stringify(bad));
      assert.throws(() => wb.sheet(0).getCellRectPt(bad, 0), TypeError, JSON.stringify(bad));
    }
    assert.throws(() => wb.sheet(0).addMerge({ ...rect, firstRow: -1 }), RangeError);
    assert.equal(wb.sheet(0).listMerges().length, 0);
    // A well-typed rectangle the engine refuses is a FormulonError, not a TypeError.
    assert.throws(() => wb.sheet(0).addMerge({ firstRow: 1, firstCol: 0, lastRow: 0, lastCol: 0 }), isFormulonError);
    assert.throws(
      () => wb.sheet(0).ooxml.setRangeXfIndex({ firstRow: 0, firstCol: 0, lastRow: 1048576, lastCol: 0 }, 0),
      isFormulonError,
    );
    wb.sheet(0).addMerge(rect);
  } finally {
    wb.dispose();
  }
});
