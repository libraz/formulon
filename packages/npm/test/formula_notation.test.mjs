// `localizeFormula` / `canonicalizeFormula` with `{ r1c1, baseRow, baseCol }`: the R1C1 spelling of each locale,
// and the options read and refused as `parseReference` reads and refuses them.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

const kInvalidArgument = 2;
const kInvalidReference = 1004;

// Mac Excel's FormulaR1C1Local of formulas entered in E5.
const AT_E5 = Object.freeze({ r1c1: true, baseRow: 4, baseCol: 4 });
const CASES = [
  ['mac-365-en_US', '=SUM(C3,$B$2)', '=SUM(R[-2]C[-2],R2C2)'],
  ['mac-365-de_DE', '=SUM(C3,$B$2)', '=SUMME(Z(-2)S(-2);Z2S2)'],
  ['mac-365-es_ES', '=SUM(C3,$B$2)', '=SUMA(F[-2]C[-2];F2C2)'],
  ['mac-365-de_DE', '=IF($A$1>0,$B$1,"x;y")', '=WENN(Z1S1>0;Z1S2;"x;y")'],
];

test('localizeFormula and canonicalizeFormula round-trip R1C1 with a base', async () => {
  const f = await getModule();
  for (const [profile, a1, local] of CASES) {
    assert.equal(f.localizeFormula(a1, profile, AT_E5), local, profile);
    assert.equal(f.canonicalizeFormula(local, profile, AT_E5), a1, profile);
  }
  // Absolute references need no base.
  assert.equal(f.localizeFormula('=$A$1+SUM($1:$2)', 'mac-365-de_DE', { r1c1: true }), '=Z1S1+SUMME(Z1:Z2)');
  assert.equal(f.canonicalizeFormula('=Z1S1+SUMME(Z1:Z2)', 'mac-365-de_DE', { r1c1: true }), '=$A$1+SUM($1:$2)');
});

test('absent, null, empty and false options keep A1', async () => {
  const f = await getModule();
  const a1 = '=SUM(C3,$B$2)';
  const shown = f.localizeFormula(a1, 'mac-365-de_DE');
  assert.equal(shown, '=SUMME(C3;$B$2)');
  for (const options of [undefined, null, {}, { r1c1: false }, { r1c1: null }, { baseRow: 4, baseCol: 4 }]) {
    assert.equal(f.localizeFormula(a1, 'mac-365-de_DE', options), shown, JSON.stringify(options));
    assert.equal(f.canonicalizeFormula(shown, 'mac-365-de_DE', options), a1, JSON.stringify(options));
  }
});

test('an R1C1 relative reference without a base fails over that reference', async () => {
  const f = await getModule();
  const failure = (fn) => {
    try {
      fn();
    } catch (e) {
      assert.equal(e.name, 'FormulonError');
      assert.equal(e.code, kInvalidReference);
      assert.ok(e.parseError, 'parseError is attached');
      return e.parseError;
    }
    assert.fail('ParserInvalidReference is expected');
  };
  const localized = failure(() => f.localizeFormula('=SUM(A1,$B$2)', 'mac-365-en_US', { r1c1: true }));
  assert.equal(localized.start, 5);
  assert.equal(localized.end, 7);
  const canonical = failure(() => f.canonicalizeFormula('=1+Z(-1)S', 'mac-365-de_DE', { r1c1: true }));
  assert.equal(canonical.start, 3);
  assert.equal(canonical.end, 9);
  const malformed = failure(() => f.canonicalizeFormula('=SUMME(Z(x)S;1)', 'mac-365-de_DE', AT_E5));
  assert.equal(malformed.start, 7);
  assert.equal(malformed.end, 11);
});

test('the options are refused as parseReference refuses them', async () => {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    for (const name of ['localizeFormula', 'canonicalizeFormula']) {
      const call = (options) => f[name]('=1', 'mac-365-de_DE', options);
      const parse = (options) => wb.parseReference('A1', options);
      for (const run of [call, parse]) {
        assert.throws(() => run('r1c1'), TypeError, name);
        assert.throws(() => run({ r1c1: true, base: 1 }), /options has no field base$/, name);
        assert.throws(() => run({ r1c1: 1 }), /options\.r1c1 must be a boolean/, name);
        assert.throws(() => run({ r1c1: true, baseRow: 1 }), /must be given together/, name);
        assert.throws(() => run({ r1c1: true, baseCol: 1 }), /must be given together/, name);
        assert.throws(() => run({ baseRow: 1.5, baseCol: 0 }), RangeError, name);
        assert.throws(
          () => run({ baseRow: 1048576, baseCol: 0 }),
          (e) => e.name === 'FormulonError' && e.code === kInvalidArgument,
          name,
        );
      }
      assert.throws(
        () => call({ R1C1: true }),
        { name: 'TypeError', message: `Formulon.${name} options has no field R1C1` },
        name,
      );
      assert.throws(
        () => f[name]('=1', 'nope', { r1c1: true }),
        (e) => e.name === 'FormulonError' && e.code === kInvalidArgument,
        name,
      );
    }
  } finally {
    wb.dispose();
  }
});
