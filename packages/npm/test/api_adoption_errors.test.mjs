// Errors and names of the facade: FormulonError with parseError and its string form, formula syntax errors
// at every formula entry, formula text through setValue / setValues, ExcelError, enumName, the ordinal tables
// and the members of the Formulon object. FORMULON_PACKAGE=native runs it against the staged
// @libraz/formulon-native; otherwise FORMULON_NPM_ENTRY selects the staged @libraz/formulon subpath.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(testDir, '..', '..', '..');

async function loadPackage() {
  if (process.env.FORMULON_PACKAGE === 'native') {
    return import(pathToFileURL(path.join(repoRoot, 'packages', 'npm-native', 'dist', 'index.mjs')).href);
  }
  const pkgRoot = path.resolve(testDir, '..');
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[process.env.FORMULON_NPM_ENTRY ?? '.']?.import;
  if (!target) throw new Error('package.json "exports" has no import target for the selected entry');
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

const mod = await loadPackage();
const formulon = await mod.createFormulon();
const { ErrorCode, ExcelError, FormulonError, ParseErrorCode, StatusCode, ValueKind, Workbook } = formulon;

const ORDINAL_TABLES = [
  'ParseErrorCode',
  'CfRuleType',
  'ComparisonOperator',
  'DatePeriod',
  'CfValueObjectType',
  'IconSetName',
  'ValidationType',
];

function withSheet(body) {
  const wb = Workbook.createDefault();
  try {
    return body(wb.sheet(0), wb);
  } finally {
    wb.dispose();
  }
}

function thrown(call) {
  try {
    call();
  } catch (error) {
    return error;
  }
  assert.fail('expected the call to throw');
}

/** The table's cases of `entries` whose formula is given inline. */
async function parseCases(...entries) {
  const table = JSON.parse(await readFile(path.join(repoRoot, 'tools', 'api', 'parse_error_cases.json'), 'utf8'));
  return table.cases.filter((c) => entries.includes(c.entry) && c.formula !== undefined);
}

function assertSyntaxError(error, c) {
  assert.equal(error.name, 'FormulonError', `${c.id}: ${error}`);
  assert.ok(error instanceof FormulonError, c.id);
  assert.equal(error.code, c.status, c.id);
  assert.ok(error.parseError, c.id);
  assert.ok(
    c.codes.map((name) => ParseErrorCode[name]).includes(error.parseError.code),
    `${c.id}: code ${error.parseError.code}`,
  );
  if (c.start !== null) {
    const { start, end, line, column } = error.parseError;
    assert.deepEqual(
      { start, end, line, column },
      { start: c.start, end: c.end, line: c.line, column: c.column },
      c.id,
    );
  }
}

test('FormulonError carries its parse error and spells its status in toString', () => {
  const plain = new FormulonError(StatusCode.NotFound, 'no such sheet');
  assert.equal(plain.name, 'FormulonError');
  assert.equal(plain.context, '');
  assert.equal(plain.parseError, undefined);
  assert.equal(String(plain), 'FormulonError [NotFound]: no such sheet');
  const where = { code: ParseErrorCode.UnexpectedEof, start: 2, end: 2, line: 1, column: 3 };
  const parsed = new FormulonError(StatusCode.ParserUnexpectedToken, 'unexpected end of input', 'row=0 col=0', where);
  assert.deepEqual(parsed.parseError, where);
  assert.equal(String(parsed), 'FormulonError [ParserUnexpectedToken]: unexpected end of input');
  assert.equal(String(new FormulonError(123456, 'x')), 'FormulonError [123456]: x');
});

test('a formula syntax error fails with its status and position at every cell and evaluation entry', async () => {
  const cases = await parseCases('set_formula', 'evaluate_formula', 'evaluate_formula_array');
  assert.ok(cases.length > 0);
  withSheet((ws) => {
    for (const c of cases) {
      if (c.entry === 'set_formula') {
        assertSyntaxError(
          thrown(() => ws.setFormula('A1', c.formula)),
          c,
        );
        if (c.formula.startsWith('='))
          assertSyntaxError(
            thrown(() => ws.setValue('A1', c.formula)),
            c,
          );
      } else if (c.entry === 'evaluate_formula') {
        assertSyntaxError(
          thrown(() => ws.evaluateFormula('A1', c.formula)),
          c,
        );
        assertSyntaxError(
          thrown(() => formulon.evaluateFormula(c.formula)),
          c,
        );
        assertSyntaxError(
          thrown(() => formulon.evaluateTypedFormula(c.formula)),
          c,
        );
      } else {
        assertSyntaxError(
          thrown(() => ws.evaluateFormulaArray('A1', c.formula)),
          c,
        );
        assertSyntaxError(
          thrown(() => ws.evaluateTypedFormulaArray('A1', c.formula)),
          c,
        );
      }
    }
    assert.equal(ws.getFormula('A1'), null);
  });
});

test('a reference that does not parse fails with 1004 over the whole text', async () => {
  const cases = await parseCases('parse_reference');
  assert.ok(cases.length > 0);
  withSheet((ws, wb) => {
    for (const c of cases)
      assertSyntaxError(
        thrown(() => wb.parseReference(c.formula)),
        c,
      );
    const error = thrown(() => ws.getValue('XFE1'));
    assert.equal(error.code, StatusCode.ParserInvalidReference);
    assert.equal(error.parseError.end, 4);
  });
});

test('a record formula that does not parse names its field', () => {
  withSheet((ws) => {
    const error = thrown(() =>
      ws.addCfRule({
        type: formulon.CfRuleType.Expression,
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }],
        formula1: '=1+　2',
      }),
    );
    assert.ok(error.parseError);
    assert.match(error.context, /field=formula1/);
  });
});

test('a failure that is not a syntax error carries no parseError', () => {
  withSheet((ws, wb) => {
    const error = thrown(() => wb.sheet(9));
    assert.equal(error.code, StatusCode.SheetIndexOutOfRange);
    assert.equal(error.parseError, undefined);
    assert.equal(ws.getValue('A1'), null);
  });
});

test('a sheet-qualified address is refused with the call that reads it', () => {
  withSheet((ws) => {
    const error = thrown(() => ws.getValue('Sheet1!A1'));
    assert.equal(error.code, StatusCode.InvalidArgument);
    assert.match(error.message, /wb\.sheet\("Sheet1"\)\.getValue\("A1"\)/);
  });
});

test('setValue reads a string starting with = as a formula and every other string as text', () => {
  withSheet((ws) => {
    ws.setValue('A1', 2);
    ws.setValue('A2', '=A1*3');
    assert.equal(ws.getFormula('A2'), '=A1*3');
    assert.equal(ws.getValue('A2'), 6);
    for (const text of ['+1', '-1', '@A1', '＝1', ' =1', '']) {
      ws.setValue('B1', text);
      assert.equal(ws.getFormula('B1'), null, JSON.stringify(text));
      assert.equal(ws.getValue('B1'), text === '' ? '' : text, JSON.stringify(text));
    }
    const lone = thrown(() => ws.setValue('B2', '='));
    assert.ok(lone.code >= 1000 && lone.code <= 1014, String(lone.code));
    assert.ok(lone.parseError);
    ws.setValue('C1', { kind: ValueKind.Text, text: '=A1' });
    ws.setText('C2', '=A1');
    assert.equal(ws.getFormula('C1'), null);
    assert.equal(ws.getValue('C1'), '=A1');
    assert.equal(ws.getValue('C2'), '=A1');
    ws.setValue('C3', new ExcelError(ErrorCode.NA));
    assert.deepEqual(ws.getTypedValue('C3'), { kind: ValueKind.Error, errorCode: ErrorCode.NA });
    for (const bad of [undefined, 1n, Symbol('x'), [1], {}, { kind: ValueKind.Array, rows: 0, cols: 0, cells: [] }]) {
      assert.throws(() => ws.setValue('D1', bad), TypeError);
    }
  });
});

test('setValues stops at a formula that does not parse and names its cell', () => {
  withSheet((ws) => {
    const error = thrown(() =>
      ws.setValues('B2', [
        [1, '=1+1'],
        ['=(', 4],
      ]),
    );
    assert.ok(error.parseError);
    assert.match(error.context, /(?:^| )row=2(?: |$)/);
    assert.match(error.context, /(?:^| )col=1(?: |$)/);
    assert.deepEqual(ws.getValues('B2:C3'), [
      [1, 2],
      [null, null],
    ]);
  });
});

test('ExcelError holds an ErrorCode and spells it as Excel displays it', () => {
  const error = new ExcelError(ErrorCode.Div0);
  assert.equal(error.code, ErrorCode.Div0);
  assert.equal(error.toString(), '#DIV/0!');
  assert.equal(`${new ExcelError(ErrorCode.NA)}`, '#N/A');
  assert.ok(Object.isFrozen(error));
  assert.throws(() => new ExcelError(99), RangeError);
  assert.throws(() => new ExcelError('#N/A'), TypeError);
  assert.equal(mod.ExcelError, ExcelError);
});

test('enumName names an ordinal of any table and is the same function as the named export', () => {
  assert.equal(formulon.enumName(formulon.CfRuleType, 3), 'DataBar');
  assert.equal(formulon.enumName(ParseErrorCode, ParseErrorCode.EntryDoubleEquals), 'EntryDoubleEquals');
  assert.equal(formulon.enumName(formulon.IconSetName, 16), 'FiveQuarters');
  assert.equal(formulon.enumName(formulon.ComparisonOperator, 99), null);
  assert.equal(mod.enumName, formulon.enumName);
  assert.throws(() => formulon.enumName(null, 0), TypeError);
});

test('the ordinal tables are frozen members and named exports', () => {
  for (const name of ORDINAL_TABLES) {
    const values = formulon[name];
    assert.ok(Object.isFrozen(values), name);
    assert.equal(mod[name], values, name);
    assert.deepEqual(
      Object.values(values),
      Object.keys(values).map((_, i) => i),
      `${name} ordinals are dense from 0`,
    );
  }
  assert.equal(Object.keys(ParseErrorCode).length, 37);
  assert.equal(formulon.CfRuleType.UniqueValues, 17);
  assert.equal(formulon.ValidationType.Custom, 7);
});

test('the Formulon object has the typed evaluation and the new helpers', () => {
  assert.ok(Object.isFrozen(formulon));
  for (const name of ['ExcelError', 'enumName', 'evaluateTypedFormula', ...ORDINAL_TABLES]) {
    assert.ok(Object.hasOwn(formulon, name), name);
  }
  assert.equal(formulon.FormulonError, mod.FormulonError);
  assert.deepEqual(formulon.convertCivilToSerial({ year: 2024, month: 3, day: 1 }), 45352);
  assert.equal(formulon.convertCivilToSerial({ year: 2024, month: 3, day: 1, hour: 12 }), 45352.5);
});
