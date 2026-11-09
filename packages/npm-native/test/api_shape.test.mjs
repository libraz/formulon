// The shape of the public surface: the createFormulon entry, FormulonError
// throws, the Value union, setValue dispatch and the A1 Sheet view.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

import { memberNames, prototypeKeys } from '../../npm/test/dts_members.mjs';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  if (!pkg.main) throw new Error('package.json is missing "main"');
  const entryPath = path.resolve(pkgRoot, pkg.main);
  const mod = await import(pathToFileURL(entryPath).href);
  const common = await import(pathToFileURL(path.join(path.dirname(entryPath), 'common.mjs')).href);
  return { mod, common };
}

const loaded = loadPackage();
let formulonPromise;
function getFormulon() {
  formulonPromise ??= loaded.then(({ mod }) => mod.createFormulon());
  return formulonPromise;
}

// Members of the Formulon object that are not value tables.
const NON_TABLE_MEMBERS = new Set([
  'Workbook',
  'FormulonError',
  'mergeFunctionMetadata',
  'evaluateTypedFormula',
  'toNative',
  'convertCivilToSerial',
  'convertSerialToCivil',
  'evaluateFormula',
  'getVersion',
  'getStatusName',
  'getFunctionMetadata',
  'listFunctionNames',
  'localizeFunctionName',
  'canonicalizeFunctionName',
  'localizeFormula',
  'canonicalizeFormula',
  'getLocaleFacts',
  'listProfileIds',
  'compareText',
  'getCollationClasses',
  'getImageInfo',
  'setLogMinLevel',
  'setLogSink',
]);

function expectFormulonError(fn, code) {
  return assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, code, `code ${e.code} (${e.message})`);
    assert.equal(typeof e.message, 'string');
    assert.equal(typeof e.context, 'string');
    return true;
  });
}

async function withWorkbook(fn) {
  const formulon = await getFormulon();
  const wb = formulon.Workbook.createDefault();
  try {
    return await fn(wb, formulon);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

// ---- Entry --------------------------------------------------------------

test('createFormulon is both the default and a named export', async () => {
  const { mod } = await loaded;
  assert.equal(typeof mod.createFormulon, 'function');
  assert.equal(mod.default, mod.createFormulon);
});

test('the resolved object is frozen and has exactly FORMULON_MEMBERS', async () => {
  const { common } = await loaded;
  const formulon = await getFormulon();
  assert.ok(Object.isFrozen(formulon));
  assert.deepEqual(Object.keys(formulon).sort(), [...common.FORMULON_MEMBERS].sort());
  for (const name of ['getLastErrorMessage', 'getLastErrorContext', 'HEAPU8', 'preRun', 'Sheet']) {
    assert.equal(name in formulon, false, name);
  }
});

test('named exports are the value tables, FormulonError, toNative, mergeFunctionMetadata and createFormulon only', async () => {
  const { mod, common } = await loaded;
  const tables = common.FORMULON_MEMBERS.filter((name) => !NON_TABLE_MEMBERS.has(name));
  const helpers = ['toNative', 'mergeFunctionMetadata'];
  const expected = ['createFormulon', 'default', 'FormulonError', ...helpers, ...tables].sort();
  assert.deepEqual(Object.keys(mod).sort(), expected);
  const formulon = await getFormulon();
  for (const name of [...tables, ...helpers]) assert.equal(mod[name], formulon[name], name);
});

test('createFormulon accepts and ignores options', async () => {
  const { mod, common } = await loaded;
  const formulon = await mod.createFormulon({ foo: 1, locateFile: () => 'unused', print() {}, printErr() {} });
  assert.deepEqual(Object.keys(formulon).sort(), [...common.FORMULON_MEMBERS].sort());
});

// ---- FormulonError ------------------------------------------------------

test('FormulonError is one class shared by the named export and the Formulon object', async () => {
  const { mod } = await loaded;
  const formulon = await getFormulon();
  assert.equal(formulon.FormulonError, mod.FormulonError);
  const e = new mod.FormulonError(mod.StatusCode.NotFound, 'missing');
  assert.ok(e instanceof Error);
  assert.equal(e.name, 'FormulonError');
  assert.equal(e.code, mod.StatusCode.NotFound);
  assert.equal(e.message, 'missing');
  assert.equal(e.context, '');
});

test('Workbook.load of a corrupt ZIP throws FormulonError(IoZipCorrupt)', async () => {
  const { mod } = await loaded;
  const formulon = await getFormulon();
  const bytes = new Uint8Array(16);
  bytes.set([0x50, 0x4b, 0x03, 0x04]);
  expectFormulonError(() => formulon.Workbook.load(bytes), mod.StatusCode.IoZipCorrupt);
  assert.throws(
    () => formulon.Workbook.load(bytes),
    (e) => e instanceof mod.FormulonError && !/fm_|::/.test(e.message) && !/fm_/.test(e.context),
  );
  // The failure is consumed by the call that raised it.
  const wb = formulon.Workbook.createDefault();
  assert.equal(wb.listSheets().length, 1);
  wb.dispose();
});

test('Workbook.load of empty bytes throws FormulonError(InvalidArgument)', async () => {
  const formulon = await getFormulon();
  expectFormulonError(() => formulon.Workbook.load(new Uint8Array(0)), formulon.StatusCode.InvalidArgument);
  assert.throws(() => formulon.Workbook.load('not bytes'), TypeError);
});

test('calls after dispose throw BindingInvalidHandle; disposed answers', async () => {
  const formulon = await getFormulon();
  const wb = formulon.Workbook.createDefault();
  const view = wb.sheet(0);
  assert.equal(wb.disposed, false);
  wb.dispose();
  assert.equal(wb.disposed, true);
  wb.dispose();
  const code = formulon.StatusCode.BindingInvalidHandle;
  expectFormulonError(() => wb.listSheets().length, code);
  expectFormulonError(() => wb.sheet(0).getTypedValue(0, 0), code);
  expectFormulonError(() => wb.sheet(0).setNumber(0, 0, 1), code);
  expectFormulonError(() => view.getTypedValue('A1'), code);
  expectFormulonError(() => view.setValue('A1', 1), code);
  expectFormulonError(() => view.getFormula('A1'), code);
  expectFormulonError(() => view.setFormula('A1', '=1'), code);
});

test('out-of-range indices and missing natural keys throw by code', async () => {
  await withWorkbook((wb, formulon) => {
    const { StatusCode } = formulon;
    expectFormulonError(() => wb.sheet(1).name, StatusCode.SheetIndexOutOfRange);
    expectFormulonError(() => wb.ooxml.getFont(wb.ooxml.listFonts().length), StatusCode.IndexOutOfRange);
    expectFormulonError(() => wb.removeDefinedName('nope', null), StatusCode.NotFound);
    assert.equal(wb.getDefinedName('nope', null), null);
  });
});

test('save takes an optional format; saveAs and isValid are gone', async () => {
  await withWorkbook((wb, formulon) => {
    assert.ok(wb.save() instanceof Uint8Array);
    assert.ok(wb.save(formulon.WorkbookFormat.Xlsx) instanceof Uint8Array);
    assert.ok(wb.save(formulon.WorkbookFormat.Xlsb) instanceof Uint8Array);
    expectFormulonError(() => wb.save(99), formulon.StatusCode.InvalidArgument);
    const diagnostics = wb.saveWithDiagnostics();
    assert.ok(diagnostics.bytes instanceof Uint8Array);
    assert.equal('status' in diagnostics, false);
    assert.equal(wb.saveAs, undefined);
    assert.equal(wb.isValid, undefined);
  });
});

test('no read method returns a status envelope', async () => {
  await withWorkbook((wb) => {
    const names = Object.getOwnPropertyNames(Object.getPrototypeOf(wb)).filter((n) => /^(get|list)[A-Z]/.test(n));
    let called = 0;
    for (const name of names) {
      for (const args of [[], [0], [0, 0], [0, 0, 0]]) {
        let result;
        try {
          result = wb[name](...args);
        } catch {
          continue;
        }
        called += 1;
        if (result !== null && typeof result === 'object') {
          assert.equal(Object.hasOwn(result, 'status'), false, `${name}(${args}) returned a status key`);
        }
        break;
      }
    }
    assert.ok(called > 20, `only ${called} read methods could be called`);
  });
});

// ---- Callbacks ----------------------------------------------------------

function iterativeWorkbook(formulon) {
  const wb = formulon.Workbook.createDefault();
  wb.updateIterative({ enabled: true, maxIterations: 100, maxChange: 0.0001 });
  wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
  return wb;
}

test('an exception thrown by a progress callback is rethrown from recalc', async () => {
  const formulon = await getFormulon();
  const wb = iterativeWorkbook(formulon);
  try {
    const thrown = new Error('stop');
    wb.setIterativeProgress(() => {
      throw thrown;
    });
    assert.throws(
      () => wb.recalc(),
      (caught) => caught === thrown,
    );
    // The aborted solve leaves its last iterate in place; re-entering the
    // formula marks it dirty so the next recalc solves it again.
    wb.setIterativeProgress(null);
    wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
    wb.recalc();
    const value = wb.sheet(0).getTypedValue(0, 0);
    assert.equal(value.kind, formulon.ValueKind.Number);
    assert.ok(Math.abs(value.number - 10) < 0.001);
  } finally {
    wb.dispose();
  }
});

test('a FormulonError raised inside a progress callback reaches the caller unchanged', async () => {
  const formulon = await getFormulon();
  const wb = iterativeWorkbook(formulon);
  const other = formulon.Workbook.createDefault();
  try {
    let raised;
    wb.setIterativeProgress(() => {
      try {
        other.ooxml.getCellXf(99999);
      } catch (e) {
        raised = e;
        throw e;
      }
    });
    assert.throws(
      () => wb.recalc(),
      (caught) => caught === raised && caught.code === formulon.StatusCode.IndexOutOfRange,
    );
    wb.setIterativeProgress(null);
    assert.equal(other.listSheets().length, 1);
  } finally {
    wb.dispose();
    other.dispose();
  }
});

// ---- Value --------------------------------------------------------------

test('each scalar kind carries only its own fields', async () => {
  await withWorkbook((wb, { ValueKind, ErrorCode }) => {
    wb.sheet(0).setNumber(0, 0, 42);
    wb.sheet(0).setBool(0, 1, true);
    wb.sheet(0).setText(0, 2, 'hi');
    wb.sheet(0).setFormula(0, 3, '=1/0');
    wb.recalc();
    assert.deepStrictEqual({ ...wb.sheet(0).getTypedValue(0, 0) }, { kind: ValueKind.Number, number: 42 });
    assert.deepStrictEqual({ ...wb.sheet(0).getTypedValue(0, 1) }, { kind: ValueKind.Bool, boolean: true });
    assert.deepStrictEqual({ ...wb.sheet(0).getTypedValue(0, 2) }, { kind: ValueKind.Text, text: 'hi' });
    assert.deepStrictEqual(
      { ...wb.sheet(0).getTypedValue(0, 3) },
      { kind: ValueKind.Error, errorCode: ErrorCode.Div0 },
    );
    assert.deepStrictEqual({ ...wb.sheet(0).getTypedValue(0, 4) }, { kind: ValueKind.Blank });
    assert.deepEqual(Object.keys(wb.sheet(0).getTypedValue(0, 3)).sort(), ['errorCode', 'kind']);
  });
});

test('evaluateTypedFormula returns a bare Value', async () => {
  const { evaluateTypedFormula, ValueKind } = await getFormulon();
  assert.deepStrictEqual({ ...evaluateTypedFormula('=SUM(1,2,3)') }, { kind: ValueKind.Number, number: 6 });
});

test('setValue(getValue(...)) round-trips every scalar kind', async () => {
  await withWorkbook((wb) => {
    wb.sheet(0).setNumber(0, 0, -1.5);
    wb.sheet(0).setBool(0, 1, false);
    wb.sheet(0).setText(0, 2, '=not a formula');
    wb.sheet(0).setFormula(0, 3, '=NA()');
    wb.recalc();
    for (let col = 0; col <= 4; col += 1) {
      const before = wb.sheet(0).getTypedValue(0, col);
      wb.sheet(0).setValue(1, col, before);
      assert.deepStrictEqual({ ...wb.sheet(0).getTypedValue(1, col) }, { ...before }, `column ${col}`);
    }
  });
});

// ---- setValue -----------------------------------------------------------

test('setValue dispatches each primitive to the matching setter', async () => {
  await withWorkbook((wb, { ValueKind, ErrorCode }) => {
    const rows = [
      [0, (r) => wb.sheet(0).setNumber(r, 0, 0)],
      [-0, (r) => wb.sheet(0).setNumber(r, 0, -0)],
      [2 ** 53, (r) => wb.sheet(0).setNumber(r, 0, 2 ** 53)],
      [2 ** 63, (r) => wb.sheet(0).setNumber(r, 0, 2 ** 63)],
      ['=A1', (r) => wb.sheet(0).setFormula(r, 0, '=A1')],
      ['', (r) => wb.sheet(0).setText(r, 0, '')],
      [true, (r) => wb.sheet(0).setBool(r, 0, true)],
      [false, (r) => wb.sheet(0).setBool(r, 0, false)],
      [null, (r) => wb.sheet(0).setBlank(r, 0)],
      [{ kind: ValueKind.Blank }, (r) => wb.sheet(0).setBlank(r, 0)],
      [{ kind: ValueKind.Number, number: 3 }, (r) => wb.sheet(0).setNumber(r, 0, 3)],
      [{ kind: ValueKind.Bool, boolean: true }, (r) => wb.sheet(0).setBool(r, 0, true)],
      [{ kind: ValueKind.Text, text: 't' }, (r) => wb.sheet(0).setText(r, 0, 't')],
      [{ kind: ValueKind.Error, errorCode: ErrorCode.Ref }, (r) => wb.sheet(0).setError(r, 0, ErrorCode.Ref)],
    ];
    rows.forEach(([value, direct], index) => {
      const row = index * 2;
      wb.sheet(0).setNumber(row, 1, 99);
      wb.sheet(0).setValue(row, 1, value);
      direct(row + 1);
      assert.deepStrictEqual(
        { ...wb.sheet(0).getTypedValue(row, 1) },
        { ...wb.sheet(0).getTypedValue(row + 1, 0) },
        `row ${index}`,
      );
      assert.equal(
        wb.sheet(0).getFormula(row, 1),
        wb.sheet(0).getFormula(row + 1, 0),
        `row ${index} stored a different formula`,
      );
    });
  });
});

test('setValue rejects non-finite numbers as the engine does', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    for (const value of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
      expectFormulonError(() => wb.sheet(0).setValue(0, 0, value), StatusCode.InvalidArgument);
    }
  });
});

test('setValue throws TypeError for values it cannot store', async () => {
  await withWorkbook((wb, { ValueKind }) => {
    const bad = [
      undefined,
      1n,
      Symbol('s'),
      {},
      [],
      [1],
      () => 1,
      { kind: ValueKind.Array },
      { kind: ValueKind.Ref },
      { kind: ValueKind.Lambda },
      { kind: 99 },
      { kind: ValueKind.Number },
      { kind: ValueKind.Number, number: '1' },
      { kind: ValueKind.Bool, boolean: 1 },
      { kind: ValueKind.Text },
      { kind: ValueKind.Error },
    ];
    for (const value of bad) {
      assert.throws(() => wb.sheet(0).setValue(0, 0, value), TypeError, String(value?.kind ?? typeof value));
    }
    assert.throws(() => wb.sheet(0).setValue(0, 0), TypeError);
  });
});

// ---- Sheet view ---------------------------------------------------------

test('sheet() resolves names and indices and rejects the rest', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    wb.addSheet('Data');
    const byName = wb.sheet('data');
    assert.equal(byName.index, 1);
    assert.equal(byName.name, 'Data');
    assert.equal(byName.workbook, wb);
    assert.equal(wb.sheet(0).name, 'Sheet1');
    expectFormulonError(() => wb.sheet('nope'), StatusCode.NotFound);
    expectFormulonError(() => wb.sheet(''), StatusCode.NotFound);
    expectFormulonError(() => wb.sheet(2), StatusCode.SheetIndexOutOfRange);
    expectFormulonError(() => wb.sheet(-1), StatusCode.SheetIndexOutOfRange);
    assert.throws(() => wb.sheet(2 ** 32), RangeError);
    assert.throws(() => wb.sheet(1.5), RangeError);
    assert.throws(() => wb.sheet(true), TypeError);
    assert.throws(() => wb.sheet(), TypeError);
    assert.throws(() => wb.sheet(null), TypeError);
  });
});

test('a Sheet keeps its index and reads its name live', async () => {
  await withWorkbook((wb) => {
    const view = wb.sheet('Sheet1');
    wb.sheet(0).rename('Renamed');
    assert.equal(view.name, 'Renamed');
    assert.equal(view.index, 0);
  });
});

test('Sheet getValue / setValue / getFormula / setFormula take one A1 cell', async () => {
  await withWorkbook((wb, { ValueKind }) => {
    const view = wb.sheet(0);
    view.setValue('B3', 5);
    assert.deepStrictEqual({ ...wb.sheet(0).getTypedValue(2, 1) }, { kind: ValueKind.Number, number: 5 });
    assert.deepStrictEqual({ ...view.getTypedValue('B3') }, { kind: ValueKind.Number, number: 5 });
    view.setValue('C1', '=B3');
    assert.equal(view.getFormula('C1'), '=B3');
    view.setValue('C2', "'=B3");
    assert.equal(view.getFormula('C2'), null);
    assert.equal(wb.sheet(0).getFormula(0, 2), '=B3');
    view.setFormula('D1', '=B3*2');
    assert.equal(wb.sheet(0).getFormula(0, 3), '=B3*2');
    assert.equal(view.getFormula('D1'), '=B3*2');
    wb.recalc();
    assert.deepStrictEqual({ ...view.getTypedValue('D1') }, { kind: ValueKind.Number, number: 10 });
    view.setValue('B3', null);
    assert.deepStrictEqual({ ...view.getTypedValue('B3') }, { kind: ValueKind.Blank });
  });
});

test('Sheet addresses equivalent spellings of one cell alike', async () => {
  await withWorkbook((wb, { ValueKind }) => {
    const view = wb.sheet(0);
    wb.sheet(0).setNumber(2, 1, 7);
    for (const a1 of ['B3', 'b3', '$B$3', '$b3', 'B$3']) {
      assert.deepStrictEqual({ ...view.getTypedValue(a1) }, { kind: ValueKind.Number, number: 7 }, a1);
    }
    wb.sheet(0).setNumber(0, 0, 1);
    for (const a1 of ['A01', 'a0000001']) {
      assert.deepStrictEqual({ ...view.getTypedValue(a1) }, { kind: ValueKind.Number, number: 1 }, a1);
    }
    wb.sheet(0).setNumber(1048575, 16383, 3);
    assert.deepStrictEqual({ ...view.getTypedValue('XFD1048576') }, { kind: ValueKind.Number, number: 3 });
  });
});

// Text that is not an A1 reference fails as a parse error (1004); a well-formed range, span or sheet-qualified
// address is refused as an argument (InvalidArgument).
const A1_CELL = /^\$?([A-Za-z]{1,3})\$?(\d+)$/;
function isCell(text) {
  const m = A1_CELL.exec(text);
  if (m === null) return false;
  let col = 0;
  for (const ch of m[1].toUpperCase()) col = col * 26 + ch.charCodeAt(0) - 64;
  const row = Number(m[2]);
  return col <= 16384 && row >= 1 && row <= 1048576;
}
function refusalStatus(a1, StatusCode) {
  const body = a1.replace(/^Sheet\d+!/, '');
  const parts = body.split(':');
  const wholeColumns = parts.length === 2 && parts.every((p) => /^\$?[A-Za-z]{1,3}$/.test(p));
  const wholeRows = parts.length === 2 && parts.every((p) => /^\$?\d+$/.test(p));
  const wellFormed = parts.length <= 2 && parts.every(isCell);
  return wellFormed || wholeColumns || wholeRows ? StatusCode.InvalidArgument : StatusCode.ParserInvalidReference;
}

test('Sheet rejects addresses that are not one unqualified cell', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const view = wb.sheet(0);
    const rejected = ['', '   ', 'A1:B2', 'Sheet2!A1', 'Sheet1!A1', 'R1C1', 'XFE1', 'A0', 'A1048577', 'B:B', '3:3'];
    rejected.push('Ａ１', 'Ａ1', '😀', 'A\uD800', ' A1', 'A1 ', 'A1\n');
    for (const a1 of rejected) {
      expectFormulonError(() => view.getTypedValue(a1), refusalStatus(a1, StatusCode));
      expectFormulonError(() => view.setValue(a1, 1), refusalStatus(a1, StatusCode));
    }
    assert.throws(() => view.getTypedValue(1), TypeError);
  });
});

// Pairwise model over column x row x `$` placement x case x decoration.
const PAIRWISE = [
  ['P1', '$z0', null],
  ['P2', 'Sheet1!xfe$9', null],
  ['P3', '$XFE1:B2', null],
  ['P4', ' Z$01 ', null],
  ['P5', 'aa$1048576:B2', null],
  ['P6', ' $a$1 ', null],
  ['P7', 'Sheet1!XFD1048576', null],
  ['P8', '$xfd$1048577', null],
  ['P9', 'Sheet1!z01', null],
  ['P10', 'Sheet1!$A1048577', null],
  ['P11', ' xfe1048577 ', null],
  ['P12', 'AA9', [8, 26]],
  ['P13', '$XFE$0:B2', null],
  ['P14', ' $aa01 ', null],
  ['P15', 'a01:B2', null],
  ['P16', 'A$1', [0, 0]],
  ['P17', '$Z$1048576', [1048575, 25]],
  ['P18', 'Z$1048577:B2', null],
  ['P19', ' $xfe1048576 ', null],
  ['P20', ' XFD$0 ', null],
  ['P21', 'Sheet1!$aa$1', null],
  ['P22', '$xfd1:B2', null],
  ['P23', 'Sheet1!a0', null],
  ['P24', 'XFD$9:B2', null],
  ['P25', 'XFE01', null],
  ['P26', 'Sheet1!aa$1048577', null],
  ['P27', ' $z$9 ', null],
  ['P28', 'Sheet1!z1', null],
  ['P29', ' AA0 ', null],
  ['P30', 'A$1048576:B2', null],
  ['P31', 'Sheet1!$xfd$01', null],
  ['P32', '$a9', [8, 0]],
];

for (const [id, a1, cell] of PAIRWISE) {
  test(`A1 acceptance ${id}: ${JSON.stringify(a1)}`, async () => {
    await withWorkbook((wb, { StatusCode, ValueKind }) => {
      const view = wb.sheet(0);
      if (cell === null) {
        expectFormulonError(() => view.getTypedValue(a1), refusalStatus(a1, StatusCode));
        expectFormulonError(() => view.setValue(a1, 1), refusalStatus(a1, StatusCode));
        return;
      }
      const [row, col] = cell;
      view.setValue(a1, 7);
      assert.deepStrictEqual({ ...wb.sheet(0).getTypedValue(row, col) }, { kind: ValueKind.Number, number: 7 });
      assert.deepStrictEqual({ ...view.getTypedValue(a1) }, { kind: ValueKind.Number, number: 7 });
    });
  });
}

// ---- Public surface (S2) --------------------------------------------------

test('S2: Workbook, Sheet and the statics are exactly the declared members', async () => {
  const dts = path.join(pkgRoot, 'dist', 'index.d.ts');
  const declared = (name) => [...memberNames(dts, name)].sort();
  await withWorkbook((wb, formulon) => {
    assert.deepEqual(prototypeKeys(Object.getPrototypeOf(wb)), declared('Workbook'));
    assert.deepEqual(prototypeKeys(Object.getPrototypeOf(wb.sheet(0))), declared('Sheet'));
    const statics = Object.getOwnPropertyNames(formulon.Workbook).filter(
      (k) => !['length', 'name', 'prototype'].includes(k),
    );
    assert.deepEqual(statics.sort(), declared('WorkbookCtor'));
    for (const name of ['delete', 'clone', 'isDeleted', 'getSheetCount', 'getSheetName', 'getSheetIndex']) {
      assert.equal(name in wb, false, name);
    }
    assert.throws(() => new formulon.Workbook(), TypeError);
  });
});

// ---- Sheet handles (H1-H15) -----------------------------------------------

// H11: the engine's own sheet list, read back through a reload, matches the handles.
function assertSheetsInSync(formulon, wb) {
  const handles = wb.listSheets();
  const reloaded = formulon.Workbook.load(wb.save());
  try {
    assert.deepEqual(
      reloaded.listSheets().map((s) => s.name),
      handles.map((s) => s.name),
    );
  } finally {
    reloaded.dispose();
  }
  for (const [i, sheet] of handles.entries()) assert.equal(sheet.index, i);
}

function threeSheets(formulon) {
  const wb = formulon.Workbook.createDefault();
  wb.sheet(0).rename('Data');
  wb.addSheet('Calc');
  wb.addSheet('Notes');
  return wb;
}

test('H1: one handle per sheet, by index or by case-insensitive name', async () => {
  await withWorkbook((wb) => {
    wb.sheet(0).rename('Data');
    const ws = wb.sheet(0);
    assert.equal(wb.sheet(0), ws);
    assert.equal(wb.sheet('Data'), ws);
    assert.equal(wb.sheet('data'), ws);
    assert.equal(wb.sheet(ws), ws);
    assert.equal(wb.listSheets()[0], ws);
  });
});

test('H2 / H14: copySheet before a handle shifts it and returns the copy; at the end it appends', async () => {
  const formulon = await getFormulon();
  const wb = threeSheets(formulon);
  try {
    const calc = wb.sheet('Calc');
    calc.setNumber(0, 0, 7);
    const copy = wb.copySheet(calc, { insertBefore: 0, newName: 'Copy' });
    assert.equal(copy.index, 0);
    assert.equal(copy.name, 'Copy');
    assert.equal(wb.sheet(0), copy);
    assert.equal(calc.index, 2);
    assert.equal(calc.getTypedValue(0, 0).number, 7);
    const last = wb.copySheet(calc, { insertBefore: wb.listSheets().length, newName: 'Last' });
    assert.equal(last.index, 4);
    expectFormulonError(() => wb.copySheet(calc, { insertBefore: 6 }), formulon.StatusCode.SheetIndexOutOfRange);
    assertSheetsInSync(formulon, wb);
  } finally {
    wb.dispose();
  }
});

test('H3 / H4 / H14: moveSheet and renameSheet keep the handle on its sheet', async () => {
  const formulon = await getFormulon();
  const wb = threeSheets(formulon);
  try {
    const data = wb.sheet('Data');
    data.setNumber(0, 0, 1);
    wb.sheet(data).move(3);
    assert.equal(data.index, 2);
    assert.equal(data.getTypedValue(0, 0).number, 1);
    wb.sheet(data).move(data.index);
    assert.equal(data.index, 2);
    assertSheetsInSync(formulon, wb);
    wb.sheet('data').rename('Renamed');
    assert.equal(data.name, 'Renamed');
    assert.equal(wb.sheet('Renamed'), data);
    assertSheetsInSync(formulon, wb);
  } finally {
    wb.dispose();
  }
});

test('H5 / H6 / H15: a removed handle throws NotFound; later handles shift down', async () => {
  const formulon = await getFormulon();
  const wb = threeSheets(formulon);
  try {
    const [data, calc, notes] = wb.listSheets();
    notes.setNumber(0, 0, 3);
    wb.removeSheet(calc);
    const { NotFound } = formulon.StatusCode;
    expectFormulonError(() => calc.index, NotFound);
    expectFormulonError(() => calc.name, NotFound);
    expectFormulonError(() => calc.setNumber(0, 0, 1), NotFound);
    assert.equal(data.index, 0);
    assert.equal(notes.index, 1);
    assert.equal(wb.sheet(1), notes);
    assert.equal(wb.sheet(1).getTypedValue(0, 0).number, 3);
    assertSheetsInSync(formulon, wb);
  } finally {
    wb.dispose();
  }
});

test('H7 / H13: after dispose, Sheet members throw BindingInvalidHandle', async () => {
  const formulon = await getFormulon();
  const wb = formulon.Workbook.createDefault();
  const ws = wb.sheet(0);
  wb.dispose();
  const code = formulon.StatusCode.BindingInvalidHandle;
  expectFormulonError(() => ws.index, code);
  expectFormulonError(() => ws.name, code);
  expectFormulonError(() => ws.setNumber(0, 0, 1), code);
});

test('H8: sheet() rejects a missing name, an out-of-range index and keys of the wrong type', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    expectFormulonError(() => wb.sheet('nope'), StatusCode.NotFound);
    expectFormulonError(() => wb.sheet(99), StatusCode.SheetIndexOutOfRange);
    expectFormulonError(() => wb.sheet(-1), StatusCode.SheetIndexOutOfRange);
    assert.throws(() => wb.sheet(1.5), RangeError);
    assert.throws(() => wb.sheet(true), TypeError);
    assert.throws(() => wb.sheet(null), TypeError);
  });
});

test('H9: showPivotDetails returns the new sheet and the existing handles follow their sheets', async () => {
  const formulon = await getFormulon();
  const { wb, pivot } = pivotWorkbook(formulon);
  try {
    const ws = wb.sheet(0);
    // The grand total's value cell, not its label.
    const cell = ws
      .pivot(pivot)
      .getLayout()
      .cells.find((c) => c.kind === formulon.PivotCellKind.GrandTotal && typeof c.value === 'number');
    assert.ok(cell, 'the pivot has a grand total value cell');
    const details = ws.pivot(pivot).showDetails(cell.row, cell.col);
    assert.equal(details.index, 0);
    assert.equal(wb.sheet(0), details);
    assert.equal(ws.index, 1);
    assert.equal(ws.listPivots().length, 1);
    assertSheetsInSync(formulon, wb);
  } finally {
    wb.dispose();
  }
});

test('H10: handles of two workbooks are never the same object', async () => {
  const formulon = await getFormulon();
  const a = formulon.Workbook.createDefault();
  const b = formulon.Workbook.createDefault();
  try {
    assert.notEqual(a.sheet(0), b.sheet(0));
  } finally {
    a.dispose();
    b.dispose();
  }
});

test('a SheetRef is a Sheet, an index or a name; a Sheet of another workbook is InvalidArgument', async () => {
  const formulon = await getFormulon();
  const wb = threeSheets(formulon);
  const other = formulon.Workbook.createDefault();
  try {
    wb.setActiveSheet(wb.sheet('Calc'));
    assert.equal(wb.activeSheet(), wb.sheet('Calc'));
    wb.setActiveSheet('notes');
    assert.equal(wb.activeSheet().name, 'Notes');
    wb.setActiveSheet(0);
    assert.equal(wb.activeSheet().index, 0);
    wb.removeSheet('Notes');
    wb.removeSheet(wb.sheet(1));
    assert.equal(wb.listSheets().length, 1);
    expectFormulonError(() => wb.removeSheet(other.sheet(0)), formulon.StatusCode.InvalidArgument);
    expectFormulonError(() => wb.sheet(other.sheet(0)), formulon.StatusCode.InvalidArgument);
    assertSheetsInSync(formulon, wb);
  } finally {
    wb.dispose();
    other.dispose();
  }
});

// ---- Argument guards (E1) -------------------------------------------------

test('E1: a positional guard names the facade method and counts its arguments', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    assert.throws(() => ws.setNumber(true, 0, 1), {
      name: 'TypeError',
      message: 'Sheet.setNumber argument 1 must be a primitive number',
    });
    expectFormulonError(() => ws.setNumber(0, -1, 1), StatusCode.InvalidArgument);
    assert.throws(() => ws.setNumber(0, 1.5, 1), RangeError);
    assert.throws(() => ws.setNumber(0, 0), {
      name: 'TypeError',
      message: 'Sheet.setNumber argument 3 must be a primitive number',
    });
    assert.throws(
      () => ws.setPrintTitles(),
      (e) => e instanceof TypeError && /^Sheet\.setPrintTitles argument 1 /.test(e.message),
    );
    assert.throws(() => wb.addSheet(5), {
      name: 'TypeError',
      message: 'Workbook.addSheet argument 1 must be a primitive string',
    });
    assert.throws(() => ws.setValue(0, 0, []), /^TypeError: Sheet\.setValue argument 3 /);
    assert.throws(() => wb.sheet({}), /^TypeError: Workbook\.sheet argument 1 /);
  });
});

// ---- Round trips of the reshaped setters ----------------------------------

test('round trip: every reshaped setter takes what its getter returns', async () => {
  const formulon = await getFormulon();
  const { wb, pivot } = pivotWorkbook(formulon);
  try {
    const ws = wb.sheet(0);
    const same = (get, set) => {
      const before = get();
      set(before);
      assert.deepEqual(get(), before);
    };
    wb.updateIterative({ enabled: true, maxIterations: 50, maxChange: 0.01 });
    same(
      () => wb.getIterative(),
      (v) => wb.updateIterative(v),
    );
    wb.setPinnedNow({ year: 2026, month: 11, day: 6, hour: 9, minute: 30, second: 15 });
    same(
      () => wb.getPinnedNow(),
      (v) => wb.setPinnedNow(v),
    );
    ws.setPrintTitles({ rows: { first: 0, last: 1 }, columns: { first: 0, last: 0 } });
    same(
      () => ws.getPrintTitles(),
      (v) => ws.setPrintTitles(v),
    );
    ws.pivot(pivot).updateGrandTotals({ rows: false, columns: true });
    same(
      () => ws.pivot(pivot).getGrandTotals(),
      (v) => ws.pivot(pivot).updateGrandTotals(v),
    );
    ws.setComment(5, 5, { author: 'Ann', text: 'note' });
    same(
      () => ws.getComment(5, 5),
      (v) => ws.setComment(5, 5, v),
    );
    wb.setWorkbookProtection({ lockStructure: true, workbookLegacyPasswordHash: 'CF1A' });
    same(
      () => wb.getWorkbookProtection(),
      (v) => wb.setWorkbookProtection(v),
    );
    ws.setSheetProtection({ allowFormatCells: true, legacyPasswordHash: 'CC1A' });
    same(
      () => ws.getSheetProtection(),
      (v) => ws.setSheetProtection(v),
    );
    ws.setText(6, 0, '漢字');
    ws.setCellPhoneticRuns(6, 0, [{ start: 0, end: 2, text: 'かんじ' }]);
    same(
      () => ws.getCellPhoneticRuns(6, 0),
      (v) => ws.setCellPhoneticRuns(6, 0, v),
    );
    same(
      () => wb.activeSheet(),
      (v) => wb.setActiveSheet(v),
    );
    wb.setThemeScriptFont({ scheme: formulon.FontScheme.Major, script: 'Zzzz', typeface: 'Face' });
    same(
      () => wb.getThemeScriptFont({ scheme: formulon.FontScheme.Major, script: 'Zzzz' }),
      (v) => wb.setThemeScriptFont(v),
    );
    const colors = wb.getThemeColors();
    wb.setThemeColors({
      colors: colors.colors.map((value) => `#${(value >>> 0).toString(16).toUpperCase().padStart(8, '0')}`),
    });
    assert.deepEqual(wb.getThemeColors().colors, colors.colors);
    wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$A$1', scope: 0, comment: 'c', hidden: true });
    same(
      () => wb.getDefinedName('Rate', 0),
      (v) => wb.setDefinedName(v),
    );
    wb.updateTableStyleDefaults({ tableStyle: 'TableStyleLight1' });
    same(
      () => wb.getTableStyleDefaults(),
      (v) => wb.updateTableStyleDefaults(v),
    );
    ws.pivot(pivot).updateOptions({ errorCaption: 'n/a', showError: true });
    same(
      () => ws.pivot(pivot).getOptions(),
      (v) => ws.pivot(pivot).updateOptions(v),
    );
    const cacheId = ws.pivot(pivot).getCache();
    same(
      () => wb.pivotCache(cacheId).getOptions(),
      (v) => wb.pivotCache(cacheId).updateOptions(v),
    );
    ws.updateFormatDefaults({ defaultColWidthChars: 12 });
    same(
      () => ws.getFormatDefaults(),
      (v) => ws.updateFormatDefaults(v),
    );
    ws.setPrintArea([{ firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 3 }]);
    same(
      () => ws.getPrintArea(),
      (v) => ws.setPrintArea(v),
    );
    ws.updateView({ viewMode: formulon.SheetViewMode.PageLayout });
    same(
      () => ws.getView().viewMode,
      (v) => ws.updateView({ viewMode: v }),
    );
    assert.equal(ws.getView().viewMode, formulon.SheetViewMode.PageLayout);
  } finally {
    wb.dispose();
  }
});

// ---- Absent values (A1-A11) -----------------------------------------------

test('A1: a value the workbook does not hold reads as null', async () => {
  const formulon = await getFormulon();
  const { wb, pivot } = pivotWorkbook(formulon);
  try {
    const ws = wb.sheet(0);
    ws.setText(9, 0, 'x');
    assert.equal(ws.getCellPhonetic(9, 0), null);
    assert.equal(wb.getFileName(), null);
    assert.equal(ws.ooxml.getAutoFilterXml(), null);
    assert.equal(ws.ooxml.getPageSetupXml(), null);
    assert.equal(ws.getPrintTitles(), null);
    ws.setPrintTitles({ rows: { first: 2, last: 3 }, columns: null });
    assert.deepEqual(ws.getPrintTitles(), { rows: { first: 2, last: 3 }, columns: null });

    wb.setDefinedName({ name: 'N', formula: '1' });
    assert.deepEqual((({ scope, comment }) => ({ scope, comment }))(wb.getDefinedName('N', null)), {
      scope: null,
      comment: null,
    });
    ws.setHyperlink(9, 1, { target: 'https://example.com/' });
    const link = ws.getHyperlink(9, 1);
    assert.deepEqual([link.location, link.display, link.tooltip], [null, null, null]);

    const view = ws.getView();
    assert.equal(view.viewMode, formulon.SheetViewMode.Normal);
    assert.equal(view.showGridLines, true);
    assert.equal(view.rightToLeft, false);
    assert.equal('visibility' in view, false);
    assert.equal(ws.getVisibility(), formulon.SheetVisibility.Visible);
    assert.equal(ws.getSheetProtection(), null);

    const field = ws.pivot(pivot).listFields()[0];
    assert.deepEqual([field.customName, field.numberFormat, field.groupBaseIndex], [null, null, null]);
    const dataField = ws.pivot(pivot).listDataFields()[0];
    assert.deepEqual([dataField.numberFormat, dataField.showAsBaseField, dataField.showAsBaseItem], [null, null, null]);
    const options = ws.pivot(pivot).getOptions();
    assert.deepEqual([options.errorCaption, options.altText, options.grandTotalCaption], [null, null, null]);
    const cacheId = ws.pivot(pivot).getCache();
    assert.equal(wb.pivotCache(cacheId).getOptions().missingItemsLimit, null);
    assert.equal(wb.pivotCache(cacheId).listFields()[0].formula, null);
    assert.equal(wb.pivotCache(wb.addPivotCache()).getWorksheetSource(), null);

    ws.addSparklineGroup({ sparklines: [{ source: 'Sheet1!A1:D1', location: 'F1' }] });
    assert.equal(ws.listSparklineGroups()[0].dateAxisRange, null);

    const font = wb.ooxml.getFont(wb.ooxml.addFont({ name: 'Arial', sizePt: 10 }));
    assert.deepEqual([font.bold, font.italic, font.family, font.charset], [null, null, null, null]);
    const xf = wb.ooxml.getCellXf(wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 0 }));
    assert.deepEqual([xf.horizontal, xf.wrapText, xf.locked, xf.hidden], [null, null, null, null]);
    ws.setRowHeight(20, 20, 30);
    const row = ws.listRowOverrides().find((r) => r.row === 20);
    assert.deepEqual([ws.ooxml.getRowXfIndex(20), row.hidden], [null, false]);
    const outcome = ws.validateValue(30, 30, { kind: formulon.ValueKind.Number, number: 1 });
    assert.deepEqual([outcome.valid, outcome.ruleId, outcome.errorStyle], [true, null, null]);
  } finally {
    wb.dispose();
  }
});

test('A2: setPinnedNow(null) releases the pin', async () => {
  await withWorkbook((wb) => {
    wb.setPinnedNow({ year: 2026, month: 1, day: 2, hour: 3, minute: 4, second: 5 });
    assert.notEqual(wb.getPinnedNow(), null);
    wb.setPinnedNow(null);
    assert.equal(wb.getPinnedNow(), null);
    wb.setPinnedNow(null);
  });
});

test('A3 / A9: setDefinedName with a null scope is workbook-wide and replaces the whole name', async () => {
  await withWorkbook((wb) => {
    wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$A$1', scope: null, comment: 'c', hidden: true });
    let name = wb.getDefinedName('Rate', null);
    assert.deepEqual([name.scope, name.comment, name.hidden], [null, 'c', true]);
    wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$B$1' });
    name = wb.getDefinedName('Rate', null);
    assert.deepEqual([name.formula, name.comment, name.hidden], ['=Sheet1!$B$1', null, false]);
    wb.removeDefinedName('Rate', null);
    assert.equal(wb.getDefinedName('Rate', null), null);
  });
});

test('A4: setPrintArea(null) and setPrintArea([]) both clear the print area', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const area = [{ firstRow: 0, firstCol: 0, lastRow: 3, lastCol: 3 }];
    for (const cleared of [null, []]) {
      ws.setPrintArea(area);
      assert.deepEqual(ws.getPrintArea(), area);
      ws.setPrintArea(cleared);
      assert.equal(ws.getPrintArea(), null);
    }
  });
});

test('A5 / N4: a null pivot axis takes the field off the report; the -1 sentinel is not public', async () => {
  const formulon = await getFormulon();
  const { wb, pivot } = pivotWorkbook(formulon);
  try {
    const ws = wb.sheet(0);
    ws.pivot(pivot).setFieldAxis(0, null);
    assert.equal(ws.pivot(pivot).listFields()[0].axis, null);
    ws.pivot(pivot).setFieldAxis(0, formulon.PivotAxis.Row);
    assert.equal(ws.pivot(pivot).listFields()[0].axis, formulon.PivotAxis.Row);
    assert.throws(
      () => ws.pivot(pivot).setFieldAxis(0, -1),
      (e) => e.name === 'FormulonError' && e.code === formulon.StatusCode.InvalidArgument,
    );
    assert.equal(ws.pivot(pivot).listFields()[0].axis, formulon.PivotAxis.Row);
    assert.throws(
      () => wb.getDefinedName('x', -1),
      (e) => e.name === 'FormulonError',
    );
  } finally {
    wb.dispose();
  }
});

test('a patch keeps omitted and undefined keys and clears only a nullable key on null', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    ws.updateFormatDefaults({ defaultColWidthChars: 12, defaultRowHeightPt: 20 });
    ws.updateFormatDefaults({ defaultRowHeightPt: undefined });
    assert.deepEqual(
      [ws.getFormatDefaults().defaultColWidthChars, ws.getFormatDefaults().defaultRowHeightPt],
      [12, 20],
    );
    ws.updateFormatDefaults({ defaultColWidthChars: null });
    assert.deepEqual(
      [ws.getFormatDefaults().defaultColWidthChars, ws.getFormatDefaults().defaultRowHeightPt],
      [null, 20],
    );
    for (const call of [
      () => ws.updateFormatDefaults({ customHeight: null }),
      () => wb.updateIterative({ enabled: null }),
      () => ws.updatePageSetup({ scalePct: null }),
      () => ws.updatePageMargins({ leftIn: null }),
    ]) {
      assert.throws(call, TypeError);
    }
    assert.equal(wb.getIterative().enabled, false);

    // Page setup: an omitted or undefined key keeps; `null` on orientation removes the attribute.
    ws.updatePageSetup({ scalePct: 150 });
    ws.updatePageSetup({ scalePct: undefined, orientation: null });
    assert.deepEqual([ws.getPageSetup().scalePct, ws.getPageSetup().orientation], [150, null]);
    ws.updatePageSetup({ orientation: 1 });
    assert.equal(ws.getPageSetup().orientation, 1);
    ws.updatePageSetup({ orientation: null });
    assert.equal(ws.getPageSetup().orientation, null);
    assert.equal(ws.getPageSetup().scalePct, 150);
    ws.updatePageMargins({ leftIn: 0.5 });
    ws.updatePageMargins({ rightIn: undefined });
    assert.equal(ws.getPageMargins().leftIn, 0.5);
  });
});

test('A7: an empty defined-name formula is refused; an empty comment text writes an empty note', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    expectFormulonError(() => wb.setDefinedName({ name: 'Empty', formula: '' }), StatusCode.InvalidArgument);
    assert.equal(wb.getDefinedName('Empty', null), null);
    const ws = wb.sheet(0);
    ws.setComment(1, 1, { author: 'Ann', text: '' });
    assert.deepEqual({ ...ws.getComment(1, 1) }, { author: 'Ann', text: '' });
  });
});

test('A8: removeComment removes a note and throws NotFound where there is none', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    ws.setComment(1, 1, { author: 'Ann', text: 'x' });
    ws.removeComment(1, 1);
    assert.equal(ws.getComment(1, 1), null);
    expectFormulonError(() => ws.removeComment(1, 1), StatusCode.NotFound);
  });
});

test('A10: null clears workbook and sheet protection and phonetic runs, as [] does for runs', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    wb.setWorkbookProtection({ lockStructure: true });
    wb.setWorkbookProtection(null);
    assert.equal(wb.getWorkbookProtection().lockStructure, false);
    ws.setSheetProtection({});
    assert.notEqual(ws.getSheetProtection(), null);
    ws.setSheetProtection(null);
    assert.equal(ws.getSheetProtection(), null);
    ws.setText(0, 0, '漢字');
    for (const cleared of [null, []]) {
      ws.setCellPhoneticRuns(0, 0, [{ start: 0, end: 2, text: 'かんじ' }]);
      assert.equal(ws.getCellPhoneticRuns(0, 0).length, 1);
      ws.setCellPhoneticRuns(0, 0, cleared);
      assert.deepEqual(ws.getCellPhoneticRuns(0, 0), []);
    }
  });
});

test('A11: an absent chart title or legend is null, and null in a patch removes it', async () => {
  const formulon = await getFormulon();
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    for (let r = 0; r < 4; r += 1) ws.setNumber(r, 1, r + 1);
    const id = ws.addChart({ type: formulon.ChartPlotType.Bar, range: 'Sheet1!$B$1:$B$4', legend: null });
    assert.equal(ws.chart(id).getProperties().legend, null);
    ws.chart(id).setTitle('T');
    assert.equal(ws.chart(id).getProperties().title.text, 'T');
    ws.chart(id).setTitle(null);
    assert.equal(ws.chart(id).getProperties().title, null);
    ws.chart(id).setLegend({ position: formulon.ChartLegendPosition.Top, overlay: false, hiddenEntries: [] });
    assert.equal(ws.chart(id).getProperties().legend.position, formulon.ChartLegendPosition.Top);
    ws.chart(id).setLegend(null);
    assert.equal(ws.chart(id).getProperties().legend, null);
    const axis = ws.chart(id).getProperties().axes[0];
    ws.chart(id).updateAxis(axis.id, { title: 'Axis' });
    assert.equal(ws.chart(id).getProperties().axes[0].title.text, 'Axis');
    ws.chart(id).updateAxis(axis.id, { title: null });
    assert.equal(ws.chart(id).getProperties().axes[0].title, null);
  });
});

// ---- Patches (P1-P4) ------------------------------------------------------

test('P1: an empty patch changes nothing', async () => {
  await withWorkbook((wb) => {
    wb.updateIterative({ enabled: true, maxIterations: 7, maxChange: 0.5 });
    const before = wb.getIterative();
    wb.updateIterative({});
    assert.deepEqual(wb.getIterative(), before);
    const ws = wb.sheet(0);
    const defaults = ws.getFormatDefaults();
    ws.updateFormatDefaults({});
    assert.deepEqual(ws.getFormatDefaults(), defaults);
  });
});

test('P3: a fractional or NaN iteration cap is a TypeError; a non-finite threshold is InvalidArgument', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const before = wb.getIterative();
    for (const maxIterations of [1.5, Number.NaN]) {
      assert.throws(() => wb.updateIterative({ maxIterations }), TypeError, String(maxIterations));
    }
    for (const maxChange of [Number.POSITIVE_INFINITY, Number.NaN]) {
      expectFormulonError(() => wb.updateIterative({ maxChange }), StatusCode.InvalidArgument);
    }
    expectFormulonError(() => wb.updateIterative({ maxChange: -1 }), StatusCode.InvalidArgument);
    assert.deepEqual(wb.getIterative(), before);
  });
});

test('P4: a fractional scope is a RangeError, a negative or past-the-end one is SheetIndexOutOfRange', async () => {
  await withWorkbook((wb, formulon) => {
    assert.throws(() => wb.setDefinedName({ name: 'N', formula: '1', scope: 1.5 }), RangeError);
    expectFormulonError(
      () => wb.setDefinedName({ name: 'N', formula: '1', scope: -1 }),
      formulon.StatusCode.SheetIndexOutOfRange,
    );
    expectFormulonError(
      () => wb.setDefinedName({ name: 'N', formula: '1', scope: 2 ** 31 }),
      formulon.StatusCode.SheetIndexOutOfRange,
    );
    assert.equal(wb.listDefinedNames().length, 0);
  });
});

// ---- Lists (L1-L2) ---------------------------------------------------------

test('L1 / L2: an empty collection lists as [], and each call returns a fresh array', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    for (const list of ['listCfRules', 'listValidations', 'listProtectedRanges', 'listSparklineGroups']) {
      assert.deepEqual(ws[list](), [], list);
    }
    assert.deepEqual(wb.sheet(0).listTables(), []);
    assert.deepEqual(wb.listExternalLinks(), []);
    ws.addValidation({ ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }], type: 1, formula1: '1' });
    const first = ws.listValidations();
    first.push(first[0]);
    first[0] = null;
    const second = ws.listValidations();
    assert.notEqual(second, first);
    assert.equal(second.length, 1);
    assert.notEqual(second[0], null);
    const sheets = wb.listSheets();
    sheets.pop();
    assert.equal(wb.listSheets().length, 1);
  });
});

// ---- Enums (N1-N4) ---------------------------------------------------------

test('N1: enums of one family share their values', async () => {
  const formulon = await getFormulon();
  assert.equal(formulon.SlicerSortOrder.Ascending, 0);
  assert.equal(formulon.PivotSortOrder.Manual, 0);
  assert.notEqual(formulon.SlicerSortOrder.Descending, formulon.PivotSortOrder.Descending);
  assert.equal(formulon.AnchorEditAs.OneCell, formulon.AnchorKind.OneCell);
  assert.equal(formulon.AnchorEditAs.TwoCell, formulon.AnchorKind.TwoCell);
  assert.equal(formulon.AnchorEditAs.Absolute, formulon.AnchorKind.Absolute);
  assert.deepEqual({ ...formulon.BlanksAs }, { Zero: 0, Gap: 1, Span: 2 });
  assert.equal(formulon.ChartAxisPosition.Bottom, 0);
  assert.notEqual(formulon.ChartAxisPosition.Top, formulon.ChartLegendPosition.Top);
  for (const [table, member] of [
    ['BlanksAs', 'Absent'],
    ['ChartAxisCrosses', 'Absent'],
    ['ShapePaint', 'Unset'],
    ['ChartTitleKind', 'None'],
    ['ChartLegendPosition', 'None'],
  ]) {
    assert.equal(formulon[table][member], undefined, `${table}.${member}`);
  }
});

test('an enum argument outside its set is refused by type, range or the engine', async () => {
  await withWorkbook((wb, { StatusCode, SheetViewMode }) => {
    const ws = wb.sheet(0);
    expectFormulonError(() => ws.updateView({ viewMode: 99 }), StatusCode.InvalidArgument);
    expectFormulonError(() => ws.updateView({ viewMode: -1 }), StatusCode.InvalidArgument);
    for (const bad of [1.5, Number.NaN]) {
      assert.throws(() => ws.updateView({ viewMode: bad }), RangeError, String(bad));
    }
    assert.throws(() => ws.updateView({ viewMode: '0' }), TypeError);
    assert.equal(ws.getView().viewMode, SheetViewMode.Normal);
  });
});

test('N3: the old ordinal of a removed member reads with its new meaning', async () => {
  const formulon = await getFormulon();
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    for (let r = 0; r < 4; r += 1) ws.setNumber(r, 1, r + 1);
    const id = ws.addChart({ type: formulon.ChartPlotType.Line, range: 'Sheet1!$B$1:$B$4' });
    ws.chart(id).updateDisplay({ blanksAs: 0 });
    assert.equal(ws.chart(id).getProperties().blanksAs, formulon.BlanksAs.Zero);
  });
});

// ---- Options objects (O1-O3) -----------------------------------------------

test('O1 / O2: an input or a patch refuses unknown keys and undefined is omitted', async () => {
  await withWorkbook((wb) => {
    wb.updateIterative({ enabled: true, maxIterations: 9 });
    assert.deepEqual([wb.getIterative().enabled, wb.getIterative().maxIterations], [true, 9]);
    wb.updateIterative({ maxIterations: undefined, maxChange: undefined });
    assert.equal(wb.getIterative().maxIterations, 9);
    const ws = wb.sheet(0);
    assert.throws(() => ws.setHyperlink(0, 0, { target: 'https://example.com/', unknown: 1 }), TypeError);
    assert.equal(ws.getHyperlink(0, 0), null);
    ws.setHyperlink(0, 0, { target: 'https://example.com/', display: undefined });
    assert.equal(ws.getHyperlink(0, 0).display, null);
    assert.throws(() => wb.updateIterative({ extra: 1 }), TypeError);
  });
});

test('O3: a missing required field is a TypeError and writes nothing', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    assert.throws(() => ws.setComment(0, 0, { author: 'Ann' }), TypeError);
    assert.equal(ws.getComment(0, 0), null);
    assert.throws(() => ws.setHyperlink(0, 0, { target: null }), TypeError);
    assert.throws(() => ws.setHyperlink(0, 0, {}), TypeError);
    assert.equal(ws.getHyperlink(0, 0), null);
    assert.throws(() => ws.addThreadedComment({ id: '{00000000-0000-4000-8000-000000000001}' }), TypeError);
    assert.deepEqual(ws.listThreadedComments(), []);
    assert.throws(() => wb.setDefinedName({ formula: '1' }), TypeError);
  });
});

// ---- Text (T1) ---------------------------------------------------------------

test('T1: non-ASCII text round-trips and a lone surrogate becomes U+FFFD', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setComment(0, 0, { author: '校閲', text: '確認😀' });
    assert.equal(ws.getComment(0, 0).text, '確認😀');
    ws.setComment(1, 0, { author: 'a', text: 'x\uD800y' });
    assert.equal(ws.getComment(1, 0).text, 'x\uFFFDy');
    ws.setHyperlink(2, 0, { location: 'Sheet1!A1', target: null, display: '表示\uDC00' });
    assert.equal(ws.getHyperlink(2, 0).display, '表示\uFFFD');
  });
});

// A one-pivot workbook: Region on rows, Sum of Amount as the data field.
function pivotWorkbook(formulon) {
  const wb = formulon.Workbook.createDefault();
  const ws = wb.sheet(0);
  ws.setText(0, 0, 'Region');
  ws.setText(0, 1, 'Amount');
  const cache = wb.addPivotCache();
  const { cacheId } = cache;
  wb.pivotCache(cacheId).setWorksheetSource({ range: 'A1:B3', sheet: 'Sheet1' });
  wb.pivotCache(cacheId).addField('Region');
  wb.pivotCache(cacheId).addField('Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    ws.setText(wb.pivotCache(cacheId).getRecordCount() + 1, 0, region);
    ws.setNumber(wb.pivotCache(cacheId).getRecordCount() + 1, 1, amount);
    const rec = wb.pivotCache(cacheId).addRecord();
    wb.pivotCache(cacheId).setRecordValue(rec, 0, region);
    wb.pivotCache(cacheId).setRecordValue(rec, 1, amount);
  }
  const pivot = ws.addPivot({ name: 'Pivot1', cacheId, anchor: { row: 0, col: 4 } });
  ws.pivot(pivot).addField({ fieldIndex: 'Region', axis: 0 });
  const amount = ws.pivot(pivot).addField({ fieldIndex: 'Amount', axis: 2 });
  ws.pivot(pivot).addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 });
  return { wb, pivot };
}

// ---- Status refusals and Excel limits ------------------------------------

test('activating a hidden or very hidden sheet is PreconditionFailed and keeps the active sheet', async () => {
  await withWorkbook((wb, { StatusCode, SheetVisibility }) => {
    wb.addSheet('Second');
    for (const visibility of [SheetVisibility.Hidden, SheetVisibility.VeryHidden]) {
      wb.sheet(1).setVisibility(visibility);
      expectFormulonError(() => wb.setActiveSheet(1), StatusCode.PreconditionFailed);
      assert.equal(wb.activeSheet().index, 0);
    }
    wb.sheet(1).setVisibility(SheetVisibility.Visible);
    wb.setActiveSheet(1);
    assert.equal(wb.activeSheet().index, 1);
  });
});

test('a duplicate threaded comment id is AlreadyExists', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const person = '{00000000-0000-4000-8000-0000000000A1}';
    const id = '{00000000-0000-4000-8000-000000000001}';
    wb.addPerson({ id: person, displayName: 'A', userId: 'a', providerId: 'None' });
    const ws = wb.sheet(0);
    ws.addThreadedComment({ id, text: 'first', cell: 'A1', personId: person });
    expectFormulonError(
      () => ws.addThreadedComment({ id, text: 'second', cell: 'B2', personId: person }),
      StatusCode.AlreadyExists,
    );
    assert.equal(ws.listThreadedComments().length, 1);
  });
});

test('an unknown function name is NotFound while an unknown profile stays InvalidArgument', async () => {
  const formulon = await getFormulon();
  const { StatusCode } = formulon;
  expectFormulonError(() => formulon.localizeFunctionName('NOPE', 'mac-365-de_DE'), StatusCode.NotFound);
  expectFormulonError(() => formulon.canonicalizeFunctionName('NOPE', 'mac-365-de_DE'), StatusCode.NotFound);
  expectFormulonError(() => formulon.localizeFunctionName('SUM', 'nope'), StatusCode.InvalidArgument);
  expectFormulonError(() => formulon.canonicalizeFunctionName('SUMME', 'nope'), StatusCode.InvalidArgument);
});

test('a column outline level above 7 is InvalidArgument and level 7 is stored', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    ws.setColumnOutline('B', 7);
    expectFormulonError(() => ws.setColumnOutline('B', 8), StatusCode.InvalidArgument);
    expectFormulonError(() => ws.setColumnOutline(1, 1, 9), StatusCode.InvalidArgument);
  });
});

test('an element position past the end is IndexOutOfRange', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    const xfCount = wb.ooxml.listCellXfs().length;
    ws.ooxml.setRangeXfIndex('A1:B2', xfCount - 1);
    expectFormulonError(() => ws.ooxml.setRangeXfIndex('A1:B2', xfCount), StatusCode.IndexOutOfRange);
    expectFormulonError(() => ws.ooxml.setRangeXfIndex('A1:B2', 9999), StatusCode.IndexOutOfRange);
    expectFormulonError(() => wb.ooxml.getFont(wb.ooxml.listFonts().length), StatusCode.IndexOutOfRange);
  });
});

test('a refusal caused by the workbook or a neighbouring object is PreconditionFailed', async () => {
  await withWorkbook((wb, { StatusCode, SheetVisibility }) => {
    const ws = wb.sheet(0);
    expectFormulonError(() => wb.removeSheet(0), StatusCode.PreconditionFailed);
    expectFormulonError(() => ws.setVisibility(SheetVisibility.Hidden), StatusCode.PreconditionFailed);
    assert.equal(wb.listSheets().length, 1);
    wb.addSheet('Second');
    wb.sheet(1).setVisibility(SheetVisibility.Hidden);
    expectFormulonError(() => wb.removeSheet(0), StatusCode.PreconditionFailed);
    assert.equal(wb.listSheets().length, 2);

    for (let r = 0; r < 4; r += 1) for (let c = 0; c < 3; c += 1) ws.setText(r, c, `v${r}${c}`);
    ws.addTable({ range: 'A1:B3', name: 'Sales' });
    expectFormulonError(() => ws.addTable({ range: 'B2:C4', name: 'Costs' }), StatusCode.PreconditionFailed);
    assert.equal(ws.listTables().length, 1);
  });
});

test('setters accept Excel limits and reject one step beyond with InvalidArgument', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    ws.setText('A1', 'x'.repeat(32767));
    expectFormulonError(() => ws.setText('A2', 'x'.repeat(32768)), StatusCode.InvalidArgument);
    for (const zoomPct of [10, 400]) ws.updateView({ zoomPct });
    for (const zoomPct of [9, 401]) {
      expectFormulonError(() => ws.updateView({ zoomPct }), StatusCode.InvalidArgument);
    }
    for (const sizePt of [1, 409]) ws.updateCellFormat('A1', { font: { sizePt } });
    for (const sizePt of [0.5, 410]) {
      expectFormulonError(() => ws.updateCellFormat('A1', { font: { sizePt } }), StatusCode.InvalidArgument);
    }
    for (const scalePct of [10, 400]) ws.updatePageSetup({ scalePct });
    for (const scalePct of [9, 401]) {
      expectFormulonError(() => ws.updatePageSetup({ scalePct }), StatusCode.InvalidArgument);
    }
  });
});

test('a row height above 409 points is InvalidArgument and 409 is stored', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    ws.setRowHeight('1', 409);
    expectFormulonError(() => ws.setRowHeight('1', 410), StatusCode.InvalidArgument);
    expectFormulonError(() => ws.setRowHeight('1', 409.5), StatusCode.InvalidArgument);
  });
});

test('a column width above 255 characters is InvalidArgument and 255 is stored', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const ws = wb.sheet(0);
    ws.setColumnWidth('B', 255);
    expectFormulonError(() => ws.setColumnWidth('B', 256), StatusCode.InvalidArgument);
    expectFormulonError(() => ws.setColumnWidth('B', 255.5), StatusCode.InvalidArgument);
  });
});
