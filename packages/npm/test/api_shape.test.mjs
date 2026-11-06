// The shape of the public surface: the createFormulon entry, FormulonError
// throws, the Value union, setValue dispatch and the A1 Sheet view.
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  const entryPath = path.resolve(pkgRoot, target);
  const mod = await import(pathToFileURL(entryPath).href);
  const common = await import(pathToFileURL(path.join(path.dirname(entryPath), 'common.js')).href);
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
  'evalFormula',
  'getVersion',
  'getStatusName',
  'getErrorDisplayName',
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

test('named exports are the value tables, FormulonError and createFormulon only', async () => {
  const { mod, common } = await loaded;
  const tables = common.FORMULON_MEMBERS.filter((name) => !NON_TABLE_MEMBERS.has(name));
  const expected = ['createFormulon', 'default', 'FormulonError', ...tables].sort();
  assert.deepEqual(Object.keys(mod).sort(), expected);
  const formulon = await getFormulon();
  for (const name of tables) assert.equal(mod[name], formulon[name], name);
});

test('createFormulon ignores unknown option keys', async () => {
  const { mod, common } = await loaded;
  const formulon = await mod.createFormulon({ foo: 1, noExitRuntime: false });
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
  assert.equal(wb.getSheetCount(), 1);
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
  expectFormulonError(() => wb.getSheetCount(), code);
  expectFormulonError(() => wb.getValue(0, 0, 0), code);
  expectFormulonError(() => wb.setNumber(0, 0, 0, 1), code);
  expectFormulonError(() => view.get('A1'), code);
  expectFormulonError(() => view.set('A1', 1), code);
  expectFormulonError(() => view.getFormula('A1'), code);
  expectFormulonError(() => view.setFormula('A1', '=1'), code);
});

test('out-of-range indices and missing natural keys throw by code', async () => {
  await withWorkbook((wb, formulon) => {
    const { StatusCode } = formulon;
    expectFormulonError(() => wb.getSheetName(1), StatusCode.SheetIndexOutOfRange);
    expectFormulonError(() => wb.getFont(wb.getFontCount()), StatusCode.IndexOutOfRange);
    expectFormulonError(() => wb.removeDefinedName('nope', -1), StatusCode.NotFound);
    assert.equal(wb.getDefinedName('nope', -1), null);
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
  wb.setIterative(true, 100, 0.0001);
  wb.setFormula(0, 0, 0, '=(A1+10)/2');
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
    wb.setFormula(0, 0, 0, '=(A1+10)/2');
    wb.recalc();
    const value = wb.getValue(0, 0, 0);
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
        other.getSheetName(5);
      } catch (e) {
        raised = e;
        throw e;
      }
    });
    assert.throws(
      () => wb.recalc(),
      (caught) => caught === raised && caught.code === formulon.StatusCode.SheetIndexOutOfRange,
    );
    wb.setIterativeProgress(null);
    assert.equal(other.getSheetCount(), 1);
  } finally {
    wb.dispose();
    other.dispose();
  }
});

// ---- Value --------------------------------------------------------------

test('each scalar kind carries only its own fields', async () => {
  await withWorkbook((wb, { ValueKind, ErrorCode }) => {
    wb.setNumber(0, 0, 0, 42);
    wb.setBool(0, 0, 1, true);
    wb.setText(0, 0, 2, 'hi');
    wb.setFormula(0, 0, 3, '=1/0');
    wb.recalc();
    assert.deepStrictEqual({ ...wb.getValue(0, 0, 0) }, { kind: ValueKind.Number, number: 42 });
    assert.deepStrictEqual({ ...wb.getValue(0, 0, 1) }, { kind: ValueKind.Bool, boolean: true });
    assert.deepStrictEqual({ ...wb.getValue(0, 0, 2) }, { kind: ValueKind.Text, text: 'hi' });
    assert.deepStrictEqual({ ...wb.getValue(0, 0, 3) }, { kind: ValueKind.Error, errorCode: ErrorCode.Div0 });
    assert.deepStrictEqual({ ...wb.getValue(0, 0, 4) }, { kind: ValueKind.Blank });
    assert.deepEqual(Object.keys(wb.getValue(0, 0, 3)).sort(), ['errorCode', 'kind']);
  });
});

test('evalFormula returns a bare Value', async () => {
  const { evalFormula, ValueKind } = await getFormulon();
  assert.deepStrictEqual({ ...evalFormula('=SUM(1,2,3)') }, { kind: ValueKind.Number, number: 6 });
});

test('setValue(getValue(...)) round-trips every scalar kind', async () => {
  await withWorkbook((wb) => {
    wb.setNumber(0, 0, 0, -1.5);
    wb.setBool(0, 0, 1, false);
    wb.setText(0, 0, 2, '=not a formula');
    wb.setFormula(0, 0, 3, '=NA()');
    wb.recalc();
    for (let col = 0; col <= 4; col += 1) {
      const before = wb.getValue(0, 0, col);
      wb.setValue(0, 1, col, before);
      assert.deepStrictEqual({ ...wb.getValue(0, 1, col) }, { ...before }, `column ${col}`);
    }
  });
});

// ---- setValue -----------------------------------------------------------

test('setValue dispatches each primitive to the matching setter', async () => {
  await withWorkbook((wb, { ValueKind, ErrorCode }) => {
    const rows = [
      [0, (r) => wb.setNumber(0, r, 0, 0)],
      [-0, (r) => wb.setNumber(0, r, 0, -0)],
      [2 ** 53, (r) => wb.setNumber(0, r, 0, 2 ** 53)],
      [2 ** 63, (r) => wb.setNumber(0, r, 0, 2 ** 63)],
      ['=A1', (r) => wb.setText(0, r, 0, '=A1')],
      ['', (r) => wb.setText(0, r, 0, '')],
      [true, (r) => wb.setBool(0, r, 0, true)],
      [false, (r) => wb.setBool(0, r, 0, false)],
      [null, (r) => wb.setBlank(0, r, 0)],
      [{ kind: ValueKind.Blank }, (r) => wb.setBlank(0, r, 0)],
      [{ kind: ValueKind.Number, number: 3 }, (r) => wb.setNumber(0, r, 0, 3)],
      [{ kind: ValueKind.Bool, boolean: true }, (r) => wb.setBool(0, r, 0, true)],
      [{ kind: ValueKind.Text, text: 't' }, (r) => wb.setText(0, r, 0, 't')],
      [{ kind: ValueKind.Error, errorCode: ErrorCode.Ref }, (r) => wb.setError(0, r, 0, ErrorCode.Ref)],
    ];
    rows.forEach(([value, direct], index) => {
      const row = index * 2;
      wb.setNumber(0, row, 1, 99);
      wb.setValue(0, row, 1, value);
      direct(row + 1);
      assert.deepStrictEqual({ ...wb.getValue(0, row, 1) }, { ...wb.getValue(0, row + 1, 0) }, `row ${index}`);
      assert.equal(wb.getFormula(0, row, 1), null, `row ${index} stored a formula`);
    });
  });
});

test('setValue rejects non-finite numbers as the engine does', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    for (const value of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]) {
      expectFormulonError(() => wb.setValue(0, 0, 0, value), StatusCode.InvalidArgument);
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
      assert.throws(() => wb.setValue(0, 0, 0, value), TypeError, String(value?.kind ?? typeof value));
    }
    assert.throws(() => wb.setValue(0, 0, 0), TypeError);
  });
});

// ---- Sheet view ---------------------------------------------------------

test('getSheetIndex matches case-insensitively and returns null on a miss', async () => {
  await withWorkbook((wb) => {
    assert.equal(wb.getSheetIndex('Sheet1'), 0);
    assert.equal(wb.getSheetIndex('SHEET1'), 0);
    assert.equal(wb.getSheetIndex('nope'), null);
    assert.equal(wb.getSheetIndex(''), null);
    assert.throws(() => wb.getSheetIndex(0), TypeError);
  });
});

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
    assert.throws(() => wb.sheet(-1), RangeError);
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
    wb.renameSheet(0, 'Renamed');
    assert.equal(view.name, 'Renamed');
    assert.equal(view.index, 0);
  });
});

test('Sheet get / set / getFormula / setFormula address one cell', async () => {
  await withWorkbook((wb, { ValueKind }) => {
    const view = wb.sheet(0);
    view.set('B3', 5);
    assert.deepStrictEqual({ ...wb.getValue(0, 2, 1) }, { kind: ValueKind.Number, number: 5 });
    assert.deepStrictEqual({ ...view.get('B3') }, { kind: ValueKind.Number, number: 5 });
    view.set('C1', '=B3');
    assert.equal(view.getFormula('C1'), null);
    view.setFormula('D1', '=B3*2');
    assert.equal(wb.getFormula(0, 0, 3), '=B3*2');
    assert.equal(view.getFormula('D1'), '=B3*2');
    wb.recalc();
    assert.deepStrictEqual({ ...view.get('D1') }, { kind: ValueKind.Number, number: 10 });
    view.set('B3', null);
    assert.deepStrictEqual({ ...view.get('B3') }, { kind: ValueKind.Blank });
  });
});

test('Sheet addresses equivalent spellings of one cell alike', async () => {
  await withWorkbook((wb, { ValueKind }) => {
    const view = wb.sheet(0);
    wb.setNumber(0, 2, 1, 7);
    for (const a1 of ['B3', 'b3', '$B$3', '$b3', 'B$3']) {
      assert.deepStrictEqual({ ...view.get(a1) }, { kind: ValueKind.Number, number: 7 }, a1);
    }
    wb.setNumber(0, 0, 0, 1);
    for (const a1 of ['A01', 'a0000001']) {
      assert.deepStrictEqual({ ...view.get(a1) }, { kind: ValueKind.Number, number: 1 }, a1);
    }
    wb.setNumber(0, 1048575, 16383, 3);
    assert.deepStrictEqual({ ...view.get('XFD1048576') }, { kind: ValueKind.Number, number: 3 });
  });
});

test('Sheet rejects addresses that are not one unqualified cell', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const view = wb.sheet(0);
    const rejected = ['', '   ', 'A1:B2', 'Sheet2!A1', 'Sheet1!A1', 'R1C1', 'XFE1', 'A0', 'A1048577', 'B:B', '3:3'];
    rejected.push('Ａ１', 'Ａ1', '😀', 'A\uD800', ' A1', 'A1 ', 'A1\n');
    for (const a1 of rejected) {
      expectFormulonError(() => view.get(a1), StatusCode.InvalidArgument);
      expectFormulonError(() => view.set(a1, 1), StatusCode.InvalidArgument);
    }
    assert.throws(() => view.get(1), TypeError);
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
        expectFormulonError(() => view.get(a1), StatusCode.InvalidArgument);
        expectFormulonError(() => view.set(a1, 1), StatusCode.InvalidArgument);
        return;
      }
      const [row, col] = cell;
      view.set(a1, 7);
      assert.deepStrictEqual({ ...wb.getValue(0, row, col) }, { kind: ValueKind.Number, number: 7 });
      assert.deepStrictEqual({ ...view.get(a1) }, { kind: ValueKind.Number, number: 7 });
    });
  });
}
