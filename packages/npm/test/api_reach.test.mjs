// Conformance of the high-level style API, toNative, whole-row / whole-column span
// strings and the lower-object handles, driven by the shared case tables under
// tools/api/. FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const pkgRoot = path.resolve(here, '..');
const repoRoot = path.resolve(pkgRoot, '..', '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

let formulonPromise;
function getFormulon() {
  formulonPromise ??= loadPackage().then((mod) => mod.createFormulon());
  return formulonPromise;
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

function expectCode(fn, code, label) {
  assert.throws(
    fn,
    (e) => {
      assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
      assert.equal(e.code, code, `code ${e.code} (${e.message})`);
      return true;
    },
    label,
  );
}

const plain = (v) => JSON.parse(JSON.stringify(v));
const readJson = (relative) => JSON.parse(readFileSync(path.join(repoRoot, relative), 'utf8'));

// ---- setCellFormat / getCellFormat: tools/api/cell_format_cases.json ---------------

const FORMAT_CASES = readJson('tools/api/cell_format_cases.json').cases;

// The enum table a case names by member, by the key that carries it.
const ENUM_KEYS = {
  underline: 'UnderlineStyle',
  vertAlign: 'FontVerticalAlign',
  pattern: 'FillPattern',
  style: 'BorderStyle',
  horizontal: 'HorizontalAlign',
  vertical: 'VerticalAlign',
  readingOrder: 'ReadingOrder',
  horizontalAlign: 'HorizontalAlign',
  verticalAlign: 'VerticalAlign',
};

// Replaces enum member names in a patch by their values; colours and other strings stay.
function resolveEnums(value, formulon, key) {
  if (Array.isArray(value)) return value.map((v) => resolveEnums(v, formulon, key));
  if (value !== null && typeof value === 'object') {
    return Object.fromEntries(Object.entries(value).map(([k, v]) => [k, resolveEnums(v, formulon, k)]));
  }
  if (typeof value === 'string' && key in ENUM_KEYS) {
    const table = formulon[ENUM_KEYS[key]];
    assert.ok(value in table, `${ENUM_KEYS[key]} has no member ${value}`);
    return table[value];
  }
  return value;
}

const argbText = (rgb) => `#${(rgb >>> 0).toString(16).toUpperCase().padStart(8, '0')}`;

// Checks that `actual` holds every key `expected` names; `#AARRGGBB` text is compared with `.rgb`.
function assertSubset(actual, expected, label) {
  if (expected !== null && typeof expected === 'object') {
    assert.ok(actual !== null && typeof actual === 'object', `${label}: expected an object, got ${actual}`);
    for (const [k, v] of Object.entries(expected)) assertSubset(actual[k], v, `${label}.${k}`);
  } else if (typeof expected === 'string' && /^#[0-9A-F]{8}$/.test(expected)) {
    assert.equal(argbText(actual.rgb), expected, label);
  } else {
    assert.equal(actual, expected, label);
  }
}

function assertThrowsLike(fn, throws, formulon, label) {
  const expected = throws.js;
  if (expected.startsWith('FormulonError:')) {
    expectCode(fn, formulon.StatusCode[expected.slice('FormulonError:'.length)], label);
  } else {
    assert.throws(fn, globalThis[expected], label);
  }
}

test('cell format cases: the table is non-empty and covers successes and failures', () => {
  assert.ok(FORMAT_CASES.length > 0, 'no cases in cell_format_cases.json');
  assert.ok(FORMAT_CASES.filter((c) => !c.throws).length >= 20);
  assert.ok(FORMAT_CASES.filter((c) => c.throws).length >= 8);
  assert.equal(new Set(FORMAT_CASES.map((c) => c.id)).size, FORMAT_CASES.length, 'duplicate case ids');
});

for (const c of FORMAT_CASES) {
  test(`cell format case: ${c.id}`, async () => {
    await withWorkbook((wb, formulon) => {
      const ws = wb.sheet(0);
      for (const step of c.before ?? []) ws.setCellFormat(step.range, resolveEnums(step.patch, formulon));
      const patch = 'patch_from' in c ? ws.getCellFormat(c.patch_from) : resolveEnums(c.patch, formulon);
      const cells = new Set([...(c.unchanged ?? []), ...Object.keys(c.expect ?? {}), 'A1']);
      const snapshot = () => new Map([...cells].map((cell) => [cell, plain(ws.getCellFormat(cell))]));
      const before = snapshot();
      const xfCount = ws.workbook.listCellXfs().length;

      if (c.throws) {
        assertThrowsLike(() => ws.setCellFormat(c.range, patch), c.throws, formulon, c.id);
        // The failing call changed nothing.
        assert.deepEqual(snapshot(), before, 'a failed call changed a cell');
        assert.equal(wb.listCellXfs().length, xfCount, 'a failed call added a cell format');
        return;
      }

      ws.setCellFormat(c.range, patch);
      const after = snapshot();
      for (const [cell, fmt] of Object.entries(c.expect ?? {})) {
        assertSubset(resolveEnums(ws.getCellFormat(cell), formulon), resolveEnums(fmt, formulon), `${c.id} ${cell}`);
      }
      for (const [cell, fields] of Object.entries(c.expect_xf ?? {})) {
        const xf = wb.getCellXf(ws.getEffectiveStyle(cell).xfIndex);
        for (const [field, want] of Object.entries(resolveEnums(fields, formulon))) {
          assert.equal(xf[field], want, `${c.id} ${cell} xf.${field}`);
        }
      }
      for (const cell of c.unchanged ?? []) assert.deepEqual(after.get(cell), before.get(cell), `${cell} changed`);
      if (c.idempotent) {
        const count = wb.listCellXfs().length;
        ws.setCellFormat(c.range, patch);
        assert.equal(wb.listCellXfs().length, count, 'a second application added a cell format');
      }
    });
  });
}

// ---- toNative: tools/api/to_native_cases.json ---------------------------------------

const NATIVE_CASES = readJson('tools/api/to_native_cases.json').cases;

test('toNative cases: the table is non-empty', () => {
  assert.ok(NATIVE_CASES.length >= 10, `only ${NATIVE_CASES.length} cases in to_native_cases.json`);
});

for (const [i, c] of NATIVE_CASES.entries()) {
  test(`toNative case ${i}: ${c.formula ?? `cell ${c.cell}`}`, async () => {
    const formulon = await getFormulon();
    let value;
    if ('formula' in c) {
      value = formulon.evaluateFormula(c.formula);
    } else {
      const wb = formulon.Workbook.createDefault();
      try {
        value = wb.sheet(0).getValue(c.cell);
      } finally {
        wb.dispose();
      }
    }
    const native = formulon.toNative(value);
    if (c.native !== null && typeof c.native === 'object' && 'error' in c.native) {
      assert.equal(native, value, 'an error maps to the ErrorValue itself');
      assert.equal(native.kind, formulon.ValueKind.Error);
      assert.equal(native.errorCode, formulon.ErrorCode[c.native.error]);
    } else if (c.native !== null && typeof c.native === 'object' && c.native.passthrough) {
      assert.equal(native, value, 'the input value itself comes back');
    } else {
      assert.strictEqual(native, c.native);
    }
  });
}

test('toNative: edge rows', async () => {
  const { toNative, ValueKind, ErrorCode } = await getFormulon();
  // Blank, Number, Bool, Text map to their JS scalars; -0 and an astral string pass through unchanged.
  assert.equal(toNative({ kind: ValueKind.Blank }), null);
  assert.ok(Object.is(toNative({ kind: ValueKind.Number, number: -0 }), -0));
  assert.ok(Object.is(toNative({ kind: ValueKind.Number, number: 0 }), 0));
  assert.equal(toNative({ kind: ValueKind.Bool, boolean: false }), false);
  assert.equal(toNative({ kind: ValueKind.Text, text: '' }), '');
  assert.equal(toNative({ kind: ValueKind.Text, text: '\u{1F600}x\u{FFFF}' }), '\u{1F600}x\u{FFFF}');
  // Array, Ref and Lambda come back as the same object.
  for (const kind of [ValueKind.Array, ValueKind.Ref, ValueKind.Lambda]) {
    const v = { kind };
    assert.equal(toNative(v), v, `kind ${kind}`);
  }
  const error = { kind: ValueKind.Error, errorCode: ErrorCode.NA };
  assert.equal(toNative(error), error);
  // Anything that is not a Value is a TypeError.
  for (const bad of [null, undefined, 1, 'x', true, [], {}, { kind: 99 }, { kind: '1' }, { kind: -1 }, { kind: 1.5 }]) {
    assert.throws(() => toNative(bad), TypeError, JSON.stringify(bad));
  }
  assert.throws(() => toNative(), TypeError);
});

// ---- Spans: SPAN_ARGS of tools/api/address_args.json ----------------------------------

const SPAN_ARGS = readJson('tools/api/address_args.json').js.span_args;

// Per `Class.method`: the arguments after the span, a setup run on both workbooks, and whether the
// call has to change what `readState` sees.
const SPAN_METHODS = {
  'Sheet.setColumnWidth': { rest: () => [12], mutates: true },
  'Sheet.setColumnHidden': { rest: () => [true], mutates: true },
  'Sheet.setColumnOutline': { rest: () => [2], mutates: true },
  'Sheet.setColumnCollapsed': { rest: () => [true], mutates: true },
  'Sheet.setColumnStyle': { rest: (ws) => [xfOf(ws)], mutates: true },
  'Sheet.clearColumnStyle': { rest: () => [], pre: (ws) => ws.setColumnStyle(1, 3, xfOf(ws)), mutates: true },
  'Sheet.clearColumnWidth': { rest: () => [], pre: (ws) => ws.setColumnWidth(1, 3, 12), mutates: true },
  'Sheet.setRowHeight': { rest: () => [30], mutates: true },
  'Sheet.setRowAutoHeight': { rest: () => [30], mutates: true },
  'Sheet.setRowHidden': { rest: () => [true], mutates: true },
  'Sheet.setRowOutline': { rest: () => [2], mutates: true },
  'Sheet.setRowCollapsed': { rest: () => [true], mutates: true },
  'Sheet.setRowStyle': { rest: (ws) => [xfOf(ws)], mutates: true },
  'Sheet.clearRowStyle': { rest: () => [], pre: (ws) => ws.setRowStyle(2, 4, xfOf(ws)), mutates: true },
  'Sheet.clearRowHeight': { rest: () => [], pre: (ws) => ws.setRowHeight(2, 4, 30), mutates: true },
  'Sheet.insertRows': { rest: () => [], mutates: true },
  'Sheet.deleteRows': { rest: () => [], mutates: true },
  'Sheet.insertColumns': { rest: () => [], mutates: true },
  'Sheet.deleteColumns': { rest: () => [], mutates: true },
};

function xfOf(ws) {
  return ws.workbook.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 2 });
}

// Column operations take B:D (1..3), row operations 3:5 (2..4).
const SPAN_FORMS = {
  col: { text: 'B:D', first: 1, last: 3 },
  row: { text: '3:5', first: 2, last: 4 },
};

function readState(ws) {
  const values = [];
  for (let r = 0; r < 8; r++) for (let c = 0; c < 6; c++) values.push(plain(ws.getValue(r, c)));
  return plain({ cols: ws.listColumnOverrides(), rows: ws.listRowOverrides(), values });
}

function seedGrid(ws) {
  for (let r = 0; r < 8; r++) for (let c = 0; c < 6; c++) ws.setNumber(r, c, r * 10 + c);
}

test('span cases: SPAN_ARGS is non-empty and every entry has an argument row', () => {
  const names = Object.keys(SPAN_ARGS);
  assert.ok(names.length > 0, 'no span_args in address_args.json');
  assert.deepEqual(
    names.filter((n) => !(n in SPAN_METHODS)),
    [],
    'span methods missing from the argument table',
  );
  assert.deepEqual(
    Object.keys(SPAN_METHODS).filter((n) => !(n in SPAN_ARGS)),
    [],
    'rows for methods that take no span',
  );
});

for (const [key, positions] of Object.entries(SPAN_ARGS)) {
  const row = SPAN_METHODS[key];
  if (!row) continue;
  test(`span case: ${key} takes the span string like its numeric form`, async () => {
    assert.equal(positions.length, 1);
    const [position, axis, form] = positions[0];
    assert.equal(position, 0);
    const span = SPAN_FORMS[axis];
    const method = key.split('.')[1];
    const run = (asString) =>
      withWorkbook((wb) => {
        const ws = wb.sheet(0);
        seedGrid(ws);
        row.pre?.(ws);
        const before = readState(ws);
        const rest = row.rest(ws);
        const args = asString
          ? [span.text, ...rest]
          : form === 'start_count'
            ? [span.first, span.last - span.first + 1, ...rest]
            : [span.first, span.last, ...rest];
        ws[method](...args);
        return { before, after: readState(ws) };
      });
    const numeric = await run(false);
    const text = await run(true);
    assert.deepEqual(text.after, numeric.after, `${key}: the string form differs from the numeric form`);
    if (row.mutates) assert.notDeepEqual(text.after, text.before, `${key} changed nothing`);
  });
}

test('span edge rows: single, bounds, whole sheet, dollar signs and rejected spellings', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const InvalidArgument = f.StatusCode.InvalidArgument;
    // "B" and "3" are one column and one row; the grid's last column and row are valid.
    ws.setColumnWidth('B', 9);
    assert.deepEqual(
      plain(ws.listColumnOverrides()).map((c) => [c.first, c.last]),
      [[1, 1]],
    );
    ws.setRowHeight('3', 25);
    assert.deepEqual(
      plain(ws.listRowOverrides()).map((r) => r.row),
      [2],
    );
    ws.setColumnWidth('XFD', 9);
    ws.setRowHeight('1048576', 25);
    assert.ok(ws.listColumnOverrides().some((c) => c.last === 16383));
    assert.ok(ws.listRowOverrides().some((r) => r.row === 1048575));
    // $ is allowed.
    ws.setColumnHidden('$B:$D', true);
    ws.setRowHidden('$3:$5', true);
    // The whole sheet is both a column span and a row span.
    ws.setColumnWidth('A:XFD', 10);
    ws.setRowHeight('1:1048576', 20);
    // Rejected spellings.
    for (const bad of ['0', '', 'XFE', '1048577', 'Ｂ:Ｄ', 'Sheet1!B:D', 'B:', ':D', 'B:3']) {
      expectCode(() => ws.setColumnWidth(bad, 9), InvalidArgument, `column ${JSON.stringify(bad)}`);
    }
    for (const bad of ['0', '', '1048577', 'B', 'Ｂ:Ｄ', 'Sheet1!3:5', '3:']) {
      expectCode(() => ws.setRowHeight(bad, 9), InvalidArgument, `row ${JSON.stringify(bad)}`);
    }
    // A column span on a row operation and the reverse.
    expectCode(() => ws.setRowHeight('B:D', 9), InvalidArgument);
    expectCode(() => ws.setColumnWidth('3:5', 9), InvalidArgument);
    expectCode(() => ws.insertRows('B:D'), InvalidArgument);
    expectCode(() => ws.deleteColumns('3:5'), InvalidArgument);
    // Types: null and undefined are not spans; a number is the numeric form.
    assert.throws(() => ws.setColumnWidth(null, 9), TypeError);
    assert.throws(() => ws.setColumnWidth(undefined, 9), TypeError);
    assert.throws(() => ws.insertRows(null), TypeError);
    assert.throws(() => ws.insertRows(), TypeError);
    assert.throws(() => ws.setColumnWidth(1, 3), TypeError);
    // A rejected call changed nothing.
    const state = readState(ws);
    expectCode(() => ws.deleteRows('0'), InvalidArgument);
    assert.deepEqual(readState(ws), state);
  });
});

test('span edge rows: insert and delete take the span as start and count', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    seedGrid(ws);
    ws.insertRows('3:5');
    // Rows 3..5 (1-based) are three new blank rows; the old row 3 moved to row 6.
    assert.equal(ws.getValue(2, 0).kind, 0);
    assert.equal(ws.getValue(4, 0).kind, 0);
    assert.equal(ws.getValue(5, 0).number, 20);
    ws.deleteRows('3:5');
    assert.equal(ws.getValue(2, 0).number, 20);
    ws.insertColumns('B');
    assert.equal(ws.getValue(0, 1).kind, 0);
    assert.equal(ws.getValue(0, 2).number, 1);
    ws.deleteColumns('B:B');
    assert.equal(ws.getValue(0, 1).number, 1);
  });
});
