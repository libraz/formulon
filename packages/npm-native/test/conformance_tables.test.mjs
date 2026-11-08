// Runs the shared conformance tables of tools/api against the staged native package:
// cell_format_cases.json (setCellFormat / getCellFormat) and to_native_cases.json (toNative).
// An empty table fails.

import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { assert, getModule, isFormulonError, pkgRoot } from './smoke_support.mjs';

const repoRoot = path.resolve(pkgRoot, '..', '..');
const readTable = (name) => JSON.parse(readFileSync(path.join(repoRoot, 'tools', 'api', name), 'utf8'));

const FORMAT_TABLE = readTable('cell_format_cases.json');
const NATIVE_TABLE = readTable('to_native_cases.json');

// Enum-valued keys of a patch or a read record: the dotted key path to the enum table that names its members.
const ENUM_KEYS = new Map([
  ['font.underline', 'UnderlineStyle'],
  ['font.vertAlign', 'FontVerticalAlign'],
  ['font.family', 'FontFamily'],
  ['font.scheme', 'FontScheme'],
  ['fill.pattern', 'FillPattern'],
  ['fill.gradientType', 'GradientType'],
  ['border.left.style', 'BorderStyle'],
  ['border.right.style', 'BorderStyle'],
  ['border.top.style', 'BorderStyle'],
  ['border.bottom.style', 'BorderStyle'],
  ['border.diagonal.style', 'BorderStyle'],
  ['alignment.horizontal', 'HorizontalAlign'],
  ['alignment.vertical', 'VerticalAlign'],
  ['alignment.readingOrder', 'ReadingOrder'],
]);
// The same keys on a CellXf read through getCellXf.
const XF_ENUM_KEYS = new Map([
  ['horizontalAlign', 'HorizontalAlign'],
  ['verticalAlign', 'VerticalAlign'],
]);
const COLOR_KEYS = new Set(['color', 'fg', 'bg']);

const isPlainObject = (v) => v !== null && typeof v === 'object' && !Array.isArray(v);

/** Replaces each enum member name in a patch with its value; a number passes as is. */
function toInput(mod, node, trail = []) {
  if (!isPlainObject(node)) return node;
  const out = {};
  for (const [key, value] of Object.entries(node)) {
    const dotted = [...trail, key].join('.');
    const table = ENUM_KEYS.get(dotted);
    if (table && typeof value === 'string') {
      assert.ok(value in mod[table], `${table}.${value}`);
      out[key] = mod[table][value];
    } else {
      out[key] = toInput(mod, value, [...trail, key]);
    }
  }
  return out;
}

/** The AARRGGBB number of an expected `#AARRGGBB` colour. */
const argb = (text) => Number.parseInt(text.slice(1), 16) >>> 0;

/** Asserts that every key named in `expected` reads back from `actual`; a nested object compares its own keys. */
function assertReads(mod, actual, expected, label, trail = [], table = ENUM_KEYS) {
  for (const [key, want] of Object.entries(expected)) {
    const where = `${label}.${key}`;
    const dotted = [...trail, key].join('.');
    const got = actual[key];
    if (COLOR_KEYS.has(key) && typeof want === 'string') {
      assert.equal(got.rgb >>> 0, argb(want), where);
    } else if (table.has(dotted) && typeof want === 'string') {
      assert.equal(got, mod[table.get(dotted)][want], where);
    } else if (isPlainObject(want)) {
      assert.ok(isPlainObject(got), `${where} is an object`);
      assertReads(mod, got, want, where, [...trail, key], table);
    } else {
      assert.deepEqual(got, want, where);
    }
  }
}

function buildSheet(mod) {
  const wb = mod.Workbook.createDefault();
  return { wb, ws: wb.sheet(0) };
}

function expectedThrow(mod, spec) {
  const kind = spec.js;
  if (kind === 'TypeError') return TypeError;
  if (kind === 'RangeError') return RangeError;
  const [, code] = kind.split(':');
  assert.equal(kind.startsWith('FormulonError:'), true, kind);
  return (e) => isFormulonError(e, mod.StatusCode[code]);
}

test('the cell format table is not empty', () => {
  assert.ok(FORMAT_TABLE.cases.length > 0);
  const ids = FORMAT_TABLE.cases.map((c) => c.id);
  assert.equal(new Set(ids).size, ids.length, 'case ids are unique');
});

for (const spec of FORMAT_TABLE.cases) {
  test(`cell format case: ${spec.id}`, async () => {
    const mod = await getModule();
    const { wb, ws } = buildSheet(mod);
    try {
      for (const step of spec.before ?? []) ws.setCellFormat(step.range, toInput(mod, step.patch));

      const patch = spec.patch_from ? ws.getCellFormat(spec.patch_from) : toInput(mod, spec.patch);
      const watched = spec.unchanged ?? [];
      const snapshot = () => watched.map((cell) => ws.getCellFormat(cell));
      const beforeCells = snapshot();
      const xfCount = wb.ooxml.listCellXfs().length;

      if (spec.throws) {
        assert.throws(() => ws.setCellFormat(spec.range, patch), expectedThrow(mod, spec.throws));
        assert.equal(wb.ooxml.listCellXfs().length, xfCount, 'a refused patch adds no XF');
        assert.deepEqual(snapshot(), beforeCells, 'a refused patch changes nothing');
        return;
      }

      ws.setCellFormat(spec.range, patch);

      for (const [cell, expected] of Object.entries(spec.expect ?? {})) {
        assertReads(mod, ws.getCellFormat(cell), expected, `${spec.id} ${cell}`);
      }
      for (const [cell, expected] of Object.entries(spec.expect_xf ?? {})) {
        const xf = wb.ooxml.getCellXf(ws.getEffectiveStyle(cell).xfIndex);
        assertReads(mod, xf, expected, `${spec.id} ${cell} xf`, [], XF_ENUM_KEYS);
      }
      assert.deepEqual(snapshot(), beforeCells, 'cells outside the range keep their format');

      if (spec.idempotent) {
        const count = wb.ooxml.listCellXfs().length;
        ws.setCellFormat(spec.range, patch);
        assert.equal(wb.ooxml.listCellXfs().length, count, 'applying the patch twice adds no XF');
      }
    } finally {
      wb.dispose();
    }
  });
}

test('the toNative table is not empty', () => {
  assert.ok(NATIVE_TABLE.cases.length > 0);
});

for (const spec of NATIVE_TABLE.cases) {
  const label = spec.formula ?? `cell ${spec.cell}`;
  test(`toNative case: ${label}`, async () => {
    const mod = await getModule();
    const { wb, ws } = buildSheet(mod);
    try {
      const value = spec.formula ? mod.evaluateTypedFormula(spec.formula) : ws.getTypedValue(spec.cell);
      const native = mod.toNative(value);
      const want = spec.native;
      if (want !== null && typeof want === 'object' && 'error' in want) {
        assert.ok(native instanceof mod.ExcelError, 'an error is returned as an ExcelError');
        assert.equal(native.code, mod.ErrorCode[want.error]);
      } else if (want !== null && typeof want === 'object' && want.passthrough) {
        assert.equal(native, value, 'an array, reference or lambda is returned as is');
      } else {
        assert.strictEqual(native, want);
      }
    } finally {
      wb.dispose();
    }
  });
}

test('toNative keeps -0 and rejects what is not a Value', async () => {
  const mod = await getModule();
  assert.ok(Object.is(mod.toNative({ kind: mod.ValueKind.Number, number: -0 }), -0));
  assert.strictEqual(mod.toNative({ kind: mod.ValueKind.Blank }), null);
  assert.strictEqual(mod.toNative({ kind: mod.ValueKind.Text, text: '' }), '');
  assert.strictEqual(mod.toNative({ kind: mod.ValueKind.Bool, boolean: true }), true);
  for (const bad of [null, undefined, 3, 'x', [], {}, { kind: 99 }, { kind: 'Number', number: 1 }]) {
    assert.throws(() => mod.toNative(bad), TypeError, String(JSON.stringify(bad)));
  }
});
