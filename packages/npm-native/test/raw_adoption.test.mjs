// Contract of the raw native addon (no JS facade): record-carrying calls,
// formula syntax errors with their position, colour specs and tab colours.

import assert from 'node:assert/strict';
import { existsSync, readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(testDir, '..', '..', '..');
const addonPath =
  process.env.FORMULON_NODE_ADDON ??
  [
    path.resolve(repoRoot, 'build', 'bin', 'formulon.node'),
    path.resolve(testDir, '..', 'dist', 'prebuilds', `${process.platform}-${process.arch}`, 'formulon.node'),
    path.resolve(testDir, '..', 'dist', 'formulon.node'),
  ].find((candidate) => existsSync(candidate));
assert.ok(addonPath, 'native addon not found; set FORMULON_NODE_ADDON or run `make node-native`');
const addon = createRequire(import.meta.url)(addonPath);

/** Records exactly the arguments the addon constructs its error with. */
class RawError extends Error {
  constructor(code, message, context, parseError) {
    super(message);
    this.code = code;
    this.context = context;
    this.parseError = parseError;
  }
}
addon.__setErrorClass(RawError);

const kColorNone = 0;
const kColorRgb = 1;
const kColorTheme = 2;

/** `FM_PARSE_ERROR_<UPPER_SNAKE>` values keyed by their PascalCase name. */
function parseErrorCodes() {
  const header = readFileSync(path.join(repoRoot, 'src', 'c_api', 'formulon_c.h'), 'utf8');
  const codes = new Map();
  for (const m of header.matchAll(/^\s*FM_PARSE_ERROR_(\w+) = (\d+),?/gm)) {
    const pascal = m[1]
      .toLowerCase()
      .split('_')
      .map((w) => w[0].toUpperCase() + w.slice(1))
      .join('');
    codes.set(pascal, Number(m[2]));
  }
  return codes;
}

/** Measured syntax-error rows with a pinned position for the cell entries. */
function positionedCases() {
  const table = JSON.parse(readFileSync(path.join(repoRoot, 'tools', 'api', 'parse_error_cases.json'), 'utf8'));
  return table.cases.filter(
    (c) => (c.entry === 'set_formula' || c.entry === 'evaluate_formula') && c.start !== null && c.formula !== undefined,
  );
}

function withWorkbook(body) {
  const wb = addon.Workbook.createDefault();
  try {
    body(wb);
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

test('record-carrying calls round-trip through struct_size-bearing records', () => {
  withWorkbook((wb) => {
    wb.setIterative({ enabled: true, maxIterations: 37, maxChange: 0.25 });
    assert.deepEqual(wb.getIterative(), { enabled: true, maxIterations: 37, maxChange: 0.25 });

    const person = {
      id: '{0A1B2C3D-0000-4000-8000-000000000001}',
      displayName: 'Alice',
      userId: 'alice@example.com',
      providerId: 'AD',
    };
    wb.addPerson(person);
    assert.deepEqual(wb.listPeople(), [person]);
  });
});

test('a formula syntax error carries its code and UTF-16 position', () => {
  const codes = parseErrorCodes();
  const cases = positionedCases();
  assert.ok(cases.length > 0);
  withWorkbook((wb) => {
    for (const c of cases) {
      const error = thrown(() =>
        c.entry === 'set_formula' ? wb.setFormula(0, 0, 0, c.formula) : wb.evaluateFormula(0, 0, 0, c.formula),
      );
      assert.ok(error instanceof RawError, `${c.id}: ${error}`);
      assert.equal(error.code, c.status, c.id);
      const { code, ...position } = error.parseError ?? {};
      assert.ok(c.codes.map((name) => codes.get(name)).includes(code), `${c.id}: code ${code}`);
      assert.deepEqual(position, { start: c.start, end: c.end, line: c.line, column: c.column }, c.id);
    }
  });
});

test('a failure that is not a syntax error carries no parseError', () => {
  withWorkbook((wb) => {
    const error = thrown(() => wb.getValue(99, 0, 0));
    assert.ok(error instanceof RawError);
    assert.equal(error.parseError, undefined);
  });
});

test('a CF rule formula syntax error names its field', () => {
  withWorkbook((wb) => {
    const error = thrown(() =>
      wb.addCfRule(0, {
        type: 0,
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }],
        formula1: '=1+\u30002',
      }),
    );
    assert.ok(error instanceof RawError);
    assert.ok(error.parseError);
    assert.match(error.context, /field=formula1/);
  });
});

test('the module evaluateFormula returns the whole result', () => {
  assert.deepEqual(addon.evaluateFormula('=SUM(1,2,3)'), { rows: 1, cols: 1, cells: [[{ kind: 1, number: 6 }]] });
  const grid = addon.evaluateFormula('={1,2;3,4}');
  assert.equal(grid.rows, 2);
  assert.equal(grid.cols, 2);
  assert.deepEqual(
    grid.cells.map((row) => row.map((v) => v.number)),
    [
      [1, 2],
      [3, 4],
    ],
  );
});

test('CF colours are colour specs', () => {
  withWorkbook((wb) => {
    const colors = [
      { kind: kColorRgb, rgb: 0xff112233, theme: 0, tint: 0, indexed: 0 },
      { kind: kColorRgb, rgb: 0xff445566, theme: 0, tint: 0, indexed: 0 },
    ];
    wb.addCfRule(0, {
      type: 2,
      ranges: [{ firstRow: 0, firstCol: 0, lastRow: 9, lastCol: 0 }],
      colorScale: { thresholds: [{ type: 3 }, { type: 4 }], colors },
    });
    assert.deepEqual(wb.getCfRule(0, 0).colorScale.colors, colors);
  });
});

test('tab colour round-trips and clears', () => {
  withWorkbook((wb) => {
    assert.equal(wb.getTabColor(0), null);
    const rgb = { kind: kColorRgb, rgb: 0xffff0000, theme: 0, tint: 0, indexed: 0 };
    wb.setTabColor(0, rgb);
    assert.deepEqual(wb.getTabColor(0), rgb);
    const theme = { kind: kColorTheme, rgb: 0, theme: 4, tint: 0.5, indexed: 0 };
    wb.setTabColor(0, theme);
    assert.deepEqual(wb.getTabColor(0), theme);
    wb.setTabColor(0, null);
    assert.equal(wb.getTabColor(0), null);
    wb.setTabColor(0, { kind: kColorNone });
    assert.equal(wb.getTabColor(0), null);
    assert.throws(() => wb.setTabColor(0), TypeError);
    assert.throws(() => wb.setTabColor(0, 'red'), TypeError);
  });
});

test('the profile id accessors use the profile name', () => {
  withWorkbook((wb) => {
    assert.equal(wb.getProfileId(), 'win-365-en_US');
    wb.setProfileId('mac-365-ja_JP');
    assert.equal(wb.getProfileId(), 'mac-365-ja_JP');
    assert.equal(wb.getExcelProfileId, undefined);
  });
});
