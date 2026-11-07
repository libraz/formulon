// Contract of the raw embind module (no JS facade): record-carrying calls,
// formula syntax errors with their position, colour specs, tab colours and
// the memory estimate. Mirrors the native addon's raw adoption test.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..');

/** Records exactly the arguments the module constructs its error with. */
class RawError extends Error {
  constructor(code, message, context, parseError) {
    super(message);
    this.code = code;
    this.context = context;
    this.parseError = parseError;
  }
}

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

function thrown(call) {
  try {
    call();
  } catch (error) {
    return error;
  }
  assert.fail('expected the call to throw');
}

/**
 * Registers the raw adoption cases. `factory` is the module factory; the
 * cases share one instance built with `RawError` as its error class.
 */
export function registerAdoption(factory, test) {
  let modulePromise;
  const raw = () => {
    modulePromise ??= factory({ FormulonError: RawError });
    return modulePromise;
  };
  const withWorkbook = async (body) => {
    const Module = await raw();
    const wb = Module.Workbook.createDefault();
    try {
      body(wb, Module);
    } finally {
      wb.delete();
    }
  };

  test('raw: record-carrying calls round-trip through struct_size-bearing records', () =>
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
    }));

  test('raw: a formula syntax error carries its code and UTF-16 position', () => {
    const codes = parseErrorCodes();
    const cases = positionedCases();
    assert.ok(cases.length > 0);
    return withWorkbook((wb) => {
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

  test('raw: a failure that is not a syntax error carries no parseError', () =>
    withWorkbook((wb) => {
      const error = thrown(() => wb.getValue(99, 0, 0));
      assert.ok(error instanceof RawError);
      assert.equal(error.parseError, undefined);
    }));

  test('raw: a CF rule formula syntax error names its field', () =>
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
    }));

  test('raw: the module evaluateFormula returns the whole result', async () => {
    const Module = await raw();
    assert.deepEqual(Module.evaluateFormula('=SUM(1,2,3)'), { rows: 1, cols: 1, cells: [[{ kind: 1, number: 6 }]] });
    const grid = Module.evaluateFormula('={1,2;3,4}');
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

  test('raw: CF colours are colour specs', () =>
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
    }));

  test('raw: tab colour round-trips and clears', () =>
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
    }));

  test('raw: the profile id accessors use the profile name', () =>
    withWorkbook((wb) => {
      assert.equal(wb.getProfileId(), 'win-365-en_US');
      wb.setProfileId('mac-365-ja_JP');
      assert.equal(wb.getProfileId(), 'mac-365-ja_JP');
      assert.equal(wb.getExcelProfileId, undefined);
    }));

  test('raw: getMemoryUsage estimates the workbook and reads 0 once disposed', async () => {
    const Module = await raw();
    const wb = Module.Workbook.createDefault();
    try {
      const empty = wb.getMemoryUsage();
      assert.equal(typeof empty, 'number');
      for (let row = 0; row < 1000; row += 1) {
        wb.setNumber(0, row, 0, row);
      }
      assert.ok(wb.getMemoryUsage() > empty);
      wb.dispose();
      assert.equal(wb.getMemoryUsage(), 0);
    } finally {
      wb.delete();
    }
  });
}
