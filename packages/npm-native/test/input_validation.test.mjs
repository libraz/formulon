import assert from 'node:assert/strict';
import { spawnSync } from 'node:child_process';
import { existsSync } from 'node:fs';
import { createRequire } from 'node:module';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';
import { FormulonError, StatusCode } from '../dist/common.mjs';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const addonPath =
  process.env.FORMULON_NODE_ADDON ??
  [
    path.resolve(testDir, '..', 'dist', 'prebuilds', `${process.platform}-${process.arch}`, 'formulon.node'),
    path.resolve(testDir, '..', 'dist', 'formulon.node'),
    path.resolve(testDir, '..', 'formulon.node'),
  ].find((candidate) => existsSync(candidate));
assert.ok(addonPath, 'native addon not found; set FORMULON_NODE_ADDON or stage dist/formulon.node');
const requireAddon = createRequire(import.meta.url);
const addon = requireAddon(addonPath);
addon.__setErrorClass(FormulonError);
const commonUrl = pathToFileURL(path.resolve(testDir, '..', 'dist', 'common.mjs')).href;

function cellNumber(wb, row = 0, col = 0) {
  const value = wb.getValue(0, row, col);
  assert.equal(value.kind, 1);
  return value.number;
}

test('coercion objects and symbols are rejected before mutating a cell', () => {
  const source = `
    (async () => {
    const addon = require(process.env.FORMULON_NODE_ADDON);
    addon.__setErrorClass((await import(process.env.FORMULON_COMMON)).FormulonError);
    const wb = addon.Workbook.createDefault();
    const failures = [];
    function probe(label, call) {
      try {
        call();
        failures.push(label + ':returned');
      } catch (error) {
        if (!(error instanceof TypeError)) failures.push(label + ':' + error.name);
      }
      const cell = wb.getValue(0, 0, 0);
      if (cell.kind !== 1 || cell.number !== 42) {
        failures.push(label + ':mutated');
      }
    }
    wb.setNumber(0, 0, 0, 42);
    probe('valueOf', () => wb.setNumber(0, 0, 0, { valueOf() { throw new Error('valueOf'); } }));
    probe('symbol-number', () => wb.setNumber(0, 0, 0, Symbol('number')));
    probe('toString', () => wb.setText(0, 0, 0, { toString() { throw new Error('toString'); } }));
    probe('symbol-string', () => wb.setText(0, 0, 0, Symbol('string')));
    probe('detached', () => wb.setNumber.call({}, 0, 0, 0, 7));
    const forged = Object.create(addon.Workbook.prototype);
    probe('forged', () => wb.setNumber.call(forged, 0, 0, 0, 7));
    wb.dispose();
    if (failures.length) {
      console.error(JSON.stringify(failures));
      process.exitCode = 1;
    }
    })();
  `;
  const result = spawnSync(process.execPath, ['--input-type=commonjs', '-e', source], {
    env: { ...process.env, FORMULON_NODE_ADDON: addonPath, FORMULON_COMMON: commonUrl },
    encoding: 'utf8',
  });
  assert.equal(result.status, 0, `subprocess exited ${result.status}: ${result.stderr}`);
});

test('module-level numeric and string positions reject coercion and overflow', () => {
  const source = `
    (async () => {
    const addon = require(process.env.FORMULON_NODE_ADDON);
    addon.__setErrorClass((await import(process.env.FORMULON_COMMON)).FormulonError);
    const failures = [];
    function probe(label, expected, call) {
      try {
        call();
        failures.push(label + ':returned');
      } catch (error) {
        if (!(error instanceof expected)) failures.push(label + ':' + error.name);
      }
    }
    probe('status-symbol', TypeError, () => addon.getStatusName(Symbol('status')));
    probe('status-fraction', RangeError, () => addon.getStatusName(1.5));
    probe('status-overflow', RangeError, () => addon.getStatusName(2 ** 31));
    probe('log-symbol', TypeError, () => addon.setLogMinLevel(Symbol('log')));
    probe('log-underflow', RangeError, () => addon.setLogMinLevel(-(2 ** 31) - 1));
    probe('formula-symbol', TypeError, () => addon.evaluateFormula(Symbol('formula')));
    probe('formula-object', TypeError, () => addon.evaluateFormula({ toString() { throw new Error('toString'); } }));
    if (failures.length) {
      console.error(JSON.stringify(failures));
      process.exitCode = 1;
    }
    })();
  `;
  const result = spawnSync(process.execPath, ['--input-type=commonjs', '-e', source], {
    env: { ...process.env, FORMULON_NODE_ADDON: addonPath, FORMULON_COMMON: commonUrl },
    encoding: 'utf8',
  });
  assert.equal(result.status, 0, `subprocess exited ${result.status}: ${result.stderr}`);
});

test('integer positional arguments reject wrapping and non-finite values', () => {
  const wb = addon.Workbook.createDefault();
  try {
    wb.setNumber(0, 0, 0, 42);
    for (const value of [2 ** 32, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      assert.throws(() => wb.setNumber(value, 0, 0, 7), RangeError, `sheet=${String(value)}`);
      assert.equal(cellNumber(wb), 42, `invalid sheet changed A1 for ${String(value)}`);
    }
    for (const value of [2 ** 32, 2 ** 31, -(2 ** 31) - 1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      assert.throws(() => wb.setError(0, 0, 0, value), RangeError, `errorCode=${String(value)}`);
      assert.equal(cellNumber(wb), 42, `invalid error code changed A1 for ${String(value)}`);
    }
  } finally {
    wb.dispose();
  }
});

test('iterative caps use the unsigned int32 contract', () => {
  const wb = addon.Workbook.createDefault();
  try {
    wb.updateIterative({ enabled: true, maxIterations: 2 ** 32 - 1, maxChange: 0.1 });
    const accepted = wb.getIterative();
    assert.throws(
      () => wb.updateIterative({ enabled: false, maxIterations: 0 }),
      (e) => e instanceof FormulonError && e.code === StatusCode.InvalidArgument,
      'a zero cap is rejected, not clamped',
    );
    assert.deepEqual({ ...wb.getIterative() }, { ...accepted }, 'a rejected patch changes no key');
    for (const value of [2 ** 32, -1]) {
      assert.throws(() => wb.updateIterative({ maxIterations: value }), RangeError, `maxIterations=${String(value)}`);
    }
    for (const value of [1.5, Number.NaN, Number.POSITIVE_INFINITY, '5', null]) {
      assert.throws(() => wb.updateIterative({ maxIterations: value }), TypeError, `maxIterations=${String(value)}`);
    }
    assert.equal(wb.getIterative().maxIterations, accepted.maxIterations);
  } finally {
    wb.dispose();
  }
});

test('sheet, structural, enum, and read positions reject non-number primitives', () => {
  const wb = addon.Workbook.createDefault();
  try {
    assert.throws(() => wb.getSheetName(Symbol('sheet')), TypeError);
    assert.throws(() => wb.insertRows(0, 0, 1.5), RangeError);
    assert.throws(() => wb.setCalcMode(1.5), RangeError);
    assert.throws(() => wb.getValue(Symbol('sheet'), 0, 0), TypeError);
    assert.throws(() => addon.localizeFunctionName('SUM', 1), TypeError);
    assert.throws(() => wb.resolveColor({ kind: 1, argb: 0 }, 2 ** 31), RangeError);
    assert.throws(
      () => wb.getCellsInRange(0, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, Symbol('cursor')),
      TypeError,
    );
  } finally {
    wb.dispose();
  }
});

test('paging cursor and limit reject values outside native integer ranges', () => {
  const wb = addon.Workbook.createDefault();
  const range = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };
  try {
    for (const value of [-1, 1.5, Number.NaN, Number.POSITIVE_INFINITY, 2 ** 53, 2 ** 64]) {
      assert.throws(() => wb.getCellsInRange(0, range, value), RangeError, `get cursor=${String(value)}`);
      assert.throws(() => wb.getInvalidCells(0, value), RangeError, `invalid cursor=${String(value)}`);
    }
    for (const value of [2 ** 32, -1, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      assert.throws(() => wb.getCellsInRange(0, range, undefined, value), RangeError, `get limit=${String(value)}`);
      assert.throws(() => wb.getInvalidCells(0, undefined, value), RangeError, `invalid limit=${String(value)}`);
    }
    const refusedByHost = (e) => e instanceof FormulonError && e.code === 2;
    assert.throws(() => wb.getCellsInRange(0, range, Number.MAX_SAFE_INTEGER), refusedByHost);
    assert.throws(() => wb.getInvalidCells(0, Number.MAX_SAFE_INTEGER), refusedByHost);
  } finally {
    wb.dispose();
  }
});

test('missing required positional arguments are rejected; optional ones keep their defaults', () => {
  const wb = addon.Workbook.createDefault();
  try {
    assert.throws(() => wb.updateIterative(), TypeError);
    assert.throws(() => wb.setVisibility(0), TypeError);
    assert.equal(wb.getIterative().maxIterations, 100, 'the rejected updateIterative must not store a cap');
    addon.getFunctionMetadata('SUM');
    const range = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };
    assert.doesNotThrow(() => wb.getCellsInRange(0, range, undefined, undefined));
    assert.doesNotThrow(() => wb.getInvalidCells(0, undefined, undefined));
    assert.doesNotThrow(() => wb.getCellsInRange(0, range, null, null));
    assert.doesNotThrow(() => wb.getInvalidCells(0, null, null));
  } finally {
    wb.dispose();
  }
});

test('getCellsInRange requires its range argument like the merge methods', () => {
  const wb = addon.Workbook.createDefault();
  try {
    wb.setNumber(0, 0, 0, 7);
    assert.throws(() => wb.getCellsInRange(0, null), TypeError);
    assert.throws(() => wb.getMergesInRange(0, null), TypeError);
    assert.throws(() => wb.getCellsInRange(0), TypeError);
    assert.throws(() => wb.getCellsInRange(0, 5), TypeError);
    assert.throws(() => wb.getMergesInRange(0, 5), TypeError);
  } finally {
    wb.dispose();
  }
});
