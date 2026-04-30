// Copyright 2026 libraz. Licensed under the MIT License.
//
// Smoke tests for the staged @libraz/formulon npm package.
//
// These tests intentionally import the *staged* dist/ artefacts (resolved
// via the package's own package.json "main" field), not the raw
// build-wasm/ output. That way we catch staging mistakes -- missing
// files, wrong relative paths, broken "exports" entries -- before the
// tarball ships.
//
// Run via `make npm-test` (which runs `make npm-package` first).

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

// fm_value_kind_t mirror (see src/c_api/formulon_c.h).
const VAL = Object.freeze({
  BLANK: 0,
  NUMBER: 1,
  BOOL: 2,
  TEXT: 3,
  ERROR: 4,
  ARRAY: 5,
  REF: 6,
  LAMBDA: 7,
});

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const pkgRoot = path.resolve(__dirname, '..');
const pkgJsonPath = path.join(pkgRoot, 'package.json');

// Resolve the staged entry point through the package's own package.json
// "main" field. Fails (rather than skips) if the package isn't staged --
// that's the whole point of running these tests against dist/.
async function loadStagedFactory() {
  const raw = await readFile(pkgJsonPath, 'utf8');
  const pkg = JSON.parse(raw);
  const main = pkg.main;
  if (!main) {
    throw new Error(`package.json is missing "main": ${pkgJsonPath}`);
  }
  const mainPath = path.resolve(pkgRoot, main);
  // Node ESM requires file:// URLs for absolute paths on Windows;
  // pathToFileURL is the portable form.
  const mod = await import(pathToFileURL(mainPath).href);
  if (typeof mod.default !== 'function') {
    throw new Error(`expected default export to be a factory, got ${typeof mod.default}`);
  }
  return mod.default;
}

// Load once and reuse across tests; the factory itself is cheap, but the
// underlying WASM instantiation costs ~30ms per invocation.
let modulePromise;
function getModule() {
  if (!modulePromise) {
    modulePromise = (async () => {
      const factory = await loadStagedFactory();
      return factory();
    })();
  }
  return modulePromise;
}

test('default export is callable factory returning a Module', async () => {
  const factory = await loadStagedFactory();
  assert.equal(typeof factory, 'function');
  const Module = await factory();
  assert.ok(Module && typeof Module === 'object');
  assert.equal(typeof Module.versionString, 'function');
});

test('versionString returns a non-empty string', async () => {
  const Module = await getModule();
  const v = Module.versionString();
  assert.equal(typeof v, 'string');
  assert.ok(v.length > 0, `expected non-empty version, got ${JSON.stringify(v)}`);
});

test('evalFormula(=SUM(1,2,3)) returns NUMBER 6', async () => {
  const Module = await getModule();
  const r = Module.evalFormula('=SUM(1,2,3)');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, VAL.NUMBER);
  assert.equal(r.value.number, 6);
});

test('evalFormula(=1/0) surfaces Excel error as a value (not failed status)', async () => {
  const Module = await getModule();
  const r = Module.evalFormula('=1/0');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, VAL.ERROR);
});

test('Workbook.createDefault produces a valid single-sheet workbook', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(wb.isValid());
    assert.equal(wb.sheetCount(), 1);
  } finally {
    wb.delete();
  }
});

test('setNumber + getValue round-trips bit-exactly', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    const x = 3.141592653589793;
    assert.ok(wb.setNumber(0, 0, 0, x).ok);
    assert.ok(wb.recalc().ok);
    const r = wb.getValue(0, 0, 0);
    assert.ok(r.status.ok);
    assert.equal(r.value.kind, VAL.NUMBER);
    assert.equal(r.value.number, x);
  } finally {
    wb.delete();
  }
});

test('setFormula + recalc computes 41 + 1 == 42 in B1', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(wb.setNumber(0, 0, 0, 41).ok);
    assert.ok(wb.setFormula(0, 0, 1, '=A1+1').ok);
    assert.ok(wb.recalc().ok);
    const b1 = wb.getValue(0, 0, 1);
    assert.ok(b1.status.ok);
    assert.equal(b1.value.kind, VAL.NUMBER);
    assert.equal(b1.value.number, 42);
  } finally {
    wb.delete();
  }
});
