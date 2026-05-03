// Copyright 2026 libraz. Licensed under the MIT License.
//
// Smoke tests for the staged @libraz/formulon-native npm package.
//
// These tests intentionally import the *staged* dist/ artefacts (resolved
// via the package's own package.json "main" field), not the raw build/
// output. That way we catch staging mistakes -- missing files, wrong
// relative paths, broken "exports" entries -- before the tarball ships.
//
// Run via `make node-test` (which runs `make node-package` first).

import test from 'node:test';
import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const pkgRoot = path.resolve(__dirname, '..');
const pkgJsonPath = path.join(pkgRoot, 'package.json');

// Resolve the staged entry point through the package's own package.json
// "main" field. Fails (rather than skips) if the package isn't staged --
// that's the whole point of running these tests against dist/.
async function loadStagedModule() {
  const raw = await readFile(pkgJsonPath, 'utf8');
  const pkg = JSON.parse(raw);
  const main = pkg.main;
  if (!main) {
    throw new Error(`package.json is missing "main": ${pkgJsonPath}`);
  }
  const mainPath = path.resolve(pkgRoot, main);
  // Node ESM requires file:// URLs for absolute paths on Windows;
  // pathToFileURL is the portable form.
  return import(pathToFileURL(mainPath).href);
}

let modPromise;
function getModule() {
  if (!modPromise) {
    modPromise = loadStagedModule();
  }
  return modPromise;
}

test('default export exposes Workbook + evalFormula + version', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.Workbook, 'function');
  assert.equal(typeof mod.evalFormula, 'function');
  assert.equal(typeof mod.version, 'function');
  assert.equal(typeof mod.statusString, 'function');
  assert.equal(typeof mod.ValueKind, 'object');
});

test('version() returns a non-empty string', async () => {
  const mod = await getModule();
  const v = mod.version();
  assert.equal(typeof v, 'string');
  assert.ok(v.length > 0, `expected non-empty version, got ${JSON.stringify(v)}`);
});

test("evalFormula('=1+2') returns NUMBER 3", async () => {
  const mod = await getModule();
  const r = mod.evalFormula('=1+2');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 3);
});

test("evalFormula('=#REF!') surfaces ERROR value (not failed status)", async () => {
  const mod = await getModule();
  const r = mod.evalFormula('=#REF!');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Error);
});

test('Workbook.createDefault + setNumber + getValue round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.sheetCount(), 1);
  const setRc = wb.setNumber(0, 0, 0, 42);
  assert.ok(setRc.ok, `setNumber: ${JSON.stringify(setRc)}`);
  const r = wb.getValue(0, 0, 0);
  assert.ok(r.status.ok, `getValue: ${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 42);
});

test("Workbook.createDefault + setFormula '=1+2' + recalc -> 3", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setFormula(0, 0, 0, '=1+2').ok);
  assert.ok(wb.recalc().ok);
  const r = wb.getValue(0, 0, 0);
  assert.ok(r.status.ok);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 3);
});

test('Workbook.createEmpty + addSheet -> sheetCount/sheetName work', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createEmpty();
  assert.ok(wb.addSheet('Sheet1').ok);
  assert.equal(wb.sheetCount(), 1);
  const sn = wb.sheetName(0);
  assert.ok(sn.status.ok);
  assert.equal(sn.value, 'Sheet1');
});

test('save() returns Uint8Array; loadBytes round-trips the value', async () => {
  const mod = await getModule();
  const wb1 = mod.Workbook.createDefault();
  assert.ok(wb1.setNumber(0, 0, 0, 42).ok);
  const sr = wb1.save();
  assert.ok(sr.status.ok, `save: ${JSON.stringify(sr.status)}`);
  assert.ok(sr.bytes instanceof Uint8Array, 'expected Uint8Array');
  assert.ok(sr.bytes.length > 0, 'expected non-empty save buffer');

  const wb2 = mod.Workbook.loadBytes(sr.bytes);
  const r = wb2.getValue(0, 0, 0);
  assert.ok(r.status.ok, `loaded getValue: ${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 42);
});
