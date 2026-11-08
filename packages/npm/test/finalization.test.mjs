// A WASM Workbook that is dropped without dispose() has its raw embind handle delete()d when the
// garbage collector reclaims it (best effort), and dispose() releases it at the call.
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".
//
// The release of a raw handle has no public signal, so this file spies on one seam: it installs a
// recording subclass of FinalizationRegistry before the entry module loads and reads the `isDeleted()`
// that embind puts on every raw handle. The registry is where the design places the release, and the
// registered value is expected to reach the raw handle (directly or as a property of what is held).
// Nothing else of the facade's internals is touched.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';
import v8 from 'node:v8';
import vm from 'node:vm';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

v8.setFlagsFromString('--expose-gc');
const gc = vm.runInNewContext('gc');

const registrations = [];
const unregistrations = [];
const RealRegistry = globalThis.FinalizationRegistry;
globalThis.FinalizationRegistry = class RecordingRegistry extends RealRegistry {
  register(target, held, token) {
    registrations.push({ held, token });
    return super.register(target, held, token);
  }

  unregister(token) {
    unregistrations.push(token);
    return super.unregister(token);
  }
};

const uncaught = [];
process.on('uncaughtException', (err) => uncaught.push(err));

async function loadFormulon() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  // A fresh module instance, so its registry is created after the recording subclass is installed.
  const href = `${pathToFileURL(path.resolve(pkgRoot, target)).href}?finalization`;
  const mod = await import(href);
  return mod.createFormulon();
}

// The raw embind handle inside a registered value: the value itself or one of its own properties.
function findRaw(held) {
  const isRaw = (v) => v !== null && typeof v === 'object' && typeof v.isDeleted === 'function';
  if (isRaw(held)) return held;
  if (held !== null && typeof held === 'object') {
    for (const v of Object.values(held)) if (isRaw(v)) return v;
    if (Array.isArray(held)) return held.find(isRaw);
  }
  return undefined;
}

const rawsSince = (start) =>
  registrations
    .slice(start)
    .map((r) => findRaw(r.held))
    .filter(Boolean);

async function collectUntil(done) {
  for (let round = 0; round < 100 && !done(); round += 1) {
    gc();
    await new Promise((resolve) => setImmediate(resolve));
  }
}

test('a Workbook dropped without dispose() has its raw handle deleted after garbage collection', async () => {
  const f = await loadFormulon();
  const start = registrations.length;
  const count = 4;
  (() => {
    for (let i = 0; i < count; i += 1) {
      const wb = f.Workbook.createDefault();
      wb.sheet(0).setNumber(0, 0, i);
    }
  })();
  const raws = rawsSince(start);
  assert.ok(raws.length >= count, `${count} workbooks register ${count} raw handles (saw ${raws.length})`);
  assert.ok(
    raws.every((raw) => !raw.isDeleted()),
    'a live workbook is not deleted',
  );
  await collectUntil(() => raws.every((raw) => raw.isDeleted()));
  assert.equal(
    raws.filter((raw) => raw.isDeleted()).length,
    raws.length,
    'every unreachable workbook must have its raw handle deleted',
  );
  assert.deepEqual(uncaught, []);
});

test('dispose() deletes the raw handle at the call, unregisters, and a later collection does nothing', async () => {
  const f = await loadFormulon();
  const start = registrations.length;
  const unregisteredBefore = unregistrations.length;
  let wb = f.Workbook.createDefault();
  const raws = rawsSince(start);
  assert.equal(raws.length, 1);
  assert.equal(raws[0].isDeleted(), false);
  wb.dispose();
  assert.equal(raws[0].isDeleted(), true, 'dispose releases immediately, not at collection');
  assert.ok(unregistrations.length > unregisteredBefore, 'dispose removes the registration');
  assert.equal(wb.disposed, true);
  wb = null;
  await collectUntil(() => false);
  assert.deepEqual(uncaught, [], 'collecting a disposed workbook must not delete twice');
});

test('a Sheet handle keeps its workbook alive, so the raw handle is not deleted while it is reachable', async () => {
  const f = await loadFormulon();
  const start = registrations.length;
  let sheet = (() => {
    const wb = f.Workbook.createDefault();
    return wb.sheet(0);
  })();
  const raws = rawsSince(start);
  assert.equal(raws.length, 1);
  for (let round = 0; round < 30; round += 1) {
    gc();
    await new Promise((resolve) => setImmediate(resolve));
  }
  assert.equal(raws[0].isDeleted(), false, 'the reachable sheet holds the workbook');
  sheet.setNumber(0, 0, 7);
  assert.equal(sheet.getTypedValue(0, 0).number, 7);
  const wb = sheet.workbook;
  sheet = null;
  wb.dispose();
  assert.equal(raws[0].isDeleted(), true);
});
