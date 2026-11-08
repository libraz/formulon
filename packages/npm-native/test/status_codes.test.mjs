import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';
import { getModule, loadStagedModule } from './smoke_support.mjs';

const headerPath = path.resolve(
  path.dirname(fileURLToPath(import.meta.url)),
  '..',
  '..',
  '..',
  'src',
  'c_api',
  'formulon_status.h',
);

async function headerCodes() {
  const text = await readFile(headerPath, 'utf8');
  return new Map([...text.matchAll(/^#define FM_STATUS_(\w+) \((\d+)\)$/gm)].map((m) => [m[1], Number(m[2])]));
}

const upperSnake = (key) => key.replace(/([a-z])([A-Z])/g, '$1_$2').toUpperCase();

test('StatusCode mirrors the C status defines', async () => {
  const mod = await loadStagedModule();
  const expected = await headerCodes();
  assert.equal(expected.size, 112);
  assert.ok(Object.isFrozen(mod.StatusCode));
  const actual = new Map(Object.entries(mod.StatusCode).map(([key, value]) => [upperSnake(key), value]));
  assert.deepEqual([...actual].sort(), [...expected].sort());
  assert.equal((await mod.createFormulon()).StatusCode, mod.StatusCode);
});

test('StatusCode keys are PascalCase code names', async () => {
  const { StatusCode } = await getModule();
  assert.equal(StatusCode.Ok, 0);
  assert.equal(StatusCode.InvalidArgument, 2);
  assert.equal(StatusCode.ParserInvalidR1C1, 1011);
  assert.equal(StatusCode.SheetIndexOutOfRange, 5050);
  assert.equal(StatusCode.BindingInvalidHandle, 7000);
  for (const key of Object.keys(StatusCode)) {
    assert.match(key, /^[A-Z][A-Za-z0-9]*$/);
  }
});

test('getStatusName names a code without its prefix and flags unknown values', async () => {
  const mod = await getModule();
  assert.equal(mod.getStatusName(mod.StatusCode.Ok), 'OK');
  assert.equal(mod.getStatusName(mod.StatusCode.SheetIndexOutOfRange), 'SHEET_INDEX_OUT_OF_RANGE');
  assert.equal(mod.getStatusName(123456), 'UNRECOGNIZED_STATUS');
});
