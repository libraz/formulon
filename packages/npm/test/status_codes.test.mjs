import test from 'node:test';
import { assert, getModule, loadStagedModule, path, pkgRoot, readFile } from './smoke_support.mjs';

const headerPath = path.resolve(pkgRoot, '..', '..', 'src', 'c_api', 'formulon_status.h');

async function headerCodes() {
  const text = await readFile(headerPath, 'utf8');
  return new Map([...text.matchAll(/^#define FM_STATUS_(\w+) \((\d+)\)$/gm)].map((m) => [m[1], Number(m[2])]));
}

const upperSnake = (key) => key.replace(/([a-z])([A-Z])/g, '$1_$2').toUpperCase();

test('StatusCode mirrors the C status defines', async () => {
  const mod = await loadStagedModule();
  const expected = await headerCodes();
  assert.equal(expected.size, 57);
  assert.ok(Object.isFrozen(mod.StatusCode));
  const actual = new Map(Object.entries(mod.StatusCode).map(([key, value]) => [upperSnake(key), value]));
  assert.deepEqual([...actual].sort(), [...expected].sort());
});

test('StatusCode keys are PascalCase code names', async () => {
  const { StatusCode } = await loadStagedModule();
  assert.equal(StatusCode.Ok, 0);
  assert.equal(StatusCode.InvalidArgument, 2);
  assert.equal(StatusCode.ParserInvalidReference, 1004);
  assert.equal(StatusCode.SheetIndexOutOfRange, 10);
  assert.equal(StatusCode.BindingInvalidHandle, 7000);
  for (const key of Object.keys(StatusCode)) {
    assert.match(key, /^[A-Z][A-Za-z0-9]*$/);
  }
});

test('renamed and renumbered codes: the old names are gone and every workbook-structural code sits in the general band', async () => {
  const { StatusCode } = await getModule();
  assert.equal(StatusCode.FunctionOverflow, 3004);
  assert.equal(StatusCode.SecurityResourceLimit, 6005);
  assert.equal('FnOverflow' in StatusCode, false);
  assert.equal('SecResourceLimit' in StatusCode, false);
  assert.deepEqual(
    [StatusCode.SheetIndexOutOfRange, StatusCode.InvalidSheetName, StatusCode.IndexOutOfRange],
    [10, 11, 13],
  );
  for (const gone of ['SheetCountLimitExceeded', 'PrintPageCountOverflow']) {
    assert.equal(gone in StatusCode, false, `${gone} is retired`);
  }
  for (const retired of [12, 9005]) {
    assert.equal(Object.values(StatusCode).includes(retired), false, `${retired} is retired`);
  }
  for (const retired of [5050, 5051, 5052, 5053, 5054]) {
    assert.equal(Object.values(StatusCode).includes(retired), false, `${retired} is retired`);
  }
});

test('getStatusName is no longer part of the module', async () => {
  const Module = await getModule();
  assert.equal(Module.getStatusName, undefined);
});
