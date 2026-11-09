// A host without a matching prebuild: the entry module still imports, and createFormulon() rejects.

import assert from 'node:assert/strict';
import { execFileSync } from 'node:child_process';
import { mkdtempSync, rmSync } from 'node:fs';
import { tmpdir } from 'node:os';
import path from 'node:path';
import test from 'node:test';

import { pkgRoot } from './smoke_support.mjs';

const probe = `
import { pathToFileURL } from 'node:url';
const entry = process.argv[1];
const mod = await import(pathToFileURL(entry).href);
const common = await import(pathToFileURL(entry.replace(/index\\.mjs$/, 'common.mjs')).href);
let outcome;
try {
  await mod.createFormulon();
  outcome = { rejected: false };
} catch (e) {
  outcome = {
    rejected: true,
    isError: e instanceof Error,
    isFormulonError: e instanceof common.FormulonError,
    message: e.message,
  };
}
console.log(JSON.stringify({ imported: typeof mod.createFormulon === 'function', ...outcome }));
`;

test('with no prebuild the entry imports and createFormulon() rejects with a plain Error naming the platform', () => {
  const emptyRoot = mkdtempSync(path.join(tmpdir(), 'formulon-no-prebuild-'));
  try {
    const entry = path.resolve(pkgRoot, 'dist', 'index.mjs');
    const out = execFileSync(process.execPath, ['--input-type=module', '-e', probe, entry], {
      env: { ...process.env, FORMULON_NATIVE_ROOT: emptyRoot },
      encoding: 'utf8',
    });
    const result = JSON.parse(out.trim().split('\n').at(-1));
    assert.equal(result.imported, true, 'the static import resolves');
    assert.equal(result.rejected, true);
    assert.equal(result.isError, true);
    assert.equal(result.isFormulonError, false, 'a missing prebuild is a host failure');
    assert.ok(result.message.includes(`${process.platform}-${process.arch}`), result.message);
  } finally {
    rmSync(emptyRoot, { recursive: true, force: true });
  }
});
