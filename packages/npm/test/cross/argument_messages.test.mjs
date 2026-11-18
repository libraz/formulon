// The WASM and native packages refuse the same bad arguments with the same error class, status code and message.
// It loads both staged packages, so it runs from `make surface-parity`, which stages both, and from neither package's
// own test target.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..', '..');

/** Both staged packages, each loaded through its own package entry. */
async function loadBothPackages() {
  const load = async (dir, pick) => {
    const root = path.join(repoRoot, 'packages', dir);
    const pkg = JSON.parse(await readFile(path.join(root, 'package.json'), 'utf8'));
    const target = pick(pkg);
    if (!target) throw new Error(`${dir}: package.json has no entry`);
    const mod = await import(pathToFileURL(path.resolve(root, target)).href);
    return mod.createFormulon();
  };
  return {
    wasm: await load('npm', (pkg) => pkg.exports?.['.']?.import),
    native: await load('npm-native', (pkg) => pkg.main),
  };
}

/** Class, status code and message of what `fn` throws on a freshly built workbook of `formulon`. */
function refusal(formulon, fn) {
  const wb = formulon.Workbook.createDefault();
  try {
    const ws = wb.sheet(0);
    const cache = wb.addPivotCache();
    const pivot = ws.addPivot({ name: 'P1', cacheId: cache.cacheId, anchor: { row: 10, col: 4 } });
    try {
      fn({ wb, ws, pivot, cache });
    } catch (e) {
      return { name: e.name, isRangeError: e instanceof RangeError, code: e.code ?? null, message: e.message };
    }
    return null;
  } finally {
    wb.dispose();
  }
}

test('argument error messages are identical on both packages', async () => {
  const { wasm, native } = await loadBothPackages();
  const battery = {
    'a fractional field': ({ ws }) => ws.updateView({ zoomPct: 1.5 }),
    'a negative field': ({ ws }) => ws.updatePageSetup({ scalePct: -5 }),
    'a table without a range': ({ ws }) => ws.addTable({}),
    'a validation without ranges': ({ ws }) => ws.addValidation({}),
    'a person with a numeric name': ({ wb }) => wb.addPerson({ displayName: 1 }),
    'a negative iteration cap': ({ wb }) => wb.updateIterative({ maxIterations: -1 }),
    'an iteration cap past the limit': ({ wb }) => wb.updateIterative({ maxIterations: 40000 }),
    'a non-boolean flag': ({ wb }) => wb.setDate1904('yes'),
    'a non-boolean field': ({ ws }) => ws.updateView({ showGridLines: 1 }),
    'an unknown key': ({ wb }) => wb.ooxml.addFont({ bogus: 1 }),
    'a pivot index past the end': ({ pivot }) => pivot.getDataField(99),
    'a pivot cache field past the end': ({ cache }) => cache.listFieldSharedItems(99),
  };
  for (const [label, fn] of Object.entries(battery)) {
    const fromWasm = refusal(wasm, fn);
    const fromNative = refusal(native, fn);
    assert.ok(fromWasm, `${label}: the WASM package accepted the call`);
    assert.ok(fromNative, `${label}: the native package accepted the call`);
    assert.ok(fromWasm.message.length > 0, label);
    assert.deepEqual(fromNative, fromWasm, `${label}: the packages disagree`);
    assert.doesNotMatch(
      fromWasm.message,
      /\bargument \d|`|\b(?:fm_|FM_)|\bNULL\b|_op\b/,
      `${label}: ${fromWasm.message}`,
    );
  }
});
