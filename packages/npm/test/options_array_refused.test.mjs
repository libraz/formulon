// An array passed as the options argument is a TypeError, not an empty options bag, for parseReference,
// localizeFormula, canonicalizeFormula, restoreDrawing and addPicture.
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

const f = await (await loadPackage()).createFormulon();

const PNG_1X1 = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='),
  (c) => c.charCodeAt(0),
);

function refused(fn, call) {
  assert.throws(fn, (e) => {
    assert.ok(e instanceof TypeError, `expected a TypeError, got ${e}`);
    assert.match(e.message, new RegExp(`^${call} options must be an object$`));
    return true;
  });
}

describe('an array options argument', () => {
  it('is refused by the five options-taking calls', () => {
    const wb = f.Workbook.createDefault();
    try {
      const ws = wb.sheet(0);
      refused(() => wb.parseReference('A1', []), 'Workbook\\.parseReference');
      refused(() => f.localizeFormula('=SUM(A1)', 'win-365-en_US', []), '(?:\\w+\\.)?localizeFormula');
      refused(() => f.canonicalizeFormula('=SUM(A1)', 'win-365-en_US', []), '(?:\\w+\\.)?canonicalizeFormula');
      refused(() => ws.restoreDrawing(PNG_1X1, []), 'Sheet\\.restoreDrawing');
      refused(() => ws.addPicture(PNG_1X1, []), 'Sheet\\.addPicture');
    } finally {
      wb.dispose();
    }
  });
});
