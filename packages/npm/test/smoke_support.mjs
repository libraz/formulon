// Shared staged-module loader and Module promise for the WASM smoke-test topics.

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

// Resolve the staged entry point through the package's own "exports" map,
// so a broken entry fails here rather than in a consumer. FORMULON_NPM_ENTRY
// selects the subpath: "." (single-threaded, default) or "./threads".
// Fails (rather than skips) if the package isn't staged -- that's the whole
// point of running these tests against dist/.
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';
const threadsEntry = entry === './threads';

async function loadStagedModule() {
  const raw = await readFile(pkgJsonPath, 'utf8');
  const pkg = JSON.parse(raw);
  const target = pkg.exports?.[entry]?.import;
  if (!target) {
    throw new Error(`package.json "exports" has no import target for ${entry}: ${pkgJsonPath}`);
  }
  // Node ESM requires file:// URLs for absolute paths on Windows;
  // pathToFileURL is the portable form.
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

async function loadStagedFactory() {
  const mod = await loadStagedModule();
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

// A self-referential average: from a blank A1 the iterative solver walks
// towards 10, and the progress callback fires after every sweep.
function makeIterativeWorkbook(Module) {
  const wb = Module.Workbook.createDefault();
  assert.ok(wb.setIterative(true, 100, 0.0001).ok);
  assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
  return wb;
}

// A structured reference the XLSB encoder cannot lower, which
// makes the writer emit one per-cell warn record.
function emitXlsbWarning(Module, xlsbFormat) {
  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(wb.setFormula(0, 0, 0, '=SUM(T[C])').ok);
    assert.ok(wb.recalc().ok);
    assert.ok(wb.saveAs(xlsbFormat).status.ok);
  } finally {
    wb.delete();
  }
}

export {
  assert,
  emitXlsbWarning,
  entry,
  getModule,
  loadStagedFactory,
  loadStagedModule,
  makeIterativeWorkbook,
  path,
  pkgJsonPath,
  pkgRoot,
  readFile,
  threadsEntry,
  VAL,
};
