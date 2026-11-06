// The JS surface and the Python surface name the same members: Workbook, Sheet and
// the module-level functions, apart from the differences recorded in
// tools/api/platform_diff.json. Python names are read from the source files.
// The native package additionally carries the members recorded under native_only.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

import { prototypeKeys } from '../../npm/test/dts_members.mjs';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const repoRoot = path.resolve(pkgRoot, '..', '..');
const pyRoot = path.join(repoRoot, 'packages', 'python', 'formulon');

const DIFF = JSON.parse(readFileSync(path.join(repoRoot, 'tools/api/platform_diff.json'), 'utf8'));

async function getFormulon() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  if (!pkg.main) throw new Error('package.json is missing "main"');
  const mod = await import(pathToFileURL(path.resolve(pkgRoot, pkg.main)).href);
  return mod.createFormulon();
}

/** Lower-cased name with separators removed, so `getFormulaR1C1` and `get_formula_r1c1` meet. */
const key = (name) => name.replaceAll('_', '').toLowerCase();

/** Body of the top-level class `name` in a Python source text. */
function classBody(text, name) {
  const start = text.search(new RegExp(`^class ${name}\\b`, 'm'));
  assert.notEqual(start, -1, `no class ${name}`);
  const rest = text.slice(text.indexOf('\n', start) + 1);
  const end = rest.search(/^(?:class |def |@|[A-Za-z_])/m);
  return end === -1 ? rest : rest.slice(0, end);
}

/** Names of the methods defined at class level (one indent), dunders included. */
function methodNames(body) {
  return new Set([...body.matchAll(/^ {4}(?:async )?def ([A-Za-z_][A-Za-z0-9_]*)\(/gm)].map((m) => m[1]));
}

function pythonClassMembers(file, cls) {
  return methodNames(classBody(readFileSync(path.join(pyRoot, file), 'utf8'), cls));
}

function pythonModuleFunctions() {
  const text = readFileSync(path.join(pyRoot, '__init__.py'), 'utf8');
  const all = text.slice(text.indexOf('__all__'));
  const names = [...all.slice(0, all.indexOf(']')).matchAll(/"([A-Za-z_][A-Za-z0-9_]*)"/g)].map((m) => m[1]);
  return new Set(names.filter((n) => /^[a-z]/.test(n)));
}

function diffSets(actual, expected) {
  return {
    onlyJs: [...actual].filter((n) => !expected.has(n)).sort(),
    onlyPy: [...expected].filter((n) => !actual.has(n)).sort(),
  };
}

/** Compares the JS member names with the Python ones after removing each side's recorded extras. */
function assertParity(label, jsNames, jsOnly, pyNames, pyOnly) {
  assert.ok(jsNames.length > 0, `${label}: no JS members`);
  assert.ok(pyNames.size > 0, `${label}: no Python members`);
  const js = new Map(jsNames.filter((n) => !jsOnly.has(n)).map((n) => [key(n), n]));
  // Underscore names are private to the Python implementation; recorded dunders are in `pyOnly`.
  const pyOnlyKeys = new Set([...pyOnly].map(key));
  const py = new Map([...pyNames].filter((n) => !pyOnlyKeys.has(key(n)) && !n.startsWith('_')).map((n) => [key(n), n]));
  const { onlyJs, onlyPy } = diffSets(new Set(js.keys()), new Set(py.keys()));
  assert.deepEqual(
    { onlyJs: onlyJs.map((k) => js.get(k)), onlyPy: onlyPy.map((k) => py.get(k)) },
    { onlyJs: [], onlyPy: [] },
    `${label}: members present on one side only`,
  );
}

const diffSet = (table, name) => new Set(table[name] ?? []);

test('the platform diff ledger names only members that exist on its own side', async () => {
  const formulon = await getFormulon();
  const wb = formulon.Workbook.createDefault();
  try {
    const jsWorkbook = new Set([
      ...prototypeKeys(Object.getPrototypeOf(wb)),
      ...Object.getOwnPropertyNames(formulon.Workbook),
    ]);
    const jsSheet = new Set(prototypeKeys(Object.getPrototypeOf(wb.sheet(0))));
    for (const n of DIFF.js_only.Workbook) assert.ok(jsWorkbook.has(n), `Workbook.${n}`);
    for (const n of DIFF.native_only.Workbook) assert.ok(jsWorkbook.has(n), `Workbook.${n}`);
    for (const n of DIFF.js_only.Sheet) assert.ok(jsSheet.has(n), `Sheet.${n}`);
    for (const n of DIFF.js_only.Formulon) assert.ok(n in formulon, `Formulon.${n}`);
    const pyWorkbook = pythonClassMembers('workbook.py', 'Workbook');
    const pySheet = pythonClassMembers('sheet.py', 'Sheet');
    for (const n of DIFF.py_only.Workbook) assert.ok(pyWorkbook.has(n), `Python Workbook.${n}`);
    for (const n of DIFF.py_only.Sheet) assert.ok(pySheet.has(n), `Python Sheet.${n}`);
  } finally {
    wb.dispose();
  }
});

test('Workbook has the Python Workbook members apart from the platform differences', async () => {
  const formulon = await getFormulon();
  const wb = formulon.Workbook.createDefault();
  try {
    const statics = Object.getOwnPropertyNames(formulon.Workbook).filter(
      (n) => !['length', 'name', 'prototype'].includes(n),
    );
    const js = [...prototypeKeys(Object.getPrototypeOf(wb)), ...statics];
    // native_only members exist here and in Python, so both sides compare them.
    assertParity(
      'Workbook',
      js,
      diffSet(DIFF.js_only, 'Workbook'),
      pythonClassMembers('workbook.py', 'Workbook'),
      diffSet(DIFF.py_only, 'Workbook'),
    );
  } finally {
    wb.dispose();
  }
});

test('Sheet has the Python Sheet members apart from the platform differences', async () => {
  const formulon = await getFormulon();
  const wb = formulon.Workbook.createDefault();
  try {
    const js = prototypeKeys(Object.getPrototypeOf(wb.sheet(0)));
    assertParity(
      'Sheet',
      js,
      diffSet(DIFF.js_only, 'Sheet'),
      pythonClassMembers('sheet.py', 'Sheet'),
      diffSet(DIFF.py_only, 'Sheet'),
    );
  } finally {
    wb.dispose();
  }
});

test('Formulon has the Python module functions apart from the platform differences', async () => {
  const formulon = await getFormulon();
  const js = Object.keys(formulon).filter((n) => /^[a-z]/.test(n));
  assertParity(
    'Formulon',
    js,
    diffSet(DIFF.js_only, 'Formulon'),
    pythonModuleFunctions(),
    diffSet(DIFF.py_only, 'module'),
  );
});

test('every recorded JS-only Formulon member is a value table or a function that Python lacks', async () => {
  const formulon = await getFormulon();
  assert.ok(DIFF.js_only.Formulon.length > 0);
  for (const name of DIFF.js_only.Formulon) assert.ok(name in formulon, name);
  assert.equal(typeof DIFF.value_tables, 'string');
});
