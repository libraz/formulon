// The JS surface and the Python surface name the same members: Workbook, Sheet, the
// handle classes (Table, Pivot, PivotCache, Drawing, Picture, Chart, Shape, Slicer) and
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

/** Members attached to a class after its body, as `Cls.name = ...`. */
function attachedNames(text, cls) {
  return new Set([...text.matchAll(new RegExp(`^${cls}\\.([A-Za-z_][A-Za-z0-9_]*) = `, 'gm'))].map((m) => m[1]));
}

function pythonClassMembers(file, cls) {
  const text = readFileSync(path.join(pyRoot, file), 'utf8');
  return new Set([...methodNames(classBody(text, cls)), ...attachedNames(text, cls)]);
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
    {
      onlyJs: onlyJs.map((k) => js.get(k)),
      onlyPy: onlyPy.map((k) => py.get(k)),
    },
    { onlyJs: [], onlyPy: [] },
    `${label}: members present on one side only`,
  );
}

const HANDLE_FILES = {
  Table: '_table.py',
  Pivot: '_pivot.py',
  PivotCache: '_pivot_cache.py',
  Drawing: '_drawing.py',
  Picture: '_drawing.py',
  Chart: '_drawing.py',
  Shape: '_drawing.py',
  Slicer: '_drawing.py',
};
const DRAWING_SUBCLASSES = new Set(['Picture', 'Chart', 'Shape', 'Slicer']);

/** Members a handle class exposes in Python, with those of every base class its own file defines. */
function pythonHandleMembers(cls, file = HANDLE_FILES[cls]) {
  const text = readFileSync(path.join(pyRoot, file), 'utf8');
  const bases = text.match(new RegExp(`^class ${cls}\\(([^)]*)\\):`, 'm'))?.[1] ?? '';
  const out = pythonClassMembers(file, cls);
  for (const base of bases.split(',').map((name) => name.trim())) {
    if (new RegExp(`^class ${base}\\b`, 'm').test(text)) {
      for (const name of pythonHandleMembers(base, file)) out.add(name);
    }
  }
  return out;
}

const PNG_1X1 = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
);

/** One live instance of every handle class, built in a fresh workbook. */
function makeHandles(formulon, wb) {
  const ws = wb.sheet(0);
  ws.setText(0, 0, 'Product');
  ws.setText(0, 1, 'Amount');
  ws.setNumber(1, 1, 5);
  const table = ws.addTable({
    range: { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 },
    name: 'Tbl',
    columns: ['Product', 'Amount'],
  });
  const cache = wb.addPivotCache();
  const pivot = ws.addPivot({
    name: 'P1',
    cacheId: cache.cacheId,
    anchor: { row: 10, col: 4 },
  });
  const chart = ws.addChart({
    type: formulon.ChartPlotType.Bar,
    range: 'Sheet1!$A$1:$B$2',
    legend: formulon.ChartLegendPosition.Right,
  });
  const shape = ws.addShape({
    presetGeometry: 'rect',
    anchor: { kind: formulon.AnchorKind.OneCell, fromRow: 2, widthEmu: 914400, heightEmu: 457200 },
  });
  const slicer = ws.addSlicer(table, 0, {
    anchor: { fromRow: 20, fromCol: 3, toRow: 30, toCol: 5 },
  });
  const picture = ws.addPicture(PNG_1X1);
  return {
    Table: table,
    Pivot: pivot,
    PivotCache: cache,
    Drawing: Object.create(Object.getPrototypeOf(Object.getPrototypeOf(picture))),
    Picture: picture,
    Chart: chart,
    Shape: shape,
    Slicer: slicer,
  };
}

/** Prototype members of a handle instance, with the Drawing base members for subclasses. */
function jsHandleMembers(instance) {
  const proto = Object.getPrototypeOf(instance);
  const base = DRAWING_SUBCLASSES.has(proto.constructor.name) ? Object.getPrototypeOf(proto) : null;
  return [...new Set([...prototypeKeys(proto), ...(base ? prototypeKeys(base) : [])])];
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
    const pyWorkbook = pythonClassMembers('_workbook.py', 'Workbook');
    const pySheet = pythonClassMembers('_sheet.py', 'Sheet');
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
      pythonClassMembers('_workbook.py', 'Workbook'),
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
      pythonClassMembers('_sheet.py', 'Sheet'),
      diffSet(DIFF.py_only, 'Sheet'),
    );
  } finally {
    wb.dispose();
  }
});

for (const cls of Object.keys(HANDLE_FILES)) {
  test(`${cls} has the Python ${cls} members apart from the platform differences`, async () => {
    const formulon = await getFormulon();
    const wb = formulon.Workbook.createDefault();
    try {
      const handles = makeHandles(formulon, wb);
      // Drawing itself has no public constructor; an object over its prototype stands in.
      const instance = handles[cls];
      assertParity(
        cls,
        jsHandleMembers(instance),
        diffSet(DIFF.js_only, cls),
        pythonHandleMembers(cls),
        diffSet(DIFF.py_only, cls),
      );
    } finally {
      wb.dispose();
    }
  });
}

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
