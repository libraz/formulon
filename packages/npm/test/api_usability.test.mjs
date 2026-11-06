// Conformance of the unified public surface: A1 spellings of cell and range
// arguments, dense reads and writes, dirty tracking, partial updates,
// number-or-name references, the moved free functions and the missing-key rules.
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const here = path.dirname(fileURLToPath(import.meta.url));
const pkgRoot = path.resolve(here, '..');
const repoRoot = path.resolve(pkgRoot, '..', '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

let formulonPromise;
function getFormulon() {
  formulonPromise ??= loadPackage().then((mod) => mod.createFormulon());
  return formulonPromise;
}

function expectCode(fn, code) {
  assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, code, `code ${e.code} (${e.message})`);
    return true;
  });
}

async function withWorkbook(fn) {
  const formulon = await getFormulon();
  const wb = formulon.Workbook.createDefault();
  try {
    return await fn(wb, formulon);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

// A TypeError raised by the argument contract, not by a missing method.
const TYPE_ERROR = (e) => e instanceof TypeError && !/is not a function/.test(e.message);

const colName = (col) => {
  let s = '';
  for (let n = col + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
const a1 = (row, col) => `${colName(col)}${row + 1}`;
const rectA1 = (r) => `${a1(r.firstRow, r.firstCol)}:${a1(r.lastRow, r.lastCol)}`;
const rect = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });
const plain = (v) => JSON.parse(JSON.stringify(v));
const blank = (formulon) => ({ kind: formulon.ValueKind.Blank });

// ---- Criteria 1 and 2: A1 spellings of cell and range arguments ------------

const ADDRESS = JSON.parse(readFileSync(path.join(repoRoot, 'tools/api/address_args.json'), 'utf8')).js;

// Cell B3 is (2, 1); the second pair of two-cell methods is A1 (0, 0).
const CELL_ROW = 2;
const CELL_COL = 1;
const HYPERLINK = { target: 'https://example.com/' };

function pivotWorkbook(wb, formulon) {
  const ws = wb.sheet(0);
  ws.setText(0, 0, 'Region');
  ws.setText(0, 1, 'Amount');
  const cacheId = wb.addPivotCache();
  wb.setPivotCacheWorksheetSource(cacheId, { present: true, ref: 'A1:B3', sheet: 'Sheet1' });
  wb.addPivotCacheField(cacheId, 'Region');
  wb.addPivotCacheField(cacheId, 'Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const row = wb.getPivotCacheRecordCount(cacheId) + 1;
    ws.setText(row, 0, region);
    ws.setNumber(row, 1, amount);
    const rec = wb.addPivotCacheRecord(cacheId);
    wb.setPivotCacheRecordValue(cacheId, rec, 0, region);
    wb.setPivotCacheRecordValue(cacheId, rec, 1, amount);
  }
  const pivot = ws.addPivot({ name: 'Pivot1', cacheId, anchorRow: 0, anchorCol: 4 });
  ws.addPivotField(pivot, { sourceName: 'Region', axis: formulon.PivotAxis.Row });
  const amountField = ws.addPivotField(pivot, { sourceName: 'Amount', axis: formulon.PivotAxis.Data });
  ws.addPivotDataField(pivot, { name: 'Sum of Amount', fieldIndex: amountField, aggregation: 0 });
  return { cacheId, pivot };
}

// One entry per JSON method. `args` is the coordinate form; the A1 form is derived from the JSON
// positions. `setup` may return context for `args`.
const CELL_METHODS = {
  addHyperlink: { args: () => [CELL_ROW, CELL_COL, HYPERLINK] },
  clearCellXfIndex: { args: () => [CELL_ROW, CELL_COL] },
  evaluateCfFormula: { args: () => [CELL_ROW, CELL_COL, 0, 0, '=TRUE'] },
  evaluateFormula: { args: () => [CELL_ROW, CELL_COL, '=1+1'] },
  evaluateFormulaArray: { args: () => [CELL_ROW, CELL_COL, '=1+1'] },
  getArrayFormula: { args: () => [CELL_ROW, CELL_COL] },
  getCellPhonetic: { args: () => [CELL_ROW, CELL_COL] },
  getCellPhoneticProperties: { args: () => [CELL_ROW, CELL_COL] },
  getCellPhoneticRuns: { args: () => [CELL_ROW, CELL_COL] },
  getCellShowPhonetic: { args: () => [CELL_ROW, CELL_COL] },
  getCellXfIndex: { args: () => [CELL_ROW, CELL_COL] },
  getComment: { args: () => [CELL_ROW, CELL_COL] },
  getCommentVisible: { args: () => [CELL_ROW, CELL_COL] },
  getDataTable: { args: () => [CELL_ROW, CELL_COL] },
  getDependents: { args: () => [CELL_ROW, CELL_COL, 1] },
  getDisplayText: { args: () => [CELL_ROW, CELL_COL] },
  getEffectiveStyle: { args: () => [CELL_ROW, CELL_COL] },
  getFormula: { args: () => [CELL_ROW, CELL_COL] },
  getFormulaR1C1: { args: () => [CELL_ROW, CELL_COL] },
  getHyperlink: { args: () => [CELL_ROW, CELL_COL] },
  getLambdaText: { args: () => [CELL_ROW, CELL_COL] },
  getPrecedents: { args: () => [CELL_ROW, CELL_COL, 1] },
  getSpillInfo: { args: () => [CELL_ROW, CELL_COL] },
  getValue: { args: () => [CELL_ROW, CELL_COL] },
  removeComment: { args: () => [CELL_ROW, CELL_COL] },
  removeDataTable: { args: () => [CELL_ROW, CELL_COL] },
  removeHyperlink: { args: () => [CELL_ROW, CELL_COL] },
  setBlank: { args: () => [CELL_ROW, CELL_COL], pre: (ws) => ws.setNumber(CELL_ROW, CELL_COL, 7), mutates: true },
  setBool: { args: () => [CELL_ROW, CELL_COL, true], mutates: true },
  setCellPhonetic: { args: () => [CELL_ROW, CELL_COL, 'kana'], mutates: true },
  setCellPhoneticProperties: { args: () => [CELL_ROW, CELL_COL, { fontId: 0, type: 2, alignment: 1 }] },
  setCellPhoneticRuns: { args: () => [CELL_ROW, CELL_COL, [{ sb: 0, eb: 1, text: 'x' }]] },
  setCellShowPhonetic: { args: () => [CELL_ROW, CELL_COL, true] },
  setCellXfIndex: { args: () => [CELL_ROW, CELL_COL, 0] },
  setComment: { args: () => [CELL_ROW, CELL_COL, { author: 'A', text: 'note' }], mutates: true },
  setCommentVisible: { args: () => [CELL_ROW, CELL_COL, true] },
  setError: { args: (f) => [CELL_ROW, CELL_COL, f.ErrorCode.Div0], mutates: true },
  setFormula: { args: () => [CELL_ROW, CELL_COL, '=1+1'], mutates: true },
  setNumber: { args: () => [CELL_ROW, CELL_COL, 5], mutates: true },
  setText: { args: () => [CELL_ROW, CELL_COL, 'hello'], mutates: true },
  setValue: { args: () => [CELL_ROW, CELL_COL, 5], mutates: true },
  setValues: { args: () => [CELL_ROW, CELL_COL, [[5, 'x']]], mutates: true },
  // F2 (1, 5) is the first value cell of the report anchored at E1.
  showPivotDetails: {
    args: (_f, ctx) => [ctx.pivot, 1, 5],
    pivot: true,
  },
  validateValue: { args: (f) => [CELL_ROW, CELL_COL, { kind: f.ValueKind.Number, number: 1 }] },
  // Methods whose cell form moved elsewhere in the same surface revision.
  addPivot: { skip: 'takes a PivotInput record' },
  setPaneTopLeft: { skip: 'folded into setView' },
  setTopLeft: { skip: 'folded into setView' },
};

const READERS = [
  'getValue',
  'getFormula',
  'getComment',
  'getHyperlink',
  'getCommentVisible',
  'getCellXfIndex',
  'getCellPhonetic',
  'getCellShowPhonetic',
  'getCellPhoneticProperties',
  'getCellPhoneticRuns',
  'getDataTable',
  'getArrayFormula',
];

function attempt(fn) {
  try {
    return normalize(fn());
  } catch (e) {
    return { thrown: e.name, code: e.code };
  }
}

function normalize(result) {
  if (result instanceof Uint8Array) return { bytes: Array.from(result) };
  if (result !== null && typeof result === 'object' && typeof result.getValue === 'function') {
    return { sheet: result.index };
  }
  return plain(result ?? null);
}

function snapshot(ws) {
  const out = [];
  for (const [r, c] of [
    [0, 0],
    [CELL_ROW, CELL_COL],
  ]) {
    for (const m of READERS) out.push([m, r, c, attempt(() => ws[m](r, c))]);
  }
  out.push(attempt(() => ws.getMergesInRange(rect(0, 0, 9, 9))));
  out.push(attempt(() => ws.getPrintArea()));
  out.push(attempt(() => ws.getRangeStyleBases(rect(0, 0, 9, 9))));
  out.push(attempt(() => ws.listHyperlinks()));
  return out;
}

function seeded(wb, formulon, needsPivot) {
  const ws = wb.sheet(0);
  const ctx = {};
  if (needsPivot) Object.assign(ctx, pivotWorkbook(wb, formulon));
  else {
    ws.setNumber(0, 0, 1);
    ws.setNumber(0, 1, 2);
    ws.setFormula(1, 1, '=A1+B1');
  }
  return { ws, ctx };
}

function a1Args(coordArgs, pairs) {
  const out = [...coordArgs];
  const drop = new Set();
  for (const [i, j] of pairs) {
    out[i] = a1(coordArgs[i], coordArgs[j]);
    drop.add(j);
  }
  return out.filter((_, k) => !drop.has(k));
}

test('criterion 1: the JSON method set is non-empty and every method has an argument-table entry', () => {
  const names = Object.keys(ADDRESS.cell_args);
  assert.ok(names.length > 0, 'no cell_args in address_args.json');
  const missing = names.filter((n) => !(n in CELL_METHODS));
  assert.deepEqual(missing, [], `methods missing from the argument table: ${missing}`);
  for (const [name, pairs] of Object.entries(ADDRESS.cell_args)) {
    assert.ok(pairs.length > 0, name);
    const entry = CELL_METHODS[name];
    if (entry.skip) continue;
    const coord = entry.args({ ErrorCode: { Div0: 1 }, ValueKind: { Number: 1 } }, { pivot: 0 });
    for (const [i, j] of pairs) {
      assert.equal(typeof coord[i], 'number', `${name}: row at ${i}`);
      assert.equal(typeof coord[j], 'number', `${name}: col at ${j}`);
    }
  }
});

for (const [name, pairs] of Object.entries(ADDRESS.cell_args)) {
  const entry = CELL_METHODS[name];
  if (!entry || entry.skip) continue;

  test(`criterion 1: ${name} takes A1 in place of (row, col) and matches the coordinate form`, async () => {
    const formulon = await getFormulon();
    const probe = formulon.Workbook.createDefault();
    const exists = typeof probe.sheet(0)[name] === 'function';
    probe.dispose();
    assert.equal(exists, true, `Sheet.${name} is missing`);
    const run = (form) =>
      withWorkbook((wb, f) => {
        const { ws, ctx } = seeded(wb, f, entry.pivot === true);
        entry.pre?.(ws);
        const coord = entry.args(f, ctx);
        const args = form === 'a1' ? a1Args(coord, pairs) : coord;
        const before = snapshot(ws);
        const result = attempt(() => ws[name](...args));
        return { result, before, after: snapshot(ws) };
      });
    const coordRun = await run('coord');
    const a1Run = await run('a1');
    assert.deepEqual(a1Run.result, coordRun.result, `${name} result`);
    assert.deepEqual(a1Run.after, coordRun.after, `${name} resulting state`);
    if (entry.mutates) {
      assert.notDeepEqual(coordRun.after, coordRun.before, `${name} changed nothing`);
    }
  });

  if (pairs.length > 1) {
    test(`criterion 1: ${name} rejects a mix of A1 and coordinate pairs with TypeError`, async () => {
      await withWorkbook((wb, f) => {
        const { ws, ctx } = seeded(wb, f, entry.pivot === true);
        const coord = entry.args(f, ctx);
        const [i, j] = pairs[0];
        const mixed = [...coord];
        mixed[i] = a1(coord[i], coord[j]);
        mixed.splice(j, 1);
        assert.throws(() => ws[name](...mixed), TYPE_ERROR);
      });
    });
  }
}

test('criterion 1: a coordinate-form call with the pair at a non-string position still works for showPivotDetails', async () => {
  // Each call inserts a detail sheet, so every form runs on its own workbook.
  const details = (call) =>
    withWorkbook((wb, f) => {
      seeded(wb, f, true);
      const sheet = call(wb.sheet(0));
      return { index: sheet.index, values: plain(sheet.getValues('A1:B3')) };
    });
  const expected = await details((ws) => ws.showPivotDetails(0, 1, 5));
  assert.equal(expected.values[1][1].number, 10);
  assert.deepEqual(await details((ws) => ws.showPivotDetails(0, 'F2')), expected);
  assert.deepEqual(await details((ws) => ws.showPivotDetails('Pivot1', 1, 5)), expected);
  assert.deepEqual(await details((ws) => ws.showPivotDetails('Pivot1', 'F2')), expected);
});

test('edge: A1 spellings of one cell address the same cell', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setNumber(0, 0, 4);
    for (const spelling of ['A1', 'a1', '$A$1', 'A01']) {
      assert.deepEqual(plain(ws.getValue(spelling)), plain(ws.getValue(0, 0)), spelling);
      assert.deepEqual(plain(ws.getValue(spelling)), { kind: f.ValueKind.Number, number: 4 }, spelling);
    }
    ws.setNumber(1048575, 16383, 3);
    assert.deepEqual(plain(ws.getValue('XFD1048576')), { kind: f.ValueKind.Number, number: 3 });
  });
});

test('edge: an A1 argument that is not one unqualified cell is InvalidArgument', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    for (const bad of ['Sheet1!A1', 'A1:B2', '', 'R1C1', 'Ａ1', 'XFE1']) {
      expectCode(() => ws.getValue(bad), f.StatusCode.InvalidArgument);
      expectCode(() => ws.setNumber(bad, 1), f.StatusCode.InvalidArgument);
    }
  });
});

test('edge: setNumber, evaluateCfFormula and showPivotDetails take A1 like their coordinate forms', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setNumber('A1', 5);
    assert.deepEqual(plain(ws.getValue(0, 0)), { kind: f.ValueKind.Number, number: 5 });
    assert.deepEqual(
      plain(ws.evaluateCfFormula('B2', 'A1', '=TRUE')),
      plain(ws.evaluateCfFormula(1, 1, 0, 0, '=TRUE')),
    );
    assert.throws(() => ws.evaluateCfFormula('B2', 0, 0, '=TRUE'), TYPE_ERROR);
  });
});

test('edge: setFormula accepts text with or without the leading equals sign', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setFormula('A1', '1+1');
    ws.setFormula('A2', '=1+1');
    assert.equal(ws.getFormula(0, 0), '=1+1');
    assert.equal(ws.getFormula(1, 0), '=1+1');
  });
});

// ---- Criterion 2 -----------------------------------------------------------

const R = rect(1, 1, 2, 2);
const RECT_METHODS = {
  addHyperlinkRange: { args: (r) => [r, HYPERLINK], mutates: true },
  addMerge: { args: (r) => [r], mutates: true },
  evaluateCfRange: { args: (r) => [r, 45000] },
  getCellRectPt: { args: (r, f) => [r, f.GeometryMode.Display] },
  getCellsInRange: { args: (r) => [r] },
  getMergesInRange: { args: (r) => [r] },
  getRangeStyleBases: { args: (r) => [r] },
  getRangeStyleBasesExcluding: { args: (r) => [r, [], []] },
  getValues: { args: (r) => [r] },
  listCells: { args: (r) => [r], pre: (ws) => ws.setNumber(2, 2, 9) },
  recalcViewport: { args: (r) => [r] },
  remapRangeStyle: { args: (r) => [r, [0], [0]] },
  removeMerge: { args: (r) => [r], pre: (ws) => ws.addMerge(R) },
  setPivotAnchor: { args: (r, _f, ctx) => [ctx.pivot, r], pivot: true },
  setPrintArea: { args: (r) => [[r]], mutates: true },
  setRangeXfIndex: { args: (r) => [r, 0] },
  snapshotRangeStyle: { args: (r) => [r] },
};

test('criterion 2: the JSON rect set is non-empty and every method has an argument-table entry', () => {
  const names = Object.keys(ADDRESS.rect_args);
  assert.ok(names.length > 0, 'no rect_args in address_args.json');
  const missing = names.filter((n) => !(n in RECT_METHODS));
  assert.deepEqual(missing, [], `methods missing from the argument table: ${missing}`);
  for (const [name, positions] of Object.entries(ADDRESS.rect_args)) {
    const sentinel = { sentinel: true };
    const args = RECT_METHODS[name].args(sentinel, { GeometryMode: { Display: 0 } }, { pivot: 0 });
    for (const p of positions) {
      const slot = Array.isArray(args[p]) ? args[p][0] : args[p];
      assert.equal(slot, sentinel, `${name}: rect at position ${p}`);
    }
  }
});

for (const [name, positions] of Object.entries(ADDRESS.rect_args)) {
  const entry = RECT_METHODS[name];
  if (!entry) continue;
  void positions;

  test(`criterion 2: ${name} takes an A1 range in place of a CellRect`, async () => {
    const run = (asString) =>
      withWorkbook((wb, f) => {
        const { ws, ctx } = seeded(wb, f, entry.pivot === true);
        entry.pre?.(ws);
        const args = entry.args(asString ? rectA1(R) : R, f, ctx);
        const before = snapshot(ws);
        const result = attempt(() => ws[name](...args));
        const extra = name === 'setPivotAnchor' ? attempt(() => ws.getPivotLayout(ctx.pivot)) : null;
        return { result, before, after: [snapshot(ws), extra] };
      });
    const rectRun = await run(false);
    const strRun = await run(true);
    assert.deepEqual(strRun.result, rectRun.result, `${name} result`);
    assert.deepEqual(strRun.after, rectRun.after, `${name} resulting state`);
    if (entry.mutates) assert.notDeepEqual(rectRun.after[0], rectRun.before, `${name} changed nothing`);
  });
}

test('edge: addMerge accepts an A1 range and a single A1 cell as 1x1', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.addMerge('A1:B2');
    assert.deepEqual(plain(ws.getMergesInRange(rect(0, 0, 9, 9))), [rect(0, 0, 1, 1)]);
    ws.addMerge('D4');
    assert.deepEqual(plain(ws.getMergesInRange(rect(3, 3, 3, 3))), [rect(3, 3, 3, 3)]);
  });
});

// ---- Criterion 3: no Sheet.get / Sheet.set ----------------------------------

test('criterion 3: Sheet has no get or set', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    for (const name of ['get', 'set']) {
      assert.equal(name in ws, false, name);
      assert.equal(ws[name], undefined, name);
    }
  });
});

// ---- Criterion 4: getValues / setValues -------------------------------------

test('criterion 4: getValues of an empty A1 range is a dense grid of Blank', async () => {
  await withWorkbook((wb, f) => {
    const out = wb.sheet(0).getValues('A1:B2');
    assert.deepEqual(plain(out), [
      [blank(f), blank(f)],
      [blank(f), blank(f)],
    ]);
    assert.deepEqual(plain(wb.sheet(0).getValues(rect(0, 0, 1, 1))), plain(out));
  });
});

test('criterion 4: getValues reads values and a single A1 cell as 1x1', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setNumber(0, 0, 3);
    ws.setText(1, 1, 'x');
    assert.deepEqual(plain(ws.getValues('A1:B2')), [
      [{ kind: f.ValueKind.Number, number: 3 }, blank(f)],
      [blank(f), { kind: f.ValueKind.Text, text: 'x' }],
    ]);
    assert.deepEqual(plain(ws.getValues('A1')), [[{ kind: f.ValueKind.Number, number: 3 }]]);
  });
});

test('criterion 4: getValues places every cell of a range larger than one page', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const rows = 700;
    const cols = 100;
    const cells = [];
    for (let r = 0; r < rows; r += 1) {
      for (let c = 0; c < cols; c += 1) {
        if ((r * 7 + c * 13) % 5 === 0) cells.push([r, c]);
      }
    }
    for (const [r, c] of cells) ws.setNumber(r, c, r * 1000 + c);
    const out = ws.getValues(rect(0, 0, rows - 1, cols - 1));
    assert.equal(rows * cols > 65536, true);
    assert.equal(out.length, rows);
    assert.ok(out.every((row) => row.length === cols));
    const filled = new Set(cells.map(([r, c]) => r * cols + c));
    for (let r = 0; r < rows; r += 1) {
      for (let c = 0; c < cols; c += 1) {
        const v = out[r][c];
        if (filled.has(r * cols + c)) {
          assert.equal(v.kind, f.ValueKind.Number, `${r},${c}`);
          assert.equal(v.number, r * 1000 + c, `${r},${c}`);
        } else {
          assert.equal(v.kind, f.ValueKind.Blank, `${r},${c}`);
        }
      }
    }
  });
});

test('criterion 4: getValues includes spilled values', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setFormula(0, 0, '=SEQUENCE(3)');
    wb.recalc();
    const out = ws.getValues('A1:A3').map((row) => row[0].number);
    assert.deepEqual(out, [1, 2, 3]);
    assert.equal(ws.getValues('A3')[0][0].kind, f.ValueKind.Number);
  });
});

test('criterion 4: getValues bounds the cell count at 2^22', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    expectCode(() => ws.getValues('A:XFD'), f.StatusCode.InvalidArgument);
    const out = ws.getValues('A1:D1048576');
    assert.equal(out.length, 1048576);
    assert.equal(out[0].length, 4);
  });
});

test('criterion 4: setValues writes a dense grid by the setValue rules', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setNumber(1, 1, 99);
    ws.setValues('A1', [
      [1, 'x'],
      [true, null],
    ]);
    assert.deepEqual(plain(ws.getValues('A1:B2')), [
      [
        { kind: f.ValueKind.Number, number: 1 },
        { kind: f.ValueKind.Text, text: 'x' },
      ],
      [{ kind: f.ValueKind.Bool, boolean: true }, blank(f)],
    ]);
    ws.setValues(4, 1, [[{ kind: f.ValueKind.Number, number: 8 }]]);
    assert.equal(ws.getValue(4, 1).number, 8);
  });
});

test('criterion 4: setValues writes nothing when validation fails', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const array = { kind: f.ValueKind.Array, rows: 1, cols: 1, values: [{ kind: f.ValueKind.Number, number: 1 }] };
    const cases = [
      ['ragged rows', 'A1', [[1], [1, 2]]],
      ['NaN', 'A1', [[Number.NaN]]],
      ['Infinity', 'A1', [[1, Number.POSITIVE_INFINITY]]],
      // biome-ignore lint/suspicious/noSparseArray: a hole is the case under test
      ['a hole', 'A1', [[1, , 3]]],
      ['a non-scalar Value', 'A1', [[array]]],
      ['past the last column', 'XFD1', [[1, 2]]],
    ];
    for (const [label, at, grid] of cases) {
      assert.throws(
        () => ws.setValues(at, grid),
        (e) => e.name === 'FormulonError' || TYPE_ERROR(e),
        label,
      );
      assert.deepEqual(
        plain(ws.getValues('A1:C2')),
        Array.from({ length: 2 }, () => [blank(f), blank(f), blank(f)]),
        label,
      );
      assert.deepEqual(plain(ws.getValue(0, 16383)), blank(f), label);
    }
  });
});

test('criterion 4: setValues of an empty grid or empty rows succeeds and does nothing', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('A1', []);
    ws.setValues('A1', [[]]);
    assert.deepEqual(plain(ws.getValue(0, 0)), blank(f));
  });
});

test('criterion 4: setValues writes a leading equals sign as text', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setValues('A1', [['=1+1']]);
    assert.deepEqual(plain(ws.getValue(0, 0)), { kind: f.ValueKind.Text, text: '=1+1' });
    assert.equal(ws.getFormula(0, 0), null);
  });
});

test('criterion 4: setValues over a data-table body throws and keeps the cells written before it', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    // B2:B3 is a one-variable data table over the column input A1; B2 refuses writes.
    ws.setNumber(0, 0, 1);
    ws.setNumber(1, 0, 2);
    ws.setNumber(2, 0, 3);
    ws.addDataTable({ rect: rect(1, 1, 2, 1), colInput: { row: 0, col: 0 } });
    assert.notEqual(ws.getDataTable(1, 1), null);
    assert.throws(
      () => ws.setValues('A2', [[7, 8]]),
      (e) => e.name === 'FormulonError',
    );
    assert.equal(ws.getValue(1, 0).number, 7);
    assert.notEqual(ws.getDataTable(1, 1), null);
  });
});

// ---- Criterion 5: hasDirtyCells ----------------------------------------------

test('criterion 5: hasDirtyCells is false for a new workbook, a value-only load and after recalc', async () => {
  await withWorkbook((wb, f) => {
    assert.equal(wb.hasDirtyCells(), false);
    wb.sheet(0).setNumber(0, 0, 1);
    wb.sheet(0).setFormula(0, 1, '=A1+1');
    assert.equal(wb.hasDirtyCells(), true);
    wb.recalc();
    assert.equal(wb.hasDirtyCells(), false);
    // A loaded formula is dirty until the first full recalc.
    const reloaded = f.Workbook.load(wb.save());
    try {
      assert.equal(reloaded.hasDirtyCells(), true);
      reloaded.recalc();
      assert.equal(reloaded.hasDirtyCells(), false);
    } finally {
      reloaded.dispose();
    }
  });
  await withWorkbook((wb, f) => {
    wb.sheet(0).setNumber(0, 0, 1);
    wb.sheet(0).setText(1, 0, 'x');
    const reloaded = f.Workbook.load(wb.save());
    try {
      assert.equal(reloaded.hasDirtyCells(), false);
    } finally {
      reloaded.dispose();
    }
  });
});

test('criterion 5: edits mark the workbook dirty', async () => {
  const edits = {
    setNumber: (wb) => wb.sheet(0).setNumber(0, 0, 2),
    setDefinedName: (wb) => wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$A$1' }),
    insertRows: (wb) => wb.sheet(0).insertRows(0, 1),
    removeSheet: (wb) => wb.removeSheet(1),
  };
  for (const [name, edit] of Object.entries(edits)) {
    await withWorkbook((wb) => {
      wb.addSheet('Second');
      wb.sheet(0).setNumber(0, 0, 1);
      wb.sheet(0).setFormula(0, 1, '=Rate+A1');
      wb.sheet(0).setFormula(1, 1, '=Second!A1');
      wb.recalc();
      assert.equal(wb.hasDirtyCells(), false, `${name}: before`);
      edit(wb);
      assert.equal(wb.hasDirtyCells(), true, name);
    });
  }
});

test('criterion 5: changing the profile marks a formula workbook dirty, the same profile does not', async () => {
  await withWorkbook((wb, f) => {
    wb.sheet(0).setFormula(0, 0, '=1+1');
    wb.recalc();
    const current = wb.getExcelProfileId();
    wb.setExcelProfileId(current);
    assert.equal(wb.hasDirtyCells(), false);
    wb.setExcelProfileId(current === 'mac-365-de_DE' ? 'win-365-en_US' : 'mac-365-de_DE');
    assert.equal(wb.hasDirtyCells(), true);
    void f;
  });
});

test('criterion 5: a volatile-only workbook is clean after recalc', async () => {
  await withWorkbook((wb) => {
    wb.sheet(0).setFormula(0, 0, '=NOW()');
    wb.recalc();
    assert.equal(wb.hasDirtyCells(), false);
  });
});

test('criterion 5: recalcViewport alone leaves the workbook dirty', async () => {
  await withWorkbook((wb) => {
    wb.sheet(0).setNumber(0, 0, 1);
    wb.sheet(0).setFormula(0, 1, '=A1+1');
    // Z100 depends on A1 but lies outside the viewport, so it stays dirty.
    wb.sheet(0).setFormula(99, 25, '=A1*2');
    assert.equal(wb.hasDirtyCells(), true);
    wb.sheet(0).recalcViewport(rect(0, 0, 9, 9));
    assert.equal(wb.hasDirtyCells(), true);
  });
});

test('criterion 5: hasDirtyCells inside an iteration progress callback is a reentrancy error', async () => {
  const f = await getFormulon();
  const wb = f.Workbook.createDefault();
  try {
    wb.setIterative({ enabled: true, maxIterations: 100, maxChange: 0.0001 });
    wb.sheet(0).setFormula(0, 0, '=(A1+10)/2');
    let raised = null;
    wb.setIterativeProgress(() => {
      try {
        wb.hasDirtyCells();
      } catch (e) {
        raised = e;
      }
      return true;
    });
    wb.recalc();
    assert.ok(raised, 'hasDirtyCells did not throw inside the callback');
    assert.equal(raised.name, 'FormulonError');
    assert.equal(raised.code, 4005);
  } finally {
    wb.dispose();
  }
});

// ---- Criterion 6: enumeration names ----------------------------------------

test('criterion 6: no public name keeps a get...Count or get...At form except the exempt record count', async () => {
  await withWorkbook(async (wb, f) => {
    const names = new Set();
    for (const proto of [Object.getPrototypeOf(wb), Object.getPrototypeOf(wb.sheet(0))]) {
      for (const n of Object.getOwnPropertyNames(proto)) names.add(n);
    }
    for (const n of Object.keys(f)) names.add(n);
    const offenders = [...names].filter((n) => /^get.*(Count|At)$/.test(n) && n !== 'getPivotCacheRecordCount');
    assert.deepEqual(offenders, []);
    assert.equal(typeof wb.getPivotCacheRecordCount, 'function');
  });
});

test('criterion 6: list methods replace the counted enumerations', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    for (const name of ['listFonts', 'listFills', 'listBorders', 'listCellXfs', 'listCellStyleXfs']) {
      const list = wb[name]();
      assert.ok(Array.isArray(list) && list.length > 0, name);
      assert.deepEqual(plain(wb[name.replace(/^list(.*)s$/, 'get$1')](0)), plain(list[0]), name);
    }
    // A new workbook has no differential formats.
    assert.deepEqual(plain(wb.listDxfs()), []);
    assert.deepEqual(plain(ws.listCells()), []);
    ws.setNumber(2, 1, 5);
    ws.setText(0, 0, 'a');
    const cells = ws.listCells();
    assert.deepEqual(cells.map((c) => [c.row, c.col]).sort(), [
      [0, 0],
      [2, 1],
    ]);
    assert.deepEqual(
      ws.listCells('A1:A1').map((c) => [c.row, c.col]),
      [[0, 0]],
    );
    assert.deepEqual(ws.listPivots(), []);
  });
});

test('criterion 6: listPivots, listPivotFieldItems and the pivot cache lists describe their elements', async () => {
  await withWorkbook((wb, f) => {
    const { cacheId, pivot } = pivotWorkbook(wb, f);
    const ws = wb.sheet(0);
    assert.deepEqual(plain(ws.listPivots()), [{ name: 'Pivot1', cacheId }]);
    assert.deepEqual(
      wb.listPivotCacheFields(cacheId).map((x) => x.name),
      ['Region', 'Amount'],
    );
    wb.addPivotCacheFieldSharedItem(cacheId, 0, 'East');
    wb.addPivotCacheFieldSharedItem(cacheId, 0, 45292, { date: true });
    const shared = wb.listPivotCacheFieldSharedItems(cacheId, 0);
    assert.equal(shared.length, 2);
    assert.deepEqual(plain(shared[0]), { value: { kind: f.ValueKind.Text, text: 'East' }, isDate: false });
    assert.equal(shared[1].isDate, true);
    assert.deepEqual(plain(ws.listPivotFieldItems(pivot, 0)), []);
    expectCode(() => ws.getPivotFieldItem(pivot, 0, 0), f.StatusCode.IndexOutOfRange);
    ws.addPivotFieldItem(pivot, 0, 'East');
    const items = ws.listPivotFieldItems(pivot, 0);
    assert.equal(items.length, 1);
    assert.deepEqual(plain(ws.getPivotFieldItem(pivot, 0, 0)), plain(items[0]));
  });
});

// ---- Criterion 7: partial updates ---------------------------------------------

test('criterion 7: setView changes only the keys given', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const before = plain(ws.getView());
    ws.setView({});
    assert.deepEqual(plain(ws.getView()), before);
    ws.setView({ showGridLines: false });
    assert.deepEqual(plain(ws.getView()), { ...before, showGridLines: false });
    assert.throws(() => ws.setView({ zoomScale: null }), TYPE_ERROR);
    assert.deepEqual(plain(ws.getView()), { ...before, showGridLines: false });
    void f;
  });
});

test('criterion 7: setView merges freeze rows and columns from the current view', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setView({ freezeRows: 0, freezeCols: 1 });
    ws.setView({ freezeRows: 2 });
    const view = ws.getView();
    assert.equal(view.freezeRows, 2);
    assert.equal(view.freezeCols, 1);
  });
});

test('criterion 7: setView rejects a freeze together with a split', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const before = plain(ws.getView());
    assert.throws(() => ws.setView({ freezeRows: 1, splitX: 100 }), TYPE_ERROR);
    assert.deepEqual(plain(ws.getView()), before);
  });
});

test('criterion 7: setView checks paneTopLeft against the merged state', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setView({ freezeRows: 1, paneTopLeftRow: 3 });
    const view = ws.getView();
    assert.equal(view.freezeRows, 1);
    assert.equal(view.paneTopLeftRow, 3);
  });
});

test('criterion 7: setView refuses to hide the only visible sheet and changes nothing', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const before = plain(ws.getView());
    assert.throws(
      () => ws.setView({ visibility: f.SheetVisibility.Hidden, showGridLines: false }),
      (e) => e.name === 'FormulonError',
    );
    assert.deepEqual(plain(ws.getView()), before);
  });
});

test('criterion 7: setView replaces the per-field view setters', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const gone = [
      'setZoom',
      'setFreeze',
      'setVisibility',
      'setShowGridLines',
      'setShowRowColHeaders',
      'setShowZeros',
      'setShowFormulas',
      'setShowOutlineSymbols',
      'setRightToLeft',
      'setTabSelected',
      'setTopLeft',
      'setPaneTopLeft',
      'setSplit',
      'setViewMode',
    ];
    for (const name of gone) assert.equal(name in ws, false, name);
    ws.setView({ zoomScale: 150, viewMode: 2, rightToLeft: true });
    const view = ws.getView();
    assert.equal(view.zoomScale, 150);
    assert.equal(view.viewMode, 2);
    assert.equal(view.rightToLeft, true);
  });
});

function chartWorkbook(wb, f) {
  const ws = wb.sheet(0);
  ['S1', 'S2'].forEach((n, c) => {
    ws.setText(0, c + 1, n);
  });
  ['a', 'b'].forEach((cat, r) => {
    ws.setText(r + 1, 0, cat);
    ws.setNumber(r + 1, 1, r + 1);
    ws.setNumber(r + 1, 2, r + 3);
  });
  const id = ws.addChart({
    kind: f.ChartInsertKind.Column,
    source: 'Sheet1!$A$1:$C$3',
    legend: f.ChartLegendPosition.Right,
  });
  return { ws, id };
}

test('criterion 7: setChartLegend keeps the keys it omits', async () => {
  await withWorkbook((wb, f) => {
    const { ws, id } = chartWorkbook(wb, f);
    ws.setChartLegend(id, { position: f.ChartLegendPosition.Bottom, overlay: true, hiddenSeries: [1] });
    ws.setChartLegend(id, { position: f.ChartLegendPosition.Right });
    const legend = plain(ws.getChart(id).legend);
    assert.deepEqual(legend, { position: f.ChartLegendPosition.Right, overlay: true, hiddenSeries: [1] });
  });
});

test('criterion 7: setChartLegend with null removes the legend; a patch on no legend creates one', async () => {
  await withWorkbook((wb, f) => {
    const { ws, id } = chartWorkbook(wb, f);
    ws.setChartLegend(id, null);
    assert.equal(ws.getChart(id).legend, null);
    ws.setChartLegend(id, { overlay: true });
    assert.deepEqual(plain(ws.getChart(id).legend), {
      position: f.ChartLegendPosition.Right,
      overlay: true,
      hiddenSeries: [],
    });
    ws.setChartLegend(id, null);
    ws.setChartLegend(id, {});
    assert.deepEqual(plain(ws.getChart(id).legend), {
      position: f.ChartLegendPosition.Right,
      overlay: false,
      hiddenSeries: [],
    });
  });
});

test('criterion 7: setChartLegend rejects null for keys that cannot be cleared', async () => {
  await withWorkbook((wb, f) => {
    const { ws, id } = chartWorkbook(wb, f);
    assert.throws(() => ws.setChartLegend(id, { position: null }), TYPE_ERROR);
    assert.throws(() => ws.setChartLegend(id, { hiddenSeries: null }), TYPE_ERROR);
  });
});

test('criterion 7: setHeaderFooter clears a text key given as null or the empty string', async () => {
  for (const cleared of [null, '']) {
    await withWorkbook((wb) => {
      const ws = wb.sheet(0);
      ws.setHeaderFooter({ oddHeader: '&CTitle', oddFooter: '&RPage' });
      ws.setHeaderFooter({ oddHeader: cleared });
      const xml = ws.getHeaderFooterXml() ?? '';
      assert.doesNotMatch(xml, /<oddHeader>/, String(cleared));
      assert.match(xml, /<oddFooter>&amp;RPage<\/oddFooter>/);
    });
  }
});

test('criterion 7: setPrintOptions rejects null', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    assert.throws(() => ws.setPrintOptions({ gridLines: null }), TYPE_ERROR);
    ws.setPrintOptions({ gridLines: true });
    ws.setPrintOptions({ horizontalCentered: true });
    const xml = ws.getPrintOptionsXml() ?? '';
    assert.match(xml, /gridLines="(1|true)"/);
    assert.match(xml, /horizontalCentered="(1|true)"/);
  });
});

test('criterion 7: updatePivotField clears customName with null and rejects keys outside the patch', async () => {
  await withWorkbook((wb, f) => {
    const { pivot } = pivotWorkbook(wb, f);
    const ws = wb.sheet(0);
    ws.updatePivotField(pivot, 0, { customName: 'Area' });
    assert.equal(ws.listPivotFields(pivot)[0].customName, 'Area');
    ws.updatePivotField(pivot, 0, { showAll: true });
    assert.equal(ws.listPivotFields(pivot)[0].customName, 'Area');
    ws.updatePivotField(pivot, 0, { customName: null });
    assert.equal(ws.listPivotFields(pivot)[0].customName, null);
    assert.throws(() => ws.updatePivotField(pivot, 0, { axis: f.PivotAxis.Row }), TYPE_ERROR);
  });
});

test('criterion 7: updatePivotItem changes the keys given and the per-field setters are gone', async () => {
  await withWorkbook((wb, f) => {
    const { cacheId, pivot } = pivotWorkbook(wb, f);
    const ws = wb.sheet(0);
    wb.addPivotCacheFieldSharedItem(cacheId, 0, 'East');
    ws.addPivotFieldItem(pivot, 0, 'East');
    ws.updatePivotItem(pivot, 0, 0, { caption: 'Eastern', visible: false });
    const item = ws.getPivotFieldItem(pivot, 0, 0);
    assert.equal(item.caption, 'Eastern');
    assert.equal(item.visible, false);
    ws.updatePivotItem(pivot, 0, 0, { caption: null });
    assert.equal(ws.getPivotFieldItem(pivot, 0, 0).caption, null);
    for (const name of [
      'setPivotFieldCustomName',
      'setPivotFieldNumberFormat',
      'setPivotFieldDefaultSubtotal',
      'setPivotFieldInsertBlankRow',
      'setPivotFieldFillDownLabels',
      'setPivotFieldShowAll',
      'setPivotFieldSubtotalTop',
      'setPivotFieldItemVisible',
      'setPivotFieldItemShowDetail',
      'setPivotFieldItemCaption',
    ]) {
      assert.equal(name in ws, false, name);
    }
  });
});

// ---- Criterion 8: number-or-name references ------------------------------------

function tableWorkbook(wb, name, displayName) {
  const ws = wb.sheet(0);
  ws.setText(0, 0, 'A');
  ws.setText(0, 1, 'B');
  return wb.addTable({
    range: { sheet: 0, rect: rect(0, 0, 2, 1) },
    name,
    displayName,
    columns: ['A', 'B'],
  });
}

test('criterion 8: removeTable resolves a displayName, even one that reads as a number, and a number', async () => {
  await withWorkbook((wb, f) => {
    tableWorkbook(wb, 'Sales', 'Sales');
    expectCode(() => wb.removeTable('Missing'), f.StatusCode.NotFound);
    expectCode(() => wb.removeTable(99), f.StatusCode.IndexOutOfRange);
    assert.throws(() => wb.removeTable(true), TYPE_ERROR);
    assert.equal(wb.listTables().length, 1);
    wb.removeTable('Sales');
    assert.equal(wb.listTables().length, 0);
  });
  await withWorkbook((wb) => {
    wb.sheet(0).setText(0, 0, 'A');
    wb.addTable({ range: { sheet: 0, rect: rect(0, 0, 1, 0) }, name: 'T0', displayName: '3', columns: ['A'] });
    wb.removeTable('3');
    assert.equal(wb.listTables().length, 0);
  });
  await withWorkbook((wb) => {
    tableWorkbook(wb, 'Sales', 'Sales');
    wb.removeTable(0);
    assert.equal(wb.listTables().length, 0);
  });
});

test('criterion 8: a table is found by displayName when name differs', async () => {
  await withWorkbook((wb) => {
    tableWorkbook(wb, 'InternalName', 'Shown');
    wb.removeTable('Shown');
    assert.equal(wb.listTables().length, 0);
  });
});

// Caller-supplied rule ids are GUIDs, as the engine checks their shape.
const CF_ID_FIRST = '{00000000-0000-4000-8000-000000000001}';
const CF_ID_SECOND = '{00000000-0000-4000-8000-000000000002}';
const CF_ID_ROUND_TRIP = '{00000000-0000-4000-8000-000000000003}';
const CF_ID_ROUND_TRIP_2 = '{00000000-0000-4000-8000-000000000004}';

function cfRule(id) {
  return { ranges: [rect(0, 0, 3, 0)], type: 0, formula1: 'TRUE', id };
}

test('criterion 8: removeCfRule takes an id; adding a duplicate id is InvalidArgument', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.addCfRule(cfRule(CF_ID_FIRST));
    ws.addCfRule(cfRule(CF_ID_SECOND));
    expectCode(() => ws.addCfRule(cfRule(CF_ID_FIRST)), f.StatusCode.InvalidArgument);
    assert.equal(ws.listCfRules().length, 2);
    ws.removeCfRule(CF_ID_SECOND);
    assert.deepEqual(
      ws.listCfRules().map((r) => r.id),
      [CF_ID_FIRST],
    );
    expectCode(() => ws.removeCfRule(CF_ID_SECOND), f.StatusCode.NotFound);
    ws.removeCfRule(0);
    assert.equal(ws.listCfRules().length, 0);
  });
});

test('criterion 8: removePivot resolves a name and takes the first of two equal names', async () => {
  await withWorkbook((wb, f) => {
    const { cacheId } = pivotWorkbook(wb, f);
    const ws = wb.sheet(0);
    const second = ws.addPivot({ name: 'Pivot1', cacheId, anchorRow: 0, anchorCol: 10 });
    assert.equal(second, 1);
    ws.removePivot('Pivot1');
    const left = ws.listPivots();
    assert.equal(left.length, 1);
    assert.equal(ws.getPivotLayout(0).rect.firstCol, 10);
    expectCode(() => ws.removePivot('Nope'), f.StatusCode.NotFound);
  });
});

test('criterion 8: removeProtectedRange resolves a name', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.addProtectedRange({ name: 'Inputs', ranges: [rect(0, 0, 0, 0)] });
    ws.addProtectedRange({ name: 'Totals', ranges: [rect(1, 0, 1, 0)] });
    ws.removeProtectedRange('Inputs');
    assert.deepEqual(
      ws.listProtectedRanges().map((r) => r.name),
      ['Totals'],
    );
    expectCode(() => ws.removeProtectedRange('Inputs'), f.StatusCode.NotFound);
  });
});

// ---- Criterion 9: the free functions live on Formulon --------------------------

const MOVED = [
  'getFunctionMetadata',
  'listFunctionNames',
  'localizeFunctionName',
  'canonicalizeFunctionName',
  'localizeFormula',
  'canonicalizeFormula',
  'getLocaleFacts',
  'listLocaleErrorNames',
  'compareText',
  'getCollationClasses',
  'getImageInfo',
];

test('criterion 9: the stateless functions are Formulon members and not Workbook members', async () => {
  await withWorkbook((wb, f) => {
    for (const name of MOVED) {
      assert.equal(typeof f[name], 'function', `Formulon.${name}`);
      assert.equal(name in wb, false, `Workbook.${name}`);
    }
    assert.equal(f.localizeFormula('=SUM(1.5,2)', 'mac-365-de_DE'), '=SUMME(1,5;2)');
    assert.equal(f.canonicalizeFormula('=SUMME(1,5;2)', 'mac-365-de_DE'), '=SUM(1.5,2)');
    assert.equal(f.localizeFunctionName('SUM', 'mac-365-de_DE'), 'SUMME');
    assert.equal(f.canonicalizeFunctionName('SUMME', 'win-365-de_DE'), 'SUM');
    assert.ok(f.listFunctionNames().includes('SUM'));
    assert.equal(f.getLocaleFacts('win-365-en_US').decimalSeparator, '.');
    assert.ok(Array.isArray(f.listLocaleErrorNames('win-365-en_US')));
    assert.equal(typeof f.compareText('a', 'b', 0, 'win-365-en_US'), 'number');
    assert.equal(typeof f.getFunctionMetadata, 'function');
    assert.equal(typeof f.getCollationClasses, 'function');
    assert.equal(typeof f.getImageInfo, 'function');
  });
});

test('criterion 9: Formulon exposes no stateless function under the retired error-name helper', async () => {
  const f = await getFormulon();
  assert.equal('getErrorDisplayName' in f, false);
});

// ---- Criterion 10: pivot cache value writers ---------------------------------------

test('criterion 10: two dispatching writers replace the twelve typed ones', async () => {
  await withWorkbook((wb) => {
    for (const kind of ['Text', 'Number', 'Date', 'Bool', 'Blank', 'Error']) {
      assert.equal(`addPivotCacheFieldSharedItem${kind}` in wb, false, kind);
      assert.equal(`setPivotCacheRecord${kind}` in wb, false, kind);
    }
    assert.equal(typeof wb.addPivotCacheFieldSharedItem, 'function');
    assert.equal(typeof wb.setPivotCacheRecordValue, 'function');
  });
});

test('criterion 10: addPivotCacheFieldSharedItem dispatches on the value type and date needs a number', async () => {
  await withWorkbook((wb, f) => {
    const cacheId = wb.addPivotCache();
    wb.addPivotCacheField(cacheId, 'Mixed');
    wb.addPivotCacheFieldSharedItem(cacheId, 0, 5, { date: true });
    wb.addPivotCacheFieldSharedItem(cacheId, 0, 5);
    wb.addPivotCacheFieldSharedItem(cacheId, 0, 'x');
    wb.addPivotCacheFieldSharedItem(cacheId, 0, true);
    wb.addPivotCacheFieldSharedItem(cacheId, 0, null);
    wb.addPivotCacheFieldSharedItem(cacheId, 0, { kind: f.ValueKind.Error, errorCode: f.ErrorCode.NA });
    const items = wb.listPivotCacheFieldSharedItems(cacheId, 0);
    assert.deepEqual(
      items.map((i) => [i.value.kind, i.isDate]),
      [
        [f.ValueKind.Number, true],
        [f.ValueKind.Number, false],
        [f.ValueKind.Text, false],
        [f.ValueKind.Bool, false],
        [f.ValueKind.Blank, false],
        [f.ValueKind.Error, false],
      ],
    );
    assert.throws(() => wb.addPivotCacheFieldSharedItem(cacheId, 0, 'x', { date: true }), TYPE_ERROR);
    assert.equal(wb.listPivotCacheFieldSharedItems(cacheId, 0).length, 6);
  });
});

test('criterion 10: setPivotCacheRecordValue dispatches on the value type', async () => {
  await withWorkbook((wb, f) => {
    const cacheId = wb.addPivotCache();
    wb.addPivotCacheField(cacheId, 'A');
    const rec = wb.addPivotCacheRecord(cacheId);
    wb.setPivotCacheRecordValue(cacheId, rec, 0, 'West');
    wb.setPivotCacheRecordValue(cacheId, rec, 0, 7);
    wb.setPivotCacheRecordValue(cacheId, rec, 0, 45300, { date: true });
    assert.throws(() => wb.setPivotCacheRecordValue(cacheId, rec, 0, true, { date: true }), TYPE_ERROR);
    assert.equal(wb.getPivotCacheRecordCount(cacheId), 1);
    void f;
  });
});

// ---- Criterion 12: missing keys ----------------------------------------------------

test('criterion 12: removing or reading a missing drawing object, chart or merge is NotFound', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const missing = 987654;
    expectCode(() => ws.removeDrawingObject(missing), f.StatusCode.NotFound);
    expectCode(() => ws.getChartData(missing), f.StatusCode.NotFound);
    expectCode(() => ws.removeMerge(rect(0, 0, 1, 1)), f.StatusCode.NotFound);
    expectCode(() => ws.removeDataTable(0, 0), f.StatusCode.NotFound);
    assert.deepEqual(ws.getMergesInRange(rect(0, 0, 9, 9)), []);
  });
});

test('criterion 12: getPrintArea is null while unset', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    assert.equal(ws.getPrintArea(), null);
    ws.setPrintArea([rect(0, 0, 4, 4)]);
    assert.deepEqual(plain(ws.getPrintArea()), [rect(0, 0, 4, 4)]);
    ws.setPrintArea(null);
    assert.equal(ws.getPrintArea(), null);
  });
});

// ---- Criterion 13: sibling alignment --------------------------------------------------

test('criterion 13: column geometry takes mode last and defaults to Display', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.equal(ws.convertColumnCharsToPt(10), ws.convertColumnCharsToPt(10, f.GeometryMode.Display));
    assert.equal(ws.convertColumnPtToChars(50), ws.convertColumnPtToChars(50, f.GeometryMode.Display));
    assert.equal(ws.getColumnWidthPt(0), ws.getColumnWidthPt(0, f.GeometryMode.Display));
    assert.notEqual(ws.convertColumnCharsToPt(10, f.GeometryMode.Print), undefined);
    assert.throws(() => ws.convertColumnCharsToPt(10, 7), RangeError);
    assert.throws(() => ws.convertColumnPtToChars(10, 7), RangeError);
    assert.throws(() => ws.getColumnWidthPt(0, 7), RangeError);
  });
});

test('criterion 13: removeMergeAt and removeHyperlinkAt are gone', async () => {
  await withWorkbook((wb) => {
    assert.equal('removeMergeAt' in wb.sheet(0), false);
    assert.equal('removeHyperlinkAt' in wb.sheet(0), false);
  });
});

test('criterion 13: defined names take an optional scope', async () => {
  await withWorkbook((wb, f) => {
    wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$A$1' });
    wb.setDefinedName({ name: 'Rate', formula: 'Sheet1!$B$1', localSheetId: 0 });
    assert.equal(wb.getDefinedName('Rate').localSheetId, null);
    assert.equal(wb.getDefinedName('Rate', null).localSheetId, null);
    assert.equal(wb.getDefinedName('Rate', 0).localSheetId, 0);
    assert.equal(wb.getDefinedName('Rate', 'Sheet1').localSheetId, 0);
    assert.equal(wb.getDefinedName('Nope'), null);
    wb.removeDefinedName('Rate', 'Sheet1');
    assert.equal(wb.getDefinedName('Rate', 0), null);
    wb.removeDefinedName('Rate');
    assert.equal(wb.getDefinedName('Rate'), null);
    expectCode(() => wb.removeDefinedName('Rate'), f.StatusCode.NotFound);
  });
});

test('criterion 13: theme script fonts are read by a key record and written by a record', async () => {
  await withWorkbook((wb) => {
    wb.setThemeScriptFont({ major: false, script: 'Zzzz', typeface: 'Test Face' });
    assert.equal(wb.getThemeScriptFont({ major: false, script: 'Zzzz' }).typeface, 'Test Face');
    assert.equal(wb.getThemeScriptFont({ major: true, script: 'Zzzz' }), null);
  });
});

test('criterion 13: setChartType defaults the grouping to Clustered', async () => {
  await withWorkbook((wb, f) => {
    const { ws, id } = chartWorkbook(wb, f);
    ws.setChartType(id, 0, f.ChartInsertKind.Line);
    const explicit = plain(ws.getChart(id));
    ws.setChartType(id, 0, f.ChartInsertKind.Column, f.ChartGrouping.Stacked);
    ws.setChartType(id, 0, f.ChartInsertKind.Line, f.ChartGrouping.Clustered);
    assert.deepEqual(plain(ws.getChart(id)), explicit);
  });
});

test('criterion 13: setPivotFieldSort and addPivotFieldItem take an options object', async () => {
  await withWorkbook((wb, f) => {
    const { cacheId, pivot } = pivotWorkbook(wb, f);
    const ws = wb.sheet(0);
    wb.addPivotCacheFieldSharedItem(cacheId, 0, 'East');
    ws.addPivotFieldItem(pivot, 0, 'East', { visible: false });
    assert.equal(ws.getPivotFieldItem(pivot, 0, 0).visible, false);
    ws.addPivotFieldItem(pivot, 0, 'West');
    assert.equal(ws.getPivotFieldItem(pivot, 0, 1).visible, true);
    ws.setPivotFieldSort(pivot, 0, f.PivotSortOrder.Ascending);
    ws.setPivotFieldSort(pivot, 0, f.PivotSortOrder.Descending, { byField: 'Sum of Amount' });
    assert.equal(ws.listPivotFields(pivot)[0].sortByField, 'Sum of Amount');
  });
});

test('criterion 13: setIterative validates maxChange and maxIterations up front', async () => {
  await withWorkbook((wb) => {
    assert.throws(() => wb.setIterative({ maxChange: -1 }), RangeError);
    assert.throws(() => wb.setIterative({ maxIterations: 1.5 }), TYPE_ERROR);
  });
});

test('criterion 13: formula text read back from the engine carries the equals sign once', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setFormula(1, 1, '=A1+B1');
    assert.ok(ws.getFormulaR1C1(1, 1).startsWith('='));
    assert.ok(!ws.getFormulaR1C1(1, 1).startsWith('=='));
    ws.setFormula(0, 0, '=1.50');
    assert.equal(ws.getFormula(0, 0), '=1.5');

    ws.addCfRule({ ranges: [rect(0, 0, 3, 0)], type: 0, formula1: '=A1>1', id: CF_ID_ROUND_TRIP });
    const rule = ws.listCfRules()[0];
    assert.ok(rule.formula1.startsWith('=') && !rule.formula1.startsWith('=='), rule.formula1);
    ws.removeCfRule(CF_ID_ROUND_TRIP);
    ws.addCfRule({ ...plain(rule), id: CF_ID_ROUND_TRIP_2 });
    assert.equal(ws.listCfRules()[0].formula1, rule.formula1);

    const reloaded = f.Workbook.load(wb.save());
    try {
      assert.equal(reloaded.sheet(0).getFormula(1, 1), '=A1+B1');
      assert.equal(reloaded.sheet(0).listCfRules()[0].formula1, rule.formula1);
    } finally {
      reloaded.dispose();
    }
  });
});

// ---- Criterion 14: ErrorDisplayName --------------------------------------------------------

test('criterion 14: ErrorDisplayName is a Formulon member and a named export', async () => {
  const mod = await loadPackage();
  const f = await getFormulon();
  assert.equal(f.ErrorDisplayName, mod.ErrorDisplayName);
  assert.equal(f.ErrorDisplayName[f.ErrorCode.Div0], '#DIV/0!');
  assert.equal(f.ErrorDisplayName[f.ErrorCode.NA], '#N/A');
  assert.equal(f.ErrorDisplayName[999], undefined);
  assert.ok(Object.isFrozen(f.ErrorDisplayName));
  for (const code of Object.values(f.ErrorCode)) {
    assert.equal(typeof f.ErrorDisplayName[code], 'string', `code ${code}`);
    assert.match(f.ErrorDisplayName[code], /^#/);
  }
  assert.equal('getErrorDisplayName' in f, false);
  assert.equal('getErrorDisplayName' in mod, false);
});
