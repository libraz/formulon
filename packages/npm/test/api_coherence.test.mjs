// API coherence contract on the staged WASM package: patch keys (unknown keys, null rules, anchor
// triggers, read-modify-write updates), date serial conversion, Shape / Slicer drawing operations,
// ChartSourceKind, geometry-mode defaults and the `wasm` option of createFormulon().
// FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const repoRoot = path.resolve(pkgRoot, '..', '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';
const threadsEntry = entry === './threads';

async function readPackageJson() {
  return JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
}

async function loadPackage() {
  const pkg = await readPackageJson();
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

let formulonPromise;
function getFormulon() {
  formulonPromise ??= loadPackage().then((mod) => mod.createFormulon());
  return formulonPromise;
}

const plain = (v) => JSON.parse(JSON.stringify(v));

function assertFormulonError(fn, code, label) {
  assert.throws(
    fn,
    (e) => {
      assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
      assert.equal(e.code, code, `code ${e.code} (${e.message})`);
      return true;
    },
    label,
  );
}

// A plain TypeError / RangeError: a misuse caught before the engine, never a FormulonError.
function assertMisuse(fn, ctor, label) {
  assert.throws(
    fn,
    (e) => {
      assert.ok(e instanceof ctor, `${label ?? ''} expected ${ctor.name}, got ${e?.name}: ${e?.message}`);
      assert.notEqual(e.name, 'FormulonError');
      return true;
    },
    label,
  );
}

async function withWorkbook(fn) {
  const f = await getFormulon();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

const rect = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });

// ---- Fixtures ----------------------------------------------------------------------------------

// A1:B3 holds "Product/Amount" and two data rows; the table covers A1:B3.
function addSalesTable(ws, name = 'Sales', lastRow = 2) {
  ws.setText(0, 0, 'Product');
  ws.setText(0, 1, 'Amount');
  ws.setText(1, 0, 'A');
  ws.setNumber(1, 1, 10);
  ws.setText(2, 0, 'B');
  ws.setNumber(2, 1, 20);
  return ws.addTable({
    range: rect(0, 0, lastRow, 1),
    name,
    columns: ['Product', 'Amount'],
    styleName: 'TableStyleMedium2',
  });
}

function addBox(f, ws, name, extra = {}) {
  return ws.addShape({
    name,
    presetGeometry: 'rect',
    anchorKind: f.AnchorKind.TwoCell,
    fromRow: 1,
    fromCol: 1,
    toRow: 5,
    toCol: 4,
    fillKind: f.ShapePaint.Solid,
    fillColor: '#FF0000',
    lineKind: f.ShapePaint.Solid,
    lineColor: '#00FF00',
    ...extra,
  });
}

const SLICER_ANCHOR = { fromRow: 4, fromCol: 3, toRow: 14, toCol: 5 };

function addSlicerOf(wb, ws) {
  return ws.addSlicer('Sales', 0, { ...SLICER_ANCHOR, caption: 'Pick', style: 'SlicerStyleLight3' });
}

function sparklineGroup(f) {
  return {
    type: f.SparklineType.Column,
    colorSeries: { kind: 1, argb: 0xff336699, theme: 0, tint: 0, indexed: 0 },
    colorNegative: { kind: 2, argb: 0, theme: 5, tint: -0.25, indexed: 0 },
    colorLow: { kind: 3, argb: 0, theme: 0, tint: 0, indexed: 12 },
    markers: true,
    negative: true,
    displayXAxis: true,
    minAxisType: f.SparklineAxisType.Custom,
    manualMin: -1.5,
    manualMax: 8.5,
    lineWeight: 1.25,
    emptyCellsAs: f.BlanksAs.Zero,
    sparklines: [
      { source: 'Sheet1!A1:D1', location: 'F1' },
      { source: 'Sheet1!A2:D2', location: 'F2' },
    ],
  };
}

const PROTECTED = {
  name: 'Totals',
  ranges: [rect(1, 1, 2, 2), rect(9, 0, 11, 0)],
  algorithmName: 'SHA-512',
  hashValue: 'aA==',
  saltValue: 'cw==',
  spinCount: 100000,
  securityDescriptor: 'O:WDG:WD',
};

// Fills Sheet1 A1:B3 with a Region / Amount table and returns a cache over it with shared items.
function addPivotOver(wb, f) {
  const ws = wb.sheet(0);
  ws.setText(0, 0, 'Region');
  ws.setText(0, 1, 'Amount');
  ws.setText(1, 0, 'East');
  ws.setNumber(1, 1, 10);
  ws.setText(2, 0, 'West');
  ws.setNumber(2, 1, 30);
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ range: 'A1:B3', sheet: 'Sheet1' });
  cache.addField('Region');
  cache.addField('Amount');
  cache.addFieldSharedItem(0, 'East');
  cache.addFieldSharedItem(0, 'West');
  const pivot = ws.addPivot({ name: 'P1', cacheId: cache.cacheId, anchor: { row: 0, col: 4 } });
  pivot.addField({ sourceName: 'Region', axis: f.PivotAxis.Row });
  return { cache, pivot, ws };
}

// ---- Criterion 3: patch keys ---------------------------------------------------------------------
//
// Parameter model (pairwise, strength 2, generated with coverwise pairwise_generate, seed 7):
//   method   : updateCellFormat | Shape.update | Slicer.update | Table.update | Table.updateStyleOptions | updateSparklineGroup | updateProtectedRange
//   keyKind  : unknown | recordOnly | nullable | nonNullable | anchorNoTrigger
//   depth    : top | nested
//   valueKind: valid | null
// Constraints:
//   anchorNoTrigger -> method in {Shape.update, Slicer.update}, valueKind = valid
//   recordOnly      -> method in {Shape.update, Slicer.update, Table.update}, valueKind = valid
//   nullable        -> method in {Shape.update, Slicer.update, Table.updateStyleOptions, updateSparklineGroup, updateCellFormat}, valueKind = null
//   updateCellFormat & nullable -> depth = nested
//   depth = nested  -> method = updateCellFormat and keyKind in {unknown, nullable, nonNullable}
//   unknown -> valid; nonNullable -> null
// The 23 rows below are the generated suite; the expectation of a row follows from keyKind alone:
//   unknown / nonNullable / anchorNoTrigger -> TypeError and nothing changed;
//   recordOnly -> accepted and nothing changed; nullable -> accepted and the key is cleared.
const PATCH_ROWS = [
  ['Shape.update', 'recordOnly', 'top', 'valid'],
  ['updateCellFormat', 'unknown', 'nested', 'valid'],
  ['Slicer.update', 'anchorNoTrigger', 'top', 'valid'],
  ['Shape.update', 'nullable', 'top', 'null'],
  ['updateSparklineGroup', 'nullable', 'top', 'null'],
  ['Shape.update', 'anchorNoTrigger', 'top', 'valid'],
  ['Table.update', 'unknown', 'top', 'valid'],
  ['updateProtectedRange', 'unknown', 'top', 'valid'],
  ['updateCellFormat', 'nonNullable', 'top', 'null'],
  ['updateProtectedRange', 'nonNullable', 'top', 'null'],
  ['updateSparklineGroup', 'unknown', 'top', 'valid'],
  ['Slicer.update', 'nullable', 'top', 'null'],
  ['Table.update', 'nonNullable', 'top', 'null'],
  ['Shape.update', 'nonNullable', 'top', 'null'],
  ['Shape.update', 'unknown', 'top', 'valid'],
  ['Slicer.update', 'unknown', 'top', 'valid'],
  ['updateCellFormat', 'nullable', 'nested', 'null'],
  ['updateSparklineGroup', 'nonNullable', 'top', 'null'],
  ['Slicer.update', 'recordOnly', 'top', 'valid'],
  ['updateCellFormat', 'nonNullable', 'nested', 'null'],
  ['Slicer.update', 'nonNullable', 'top', 'null'],
  ['Table.updateStyleOptions', 'nullable', 'top', 'null'],
  ['Table.update', 'recordOnly', 'top', 'valid'],
];

// Per method: a fixture, how to apply a patch, how to read the state back, and the patches per key kind.
// `nullable.cleared` checks the state after the clear.
const SUBJECTS = {
  updateCellFormat: {
    setup(f, wb) {
      const ws = wb.sheet(0);
      ws.updateCellFormat('A1', { numberFormat: '0.00', font: { name: 'Arial', sizePt: 14, bold: true } });
      return ws;
    },
    apply: (t, patch) => t.updateCellFormat('A1', patch),
    read: (t) => plain(t.getCellFormat('A1')),
    unknown: { top: { bogus: 1 }, nested: { font: { bogus: 1 } } },
    nonNullable: { top: { numberFormat: null }, nested: { font: { name: null } } },
    // A cell font always carries a size, so the clearable nested key is a flag.
    nullable: { nested: { font: { bold: null } } },
    cleared: (state) => assert.equal(state.font.bold, null),
  },
  'Shape.update': {
    setup(f, wb) {
      return addBox(f, wb.sheet(0), 'S1');
    },
    apply: (t, patch) => t.update(patch),
    read: (t) => plain(t.getProperties()),
    unknown: { top: { bogus: 1 } },
    recordOnly: { top: { objectId: 9999, kind: 12345 } },
    nullable: { top: { fillKind: null, lineKind: null, fillColor: null, lineColor: null } },
    cleared: (state) => {
      assert.equal(state.fillKind, null);
      assert.equal(state.fillColor, null);
      assert.equal(state.lineKind, null);
      assert.equal(state.lineColor, null);
    },
    nonNullable: { top: { name: null } },
    anchorNoTrigger: [
      { editAs: 1 },
      { fromRow: 3 },
      { fromCol: 3 },
      { toRow: 6 },
      { toCol: 6 },
      { widthEmu: 100 },
      { heightEmu: 100 },
      { fromColOffsetEmu: 5 },
    ],
  },
  'Slicer.update': {
    setup(f, wb) {
      addSalesTable(wb.sheet(0));
      return addSlicerOf(wb, wb.sheet(0));
    },
    apply: (t, patch) => t.update(patch),
    read: (t) => plain(t.getProperties()),
    unknown: { top: { bogus: 1 } },
    recordOnly: {
      top: { name: 'Other', objectId: 777, cacheName: 'x', tableName: 'x', columnName: 'x', anchorKind: 0 },
    },
    nullable: { top: { style: null } },
    cleared: (state) => assert.equal(state.style, null),
    nonNullable: { top: { caption: null } },
    anchorNoTrigger: [{ toRow: 20 }, { toCol: 9 }, { toRowOffsetEmu: 5 }, { fromColOffsetEmu: 5 }],
  },
  'Table.update': {
    setup(f, wb) {
      return addSalesTable(wb.sheet(0));
    },
    apply: (t, patch) => t.update(patch),
    read: (t) => plain(t.getProperties()),
    unknown: { top: { bogus: 1 } },
    recordOnly: { top: { columns: [{ name: 'X' }, { name: 'Y' }] } },
    nonNullable: { top: { headerRow: null } },
  },
  'Table.updateStyleOptions': {
    setup(f, wb) {
      return addSalesTable(wb.sheet(0));
    },
    apply: (t, patch) => t.updateStyleOptions(patch),
    read: (t) => plain(t.getStyleOptions()),
    unknown: { top: { bogus: 1 } },
    nullable: { top: { styleName: null } },
    cleared: (state) => assert.equal(state.styleName, null),
  },
  updateSparklineGroup: {
    setup(f, wb) {
      const ws = wb.sheet(0);
      ws.addSparklineGroup(sparklineGroup(f));
      return ws;
    },
    apply: (t, patch) => t.updateSparklineGroup(0, patch),
    read: (t) => plain(t.listSparklineGroups()[0]),
    unknown: { top: { bogus: 1 } },
    nullable: { top: { manualMin: null } },
    cleared: (state) => assert.equal(state.manualMin, null),
    nonNullable: { top: { markers: null } },
  },
  updateProtectedRange: {
    setup(f, wb) {
      const ws = wb.sheet(0);
      ws.addProtectedRange(PROTECTED);
      return ws;
    },
    apply: (t, patch) => t.updateProtectedRange(0, patch),
    read: (t) => plain(t.listProtectedRanges()[0]),
    unknown: { top: { bogus: 1 } },
    nonNullable: { top: { name: null } },
  },
};

for (const [method, keyKind, depth, valueKind] of PATCH_ROWS) {
  test(`patch: ${method} / ${keyKind} / ${depth} / ${valueKind}`, async () => {
    await withWorkbook((wb, f) => {
      const subject = SUBJECTS[method];
      const target = subject.setup(f, wb);
      const before = subject.read(target);
      const patches = [].concat(subject[keyKind]?.[depth] ?? subject[keyKind]);
      assert.ok(patches.length > 0 && patches[0] !== undefined, 'the model row has a patch');
      for (const patch of patches) {
        const label = JSON.stringify(patch);
        switch (keyKind) {
          case 'unknown':
          case 'nonNullable':
          case 'anchorNoTrigger':
            assertMisuse(() => subject.apply(target, patch), TypeError, label);
            assert.deepEqual(subject.read(target), before, `${label} left the state unchanged`);
            break;
          case 'recordOnly':
            assert.doesNotThrow(() => subject.apply(target, patch), label);
            assert.deepEqual(subject.read(target), before, `${label} is ignored`);
            break;
          case 'nullable':
            assert.doesNotThrow(() => subject.apply(target, patch), label);
            subject.cleared(subject.read(target));
            break;
          default:
            assert.fail(`unmodelled key kind ${keyKind}`);
        }
      }
    });
  });
}

test('patch: an unknown key is a TypeError at the top level of updateCellFormat and nested in every group', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    const before = plain(ws.getCellFormat('A1'));
    assertMisuse(() => ws.updateCellFormat('A1', { bogus: 1 }), TypeError);
    for (const group of ['font', 'fill', 'border', 'alignment']) {
      assertMisuse(() => ws.updateCellFormat('A1', { [group]: { bogus: 1 } }), TypeError, group);
    }
    assertMisuse(() => ws.updateCellFormat('A1', { font: { bold: true, bogus: 1 } }), TypeError);
    assert.deepEqual(plain(ws.getCellFormat('A1')), before);
    // The same keys without the bogus one are accepted.
    ws.updateCellFormat('A1', { font: { bold: true } });
    assert.equal(ws.getCellFormat('A1').font.bold, true);
  });
});

test('patch: an unknown key is a TypeError on every other patch entry point', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const table = addSalesTable(ws);
    const shape = addBox(f, ws, 'S1');
    ws.addSparklineGroup(sparklineGroup(f));
    ws.addProtectedRange(PROTECTED);
    const slicer = addSlicerOf(wb, ws);
    assertMisuse(() => ws.updateView({ bogus: 1 }), TypeError, 'updateView');
    assertMisuse(() => table.update({ bogus: 1 }), TypeError, 'Table.update');
    assertMisuse(() => shape.update({ bogus: 1 }), TypeError, 'Shape.update');
    assertMisuse(() => slicer.update({ bogus: 1 }), TypeError, 'Slicer.update');
    assertMisuse(() => ws.updateSparklineGroup(0, { bogus: 1 }), TypeError, 'updateSparklineGroup');
    assertMisuse(() => ws.updateProtectedRange(0, { bogus: 1 }), TypeError, 'updateProtectedRange');
  });
});

test('patch: Shape.update clears fill and line kind and colour on null and rejects null elsewhere', async () => {
  await withWorkbook((wb, f) => {
    const shape = addBox(f, wb.sheet(0), 'S1');
    shape.update({ fillKind: null, lineKind: null, fillColor: null, lineColor: null, name: 'Renamed' });
    const props = shape.getProperties();
    assert.equal(props.name, 'Renamed');
    assert.equal(props.fillKind, null);
    assert.equal(props.fillColor, null);
    assert.equal(props.lineKind, null);
    assert.equal(props.lineColor, null);
    for (const key of ['presetGeometry', 'text', 'hidden', 'rotationDeg', 'anchorKind']) {
      assertMisuse(() => shape.update({ [key]: null }), TypeError, key);
    }
    assert.equal(shape.getProperties().name, 'Renamed');
  });
});

test('patch: Shape.update and Slicer.update accept the getter record back and keep every field', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const shape = addBox(f, ws, 'S1');
    const shapeBefore = plain(shape.getProperties());
    assert.doesNotThrow(() => shape.update(shape.getProperties()));
    assert.deepEqual(plain(shape.getProperties()), shapeBefore);

    addSalesTable(ws);
    const slicer = addSlicerOf(wb, ws);
    const slicerBefore = plain(slicer.getProperties());
    assert.doesNotThrow(() => slicer.update(slicer.getProperties()));
    assert.deepEqual(plain(slicer.getProperties()), slicerBefore);
  });
});

test('patch: anchor keys are accepted with their trigger key', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const shape = addBox(f, ws, 'S1');
    shape.update({ anchorKind: f.AnchorKind.TwoCell, fromRow: 2, fromCol: 2, toRow: 8, toCol: 6 });
    const moved = shape.getProperties();
    assert.deepEqual([moved.fromRow, moved.fromCol, moved.toRow, moved.toCol], [2, 2, 8, 6]);
    // An update cannot change the anchor kind; a one-cell anchor takes its size from the patch.
    const oneCell = addBox(f, ws, 'S2', { anchorKind: f.AnchorKind.OneCell, widthEmu: 100, heightEmu: 100 });
    oneCell.update({ anchorKind: f.AnchorKind.OneCell, fromRow: 3, widthEmu: 914400, heightEmu: 457200 });
    assert.equal(oneCell.getProperties().widthEmu, 914400);
    assert.equal(oneCell.getProperties().fromRow, 3);
    assert.equal(oneCell.getProperties().anchorKind, f.AnchorKind.OneCell);

    addSalesTable(ws);
    const slicer = addSlicerOf(wb, ws);
    slicer.update({ fromRow: 6, fromCol: 3, toRow: 20, toCol: 6 });
    const s = slicer.getProperties();
    assert.deepEqual([s.fromRow, s.fromCol, s.toRow, s.toCol], [6, 3, 20, 6]);
    // A patch without anchor keys keeps the anchor.
    slicer.update({ caption: 'Renamed' });
    assert.equal(slicer.getProperties().fromRow, 6);
    assert.equal(slicer.getProperties().caption, 'Renamed');
  });
});

test('patch: Slicer.update(style: null) clears the style; an empty patch changes nothing', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    addSalesTable(ws);
    const slicer = addSlicerOf(wb, ws);
    const before = plain(slicer.getProperties());
    assert.equal(before.style, 'SlicerStyleLight3');
    slicer.update({});
    assert.deepEqual(plain(slicer.getProperties()), before);
    slicer.update({ style: null });
    assert.deepEqual(plain(slicer.getProperties()), { ...before, style: null });
  });
});

test('patch: updateSparklineGroup is read-modify-write and every omitted field survives', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.addSparklineGroup(sparklineGroup(f));
    let current = plain(ws.listSparklineGroups()[0]);
    const colorGreen = { kind: 1, argb: 0xff00ff00, theme: 0, tint: 0, indexed: 0 };
    const patches = [
      {},
      { type: f.SparklineType.Stacked },
      { high: true, low: true },
      { manualMin: null },
      { manualMax: 9 },
      { dateAxis: true, dateAxisRange: { sheet: 0, rect: rect(0, 0, 0, 3) } },
      { dateAxisRange: null },
      { colorSeries: colorGreen },
      { lineWeight: 2.25 },
      { emptyCellsAs: f.BlanksAs.Gap },
      { sparklines: [{ source: { sheet: 0, rect: rect(0, 1, 4, 1) }, location: { row: 0, col: 7 } }] },
    ];
    for (const patch of patches) {
      ws.updateSparklineGroup(0, patch);
      current = { ...current, ...patch };
      assert.deepEqual(plain(ws.listSparklineGroups()[0]), current, JSON.stringify(patch));
    }
    // The getter's record goes back unchanged.
    ws.updateSparklineGroup(0, ws.listSparklineGroups()[0]);
    assert.deepEqual(plain(ws.listSparklineGroups()[0]), current);
  });
});

test('patch: updateSparklineGroup with sparklines: [] is InvalidArgument and changes nothing', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.addSparklineGroup(sparklineGroup(f));
    const before = plain(ws.listSparklineGroups()[0]);
    assertFormulonError(() => ws.updateSparklineGroup(0, { sparklines: [] }), f.StatusCode.InvalidArgument);
    assert.deepEqual(plain(ws.listSparklineGroups()[0]), before);
  });
});

test('patch: updateProtectedRange is read-modify-write and every omitted field survives', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.addProtectedRange(PROTECTED);
    ws.addProtectedRange({ name: 'Other', ranges: [rect(20, 0, 20, 0)] });
    let current = plain(ws.listProtectedRanges()[0]);
    const other = plain(ws.listProtectedRanges()[1]);
    for (const patch of [
      {},
      { name: 'Renamed' },
      { spinCount: 5 },
      { securityDescriptor: null },
      { ranges: [rect(3, 3, 4, 4)] },
    ]) {
      ws.updateProtectedRange(0, patch);
      current = { ...current, ...patch };
      assert.deepEqual(plain(ws.listProtectedRanges()[0]), current, JSON.stringify(patch));
    }
    assert.deepEqual(plain(ws.listProtectedRanges()[1]), other, 'the other range is untouched');
    ws.updateProtectedRange(0, ws.listProtectedRanges()[0]);
    assert.deepEqual(plain(ws.listProtectedRanges()[0]), current);
  });
});

test('patch: updateProtectedRange with ranges: [] is InvalidArgument; an omitted ranges keeps them', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.addProtectedRange(PROTECTED);
    const before = plain(ws.listProtectedRanges()[0]);
    assertFormulonError(() => ws.updateProtectedRange(0, { ranges: [] }), f.StatusCode.InvalidArgument);
    assert.deepEqual(plain(ws.listProtectedRanges()[0]), before);
    ws.updateProtectedRange(0, { spinCount: 7 });
    assert.deepEqual(plain(ws.listProtectedRanges()[0].ranges), before.ranges);
    // A null is a misuse on name, ranges and spinCount.
    for (const key of ['name', 'ranges', 'spinCount']) {
      assertMisuse(() => ws.updateProtectedRange(0, { [key]: null }), TypeError, key);
    }
    // The five password and descriptor strings are clearable: null removes them.
    ws.updateProtectedRange(0, {
      legacyPasswordHash: null,
      algorithmName: null,
      hashValue: null,
      saltValue: null,
      securityDescriptor: null,
    });
    const cleared = ws.listProtectedRanges()[0];
    for (const key of ['legacyPasswordHash', 'algorithmName', 'hashValue', 'saltValue', 'securityDescriptor']) {
      assert.equal(cleared[key], null, key);
    }
    assert.equal(ws.listProtectedRanges()[0].spinCount, 7);
  });
});

test('patch: Table.update takes an A1 range and boolean headerRow / totalsRow', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const table = addSalesTable(ws);
    const props = table.getProperties();
    assert.equal(props.headerRow, true);
    assert.equal(props.totalsRow, false);
    assert.equal('headerRowCount' in props, false);
    assert.equal('totalsRowCount' in props, false);

    table.update({ range: 'A1:B5' });
    assert.deepEqual(plain(table.getProperties().range), { sheet: 0, rect: rect(0, 0, 4, 1) });
    assert.equal(table.getProperties().name, 'Sales');
    assert.equal(table.getStyleOptions().styleName, 'TableStyleMedium2');

    table.update({ totalsRow: true });
    assert.equal(table.getProperties().totalsRow, true);
    assert.equal(table.getProperties().headerRow, true);
    table.update({ totalsRow: false, headerRow: false });
    assert.equal(table.getProperties().totalsRow, false);
    assert.equal(table.getProperties().headerRow, false);
    table.update({ headerRow: true });

    // The getter's record goes back unchanged; a SheetRect range is accepted for the table's own sheet.
    const before = plain(table.getProperties());
    table.update(table.getProperties());
    assert.deepEqual(plain(table.getProperties()), before);
    table.update({ range: { sheet: 0, rect: rect(0, 0, 3, 1) } });
    assert.equal(table.getProperties().range.rect.lastRow, 3);

    // A SheetRect on another sheet is refused by the engine.
    wb.addSheet('Other');
    assertFormulonError(
      () => table.update({ range: { sheet: 1, rect: rect(0, 0, 3, 1) } }),
      f.StatusCode.InvalidArgument,
    );
    assert.equal(table.getProperties().range.sheet, 0);
    // The old key and the old count fields are gone.
    assertMisuse(() => table.update({ rect: rect(0, 0, 3, 1) }), TypeError, 'rect');
    assertMisuse(() => table.update({ headerRowCount: 1 }), TypeError, 'headerRowCount');
    assertMisuse(() => table.update({ totalsRow: 1 }), TypeError, 'numeric totalsRow');
    // The style keys and the display name left the table patch: style goes through updateStyleOptions.
    for (const key of ['styleName', 'showRowStripes', 'displayName']) {
      assertMisuse(() => table.update({ [key]: true }), TypeError, key);
    }
  });
});

test('patch: PivotCache.setWorksheetSource takes null to clear and {} for an attribute-less source', async () => {
  await withWorkbook((wb) => {
    const cache = wb.addPivotCache();
    cache.setWorksheetSource({ range: 'A1:B3', sheet: 'Sheet1' });
    assert.deepEqual(plain(cache.getWorksheetSource()), { range: { sheet: 0, rect: rect(0, 0, 2, 1) }, name: null });

    cache.setWorksheetSource({});
    assert.deepEqual(plain(cache.getWorksheetSource()), { range: null, name: null });

    cache.setWorksheetSource({ name: 'Tbl' });
    assert.deepEqual(plain(cache.getWorksheetSource()), { range: null, name: 'Tbl' });

    cache.setWorksheetSource(null);
    assert.equal(cache.getWorksheetSource(), null);
    assertMisuse(() => cache.setWorksheetSource(), TypeError, 'no argument');
    assertMisuse(() => cache.setWorksheetSource({ bogus: 1 }), TypeError, 'unknown key');
  });
});

test('patch: Pivot.addFieldItem binds a cache item by cacheIndex and takes an options bag', async () => {
  await withWorkbook((wb, f) => {
    const { pivot } = addPivotOver(wb, f);
    pivot.addFieldItem(0, { cacheIndex: 0 });
    pivot.addFieldItem(0, { cacheIndex: 1 }, { visible: false });
    const items = pivot.listFieldItems(0);
    assert.equal(items.length, 2);
    assert.deepEqual([items[0].cacheIndex, items[0].visible], [0, true]);
    assert.deepEqual([items[1].cacheIndex, items[1].visible], [1, false]);
    pivot.addFieldItem(0, 'Extra', { visible: false });
    assert.equal(pivot.listFieldItems(0)[2].visible, false);
    pivot.addFieldItem(0, { cacheIndex: 0 }, {});
    assert.equal(pivot.listFieldItems(0)[3].visible, true);
    assertMisuse(() => pivot.addFieldItem(0, { cacheIndex: 0 }, { visible: 'no' }), TypeError, 'non-boolean visible');
    assertMisuse(() => pivot.addFieldItem(0, { cacheIndex: 'x' }), TypeError, 'non-numeric cacheIndex');
    assert.equal('addFieldSharedItem' in pivot, false);
  });
});

test('addDataTable takes A1 strings for its rectangle and input cells', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    ws.setNumber(0, 0, 1);
    ws.setNumber(1, 0, 2);
    ws.setNumber(2, 0, 3);
    ws.addDataTable({ range: 'B2:B3', columnInput: 'A1' });
    const table = ws.getDataTable(1, 1);
    assert.notEqual(table, null);
    assert.equal(table.columnInput.row, 0);
    assert.equal(table.columnInput.col, 0);
    ws.addDataTable({ range: rect(1, 3, 2, 3), columnInput: { row: 0, col: 0 } });
    assert.notEqual(ws.getDataTable(1, 3), null);
  });
});

// ---- Criterion 8: date serial conversion ---------------------------------------------------------

const DATE_CASES = JSON.parse(
  readFileSync(path.join(repoRoot, 'tools', 'api', 'date_serial_cases.json'), 'utf8'),
).cases;
const NUMERIC_TOKENS = {
  NaN: Number.NaN,
  Infinity: Number.POSITIVE_INFINITY,
  '-Infinity': Number.NEGATIVE_INFINITY,
  '-0': -0,
};
const TOLERANCE = 1e-9;

// A numeric field given as a string stands for a non-finite or signed-zero number.
const numeric = (v) => (typeof v === 'string' ? NUMERIC_TOKENS[v] : v);
const jsCases = DATE_CASES.filter((c) => c.only === undefined || c.only.includes('js'));

test('date cases: the shared table has rows for both directions and both epochs', () => {
  assert.ok(jsCases.length >= 40, `rows for js: ${jsCases.length}`);
  for (const direction of ['civil_to_serial', 'serial_to_civil']) {
    for (const date1904 of [false, true]) {
      assert.ok(
        jsCases.some((c) => c.direction === direction && c.date1904 === date1904 && c.expect.js.value !== undefined),
        `${direction} date1904=${date1904} has a value row`,
      );
    }
  }
  assert.ok(jsCases.some((c) => c.expect.js.layer === 1));
  assert.ok(jsCases.some((c) => c.expect.js.layer === 2));
});

// Evaluates the engine's own DATE()/TIME() (civil to serial) or YEAR()..SECOND() (serial to civil).
function evaluateWithEngine(f, row) {
  const wb = f.Workbook.createDefault();
  try {
    wb.setDate1904(row.date1904);
    const ws = wb.sheet(0);
    if (row.direction === 'civil_to_serial') {
      const { year, month, day, hour, minute, second } = row.input;
      ws.setFormula(0, 0, `=DATE(${year},${month},${day})+TIME(${hour},${minute},${second})`);
      wb.recalc();
      return ws.getTypedValue(0, 0).number;
    }
    ws.setNumber(0, 0, numeric(row.input.serial));
    for (const [i, fn] of ['YEAR', 'MONTH', 'DAY', 'HOUR', 'MINUTE', 'SECOND'].entries()) {
      ws.setFormula(0, i + 1, `=${fn}(A1)`);
    }
    wb.recalc();
    const keys = ['year', 'month', 'day', 'hour', 'minute', 'second'];
    return Object.fromEntries(keys.map((key, i) => [key, ws.getTypedValue(0, i + 1).number]));
  } finally {
    wb.dispose();
  }
}

for (const row of jsCases) {
  test(`date cases: ${row.id}`, async () => {
    const f = await getFormulon();
    const input = row.input;
    const run = () => {
      if (row.direction === 'civil_to_serial') {
        const civil = Object.fromEntries(Object.entries(input).map(([k, v]) => [k, numeric(v)]));
        return f.convertCivilToSerial(civil, { date1904: row.date1904 });
      }
      if (input.serial_string !== undefined) {
        return f.convertSerialToCivil(input.serial_string, { date1904: row.date1904 });
      }
      return f.convertSerialToCivil(numeric(input.serial), { date1904: row.date1904 });
    };
    const expected = row.expect.js;
    if (expected.error === undefined) {
      const actual = run();
      if (row.direction === 'civil_to_serial') {
        assert.ok(Math.abs(actual - expected.value) <= TOLERANCE, `${actual} vs ${expected.value}`);
        const evaluated = evaluateWithEngine(f, row);
        assert.ok(Math.abs(actual - evaluated) <= TOLERANCE, `DATE()+TIME() gives ${evaluated}, got ${actual}`);
      } else {
        assert.deepEqual(plain(actual), expected.value);
        assert.deepEqual(evaluateWithEngine(f, row), expected.value, 'YEAR()..SECOND() of the same serial');
      }
    } else if (expected.layer === 1) {
      assertMisuse(run, expected.error === 'RangeError' ? RangeError : TypeError, row.id);
    } else {
      assert.equal(expected.error, 'InvalidArgument');
      assertFormulonError(run, f.StatusCode.InvalidArgument, row.id);
    }
  });
}

test('date functions: members of the Formulon object, date1904 defaults to false', async () => {
  const f = await getFormulon();
  assert.equal(typeof f.convertCivilToSerial, 'function');
  assert.equal(typeof f.convertSerialToCivil, 'function');
  assert.ok(Object.isFrozen(f));
  const civil = { year: 2026, month: 11, day: 8, hour: 6, minute: 7, second: 8 };
  assert.equal(f.convertCivilToSerial(civil), f.convertCivilToSerial(civil, { date1904: false }));
  assert.equal(f.convertCivilToSerial(civil), f.convertCivilToSerial(civil, {}));
  assert.deepEqual(plain(f.convertSerialToCivil(46334.5)), plain(f.convertSerialToCivil(46334.5, { date1904: false })));
  assert.notEqual(
    f.convertCivilToSerial(civil, { date1904: true }),
    f.convertCivilToSerial(civil, { date1904: false }),
  );
});

test('date functions: round trip civil -> serial -> civil on both epochs', async () => {
  const f = await getFormulon();
  for (const date1904 of [false, true]) {
    for (const civil of [
      { year: 1904, month: 1, day: 1, hour: 0, minute: 0, second: 0 },
      { year: 2000, month: 2, day: 29, hour: 23, minute: 59, second: 59 },
      { year: 2026, month: 11, day: 8, hour: 6, minute: 7, second: 8 },
      { year: 9999, month: 12, day: 31, hour: 23, minute: 59, second: 59 },
    ]) {
      const options = { date1904 };
      assert.deepEqual(plain(f.convertSerialToCivil(f.convertCivilToSerial(civil, options), options)), civil);
    }
  }
});

test('date functions: argument misuse is a TypeError before the engine', async () => {
  const f = await getFormulon();
  const civil = { year: 2026, month: 1, day: 1, hour: 0, minute: 0, second: 0 };
  for (const flag of [1, 0, 'true', {}, []]) {
    const options = { date1904: flag };
    assertMisuse(() => f.convertCivilToSerial(civil, options), TypeError, `civil date1904=${String(flag)}`);
    assertMisuse(() => f.convertSerialToCivil(46000, options), TypeError, `serial date1904=${String(flag)}`);
  }
  for (const bare of [true, false, 1, 'x']) {
    assertMisuse(() => f.convertCivilToSerial(civil, bare), TypeError, `civil bare ${String(bare)}`);
    assertMisuse(() => f.convertSerialToCivil(46000, bare), TypeError, `serial bare ${String(bare)}`);
  }
  assertMisuse(() => f.convertCivilToSerial(null), TypeError, 'null civil');
  assertMisuse(() => f.convertCivilToSerial(), TypeError, 'no civil');
  assertMisuse(() => f.convertCivilToSerial({ year: 2026 }), TypeError, 'missing fields');
  assertMisuse(() => f.convertSerialToCivil(null), TypeError, 'null serial');
  assertMisuse(() => f.convertSerialToCivil(), TypeError, 'no serial');
  assertMisuse(() => f.convertSerialToCivil(1n), TypeError, 'bigint serial');
});

// ---- Criterion 9: Shape / Slicer drawing operations ---------------------------------------------

const orderOf = (ws) => ws.listDrawings().map((d) => d.objectId);

test('Shape.rename, setZOrder and snapshot work and Sheet.restoreDrawing brings a shape back', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const a = addBox(f, ws, 'A');
    const b = addBox(f, ws, 'B', { fromRow: 7, toRow: 9 });
    const c = addBox(f, ws, 'C', { fromRow: 11, toRow: 13 });
    assert.deepEqual(orderOf(ws), [a.objectId, b.objectId, c.objectId]);

    c.setZOrder(0);
    assert.deepEqual(orderOf(ws), [c.objectId, a.objectId, b.objectId]);
    a.setZOrder(2);
    assert.deepEqual(orderOf(ws), [c.objectId, b.objectId, a.objectId]);
    assertMisuse(() => a.setZOrder(1.5), RangeError, 'fractional index');
    assertMisuse(() => a.setZOrder('0'), TypeError, 'string index');
    assertFormulonError(() => a.setZOrder(3), f.StatusCode.IndexOutOfRange);
    assert.deepEqual(orderOf(ws), [c.objectId, b.objectId, a.objectId]);

    b.rename('Renamed');
    assert.equal(b.getProperties().name, 'Renamed');
    assert.equal(ws.shape(b.objectId).getProperties().name, 'Renamed');
    assertMisuse(() => b.rename(5), TypeError, 'non-string name');

    const bytes = b.snapshot();
    assert.ok(bytes instanceof Uint8Array);
    assert.ok(bytes.length > 0);
    const id = b.objectId;
    ws.removeDrawing(b);
    assertFormulonError(() => b.getProperties(), f.StatusCode.NotFound);
    const restored = ws.restoreDrawing(bytes);
    assert.equal(restored, b, 'restoring under the old id revives the handle');
    assert.equal(restored.objectId, id);
    assert.equal(b.getProperties().name, 'Renamed');
    assert.equal(b.getProperties().fromRow, 7);
    assert.equal(orderOf(ws).length, 3);
  });
});

test('Shape rename and z-order survive save and load', async () => {
  await withWorkbook(async (wb, f) => {
    const ws = wb.sheet(0);
    const a = addBox(f, ws, 'A');
    const b = addBox(f, ws, 'B', { fromRow: 7, toRow: 9 });
    const c = addBox(f, ws, 'C', { fromRow: 11, toRow: 13 });
    c.setZOrder(0);
    a.rename('First');
    const namesInOrder = (sheet) => orderOf(sheet).map((id) => sheet.shape(id).getProperties().name);
    const expected = namesInOrder(ws);
    assert.deepEqual(expected, ['C', 'First', 'B']);
    void b;

    const loaded = f.Workbook.load(wb.save());
    try {
      assert.deepEqual(namesInOrder(loaded.sheet(0)), expected);
    } finally {
      loaded.dispose();
    }
  });
});

test('Sheet.restoreDrawing puts a snapshot of a shape back after the sheet was saved and reloaded', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    const shape = addBox(f, ws, 'Keep', { text: 'Hello' });
    const bytes = shape.snapshot();
    ws.removeDrawing(shape);
    assert.equal(ws.listShapes().length, 0);
    const loaded = f.Workbook.load(wb.save());
    try {
      assert.equal(loaded.sheet(0).listShapes().length, 0);
      const back = loaded.sheet(0).restoreDrawing(bytes);
      assert.equal(back.getProperties().name, 'Keep');
      assert.equal(back.getProperties().text, 'Hello');
      const again = f.Workbook.load(loaded.save());
      try {
        assert.equal(again.sheet(0).listShapes()[0].getProperties().name, 'Keep');
      } finally {
        again.dispose();
      }
    } finally {
      loaded.dispose();
    }
  });
});

test('Slicer.setZOrder reorders the drawing and survives save and load; Slicer has no rename or snapshot', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    addSalesTable(ws);
    const shape = addBox(f, ws, 'Box');
    const slicer = addSlicerOf(wb, ws);
    const ids = orderOf(ws);
    assert.ok(ids.includes(shape.objectId) && ids.includes(slicer.objectId));

    slicer.setZOrder(0);
    assert.equal(orderOf(ws)[0], slicer.objectId);
    shape.setZOrder(0);
    assert.equal(orderOf(ws)[0], shape.objectId);
    slicer.setZOrder(0);
    assert.equal(orderOf(ws)[0], slicer.objectId);
    assertFormulonError(() => slicer.setZOrder(9), f.StatusCode.IndexOutOfRange);

    const loaded = f.Workbook.load(wb.save());
    try {
      const reloaded = loaded.sheet(0);
      assert.equal(reloaded.listDrawings()[0].kind, slicer.kind);
      assert.equal(reloaded.listSlicers().length, 1);
      assert.equal(reloaded.listSlicers()[0].getProperties().caption, 'Pick');
    } finally {
      loaded.dispose();
    }

    assert.equal(typeof slicer.rename, 'undefined');
    assert.equal(typeof slicer.snapshot, 'undefined');
  });
});

// ---- Smaller contracts -------------------------------------------------------------------------

test('ChartSourceKind is a numeric table and ChartDataSource.kind uses it', async () => {
  await withWorkbook((wb, f) => {
    assert.deepEqual(plain(f.ChartSourceKind), { None: 0, Ref: 1, Literal: 2 });
    assert.ok(Object.isFrozen(f.ChartSourceKind));
    const ws = wb.sheet(0);
    ws.setText(0, 1, 'S1');
    ['a', 'b', 'c'].forEach((cat, r) => {
      ws.setText(r + 1, 0, cat);
      ws.setNumber(r + 1, 1, r + 1);
    });
    const chart = ws.addChart({
      type: f.ChartPlotType.Bar,
      range: 'Sheet1!$A$1:$B$4',
      legend: f.ChartLegendPosition.Right,
    });
    const series = chart.getProperties().plots[0].series[0];
    assert.equal(series.name.kind, f.ChartSourceKind.Ref);
    assert.equal(series.categories.kind, f.ChartSourceKind.Ref);
    assert.equal(series.values.kind, f.ChartSourceKind.Ref);
    assert.equal(series.values.formula, 'Sheet1!$B$2:$B$4');
    assert.equal(series.x.kind, f.ChartSourceKind.None);
    assert.equal(typeof series.x.kind, 'number');
  });
});

test('getCellRectPt and getWidthModel default to the Display mode', async () => {
  await withWorkbook((wb, f) => {
    const ws = wb.sheet(0);
    assert.deepEqual(plain(ws.getCellRectPt('B2:C3')), plain(ws.getCellRectPt('B2:C3', f.GeometryMode.Display)));
    assert.deepEqual(
      plain(ws.getCellRectPt(rect(1, 1, 2, 2))),
      plain(ws.getCellRectPt('B2:C3', f.GeometryMode.Display)),
    );
    assert.deepEqual(plain(ws.getWidthModel()), plain(ws.getWidthModel(f.GeometryMode.Display)));
    assert.throws(
      () => ws.getCellRectPt('B2:C3', 99),
      (e) => e.name === 'FormulonError' && e.code === f.StatusCode.InvalidArgument,
    );
    assert.throws(() => ws.getCellRectPt('B2:C3', '1'), TypeError);
  });
});

test('Pivot.setAnchor takes one cell as a CellAddress or an A1 string', async () => {
  await withWorkbook((wb, f) => {
    const { pivot } = addPivotOver(wb, f);
    assert.doesNotThrow(() => pivot.setAnchor({ row: 3, col: 6 }));
    assert.doesNotThrow(() => pivot.setAnchor('H1'));
    assertMisuse(() => pivot.setAnchor(5), TypeError, 'number');
    assertMisuse(() => pivot.setAnchor(rect(3, 6, 3, 6)), TypeError, 'rectangle');
  });
});

// ---- FormulonOptions.wasm ------------------------------------------------------------------------

// The staged binary that belongs to the entry under test.
async function readWasmBytes() {
  const key = threadsEntry ? './formulon_threads.wasm' : './formulon.wasm';
  const pkg = await readPackageJson();
  const target = pkg.exports?.[key];
  if (!target) throw new Error(`package.json "exports" has no ${key}`);
  const file = typeof target === 'string' ? target : (target.default ?? target.import);
  return new Uint8Array(await readFile(path.resolve(pkgRoot, file)));
}

// A wide formula DAG so recalcParallel has independent work; checks the values it must produce.
function checkParallelRecalc(f) {
  const wb = f.Workbook.createDefault();
  try {
    const ws = wb.sheet(0);
    const branches = 24;
    ws.setNumber(0, 0, 1);
    for (let i = 0; i < branches; i += 1) {
      ws.setFormula(i + 1, 1, `=A1+${i + 2}`);
      ws.setFormula(i + 1, 2, `=B${i + 2}*2`);
    }
    const stats = wb.recalcParallel(4);
    assert.ok(stats.cellsEvaluated > 0);
    if (stats.workerThreadsStarted > 1) {
      // Workers started: the module reached them and ran parallel steps.
      assert.ok(stats.parallelSteps > 0, JSON.stringify(stats));
    } else {
      // A documented serial degradation (launch refused, or the single-threaded build).
      assert.equal(stats.workerThreadsUsed, 0);
    }
    assert.equal(ws.getTypedValue(1, 1).number, 3);
    assert.equal(ws.getTypedValue(branches, 2).number, (1 + branches + 1) * 2);
  } finally {
    wb.dispose();
  }
}

function checkEvaluates(f) {
  const wb = f.Workbook.createDefault();
  try {
    wb.sheet(0).setNumber(0, 0, 20);
    wb.sheet(0).setFormula(0, 1, '=A1+22');
    wb.recalc();
    assert.equal(wb.sheet(0).getTypedValue(0, 1).number, 42);
  } finally {
    wb.dispose();
  }
}

test('createFormulon({ wasm }) accepts the binary as a Uint8Array', async () => {
  const mod = await loadPackage();
  const bytes = await readWasmBytes();
  const f = await mod.createFormulon({ wasm: bytes });
  checkEvaluates(f);
  checkParallelRecalc(f);
});

test('createFormulon({ wasm }) accepts a WebAssembly.Module and recalcParallel runs on it', async () => {
  const mod = await loadPackage();
  const compiled = await WebAssembly.compile(await readWasmBytes());
  const f = await mod.createFormulon({ wasm: compiled });
  checkEvaluates(f);
  checkParallelRecalc(f);
});

test('createFormulon({ wasm: ArrayBuffer }) is a BufferSource too, and null means "not given"', async () => {
  const mod = await loadPackage();
  const bytes = await readWasmBytes();
  const buffer = bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
  checkEvaluates(await mod.createFormulon({ wasm: buffer }));
  checkEvaluates(await mod.createFormulon({ wasm: null }));
});

test('createFormulon({ wasm }) refuses an empty buffer and anything that is neither a BufferSource nor a Module', async () => {
  const mod = await loadPackage();
  for (const bad of [new Uint8Array(0), new ArrayBuffer(0), 'formulon.wasm', 42, true, {}, [], () => 1, new Map()]) {
    await assert.rejects(
      async () => mod.createFormulon({ wasm: bad }),
      (e) => {
        assert.ok(e instanceof TypeError, `${String(bad)}: ${e?.name}: ${e?.message}`);
        return true;
      },
    );
  }
});

test('createFormulon ignores unknown option keys', async () => {
  const mod = await loadPackage();
  checkEvaluates(await mod.createFormulon({ bogus: 1 }));
});
