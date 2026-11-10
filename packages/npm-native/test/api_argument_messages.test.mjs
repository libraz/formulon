// Argument errors read `<Class>.<method> <path> <constraint>`: the class and status code of each refusal and its
// exact message, for facade, reader and engine errors alike.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');

// The staged entry named by the package's own "main" field.
async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  if (!pkg.main) throw new Error('package.json is missing "main"');
  return import(pathToFileURL(path.resolve(pkgRoot, pkg.main)).href);
}

let formulonPromise;
function getFormulon() {
  formulonPromise ??= loadPackage().then((mod) => mod.createFormulon());
  return formulonPromise;
}

const PNG_1X1 = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
  'base64',
);

async function withWorkbook(fn) {
  const f = await getFormulon();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

// Words no argument message may contain: positional labels, C spellings and patch-op words.
const BANNED = [
  /\bargument \d/,
  /`/,
  /\b(?:SET|KEEP|CLEAR|ABSENT)\b/,
  /\b(?:FM_|fm_)[A-Za-z_]+/,
  /\bNULL\b/,
  /struct_size/,
  /_op\b/,
  /spec->/,
  /\b(?:uint32|int32|2\^63)\b/,
  /4294967295/,
  /\ba (?:object|array)\b/,
];

function assertPublicWords(message) {
  for (const banned of BANNED) assert.doesNotMatch(message, banned, message);
  // Only the leading `Class.method` names the call; one inside the body is a relabelled word.
  assert.doesNotMatch(message.replace(/^[A-Z]\w*\.\w+ /, ''), /\b[A-Z]\w*\.[a-z]\w* /, message);
}

/** `fn` throws `ctor` (TypeError / RangeError) with exactly `message`. */
function refusesWith(fn, ctor, message) {
  assert.throws(
    fn,
    (e) => {
      assert.ok(e instanceof ctor, `expected ${ctor.name}, got ${e?.name}: ${e?.message}`);
      assert.notEqual(e.name, 'FormulonError', e.message);
      assert.equal(e.message, message);
      assertPublicWords(e.message);
      return true;
    },
    message,
  );
}

/** `fn` throws a FormulonError with `code` and exactly `message`. */
function failsWith(fn, code, message) {
  assert.throws(
    fn,
    (e) => {
      assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e?.name}: ${e?.message}`);
      assert.equal(e.code, code, `code ${e.code} (${e.message})`);
      assert.equal(e.message, message);
      assertPublicWords(e.message);
      return true;
    },
    message,
  );
}

/** Error thrown by `fn`; fails when `fn` does not throw. */
function thrownBy(fn) {
  try {
    fn();
  } catch (e) {
    return e;
  }
  assert.fail('expected the call to throw');
}

/** One live handle of each kind the messages name, built in `wb`. */
function makeHandles(f, wb) {
  const ws = wb.sheet(0);
  ws.setText(0, 0, 'Product');
  ws.setText(0, 1, 'Amount');
  ws.setNumber(1, 1, 5);
  const table = ws.addTable({ range: 'A1:B2', name: 'Tbl' });
  const cache = wb.addPivotCache();
  const pivot = ws.addPivot({ name: 'P1', cacheId: cache.cacheId, anchor: { row: 10, col: 4 } });
  const chart = ws.addChart({
    type: f.ChartPlotType.Bar,
    range: 'Sheet1!$A$1:$B$2',
    legend: f.ChartLegendPosition.Right,
  });
  const shape = ws.addShape({
    presetGeometry: 'rect',
    anchor: { kind: f.AnchorKind.OneCell, fromRow: 2, widthEmu: 914400, heightEmu: 457200 },
  });
  const picture = ws.addPicture(PNG_1X1);
  return { ws, table, cache, pivot, chart, shape, picture };
}

test('argument errors read Class.method param constraint', async () => {
  await withWorkbook((wb, f) => {
    const { ws, pivot, chart, shape } = makeHandles(f, wb);
    const shape3 = /^[A-Za-z]+(?:\.[A-Za-z0-9]+)+ \S+ \S/;
    const battery = [
      // facade-checked
      [() => ws.updateView({ zoomPct: 1.5 }), RangeError, 'Sheet.updateView patch.zoomPct must be an integer'],
      [() => ws.updateView({ zoomPct: 'x' }), TypeError, 'Sheet.updateView patch.zoomPct must be a number'],
      [() => wb.addPerson({ displayName: 1 }), TypeError, 'Workbook.addPerson person.displayName must be a string'],
      [() => ws.addValidation({}), TypeError, 'Sheet.addValidation validation.ranges must be an array'],
      // reader-checked
      [() => wb.ooxml.addFont({ bogus: 1 }), TypeError, 'WorkbookOoxml.addFont record has no field bogus'],
      // engine-checked
      [
        () => pivot.getDataField(99),
        f.StatusCode.IndexOutOfRange,
        'Pivot.getDataField dataFieldIndex 99 is out of range (0 items)',
      ],
      [
        () => wb.updateIterative({ maxIterations: -1 }),
        f.StatusCode.InvalidArgument,
        'Workbook.updateIterative patch.maxIterations must not be negative',
      ],
      [
        () => chart.update({ anchor: { fromRow: -1, fromCol: 0, toRow: 5, toCol: 5 } }),
        f.StatusCode.InvalidArgument,
        'Chart.update patch.anchor.fromRow must not be negative',
      ],
      [() => shape.update({ name: '' }), f.StatusCode.InvalidArgument, 'Shape.update patch.name must not be empty'],
    ];
    for (const [fn, expected, message] of battery) {
      if (typeof expected === 'number') failsWith(fn, expected, message);
      else refusesWith(fn, expected, message);
      assert.match(message, shape3);
    }
  });
});

test('ooxml and drawing calls name their own class', async () => {
  await withWorkbook((wb, f) => {
    const { ws, table, chart, shape, picture } = makeHandles(f, wb);
    const cases = [
      [() => wb.ooxml.addFont({ bogus: 1 }), 'WorkbookOoxml.addFont '],
      [() => ws.ooxml.getCellXfIndex(99999999, 0), 'SheetOoxml.getCellXfIndex '],
      [() => chart.update({ bogus: 1 }), 'Chart.update '],
      [() => shape.update({ bogus: 1 }), 'Shape.update '],
      [() => picture.update({ bogus: 1 }), 'Picture.update '],
      [() => table.update({ bogus: 1 }), 'Table.update '],
    ];
    for (const [fn, prefix] of cases) {
      const e = thrownBy(fn);
      assert.ok(e.message.startsWith(prefix), `${prefix.trim()}: ${e.message}`);
      assert.doesNotMatch(e.message, /^Drawing\./, e.message);
      assertPublicWords(e.message);
    }
  });
});

test('PivotCache argument errors name PivotCache methods', async () => {
  await withWorkbook((wb, f) => {
    const { cache } = makeHandles(f, wb);
    const cases = [
      [() => cache.listFieldSharedItems(99), 'PivotCache.listFieldSharedItems '],
      [() => cache.getRecordValue(99, 0), 'PivotCache.getRecordValue '],
      [() => cache.addFieldSharedItem(99, 1), 'PivotCache.addFieldSharedItem '],
    ];
    for (const [fn, prefix] of cases) {
      const e = thrownBy(fn);
      assert.ok(e.message.startsWith(prefix), `${prefix.trim()}: ${e.message}`);
      assert.doesNotMatch(e.message, /Workbook\./, e.message);
      assertPublicWords(e.message);
    }
  });
});

test('engine refusals name the public field', async () => {
  await withWorkbook((wb, f) => {
    const { table, shape, pivot } = makeHandles(f, wb);
    const invalid = f.StatusCode.InvalidArgument;
    failsWith(
      () => table.updateStyleOptions({ styleName: '' }),
      invalid,
      'Table.updateStyleOptions patch.styleName must not be empty',
    );
    failsWith(() => shape.update({ name: '' }), invalid, 'Shape.update patch.name must not be empty');
    failsWith(
      () => pivot.getDataField(99),
      f.StatusCode.IndexOutOfRange,
      'Pivot.getDataField dataFieldIndex 99 is out of range (0 items)',
    );
    failsWith(
      () => wb.updateIterative({ maxIterations: 40000 }),
      invalid,
      'Workbook.updateIterative patch.maxIterations must be in 1..32767',
    );
  });
});

test('a parameter name inside a message body is not relabelled', async () => {
  await withWorkbook((wb, f) => {
    const { table } = makeHandles(f, wb);
    failsWith(() => table.insertRows(1, 0), f.StatusCode.InvalidArgument, 'Table.insertRows last 0 is before first 1');
    failsWith(() => table.insertRows(undefined, 1), f.StatusCode.InvalidArgument, 'Table.insertRows last needs first');
  });
});

test('updateIterative maxIterations 40000 and 0 are InvalidArgument', async () => {
  await withWorkbook((wb, f) => {
    const message = 'Workbook.updateIterative patch.maxIterations must be in 1..32767';
    failsWith(() => wb.updateIterative({ maxIterations: 40000 }), f.StatusCode.InvalidArgument, message);
    failsWith(() => wb.updateIterative({ maxIterations: 0 }), f.StatusCode.InvalidArgument, message);
    wb.updateIterative({ maxIterations: 32767 });
    assert.equal(wb.getIterative().maxIterations, 32767);
  });
});

test('updateIterative maxIterations -1 is InvalidArgument and 1.5 is RangeError', async () => {
  await withWorkbook((wb, f) => {
    failsWith(
      () => wb.updateIterative({ maxIterations: -1 }),
      f.StatusCode.InvalidArgument,
      'Workbook.updateIterative patch.maxIterations must not be negative',
    );
    refusesWith(
      () => wb.updateIterative({ maxIterations: 1.5 }),
      RangeError,
      'Workbook.updateIterative patch.maxIterations must be an integer',
    );
  });
});

test('boolean arguments and fields refuse non-booleans', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    refusesWith(() => wb.setDate1904('yes'), TypeError, 'Workbook.setDate1904 date1904 must be a boolean');
    refusesWith(() => wb.setCalcOnSave(null), TypeError, 'Workbook.setCalcOnSave calcOnSave must be a boolean');
    refusesWith(() => wb.setFullPrecision(1), TypeError, 'Workbook.setFullPrecision fullPrecision must be a boolean');
    refusesWith(
      () => wb.parseReference('R1C1', { r1c1: 1 }),
      TypeError,
      'Workbook.parseReference options.r1c1 must be a boolean',
    );
    refusesWith(
      () => wb.formatReference({ range: 'A1' }, { absolute: 1 }),
      TypeError,
      'Workbook.formatReference options.absolute must be a boolean',
    );
    refusesWith(
      () => ws.updateView({ showGridLines: 1 }),
      TypeError,
      'Sheet.updateView patch.showGridLines must be a boolean',
    );
    wb.setDate1904(true);
    wb.setDate1904(false);
  });
});

test('an engine message that leads with no argument keeps its wording under the call label', async () => {
  await withWorkbook((wb, f) => {
    const invalid = f.StatusCode.InvalidArgument;
    failsWith(
      () => wb.setDefinedName({ name: 'A1', formula: '1' }),
      invalid,
      'Workbook.setDefinedName definedName.name is not a valid defined name',
    );
    failsWith(
      () => wb.sheet(0).addTable({ range: 'A1:B3', name: 'Has Space', columns: ['a', 'b'] }),
      invalid,
      'Sheet.addTable input.name is not a valid table name',
    );
  });
});

test('updateAxis with an unknown id is NotFound and names the kind form', async () => {
  await withWorkbook((wb, f) => {
    const { chart } = makeHandles(f, wb);
    failsWith(
      () => chart.updateAxis(99999, { visible: false }),
      f.StatusCode.NotFound,
      'Chart.updateAxis axis has no axis with id 99999. Pass { kind: ChartAxisKind.Value } to address an axis by kind.',
    );
  });
});

test('a negative integer in a record is InvalidArgument', async () => {
  await withWorkbook((wb, f) => {
    const { ws, pivot, chart } = makeHandles(f, wb);
    const invalid = f.StatusCode.InvalidArgument;
    failsWith(() => ws.updateView({ zoomPct: -5 }), invalid, 'Sheet.updateView patch.zoomPct must not be negative');
    failsWith(
      () => ws.updatePageSetup({ scalePct: -5 }),
      invalid,
      'Sheet.updatePageSetup patch.scalePct must not be negative',
    );
    failsWith(
      () => ws.updateSheetProtection({ spinCount: -3 }),
      invalid,
      'Sheet.updateSheetProtection patch.spinCount must not be negative',
    );
    failsWith(
      () => pivot.updateOptions({ indent: -5 }),
      invalid,
      'Pivot.updateOptions patch.indent must not be negative',
    );
    failsWith(
      () => chart.update({ anchor: { fromRow: -1, fromCol: 0, toRow: 5, toCol: 5 } }),
      invalid,
      'Chart.update patch.anchor.fromRow must not be negative',
    );
  });
});

test('a fractional integer in a record is RangeError', async () => {
  await withWorkbook((wb, f) => {
    const { ws, pivot, chart } = makeHandles(f, wb);
    refusesWith(() => ws.updateView({ zoomPct: 1.5 }), RangeError, 'Sheet.updateView patch.zoomPct must be an integer');
    refusesWith(
      () => ws.updatePageSetup({ scalePct: 1.5 }),
      RangeError,
      'Sheet.updatePageSetup patch.scalePct must be an integer',
    );
    refusesWith(
      () => ws.updateSheetProtection({ spinCount: 1.5 }),
      RangeError,
      'Sheet.updateSheetProtection patch.spinCount must be an integer',
    );
    refusesWith(
      () => pivot.updateOptions({ indent: 1.5 }),
      RangeError,
      'Pivot.updateOptions patch.indent must be an integer',
    );
    refusesWith(
      () => chart.update({ anchor: { fromRow: 1.5, fromCol: 0, toRow: 5, toCol: 5 } }),
      RangeError,
      'Chart.update patch.anchor.fromRow must be an integer',
    );
  });
});

test('static factories and module functions refuse extra arguments', async () => {
  const f = await getFormulon();
  refusesWith(() => f.Workbook.createDefault('x'), TypeError, 'Workbook.createDefault takes at most 0 arguments');
  refusesWith(() => f.Workbook.createEmpty('x'), TypeError, 'Workbook.createEmpty takes at most 0 arguments');
  refusesWith(
    () => f.Workbook.createForProfile('win-365-en_US', 'x'),
    TypeError,
    'Workbook.createForProfile takes at most 1 argument',
  );
  refusesWith(
    () => f.Workbook.load(new Uint8Array(0), { profile: 'win-365-en_US' }),
    TypeError,
    'Workbook.load takes at most 1 argument',
  );
  refusesWith(() => f.getVersion(1), TypeError, 'Formulon.getVersion takes at most 0 arguments');
  refusesWith(() => f.evaluateFormula('=1', 2), TypeError, 'Formulon.evaluateFormula takes at most 1 argument');
});
