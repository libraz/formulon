//
// Smoke tests for the staged @libraz/formulon npm package.
//
// These tests intentionally import the *staged* dist/ artefacts (resolved
// via the package's own package.json "main" field), not the raw
// build-wasm/ output. That way we catch staging mistakes -- missing
// files, wrong relative paths, broken "exports" entries -- before the
// tarball ships.
//
// Run via `make npm-test` (which runs `make npm-package` first).

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
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

// Resolve the staged entry point through the package's own package.json
// "main" field. Fails (rather than skips) if the package isn't staged --
// that's the whole point of running these tests against dist/.
async function loadStagedFactory() {
  const raw = await readFile(pkgJsonPath, 'utf8');
  const pkg = JSON.parse(raw);
  const main = pkg.main;
  if (!main) {
    throw new Error(`package.json is missing "main": ${pkgJsonPath}`);
  }
  const mainPath = path.resolve(pkgRoot, main);
  // Node ESM requires file:// URLs for absolute paths on Windows;
  // pathToFileURL is the portable form.
  const mod = await import(pathToFileURL(mainPath).href);
  if (typeof mod.default !== 'function') {
    throw new Error(`expected default export to be a factory, got ${typeof mod.default}`);
  }
  return mod.default;
}

async function loadStagedModule() {
  const raw = await readFile(pkgJsonPath, 'utf8');
  const pkg = JSON.parse(raw);
  return import(pathToFileURL(path.resolve(pkgRoot, pkg.main)).href);
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

// An implicit-intersection formula the XLSB encoder cannot lower, which
// makes the writer emit one per-cell warn record.
function emitXlsbWarning(Module, xlsbFormat) {
  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(wb.setFormula(0, 0, 0, '=@A1:A10').ok);
    assert.ok(wb.recalc().ok);
    assert.ok(wb.saveAs(xlsbFormat).status.ok);
  } finally {
    wb.delete();
  }
}

test('default export is callable factory returning a Module', async () => {
  const factory = await loadStagedFactory();
  assert.equal(typeof factory, 'function');
  const Module = await factory();
  assert.ok(Module && typeof Module === 'object');
  assert.equal(typeof Module.versionString, 'function');
});

test('all declared enum and constant exports resolve at runtime', async () => {
  const mod = await loadStagedModule();
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
  const names = [...dts.matchAll(/^export (?:const )?(?:enum|const) (\w+)/gm)].map((match) => match[1]);
  for (const name of names) {
    assert.ok(name in mod, `missing runtime export ${name}`);
  }
  assert.equal(mod.ValueKind.Error, 4);
  assert.equal(mod.PivotAxis.Row, 0);
  assert.equal(mod.ExternalLinkKind.Dde, 3);
  // Named constants a consumer needs in order to interpret shared record
  // fields (`Value.errorCode`, `ExternalLinkRecord.kind`) or to drive the
  // calc policy. The native package exports the same names and ordinals;
  // `check_binding_drift dts-enums` holds the two in step.
  assert.equal(mod.ErrorCode.Div0, 1);
  assert.equal(mod.ErrorCode.Unknown, 16);
  assert.equal(mod.CalcMode.Manual, 1);
  for (const name of ['ErrorCode', 'CalcMode', 'ExternalLinkKind']) {
    assert.ok(Object.isFrozen(mod[name]), `${name} must be frozen`);
  }
});

test('setLogMinLevel / setLogSink are module-level controls on the Module', async () => {
  const Module = await getModule();
  assert.equal(typeof Module.setLogMinLevel, 'function');
  assert.equal(typeof Module.setLogSink, 'function');
  const mod = await loadStagedModule();
  for (const level of Object.values(mod.LogLevel)) {
    assert.ok(Module.setLogMinLevel(level).ok, `level=${level}`);
  }
  for (const bad of [-1, 5, 99]) {
    const r = Module.setLogMinLevel(bad);
    assert.equal(r.ok, false, `level=${bad} should be rejected`);
    assert.equal(r.status, 2);
  }
  assert.ok(Module.setLogMinLevel(mod.LogLevel.Off).ok);
});

test('a registered sink receives records only once the threshold admits them', async () => {
  const Module = await getModule();
  const mod = await loadStagedModule();
  const records = [];
  assert.ok(Module.setLogSink((record) => records.push(new Uint8Array(record))).ok);
  try {
    // Default threshold is Off: the same save must produce nothing.
    emitXlsbWarning(Module, mod.WorkbookFormat.Xlsb);
    assert.deepEqual(records, []);

    assert.ok(Module.setLogMinLevel(mod.LogLevel.Warn).ok);
    emitXlsbWarning(Module, mod.WorkbookFormat.Xlsb);
    assert.ok(records.length > 0, 'expected at least one record');
    const text = records.map((r) => new TextDecoder().decode(r)).join('');
    assert.match(text, /xlsb\.writer\.formula_downgraded/);
    assert.match(text, /"level":"warn"/);
  } finally {
    assert.ok(Module.setLogMinLevel(mod.LogLevel.Off).ok);
    assert.ok(Module.setLogSink(null).ok);
  }
});

test('a throwing log sink is dropped and leaves the engine usable', async () => {
  const Module = await getModule();
  const mod = await loadStagedModule();
  let calls = 0;
  assert.ok(
    Module.setLogSink(() => {
      calls += 1;
      throw new Error('sink failure');
    }).ok,
  );
  try {
    assert.ok(Module.setLogMinLevel(mod.LogLevel.Warn).ok);
    // The sink runs from deep inside the writer. If its exception reached
    // the C++ frames it would tear them down without running a single
    // destructor, because the binary is linked without exception support.
    emitXlsbWarning(Module, mod.WorkbookFormat.Xlsb);
    assert.ok(calls > 0, 'expected the sink to be invoked');
  } finally {
    assert.ok(Module.setLogMinLevel(mod.LogLevel.Off).ok);
    assert.ok(Module.setLogSink(null).ok);
  }

  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(wb.setFormula(0, 0, 0, '=1+1').ok);
    const after = wb.recalc();
    assert.ok(after.ok, `recalc after a throwing sink: ${JSON.stringify(after)}`);
    assert.equal(wb.getValue(0, 0, 0).value.number, 2);
  } finally {
    wb.delete();
  }
});

// One-shot evaluation is anchored at Sheet1!A1 and read-only on every
// shipped surface, so these must agree value-for-value with the WASM
// package and the Python wheel. The anchor-referencing cases are the ones
// that diverge if a surface writes the formula into A1 and recalcs.
const ONE_SHOT_CASES = [
  { formula: '=1+2', kind: 1, number: 3 },
  { formula: '=A1', kind: 1, number: 0 },
  { formula: '=COUNTA(A1)', kind: 1, number: 0 },
  { formula: '=ISBLANK(A1)', kind: 2, boolean: 1 },
  { formula: '=SUM(A1:A3)', kind: 1, number: 0 },
  { formula: '=SEQUENCE(3)', kind: 1, number: 1 },
  { formula: '=ROWS(SEQUENCE(3))', kind: 1, number: 3 },
];
test('evalFormula evaluates read-only against a blank anchor', async () => {
  const Module = await getModule();
  for (const spec of ONE_SHOT_CASES) {
    const r = Module.evalFormula(spec.formula);
    assert.ok(r.status.ok, `${spec.formula}: ${JSON.stringify(r.status)}`);
    assert.equal(r.value.kind, spec.kind, spec.formula);
    if (spec.number !== undefined) {
      assert.equal(r.value.number, spec.number, spec.formula);
    }
    if (spec.boolean !== undefined) {
      assert.equal(r.value.boolean, spec.boolean, spec.formula);
    }
    assert.equal(r.value.errorCode, 0, spec.formula);
  }
});

// Parses the staged `export enum X { A = 0, ... }` declarations into
// {name: {member: ordinal}}. Name presence alone is not enough: a staged
// bundle carrying a stale ordinal ships a constant that silently means
// something else, and nothing upstream of dist/ can see that.
function declaredOrdinalTables(dts) {
  const tables = {};
  const blocks = dts.matchAll(/^export enum (\w+) \{([^}]*)\}/gm);
  for (const [, name, body] of blocks) {
    const members = {};
    let next = 0;
    for (const line of body.split(',')) {
      const m = line.trim().match(/^(\w+)\s*(?:=\s*(-?\d+))?$/);
      if (!m) continue;
      if (m[2] !== undefined) next = Number(m[2]);
      members[m[1]] = next;
      next += 1;
    }
    tables[name] = members;
  }
  return tables;
}

test('staged constant tables carry the ordinals the staged .d.ts declares', async () => {
  const mod = await loadStagedModule();
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
  const tables = declaredOrdinalTables(dts);
  assert.ok(Object.keys(tables).length > 0, 'expected declared ordinal tables');
  for (const [name, members] of Object.entries(tables)) {
    const runtime = mod[name];
    assert.ok(runtime, `missing runtime table ${name}`);
    assert.deepEqual({ ...runtime }, members, `${name} ordinals differ from dist/formulon.d.ts`);
    assert.ok(Object.isFrozen(runtime), `${name} must be frozen`);
  }
});

// Builds a workbook with one pivot, ready for the filter-contract cases.
// `delete` is embind's auto-added finaliser; the caller must call it.
function makePivotWorkbook(Workbook) {
  const wb = Workbook.createDefault();
  const cacheId = wb.pivotCacheCreate(0).index;
  // A cache without a worksheet source cannot be saved, so the fixture
  // declares the range its two fields and two records occupy.
  assert.ok(wb.pivotCacheSetWorksheetSource(cacheId, { present: true, ref: 'A1:B3', sheet: 'Sheet1' }).ok);
  assert.ok(wb.pivotCacheFieldAdd(cacheId, 'Region').status.ok);
  assert.ok(wb.pivotCacheFieldAdd(cacheId, 'Amount').status.ok);
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const rec = wb.pivotCacheRecordAdd(cacheId).index;
    assert.ok(wb.pivotCacheRecordSetText(cacheId, rec, 0, region).ok);
    assert.ok(wb.pivotCacheRecordSetNumber(cacheId, rec, 1, amount).ok);
  }
  const pivot = wb.pivotCreate(0, 'Pivot1', cacheId, 0, 4);
  assert.ok(pivot.status.ok, JSON.stringify(pivot.status));
  assert.ok(wb.pivotFieldAdd(0, pivot.index, { sourceName: 'Region', axis: 0 }).status.ok);
  const amountField = wb.pivotFieldAdd(0, pivot.index, { sourceName: 'Amount', axis: 2 });
  assert.ok(amountField.status.ok);
  assert.ok(
    wb.pivotDataFieldAdd(0, pivot.index, {
      name: 'Sum of Amount',
      fieldIndex: amountField.index,
      aggregation: 0,
    }).status.ok,
  );
  return { wb, pivot: pivot.index };
}

// greaterThan on a double payload: the shape pivotFilterAt must read back.
const PIVOT_FILTER = Object.freeze({
  axis: 0,
  fieldName: 'Region',
  type: 1,
  valueKind: 1,
  valueDouble: 15,
});

test('pivotFilterAt reads back an added filter field for field', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    assert.ok(wb.pivotFilterAdd(0, pivot, PIVOT_FILTER).ok);
    assert.equal(wb.pivotFilterCount(0, pivot), 1);
    const got = wb.pivotFilterAt(0, pivot, 0);
    assert.ok(got.status.ok, JSON.stringify(got.status));
    assert.equal(got.axis, PIVOT_FILTER.axis);
    assert.equal(got.fieldName, PIVOT_FILTER.fieldName);
    assert.equal(got.type, PIVOT_FILTER.type);
    assert.equal(got.valueKind, PIVOT_FILTER.valueKind);
    assert.equal(got.valueDouble, PIVOT_FILTER.valueDouble);
  } finally {
    wb.delete();
  }
});

test('pivotFilterCount reports only what this session added', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    assert.equal(wb.pivotFilterCount(0, pivot), 0);
    assert.ok(wb.pivotFilterAdd(0, pivot, PIVOT_FILTER).ok);
    assert.equal(wb.pivotFilterCount(0, pivot), 1);
  } finally {
    wb.delete();
  }
});

test('active filters are session state and do not survive save/load', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  let bytes;
  try {
    assert.ok(wb.pivotFilterAdd(0, pivot, PIVOT_FILTER).ok);
    assert.equal(wb.pivotFilterCount(0, pivot), 1);
    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    bytes = saved.bytes;
  } finally {
    wb.delete();
  }
  const reloaded = Module.Workbook.loadBytes(bytes);
  try {
    assert.equal(reloaded.pivotCount(0), 1);
    // The pivot round-trips; its active-filter list deliberately does not.
    assert.equal(reloaded.pivotFilterCount(0, 0), 0);
  } finally {
    reloaded.delete();
  }
});

test('pivotFilterAt rejects an out-of-range index', async () => {
  const Module = await getModule();
  const { wb, pivot } = makePivotWorkbook(Module.Workbook);
  try {
    const got = wb.pivotFilterAt(0, pivot, 99);
    assert.equal(got.status.ok, false);
  } finally {
    wb.delete();
  }
});

test('versionString returns a non-empty string', async () => {
  const Module = await getModule();
  const v = Module.versionString();
  assert.equal(typeof v, 'string');
  assert.ok(v.length > 0, `expected non-empty version, got ${JSON.stringify(v)}`);
});

test('evalFormula(=SUM(1,2,3)) returns NUMBER 6', async () => {
  const Module = await getModule();
  const r = Module.evalFormula('=SUM(1,2,3)');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, VAL.NUMBER);
  assert.equal(r.value.number, 6);
});

test('evalFormula(=1/0) surfaces Excel error as a value (not failed status)', async () => {
  const Module = await getModule();
  const r = Module.evalFormula('=1/0');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, VAL.ERROR);
});

test('evalFormula uses read-only evaluation for the intersection operator', async () => {
  const Module = await getModule();
  const r = Module.evalFormula('=A1 B1');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, VAL.ERROR);
  assert.equal(r.value.errorCode, 0); // ErrorCode::Null / #NULL!
});

test('Workbook.createDefault produces a valid single-sheet workbook', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(wb.isValid());
    assert.equal(wb.sheetCount(), 1);
  } finally {
    wb.delete();
  }
});

test('print settings: raw XML round-trips and rejects malformed fragments', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    // An absent element reads back as the empty string, never null.
    assert.equal(wb.getSheetPageSetupXml(0).xml, '');

    const fragment = '<pageSetup paperSize="9" orientation="portrait" scale="85"/>';
    assert.ok(wb.setSheetPageSetupXml(0, fragment).ok);
    assert.equal(wb.getSheetPageSetupXml(0).xml, fragment);

    for (const bad of [
      '<pageSetup/><pageSetup/>',
      '<pageMargins left="1"/>',
      '<pageSetup orientation="portrait"',
      '<x:pageSetup/>',
    ]) {
      assert.equal(wb.setSheetPageSetupXml(0, bad).ok, false, `expected rejection for ${bad}`);
    }
    // A rejected set leaves the stored fragment untouched.
    assert.equal(wb.getSheetPageSetupXml(0).xml, fragment);

    assert.ok(wb.setSheetPageSetupXml(0, '').ok);
    assert.equal(wb.getSheetPageSetupXml(0).xml, '');
    const cleared = wb.getSheetPageSetup(0);
    assert.ok(cleared.status.ok);
    assert.equal(cleared.scale, 100);
    assert.equal(cleared.scaleStated, false);
  } finally {
    wb.delete();
  }
});

test('print settings: typed patch leaves unstated attributes alone', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(
      wb.setSheetPageSetupXml(0, '<pageSetup paperSize="9" orientation="portrait" horizontalDpi="600" copies="3"/>').ok,
    );
    assert.ok(wb.setSheetPageSetup(0, { orientation: 2 }).ok);

    const xml = wb.getSheetPageSetupXml(0).xml;
    assert.match(xml, /orientation="landscape"/);
    // Attributes the engine does not model must survive the patch, or the
    // "open a template and change one thing" route loses data.
    assert.match(xml, /horizontalDpi="600"/);
    assert.match(xml, /copies="3"/);
    assert.match(xml, /paperSize="9"/);

    // A print scale outside Excel's range is rejected, not clamped.
    assert.equal(wb.setSheetPageSetup(0, { scale: 9 }).ok, false);
    assert.equal(wb.setSheetPageSetup(0, { scale: 401 }).ok, false);
    assert.ok(wb.setSheetPageSetup(0, { scale: 400 }).ok);
  } finally {
    wb.delete();
  }
});

test('print settings: a change reaches paginate with no save cycle', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    for (let row = 0; row < 200; row += 1) {
      for (let col = 0; col < 20; col += 1) {
        assert.ok(wb.setNumber(0, row, col, 1).ok);
      }
    }
    const portrait = wb.paginate(0).pageCount;
    assert.ok(wb.setSheetPageSetup(0, { orientation: 2 }).ok);
    assert.notEqual(wb.paginate(0).pageCount, portrait);
  } finally {
    wb.delete();
  }
});

test('print settings: an authored report survives save and load', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  let bytes;
  try {
    assert.ok(wb.setText(0, 0, 0, '売上').ok);
    assert.ok(
      wb.setSheetPageSetup(0, {
        paperSize: 9,
        orientation: 1,
        fitToPage: true,
        fitToWidth: 1,
        fitToHeight: 0,
      }).ok,
    );
    assert.ok(wb.setSheetPageMargins(0, { left: 0.5, right: 0.5, top: 0.8, bottom: 0.8 }).ok);
    assert.ok(wb.setSheetPrintOptions(0, { horizontalCentered: true }).ok);
    assert.ok(wb.setSheetHeaderFooter(0, { oddHeader: '&C月次報告', oddFooter: '&R&P / &N' }).ok);
    assert.ok(wb.setSheetPrintArea(0, 'A1:F80').ok);
    assert.ok(wb.setSheetPrintTitles(0, '1:2', '').ok);
    assert.ok(wb.addSheetRowBreak(0, 39, true).ok);

    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    bytes = saved.bytes;
  } finally {
    wb.delete();
  }

  const reloaded = Module.Workbook.loadBytes(bytes);
  try {
    const setup = reloaded.getSheetPageSetup(0);
    assert.ok(setup.status.ok);
    assert.equal(setup.paperSize, 9);
    assert.equal(setup.orientation, 1);
    assert.equal(setup.fitToPage, true);
    assert.equal(setup.fitToHeight, 0);
    assert.equal(reloaded.getSheetPageMargins(0).left, 0.5);
    assert.equal(reloaded.getSheetPrintArea(0).ranges, 'A1:F80');
    const titles = reloaded.getSheetPrintTitles(0);
    assert.equal(titles.repeatRows, '1:2');
    assert.equal(titles.repeatCols, '');
    // The header code is stored escaped and decodes back to `&C...`.
    assert.match(reloaded.getSheetHeaderFooterXml(0).xml, /<oddHeader>&amp;C月次報告<\/oddHeader>/);
    const breaks = reloaded.getSheetRowBreaks(0);
    assert.deepEqual(
      breaks.breaks.map((b) => b.id),
      [39],
    );
    assert.equal(breaks.breaks[0].manual, true);
  } finally {
    reloaded.delete();
  }
});

test('setRangeXfIndex applies one xf across a rectangle', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    const xf = wb.addXf({ fontIndex: 0, fillIndex: 0, borderIndex: 0, numFmtId: 0, wrapText: true, hasWrapText: true });
    assert.ok(xf.status.ok, JSON.stringify(xf.status));
    assert.ok(wb.setRangeXfIndex(0, 0, 0, 2, 2, xf.index).ok);
    // Cells that held nothing are materialised so the ruled box renders.
    for (const [row, col] of [
      [0, 0],
      [1, 1],
      [2, 2],
    ]) {
      assert.equal(wb.getCellXfIndex(0, row, col).xfIndex, xf.index);
    }
    assert.equal(wb.getCellXfIndex(0, 3, 3).xfIndex, 0);
  } finally {
    wb.delete();
  }
});

test('sheet layout getters preserve explicit width presence flags', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(wb.setColumnWidth(0, 0, 0, 0).ok);
    assert.ok(wb.setColumnHidden(0, 1, 1, true).ok);
    assert.ok(wb.setRowHeight(0, 3, 30).ok);

    const columns = wb.getSheetColumns(0);
    assert.ok(columns.status.ok, `getSheetColumns: ${JSON.stringify(columns.status)}`);
    const widthZero = columns.columns.find((column) => column.first === 0 && column.last === 0);
    assert.ok(widthZero);
    assert.equal(widthZero.width, 0);
    assert.equal(typeof widthZero.hasWidth, 'number');
    assert.equal(widthZero.hasWidth, 1);
    assert.equal(typeof widthZero.hasStyle, 'number');
    assert.equal(widthZero.hasStyle, 0);
    const hidden = columns.columns.find((column) => column.first === 1 && column.last === 1);
    assert.ok(hidden);
    assert.equal(hidden.hidden, 1);
    assert.equal(typeof hidden.hasWidth, 'number');
    assert.equal(hidden.hasWidth, 0);

    const rows = wb.getSheetRowOverrides(0);
    assert.ok(rows.status.ok, `getSheetRowOverrides: ${JSON.stringify(rows.status)}`);
    const row = rows.rows.find((entry) => entry.row === 3);
    assert.ok(row);
    assert.equal(row.height, 30);
    assert.equal(row.hasStyle, 0);
    assert.equal(row.styleXf, 0);
  } finally {
    wb.delete();
  }
});

test('Unicode sheet identity rejects folded duplicate and permits casing rename', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createEmpty();
  try {
    assert.ok(wb.addSheet('Ä').ok);
    assert.ok(wb.addSheet('Ö').ok);
    assert.equal(wb.addSheet('ä').ok, false);
    const collision = wb.renameSheet(1, 'ä');
    assert.equal(collision.ok, false);
    assert.notEqual(collision.status, 0);
    assert.equal(wb.sheetName(0).value, 'Ä');
    assert.equal(wb.sheetName(1).value, 'Ö');
    assert.ok(wb.renameSheet(0, 'ä').ok);
    const sn = wb.sheetName(0);
    assert.ok(sn.status.ok);
    assert.equal(sn.value, 'ä');
    assert.equal(wb.sheetName(1).value, 'Ö');
  } finally {
    wb.delete();
  }
});

test('setNumber + getValue round-trips bit-exactly', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    const x = 3.141592653589793;
    assert.ok(wb.setNumber(0, 0, 0, x).ok);
    assert.ok(wb.recalc().ok);
    const r = wb.getValue(0, 0, 0);
    assert.ok(r.status.ok);
    assert.equal(r.value.kind, VAL.NUMBER);
    assert.equal(r.value.number, x);
  } finally {
    wb.delete();
  }
});

test('setFormula + recalc computes 41 + 1 == 42 in B1', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    assert.ok(wb.setNumber(0, 0, 0, 41).ok);
    assert.ok(wb.setFormula(0, 0, 1, '=A1+1').ok);
    assert.ok(wb.recalc().ok);
    const b1 = wb.getValue(0, 0, 1);
    assert.ok(b1.status.ok);
    assert.equal(b1.value.kind, VAL.NUMBER);
    assert.equal(b1.value.number, 42);
  } finally {
    wb.delete();
  }
});

test('Workbook.recalcParallel evaluates a wide DAG and reports bounded telemetry', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  const branchCount = 32;
  try {
    assert.ok(wb.setNumber(0, 0, 0, 1).ok);
    for (let i = 0; i < branchCount; i += 1) {
      const row = i + 1;
      assert.ok(wb.setFormula(0, row, 1, `=A1+${i + 2}`).ok);
      assert.ok(wb.setFormula(0, row, 2, `=B${row + 1}*2`).ok);
    }

    const parallel = wb.recalcParallel(4);
    assert.ok(parallel.status.ok, `status=${JSON.stringify(parallel.status)}`);
    assert.equal(typeof parallel.stats.cellsEvaluated, 'number');
    assert.equal(typeof parallel.stats.sccsProcessed, 'number');
    assert.equal(typeof parallel.stats.parallelSteps, 'number');
    assert.ok(parallel.stats.cellsEvaluated > 0);
    assert.ok(parallel.stats.sccsProcessed > 0);
    if (parallel.stats.workerThreadsStarted <= 1) {
      // OS launch refusal or a partial launch of one worker is a documented
      // successful serial degradation.
      assert.equal(parallel.stats.workerThreadsUsed, 0);
      assert.equal(parallel.stats.parallelSteps, 0);
      assert.ok(parallel.stats.serialFallbackSteps > 0, `stats=${JSON.stringify(parallel.stats)}`);
    } else {
      assert.ok(parallel.stats.parallelSteps > 0, `stats=${JSON.stringify(parallel.stats)}`);
      assert.ok(parallel.stats.workerThreadsStarted >= 2);
      assert.ok(parallel.stats.workerThreadsStarted <= 4);
      assert.ok(parallel.stats.workerThreadsUsed > 0);
      assert.ok(parallel.stats.workerThreadsUsed <= parallel.stats.workerThreadsStarted);
    }

    const first = wb.getValue(0, 1, 1);
    const last = wb.getValue(0, branchCount, 2);
    assert.ok(first.status.ok);
    assert.ok(last.status.ok);
    assert.equal(first.value.kind, VAL.NUMBER);
    assert.equal(last.value.kind, VAL.NUMBER);
    assert.equal(first.value.number, 3);
    assert.equal(last.value.number, (1 + branchCount + 1) * 2);

    assert.ok(wb.setNumber(0, 0, 0, 5).ok);
    const callerOnly = wb.recalcParallel(1);
    assert.ok(callerOnly.status.ok, `status=${JSON.stringify(callerOnly.status)}`);
    assert.equal(callerOnly.stats.parallelSteps, 0);
    assert.ok(callerOnly.stats.serialFallbackSteps > 0);
    assert.equal(callerOnly.stats.workerThreadsStarted, 0);
    assert.equal(callerOnly.stats.workerThreadsUsed, 0);
    assert.equal(wb.getValue(0, branchCount, 2).value.number, (5 + branchCount + 1) * 2);

    const invalid = wb.recalcParallel(9);
    assert.equal(invalid.status.ok, false);
    assert.notEqual(invalid.status.status, 0);
    assert.equal(invalid.stats.cellsEvaluated, 0);
    assert.equal(invalid.stats.sccsProcessed, 0);
    assert.equal(invalid.stats.parallelSteps, 0);
    assert.equal(invalid.stats.serialFallbackSteps, 0);
    assert.equal(invalid.stats.cycleRecoveries, 0);
    assert.equal(invalid.stats.workerThreadsStarted, 0);
    assert.equal(invalid.stats.workerThreadsUsed, 0);

    for (const invalidThreadCount of [
      -0.5,
      1.5,
      Number.NaN,
      Number.POSITIVE_INFINITY,
      Number.NEGATIVE_INFINITY,
      2 ** 32 + 1,
      null,
      undefined,
    ]) {
      const invalidShape = wb.recalcParallel(invalidThreadCount);
      assert.equal(invalidShape.status.ok, false, `threadCount=${String(invalidThreadCount)}`);
      assert.notEqual(invalidShape.status.status, 0);
      assert.equal(invalidShape.stats.cellsEvaluated, 0);
      assert.equal(invalidShape.stats.sccsProcessed, 0);
      assert.equal(invalidShape.stats.parallelSteps, 0);
      assert.equal(invalidShape.stats.serialFallbackSteps, 0);
      assert.equal(invalidShape.stats.cycleRecoveries, 0);
      assert.equal(invalidShape.stats.workerThreadsStarted, 0);
      assert.equal(invalidShape.stats.workerThreadsUsed, 0);
    }
  } finally {
    wb.delete();
  }
});

// Status ordinals the callback contract is stated in terms of; see
// `src/utils/error.h`. Spelled out here because the binding does not
// export a status enum.
const STATUS = Object.freeze({
  RECALC_REENTRANT: 4005,
  CALLBACK_EXCEPTION: 7003,
});

// A self-referential average: from a blank A1 the iterative solver walks
// towards 10, and the progress callback fires after every sweep.
function makeIterativeWorkbook(Module) {
  const wb = Module.Workbook.createDefault();
  assert.ok(wb.setIterative(true, 100, 0.0001).ok);
  assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
  return wb;
}

test('a throwing iterative progress callback aborts the solve and reports a status', async () => {
  const Module = await getModule();
  const wb = makeIterativeWorkbook(Module);
  try {
    let calls = 0;
    assert.ok(
      wb.setIterativeProgress(() => {
        calls += 1;
        throw new Error('callback failure');
      }).ok,
    );

    // The binding never lets an exception cross back into WASM, where no
    // landing pad exists to restore the engine's RAII state: the throw
    // comes back as a status on the same envelope every other call uses.
    const aborted = wb.recalc();
    assert.equal(aborted.ok, false, `expected a failure envelope: ${JSON.stringify(aborted)}`);
    assert.equal(aborted.status, STATUS.CALLBACK_EXCEPTION);
    assert.equal(calls, 1, 'the solve must stop at the first throw');

    // The workbook is still readable, and the recalc re-entry guard was
    // released on the way out.
    const mid = wb.getValue(0, 0, 0);
    assert.ok(mid.status.ok, JSON.stringify(mid.status));
    assert.equal(mid.value.kind, VAL.NUMBER);

    const second = wb.recalc();
    assert.notEqual(second.status, STATUS.RECALC_REENTRANT, 'the re-entry flag stayed set');

    assert.ok(wb.setIterativeProgress(null).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
    const recovered = wb.recalc();
    assert.ok(recovered.ok, `recalc after the throwing callback: ${JSON.stringify(recovered)}`);
    const converged = wb.getValue(0, 0, 0);
    assert.ok(converged.status.ok, JSON.stringify(converged.status));
    assert.equal(converged.value.kind, VAL.NUMBER);
    assert.ok(Math.abs(converged.value.number - 10) < 0.001, `A1=${converged.value.number}`);
  } finally {
    wb.setIterativeProgress(null);
    wb.delete();
  }
});

test('an iterative progress callback returning a non-boolean is read by truthiness', async () => {
  const Module = await getModule();
  for (const [returned, expectedSweeps] of [
    ['keep going', 0],
    [0, 1],
    [undefined, 0],
  ]) {
    const wb = makeIterativeWorkbook(Module);
    try {
      let calls = 0;
      assert.ok(
        wb.setIterativeProgress(() => {
          calls += 1;
          return returned;
        }).ok,
      );
      assert.ok(wb.recalc().ok);
      // A falsy return aborts after the first sweep; anything else lets
      // the solve run to convergence.
      if (expectedSweeps === 1) {
        assert.equal(calls, 1, `returned=${String(returned)}`);
      } else {
        assert.ok(calls > 1, `returned=${String(returned)}, calls=${calls}`);
      }
    } finally {
      wb.setIterativeProgress(null);
      wb.delete();
    }
  }
});

// ---- Result-envelope key sets ------------------------------------------
//
// A result type whose payload key is declared non-optional must carry that
// key on every exit path, or `r.name.toUpperCase()` type-checks and then
// throws at runtime for a caller who skipped the status check. The probe
// table below is checked for completeness against the declaration file, so
// a new result type cannot be added without a failure-path probe.

function interfaceBody(dts, name) {
  const m = dts.match(new RegExp(`^export interface ${name}\\b[^{]*\\{`, 'm'));
  if (!m) return null;
  let depth = 0;
  const open = dts.indexOf('{', m.index);
  for (let i = open; i < dts.length; i += 1) {
    if (dts[i] === '{') depth += 1;
    else if (dts[i] === '}') {
      depth -= 1;
      if (depth === 0) return dts.slice(open + 1, i);
    }
  }
  return null;
}

function interfaceBases(dts, name) {
  const m = dts.match(new RegExp(`^export interface ${name}\\b([^{]*)\\{`, 'm'));
  if (!m) return [];
  const clause = m[1].trim();
  if (!clause.startsWith('extends')) return [];
  return clause
    .slice('extends'.length)
    .split(',')
    .map((s) => s.trim())
    .filter(Boolean);
}

// [name, required] for every member, following `extends`.
function interfaceMembers(dts, name, seen = new Set()) {
  if (seen.has(name)) return [];
  seen.add(name);
  const body = interfaceBody(dts, name);
  if (body === null) return [];
  const out = [];
  for (const base of interfaceBases(dts, name)) out.push(...interfaceMembers(dts, base, seen));
  for (const m of body.matchAll(/^ {2}(?:readonly\s+)?([A-Za-z_$][\w$]*)(\??)\s*:/gm)) {
    out.push([m[1], m[2] === '']);
  }
  return out;
}

// Every declared result type that carries a `status` plus at least one
// required payload key. `Status` itself is the envelope, not a result.
function resultTypesWithRequiredPayload(dts) {
  const out = new Map();
  for (const m of dts.matchAll(/^export interface ([A-Za-z0-9_]+)/gm)) {
    const name = m[1];
    if (name === 'Status') continue;
    const members = interfaceMembers(dts, name);
    if (!members.some(([key]) => key === 'status')) continue;
    const required = members.filter(([key, req]) => req && key !== 'status').map(([key]) => key);
    if (required.length > 0) out.set(name, required);
  }
  return out;
}

// One failing (or, where no failure is reachable, succeeding) call per
// result type. `fails: false` marks a call whose only failure path is a
// destroyed handle, which `delete()` makes unusable to probe.
function envelopeProbes(wb) {
  return [
    ['ParallelRecalcResult', true, () => wb.recalcParallel(9)],
    ['CellResult', true, () => wb.getValue(99, 0, 0)],
    ['EvalResult', true, () => wb.evaluateFormulaText(99, 0, 0, '=1')],
    ['EvalArrayResult', true, () => wb.evaluateFormulaArray(99, 0, 0, '=1')],
    ['SaveResult', true, () => wb.saveAs(99)],
    ['SaveDiagnosticsResult', true, () => wb.saveWithDiagnostics(99)],
    ['ReadDiagnosticsResult', false, () => wb.readDiagnostics()],
    ['StringResult', true, () => wb.sheetName(99)],
    ['PivotLayoutResult', true, () => wb.pivotLayout(99, 0)],
    ['PivotWorksheetSourceResult', true, () => wb.pivotCacheGetWorksheetSource(9999)],
    ['PivotReportLayoutResult', true, () => wb.pivotGetLayout(99, 0)],
    ['PivotFilterResult', true, () => wb.pivotFilterAt(99, 0, 0)],
    ['CfRangeResult', true, () => wb.evaluateCfRange(99, 0, 0, 1, 1, Number.NaN)],
    ['PaginationResult', true, () => wb.paginate(99)],
    ['SheetViewResult', true, () => wb.getSheetView(99)],
    ['SheetProtectionResult', true, () => wb.getSheetProtection(99)],
    ['ColumnsResult', true, () => wb.getSheetColumns(99)],
    ['RowsResult', true, () => wb.getSheetRowOverrides(99)],
    ['CommentResult', true, () => wb.getCommentResult(99, 0, 0)],
    ['CellXfIndexResult', true, () => wb.getCellXfIndex(99, 0, 0)],
    ['CellXfResult', true, () => wb.getCellXf(9999)],
    ['FontResult', true, () => wb.getFont(9999)],
    ['FillResult', true, () => wb.getFill(9999)],
    ['BorderResult', true, () => wb.getBorder(9999)],
    ['NumFmtResult', true, () => wb.getNumFmt(59999)],
    ['LambdaTextResult', true, () => wb.getLambdaText(99, 0, 0)],
    ['CellStyleResult', true, () => wb.getCellStyle(9999)],
    ['AddStyleResult', true, () => wb.addXf({ fontIndex: 9999 })],
    ['AddNumFmtResult', false, () => wb.addNumFmt('0.00')],
    [
      'PartialRecalcResult',
      true,
      () => wb.partialRecalc({ sheet: 99, firstRow: 0, lastRow: 1, firstCol: 0, lastCol: 1 }),
    ],
    ['SheetAutoFilterXmlResult', true, () => wb.getSheetAutoFilterXml(99)],
    ['SheetPrintXmlResult', true, () => wb.getSheetPageSetupXml(99)],
    ['SheetPrintAreaResult', true, () => wb.getSheetPrintArea(99)],
    ['SheetPrintTitlesResult', true, () => wb.getSheetPrintTitles(99)],
    // `getSheetRowBreaks` / `getSheetColBreaks` answer an out-of-range
    // sheet with an empty list rather than an error, unlike every sibling
    // print getter, so no failure path is reachable from here.
    ['SheetPageBreaksResult', false, () => wb.getSheetRowBreaks(99)],
    ['SheetPageSetupResult', true, () => wb.getSheetPageSetup(99)],
    ['SheetPageMarginsResult', true, () => wb.getSheetPageMargins(99)],
  ];
}

test('every declared result type is probed for its key set', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
  const declared = resultTypesWithRequiredPayload(dts);
  const probed = new Set(envelopeProbes({}).map(([name]) => name));
  const unprobed = [...declared.keys()].filter((name) => !probed.has(name));
  assert.deepEqual(unprobed, [], 'result types with no failure-path probe');
  const stale = [...probed].filter((name) => !declared.has(name));
  assert.deepEqual(stale, [], 'probes for result types that no longer declare a required payload');
});

test('result envelopes keep their declared keys on failure paths', async () => {
  const Module = await getModule();
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
  const declared = resultTypesWithRequiredPayload(dts);
  const wb = Module.Workbook.createDefault();
  try {
    for (const [typeName, expectFailure, call] of envelopeProbes(wb)) {
      const r = call();
      if (expectFailure) {
        assert.equal(r.status.ok, false, `${typeName}: probe was expected to fail`);
      }
      const missing = declared.get(typeName).filter((key) => !(key in r));
      assert.deepEqual(missing, [], `${typeName} dropped declared keys`);
    }
  } finally {
    wb.delete();
  }
});

test('iterative progress callbacks are per Workbook, not per module', async () => {
  const Module = await getModule();
  const first = makeIterativeWorkbook(Module);
  const second = makeIterativeWorkbook(Module);
  try {
    let firstCalls = 0;
    let secondCalls = 0;
    assert.ok(
      first.setIterativeProgress(() => {
        firstCalls += 1;
        return true;
      }).ok,
    );
    // Installing on `second` must not displace the registration on `first`.
    assert.ok(
      second.setIterativeProgress(() => {
        secondCalls += 1;
        return true;
      }).ok,
    );

    assert.ok(first.recalc().ok);
    assert.ok(firstCalls > 0, 'the first workbook kept its own callback');
    assert.equal(secondCalls, 0);

    firstCalls = 0;
    assert.ok(second.recalc().ok);
    assert.ok(secondCalls > 0);
    assert.equal(firstCalls, 0);
  } finally {
    first.setIterativeProgress(null);
    second.setIterativeProgress(null);
    first.delete();
    second.delete();
  }
});
