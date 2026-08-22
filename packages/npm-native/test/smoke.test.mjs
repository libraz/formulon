//
// Smoke tests for the staged @libraz/formulon-native npm package.
//
// These tests intentionally import the *staged* dist/ artefacts (resolved
// via the package's own package.json "main" field), not the raw build/
// output. That way we catch staging mistakes -- missing files, wrong
// relative paths, broken "exports" entries -- before the tarball ships.
//
// Run via `make node-test` (which runs `make node-package` first).

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const pkgRoot = path.resolve(__dirname, '..');
const pkgJsonPath = path.join(pkgRoot, 'package.json');

// Resolve the staged entry point through the package's own package.json
// "main" field. Fails (rather than skips) if the package isn't staged --
// that's the whole point of running these tests against dist/.
async function loadStagedModule() {
  const raw = await readFile(pkgJsonPath, 'utf8');
  const pkg = JSON.parse(raw);
  const main = pkg.main;
  if (!main) {
    throw new Error(`package.json is missing "main": ${pkgJsonPath}`);
  }
  const mainPath = path.resolve(pkgRoot, main);
  // Node ESM requires file:// URLs for absolute paths on Windows;
  // pathToFileURL is the portable form.
  return import(pathToFileURL(mainPath).href);
}

let modPromise;
function getModule() {
  if (!modPromise) {
    modPromise = loadStagedModule();
  }
  return modPromise;
}

function appendEmptyZipEntry(bytes, name) {
  const input = new Uint8Array(bytes);
  const view = new DataView(input.buffer, input.byteOffset, input.byteLength);
  let eocd = input.length - 22;
  while (eocd >= 0 && view.getUint32(eocd, true) !== 0x06054b50) eocd -= 1;
  assert.ok(eocd >= 0, 'missing ZIP end record');
  const count = view.getUint16(eocd + 10, true);
  const centralSize = view.getUint32(eocd + 12, true);
  const centralOffset = view.getUint32(eocd + 16, true);
  const encodedName = new TextEncoder().encode(name);
  const u16 = (out, n) => out.push(n & 0xff, (n >>> 8) & 0xff);
  const u32 = (out, n) => out.push(n & 0xff, (n >>> 8) & 0xff, (n >>> 16) & 0xff, (n >>> 24) & 0xff);
  const bytesTo = (out, source) => {
    source.forEach((b) => {
      out.push(b);
    });
  };
  const local = [];
  u32(local, 0x04034b50);
  u16(local, 20);
  u16(local, 0);
  u16(local, 0);
  u16(local, 0);
  u16(local, 0);
  u32(local, 0);
  u32(local, 0);
  u32(local, 0);
  u16(local, encodedName.length);
  u16(local, 0);
  bytesTo(local, encodedName);
  const central = [];
  u32(central, 0x02014b50);
  u16(central, 20);
  u16(central, 20);
  u16(central, 0);
  u16(central, 0);
  u16(central, 0);
  u16(central, 0);
  u32(central, 0);
  u32(central, 0);
  u32(central, 0);
  u16(central, encodedName.length);
  u16(central, 0);
  u16(central, 0);
  u16(central, 0);
  u16(central, 0);
  u32(central, 0);
  u32(central, centralOffset);
  bytesTo(central, encodedName);
  const out = [];
  bytesTo(out, input.slice(0, centralOffset));
  bytesTo(out, local);
  bytesTo(out, input.slice(centralOffset, centralOffset + centralSize));
  bytesTo(out, central);
  u32(out, 0x06054b50);
  u16(out, 0);
  u16(out, 0);
  u16(out, count + 1);
  u16(out, count + 1);
  u32(out, centralSize + central.length);
  u32(out, centralOffset + local.length);
  u16(out, 0);
  return Uint8Array.from(out);
}

test('default export exposes Workbook + evalFormula + version', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.Workbook, 'function');
  assert.equal(typeof mod.evalFormula, 'function');
  assert.equal(typeof mod.version, 'function');
  assert.equal(typeof mod.statusString, 'function');
  assert.equal(typeof mod.ValueKind, 'object');
  assert.equal(typeof mod.PivotCellKind, 'object');
});

test('all declared runtime constants resolve from the staged ESM entry point', async () => {
  const mod = await getModule();
  const declarations = await readFile(path.join(pkgRoot, 'index.d.ts'), 'utf8');
  const names = [...declarations.matchAll(/^export const (\w+)/gm)].map((match) => match[1]);
  assert.ok(names.length > 0, 'expected runtime constant declarations');
  for (const name of names) {
    assert.ok(name in mod, `missing runtime export declared in index.d.ts: ${name}`);
  }
  assert.equal(mod.PivotAxis.Row, 0);
  assert.equal(mod.PIVOT_SHOW_AS_BASE_PREVIOUS, 1048828);
  // Named constants a consumer needs in order to interpret shared record
  // fields (`Value.errorCode`, `ExternalLinkRecord.kind`) or to drive the
  // calc policy. The WASM package exports the same names and ordinals;
  // `check_binding_drift dts-enums` holds the two in step.
  assert.equal(mod.ErrorCode.Div0, 1);
  assert.equal(mod.ErrorCode.Unknown, 16);
  assert.equal(mod.CalcMode.Manual, 1);
  assert.equal(mod.ExternalLinkKind.Dde, 3);
  for (const name of ['ErrorCode', 'CalcMode', 'ExternalLinkKind']) {
    assert.ok(Object.isFrozen(mod[name]), `${name} must be frozen`);
  }
});

test('default export type and runtime default export declare the same keys', async () => {
  const mod = await getModule();
  const declarations = await readFile(path.join(pkgRoot, 'index.d.ts'), 'utf8');
  const typeBlock = declarations.match(/declare const _default: \{([\s\S]*?)\n\};/);
  assert.ok(typeBlock, 'expected a `declare const _default: {...}` block in index.d.ts');
  const declared = new Set([...typeBlock[1].matchAll(/^\s{2}(\w+):/gm)].map((match) => match[1]));
  const runtime = new Set(Object.keys(mod.default));
  const missingFromType = [...runtime].filter((key) => !declared.has(key));
  const missingFromRuntime = [...declared].filter((key) => !runtime.has(key));
  assert.deepEqual(missingFromType, [], 'runtime default export keys absent from the _default type');
  assert.deepEqual(missingFromRuntime, [], '_default type keys absent from the runtime default export');
  assert.equal(mod.default.PivotAxis.Row, 0);
  assert.equal(typeof mod.default.errorDisplayName, 'function');
});

test('every named export is reachable through the default export', async () => {
  const mod = await getModule();
  const named = Object.keys(mod).filter((key) => key !== 'default');
  const missing = named.filter((key) => !(key in mod.default));
  assert.deepEqual(missing, [], 'named exports absent from the default export object');
});

test('setLogMinLevel is a module-level control and validates its ordinal', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.setLogMinLevel, 'function');
  assert.equal(typeof mod.setLogSink, 'function');
  // Process-wide state, so it must not be a Workbook method.
  const wb = mod.Workbook.createDefault();
  assert.equal(typeof wb.setLogMinLevel, 'undefined');
  wb.dispose();

  for (const level of Object.values(mod.LogLevel)) {
    assert.ok(mod.setLogMinLevel(level).ok, `level=${level}`);
  }
  for (const bad of [-1, 5, 99]) {
    const r = mod.setLogMinLevel(bad);
    assert.equal(r.ok, false, `level=${bad} should be rejected`);
    assert.equal(r.status, 2);
  }
  assert.ok(mod.setLogMinLevel(mod.LogLevel.Off).ok);
});

test('setLogSink accepts a function and null, and rejects a non-function', async () => {
  const mod = await getModule();
  assert.ok(mod.setLogSink(() => {}).ok);
  assert.ok(mod.setLogSink(null).ok);
  assert.ok(mod.setLogSink().ok);
  const bad = mod.setLogSink(42);
  assert.equal(bad.ok, false);
  assert.equal(bad.status, 7001);
  assert.ok(mod.setLogSink(null).ok);
});

// Sink delivery goes through a thread-safe function, so records land on a
// later turn of the libuv loop rather than inside the native call. Poll
// with a bounded budget instead of guessing a single turn.
const SINK_SETTLE_MS = 500;

async function waitForRecords(records, wanted) {
  const deadline = Date.now() + SINK_SETTLE_MS;
  while (records.length < wanted && Date.now() < deadline) {
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  return records;
}

test('a registered sink receives the record as raw bytes', async () => {
  const mod = await getModule();
  const records = [];
  assert.ok(mod.setLogSink((record) => records.push(record)).ok);
  assert.ok(mod.setLogMinLevel(mod.LogLevel.Warn).ok);
  try {
    const wb = mod.Workbook.createDefault();
    // An implicit-intersection formula the XLSB encoder cannot lower,
    // which makes the writer emit a per-cell warn record.
    assert.ok(wb.setFormula(0, 0, 0, '=@A1:A10').ok);
    assert.ok(wb.recalc().ok);
    assert.ok(wb.saveAs(mod.WorkbookFormat.Xlsb).status.ok);
    wb.dispose();
    await waitForRecords(records, 1);
    assert.ok(records.length > 0, 'expected at least one record');
    for (const record of records) {
      assert.ok(ArrayBuffer.isView(record), 'record must be a byte view');
    }
    const text = records.map((r) => new TextDecoder().decode(r)).join('');
    assert.match(text, /xlsb\.writer\.formula_downgraded/);
    assert.match(text, /"level":"warn"/);
  } finally {
    assert.ok(mod.setLogMinLevel(mod.LogLevel.Off).ok);
    assert.ok(mod.setLogSink(null).ok);
  }
});

test('the default threshold delivers nothing to a registered sink', async () => {
  const mod = await getModule();
  const records = [];
  assert.ok(mod.setLogSink((record) => records.push(record)).ok);
  try {
    const wb = mod.Workbook.createDefault();
    assert.ok(wb.setFormula(0, 0, 0, '=@A1:A10').ok);
    assert.ok(wb.recalc().ok);
    assert.ok(wb.saveAs(mod.WorkbookFormat.Xlsb).status.ok);
    wb.dispose();
    // Wait the same budget the positive case needs, so silence here is a
    // result rather than an artefact of not having waited long enough.
    await new Promise((resolve) => setTimeout(resolve, SINK_SETTLE_MS));
    assert.deepEqual(records, []);
  } finally {
    assert.ok(mod.setLogSink(null).ok);
  }
});

test('memoryUsage and the Python wheel agree that the estimate grows', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const empty = wb.memoryUsage();
  for (let row = 0; row < 2000; row += 1) {
    wb.setNumber(0, row, 0, row);
  }
  assert.ok(wb.memoryUsage() > empty);
  wb.dispose();
});

// Parses the staged `export const X: Readonly<{ A: 0; ... }>` declarations
// into {name: {member: ordinal}}. Name presence alone is not enough: a
// staged bundle carrying a stale ordinal ships a constant that silently
// means something else, and nothing upstream of dist/ can see that.
function declaredOrdinalTables(dts) {
  const tables = {};
  const blocks = dts.matchAll(/^export const (\w+): Readonly<\{([^}]*)\}>;/gm);
  for (const [, name, body] of blocks) {
    const members = {};
    for (const [, member, value] of body.matchAll(/(\w+):\s*(-?\d+)/g)) {
      members[member] = Number(value);
    }
    tables[name] = members;
  }
  return tables;
}

test('staged constant tables carry the ordinals the staged .d.ts declares', async () => {
  const mod = await getModule();
  const dts = await readFile(path.join(pkgRoot, 'dist', 'index.d.ts'), 'utf8');
  const tables = declaredOrdinalTables(dts);
  assert.ok(Object.keys(tables).length > 0, 'expected declared ordinal tables');
  for (const [name, members] of Object.entries(tables)) {
    const runtime = mod[name];
    assert.ok(runtime, `missing runtime table ${name}`);
    assert.deepEqual({ ...runtime }, members, `${name} ordinals differ from dist/index.d.ts`);
    assert.ok(Object.isFrozen(runtime), `${name} must be frozen`);
  }
});

test('version() returns a non-empty string', async () => {
  const mod = await getModule();
  const v = mod.version();
  assert.equal(typeof v, 'string');
  assert.ok(v.length > 0, `expected non-empty version, got ${JSON.stringify(v)}`);
});

test('errorDisplayName returns Excel literals', async () => {
  const mod = await getModule();
  assert.equal(mod.errorDisplayName(1), '#DIV/0!');
  assert.equal(mod.errorDisplayName(999), '#UNKNOWN!');
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
  const mod = await getModule();
  for (const spec of ONE_SHOT_CASES) {
    const r = mod.evalFormula(spec.formula);
    assert.ok(r.status.ok, `${spec.formula}: ${JSON.stringify(r.status)}`);
    assert.equal(r.value.kind, spec.kind, spec.formula);
    if (spec.number !== undefined) {
      assert.equal(r.value.number, spec.number, spec.formula);
    }
    if (spec.boolean !== undefined) {
      assert.equal(r.value.boolean, spec.boolean, spec.formula);
    }
    // Writing into the anchor and recalcing would make these #REF!.
    assert.equal(r.value.errorCode, 0, spec.formula);
  }
});

test("evalFormula('=1+2') returns NUMBER 3", async () => {
  const mod = await getModule();
  const r = mod.evalFormula('=1+2');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 3);
});

test("evalFormula('=#REF!') surfaces ERROR value (not failed status)", async () => {
  const mod = await getModule();
  const r = mod.evalFormula('=#REF!');
  assert.ok(r.status.ok, `status=${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Error);
});

test('Workbook.createDefault + setNumber + getValue round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.sheetCount(), 1);
  const setRc = wb.setNumber(0, 0, 0, 42);
  assert.ok(setRc.ok, `setNumber: ${JSON.stringify(setRc)}`);
  const r = wb.getValue(0, 0, 0);
  assert.ok(r.status.ok, `getValue: ${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 42);
});

test('paginate exposes a status envelope and page geometry', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Asserts the envelope's shape, not the page geometry: how many rows fit a
  // page is the print engine's business and is pinned by the native print
  // suite against Excel. Row 400 is far enough down to force a break under
  // any plausible body height.
  assert.ok(wb.setNumber(0, 0, 0, 1).ok);
  assert.ok(wb.setNumber(0, 400, 0, 2).ok);
  const result = wb.paginate(0);
  assert.ok(result.status.ok, `paginate: ${JSON.stringify(result.status)}`);
  assert.ok(result.pageCount >= 2, `pageCount: ${result.pageCount}`);
  assert.deepEqual(result.printArea, []);
  assert.ok(result.horizontalBreaks.length > 0);
  assert.deepEqual(
    result.horizontalBreaks,
    [...new Set(result.horizontalBreaks)].sort((a, b) => a - b),
  );
  assert.deepEqual(result.verticalBreaks, []);
});

test('print settings: raw XML round-trips and rejects malformed fragments', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // An absent element reads back as the empty string, never null.
  assert.equal(wb.getSheetPageSetupXml(0).xml, '');

  const fragment = '<pageSetup paperSize="9" orientation="portrait" scale="85"/>';
  assert.ok(wb.setSheetPageSetupXml(0, fragment).ok);
  assert.equal(wb.getSheetPageSetupXml(0).xml, fragment);

  // Two top-level elements, the wrong root name, and a truncated fragment
  // are all rejected at set time rather than producing a file Excel repairs.
  for (const bad of ['<pageSetup/><pageSetup/>', '<pageMargins left="1"/>', '<pageSetup orientation="portrait"']) {
    const rejected = wb.setSheetPageSetupXml(0, bad);
    assert.equal(rejected.ok, false, `expected rejection for ${bad}`);
  }
  assert.equal(wb.getSheetPageSetupXml(0).xml, fragment);

  // The empty string removes the element and restores the defaults.
  assert.ok(wb.setSheetPageSetupXml(0, '').ok);
  assert.equal(wb.getSheetPageSetupXml(0).xml, '');
  const cleared = wb.getSheetPageSetup(0);
  assert.ok(cleared.status.ok);
  assert.equal(cleared.scale, 100);
  assert.equal(cleared.scaleStated, false);
});

test('print settings: typed patch touches only the keys it states', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(
    wb.setSheetPageSetupXml(0, '<pageSetup paperSize="9" orientation="portrait" horizontalDpi="600" copies="3"/>').ok,
  );
  assert.ok(wb.setSheetPageSetup(0, { orientation: 2 }).ok);

  const xml = wb.getSheetPageSetupXml(0).xml;
  assert.match(xml, /orientation="landscape"/);
  // Attributes the engine does not model must survive a patch.
  assert.match(xml, /horizontalDpi="600"/);
  assert.match(xml, /copies="3"/);
  assert.match(xml, /paperSize="9"/);

  const read = wb.getSheetPageSetup(0);
  assert.ok(read.status.ok);
  assert.equal(read.orientation, 2);
  assert.equal(read.paperSize, 9);
  assert.equal(read.orientationStated, true);

  assert.ok(wb.setSheetPageMargins(0, { left: 0.25, top: 1.5 }).ok);
  const margins = wb.getSheetPageMargins(0);
  assert.ok(margins.status.ok);
  assert.equal(margins.left, 0.25);
  assert.equal(margins.leftStated, true);
  // An unstated margin reports the OOXML default with its flag clear.
  assert.equal(margins.right, 0.7);
  assert.equal(margins.rightStated, false);
});

test('print settings: fitToPage keeps the rest of sheetPr', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setSheetSheetPrXml(0, '<sheetPr codeName="Sheet1"><tabColor rgb="FFFF0000"/></sheetPr>').ok);
  assert.ok(wb.setSheetFitToPage(0, true).ok);
  const xml = wb.getSheetSheetPrXml(0).xml;
  assert.match(xml, /codeName="Sheet1"/);
  assert.match(xml, /<tabColor rgb="FFFF0000"\/>/);
  assert.match(xml, /fitToPage="true"/);
});

test('print settings: header/footer sections take decoded text', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setSheetHeaderFooter(0, { oddHeader: '&C\u5e33\u7968', oddFooter: '&R&P / &N' }).ok);
  // The codes reach the file escaped; a parser hands back `&C...`.
  assert.equal(
    wb.getSheetHeaderFooterXml(0).xml,
    '<headerFooter><oddHeader>&amp;C\u5e33\u7968</oddHeader><oddFooter>&amp;R&amp;P / &amp;N</oddFooter></headerFooter>',
  );

  // Omitting a key leaves that section alone; an empty string clears it.
  assert.ok(wb.setSheetHeaderFooter(0, { oddFooter: '' }).ok);
  assert.equal(
    wb.getSheetHeaderFooterXml(0).xml,
    '<headerFooter><oddHeader>&amp;C\u5e33\u7968</oddHeader></headerFooter>',
  );
});

test('print settings: print area, titles and manual breaks', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setSheetPrintArea(0, 'A1:F8').ok);
  assert.equal(wb.getSheetPrintArea(0).ranges, 'A1:F8');
  assert.equal(wb.setSheetPrintArea(0, 'not-a-range').ok, false);
  assert.equal(wb.getSheetPrintArea(0).ranges, 'A1:F8');
  assert.ok(wb.setSheetPrintArea(0, '').ok);
  assert.equal(wb.getSheetPrintArea(0).ranges, '');

  assert.ok(wb.setSheetPrintTitles(0, '1:2', 'A:A').ok);
  const titles = wb.getSheetPrintTitles(0);
  assert.equal(titles.repeatRows, '1:2');
  assert.equal(titles.repeatCols, 'A:A');

  assert.ok(wb.addSheetRowBreak(0, 40, true).ok);
  assert.ok(wb.addSheetRowBreak(0, 10, true).ok);
  const breaks = wb.getSheetRowBreaks(0);
  assert.ok(breaks.status.ok);
  // Kept sorted by index regardless of insertion order.
  assert.deepEqual(
    breaks.breaks.map((b) => b.id),
    [10, 40],
  );
  assert.equal(breaks.breaks[0].max, 16383);
  assert.equal(breaks.breaks[0].manual, true);

  assert.ok(wb.removeSheetRowBreak(0, 10).ok);
  assert.equal(wb.getSheetRowBreaks(0).breaks.length, 1);
  assert.ok(wb.clearSheetBreaks(0).ok);
  assert.equal(wb.getSheetRowBreaks(0).breaks.length, 0);
});

test('print settings: a change is visible to paginate without a save cycle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  for (let row = 0; row < 200; row += 1) {
    for (let col = 0; col < 20; col += 1) {
      assert.ok(wb.setNumber(0, row, col, 1).ok);
    }
  }
  const portrait = wb.paginate(0).pageCount;
  assert.ok(wb.setSheetPageSetup(0, { orientation: 2 }).ok);
  assert.notEqual(wb.paginate(0).pageCount, portrait);
});

test('setRangeXfIndex applies one xf across a rectangle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const border = wb.addBorder({
    left: { style: 1, colorArgb: 0xff000000 },
    right: { style: 1, colorArgb: 0xff000000 },
    top: { style: 1, colorArgb: 0xff000000 },
    bottom: { style: 1, colorArgb: 0xff000000 },
    diagonal: { style: 0, colorArgb: 0 },
    diagonalUp: false,
    diagonalDown: false,
  });
  assert.ok(border.status.ok);
  const xf = wb.addXf({ fontIndex: 0, fillIndex: 0, borderIndex: border.index, numFmtId: 0 });
  assert.ok(xf.status.ok);

  assert.ok(wb.setRangeXfIndex(0, 0, 0, 2, 2, xf.index).ok);
  // Every cell in the rectangle is materialised, including ones that held
  // no value, so the box renders.
  for (const [row, col] of [
    [0, 0],
    [1, 1],
    [2, 2],
  ]) {
    assert.equal(wb.getCellXfIndex(0, row, col).xfIndex, xf.index);
  }
  assert.equal(wb.getCellXfIndex(0, 3, 3).xfIndex, 0);
});

test('pivotCount + pivotLayout expose PivotTable projection status', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.pivotCount(0), 0);

  const missing = wb.pivotLayout(0, 0);
  assert.equal(missing.status.ok, false);
  assert.notEqual(missing.status.status, 0);
  assert.equal(missing.top, 0);
  assert.equal(missing.left, 0);
  assert.equal(missing.rows, 0);
  assert.equal(missing.cols, 0);
  assert.deepEqual(missing.cells, []);
});

test("Workbook.createDefault + setFormula '=1+2' + recalc -> 3", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setFormula(0, 0, 0, '=1+2').ok);
  assert.ok(wb.recalc().ok);
  const r = wb.getValue(0, 0, 0);
  assert.ok(r.status.ok);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 3);
});

test('recalcParallel evaluates a wide DAG and reports bounded telemetry', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const width = 128;
  for (let row = 0; row < width; row += 1) {
    assert.ok(wb.setNumber(0, row, 0, row + 1).ok);
    assert.ok(wb.setFormula(0, row, 1, `=A${row + 1}*2`).ok);
  }

  const serial = wb.recalcParallel(1);
  assert.ok(serial.status.ok, `recalcParallel(1): ${JSON.stringify(serial.status)}`);
  assert.equal(serial.stats.workerThreadsStarted, 0);
  assert.equal(serial.stats.workerThreadsUsed, 0);
  assert.equal(serial.stats.cellsEvaluated, width);
  for (let row = 0; row < width; row += 1) {
    const value = wb.getValue(0, row, 1);
    assert.ok(value.status.ok, `serial value row ${row}: ${JSON.stringify(value.status)}`);
    assert.equal(value.value.kind, mod.ValueKind.Number);
    assert.equal(value.value.number, (row + 1) * 2);
  }

  for (let row = 0; row < width; row += 1) {
    assert.ok(wb.setNumber(0, row, 0, row + 2).ok);
  }
  const parallel = wb.recalcParallel(8);
  assert.ok(parallel.status.ok, `recalcParallel(8): ${JSON.stringify(parallel.status)}`);
  if (parallel.stats.workerThreadsStarted <= 1) {
    assert.equal(parallel.stats.workerThreadsUsed, 0);
    assert.equal(parallel.stats.parallelSteps, 0);
    assert.ok(parallel.stats.serialFallbackSteps > 0, `expected serial fallback: ${JSON.stringify(parallel)}`);
  } else {
    assert.ok(parallel.stats.parallelSteps > 0, `expected parallel steps: ${JSON.stringify(parallel)}`);
    assert.ok(parallel.stats.workerThreadsUsed > 0, `expected an effective worker: ${JSON.stringify(parallel)}`);
    assert.ok(parallel.stats.workerThreadsStarted <= 8);
    assert.ok(parallel.stats.workerThreadsUsed <= parallel.stats.workerThreadsStarted);
  }
  assert.equal(parallel.stats.cellsEvaluated, width);
  for (let row = 0; row < width; row += 1) {
    const value = wb.getValue(0, row, 1);
    assert.ok(value.status.ok, `parallel value row ${row}: ${JSON.stringify(value.status)}`);
    assert.equal(value.value.kind, mod.ValueKind.Number);
    assert.equal(value.value.number, (row + 2) * 2);
  }

  const assertInvalid = (input, label) => {
    const invalid = input === undefined ? wb.recalcParallel() : wb.recalcParallel(input);
    assert.equal(invalid.status.ok, false, label);
    assert.notEqual(invalid.status.status, 0, label);
    assert.equal(invalid.stats.cellsEvaluated, 0, label);
    assert.equal(invalid.stats.sccsProcessed, 0, label);
    assert.equal(invalid.stats.parallelSteps, 0, label);
    assert.equal(invalid.stats.serialFallbackSteps, 0, label);
    assert.equal(invalid.stats.cycleRecoveries, 0, label);
    assert.equal(invalid.stats.workerThreadsStarted, 0, label);
    assert.equal(invalid.stats.workerThreadsUsed, 0, label);
    return invalid;
  };
  assertInvalid(undefined, 'missing threadCount');
  assertInvalid(null, 'null threadCount');
  assertInvalid(NaN, 'NaN threadCount');
  assertInvalid(Infinity, 'Infinity threadCount');
  assertInvalid(-Infinity, '-Infinity threadCount');
  assertInvalid(1.5, 'fractional threadCount');
  assertInvalid(-1, 'negative threadCount');
  const invalid = assertInvalid(9, 'threadCount above cap');
  assert.equal(invalid.status.message, 'fm_workbook_recalc_parallel: thread_count must be 0..8');
  assert.equal(invalid.status.context, 'thread_count=9 max=8');
  wb.dispose();
});

test("evaluateFormulaArray('=SEQUENCE(2,3)') returns a 2x3 grid", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const r = wb.evaluateFormulaArray(0, 0, 0, '=SEQUENCE(2,3)');
  assert.ok(r.status.ok);
  assert.equal(r.rows, 2);
  assert.equal(r.cols, 3);
  assert.equal(r.cells.length, 2);
  assert.equal(r.cells[0].length, 3);
  // Row-major 1..6.
  assert.equal(r.cells[0][0].kind, mod.ValueKind.Number);
  assert.equal(r.cells[0][0].number, 1);
  assert.equal(r.cells[1][2].number, 6);
});

test("evaluateFormulaArray('=1+2') reports a 1x1 scalar array", async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const r = wb.evaluateFormulaArray(0, 0, 0, '=1+2');
  assert.ok(r.status.ok);
  assert.equal(r.rows, 1);
  assert.equal(r.cols, 1);
  assert.equal(r.cells[0][0].kind, mod.ValueKind.Number);
  assert.equal(r.cells[0][0].number, 3);
});

test('Workbook.createEmpty + addSheet -> sheetCount/sheetName work', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createEmpty();
  assert.ok(wb.addSheet('Sheet1').ok);
  assert.equal(wb.sheetCount(), 1);
  const sn = wb.sheetName(0);
  assert.ok(sn.status.ok);
  assert.equal(sn.value, 'Sheet1');
});

test('Unicode sheet identity rejects folded duplicate and permits casing rename', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createEmpty();
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
  wb.dispose();
});

test('save() returns Uint8Array; loadBytes round-trips the value', async () => {
  const mod = await getModule();
  const wb1 = mod.Workbook.createDefault();
  assert.ok(wb1.setNumber(0, 0, 0, 42).ok);
  const sr = wb1.save();
  assert.ok(sr.status.ok, `save: ${JSON.stringify(sr.status)}`);
  assert.ok(sr.bytes instanceof Uint8Array, 'expected Uint8Array');
  assert.ok(sr.bytes.length > 0, 'expected non-empty save buffer');

  const wb2 = mod.Workbook.loadBytes(sr.bytes);
  const r = wb2.getValue(0, 0, 0);
  assert.ok(r.status.ok, `loaded getValue: ${JSON.stringify(r.status)}`);
  assert.equal(r.value.kind, mod.ValueKind.Number);
  assert.equal(r.value.number, 42);
});

test('loadBytes rejects invalid input with a fresh diagnostic', async () => {
  const mod = await getModule();
  const loaded = mod.Workbook.loadBytes(new Uint16Array([1]));
  assert.equal(loaded.isValid(), false);
  assert.match(mod.lastErrorMessage(), /NULL or empty input/);
});

// -- Expanded surface coverage -----------------------------------------
// These tests exercise the methods the addon mirrors from the embind
// binding. They are deliberately shallow: each call should round-trip
// some observable state without crashing the Node process.

test('isValid returns true for a live workbook', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.isValid(), true);
});

test('dispose deterministically releases a workbook and is idempotent', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.isValid(), true);
  assert.equal(wb.dispose(), undefined);
  assert.equal(wb.isValid(), false);
  assert.equal(wb.getValue(0, 0, 0).status.ok, false);
  assert.equal(wb.dispose(), undefined);
});

test('memoryUsage grows with the workbook and reads zero after dispose', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();

  const empty = wb.memoryUsage();
  assert.ok(empty > 0, `an empty workbook still owns storage, got ${empty}`);

  // Distinct strings so the shared-string storage cannot fold them into
  // one entry and hide the growth.
  for (let row = 0; row < 2000; row += 1) {
    wb.setText(0, row, 0, `payload-${row}-${'x'.repeat(64)}`);
  }
  const filled = wb.memoryUsage();
  assert.ok(filled > empty, `expected growth, got ${empty} -> ${filled}`);

  wb.dispose();
  assert.equal(wb.memoryUsage(), 0);
});

test('memoryUsage is stable across repeated calls and survives a recalc', async () => {
  // The external-memory report is computed as a delta against the last
  // reported figure, so a repeated call on an unchanged workbook must be
  // a no-op rather than double-counting, and the operations that trigger
  // their own report must not disturb the estimate either.
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  for (let row = 0; row < 500; row += 1) {
    wb.setFormula(0, row, 0, `=${row}+1`);
  }

  const first = wb.memoryUsage();
  assert.equal(wb.memoryUsage(), first);

  assert.equal(wb.recalc().ok, true);
  const afterRecalc = wb.memoryUsage();
  assert.ok(afterRecalc >= first, `recalc must not shrink the estimate: ${first} -> ${afterRecalc}`);
  assert.equal(wb.memoryUsage(), afterRecalc);

  wb.dispose();
});

test('Workbook factories return instances of the exported Workbook class', async () => {
  const mod = await getModule();
  const defaultBook = mod.Workbook.createDefault();
  const emptyBook = mod.Workbook.createEmpty();
  assert.ok(defaultBook instanceof mod.Workbook);
  assert.ok(emptyBook instanceof mod.Workbook);
});

test('addMerge + getMerges round-trip a single range', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const range = { firstRow: 1, firstCol: 1, lastRow: 2, lastCol: 3 };
  const ar = wb.addMerge(0, range);
  assert.ok(ar.ok, `addMerge: ${JSON.stringify(ar)}`);
  const list = wb.getMerges(0);
  assert.ok(Array.isArray(list), `expected Array, got ${typeof list}`);
  assert.equal(list.length, 1);
  assert.deepEqual(
    {
      firstRow: list[0].firstRow,
      firstCol: list[0].firstCol,
      lastRow: list[0].lastRow,
      lastCol: list[0].lastCol,
    },
    range,
  );
});

test('addHyperlink + getHyperlinks round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getHyperlinks(0).length, 0);
  // Three entries with progressively more optional fields populated.
  assert.ok(wb.addHyperlink(0, 1, 2, 'https://example.com/', '', '', '').ok);
  assert.ok(wb.addHyperlink(0, 3, 4, 'mailto:hello@example.com', 'Hello', '', '').ok);
  assert.ok(wb.addHyperlink(0, 5, 6, '', 'See X', 'Internal link', 'Sheet1!A1').ok);
  assert.ok(wb.addHyperlinkRange(0, 7, 8, 9, 10, 'https://example.com/range', 'Range', 'Range tip', '').ok);
  const list = wb.getHyperlinks(0);
  assert.equal(list.length, 4);
  assert.equal(list[0].row, 1);
  assert.equal(list[0].col, 2);
  assert.equal(list[0].lastRow, 1);
  assert.equal(list[0].lastCol, 2);
  assert.equal(list[0].target, 'https://example.com/');
  assert.equal(list[0].display, '');
  assert.equal(list[0].tooltip, '');
  assert.equal(list[1].target, 'mailto:hello@example.com');
  assert.equal(list[1].display, 'Hello');
  assert.equal(list[2].display, 'See X');
  assert.equal(list[2].tooltip, 'Internal link');
  assert.equal(list[2].location, 'Sheet1!A1');
  assert.deepEqual(
    {
      row: list[3].row,
      col: list[3].col,
      lastRow: list[3].lastRow,
      lastCol: list[3].lastCol,
      target: list[3].target,
      display: list[3].display,
      tooltip: list[3].tooltip,
    },
    {
      row: 7,
      col: 8,
      lastRow: 9,
      lastCol: 10,
      target: 'https://example.com/range',
      display: 'Range',
      tooltip: 'Range tip',
    },
  );
  // Sheet-out-of-range is rejected.
  assert.ok(!wb.addHyperlink(999, 0, 0, 'https://x/', '', '', '').ok);
  // clearHyperlinks drops everything.
  assert.ok(wb.clearHyperlinks(0).ok);
  assert.equal(wb.getHyperlinks(0).length, 0);
});

test('removeHyperlink / removeHyperlinkAt / clearHyperlinks surface on an empty sheet', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // No-op variants on an empty list still return kOk.
  assert.ok(wb.removeHyperlink(0, 0, 0).ok);
  assert.ok(wb.clearHyperlinks(0).ok);
  // Out-of-range index is rejected.
  assert.ok(!wb.removeHyperlinkAt(0, 0).ok);
  // Sheet-index-out-of-range is rejected on every variant.
  assert.ok(!wb.removeHyperlink(999, 0, 0).ok);
  assert.ok(!wb.removeHyperlinkAt(999, 0).ok);
  assert.ok(!wb.clearHyperlinks(999).ok);
  // The hyperlink list is still empty afterwards.
  assert.equal(wb.getHyperlinks(0).length, 0);
});

test('removeMerge + removeMergeAt + clearMerges step-wise prune the merge list', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const a = { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 };
  const b = { firstRow: 4, firstCol: 4, lastRow: 5, lastCol: 5 };
  assert.ok(wb.addMerge(0, a).ok);
  assert.ok(wb.addMerge(0, b).ok);
  // removeMerge with an overlap that hits `a` only.
  const overlap = { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 };
  assert.ok(wb.removeMerge(0, overlap).ok);
  let list = wb.getMerges(0);
  assert.equal(list.length, 1);
  assert.equal(list[0].firstRow, 4);
  // removeMergeAt drops the survivor by index.
  assert.ok(wb.removeMergeAt(0, 0).ok);
  list = wb.getMerges(0);
  assert.equal(list.length, 0);
  // clearMerges is a no-op on an empty list and stays kOk.
  assert.ok(wb.addMerge(0, a).ok);
  assert.ok(wb.addMerge(0, b).ok);
  assert.ok(wb.clearMerges(0).ok);
  list = wb.getMerges(0);
  assert.equal(list.length, 0);
});

test('setSheetZoom + getSheetView surface the new zoom', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setSheetZoom(0, 175).ok);
  const v = wb.getSheetView(0);
  assert.ok(v.status.ok, `getSheetView: ${JSON.stringify(v.status)}`);
  assert.equal(v.view.zoomScale, 175);
  // Default freeze / tab-hidden state survives a zoom change.
  assert.equal(v.view.freezeRows, 0);
  assert.equal(v.view.freezeCols, 0);
  assert.equal(v.view.tabHidden, 0);
});

test('sheet layout getters preserve explicit width presence flags', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setColumnWidth(0, 0, 0, 0).ok);
  assert.ok(wb.setColumnHidden(0, 1, 1, true).ok);
  assert.ok(wb.setRowHeight(0, 3, 30).ok);

  const columns = wb.getSheetColumns(0);
  assert.ok(columns.status.ok, `getSheetColumns: ${JSON.stringify(columns.status)}`);
  const widthZero = columns.columns.find((column) => column.first === 0 && column.last === 0);
  assert.ok(widthZero);
  assert.equal(widthZero.width, 0);
  assert.equal(widthZero.hasWidth, 1);
  assert.equal(widthZero.hasStyle, 0);
  const hidden = columns.columns.find((column) => column.first === 1 && column.last === 1);
  assert.ok(hidden);
  assert.equal(hidden.hidden, 1);
  assert.equal(hidden.hasWidth, 0);

  const rows = wb.getSheetRowOverrides(0);
  assert.ok(rows.status.ok, `getSheetRowOverrides: ${JSON.stringify(rows.status)}`);
  const row = rows.rows.find((entry) => entry.row === 3);
  assert.ok(row);
  assert.equal(row.height, 30);
  assert.equal(row.hasStyle, 0);
  assert.equal(row.styleXf, 0);
});

test('getSheetView surfaces display / orientation defaults', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const v = wb.getSheetView(0);
  assert.ok(v.status.ok, `getSheetView: ${JSON.stringify(v.status)}`);
  assert.equal(v.view.showGridLines, 1);
  assert.equal(v.view.showRowColHeaders, 1);
  assert.equal(v.view.showZeros, 1);
  assert.equal(v.view.rightToLeft, 0);
  assert.equal(v.view.tabSelected, 0);
  assert.equal(v.view.viewMode, '');
});

test('setSheetShowGridLines / setSheetRightToLeft / setSheetTabSelected / setSheetViewMode round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setSheetShowGridLines(0, false).ok);
  assert.ok(wb.setSheetShowRowColHeaders(0, false).ok);
  assert.ok(wb.setSheetShowZeros(0, false).ok);
  assert.ok(wb.setSheetRightToLeft(0, true).ok);
  assert.ok(wb.setSheetTabSelected(0, true).ok);
  assert.ok(wb.setSheetViewMode(0, 'pageBreakPreview').ok);
  const v = wb.getSheetView(0);
  assert.ok(v.status.ok, `getSheetView: ${JSON.stringify(v.status)}`);
  assert.equal(v.view.showGridLines, 0);
  assert.equal(v.view.showRowColHeaders, 0);
  assert.equal(v.view.showZeros, 0);
  assert.equal(v.view.rightToLeft, 1);
  assert.equal(v.view.tabSelected, 1);
  assert.equal(v.view.viewMode, 'pageBreakPreview');
});

test('insertRows shifts an existing literal forward', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Populate row 0; insert one row at row 0; the literal is now at row 1.
  assert.ok(wb.setNumber(0, 0, 0, 99).ok);
  assert.ok(wb.insertRows(0, 0, 1).ok);
  const r0 = wb.getValue(0, 0, 0);
  assert.ok(r0.status.ok);
  // The freshly-inserted row 0 is blank.
  assert.equal(r0.value.kind, mod.ValueKind.Blank);
  const r1 = wb.getValue(0, 1, 0);
  assert.ok(r1.status.ok);
  assert.equal(r1.value.kind, mod.ValueKind.Number);
  assert.equal(r1.value.number, 99);
});

test('setCellXfIndex + getCellXfIndex round-trip the xf id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Seed the cell so the xf attaches to a stored slot.
  assert.ok(wb.setNumber(0, 0, 0, 1).ok);
  // xfIndex 0 is always the default xf and is guaranteed to be valid.
  assert.ok(wb.setCellXfIndex(0, 0, 0, 0).ok);
  const r = wb.getCellXfIndex(0, 0, 0);
  assert.ok(r.status.ok, `getCellXfIndex: ${JSON.stringify(r.status)}`);
  assert.equal(typeof r.xfIndex, 'number');
  assert.equal(r.xfIndex, 0);
});

test('style building blocks: addFont -> addXf -> setCellXfIndex -> getXf', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A fresh workbook seeds Excel's minimum style table: one default font,
  // one empty border, one default xf, and the two fills Excel reserves at
  // the front of every file (`none`, `gray125`). A caller's first record
  // therefore lands after them instead of displacing `none`.
  assert.equal(wb.fontCount(), 1);
  assert.equal(wb.fillCount(), 2);
  assert.equal(wb.borderCount(), 1);
  assert.equal(wb.xfCount(), 1);

  const fontResult = wb.addFont({
    name: 'Arial',
    size: 12,
    bold: true,
    italic: false,
    strike: false,
    underline: 0,
    colorArgb: 0xff112233,
  });
  assert.ok(fontResult.status.ok, `addFont: ${JSON.stringify(fontResult.status)}`);
  // Linear-search dedup: the second call with the same payload returns the same index.
  const fontDup = wb.addFont({
    name: 'Arial',
    size: 12,
    bold: true,
    italic: false,
    strike: false,
    underline: 0,
    colorArgb: 0xff112233,
  });
  assert.equal(fontDup.index, fontResult.index);
  assert.equal(wb.fontCount(), 2);

  const fill = wb.addFill({ pattern: 1, fgArgb: 0xffff0000, bgArgb: 0xff000000 });
  assert.ok(fill.status.ok);
  const border = wb.addBorder({
    left: { style: 1, colorArgb: 0xff000000 },
    right: { style: 1, colorArgb: 0xff000000 },
    top: { style: 1, colorArgb: 0xff000000 },
    bottom: { style: 1, colorArgb: 0xff000000 },
    diagonal: { style: 0, colorArgb: 0 },
    diagonalUp: false,
    diagonalDown: false,
  });
  assert.ok(border.status.ok);

  const builtin = wb.addNumFmt('General');
  assert.ok(builtin.status.ok);
  assert.equal(builtin.numFmtId, 0);

  const custom = wb.addNumFmt('"USD" #,##0');
  assert.ok(custom.status.ok);
  assert.ok(custom.numFmtId >= 164, `expected custom id >= 164, got ${custom.numFmtId}`);

  const xf = wb.addXf({
    fontIndex: fontResult.index,
    fillIndex: fill.index,
    borderIndex: border.index,
    numFmtId: custom.numFmtId,
    horizontalAlign: 1,
    verticalAlign: 2,
    wrapText: true,
  });
  assert.ok(xf.status.ok, `addXf: ${JSON.stringify(xf.status)}`);

  assert.ok(wb.setNumber(0, 0, 0, 1).ok);
  assert.ok(wb.setCellXfIndex(0, 0, 0, xf.index).ok);

  const reread = wb.getCellXf(xf.index);
  assert.ok(reread.status.ok);
  assert.equal(reread.fontIndex, fontResult.index);
  assert.equal(reread.numFmtId, custom.numFmtId);
  assert.equal(reread.wrapText, true);
  assert.equal(reread.hasAlignment, true);
  assert.equal(reread.hasHorizontalAlign, true);
  assert.equal(reread.hasVerticalAlign, true);
  assert.equal(reread.hasWrapText, true);
  assert.equal(reread.hasJustifyLastLine, false);

  const optional = wb.addXf({
    fontIndex: fontResult.index,
    fillIndex: fill.index,
    borderIndex: border.index,
    numFmtId: custom.numFmtId,
    textRotation: 255,
    indent: 0,
    relativeIndent: -3,
    shrinkToFit: false,
    readingOrder: 0,
    justifyLastLine: true,
  });
  assert.ok(optional.status.ok, `optional alignment: ${JSON.stringify(optional.status)}`);
  const optionalRead = wb.getCellXf(optional.index);
  assert.ok(optionalRead.status.ok);
  assert.equal(optionalRead.textRotation, 255);
  assert.equal(optionalRead.indent, 0);
  assert.equal(optionalRead.relativeIndent, -3);
  assert.equal(optionalRead.shrinkToFit, false);
  assert.equal(optionalRead.readingOrder, 0);
  assert.equal(optionalRead.justifyLastLine, true);
  assert.equal(optionalRead.hasAlignment, true);
  assert.equal(optionalRead.hasJustifyLastLine, true);
  assert.equal(wb.addXf(optionalRead).index, optional.index);

  const omitted = wb.addXf({
    fontIndex: fontResult.index,
    fillIndex: fill.index,
    borderIndex: border.index,
    numFmtId: custom.numFmtId,
  });
  assert.ok(omitted.status.ok);
  const omittedRead = wb.getCellXf(omitted.index);
  assert.ok(omittedRead.status.ok);
  assert.equal(omittedRead.hasAlignment, false);

  const explicitEmpty = wb.addXf({
    fontIndex: fontResult.index,
    fillIndex: fill.index,
    borderIndex: border.index,
    numFmtId: custom.numFmtId,
    hasAlignment: true,
  });
  assert.ok(explicitEmpty.status.ok);
  assert.notEqual(explicitEmpty.index, omitted.index);
  const explicitEmptyRead = wb.getCellXf(explicitEmpty.index);
  assert.ok(explicitEmptyRead.status.ok);
  assert.equal(explicitEmptyRead.hasAlignment, true);

  const rfont = wb.getFont(fontResult.index);
  assert.ok(rfont.status.ok);
  assert.equal(rfont.name, 'Arial');
  assert.equal(rfont.bold, true);

  const rfill = wb.getFill(fill.index);
  assert.ok(rfill.status.ok);
  assert.equal(rfill.pattern, 1);

  const rborder = wb.getBorder(border.index);
  assert.ok(rborder.status.ok);
  assert.equal(rborder.left.style, 1);
  assert.equal(rborder.diagonalUp, false);

  const rfmt = wb.getNumFmt(custom.numFmtId);
  assert.ok(rfmt.status.ok);
  assert.equal(rfmt.formatCode, '"USD" #,##0');
});

test('selector colours survive get/add identity and save/load', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const theme = { kind: 2, rgb: 0, theme: 3, tint: 0.5, indexed: 0 };
  const indexed = { kind: 3, rgb: 0, theme: 0, tint: 0, indexed: 9 };
  const automatic = { kind: 4, rgb: 0, theme: 0, tint: 0, indexed: 0 };

  const font = wb.addFont({ name: 'SelectorFont', size: 11, colorArgb: 0x01020304, color: theme });
  assert.ok(font.status.ok);
  const gotFont = wb.getFont(font.index);
  assert.ok(gotFont.status.ok);
  assert.equal(gotFont.color.kind, 2);
  assert.equal(gotFont.color.theme, 3);
  assert.equal(gotFont.color.tint, 0.5);
  assert.equal(gotFont.colorArgb, 0x01020304);
  const fontAgain = wb.addFont(gotFont);
  assert.ok(fontAgain.status.ok);
  assert.equal(fontAgain.index, font.index);

  const fill = wb.addFill({
    pattern: 1,
    fgArgb: 0x05060708,
    bgArgb: 0x090a0b0c,
    fg: indexed,
    bg: automatic,
  });
  assert.ok(fill.status.ok);
  const gotFill = wb.getFill(fill.index);
  assert.ok(gotFill.status.ok);
  assert.equal(gotFill.fg.kind, 3);
  assert.equal(gotFill.fg.indexed, 9);
  assert.equal(gotFill.bg.kind, 4);
  const fillAgain = wb.addFill(gotFill);
  assert.ok(fillAgain.status.ok);
  assert.equal(fillAgain.index, fill.index);

  const border = wb.addBorder({
    left: { style: 1, colorArgb: 0x01020304, color: theme },
    right: { style: 1, colorArgb: 0x05060708, color: indexed },
    top: { style: 1, colorArgb: 0x090a0b0c, color: automatic },
    bottom: { style: 0, colorArgb: 0 },
    diagonal: { style: 0, colorArgb: 0 },
  });
  assert.ok(border.status.ok);
  const gotBorder = wb.getBorder(border.index);
  assert.ok(gotBorder.status.ok);
  assert.equal(gotBorder.left.color.kind, 2);
  assert.equal(gotBorder.left.color.theme, 3);
  assert.equal(gotBorder.right.color.kind, 3);
  assert.equal(gotBorder.right.color.indexed, 9);
  assert.equal(gotBorder.top.color.kind, 4);
  const borderAgain = wb.addBorder(gotBorder);
  assert.ok(borderAgain.status.ok);
  assert.equal(borderAgain.index, border.index);

  const dxf = wb.addDxf({
    font: { name: 'DxfSelector', size: 9, colorArgb: 0x11121314, color: automatic },
    fill: { pattern: 1, fgArgb: 0x15161718, fg: indexed },
    border: { left: { style: 1, colorArgb: 0x191a1b1c, color: theme } },
  });
  assert.ok(dxf.status.ok);
  const gotDxf = wb.getDxf(dxf.index);
  assert.ok(gotDxf.status.ok);
  assert.equal(gotDxf.font.color.kind, 4);
  assert.equal(gotDxf.fill.fg.kind, 3);
  assert.equal(gotDxf.border.left.color.kind, 2);
  const dxfAgain = wb.addDxf(gotDxf);
  assert.ok(dxfAgain.status.ok);
  assert.equal(dxfAgain.index, dxf.index);

  const saved = wb.save();
  assert.ok(saved.status.ok);
  const loaded = mod.Workbook.loadBytes(saved.bytes);
  const loadedFont = loaded.getFont(font.index);
  const loadedFill = loaded.getFill(fill.index);
  const loadedBorder = loaded.getBorder(border.index);
  const loadedDxf = loaded.getDxf(dxf.index);
  assert.ok(loadedFont.status.ok);
  assert.ok(loadedFill.status.ok);
  assert.ok(loadedBorder.status.ok);
  assert.ok(loadedDxf.status.ok);
  assert.equal(loadedFont.color.kind, 2);
  assert.equal(loadedFill.fg.kind, 3);
  assert.equal(loadedBorder.left.color.kind, 2);
  assert.equal(loadedDxf.font.color.kind, 4);
  loaded.dispose();
  wb.dispose();
});

test('dxf alignment and protection XML survive get/add identity and save/load', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const alignmentXml = '<alignment horizontal="center" wrapText="1"/>';
  const protectionXml = '<protection locked="0" hidden="1"/>';
  const alignment = wb.addDxf({ alignmentXml });
  assert.ok(alignment.status.ok);
  const protection = wb.addDxf({ protectionXml });
  assert.ok(protection.status.ok);
  assert.notEqual(alignment.index, protection.index);

  const gotAlignment = wb.getDxf(alignment.index);
  assert.ok(gotAlignment.status.ok);
  assert.equal(gotAlignment.alignmentXml, alignmentXml);
  assert.equal(gotAlignment.protectionXml, undefined);
  const gotProtection = wb.getDxf(protection.index);
  assert.ok(gotProtection.status.ok);
  assert.equal(gotProtection.alignmentXml, undefined);
  assert.equal(gotProtection.protectionXml, protectionXml);

  const alignmentAgain = wb.addDxf(gotAlignment);
  assert.ok(alignmentAgain.status.ok);
  assert.equal(alignmentAgain.index, alignment.index);
  const protectionAgain = wb.addDxf(gotProtection);
  assert.ok(protectionAgain.status.ok);
  assert.equal(protectionAgain.index, protection.index);

  const saved = wb.save();
  assert.ok(saved.status.ok);
  const loaded = mod.Workbook.loadBytes(saved.bytes);
  assert.ok(loaded.isValid());
  const loadedAlignment = loaded.getDxf(alignment.index);
  assert.ok(loadedAlignment.status.ok);
  assert.equal(loadedAlignment.alignmentXml, alignmentXml);
  assert.equal(loadedAlignment.protectionXml, undefined);
  const loadedProtection = loaded.getDxf(protection.index);
  assert.ok(loadedProtection.status.ok);
  assert.equal(loadedProtection.alignmentXml, undefined);
  assert.equal(loadedProtection.protectionXml, protectionXml);
  const loadedAlignmentAgain = loaded.addDxf(loadedAlignment);
  assert.ok(loadedAlignmentAgain.status.ok);
  assert.equal(loadedAlignmentAgain.index, alignment.index);
  const loadedProtectionAgain = loaded.addDxf(loadedProtection);
  assert.ok(loadedProtectionAgain.status.ok);
  assert.equal(loadedProtectionAgain.index, protection.index);
  loaded.dispose();
  wb.dispose();
});

test('addFont / getFont preserve superscript and round-trip to the same index', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const superscript = wb.addFont({ name: 'Arial', size: 12, vertAlign: 1, colorArgb: 0xff112233 });
  assert.ok(superscript.status.ok, `addFont: ${JSON.stringify(superscript.status)}`);
  assert.equal(wb.getFont(superscript.index).vertAlign, 1);

  const before = wb.fontCount();
  const again = wb.addFont(wb.getFont(superscript.index));
  assert.ok(again.status.ok);
  assert.equal(again.index, superscript.index);
  assert.equal(wb.fontCount(), before);

  const edited = wb.getFont(superscript.index);
  edited.colorArgb = 0xff00ff00;
  const recolored = wb.addFont(edited);
  assert.ok(recolored.status.ok);
  assert.notEqual(recolored.index, superscript.index);
  const reread = wb.getFont(recolored.index);
  assert.equal(reread.vertAlign, 1);
  assert.equal(reread.colorArgb, 0xff00ff00);
});

test('phonetic runs keep their spans through a save/load round trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.setText(0, 0, 0, '東京都');
  assert.deepEqual(wb.getCellPhoneticRuns(0, 0, 0).runs, []);

  const runs = [
    { sb: 0, eb: 2, text: 'トウキョウ' },
    { sb: 2, eb: 3, text: 'ト' },
  ];
  const stored = wb.setCellPhoneticRuns(0, 0, 0, runs);
  assert.ok(stored.ok, `setCellPhoneticRuns: ${JSON.stringify(stored)}`);
  assert.deepEqual(wb.getCellPhoneticRuns(0, 0, 0).runs, runs);
  // The flattening getter still reports the concatenation.
  assert.equal(wb.getCellPhonetic(0, 0, 0).value, 'トウキョウト');

  const saved = wb.save();
  assert.ok(saved.status.ok, `save: ${JSON.stringify(saved.status)}`);
  const loaded = mod.Workbook.loadBytes(saved.bytes);
  assert.deepEqual(loaded.getCellPhoneticRuns(0, 0, 0).runs, runs);

  // Writing the flattened reading back is the collapse the run API avoids.
  wb.setCellPhonetic(0, 0, 0, 'トウキョウト');
  assert.deepEqual(wb.getCellPhoneticRuns(0, 0, 0).runs, [{ sb: 0, eb: 3, text: 'トウキョウト' }]);

  const rejected = wb.setCellPhoneticRuns(0, 0, 0, [
    { sb: 2, eb: 3, text: 'ト' },
    { sb: 0, eb: 2, text: 'トウ' },
  ]);
  assert.equal(rejected.ok, false);

  loaded.dispose();
  wb.dispose();
});

test('setDefaultFont declares what an unstyled cell is saved as', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getFont(0).name, 'Calibri');

  // addFont can only ever append beside the seeded default.
  const appended = wb.addFont({ name: '游ゴシック', size: 11 });
  assert.ok(appended.status.ok);
  assert.ok(appended.index > 0);
  assert.equal(wb.getFont(0).name, 'Calibri');

  const declared = wb.setDefaultFont({ name: '游ゴシック', size: 11, hasCharset: true, charset: 128 });
  assert.ok(declared.ok, `setDefaultFont: ${JSON.stringify(declared)}`);
  assert.equal(wb.getFont(0).name, '游ゴシック');
  assert.equal(wb.getFont(0).charset, 128);

  wb.dispose();
});

test('setFont overwrites an existing slot and refuses an absent index', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const added = wb.addFont({ name: 'Meiryo', size: 12 });
  assert.ok(added.status.ok);

  const replaced = wb.setFont(added.index, { name: 'MS Gothic', size: 9 });
  assert.ok(replaced.ok, `setFont: ${JSON.stringify(replaced)}`);
  assert.equal(wb.getFont(added.index).name, 'MS Gothic');

  const before = wb.fontCount();
  assert.equal(wb.setFont(before, { name: 'MS Gothic', size: 9 }).ok, false);
  assert.equal(wb.fontCount(), before);

  wb.dispose();
});

test('addDxf / getDxf round-trip a superscript differential font', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const added = wb.addDxf({ font: { name: 'Calibri', size: 9, vertAlign: 1 } });
  assert.ok(added.status.ok, `addDxf: ${JSON.stringify(added.status)}`);
  const readBack = wb.getDxf(added.index);
  assert.ok(readBack.status.ok);
  assert.equal(readBack.font.vertAlign, 1);

  const before = wb.dxfCount();
  const again = wb.addDxf(readBack);
  assert.ok(again.status.ok);
  assert.equal(again.index, added.index);
  assert.equal(wb.dxfCount(), before);
});

test('addXf rejects out-of-range font_index', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const r = wb.addXf({
    fontIndex: 99,
    fillIndex: 0,
    borderIndex: 0,
    numFmtId: 0,
    horizontalAlign: 0,
    verticalAlign: 0,
    wrapText: false,
  });
  assert.ok(!r.status.ok, 'expected addXf to reject out-of-range font_index');
});

test('setIterativeProgress accepts a function and accepts null to clear', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Registration roundtrip; we don't assert the callback fires because
  // a non-iterative recalc never invokes the trampoline. The smoke
  // signal we want is "the addon does not crash on register / clear".
  const cb = () => true;
  const reg = wb.setIterativeProgress(cb);
  assert.ok(reg.ok, `setIterativeProgress(fn): ${JSON.stringify(reg)}`);
  const clr = wb.setIterativeProgress(null);
  assert.ok(clr.ok, `setIterativeProgress(null): ${JSON.stringify(clr)}`);
});

test('setIterativeProgress keeps callbacks isolated per Workbook', async () => {
  const mod = await getModule();
  const first = mod.Workbook.createDefault();
  const second = mod.Workbook.createDefault();
  let firstCalls = 0;
  let secondCalls = 0;

  for (const wb of [first, second]) {
    assert.ok(wb.setIterative(true, 10, 0.001).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
  }
  assert.ok(
    first.setIterativeProgress(() => {
      firstCalls += 1;
      return true;
    }).ok,
  );
  assert.ok(
    second.setIterativeProgress(() => {
      secondCalls += 1;
      return true;
    }).ok,
  );

  assert.ok(first.recalc().ok);
  assert.ok(firstCalls > 0);
  assert.equal(secondCalls, 0);
});

test('dispose is rejected while this workbook is executing its progress callback', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setIterative(true, 10, 0.001).ok);
  assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
  let calls = 0;
  assert.ok(
    wb.setIterativeProgress(() => {
      calls += 1;
      assert.throws(() => wb.dispose(), /cannot dispose a Workbook/);
      return true;
    }).ok,
  );
  assert.ok(wb.recalc().ok);
  assert.ok(calls > 0);
  assert.equal(wb.isValid(), true);
  wb.dispose();
});

test('evaluateCfRange returns ok envelope with empty cells for a CF-less workbook', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // The default workbook has no CF rules, so the call should succeed
  // and return an empty cell list. NaN disables `TimePeriod` rules.
  const r = wb.evaluateCfRange(0, 0, 0, 4, 4, Number.NaN);
  assert.ok(r.status.ok, `evaluateCfRange: ${JSON.stringify(r.status)}`);
  assert.ok(Array.isArray(r.cells));
  assert.equal(r.cells.length, 0);
});

test('evaluateCfRange with todaySerial omitted disables timePeriod rules', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // timePeriod rule (type 15) for "today" (period 0) over A1. The cell
  // holds the serial for 1899-12-30, which a 0.0 basis would match.
  assert.ok(wb.setNumber(0, 0, 0, 0).ok);
  const dxf = wb.addDxf({ font: { bold: true } });
  assert.ok(dxf.status.ok, JSON.stringify(dxf.status));
  const add = wb.addConditionalFormat(0, {
    sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
    type: 15,
    timePeriod: 0,
    dxfId: dxf.index,
  });
  assert.ok(add.status.ok, `addConditionalFormat: ${JSON.stringify(add)}`);
  assert.ok(wb.recalc().ok);

  // Omitting todaySerial must behave like passing NaN (the documented
  // disabling value), not like passing 0 -- which is the valid serial for
  // 1899-12-30 and would make the rule match.
  const omitted = wb.evaluateCfRange(0, 0, 0, 0, 0);
  assert.ok(omitted.status.ok, `evaluateCfRange: ${JSON.stringify(omitted.status)}`);
  const explicitNaN = wb.evaluateCfRange(0, 0, 0, 0, 0, Number.NaN);
  assert.ok(explicitNaN.status.ok, `evaluateCfRange: ${JSON.stringify(explicitNaN.status)}`);
  assert.equal(omitted.cells.length, explicitNaN.cells.length);
  assert.equal(omitted.cells.length, 0);

  // A concrete basis of 0 does engage the rule, proving the default is
  // not merely a no-op for this workbook.
  const zeroBasis = wb.evaluateCfRange(0, 0, 0, 0, 0, 0);
  assert.ok(zeroBasis.status.ok, `evaluateCfRange: ${JSON.stringify(zeroBasis.status)}`);
  assert.equal(zeroBasis.cells.length, 1);
});

// Builds a workbook with one pivot, ready for the filter-contract cases.
// `dispose` is the caller's job; the WASM package uses `delete` instead.
function makePivotWorkbook(Workbook) {
  const wb = Workbook.createDefault();
  const cacheId = wb.pivotCacheCreate(0).index;
  // A cache with no declared source cannot be saved: Excel offers to repair
  // any package containing one, so the writer refuses rather than emit it.
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
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
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
    wb.dispose();
  }
});

test('pivotFilterCount reports only what this session added', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    assert.equal(wb.pivotFilterCount(0, pivot), 0);
    assert.ok(wb.pivotFilterAdd(0, pivot, PIVOT_FILTER).ok);
    assert.equal(wb.pivotFilterCount(0, pivot), 1);
  } finally {
    wb.dispose();
  }
});

test('active filters are session state and do not survive save/load', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  let bytes;
  try {
    assert.ok(wb.pivotFilterAdd(0, pivot, PIVOT_FILTER).ok);
    assert.equal(wb.pivotFilterCount(0, pivot), 1);
    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    bytes = saved.bytes;
  } finally {
    wb.dispose();
  }
  const reloaded = mod.Workbook.loadBytes(bytes);
  try {
    assert.equal(reloaded.pivotCount(0), 1);
    // The pivot round-trips; its active-filter list deliberately does not.
    assert.equal(reloaded.pivotFilterCount(0, 0), 0);
  } finally {
    reloaded.dispose();
  }
});

test('pivotFilterAt rejects an out-of-range index', async () => {
  const mod = await getModule();
  const { wb, pivot } = makePivotWorkbook(mod.Workbook);
  try {
    const got = wb.pivotFilterAt(0, pivot, 99);
    assert.equal(got.status.ok, false);
  } finally {
    wb.dispose();
  }
});

test('comments round-trip: setComment + getComment', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Pre-existing cell so the comment attaches to a stored slot.
  assert.ok(wb.setNumber(0, 1, 1, 7).ok);
  const sc = wb.setComment(0, 1, 1, 'libraz', 'hello world');
  assert.ok(sc.ok, `setComment: ${JSON.stringify(sc)}`);
  const c = wb.getComment(0, 1, 1);
  assert.ok(c !== null, 'expected non-null CommentEntry');
  assert.equal(c.author, 'libraz');
  assert.equal(c.text, 'hello world');
  // Removing surfaces null on the next read.
  assert.ok(wb.setComment(0, 1, 1, '', '').ok);
  assert.equal(wb.getComment(0, 1, 1), null);
});

test('getCommentResult distinguishes absence from an invalid sheet', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const missing = wb.getCommentResult(0, 1, 1);
  assert.equal(missing.status.ok, false);
  assert.equal(missing.comment, null);
  const invalid = wb.getCommentResult(99, 1, 1);
  assert.equal(invalid.status.ok, false);
  assert.equal(invalid.comment, null);
  assert.notEqual(missing.status.status, invalid.status.status);
  assert.ok(wb.setComment(0, 1, 1, 'libraz', 'hello').ok);
  const found = wb.getCommentResult(0, 1, 1);
  assert.equal(found.status.ok, true);
  assert.equal(found.comment.author, 'libraz');
  assert.equal(found.comment.text, 'hello');
});

test('cellCount + cellAt enumerate stored cells', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setNumber(0, 0, 0, 1).ok);
  assert.ok(wb.setNumber(0, 0, 1, 2).ok);
  assert.ok(wb.setFormula(0, 1, 0, '=A1+B1').ok);
  assert.ok(wb.recalc().ok);
  const n = wb.cellCount(0);
  assert.equal(typeof n, 'number');
  assert.ok(n >= 3, `expected >= 3 stored cells, got ${n}`);
  // Iterate; one of the entries should carry our formula.
  let foundFormula = false;
  for (let i = 0; i < n; ++i) {
    const e = wb.cellAt(0, i);
    assert.ok(e.status.ok, `cellAt(${i}): ${JSON.stringify(e.status)}`);
    assert.equal(typeof e.row, 'number');
    assert.equal(typeof e.col, 'number');
    if (e.formula === '=A1+B1') {
      foundFormula = true;
      assert.equal(e.value.kind, mod.ValueKind.Number);
      assert.equal(e.value.number, 3);
    }
  }
  assert.ok(foundFormula, 'expected to find the =A1+B1 cell during iteration');
});

test('default export exposes CfMatchKind ordinals', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.CfMatchKind, 'object');
  assert.equal(mod.CfMatchKind.DifferentialFormat, 0);
  assert.equal(mod.CfMatchKind.ColorScale, 1);
  assert.equal(mod.CfMatchKind.DataBar, 2);
  assert.equal(mod.CfMatchKind.IconSet, 3);
});

test('default export exposes ErrorCode ordinals and versionString alias', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.ErrorCode, 'object');
  assert.equal(mod.ErrorCode.Div0, 1);
  assert.equal(mod.ErrorCode.Ref, 3);
  assert.equal(mod.ErrorCode.NA, 6);
  assert.equal(typeof mod.versionString, 'function');
  assert.equal(mod.versionString(), mod.version());
});

test('setError() rejects a missing errorCode argument', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.throws(() => wb.setError(0, 0, 0), TypeError);
});

test('saveAs() writes xlsx and xlsb bytes; rejects a missing format argument', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.WorkbookFormat, 'object');
  const wb = mod.Workbook.createDefault();
  const xlsx = wb.saveAs(mod.WorkbookFormat.Xlsx);
  assert.ok(xlsx.status.ok, JSON.stringify(xlsx.status));
  assert.ok(xlsx.bytes.length > 0);
  const xlsb = wb.saveAs(mod.WorkbookFormat.Xlsb);
  assert.ok(xlsb.status.ok, JSON.stringify(xlsb.status));
  assert.ok(xlsb.bytes.length > 0);
  assert.throws(() => wb.saveAs(), TypeError);
});

test('saveWithDiagnostics() reports counters and readDiagnostics keeps a stable shape', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setFormula(0, 0, 0, '=@A1:A10').ok);
  assert.ok(
    wb.addValidation(0, {
      ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
      type: 3,
      formula1: '"Yes,No"',
    }).ok,
  );
  // A CF rule's dxfId must resolve against a registered dxf.
  const dxf = wb.addDxf({ font: { bold: true } });
  assert.ok(dxf.status.ok, JSON.stringify(dxf.status));
  assert.ok(
    wb.addConditionalFormat(0, {
      sqref: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
      type: 1,
      op: 5,
      formula1: '10',
      dxfId: dxf.index,
      stopIfTrue: false,
    }).status.ok,
  );

  const xlsx = wb.saveWithDiagnostics(mod.WorkbookFormat.Xlsx);
  assert.ok(xlsx.status.ok, JSON.stringify(xlsx.status));
  assert.ok(xlsx.bytes instanceof Uint8Array);
  assert.equal(xlsx.downgradedFormulaCount, 0);
  assert.equal(xlsx.deferredFeatureCount, 0);
  assert.equal(xlsx.droppedPartCount, 0);
  assert.equal(xlsx.droppedRelationshipCount, 0);
  assert.equal(xlsx.renumberedPartCount, 0);

  const xlsb = wb.saveWithDiagnostics(mod.WorkbookFormat.Xlsb);
  assert.ok(xlsb.status.ok, JSON.stringify(xlsb.status));
  assert.ok(xlsb.bytes instanceof Uint8Array);
  assert.equal(xlsb.downgradedFormulaCount, 1);
  assert.ok(xlsb.deferredFeatureCount >= 2);
  // The binary writer never reassigns a part id.
  assert.equal(xlsb.renumberedPartCount, 0);

  const loaded = mod.Workbook.loadBytes(xlsb.bytes);
  const read = loaded.readDiagnostics();
  assert.ok(read.status.ok, JSON.stringify(read.status));
  assert.equal(read.undecodedFormulaCount, 0);
  assert.equal(read.undecodedDefinedNameCount, 0);
  assert.equal(read.undecodedPartCount, 0);
  assert.equal(read.skippedFeatureCount, 0);
  assert.equal(read.unknownContentTypeCount, 0);
  const preservedLoaded = mod.Workbook.loadBytes(appendEmptyZipEntry(xlsb.bytes, 'xl/preserved.bin'));
  const preserved = preservedLoaded.readDiagnostics();
  assert.ok(preserved.status.ok, JSON.stringify(preserved.status));
  assert.equal(preserved.undecodedPartCount, 0);
  preservedLoaded.dispose();
  const invalidSave = wb.saveWithDiagnostics(0);
  assert.ok(!invalidSave.status.ok, JSON.stringify(invalidSave));
  assert.equal(invalidSave.bytes, null);
  assert.equal(invalidSave.downgradedFormulaCount, 0);
  assert.equal(invalidSave.deferredFeatureCount, 0);
  assert.equal(invalidSave.droppedPartCount, 0);
  assert.equal(invalidSave.droppedRelationshipCount, 0);
  assert.equal(invalidSave.renumberedPartCount, 0);
  assert.ok(wb.saveAs(mod.WorkbookFormat.Xlsb).bytes instanceof Uint8Array);
});

test('invalid workbook read diagnostics return a zeroed failure envelope', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.loadBytes(new Uint8Array());
  const read = wb.readDiagnostics();
  assert.ok(!read.status.ok, JSON.stringify(read));
  assert.equal(read.undecodedFormulaCount, 0);
  assert.equal(read.undecodedDefinedNameCount, 0);
  assert.equal(read.undecodedPartCount, 0);
  assert.equal(read.skippedFeatureCount, 0);
  assert.equal(read.unknownContentTypeCount, 0);
  wb.dispose();
});

test('addValidation / getValidations / removeValidationAt / clearValidations round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Empty by default.
  assert.equal(wb.getValidations(0).length, 0);

  // List-type validation at A1:B3.
  const listRule = {
    ranges: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 }],
    type: 3,
    op: 0,
    errorStyle: 1,
    allowBlank: true,
    showInputMessage: true,
    showErrorMessage: true,
    formula1: '"Yes,No,Maybe"',
    promptTitle: 'Choose',
    promptMessage: 'Pick one',
    errorTitle: 'Bad value',
    errorMessage: 'Pick from the list',
  };
  assert.ok(wb.addValidation(0, listRule).ok);

  // Decimal between [0, 100] across two rectangles.
  const decimalRule = {
    ranges: [
      { firstRow: 4, firstCol: 0, lastRow: 4, lastCol: 0 },
      { firstRow: 6, firstCol: 0, lastRow: 9, lastCol: 0 },
    ],
    type: 2,
    op: 0,
    formula1: '0',
    formula2: '100',
    allowBlank: false,
  };
  assert.ok(wb.addValidation(0, decimalRule).ok);

  const list = wb.getValidations(0);
  assert.equal(list.length, 2);
  // Booleans must arrive as JS booleans, not 0/1.
  assert.equal(typeof list[0].allowBlank, 'boolean');
  assert.equal(list[0].allowBlank, true);
  assert.equal(typeof list[0].showInputMessage, 'boolean');
  assert.equal(list[0].showInputMessage, true);
  assert.equal(typeof list[1].allowBlank, 'boolean');
  assert.equal(list[1].allowBlank, false);
  assert.equal(list[0].formula1, '"Yes,No,Maybe"');
  assert.equal(list[1].formula1, '0');
  assert.equal(list[1].formula2, '100');
  assert.equal(list[1].ranges.length, 2);
  assert.equal(list[1].ranges[1].firstRow, 6);

  // removeValidationAt removes the first rule.
  assert.ok(wb.removeValidationAt(0, 0).ok);
  const after = wb.getValidations(0);
  assert.equal(after.length, 1);
  assert.equal(after[0].type, 2);

  // Out-of-range index is rejected.
  assert.ok(!wb.removeValidationAt(0, 99).ok);
  // Sheet-out-of-range is rejected on every entry.
  assert.ok(!wb.addValidation(999, listRule).ok);
  assert.ok(!wb.removeValidationAt(999, 0).ok);
  assert.ok(!wb.clearValidations(999).ok);

  // clearValidations drops everything; safe to call again.
  assert.ok(wb.clearValidations(0).ok);
  assert.equal(wb.getValidations(0).length, 0);
  assert.ok(wb.clearValidations(0).ok);
});

// -- Newly wired surface (audit parity with the WASM binding) ----------
// These tests exercise the methods that were previously only reachable
// from the WASM package. They assert sensible return shapes rather than
// deep semantics; the point is "the addon wires them and does not throw".

test('addConditionalFormat + getConditionalFormats + clearConditionalFormats round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.getConditionalFormats(0).length, 0);
  // A CF rule's dxfId must resolve against a registered dxf.
  const dxf = wb.addDxf({ font: { bold: true } });
  assert.ok(dxf.status.ok, JSON.stringify(dxf.status));
  // cellIs rule (type 1) comparing the cell to a literal.
  const add = wb.addConditionalFormat(0, {
    sqref: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
    type: 1,
    op: 5,
    formula1: '10',
    dxfId: dxf.index,
    stopIfTrue: false,
  });
  assert.ok(add.status.ok, `addConditionalFormat: ${JSON.stringify(add)}`);
  assert.equal(typeof add.index, 'number');
  const list = wb.getConditionalFormats(0);
  assert.equal(list.length, 1);
  assert.equal(list[0].type, 1);
  assert.equal(typeof list[0].id, 'string');
  assert.equal(list[0].sqref.length, 1);
  assert.equal(list[0].sqref[0].lastRow, 4);
  assert.equal(list[0].formula1, '10');
  // removeConditionalFormatAt drops the only rule.
  assert.ok(wb.removeConditionalFormatAt(0, 0).ok);
  assert.equal(wb.getConditionalFormats(0).length, 0);
  // clearConditionalFormats is a safe no-op on an empty sheet.
  assert.ok(wb.clearConditionalFormats(0).ok);
  // Sheet-out-of-range is rejected.
  assert.ok(!wb.clearConditionalFormats(999).ok);
});

test('data bar x14 fields survive save and load', async () => {
  // The six live in the `x14` extension rather than the legacy `<dataBar>`
  // element, so an in-session round-trip would still pass with a writer
  // that never emits the extension -- which is exactly how they used to be
  // lost. Only a save/load cycle pins them.
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const add = wb.addConditionalFormat(0, {
    sqref: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
    type: 3,
    dataBar: {
      min: { type: 3 },
      max: { type: 4 },
      fill: { r: 0, g: 0, b: 255 },
      gradient: false,
      axisPosition: 1,
      negativeFill: { r: 255, g: 0, b: 0 },
      border: { r: 9, g: 9, b: 9 },
      negativeBorder: { r: 8, g: 8, b: 8 },
      axisColor: { r: 1, g: 2, b: 3 },
    },
  });
  assert.ok(add.status.ok, `addConditionalFormat: ${JSON.stringify(add)}`);

  const saved = wb.save();
  assert.ok(saved.status.ok);
  const loaded = mod.Workbook.loadBytes(saved.bytes);
  const bar = loaded.getConditionalFormats(0)[0].dataBar;
  assert.equal(bar.gradient, false);
  assert.equal(bar.axisPosition, 1);
  assert.equal(bar.negativeFill.r, 255);
  assert.equal(bar.border.r, 9);
  assert.equal(bar.negativeBorder.r, 8);
  assert.equal(bar.axisColor.b, 3);

  // The decoded bar fed straight back must reproduce the same rule.
  const again = loaded.addConditionalFormat(0, {
    sqref: [{ firstRow: 4, firstCol: 0, lastRow: 6, lastCol: 0 }],
    type: 3,
    dataBar: bar,
  });
  assert.ok(again.status.ok, `re-add: ${JSON.stringify(again)}`);
  const reread = loaded.getConditionalFormats(0)[1].dataBar;
  assert.equal(reread.gradient, false);
  assert.equal(reread.axisPosition, 1);
  assert.equal(reread.axisColor.b, 3);

  loaded.dispose();
  wb.dispose();
});

test('omitted data bar x14 fields keep the model defaults', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const add = wb.addConditionalFormat(0, {
    sqref: [{ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 0 }],
    type: 3,
    dataBar: { min: { type: 3 }, max: { type: 4 }, fill: { r: 0, g: 0, b: 255 } },
  });
  assert.ok(add.status.ok, `addConditionalFormat: ${JSON.stringify(add)}`);
  const bar = wb.getConditionalFormats(0)[0].dataBar;
  // The getter engages all six, so they read back as the defaults rather
  // than as absent keys.
  assert.equal(bar.gradient, true);
  assert.equal(bar.axisPosition, 0);
  assert.equal(bar.negativeFill.b, 255);
  wb.dispose();
});

test('named cell styles: cellStyleCount / getCellStyle / cellStyleXfCount', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A fresh workbook may carry zero named styles; the accessors must
  // still return well-formed values rather than throwing.
  const count = wb.cellStyleCount();
  assert.equal(typeof count, 'number');
  assert.ok(count >= 0);
  assert.equal(typeof wb.cellStyleXfCount(), 'number');
  if (count > 0) {
    const cs = wb.getCellStyle(0);
    assert.ok(cs.status.ok, `getCellStyle: ${JSON.stringify(cs.status)}`);
    assert.equal(typeof cs.name, 'string');
    assert.equal(typeof cs.xfId, 'number');
  } else {
    // Out-of-range index surfaces an error status, not a throw.
    const cs = wb.getCellStyle(0);
    assert.equal(cs.status.ok, false);
  }
});

test('setCalcMode / calcMode round-trip the calc policy', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Default is automatic (0).
  assert.equal(wb.calcMode(), 0);
  assert.ok(wb.setCalcMode(1).ok);
  assert.equal(wb.calcMode(), 1);
  assert.ok(wb.setCalcMode(2).ok);
  assert.equal(wb.calcMode(), 2);
});

test('setPinnedNow / pinnedNow / clearPinnedNow drive the clock seam', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Unpinned by default: the workbook follows the host clock.
  assert.equal(wb.pinnedNow(), null);

  assert.ok(wb.setPinnedNow(2026, 4, 23, 15, 30, 45).ok);
  assert.deepEqual(wb.pinnedNow(), {
    year: 2026,
    month: 4,
    day: 23,
    hour: 15,
    minute: 30,
    second: 45,
  });

  // 2026-04-23 is serial 46135 under the 1900 date system.
  const today = wb.evaluateFormulaText(0, 0, 0, '=TODAY()');
  assert.ok(today.status.ok, `evaluateFormulaText: ${JSON.stringify(today.status)}`);
  assert.equal(today.value.number, 46135);

  // The pin is a calendar instant, not a normalising constructor.
  assert.equal(wb.setPinnedNow(2026, 13, 1, 0, 0, 0).ok, false);
  assert.equal(wb.setPinnedNow(2025, 2, 29, 0, 0, 0).ok, false);

  assert.ok(wb.clearPinnedNow().ok);
  assert.equal(wb.pinnedNow(), null);
});

test('excelProfileId / setExcelProfileId round-trip the profile id', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const def = wb.excelProfileId();
  assert.equal(typeof def, 'string');
  assert.ok(def.length > 0);
  assert.ok(wb.setExcelProfileId('mac-365-ja_JP').ok);
  assert.equal(wb.excelProfileId(), 'mac-365-ja_JP');
});

test('functionNames + functionMetadata expose the catalog', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const names = wb.functionNames();
  assert.ok(Array.isArray(names));
  assert.ok(names.length > 0, 'expected a non-empty function catalog');
  assert.ok(names.includes('SUM'), 'expected SUM in the catalog');
  const md = wb.functionMetadata('SUM', 0);
  assert.equal(md.ok, true);
  assert.equal(md.name, 'SUM');
  assert.equal(typeof md.minArity, 'number');
  // Unknown function returns { ok: false }.
  const miss = wb.functionMetadata('NOT_A_REAL_FUNCTION_XYZ', 0);
  assert.equal(miss.ok, false);
});

test('mergeFunctionMetadata overlays provider metadata; is identity without a provider', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.mergeFunctionMetadata, 'function');
  const base = mod.Workbook.createDefault().functionMetadata('XLOOKUP', 0);
  assert.equal(base.ok, true);
  // The engine leaves display metadata empty.
  assert.equal(base.signatureTemplate, undefined);
  assert.equal(base.description, undefined);

  const entry = {
    signature: 'XLOOKUP(lookup_value, lookup_array, return_array)',
    description: 'Searches a range or an array.',
    aliases: { 'fr-FR': 'RECHERCHEX' },
    localized: {
      'fr-FR': { signature: 'RECHERCHEX(...)', description: 'Recherche.' },
    },
  };

  // Localized override wins for the matching locale.
  const fr = mod.mergeFunctionMetadata(base, entry, 'fr-FR');
  assert.equal(fr.signatureTemplate, 'RECHERCHEX(...)');
  assert.equal(fr.description, 'Recherche.');
  assert.equal(fr.localizedName, 'RECHERCHEX');
  // Structural fields survive the merge.
  assert.equal(fr.ok, true);
  assert.equal(fr.name, 'XLOOKUP');

  // A locale with no localized/alias entry falls back to the default
  // signature/description and the canonical display name.
  const de = mod.mergeFunctionMetadata(base, entry, 'de-DE');
  assert.equal(de.signatureTemplate, 'XLOOKUP(lookup_value, lookup_array, return_array)');
  assert.equal(de.description, 'Searches a range or an array.');
  assert.equal(de.localizedName, 'XLOOKUP');

  // No provider entry -> base returned verbatim, display metadata stays NULL.
  const none = mod.mergeFunctionMetadata(base, undefined, 'fr-FR');
  assert.equal(none, base);
  assert.equal(none.signatureTemplate, undefined);
  assert.equal(none.description, undefined);
});

test('localizeFunctionName / canonicalizeFunctionName round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // en-US locale (0): the localized name is the canonical name itself.
  assert.equal(wb.localizeFunctionName('SUM', 0), 'SUM');
  assert.equal(wb.canonicalizeFunctionName('SUM', 0), 'SUM');
  // An unknown name returns the empty string.
  assert.equal(wb.canonicalizeFunctionName('NOPE_XYZ', 0), '');
});

test('precedents / dependents return arrays for a small formula graph', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.ok(wb.setNumber(0, 0, 0, 1).ok); // A1
  assert.ok(wb.setNumber(0, 0, 1, 2).ok); // B1
  assert.ok(wb.setFormula(0, 1, 0, '=A1+B1').ok); // A2 = A1 + B1
  assert.ok(wb.recalc().ok);
  // A2 reads A1 and B1.
  const prec = wb.precedents(0, 1, 0, 1);
  assert.ok(Array.isArray(prec));
  assert.equal(prec.length, 2);
  for (const node of prec) {
    assert.equal(typeof node.sheet, 'number');
    assert.equal(typeof node.row, 'number');
    assert.equal(typeof node.col, 'number');
  }
  // A1 is read by A2.
  const dep = wb.dependents(0, 0, 0, 1);
  assert.ok(Array.isArray(dep));
  assert.equal(dep.length, 1);
  assert.equal(dep[0].row, 1);
  assert.equal(dep[0].col, 0);
});

test('spillInfo reports a dynamic-array region', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Non-spill cell: engaged is false, fields zeroed.
  const none = wb.spillInfo(0, 0, 0);
  assert.equal(typeof none.engaged, 'boolean');
  assert.equal(none.engaged, false);
  // A spilling dynamic array anchored at A1.
  assert.ok(wb.setFormula(0, 0, 0, '=SEQUENCE(3,1)').ok);
  assert.ok(wb.recalc().ok);
  const info = wb.spillInfo(0, 0, 0);
  assert.equal(info.engaged, true);
  assert.equal(info.anchorRow, 0);
  assert.equal(info.anchorCol, 0);
  assert.equal(info.rows, 3);
  assert.equal(info.cols, 1);
});

test('getLambdaText is wired and reports the documented contract', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // A bare top-level LAMBDA renders as #CALC! in Excel 365 and does not
  // retain a renderable closure; the engine mirrors that, so the lambda
  // text surface returns an error status (not a throw). A plain number
  // cell is likewise not a lambda. Both prove the method marshals
  // arguments and propagates the Expected failure as a JS error status.
  assert.ok(wb.setFormula(0, 0, 0, '=LAMBDA(x,x+1)').ok);
  assert.ok(wb.recalc().ok);
  const bare = wb.getLambdaText(0, 0, 0);
  assert.equal(bare.status.ok, false);
  assert.equal(typeof bare.text, 'string');
  assert.equal(bare.text, '');

  assert.ok(wb.setNumber(0, 1, 0, 5).ok);
  const num = wb.getLambdaText(0, 1, 0);
  assert.equal(num.status.ok, false);
});

test('getExternalLinks returns an array (empty for a fresh workbook)', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const links = wb.getExternalLinks();
  assert.ok(Array.isArray(links));
  assert.equal(links.length, 0);
});

test('getSheetProtection / setSheetProtection round-trip the protection flags', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const before = wb.getSheetProtection(0);
  assert.ok(before.status.ok, `getSheetProtection: ${JSON.stringify(before.status)}`);
  assert.equal(before.protection.enabled, 0);
  // Enable protection and lock cell selection.
  const next = {
    ...before.protection,
    enabled: 1,
    selectLockedCells: 1,
    sort: 1,
  };
  assert.ok(wb.setSheetProtection(0, next).ok);
  const after = wb.getSheetProtection(0);
  assert.ok(after.status.ok);
  assert.equal(after.protection.enabled, 1);
  assert.equal(after.protection.selectLockedCells, 1);
  assert.equal(after.protection.sort, 1);
  // Sheet-out-of-range is rejected.
  assert.ok(!wb.getSheetProtection(999).status.ok);
  assert.ok(!wb.setSheetProtection(999, next).ok);
});

// ---- Result-envelope key sets ------------------------------------------
//
// A result type whose payload key is declared non-optional must carry that
// key on every exit path, or `r.name.toUpperCase()` type-checks and then
// throws at runtime for a caller who skipped the status check. The probe
// table is checked for completeness against the declaration file, so a new
// result type cannot be added without a failure-path probe. The WASM
// package runs the mirror image of this test.

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

// `fails: false` marks a call whose only failure path is a disposed
// handle, which is not usable to probe.
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
    ['PhoneticRunsResult', true, () => wb.getCellPhoneticRuns(99, 0, 0)],
    ['CellStyleResult', true, () => wb.getCellStyle(9999)],
    ['AddStyleResult', true, () => wb.addXf({ fontIndex: 9999 })],
    ['AddNumFmtResult', false, () => wb.addNumFmt('0.00')],
    ['IterativeSettingsResult', false, () => wb.getIterative()],
    [
      'PartialRecalcResult',
      true,
      () => wb.partialRecalc({ sheet: 99, firstRow: 0, lastRow: 1, firstCol: 0, lastCol: 1 }),
    ],
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
  const dts = await readFile(path.join(pkgRoot, 'dist', 'index.d.ts'), 'utf8');
  const declared = resultTypesWithRequiredPayload(dts);
  const probed = new Set(envelopeProbes({}).map(([name]) => name));
  const unprobed = [...declared.keys()].filter((name) => !probed.has(name));
  assert.deepEqual(unprobed, [], 'result types with no failure-path probe');
  const stale = [...probed].filter((name) => !declared.has(name));
  assert.deepEqual(stale, [], 'probes for result types that no longer declare a required payload');
});

test('result envelopes keep their declared keys on failure paths', async () => {
  const mod = await getModule();
  const dts = await readFile(path.join(pkgRoot, 'dist', 'index.d.ts'), 'utf8');
  const declared = resultTypesWithRequiredPayload(dts);
  const wb = mod.Workbook.createDefault();
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
    wb.dispose();
  }
});

test('list getters carry a status and stop at the first element failure', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    // A readable sheet: every list is empty but the enumeration succeeded,
    // which is what distinguishes it from a sheet that could not be read.
    for (const [name, call] of [
      ['getMerges', () => wb.getMerges(0)],
      ['getComments', () => wb.getComments(0)],
      ['getHyperlinks', () => wb.getHyperlinks(0)],
      ['getValidations', () => wb.getValidations(0)],
      ['getConditionalFormats', () => wb.getConditionalFormats(0)],
      ['getExternalLinks', () => wb.getExternalLinks()],
    ]) {
      const list = call();
      assert.ok(Array.isArray(list), `${name} must stay an array`);
      assert.equal(list.length, 0, name);
      assert.ok(list.status.ok, `${name}: ${JSON.stringify(list.status)}`);
    }

    // An unreadable sheet is reported, not silently rendered as "empty".
    for (const [name, call] of [
      ['getMerges', () => wb.getMerges(99)],
      ['getComments', () => wb.getComments(99)],
      ['getHyperlinks', () => wb.getHyperlinks(99)],
      ['getValidations', () => wb.getValidations(99)],
      ['getConditionalFormats', () => wb.getConditionalFormats(99)],
    ]) {
      const list = call();
      assert.ok(Array.isArray(list), `${name} must stay an array`);
      assert.equal(list.status.ok, false, `${name} must signal an unreadable sheet`);
      assert.notEqual(list.status.status, 0, name);
    }
  } finally {
    wb.dispose();
  }
});

test('a throwing iterative progress callback aborts the solve and reports a status', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(wb.setIterative(true, 100, 0.0001).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
    let calls = 0;
    assert.ok(
      wb.setIterativeProgress(() => {
        calls += 1;
        throw new Error('callback failure');
      }).ok,
    );

    // Reported as a status rather than propagated, and with the same
    // ordinal the WASM binding uses (`kBindingCallbackException`).
    const aborted = wb.recalc();
    assert.equal(aborted.ok, false, `expected a failure envelope: ${JSON.stringify(aborted)}`);
    assert.equal(aborted.status, 7003);
    assert.equal(calls, 1, 'the solve must stop at the first throw');

    const mid = wb.getValue(0, 0, 0);
    assert.ok(mid.status.ok, JSON.stringify(mid.status));

    assert.ok(wb.setIterativeProgress(null).ok);
    assert.ok(wb.setFormula(0, 0, 0, '=(A1+10)/2').ok);
    const recovered = wb.recalc();
    assert.ok(recovered.ok, `recalc after the throwing callback: ${JSON.stringify(recovered)}`);
    const converged = wb.getValue(0, 0, 0);
    assert.ok(Math.abs(converged.value.number - 10) < 0.001, `A1=${converged.value.number}`);
  } finally {
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

// `CfValueObjectInput.type`: 0 num, 1 percent, 2 percentile, 3 min,
// 4 max, 5 formula, 6 autoMin, 7 autoMax. Every threshold below states a
// `value`, which is what puts a borrowed string behind each CFVO.
const COLOR_SCALE_RULE = Object.freeze({
  sqref: [{ firstRow: 0, firstCol: 0, lastRow: 4, lastCol: 0 }],
  type: 2,
  colorScale: {
    thresholds: [
      { type: 0, value: '0' },
      { type: 1, value: '50' },
      { type: 0, value: '100' },
    ],
    colors: [
      { r: 248, g: 105, b: 107, a: 255 },
      { r: 255, g: 235, b: 132, a: 255 },
      { r: 99, g: 190, b: 123, a: 255 },
    ],
  },
});

const DATA_BAR_RULE = Object.freeze({
  sqref: [{ firstRow: 0, firstCol: 1, lastRow: 4, lastCol: 1 }],
  type: 3,
  dataBar: {
    min: { type: 0, value: '5' },
    max: { type: 0, value: '95' },
    fill: { r: 0, g: 112, b: 192, a: 255 },
  },
});

const ICON_SET_RULE = Object.freeze({
  sqref: [{ firstRow: 0, firstCol: 2, lastRow: 4, lastCol: 2 }],
  type: 4,
  iconSet: {
    name: 0,
    thresholds: [
      { type: 1, value: '0' },
      { type: 1, value: '33' },
      { type: 1, value: '67' },
    ],
  },
});

const EXPECTED_CFVO_VALUES = Object.freeze(['0', '50', '100', '5', '95', '0', '33', '67']);

// The threshold value strings of the three visual rules, flattened.
function cfvoValues(rules) {
  const byType = new Map(rules.map((rule) => [rule.type, rule]));
  const colorScale = byType.get(2);
  const dataBar = byType.get(3);
  const iconSet = byType.get(4);
  assert.ok(colorScale?.colorScale, 'colorScale payload missing');
  assert.ok(dataBar?.dataBar, 'dataBar payload missing');
  assert.ok(iconSet?.iconSet, 'iconSet payload missing');
  return [
    ...colorScale.colorScale.thresholds.map((t) => t.value),
    dataBar.dataBar.min.value,
    dataBar.dataBar.max.value,
    ...iconSet.iconSet.thresholds.map((t) => t.value),
  ];
}

test('CFVO value strings survive getConditionalFormats and a save/load cycle', async () => {
  // Each CFVO `value` crosses the C ABI as a borrowed `const char*`. A
  // store that relocates while the later thresholds are pulled publishes
  // a pointer into freed bytes, so the strings come back wrong or empty
  // even though the rule count and colors still look right. Rules whose
  // thresholds are all `min` / `max` never take that path, because those
  // types carry no value string.
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  let bytes;
  try {
    for (const rule of [COLOR_SCALE_RULE, DATA_BAR_RULE, ICON_SET_RULE]) {
      const added = wb.addConditionalFormat(0, rule);
      assert.ok(added.status.ok, `addConditionalFormat type=${rule.type}: ${JSON.stringify(added.status)}`);
    }

    const inSession = wb.getConditionalFormats(0);
    assert.ok(inSession.status.ok, JSON.stringify(inSession.status));
    assert.equal(inSession.length, 3);
    assert.deepEqual(cfvoValues(inSession), EXPECTED_CFVO_VALUES);
    const byType = new Map(inSession.map((rule) => [rule.type, rule]));
    // The type travels with the value: a percent threshold read back as
    // type 0 would mean an absolute number instead.
    assert.deepEqual(
      byType.get(2).colorScale.thresholds.map((t) => t.type),
      [0, 1, 0],
    );
    assert.deepEqual(
      byType.get(4).iconSet.thresholds.map((t) => t.type),
      [1, 1, 1],
    );
    assert.equal(byType.get(2).colorScale.colors.length, 3);
    assert.equal(byType.get(2).colorScale.colors[2].g, 190);
    assert.equal(byType.get(3).dataBar.fill.b, 192);

    const saved = wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    bytes = saved.bytes;
  } finally {
    wb.dispose();
  }

  const reloaded = mod.Workbook.loadBytes(bytes);
  try {
    const rules = reloaded.getConditionalFormats(0);
    assert.equal(rules.length, 3);
    assert.deepEqual(cfvoValues(rules), EXPECTED_CFVO_VALUES);
  } finally {
    reloaded.dispose();
  }
});

test('getIterative reads back what setIterative stored', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const initial = wb.getIterative();
    assert.ok(initial.status.ok, `getIterative: ${JSON.stringify(initial.status)}`);
    assert.equal(typeof initial.enabled, 'boolean');
    assert.equal(typeof initial.maxIterations, 'number');
    assert.equal(typeof initial.maxChange, 'number');

    assert.ok(wb.setIterative(true, 42, 0.25).ok);
    const enabled = wb.getIterative();
    assert.ok(enabled.status.ok);
    assert.equal(enabled.enabled, true);
    assert.equal(enabled.maxIterations, 42);
    assert.equal(enabled.maxChange, 0.25);

    // The cap and threshold survive switching iteration back off, which is
    // what lets a host render the dialog with the stored values.
    assert.ok(wb.setIterative(false, 42, 0.25).ok);
    const disabled = wb.getIterative();
    assert.ok(disabled.status.ok);
    assert.equal(disabled.enabled, false);
    assert.equal(disabled.maxIterations, 42);
    assert.equal(disabled.maxChange, 0.25);
  } finally {
    wb.dispose();
  }
});

test('addValidation defaults allowBlank to false', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    assert.ok(
      wb.addValidation(0, {
        ranges: [{ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }],
        type: 3,
        formula1: '"a,b"',
      }).ok,
    );
    const list = wb.getValidations(0);
    assert.equal(list.length, 1);
    assert.equal(list[0].allowBlank, false);
  } finally {
    wb.dispose();
  }
});

test('getCellStyleXf reports xfId alongside the cell-format fields', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const read = wb.getCellStyleXf(0);
    assert.ok(read.status.ok, `getCellStyleXf: ${JSON.stringify(read.status)}`);
    // A named-style xf is its own parent, so it never inherits one.
    assert.equal(read.xfId, 0);
    // The record shape matches `getCellXf`, which the documented
    // read / edit / addCellStyleXf round-trip depends on.
    const cellXf = wb.getCellXf(0);
    assert.ok(cellXf.status.ok);
    for (const key of Object.keys(cellXf)) {
      assert.ok(key in read, `getCellStyleXf is missing '${key}'`);
    }
  } finally {
    wb.dispose();
  }
});

test('pivotFieldAddItemAt hides the blank item an empty label cannot name', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const cache = wb.pivotCacheCreate(0);
    assert.ok(cache.status.ok, `pivotCacheCreate: ${JSON.stringify(cache.status)}`);
    const region = wb.pivotCacheFieldAdd(cache.index, 'Region');
    assert.ok(region.status.ok);
    const amount = wb.pivotCacheFieldAdd(cache.index, 'Amount');
    assert.ok(amount.status.ok);
    // Shared item 0 renders as "North"; shared item 1 is the blank, which
    // carries no label of its own and can only be named by its index.
    assert.ok(wb.pivotCacheFieldAddSharedItemText(cache.index, region.index, 'North').ok);
    assert.ok(wb.pivotCacheFieldAddSharedItemBlank(cache.index, region.index).ok);

    const first = wb.pivotCacheRecordAdd(cache.index);
    assert.ok(first.status.ok);
    assert.ok(wb.pivotCacheRecordSetNumber(cache.index, first.index, region.index, 0).ok);
    assert.ok(wb.pivotCacheRecordSetNumber(cache.index, first.index, amount.index, 100).ok);
    const second = wb.pivotCacheRecordAdd(cache.index);
    assert.ok(second.status.ok);
    assert.ok(wb.pivotCacheRecordSetBlank(cache.index, second.index, region.index).ok);
    assert.ok(wb.pivotCacheRecordSetNumber(cache.index, second.index, amount.index, 200).ok);

    const pivot = wb.pivotCreate(0, 'PT', cache.index, 0, 0);
    assert.ok(pivot.status.ok, `pivotCreate: ${JSON.stringify(pivot.status)}`);
    const regionField = wb.pivotFieldAdd(0, pivot.index, { sourceName: 'Region', axis: 0 });
    assert.ok(regionField.status.ok);
    const amountField = wb.pivotFieldAdd(0, pivot.index, { sourceName: 'Amount', axis: 2 });
    assert.ok(amountField.status.ok);
    assert.ok(wb.pivotSetRowFieldOrder(0, pivot.index, [regionField.index]).ok);
    assert.ok(
      wb.pivotDataFieldAdd(0, pivot.index, {
        name: 'Sum of Amount',
        fieldIndex: amountField.index,
        aggregation: 0,
      }).status.ok,
    );

    const sumDataCells = () => {
      const layout = wb.pivotLayout(0, pivot.index);
      assert.ok(layout.status.ok, `pivotLayout: ${JSON.stringify(layout.status)}`);
      let total = 0;
      for (const cell of layout.cells) {
        if (cell.kind === mod.PivotCellKind.Data && cell.value.kind === mod.ValueKind.Number) {
          total += cell.value.number;
        }
      }
      return total;
    };
    assert.equal(sumDataCells(), 300);

    assert.ok(wb.pivotFieldAddItem(0, pivot.index, regionField.index, 'North', true).ok);
    assert.ok(wb.pivotFieldAddItemAt(0, pivot.index, regionField.index, 1, false).ok);
    assert.equal(sumDataCells(), 100);
  } finally {
    wb.dispose();
  }
});

test('a rejected argument reports its own diagnostic, not the previous call', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    // Leave a failure of a different call in the thread-local diagnostics.
    const stale = wb.getValue(99, 0, 0);
    assert.equal(stale.status.ok, false);
    assert.ok(stale.status.message.length > 0);

    // An argument the binding rejects before any C call carries its own
    // message, and a code that names an argument fault rather than the
    // destroyed-handle code.
    const missingSpec = wb.pivotFieldAdd(0, 0);
    assert.equal(missingSpec.status.ok, false);
    assert.equal(missingSpec.status.status, 7001);
    assert.match(missingSpec.status.message, /pivotFieldAdd/);
    assert.notEqual(missingSpec.status.message, stale.status.message);
    assert.equal(missingSpec.status.context, '');

    const badCallback = wb.setIterativeProgress(42);
    assert.equal(badCallback.ok, false);
    assert.equal(badCallback.status, 7001);
    assert.match(badCallback.message, /callback/);
  } finally {
    wb.setIterativeProgress(null);
    wb.dispose();
  }
});

test('a call on a disposed workbook names the destroyed handle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.dispose();
  // 7000 is reserved for this: an argument fault reports 7001 instead, so
  // "make a new workbook" stays the correct recovery for exactly this code.
  const after = wb.recalc();
  assert.equal(after.ok, false);
  assert.equal(after.status, 7000);
  assert.ok(after.message.length > 0, 'the destroyed-handle status must carry a message');
});

test('sheet-layout setters refuse coordinates and metrics Excel cannot open', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    // These six setters marshal straight to the C ABI, which is where the
    // validation lives. The assertions below are what keeps that true: a
    // binding that grew its own path to the model would still pass every
    // round-trip test while quietly authoring a file Excel refuses, since
    // NaN and Infinity serialise to an empty string -- `<col width=""/>`
    // is not a lexical xsd:double at all, and a repair prompt rather than
    // an out-of-range value.
    const kInvalidArgument = 2;
    // The message is pinned, not just the code. Both JS bindings route
    // these calls to one set of C ABI helpers, so both report that shared
    // wording; a binding that grew its own check would keep returning
    // `kInvalidArgument` while the text stopped matching. The status code
    // alone cannot tell a shared implementation from two that agree today.
    const kSpanRefused = 'column span out of range';
    const kRowRefused = 'row index out of range';
    const kMetricRefused = 'value must be finite and non-negative';
    // These six setters answer with a bare Status, not a `{status}`
    // envelope, so the diagnostics are read off the returned object.
    const refused = (status, expectedMessage, label) => {
      assert.equal(status.ok, false, `${label} should have been refused`);
      assert.equal(status.status, kInvalidArgument, `${label}: ${JSON.stringify(status)}`);
      assert.equal(status.message, expectedMessage, `${label}: ${JSON.stringify(status)}`);
    };

    // One past Excel's last column / row.
    refused(wb.setColumnWidth(0, 0, 16384, 10), kSpanRefused, 'setColumnWidth past the grid');
    refused(wb.setColumnHidden(0, 0, 16384, true), kSpanRefused, 'setColumnHidden past the grid');
    refused(wb.setColumnOutline(0, 0, 16384, 1), kSpanRefused, 'setColumnOutline past the grid');
    refused(wb.setRowHeight(0, 1048576, 10), kRowRefused, 'setRowHeight past the grid');
    refused(wb.setRowHidden(0, 1048576, true), kRowRefused, 'setRowHidden past the grid');
    refused(wb.setRowOutline(0, 1048576, 1), kRowRefused, 'setRowOutline past the grid');

    for (const bad of [Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY, -1]) {
      refused(wb.setColumnWidth(0, 0, 2, bad), kMetricRefused, `setColumnWidth(${bad})`);
      refused(wb.setRowHeight(0, 1, bad), kMetricRefused, `setRowHeight(${bad})`);
    }

    // The context names the C ABI entry point that refused the call and
    // the values it saw, which is what makes a rejection actionable.
    const span = wb.setColumnWidth(0, 0, 16384, 10);
    assert.match(span.context, /fm_sheet_set_column_width: first=0 last=16384/);
    const metric = wb.setRowHeight(0, 1, Number.NaN);
    assert.match(metric.context, /fm_sheet_set_row_height: height=/);

    // Zero is a real width and a real height, not a rejected one.
    assert.ok(wb.setColumnWidth(0, 0, 2, 0).ok);
    assert.ok(wb.setRowHeight(0, 1, 0).ok);

    // The last coordinates Excel does address stay accepted -- an
    // off-by-one in either bound would refuse them.
    assert.ok(wb.setColumnWidth(0, 16383, 16383, 10).ok);
    assert.ok(wb.setColumnHidden(0, 16383, 16383, true).ok);
    assert.ok(wb.setColumnOutline(0, 16383, 16383, 1).ok);
    assert.ok(wb.setRowHeight(0, 1048575, 10).ok);
    assert.ok(wb.setRowHidden(0, 1048575, true).ok);
    assert.ok(wb.setRowOutline(0, 1048575, 1).ok);
  } finally {
    wb.dispose();
  }
});

test('setSheetVisibility states veryHidden, which the bool setter cannot', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const read = () => {
      const r = wb.getSheetView(0);
      assert.ok(r.status.ok, `getSheetView: ${JSON.stringify(r.status)}`);
      return r.view;
    };
    assert.equal(read().visibility, mod.SheetVisibility.Visible);
    assert.equal(read().tabHidden, 0);

    assert.ok(wb.setSheetVisibility(0, mod.SheetVisibility.VeryHidden).ok);
    // The two-state view stays consistent: a very-hidden sheet reads as
    // hidden to a caller that knows only the bool, never as visible.
    assert.equal(read().visibility, mod.SheetVisibility.VeryHidden);
    assert.equal(read().tabHidden, 1);

    // Asking for plain "hidden" must not weaken the stronger state.
    assert.ok(wb.setSheetTabHidden(0, true).ok);
    assert.equal(read().visibility, mod.SheetVisibility.VeryHidden);

    // Demotion is the other direction the bool cannot express.
    assert.ok(wb.setSheetVisibility(0, mod.SheetVisibility.Hidden).ok);
    assert.equal(read().visibility, mod.SheetVisibility.Hidden);
    assert.equal(read().tabHidden, 1);

    // Showing the sheet clears either hidden state.
    assert.ok(wb.setSheetTabHidden(0, false).ok);
    assert.equal(read().visibility, mod.SheetVisibility.Visible);
    assert.equal(read().tabHidden, 0);

    // An unknown ordinal is refused and leaves the sheet alone.
    assert.ok(wb.setSheetVisibility(0, mod.SheetVisibility.VeryHidden).ok);
    assert.equal(wb.setSheetVisibility(0, 3).ok, false);
    assert.equal(read().visibility, mod.SheetVisibility.VeryHidden);
  } finally {
    wb.dispose();
  }
});
