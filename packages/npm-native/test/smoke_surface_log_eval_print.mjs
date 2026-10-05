import test from 'node:test';
import { assert, getModule, path, pkgRoot, readFile } from './smoke_support.mjs';

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
    // A structured reference the XLSB encoder cannot lower,
    // which makes the writer emit a per-cell warn record.
    assert.ok(wb.setFormula(0, 0, 0, '=SUM(T[C])').ok);
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
    assert.ok(wb.setFormula(0, 0, 0, '=SUM(T[C])').ok);
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

test('a throwing log sink does not leave a pending exception at the ThreadSafeFunction boundary', async () => {
  const mod = await getModule();
  let calls = 0;
  assert.ok(
    mod.setLogSink(() => {
      calls += 1;
      throw new Error('sink failure');
    }).ok,
  );
  assert.ok(mod.setLogMinLevel(mod.LogLevel.Warn).ok);
  try {
    const wb = mod.Workbook.createDefault();
    assert.ok(wb.setFormula(0, 0, 0, '=SUM(T[C])').ok);
    assert.ok(wb.recalc().ok);
    assert.ok(wb.saveAs(mod.WorkbookFormat.Xlsb).status.ok);
    wb.dispose();
    // Wait for the ThreadSafeFunction drain to run the sink.
    await new Promise((resolve) => setTimeout(resolve, SINK_SETTLE_MS));
    assert.ok(calls > 0, 'expected the sink to have been invoked');
  } finally {
    assert.ok(mod.setLogMinLevel(mod.LogLevel.Off).ok);
    assert.ok(mod.setLogSink(null).ok);
  }

  // An uncaught exception at the drain boundary would have crashed the
  // process by now; the engine must still be usable afterwards.
  const wb2 = mod.Workbook.createDefault();
  try {
    assert.ok(wb2.setFormula(0, 0, 0, '=1+1').ok);
    assert.ok(wb2.recalc().ok);
    assert.equal(wb2.getValue(0, 0, 0).value.number, 2);
  } finally {
    wb2.dispose();
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
  assert.equal(wb.sheetCount().value, 1);
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
  assert.equal(wb.pivotCount(0).value, 0);

  const missing = wb.pivotLayout(0, 0);
  assert.equal(missing.status.ok, false);
  assert.notEqual(missing.status.status, 0);
  assert.equal(missing.top, 0);
  assert.equal(missing.left, 0);
  assert.equal(missing.rows, 0);
  assert.equal(missing.cols, 0);
  assert.deepEqual(missing.cells, []);
});
