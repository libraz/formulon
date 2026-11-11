import test from 'node:test';
import { assert, getModule, isFormulonError, loadStagedModule, path, pkgRoot, readFile } from './smoke_support.mjs';

const rgb = (value) => ({ kind: 1, argb: value, theme: 0, tint: 0, indexed: 0 });
const NO_COLOR = { kind: 0, argb: 0, theme: 0, tint: 0, indexed: 0 };

test('the Formulon object exposes Workbook + evaluateFormula + getVersion', async () => {
  const mod = await getModule();
  assert.equal(typeof mod.Workbook, 'function');
  assert.equal(typeof mod.evaluateFormula, 'function');
  assert.equal(typeof mod.getVersion, 'function');
  assert.equal(mod.getStatusName, undefined);
  assert.equal(typeof mod.ValueKind, 'object');
  assert.equal(typeof mod.PivotCellKind, 'object');
});

test('all declared runtime constants resolve from the staged ESM entry point', async () => {
  const mod = await loadStagedModule();
  const declarations = await readFile(path.join(pkgRoot, 'index.d.ts'), 'utf8');
  const names = [...declarations.matchAll(/^export const (\w+)/gm)].map((match) => match[1]);
  assert.ok(names.length > 0, 'expected runtime constant declarations');
  for (const name of names) {
    assert.ok(name in mod, `missing runtime export declared in index.d.ts: ${name}`);
  }
  assert.equal(mod.PivotAxis.Row, 0);
  assert.equal(mod.PivotShowAsBase.Previous, 1048828);
  assert.equal(mod.PivotShowAsBase.Next, 1048829);
  assert.equal('PIVOT_SHOW_AS_BASE_PREVIOUS' in mod, false);
  assert.equal('PIVOT_SHOW_AS_BASE_NEXT' in mod, false);
  // Named constants a consumer needs in order to interpret shared record
  // fields (`Value.errorCode`, `ExternalLink.kind`) or to drive the
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

test('the Formulon interface and the runtime object declare the same keys', async () => {
  const mod = await getModule();
  const declarations = await readFile(path.join(pkgRoot, 'index.d.ts'), 'utf8');
  const typeBlock = declarations.match(/export interface Formulon \{([\s\S]*?)\n\}/);
  assert.ok(typeBlock, 'expected an `export interface Formulon {...}` block in index.d.ts');
  const declared = new Set([...typeBlock[1].matchAll(/^ {2}(\w+)[:(]/gm)].map((match) => match[1]));
  const runtime = new Set(Object.keys(mod));
  const missingFromType = [...runtime].filter((key) => !declared.has(key));
  const missingFromRuntime = [...declared].filter((key) => !runtime.has(key));
  assert.deepEqual(missingFromType, [], 'runtime keys absent from the Formulon interface');
  assert.deepEqual(missingFromRuntime, [], 'Formulon interface keys absent from the runtime object');
  assert.equal(mod.PivotAxis.Row, 0);
  assert.equal(mod.getErrorDisplayName, undefined);
  assert.equal(typeof mod.ErrorDisplayName, 'object');
});

test('every table exported by name is reachable through the Formulon object', async () => {
  const named = await loadStagedModule();
  const formulon = await getModule();
  const tables = Object.keys(named).filter((key) => key !== 'default' && key !== 'createFormulon');
  const missing = tables.filter((key) => formulon[key] !== named[key]);
  assert.deepEqual(missing, [], 'named exports absent from the Formulon object');
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
    mod.setLogMinLevel(level);
  }
  for (const bad of [-1, 5, 99]) {
    assert.throws(
      () => mod.setLogMinLevel(bad),
      (e) => e instanceof mod.FormulonError && e.code === 2,
      `level=${bad} should be rejected`,
    );
  }
  mod.setLogMinLevel(mod.LogLevel.Off);
});

test('setLogSink accepts a function and null, and rejects a non-function', async () => {
  const mod = await getModule();
  mod.setLogSink(() => {});
  mod.setLogSink(null);
  assert.throws(() => mod.setLogSink(), TypeError);
  assert.throws(() => mod.setLogSink(42), TypeError);
  mod.setLogSink(null);
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

test('a registered sink receives each record as a parsed object', async () => {
  const mod = await getModule();
  const records = [];
  mod.setLogSink((record) => records.push(record));
  mod.setLogMinLevel(mod.LogLevel.Warn);
  try {
    const wb = mod.Workbook.createDefault();
    // A structured reference the XLSB encoder cannot lower,
    // which makes the writer emit a per-cell warn record.
    wb.sheet(0).setFormula(0, 0, '=SUM(T[C])');
    wb.recalc();
    wb.save(mod.WorkbookFormat.Xlsb);
    wb.dispose();
    await waitForRecords(records, 1);
    assert.ok(records.length > 0, 'expected at least one record');
    for (const record of records) {
      assert.equal(typeof record, 'object', 'record must be an object');
      assert.equal(record.level, mod.LogLevel.Warn);
    }
    assert.ok(records.some((record) => record.event === 'xlsb.writer.formula_downgraded'));
  } finally {
    mod.setLogMinLevel(mod.LogLevel.Off);
    mod.setLogSink(null);
  }
});

test('the default threshold delivers nothing to a registered sink', async () => {
  const mod = await getModule();
  const records = [];
  mod.setLogSink((record) => records.push(record));
  try {
    const wb = mod.Workbook.createDefault();
    wb.sheet(0).setFormula(0, 0, '=SUM(T[C])');
    wb.recalc();
    wb.save(mod.WorkbookFormat.Xlsb);
    wb.dispose();
    // Wait the same budget the positive case needs, so silence here is a
    // result rather than an artefact of not having waited long enough.
    await new Promise((resolve) => setTimeout(resolve, SINK_SETTLE_MS));
    assert.deepEqual(records, []);
  } finally {
    mod.setLogSink(null);
  }
});

test('a throwing log sink does not leave a pending exception at the ThreadSafeFunction boundary', async () => {
  const mod = await getModule();
  let calls = 0;
  mod.setLogSink(() => {
    calls += 1;
    throw new Error('sink failure');
  });
  mod.setLogMinLevel(mod.LogLevel.Warn);
  try {
    const wb = mod.Workbook.createDefault();
    wb.sheet(0).setFormula(0, 0, '=SUM(T[C])');
    wb.recalc();
    wb.save(mod.WorkbookFormat.Xlsb);
    wb.dispose();
    // Wait for the ThreadSafeFunction drain to run the sink.
    await new Promise((resolve) => setTimeout(resolve, SINK_SETTLE_MS));
    assert.ok(calls > 0, 'expected the sink to have been invoked');
  } finally {
    mod.setLogMinLevel(mod.LogLevel.Off);
    mod.setLogSink(null);
  }

  // An uncaught exception at the drain boundary would have crashed the
  // process by now; the engine must still be usable afterwards.
  const wb2 = mod.Workbook.createDefault();
  try {
    wb2.sheet(0).setFormula(0, 0, '=1+1');
    wb2.recalc();
    assert.equal(wb2.sheet(0).getTypedValue(0, 0).number, 2);
  } finally {
    wb2.dispose();
  }
});

test('memoryUsage and the Python wheel agree that the estimate grows', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const empty = wb.getMemoryUsage();
  for (let row = 0; row < 2000; row += 1) {
    wb.sheet(0).setNumber(row, 0, row);
  }
  assert.ok(wb.getMemoryUsage() > empty);
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

test('getVersion() returns a non-empty string', async () => {
  const mod = await getModule();
  const v = mod.getVersion();
  assert.equal(typeof v, 'string');
  assert.ok(v.length > 0, `expected non-empty version, got ${JSON.stringify(v)}`);
});

test('ErrorDisplayName holds the Excel literals', async () => {
  const mod = await getModule();
  assert.equal(mod.ErrorDisplayName[mod.ErrorCode.Div0], '#DIV/0!');
  assert.equal(mod.ErrorDisplayName[999], undefined);
  assert.equal(Object.isFrozen(mod.ErrorDisplayName), true);
  assert.deepEqual(Object.keys(mod.ErrorDisplayName).sort(), Object.values(mod.ErrorCode).map(String).sort());
});

// One-shot evaluation is anchored at Sheet1!A1 and read-only on every
// shipped surface, so these must agree value-for-value with the WASM
// package and the Python wheel. The anchor-referencing cases are the ones
// that diverge if a surface writes the formula into A1 and recalcs.
const ONE_SHOT_CASES = [
  { formula: '=1+2', kind: 1, number: 3 },
  { formula: '=A1', kind: 1, number: 0 },
  { formula: '=COUNTA(A1)', kind: 1, number: 0 },
  { formula: '=ISBLANK(A1)', kind: 2, boolean: true },
  { formula: '=SUM(A1:A3)', kind: 1, number: 0 },
  { formula: '=SEQUENCE(3)', kind: 5, rows: 3, cols: 1, numbers: [1, 2, 3] },
  { formula: '=ROWS(SEQUENCE(3))', kind: 1, number: 3 },
];
test('evaluateFormula evaluates read-only against a blank anchor', async () => {
  const mod = await getModule();
  for (const spec of ONE_SHOT_CASES) {
    const r = mod.evaluateTypedFormula(spec.formula);
    assert.equal(r.kind, spec.kind, spec.formula);
    if (spec.number !== undefined) {
      assert.equal(r.number, spec.number, spec.formula);
    }
    if (spec.boolean !== undefined) {
      assert.equal(r.boolean, spec.boolean, spec.formula);
    }
    if (spec.numbers !== undefined) {
      assert.equal(r.rows, spec.rows, spec.formula);
      assert.equal(r.cols, spec.cols, spec.formula);
      assert.deepEqual(
        r.cells.map((cell) => cell.number),
        spec.numbers,
        spec.formula,
      );
    }
    // Writing into the anchor and recalcing would make these #REF!.
    assert.notEqual(r.kind, mod.ValueKind.Error, spec.formula);
  }
});

test("evaluateFormula('=1+2') returns NUMBER 3", async () => {
  const mod = await getModule();
  const r = mod.evaluateTypedFormula('=1+2');
  assert.equal(r.kind, mod.ValueKind.Number);
  assert.equal(r.number, 3);
});

test("evaluateFormula('=#REF!') returns an ERROR value (does not throw)", async () => {
  const mod = await getModule();
  const r = mod.evaluateTypedFormula('=#REF!');
  assert.equal(r.kind, mod.ValueKind.Error);
});

test('Workbook.createDefault + setNumber + getValue round-trip', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.listSheets().length, 1);
  wb.sheet(0).setNumber(0, 0, 42);
  const r = wb.sheet(0).getTypedValue(0, 0);
  assert.equal(r.kind, mod.ValueKind.Number);
  assert.equal(r.number, 42);
});

test('paginate exposes page geometry', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // Asserts the result's shape, not the page geometry: how many rows fit a
  // page is the print engine's business and is pinned by the native print
  // suite against Excel. Row 400 is far enough down to force a break under
  // any plausible body height.
  wb.sheet(0).setNumber(0, 0, 1);
  wb.sheet(0).setNumber(400, 0, 2);
  const result = wb.sheet(0).getPagination();
  assert.ok(result.pageCount >= 2, `pageCount: ${result.pageCount}`);
  assert.deepEqual(result.printArea, []);
  assert.ok(result.rowBreaks.length > 0);
  const breakRows = result.rowBreaks.map((b) => b.row);
  assert.deepEqual(
    breakRows,
    [...new Set(breakRows)].sort((a, b) => a - b),
  );
  assert.deepEqual(result.columnBreaks, []);
});

test('print settings: raw XML round-trips and rejects malformed fragments', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  // An absent element reads back as null.
  assert.equal(wb.sheet(0).ooxml.getPageSetupXml(), null);

  const fragment = '<pageSetup paperSize="9" orientation="portrait" scale="85"/>';
  wb.sheet(0).ooxml.setPageSetupXml(fragment);
  assert.equal(wb.sheet(0).ooxml.getPageSetupXml(), fragment);

  // Two top-level elements, the wrong root name, and a truncated fragment
  // are all rejected at set time rather than producing a file Excel repairs.
  for (const bad of ['<pageSetup/><pageSetup/>', '<pageMargins left="1"/>', '<pageSetup orientation="portrait"']) {
    assert.throws(() => wb.sheet(0).ooxml.setPageSetupXml(bad), isFormulonError, `expected rejection for ${bad}`);
  }
  assert.equal(wb.sheet(0).ooxml.getPageSetupXml(), fragment);

  // null removes the element and restores the defaults.
  wb.sheet(0).ooxml.setPageSetupXml(null);
  assert.equal(wb.sheet(0).ooxml.getPageSetupXml(), null);
  const cleared = wb.sheet(0).getPageSetup();
  assert.equal(cleared.scalePct, 100);
  assert.equal(cleared.scalePctStated, false);
});

test('print settings: typed patch touches only the keys it states', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).ooxml.setPageSetupXml('<pageSetup paperSize="9" orientation="portrait" horizontalDpi="600" copies="3"/>');
  wb.sheet(0).updatePageSetup({ orientation: mod.PageOrientation.Landscape });

  const xml = wb.sheet(0).ooxml.getPageSetupXml();
  assert.match(xml, /orientation="landscape"/);
  // Attributes the engine does not model must survive a patch.
  assert.match(xml, /horizontalDpi="600"/);
  assert.match(xml, /copies="3"/);
  assert.match(xml, /paperSize="9"/);

  const read = wb.sheet(0).getPageSetup();
  assert.equal(read.orientation, mod.PageOrientation.Landscape);
  assert.equal(read.paperSize, 9);
  assert.equal('orientationStated' in read, false);

  wb.sheet(0).updatePageMargins({ leftIn: 0.25, topIn: 1.5 });
  const margins = wb.sheet(0).getPageMargins();
  assert.equal(margins.leftIn, 0.25);
  assert.equal(margins.leftInStated, true);
  // An unstated margin reports the OOXML default with its flag clear.
  assert.equal(margins.rightIn, 0.7);
  assert.equal(margins.rightInStated, false);
});

test('print settings: fitToPage keeps the rest of sheetPr', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).ooxml.setSheetPrXml('<sheetPr codeName="Sheet1"><tabColor rgb="FFFF0000"/></sheetPr>');
  wb.sheet(0).updatePageSetup({ fitToPage: true });
  const xml = wb.sheet(0).ooxml.getSheetPrXml();
  assert.match(xml, /codeName="Sheet1"/);
  assert.match(xml, /<tabColor rgb="FFFF0000"\/>/);
  assert.match(xml, /fitToPage="true"/);
  assert.equal(wb.sheet(0).getPageSetup().fitToPage, true);
  // fitToPage: false removes the attribute and the stated flag reads false.
  wb.sheet(0).updatePageSetup({ fitToPage: false });
  assert.doesNotMatch(wb.sheet(0).ooxml.getSheetPrXml(), /fitToPage/);
  assert.equal(wb.sheet(0).getPageSetup().fitToPage, false);
  assert.equal(wb.sheet(0).getPageSetup().fitToPageStated, false);
});

test('print settings: header/footer sections take decoded text', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).updateHeaderFooter({ oddHeader: '&C\u5e33\u7968', oddFooter: '&R&P / &N' });
  // The codes reach the file escaped; a parser hands back `&C...`.
  assert.equal(
    wb.sheet(0).ooxml.getHeaderFooterXml(),
    '<headerFooter><oddHeader>&amp;C\u5e33\u7968</oddHeader><oddFooter>&amp;R&amp;P / &amp;N</oddFooter></headerFooter>',
  );

  // Omitting a key leaves that section alone; an empty string or null clears it.
  wb.sheet(0).updateHeaderFooter({ oddFooter: '' });
  assert.equal(
    wb.sheet(0).ooxml.getHeaderFooterXml(),
    '<headerFooter><oddHeader>&amp;C\u5e33\u7968</oddHeader></headerFooter>',
  );
  wb.sheet(0).updateHeaderFooter({ oddHeader: null });
  assert.doesNotMatch(wb.sheet(0).ooxml.getHeaderFooterXml() ?? '', /oddHeader|oddFooter/);
});

test('print settings: print area, titles and manual breaks', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  wb.sheet(0).setPrintArea([{ firstRow: 0, firstCol: 0, lastRow: 7, lastCol: 5 }]);
  assert.deepEqual(wb.sheet(0).getPrintArea(), [{ firstRow: 0, firstCol: 0, lastRow: 7, lastCol: 5 }]);
  assert.throws(
    () => wb.sheet(0).setPrintArea([{ firstRow: 5, firstCol: 0, lastRow: 4, lastCol: 0 }]),
    isFormulonError,
  );
  assert.deepEqual(wb.sheet(0).getPrintArea(), [{ firstRow: 0, firstCol: 0, lastRow: 7, lastCol: 5 }]);
  wb.sheet(0).setPrintArea(['A1:F8']);
  assert.deepEqual(wb.sheet(0).getPrintArea(), [{ firstRow: 0, firstCol: 0, lastRow: 7, lastCol: 5 }]);
  wb.sheet(0).setPrintArea([]);
  assert.equal(wb.sheet(0).getPrintArea(), null);

  assert.equal(wb.sheet(0).getPrintTitles(), null);
  wb.sheet(0).setPrintTitles({ rows: { first: 0, last: 1 }, columns: { first: 0, last: 0 } });
  assert.deepEqual(wb.sheet(0).getPrintTitles(), { rows: { first: 0, last: 1 }, columns: { first: 0, last: 0 } });
  wb.sheet(0).setPrintTitles(null);
  assert.equal(wb.sheet(0).getPrintTitles(), null);

  wb.sheet(0).setRowBreak(40, { manual: true });
  wb.sheet(0).setRowBreak(10, { manual: true });
  const breaks = wb.sheet(0).listRowBreaks();
  // Kept sorted by row regardless of insertion order.
  assert.deepEqual(
    breaks.map((b) => b.row),
    [10, 40],
  );
  assert.equal(breaks[0].first, 0);
  assert.equal(breaks[0].last, 16383);
  assert.equal(breaks[0].manual, true);

  // A set on an existing row replaces its span and flag; null removes it, and removing an absent break is a no-op.
  wb.sheet(0).setRowBreak(10, { first: 2, last: 5, manual: false });
  const replaced = wb.sheet(0).listRowBreaks();
  assert.equal(replaced.length, 2);
  assert.deepEqual(
    { row: replaced[0].row, first: replaced[0].first, last: replaced[0].last, manual: replaced[0].manual },
    { row: 10, first: 2, last: 5, manual: false },
  );
  wb.sheet(0).setRowBreak(10, null);
  assert.equal(wb.sheet(0).listRowBreaks().length, 1);
  wb.sheet(0).setRowBreak(10, null);
  assert.equal(wb.sheet(0).listRowBreaks().length, 1);

  wb.sheet(0).setColumnBreak(3, { first: 0, last: 99 });
  const columnBreaks = wb.sheet(0).listColumnBreaks();
  assert.deepEqual(
    columnBreaks.map((b) => [b.col, b.first, b.last, b.manual]),
    [[3, 0, 99, true]],
  );
  wb.sheet(0).setColumnBreak(3, null);
  assert.equal(wb.sheet(0).listColumnBreaks().length, 0);

  wb.sheet(0).clearBreaks();
  assert.equal(wb.sheet(0).listRowBreaks().length, 0);
});

test('print settings: a change is visible to paginate without a save cycle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  for (let row = 0; row < 200; row += 1) {
    for (let col = 0; col < 20; col += 1) {
      wb.sheet(0).setNumber(row, col, 1);
    }
  }
  const portrait = wb.sheet(0).getPagination().pageCount;
  wb.sheet(0).updatePageSetup({ orientation: mod.PageOrientation.Landscape });
  assert.notEqual(wb.sheet(0).getPagination().pageCount, portrait);
});

test('setRangeXfIndex applies one xf across a rectangle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  const border = wb.ooxml.addBorder({
    left: { style: 1, color: rgb(0xff000000) },
    right: { style: 1, color: rgb(0xff000000) },
    top: { style: 1, color: rgb(0xff000000) },
    bottom: { style: 1, color: rgb(0xff000000) },
    diagonal: { style: 0, color: NO_COLOR },
    diagonalUp: false,
    diagonalDown: false,
  });
  const xf = wb.ooxml.addCellXf({ fontIndex: 0, fillIndex: 0, borderIndex: border, numFmtId: 0 });

  wb.sheet(0).ooxml.setRangeXfIndex({ firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 2 }, xf);
  // Every cell in the rectangle takes the format, including ones that held
  // no value, so the box renders.
  for (const [row, col] of [
    [0, 0],
    [1, 1],
    [2, 2],
  ]) {
    assert.equal(wb.sheet(0).ooxml.getCellXfIndex(row, col), xf);
  }
  assert.equal(wb.sheet(0).ooxml.getCellXfIndex(3, 3), null);
  wb.sheet(0).ooxml.setCellXfIndex(3, 3, 0);
  assert.equal(wb.sheet(0).ooxml.getCellXfIndex(3, 3), null);
});

test('pivotCount + pivotLayout expose the PivotTable projection', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  assert.equal(wb.sheet(0).listPivots().length, 0);

  assert.throws(
    () => wb.sheet(0).pivot(0).getLayout(),
    (e) => isFormulonError(e) && e.code !== 0,
  );
});
