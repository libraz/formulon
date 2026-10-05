import test from 'node:test';
import { assert, getModule, path, pkgRoot, readFile } from './smoke_support.mjs';

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
    ['PhoneticPropertiesResult', true, () => wb.getCellPhoneticProperties(99, 0, 0)],
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
    ['NumberResult', true, () => wb.cellCount(99)],
    // A released handle is the only failure path, which is not usable here.
    ['PinnedNowResult', false, () => wb.pinnedNow()],
    ['SpillInfo', true, () => wb.spillInfo(99, 0, 0)],
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
