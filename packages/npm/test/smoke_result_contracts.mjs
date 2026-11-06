import test from 'node:test';
import { assert, expectFailure, getModule, path, pkgRoot, readFile } from './smoke_support.mjs';

// ---- Failure paths and payload shapes ----------------------------------
//
// Every failing call throws a FormulonError, and no declared type carries a
// `status` member. The probe table below has one call per payload shape.

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

// Every declared interface with a `status` member.
function typesWithStatus(dts) {
  const out = [];
  for (const m of dts.matchAll(/^export interface ([A-Za-z0-9_]+)/gm)) {
    if (interfaceMembers(dts, m[1]).some(([key]) => key === 'status')) out.push(m[1]);
  }
  return out;
}

// One failing (or, where no failure is reachable, succeeding) call per
// payload shape. `fails: false` marks a call whose only failure path is a
// destroyed handle, which `delete()` makes unusable to probe.
function failureProbes(wb) {
  return [
    [true, () => wb.recalcParallel(9)],
    [true, () => wb.getValue(99, 0, 0)],
    [true, () => wb.getPivotCacheFieldSharedItemValue(9999, 0, 0)],
    [true, () => wb.evaluateFormula(99, 0, 0, '=1')],
    [true, () => wb.evaluateFormulaArray(99, 0, 0, '=1')],
    [true, () => wb.getShape(99, 1)],
    [true, () => wb.save(99)],
    [true, () => wb.saveWithDiagnostics(99)],
    [false, () => wb.getDiagnostics()],
    [true, () => wb.getSheetName(99)],
    [true, () => wb.getPivotLayout(99, 0)],
    [true, () => wb.getPivotCacheWorksheetSource(9999)],
    [true, () => wb.getPivotReportLayout(99, 0)],
    [true, () => wb.getPivotFilter(99, 0, 0)],
    [true, () => wb.getPivotGrandTotals(99, 0)],
    [true, () => wb.getPivotOptions(99, 0)],
    [true, () => wb.getPivotCacheOptions(9999)],
    [true, () => wb.getPivotCacheCalculatedItem(9999, 0)],
    [true, () => wb.isPivotProjectionPending(99, 0)],
    [true, () => wb.isPivotCacheFieldSharedItemDate(9999, 0, 0)],
    [true, () => wb.getPivotStyleInfo(99, 0)],
    [true, () => wb.getPivotRowFieldOrder(99, 0)],
    [true, () => wb.getPivotField(99, 0, 0)],
    [true, () => wb.getPivotFieldItem(99, 0, 0, 0)],
    [true, () => wb.getPivotDataField(99, 0, 0)],
    [true, () => wb.evaluateCfRange(99, { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }, Number.NaN)],
    [true, () => wb.getPagination(99)],
    [true, () => wb.getView(99)],
    [true, () => wb.getSheetProtection(99)],
    // A released handle is the only failure path of the workbook-level getter.
    [false, () => wb.getWorkbookProtection()],
    [true, () => wb.listColumnOverrides(99)],
    [true, () => wb.listRowOverrides(99)],
    [true, () => wb.getComment(99, 0, 0)],
    [true, () => wb.getCellXfIndex(99, 0, 0)],
    [true, () => wb.getCellXf(9999)],
    [true, () => wb.getFont(9999)],
    [true, () => wb.getFill(9999)],
    [true, () => wb.getBorder(9999)],
    [false, () => wb.getNumFmtString(59999)],
    [true, () => wb.getLambdaText(99, 0, 0)],
    [true, () => wb.getCellPhoneticRun(99, 0, 0, 0)],
    [true, () => wb.getCfRule(99, 0)],
    [true, () => wb.getValidation(99, 0)],
    [true, () => wb.getProtectedRange(99, 0)],
    [true, () => wb.getSparklineGroup(99, 0)],
    [true, () => wb.getExternalLink(0)],
    [true, () => wb.getCellPhoneticProperties(99, 0, 0)],
    [true, () => wb.getCellShowPhonetic(99, 0, 0)],
    [true, () => wb.getSelection(99)],
    // A released handle is the only failure path, which is not usable here.
    [false, () => wb.getTableStyleDefaults()],
    [true, () => wb.addCellXf({ fontIndex: 9999 })],
    [false, () => wb.addNumFmt('0.00')],
    [false, () => wb.getIterative()],
    [true, () => wb.recalcViewport(99, { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 })],
    [true, () => wb.getAutoFilterXml(99)],
    [true, () => wb.getPageSetupXml(99)],
    [true, () => wb.getPrintArea(99)],
    [true, () => wb.getPrintTitles(99)],
    [true, () => wb.listRowBreaks(99)],
    [true, () => wb.getPageSetup(99)],
    [true, () => wb.getPageMargins(99)],
    [true, () => wb.getCellCount(99)],
    // A released handle is the only failure path, which is not usable here.
    [false, () => wb.getPinnedNow()],
    [true, () => wb.getLocaleFacts('nope')],
    [true, () => wb.compareText('a', 'b', 0, 'nope')],
    [true, () => wb.getCollationClasses([{ kind: 1, number: 1 }], 0, 'nope')],
    [true, () => wb.getSpillInfo(99, 0, 0)],
    [true, () => wb.getArrayFormula(99, 0, 0)],
    [true, () => wb.getFormatDefaults(99)],
    [true, () => wb.getCellRectPt(99, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, 0)],
    [true, () => wb.getWidthModel(99, 0)],
    [true, () => wb.getFormula(99, 0, 0)],
    [true, () => wb.getCellsInRange(99, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 })],
    [true, () => wb.getAutoFilter(99)],
    [true, () => wb.getDataTable(99, 0, 0)],
    [true, () => wb.evaluateAutoFilter(99)],
    [true, () => wb.validateValue(99, 0, 0, { kind: 1, number: 1 })],
    [true, () => wb.listThreadedComments(99)],
    // A released handle is the only failure path, which is not usable here.
    [false, () => wb.listPeople()],
    [true, () => wb.getImageInfo(new Uint8Array(4))],
    [true, () => wb.getImage(99, 1)],
    [true, () => wb.addImage(99, new Uint8Array(4))],
    [true, () => wb.snapshotDrawing(99, 1)],
    [true, () => wb.getChart(99, 1)],
    [true, () => wb.getChartData(99, 1)],
    [true, () => wb.addChart(99, { kind: 0, source: 'A1:B2' })],
    [true, () => wb.addChartSeries(99, 1, 0, { values: 'A1:A2' })],
    [true, () => wb.getRangeStyleBases(99, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 })],
    [true, () => wb.snapshotRangeStyle(99, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 })],
    [true, () => wb.getDisplayText(99, 0, 0)],
    // A released handle is the only failure path, which is not usable here.
    [false, () => wb.getThemeFonts()],
    [true, () => wb.resolveColor({ kind: 2, theme: 0, tint: Number.NaN }, 0)],
    [true, () => wb.getEffectiveStyle(99, 0, 0)],
  ];
}

test('no declared type carries a status member', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
  assert.deepEqual(typesWithStatus(dts), []);
});

test('failing calls throw a FormulonError and succeeding ones return a payload without status', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    for (const [fails, call] of failureProbes(wb)) {
      if (fails) {
        expectFailure(call);
      } else {
        const r = call();
        assert.equal(r !== null && typeof r === 'object' && 'status' in r, false, String(call));
      }
    }
  } finally {
    wb.delete();
  }
});

test('iterative caps use the signed int32 contract', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    wb.setIterative(true, 2 ** 31 - 1, 0.1);
    wb.setIterative(false, -1, 0.1);
    assert.equal(wb.getIterative().maxIterations, 1, 'negative caps retain the native clamp');
    for (const value of [2 ** 31, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      assert.throws(() => wb.setIterative(true, value, 0.1), RangeError, `maxIterations=${String(value)}`);
    }
  } finally {
    wb.delete();
  }
});
