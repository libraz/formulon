import test from 'node:test';
import { assert, getModule, isFormulonError, path, pkgRoot, readFile } from './smoke_support.mjs';

// ---- Failure paths --------------------------------------------------------
//
// The declaration file declares no result envelope: a failed call throws
// `FormulonError` and every other call returns its payload directly. The
// probe table below is the list of calls that fail for a reachable reason,
// so each one is checked to throw with a string message and context.

// `fails: false` marks a call whose only failure path is a disposed
// handle, which is not usable to probe.
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
    [true, () => wb.getPivotStyleInfo(99, 0)],
    [true, () => wb.getPivotOptions(99, 0)],
    [true, () => wb.getPivotCacheOptions(9999)],
    [true, () => wb.getPivotCacheCalculatedItem(9999, 0)],
    [true, () => wb.isPivotProjectionPending(99, 0)],
    [true, () => wb.isPivotCacheFieldSharedItemDate(9999, 0, 0)],
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
    [true, () => wb.getPageSetupXml(99)],
    [true, () => wb.getPrintArea(99)],
    [true, () => wb.getPrintTitles(99)],
    // `listRowBreaks` / `listColumnBreaks` answer an out-of-range
    // sheet with an empty list rather than an error, unlike every sibling
    // print getter, so no failure path is reachable from here.
    [false, () => wb.listRowBreaks(99)],
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
    [true, () => wb.getDisplayText(99, 0, 0)],
    // A released handle is the only failure path of `getThemeFonts`.
    [false, () => wb.getThemeFonts()],
    [true, () => wb.resolveColor({ kind: 0 }, 99)],
    [true, () => wb.getEffectiveStyle(99, 0, 0)],
    [true, () => wb.getAutoFilter(99)],
    [true, () => wb.getDataTable(99, 0, 0)],
    [true, () => wb.evaluateAutoFilter(99)],
    [true, () => wb.listThreadedComments(99)],
    [false, () => wb.listPeople()],
    [true, () => wb.getImageInfo(new Uint8Array(0))],
    [true, () => wb.getImage(99, 1)],
    [true, () => wb.addImage(99, new Uint8Array(0), {})],
    [true, () => wb.snapshotDrawing(99, 1)],
    [true, () => wb.getChart(99, 1)],
    [true, () => wb.getChartData(99, 1)],
    [true, () => wb.addChart(99, { kind: 0, source: 'A1:B2' })],
    [true, () => wb.addChartSeries(99, 1, 0, { values: 'A1:A2' })],
    [true, () => wb.getRangeStyleBases(99, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 })],
    [true, () => wb.getRangeStyleBasesExcluding(99, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, [], [])],
    [true, () => wb.snapshotRangeStyle(99, { firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 })],
    [true, () => wb.validateValue(99, 0, 0, { kind: 1, number: 1 })],
  ];
}

test('the declaration file declares no result envelope', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'index.d.ts'), 'utf8');
  const envelopes = [...dts.matchAll(/^export interface (\w+) [^{]*\{([\s\S]*?)\n\}/gm)]
    .filter(([, , body]) => /^ {2}(?:readonly )?status\??:/m.test(body))
    .map(([, name]) => name);
  assert.deepEqual(envelopes, [], 'interfaces that still carry a status member');
  assert.doesNotMatch(dts, /^export (?:interface|type) \w*Result\b/m);
});

test('failure paths throw FormulonError carrying a message and a context', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    for (const [expectFailure, call] of failureProbes(wb)) {
      if (!expectFailure) continue;
      assert.throws(
        () => call(),
        (e) => {
          assert.ok(isFormulonError(e), `${call}: expected a FormulonError, got ${e}`);
          assert.equal(typeof e.message, 'string');
          assert.equal(typeof e.context, 'string');
          return true;
        },
        `${call}: probe was expected to fail`,
      );
    }
  } finally {
    wb.dispose();
  }
});
