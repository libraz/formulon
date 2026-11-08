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
function failureProbes(wb, mod) {
  return [
    [true, () => wb.recalcParallel(9)],
    [true, () => wb.sheet(99).getTypedValue(0, 0)],
    [true, () => wb.pivotCache(9999).listFieldSharedItems(0)],
    [true, () => wb.sheet(99).evaluateTypedFormula(0, 0, '=1')],
    [true, () => wb.sheet(99).evaluateTypedFormulaArray(0, 0, '=1')],
    [true, () => wb.sheet(99).shape(1).getProperties()],
    [true, () => wb.save(99)],
    [true, () => wb.saveWithDiagnostics(99)],
    [false, () => wb.getDiagnostics()],
    [true, () => wb.sheet(99).name],
    [true, () => wb.sheet(99).pivot(0).getLayout()],
    [true, () => wb.pivotCache(9999).getWorksheetSource()],
    [true, () => wb.sheet(99).pivot(0).getReportLayout()],
    [true, () => wb.sheet(99).pivot(0).listFilters()],
    [true, () => wb.sheet(99).pivot(0).getGrandTotals()],
    [true, () => wb.sheet(99).pivot(0).getStyleInfo()],
    [true, () => wb.sheet(99).pivot(0).getOptions()],
    [true, () => wb.pivotCache(9999).getOptions()],
    [true, () => wb.pivotCache(9999).listCalculatedItems()],
    [true, () => wb.sheet(99).pivot(0).isProjectionPending()],
    [true, () => wb.pivotCache(9999).listFields()],
    [true, () => wb.sheet(99).pivot(0).getRowFieldOrder()],
    [true, () => wb.sheet(99).pivot(0).listFields()],
    [true, () => wb.sheet(99).pivot(0).getFieldItem(0, 0)],
    [true, () => wb.sheet(99).pivot(0).listDataFields()],
    [true, () => wb.sheet(99).evaluateCfRange({ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }, Number.NaN)],
    [true, () => wb.sheet(99).getPagination()],
    [true, () => wb.sheet(99).getView()],
    [true, () => wb.sheet(99).getSheetProtection()],
    // A released handle is the only failure path of the workbook-level getter.
    [false, () => wb.getWorkbookProtection()],
    [true, () => wb.sheet(99).listColumnOverrides()],
    [true, () => wb.sheet(99).listRowOverrides()],
    [true, () => wb.sheet(99).getComment(0, 0)],
    [true, () => wb.sheet(99).ooxml.getCellXfIndex(0, 0)],
    [true, () => wb.ooxml.getCellXf(9999)],
    [true, () => wb.ooxml.getFont(9999)],
    [true, () => wb.ooxml.getFill(9999)],
    [true, () => wb.ooxml.getBorder(9999)],
    [false, () => wb.ooxml.getNumFmtString(59999)],
    [true, () => wb.sheet(99).getLambdaText(0, 0)],
    [true, () => wb.sheet(99).getCellPhoneticRuns(0, 0)],
    [true, () => wb.sheet(99).listCfRules()],
    [true, () => wb.sheet(99).listValidations()],
    [true, () => wb.sheet(99).listProtectedRanges()],
    [true, () => wb.sheet(99).listSparklineGroups()],
    // A released handle is the only failure path, which is not usable here.
    [false, () => wb.listExternalLinks()],
    [true, () => wb.sheet(99).getCellPhoneticProperties(0, 0)],
    [true, () => wb.sheet(99).getCellShowPhonetic(0, 0)],
    [true, () => wb.sheet(99).getSelection()],
    // A released handle is the only failure path, which is not usable here.
    [false, () => wb.getTableStyleDefaults()],
    [true, () => wb.ooxml.addCellXf({ fontIndex: 9999 })],
    [false, () => wb.ooxml.addNumFmt('0.00')],
    [false, () => wb.getIterative()],
    [true, () => wb.sheet(99).recalcViewport({ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 })],
    [true, () => wb.sheet(99).ooxml.getPageSetupXml()],
    [true, () => wb.sheet(99).getPrintArea()],
    [true, () => wb.sheet(99).getPrintTitles()],
    // `listRowBreaks` / `listColumnBreaks` answer an out-of-range
    // sheet with an empty list rather than an error, unlike every sibling
    // print getter, so no failure path is reachable from here.
    [false, () => wb.sheet(99).listRowBreaks()],
    [true, () => wb.sheet(99).getPageSetup()],
    [true, () => wb.sheet(99).getPageMargins()],
    [true, () => wb.sheet(99).listCells().length],
    // A released handle is the only failure path, which is not usable here.
    [false, () => wb.getPinnedNow()],
    [true, () => mod.getLocaleFacts('nope')],
    [true, () => mod.compareText('a', 'b', 0, 'nope')],
    [true, () => mod.getCollationClasses([{ kind: 1, number: 1 }], 0, 'nope')],
    [true, () => wb.sheet(99).getSpillInfo(0, 0)],
    [true, () => wb.sheet(99).getArrayFormula(0, 0)],
    [true, () => wb.sheet(99).getFormatDefaults()],
    [true, () => wb.sheet(99).getCellRectPt({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, 0)],
    [true, () => wb.sheet(99).getWidthModel(0)],
    [true, () => wb.sheet(99).getFormula(0, 0)],
    [true, () => wb.sheet(99).getCellsInRange({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 })],
    [true, () => wb.sheet(99).getDisplayText(0, 0)],
    // A released handle is the only failure path of the theme getters.
    [false, () => wb.getThemeFonts()],
    [false, () => wb.getThemeColors()],
    [true, () => wb.resolveColor({ kind: 0 }, 99)],
    [true, () => wb.sheet(99).getEffectiveStyle(0, 0)],
    [true, () => wb.sheet(99).getAutoFilter()],
    [true, () => wb.sheet(99).getDataTable(0, 0)],
    [true, () => wb.sheet(99).evaluateAutoFilter()],
    [true, () => wb.sheet(99).listThreadedComments()],
    [false, () => wb.listPeople()],
    [true, () => mod.getImageInfo(new Uint8Array(0))],
    [true, () => wb.sheet(99).picture(1).getProperties()],
    [true, () => wb.sheet(99).addPicture(new Uint8Array(0), {})],
    [true, () => wb.sheet(99).chart(1).snapshot()],
    [true, () => wb.sheet(99).chart(1).getProperties()],
    [true, () => wb.sheet(99).chart(1).getData()],
    [true, () => wb.sheet(99).addChart({ kind: 0, source: 'A1:B2' })],
    [true, () => wb.sheet(99).chart(1).addSeries(0, { values: 'A1:A2' })],
    [true, () => wb.sheet(99).ooxml.getRangeStyleBases({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 })],
    [
      true,
      () =>
        wb.sheet(99).ooxml.getRangeStyleBasesExcluding({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, [], []),
    ],
    [true, () => wb.sheet(99).snapshotRangeStyle({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 })],
    [true, () => wb.sheet(99).validateValue(0, 0, { kind: 1, number: 1 })],
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
    for (const [expectFailure, call] of failureProbes(wb, mod)) {
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
