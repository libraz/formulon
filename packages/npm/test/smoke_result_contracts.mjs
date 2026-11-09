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
function failureProbes(wb, Module) {
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
    [true, () => wb.sheet(99).pivot(0).getLayout()],
    [true, () => wb.pivotCache(9999).getWorksheetSource()],
    [true, () => wb.sheet(99).pivot(0).getReportLayout()],
    [true, () => wb.sheet(0).pivot(99).listFilters()],
    [true, () => wb.sheet(99).pivot(0).getGrandTotals()],
    [true, () => wb.sheet(99).pivot(0).getOptions()],
    [true, () => wb.pivotCache(9999).getOptions()],
    [true, () => wb.pivotCache(9999).listCalculatedItems()],
    [true, () => wb.sheet(99).pivot(0).isProjectionPending()],
    [true, () => wb.pivotCache(9999).listFields()],
    [true, () => wb.sheet(99).pivot(0).getStyleOptions()],
    [true, () => wb.sheet(99).pivot(0).getRowFieldOrder()],
    [true, () => wb.sheet(0).pivot(99).listFields()],
    [true, () => wb.sheet(99).pivot(0).getFieldItem(0, 0)],
    [true, () => wb.sheet(0).pivot(99).listDataFields()],
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
    [false, () => wb.ooxml.getNumberFormat(59999)],
    [true, () => wb.sheet(99).getLambdaText(0, 0)],
    [true, () => wb.sheet(99).getCellPhoneticRuns(0, 0)],
    [true, () => wb.sheet(99).listCfRules()],
    [true, () => wb.sheet(99).listValidations()],
    [true, () => wb.sheet(99).listProtectedRanges()],
    [true, () => wb.sheet(99).listSparklineGroups()],
    [false, () => wb.listExternalLinks()],
    [true, () => wb.sheet(99).getCellPhoneticProperties(0, 0)],
    [true, () => wb.sheet(99).getCellShowPhonetic(0, 0)],
    [true, () => wb.sheet(99).getSelection()],
    // A released handle is the only failure path, which is not usable here.
    [false, () => wb.getTableStyleDefaults()],
    [true, () => wb.ooxml.addCellXf({ fontIndex: 9999 })],
    [false, () => wb.ooxml.addNumberFormat('0.00')],
    [false, () => wb.getIterative()],
    [true, () => wb.sheet(99).recalcViewport({ firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 })],
    [true, () => wb.sheet(99).ooxml.getAutoFilterXml()],
    [true, () => wb.sheet(99).ooxml.getPageSetupXml()],
    [true, () => wb.sheet(99).getPrintArea()],
    [true, () => wb.sheet(99).getPrintTitles()],
    [true, () => wb.sheet(99).listRowBreaks()],
    [true, () => wb.sheet(99).getPageSetup()],
    [true, () => wb.sheet(99).getPageMargins()],
    [true, () => wb.sheet(99).listCells()],
    // A released handle is the only failure path, which is not usable here.
    [false, () => wb.getPinnedNow()],
    [true, () => Module.getLocaleFacts('nope')],
    [true, () => Module.compareText('a', 'b', 0, 'nope')],
    [true, () => Module.getCollationClasses([{ kind: 1, number: 1 }], 0, 'nope')],
    [true, () => wb.sheet(99).getSpill(0, 0)],
    [true, () => wb.sheet(99).getArrayFormula(0, 0)],
    [true, () => wb.sheet(99).getFormatDefaults()],
    [true, () => wb.sheet(99).getCellRectPt({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 }, 0)],
    [true, () => wb.sheet(99).getWidthModel(0)],
    [true, () => wb.sheet(99).getFormula(0, 0)],
    [true, () => wb.sheet(99).getCellsInRange({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 })],
    [true, () => wb.sheet(99).getAutoFilter()],
    [true, () => wb.sheet(99).getDataTable(0, 0)],
    [true, () => wb.sheet(99).evaluateAutoFilter()],
    [true, () => wb.sheet(99).validateValue(0, 0, { kind: 1, number: 1 })],
    [true, () => wb.sheet(99).listThreadedComments()],
    // A released handle is the only failure path, which is not usable here.
    [false, () => wb.listPeople()],
    [true, () => Module.getImageInfo(new Uint8Array(4))],
    [true, () => wb.sheet(99).picture(1).getProperties()],
    [true, () => wb.sheet(99).addPicture(new Uint8Array(4))],
    [true, () => wb.sheet(99).chart(1).snapshot()],
    [true, () => wb.sheet(99).chart(1).getProperties()],
    [true, () => wb.sheet(99).chart(1).getData()],
    [true, () => wb.sheet(99).addChart({ type: 0, range: 'A1:B2' })],
    [true, () => wb.sheet(99).chart(1).addSeries(0, { values: 'A1:A2' })],
    [true, () => wb.sheet(99).ooxml.getRangeXfBases({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 })],
    [true, () => wb.sheet(99).ooxml.snapshotRangeXfs({ firstRow: 0, firstCol: 0, lastRow: 0, lastCol: 0 })],
    [true, () => wb.sheet(99).getDisplayText(0, 0)],
    // A released handle is the only failure path, which is not usable here.
    [false, () => wb.getThemeFonts()],
    [false, () => wb.getThemeColors()],
    [true, () => wb.resolveColor({ kind: 2, theme: 0, tint: Number.NaN }, 0)],
    [true, () => wb.sheet(99).ooxml.getEffectiveStyle(0, 0)],
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
    for (const [fails, call] of failureProbes(wb, Module)) {
      if (fails) {
        expectFailure(call);
      } else {
        const r = call();
        assert.equal(r !== null && typeof r === 'object' && 'status' in r, false, String(call));
      }
    }
  } finally {
    wb.dispose();
  }
});

test('iterative caps accept 1..32767 and refuse the rest by class', async () => {
  const Module = await getModule();
  const wb = Module.Workbook.createDefault();
  try {
    wb.updateIterative({ enabled: true, maxIterations: 32767, maxChange: 0.1 });
    const accepted = wb.getIterative();
    assert.throws(
      () => wb.updateIterative({ enabled: false, maxIterations: 0 }),
      (e) => e.name === 'FormulonError' && e.code === Module.StatusCode.InvalidArgument,
      'a zero cap is rejected, not clamped',
    );
    assert.deepEqual({ ...wb.getIterative() }, { ...accepted }, 'a rejected patch changes no key');
    for (const value of [32768, 2 ** 32 - 1, -1]) {
      assert.throws(
        () => wb.updateIterative({ maxIterations: value }),
        (e) => e.name === 'FormulonError' && e.code === Module.StatusCode.InvalidArgument,
        `maxIterations=${String(value)}`,
      );
    }
    for (const value of [2 ** 32, 1.5, Number.NaN, Number.POSITIVE_INFINITY]) {
      assert.throws(() => wb.updateIterative({ maxIterations: value }), RangeError, `maxIterations=${String(value)}`);
    }
    for (const value of ['5', null]) {
      assert.throws(() => wb.updateIterative({ maxIterations: value }), TypeError, `maxIterations=${String(value)}`);
    }
    assert.equal(wb.getIterative().maxIterations, accepted.maxIterations);
  } finally {
    wb.dispose();
  }
});
