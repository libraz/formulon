// Sheet.kind and Sheet.getChart on a worksheet and on an Excel-saved chartsheet, from both containers.

import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath } from 'node:url';

import { getModule, isFormulonError } from './smoke_support.mjs';

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..', '..', '..');
const fixtureDir = path.join(repoRoot, 'tests', 'fixtures', 'excel', 'chartsheet_twins');

// `Sheet1` (a worksheet) then `Chart1` (a chartsheet holding one chart).
function fixture(ext) {
  return new Uint8Array(readFileSync(path.join(fixtureDir, `tab_color_rgb_red.${ext}`)));
}

test('SheetKind is the frozen numeric enum', async () => {
  const f = await getModule();
  assert.deepEqual({ ...f.SheetKind }, { Worksheet: 0, Chartsheet: 1, Dialogsheet: 2, Macrosheet: 3, Other: 4 });
  assert.ok(Object.isFrozen(f.SheetKind));
});

for (const ext of ['xlsx', 'xlsb']) {
  test(`kind and getChart tell a worksheet from a chartsheet (${ext})`, async () => {
    const f = await getModule();
    const wb = f.Workbook.load(fixture(ext));
    try {
      const [worksheet, chartsheet] = wb.listSheets();
      assert.equal(worksheet.kind, f.SheetKind.Worksheet);
      assert.equal(chartsheet.kind, f.SheetKind.Chartsheet);

      assert.equal(worksheet.getChart(), null);
      const chart = chartsheet.getChart();
      assert.notEqual(chart, null);
      assert.equal(chart.getProperties().onChartsheet, true);
      // The handle `chart(objectId)` returns for that chart.
      assert.equal(chartsheet.chart(chart.objectId), chart);
      assert.equal(chartsheet.getChart(), chart);
    } finally {
      wb.dispose();
    }
  });
}

for (const ext of ['xlsx', 'xlsb']) {
  test(`zoomToFit reads a chartsheet's stored fit and a new zoom ends it (${ext})`, async () => {
    const f = await getModule();
    const wb = f.Workbook.load(fixture(ext));
    try {
      const [worksheet, chartsheet] = wb.listSheets();
      assert.equal(worksheet.getView().zoomToFit, false);
      const view = chartsheet.getView();
      assert.equal(view.zoomToFit, true);
      // The read view goes back unchanged: zoomToFit is read-only and the rest restates the view.
      chartsheet.updateView(view);
      assert.deepEqual(chartsheet.getView(), view);
      chartsheet.updateView({ zoomPct: 80 });
      assert.equal(chartsheet.getView().zoomPct, 80);
      assert.equal(chartsheet.getView().zoomToFit, false);
    } finally {
      wb.dispose();
    }
  });
}

test('a new workbook holds worksheets and kind is read-only', async () => {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    const added = wb.addSheet('Second');
    assert.equal(wb.sheet(0).kind, f.SheetKind.Worksheet);
    assert.equal(added.kind, f.SheetKind.Worksheet);
    assert.equal(added.getChart(), null);
    assert.throws(() => {
      added.kind = f.SheetKind.Chartsheet;
    }, TypeError);
    wb.removeSheet(added);
    assert.throws(
      () => added.kind,
      (e) => isFormulonError(e, f.StatusCode.NotFound),
    );
    assert.throws(
      () => added.getChart(),
      (e) => isFormulonError(e, f.StatusCode.NotFound),
    );
  } finally {
    wb.dispose();
  }
});

for (const ext of ['xlsx', 'xlsb']) {
  test(`a chartsheet answers from its part and refuses worksheet content (${ext})`, async () => {
    const f = await getModule();
    const wb = f.Workbook.load(fixture(ext));
    const invalid = (e) => isFormulonError(e, f.StatusCode.InvalidArgument);
    try {
      const chartsheet = wb.sheet(1);
      // Cell reads and writes need a worksheet.
      assert.throws(() => chartsheet.getValue('A1'), invalid);
      assert.throws(() => chartsheet.setValue('A1', 1), invalid);
      // A list of worksheet content is empty.
      assert.deepEqual(chartsheet.listMerges(), []);
      // The tab colour is read from the part; a new one is refused.
      assert.notEqual(chartsheet.getTabColor(), null);
      assert.throws(() => chartsheet.setTabColor(null), invalid);
      // Visibility serves every kind.
      assert.equal(chartsheet.getVisibility(), f.SheetVisibility.Visible);
      // The view takes a new zoom and tab selection, and refuses any other change.
      assert.throws(() => chartsheet.updateView({ showGridLines: !chartsheet.getView().showGridLines }), invalid);
      chartsheet.updateView({ tabSelected: true });
      assert.equal(chartsheet.getView().tabSelected, true);
    } finally {
      wb.dispose();
    }
  });
}

function protectedFixture(name, ext) {
  return new Uint8Array(readFileSync(path.join(fixtureDir, `${name}.${ext}`)));
}

const hashedProtection = {
  algorithmName: 'SHA-512',
  hashValue: 'FEOoOgNU+lu/gQaYkcJsMsJCAdolOUZsju7crdLGfmo6yGCo62jHPQ2LfHdvg+umnjLNWsGZoZgG16ju/+taFQ==',
  saltValue: 'BAMwdck9CSCcwdaFLEANaA==',
  spinCount: 100000,
  legacyPasswordHash: null,
  allowEditContent: false,
  allowEditObjects: false,
};

for (const ext of ['xlsx', 'xlsb']) {
  test(`getChartsheetProtection reads what a chartsheet's part states (${ext})`, async () => {
    const f = await getModule();
    const contentsOnly = f.Workbook.load(protectedFixture('prot_contents_only', ext));
    const both = f.Workbook.load(protectedFixture('prot_both_pw_abc', ext));
    const plain = f.Workbook.load(fixture(ext));
    try {
      assert.deepEqual(contentsOnly.sheet(1).getChartsheetProtection(), {
        algorithmName: null,
        hashValue: null,
        saltValue: null,
        spinCount: 0,
        legacyPasswordHash: null,
        allowEditContent: false,
        allowEditObjects: true,
      });
      assert.deepEqual(both.sheet(1).getChartsheetProtection(), hashedProtection);
      // A worksheet, and a chartsheet that states none, read null.
      assert.equal(contentsOnly.sheet(0).getChartsheetProtection(), null);
      assert.equal(plain.sheet(1).getChartsheetProtection(), null);
      // A worksheet's protection is read by getSheetProtection, which a chartsheet refuses.
      assert.throws(
        () => contentsOnly.sheet(1).getSheetProtection(),
        (e) => isFormulonError(e, f.StatusCode.InvalidArgument),
      );
      assert.throws(
        () => contentsOnly.sheet(1).setSheetProtection(null),
        (e) => isFormulonError(e, f.StatusCode.InvalidArgument),
      );
      // The record is not a SheetProtectionInput.
      assert.throws(() => plain.sheet(0).setSheetProtection(both.sheet(1).getChartsheetProtection()), TypeError);
    } finally {
      contentsOnly.dispose();
      both.dispose();
      plain.dispose();
    }
  });
}

test('getChartsheetProtection reads a worksheet as null and refuses a removed sheet', async () => {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    const added = wb.addSheet('Second');
    assert.equal(wb.sheet(0).getChartsheetProtection(), null);
    added.setSheetProtection({});
    assert.equal(added.getChartsheetProtection(), null);
    wb.removeSheet(added);
    assert.throws(
      () => added.getChartsheetProtection(),
      (e) => isFormulonError(e, f.StatusCode.NotFound),
    );
  } finally {
    wb.dispose();
  }
});

const authoredDir = path.join(repoRoot, 'tests', 'fixtures', 'chartsheet_protection');
const unlocked = { algorithmName: null, hashValue: null, saltValue: null, spinCount: 0, legacyPasswordHash: null };

test('getChartsheetProtection reads objects-only, legacy-only and bare parts, and the bare one reads null after an xlsb save', async () => {
  const f = await getModule();
  const cases = [
    ['objects_only', { ...unlocked, allowEditContent: true, allowEditObjects: false }, true],
    ['legacy_only', { ...unlocked, legacyPasswordHash: 'CBEB', allowEditContent: false, allowEditObjects: true }, true],
    ['bare', { ...unlocked, allowEditContent: true, allowEditObjects: true }, false],
  ];
  for (const [name, expected, survivesXlsb] of cases) {
    const wb = f.Workbook.load(new Uint8Array(readFileSync(path.join(authoredDir, `${name}.xlsx`))));
    try {
      assert.deepEqual(wb.sheet(1).getChartsheetProtection(), expected, name);
      const reloaded = f.Workbook.load(wb.save(f.WorkbookFormat.Xlsb));
      try {
        assert.deepEqual(reloaded.sheet(1).getChartsheetProtection(), survivesXlsb ? expected : null, name);
      } finally {
        reloaded.dispose();
      }
    } finally {
      wb.dispose();
    }
  }
});
