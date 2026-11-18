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
