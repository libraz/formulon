// Acceptance of the renamed, merged and unit-suffixed public surface: method names (old spellings gone),
// update* merging against set*(null) removal of singletons, protection polarity and defaults, validation
// defaults, unit-suffixed record fields, break records, the sheet-owned table lifecycle, reset* methods and the
// renamed enum members. FORMULON_NPM_ENTRY selects the staged subpath: "." (default) or "./threads".

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const entry = process.env.FORMULON_NPM_ENTRY ?? '.';

async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[entry]?.import;
  if (!target) throw new Error(`package.json "exports" has no import target for ${entry}`);
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

const f = await (await loadPackage()).createFormulon();

const PNG_1X1 = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='),
  (c) => c.charCodeAt(0),
);

/** Runs `body(ws, wb)` on a fresh default workbook and disposes it. */
function withSheet(body) {
  const wb = f.Workbook.createDefault();
  try {
    return body(wb.sheet(0), wb);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

function expectCode(fn, code) {
  assert.throws(fn, (e) => {
    assert.equal(e.name, 'FormulonError', `expected a FormulonError, got ${e}`);
    assert.equal(e.code, code, `code ${e.code}`);
    return true;
  });
}

const present = (target, names) => {
  for (const name of names) assert.equal(typeof target[name], 'function', `${name} is a method`);
};
const absent = (target, names) => {
  for (const name of names) assert.equal(typeof target[name], 'undefined', `${name} is gone`);
};
const hasKeys = (record, names) => {
  for (const name of names) assert.ok(name in record, `${name} is a field`);
};
const lacksKeys = (record, names) => {
  for (const name of names) assert.ok(!(name in record), `${name} is not a field`);
};

function fillTable(ws) {
  ws.setText('A1', 'Region');
  ws.setText('B1', 'Amount');
  ws.setText('A2', 'East');
  ws.setNumber('B2', 10);
  ws.setText('A3', 'West');
  ws.setNumber('B3', 30);
}

function addCache(wb) {
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ range: 'A1:B3', sheet: 'Sheet1' });
  cache.addField('Region');
  cache.addField('Amount');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const record = cache.addRecord();
    cache.setRecordValue(record, 0, region);
    cache.setRecordValue(record, 1, amount);
  }
  return cache;
}

function addPivot(ws, wb) {
  fillTable(ws);
  const cache = addCache(wb);
  const pivot = ws.addPivot({ name: 'ByRegion', cacheId: cache.cacheId, anchor: 'E1' });
  pivot.setFieldAxis('Region', f.PivotAxis.Row);
  pivot.setFieldAxis('Amount', f.PivotAxis.Value);
  const amount = 1;
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 });
  return { pivot, cache };
}

function addChart(ws, kind = f.ChartPlotType.Bar) {
  fillTable(ws);
  return ws.addChart({ type: kind, range: 'Sheet1!$A$1:$B$3', legend: f.ChartLegendPosition.Right });
}

describe('method names', () => {
  it('Sheet has update* members for the patch-merging setters and no set* spelling of them', () => {
    withSheet((ws) => {
      present(ws, [
        'updatePageSetup',
        'updatePageMargins',
        'updatePrintOptions',
        'updateHeaderFooter',
        'updateView',
        'updateFormatDefaults',
        'updateCellFormat',
        'updateSheetProtection',
      ]);
      absent(ws, [
        'setPageSetup',
        'setPageMargins',
        'setPrintOptions',
        'setHeaderFooter',
        'setView',
        'setFormatDefaults',
        'setCellFormat',
      ]);
    });
  });

  it('Workbook has update* members for iterative settings and table style defaults', () => {
    withSheet((_ws, wb) => {
      present(wb, ['updateIterative', 'updateTableStyleDefaults', 'resetTheme', 'activeSheet']);
      absent(wb, ['setIterative', 'setTableStyleDefaults', 'clearTheme', 'clearPinnedNow', 'getActiveSheet']);
    });
  });

  it('activeSheet returns the Sheet handle itself', () => {
    withSheet((ws, wb) => {
      assert.equal(wb.activeSheet(), ws);
      assert.equal(typeof wb.activeSheet().setText, 'function');
    });
  });

  it('Pivot and PivotCache carry update* options and spelled-out subtotal function members', () => {
    withSheet((ws, wb) => {
      const { pivot, cache } = addPivot(ws, wb);
      present(pivot, [
        'updateOptions',
        'updateGrandTotals',
        'updateStyleOptions',
        'updateFieldItem',
        'setFieldSubtotalFunctions',
        'addFieldItem',
        'rename',
      ]);
      absent(pivot, [
        'setOptions',
        'setGrandTotals',
        'setStyleInfo',
        'getStyleInfo',
        'updateStyleInfo',
        'updateItem',
        'addFieldSubtotalFn',
        'clearFieldSubtotalFns',
        'addFieldSubtotalFunction',
        'clearFieldSubtotalFunctions',
        'addFieldSharedItem',
      ]);
      present(cache, ['updateOptions', 'addFieldSharedItem']);
      absent(cache, ['setOptions']);
    });
  });

  it('Chart merges with update* and replaces or removes with set* for the legend and points', () => {
    withSheet((ws) => {
      const chart = addChart(ws);
      present(chart, [
        'updateAxis',
        'updateDisplay',
        'updateSeries',
        'updatePlot',
        'updateLegend',
        'updatePoint',
        'setLegend',
        'setPoint',
      ]);
      absent(chart, ['setAxis', 'setDisplay', 'setSeries', 'setPlot']);
    });
  });

  it('Table has clearAutoFilterCriteria, rename and style options members', () => {
    withSheet((ws) => {
      fillTable(ws);
      const table = ws.addTable({ range: 'A1:B3', name: 'Sales' });
      present(table, ['clearAutoFilterCriteria', 'rename', 'getStyleOptions', 'updateStyleOptions', 'setAutoFilter']);
      absent(table, ['clearAutoFilter', 'removeAutoFilter', 'displayName', 'getStyleInfo', 'updateStyleInfo']);
    });
  });

  it('Sheet removal and reset verbs: reset* restore defaults and the remove*/clear* spellings of singletons are gone', () => {
    withSheet((ws) => {
      present(ws, ['resetRowHeight', 'resetColumnWidth', 'clearAutoFilterCriteria', 'setAutoFilter']);
      present(ws.ooxml, [
        'resetCellXfIndex',
        'resetColumnXfIndex',
        'resetRowXfIndex',
        'getColumnXfIndex',
        'setColumnXfIndex',
        'getRowXfIndex',
        'setRowXfIndex',
      ]);
      absent(ws, ['clearRowHeight', 'clearColumnWidth', 'clearAutoFilter', 'removeAutoFilter']);
      absent(ws.ooxml, [
        'clearCellXfIndex',
        'clearColumnStyle',
        'clearRowStyle',
        'getColumnStyle',
        'setColumnStyle',
        'getRowStyle',
        'setRowStyle',
      ]);
    });
  });

  it('Sheet break and hyperlink upserts are set* and the add*/remove* spellings are gone', () => {
    withSheet((ws) => {
      present(ws, [
        'setRowBreak',
        'setColumnBreak',
        'listRowBreaks',
        'listColumnBreaks',
        'setHyperlink',
        'insertHyperlinkAt',
        'removeHyperlinkAt',
      ]);
      absent(ws, [
        'addRowBreak',
        'addColumnBreak',
        'removeRowBreak',
        'removeColumnBreak',
        'addHyperlink',
        'addHyperlinkRange',
        'setHyperlinkRange',
      ]);
    });
  });

  it('range XF members and effective-style reads sit on the ooxml object', () => {
    withSheet((ws) => {
      present(ws.ooxml, [
        'getRangeXfBases',
        'getRangeXfBasesExcluding',
        'getCellXfIndices',
        'remapRangeXfs',
        'snapshotRangeXfs',
        'restoreRangeXfs',
        'getEffectiveStyle',
      ]);
      absent(ws, ['getEffectiveStyle', 'snapshotRangeXfs', 'restoreRangeXfs', 'getRangeStyleBases', 'rangeStyleBases']);
    });
  });

  it('number-format and passthrough members use the spelled-out names on the ooxml object', () => {
    withSheet((_ws, wb) => {
      present(wb.ooxml, ['getNumberFormat', 'addNumberFormat', 'hasPassthrough']);
      absent(wb.ooxml, ['getNumFmtString', 'addNumFmt', 'getPassthrough']);
    });
  });

  it('cell style registry sits on the Workbook beside the table style registry', () => {
    withSheet((_ws, wb) => {
      present(wb, [
        'listCellStyles',
        'getCellStyle',
        'setCellStyle',
        'removeCellStyle',
        'listTableStyles',
        'getTableStyle',
        'setTableStyle',
        'removeTableStyle',
        'getDefaultFont',
        'setDefaultFont',
      ]);
      absent(wb.ooxml, ['listCellStyles', 'getCellStyle', 'setCellStyle', 'removeCellStyle']);
      present(wb.ooxml, ['getCellStyleXf', 'listCellStyleXfs']);
    });
  });
});

describe('update merges and set(null) removes a singleton', () => {
  it('updatePageSetup keeps the keys it does not name', () => {
    withSheet((ws) => {
      ws.updatePageSetup({ scalePct: 150 });
      ws.updatePageSetup({ orientation: f.PageOrientation.Landscape });
      const setup = ws.getPageSetup();
      assert.equal(setup.scalePct, 150);
      assert.equal(setup.orientation, f.PageOrientation.Landscape);
    });
  });

  it('updatePageMargins keeps the margins it does not name', () => {
    withSheet((ws) => {
      ws.updatePageMargins({ leftIn: 1.25 });
      ws.updatePageMargins({ topIn: 0.5 });
      const margins = ws.getPageMargins();
      assert.equal(margins.leftIn, 1.25);
      assert.equal(margins.topIn, 0.5);
    });
  });

  it('updatePrintOptions and updateHeaderFooter merge', () => {
    withSheet((ws) => {
      ws.updatePrintOptions({ gridLines: true });
      ws.updatePrintOptions({ headings: true });
      assert.equal(ws.getPrintOptions().gridLines, true);
      assert.equal(ws.getPrintOptions().headings, true);
      ws.updateHeaderFooter({ oddHeader: '&CTitle' });
      ws.updateHeaderFooter({ oddFooter: '&P' });
      assert.equal(ws.getHeaderFooter().oddHeader, '&CTitle');
      assert.equal(ws.getHeaderFooter().oddFooter, '&P');
    });
  });

  it('updateView merges and a view patch with an unknown key is a TypeError', () => {
    withSheet((ws) => {
      ws.updateView({ zoomPct: 150 });
      ws.updateView({ showGridLines: false });
      assert.equal(ws.getView().zoomPct, 150);
      assert.equal(ws.getView().showGridLines, false);
      assert.throws(() => ws.updateView({ zoomScale: 120 }), TypeError);
    });
  });

  it('updateFormatDefaults merges', () => {
    withSheet((ws) => {
      ws.updateFormatDefaults({ defaultColumnWidthChars: 12 });
      ws.updateFormatDefaults({ defaultRowHeightPt: 20 });
      const defaults = ws.getFormatDefaults();
      assert.equal(defaults.defaultColumnWidthChars, 12);
      assert.equal(defaults.defaultRowHeightPt, 20);
    });
  });

  it('updateCellFormat patches a range and keeps the other format keys', () => {
    withSheet((ws) => {
      ws.updateCellFormat('A1', { font: { bold: true } });
      ws.updateCellFormat('A1', { font: { sizePt: 14 } });
      const font = ws.getCellFormat('A1').font;
      assert.equal(font.bold, true);
      assert.equal(font.sizePt, 14);
    });
  });

  it('updateIterative and updateTableStyleDefaults merge', () => {
    withSheet((_ws, wb) => {
      wb.updateIterative({ enabled: true });
      wb.updateIterative({ maxIterations: 7 });
      assert.equal(wb.getIterative().enabled, true);
      assert.equal(wb.getIterative().maxIterations, 7);
      wb.updateTableStyleDefaults({ tableStyleName: 'TableStyleLight9' });
      wb.updateTableStyleDefaults({ pivotStyleName: 'PivotStyleLight16' });
      assert.equal(wb.getTableStyleDefaults().tableStyleName, 'TableStyleLight9');
      assert.equal(wb.getTableStyleDefaults().pivotStyleName, 'PivotStyleLight16');
    });
  });

  it('Chart.updateLegend merges while Chart.setLegend replaces and setLegend(null) removes', () => {
    withSheet((ws) => {
      const chart = addChart(ws);
      chart.setLegend({ position: f.ChartLegendPosition.Bottom, overlay: true, hiddenEntries: [] });
      chart.updateLegend({ overlay: false });
      let legend = chart.getProperties().legend;
      assert.equal(legend.position, f.ChartLegendPosition.Bottom);
      assert.equal(legend.overlay, false);
      chart.setLegend({ position: f.ChartLegendPosition.Top, overlay: true, hiddenEntries: [] });
      legend = chart.getProperties().legend;
      assert.equal(legend.position, f.ChartLegendPosition.Top);
      chart.setLegend(null);
      assert.equal(chart.getProperties().legend, null);
      chart.setLegend(null);
      assert.equal(chart.getProperties().legend, null);
    });
  });

  it('Chart.updatePoint merges a point override and setPoint(null) removes it', () => {
    withSheet((ws) => {
      const chart = addChart(ws);
      const before = chart.getProperties().plots[0].series[0].points.length;
      chart.updatePoint(0, 0, 1, { line: { color: '#00FF00' } });
      assert.equal(chart.getProperties().plots[0].series[0].points.length, before + 1);
      chart.setPoint(0, 0, 1, null);
      assert.equal(chart.getProperties().plots[0].series[0].points.length, before);
    });
  });

  it('Chart.updateAxis, updateDisplay, updateSeries and updatePlot each merge into the chart', () => {
    withSheet((ws) => {
      const chart = addChart(ws);
      const axis = chart.getProperties().axes.find((a) => a.kind === f.ChartAxisKind.Value);
      chart.updateAxis(axis.id, { min: 0 });
      chart.updateAxis(axis.id, { max: 20 });
      const read = chart.getProperties().axes.find((a) => a.id === axis.id);
      assert.deepEqual([read.min, read.max], [0, 20]);
      chart.updateDisplay({ plotVisibleOnly: false });
      assert.equal(chart.getProperties().plotVisibleOnly, false);
      chart.updatePlot(0, { overlapPct: 10 });
      assert.equal(chart.getProperties().plots[0].overlapPct, 10);
    });
  });

  it('Pivot.updateOptions, updateGrandTotals and updateStyleOptions merge', () => {
    withSheet((ws, wb) => {
      const { pivot, cache } = addPivot(ws, wb);
      pivot.updateGrandTotals({ rows: false });
      pivot.updateGrandTotals({ columns: false });
      assert.equal(pivot.getGrandTotals().rows, false);
      assert.equal(pivot.getGrandTotals().columns, false);
      pivot.updateStyleOptions({ styleName: 'PivotStyleLight16' });
      pivot.updateStyleOptions({ showRowStripes: true });
      assert.equal(pivot.getStyleOptions().styleName, 'PivotStyleLight16');
      assert.equal(pivot.getStyleOptions().showRowStripes, true);
      pivot.updateOptions({ showHeaders: false });
      pivot.updateOptions({ showDrill: false });
      assert.equal(pivot.getOptions().showHeaders, false);
      assert.equal(pivot.getOptions().showDrill, false);
      cache.updateOptions({ refreshOnLoad: true });
      assert.equal(cache.getOptions().refreshOnLoad, true);
    });
  });

  it('setAutoFilter(null) removes the filter and is a no-op when none exists', () => {
    withSheet((ws) => {
      fillTable(ws);
      ws.setAutoFilter(null);
      assert.equal(ws.getAutoFilter(), null);
      ws.setAutoFilter({ range: 'A1:B3' });
      assert.notEqual(ws.getAutoFilter(), null);
      ws.setAutoFilter(null);
      assert.equal(ws.getAutoFilter(), null);
    });
  });

  it('Table.setAutoFilter(null) removes the table filter and is a no-op when none exists', () => {
    withSheet((ws) => {
      fillTable(ws);
      const table = ws.addTable({ range: 'A1:B3', name: 'Sales' });
      table.setAutoFilter(null);
      table.setAutoFilter({ range: 'A1:B3' });
      assert.notEqual(table.getAutoFilter(), null);
      table.setAutoFilter(null);
      assert.equal(table.getAutoFilter(), null);
    });
  });

  it('clearAutoFilterCriteria empties the column criteria and keeps the filter', () => {
    withSheet((ws) => {
      fillTable(ws);
      ws.setAutoFilter({ range: 'A1:B3', columns: [{ columnOffset: 0, kind: f.FilterKind.Values, values: ['East'] }] });
      assert.equal(ws.getAutoFilter().columns.length, 1);
      ws.clearAutoFilterCriteria();
      const filter = ws.getAutoFilter();
      assert.notEqual(filter, null);
      assert.equal(filter.columns.length, 0);
    });
  });

  it('Table.clearAutoFilterCriteria empties the table filter criteria and keeps the filter', () => {
    withSheet((ws) => {
      fillTable(ws);
      const table = ws.addTable({ range: 'A1:B3', name: 'Sales' });
      table.setAutoFilter({
        range: 'A1:B3',
        columns: [{ columnOffset: 0, kind: f.FilterKind.Values, values: ['East'] }],
      });
      table.clearAutoFilterCriteria();
      assert.notEqual(table.getAutoFilter(), null);
      assert.equal(table.getAutoFilter().columns.length, 0);
    });
  });

  it('setPinnedNow(null) unpins the clock', () => {
    withSheet((_ws, wb) => {
      wb.setPinnedNow(new Date(Date.UTC(2024, 0, 15, 12, 0, 0)));
      wb.setPinnedNow(null);
      wb.setPinnedNow(null);
    });
  });

  it('resetTheme restores the default theme', () => {
    withSheet((_ws, wb) => {
      const original = JSON.stringify(wb.getThemeColors());
      const colors = JSON.parse(original);
      assert.ok(Array.isArray(colors.colors ?? colors));
      wb.resetTheme();
      assert.equal(JSON.stringify(wb.getThemeColors()), original);
    });
  });

  it('setHyperlink upserts one link per cell and setHyperlink(cell, null) removes it', () => {
    withSheet((ws) => {
      ws.setHyperlink('A1', { target: 'https://example.com/a', display: 'A' });
      ws.setHyperlink('A1', { target: 'https://example.com/b', display: 'B' });
      assert.equal(ws.listHyperlinks().length, 1);
      assert.equal(ws.getHyperlink('A1').target, 'https://example.com/b');
      ws.setHyperlink('A1', null);
      assert.equal(ws.getHyperlink('A1'), null);
      ws.setHyperlink('A1', null);
      assert.equal(ws.listHyperlinks().length, 0);
    });
  });

  it('setHyperlink with a range links the range from its top-left cell and null removes it', () => {
    withSheet((ws) => {
      ws.setHyperlink('B1', { target: 'https://example.com/r', range: 'B1:B2' });
      assert.deepEqual({ ...ws.getHyperlink('B1').range }, { firstRow: 0, firstCol: 1, lastRow: 1, lastCol: 1 });
      assert.equal(ws.getHyperlink('B2'), null);
      ws.setHyperlink('B1', null);
      assert.equal(ws.listHyperlinks().length, 0);
    });
  });
});

describe('sheet protection', () => {
  it('an empty input takes the dialog defaults: select locked and unlocked cells allowed, the rest not', () => {
    withSheet((ws) => {
      ws.setSheetProtection({});
      const p = ws.getSheetProtection();
      assert.equal(p.allowSelectLockedCells, true);
      assert.equal(p.allowSelectUnlockedCells, true);
      for (const name of [
        'allowEditObjects',
        'allowEditScenarios',
        'allowFormatCells',
        'allowFormatColumns',
        'allowFormatRows',
        'allowInsertColumns',
        'allowInsertRows',
        'allowInsertHyperlinks',
        'allowDeleteColumns',
        'allowDeleteRows',
        'allowSort',
        'allowAutoFilter',
        'allowPivotTables',
      ]) {
        assert.equal(p[name], false, name);
      }
    });
  });

  it('the record carries allow* flags and legacyPasswordHash, with no enabled, sheet or lock-polarity fields', () => {
    withSheet((ws) => {
      ws.setSheetProtection({ legacyPasswordHash: 'CC1A' });
      const p = ws.getSheetProtection();
      assert.equal(p.legacyPasswordHash, 'CC1A');
      lacksKeys(p, [
        'enabled',
        'sheet',
        'legacyPassword',
        'objects',
        'scenarios',
        'formatCells',
        'formatColumns',
        'formatRows',
        'insertColumns',
        'insertRows',
        'insertHyperlinks',
        'deleteColumns',
        'deleteRows',
        'selectLockedCells',
        'selectUnlockedCells',
        'sort',
        'autoFilter',
        'pivotTables',
      ]);
    });
  });

  it('setSheetProtection replaces the whole record and updateSheetProtection merges', () => {
    withSheet((ws) => {
      ws.setSheetProtection({ allowFormatCells: true });
      ws.setSheetProtection({ allowSort: true });
      assert.equal(ws.getSheetProtection().allowFormatCells, false);
      assert.equal(ws.getSheetProtection().allowSort, true);
      ws.updateSheetProtection({ allowFormatCells: true });
      assert.equal(ws.getSheetProtection().allowFormatCells, true);
      assert.equal(ws.getSheetProtection().allowSort, true);
    });
  });

  it('setSheetProtection(null) unprotects and is a no-op when already unprotected', () => {
    withSheet((ws) => {
      assert.equal(ws.getSheetProtection(), null);
      ws.setSheetProtection(null);
      assert.equal(ws.getSheetProtection(), null);
      ws.setSheetProtection({ allowSort: true });
      assert.notEqual(ws.getSheetProtection(), null);
      ws.setSheetProtection(null);
      assert.equal(ws.getSheetProtection(), null);
    });
  });

  it('the removed enabled and sheet keys are rejected with a TypeError', () => {
    withSheet((ws) => {
      assert.throws(() => ws.setSheetProtection({ enabled: true }), TypeError);
      assert.throws(() => ws.setSheetProtection({ formatCells: true }), TypeError);
      assert.throws(() => ws.updateSheetProtection({ sort: true }), TypeError);
    });
  });

  it('workbook protection passwords are named legacy password hashes', () => {
    withSheet((_ws, wb) => {
      wb.setWorkbookProtection({ lockStructure: true });
      const record = wb.getWorkbookProtection();
      hasKeys(record, ['workbookLegacyPasswordHash', 'revisionsLegacyPasswordHash']);
      lacksKeys(record, ['workbookPassword', 'revisionsPassword']);
    });
  });
});

describe('data validation', () => {
  it('allowBlank defaults to true and showInCellDropdown defaults to true', () => {
    withSheet((ws) => {
      ws.addValidation({ ranges: 'B1:B10', type: f.ValidationType.List, formula1: '"Open,Closed"' });
      const dv = ws.listValidations()[0];
      assert.equal(dv.allowBlank, true);
      assert.equal(dv.showInCellDropdown, true);
      lacksKeys(dv, ['showDropDown', 'op']);
    });
  });

  it('allowBlank false and showInCellDropdown false are kept as given', () => {
    withSheet((ws) => {
      ws.addValidation({
        ranges: 'B1:B10',
        type: f.ValidationType.List,
        formula1: '"a,b"',
        allowBlank: false,
        showInCellDropdown: false,
      });
      const dv = ws.listValidations()[0];
      assert.equal(dv.allowBlank, false);
      assert.equal(dv.showInCellDropdown, false);
    });
  });

  it('the comparison is named operator on read and on input', () => {
    withSheet((ws) => {
      ws.addValidation({
        ranges: 'C1:C3',
        type: f.ValidationType.Whole,
        operator: f.ComparisonOperator.GreaterThan,
        formula1: '1',
      });
      const dv = ws.listValidations()[0];
      assert.equal(dv.operator, f.ComparisonOperator.GreaterThan);
    });
  });
});

describe('conditional formats and auto filters', () => {
  it('a CF rule names its differential format dxfIndex and its comparison operator', () => {
    withSheet((ws, wb) => {
      const dxf = wb.ooxml.addDxf({});
      ws.addCfRule({
        type: f.CfRuleType.CellIs,
        ranges: 'A1:A3',
        operator: f.ComparisonOperator.GreaterThanOrEqual,
        formula1: '1',
        dxfIndex: dxf,
      });
      const rule = ws.listCfRules()[0];
      assert.equal(rule.operator, f.ComparisonOperator.GreaterThanOrEqual);
      assert.equal(rule.dxfIndex, dxf);
      lacksKeys(rule, ['op', 'dxfId']);
    });
  });

  it('a CF rule accepts an inline format resolved to a differential format, exclusive with dxfIndex', () => {
    withSheet((ws, wb) => {
      ws.addCfRule({
        type: f.CfRuleType.Expression,
        ranges: 'A1:A3',
        formula1: 'A1>0',
        format: { font: { bold: true } },
      });
      const rule = ws.listCfRules()[0];
      assert.equal(typeof rule.dxfIndex, 'number');
      assert.equal(wb.ooxml.getDxf(rule.dxfIndex).font.bold, true);
      assert.throws(
        () =>
          ws.addCfRule({
            type: f.CfRuleType.Expression,
            ranges: 'B1:B3',
            formula1: 'B1>0',
            format: { font: { bold: true } },
            dxfIndex: rule.dxfIndex,
          }),
        TypeError,
      );
    });
  });

  it('an auto filter column uses spelled-out field names', () => {
    withSheet((ws) => {
      fillTable(ws);
      ws.setAutoFilter({
        range: 'A1:B3',
        columns: [
          {
            columnOffset: 1,
            kind: f.FilterKind.Custom,
            customAnd: true,
            conditions: [
              { operator: f.ComparisonOperator.LessThan, value: '20' },
              { operator: f.ComparisonOperator.NotEqual, value: '5' },
            ],
          },
        ],
      });
      const column = ws.getAutoFilter().columns[0];
      assert.equal(column.columnOffset, 1);
      assert.deepEqual(JSON.parse(JSON.stringify(column.conditions)), [
        { operator: f.ComparisonOperator.LessThan, value: '20' },
        { operator: f.ComparisonOperator.NotEqual, value: '5' },
      ]);
      hasKeys(column, [
        'topValue',
        'filterValue',
        'dynamicValue',
        'dynamicMaxValue',
        'valueIso',
        'maxValueIso',
        'dxfIndex',
        'iconIndex',
      ]);
      lacksKeys(column, [
        'colId',
        'op1',
        'val1',
        'op2',
        'val2',
        'topVal',
        'filterVal',
        'dynVal',
        'dynMaxVal',
        'valIso',
        'maxValIso',
        'dxfId',
        'iconId',
        'customCount',
        'operator1',
        'value1',
        'operator2',
        'value2',
      ]);
    });
  });
});

describe('unit-suffixed fields', () => {
  it('page margins are inches with the In suffix on the record and its Stated flags', () => {
    withSheet((ws) => {
      ws.updatePageMargins({ leftIn: 1, rightIn: 1, topIn: 1, bottomIn: 1, headerIn: 0.4, footerIn: 0.4 });
      const margins = ws.getPageMargins();
      assert.deepEqual(
        [margins.leftIn, margins.rightIn, margins.topIn, margins.bottomIn, margins.headerIn, margins.footerIn],
        [1, 1, 1, 1, 0.4, 0.4],
      );
      assert.equal(margins.leftInStated, true);
      lacksKeys(margins, ['left', 'right', 'top', 'bottom', 'header', 'footer', 'leftStated']);
      assert.throws(() => ws.updatePageMargins({ left: 1 }), TypeError);
    });
  });

  it('page setup scale is scalePct with a scalePctStated flag', () => {
    withSheet((ws) => {
      assert.equal(ws.getPageSetup().scalePctStated, false);
      ws.updatePageSetup({ scalePct: 80 });
      const setup = ws.getPageSetup();
      assert.equal(setup.scalePct, 80);
      assert.equal(setup.scalePctStated, true);
      lacksKeys(setup, ['scale', 'scaleStated']);
      assert.throws(() => ws.updatePageSetup({ scale: 80 }), TypeError);
    });
  });

  it('pagination reports scalePct as a percent and margins in inches', () => {
    withSheet((ws) => {
      ws.setText('A1', 'x');
      assert.equal(ws.getPagination().scalePct, 100);
      ws.updatePageSetup({ scalePct: 50 });
      ws.updatePageMargins({ leftIn: 1.5 });
      const pagination = ws.getPagination();
      assert.equal(pagination.scalePct, 50);
      assert.ok(Math.abs(pagination.margins.leftIn - 1.5) < 1e-9);
      lacksKeys(pagination, ['scale']);
      lacksKeys(pagination.margins, ['left', 'top']);
      hasKeys(pagination.margins, ['leftIn', 'rightIn', 'topIn', 'bottomIn', 'headerIn', 'footerIn']);
    });
  });

  it('pagination printable rectangle carries Pt suffixes', () => {
    withSheet((ws) => {
      ws.setText('A1', 'x');
      const printable = ws.getPagination().printable;
      hasKeys(printable, ['x', 'y', 'widthPt', 'heightPt']);
      lacksKeys(printable, ['width', 'height']);
    });
  });

  it('sheet view zoom is zoomPct and split positions are in twips', () => {
    withSheet((ws) => {
      ws.updateView({ zoomPct: 125 });
      const view = ws.getView();
      assert.equal(view.zoomPct, 125);
      hasKeys(view, ['splitXTwips', 'splitYTwips']);
      lacksKeys(view, ['zoomScale', 'splitX', 'splitY', 'showRowColHeaders']);
    });
  });

  it('column widths are characters and row heights are points on the layout records', () => {
    withSheet((ws) => {
      ws.setColumnWidth('B:C', 14);
      ws.setRowHeight('3', 30);
      const column = ws.listColumnOverrides()[0];
      const row = ws.listRowOverrides()[0];
      assert.equal(column.widthChars, 14);
      assert.equal(row.heightPt, 30);
      lacksKeys(column, ['width']);
      lacksKeys(row, ['height']);
    });
  });

  it('setRowAutoHeight stores points and the row keeps following its content', () => {
    withSheet((ws) => {
      ws.setRowAutoHeight('2', 22);
      const row = ws.listRowOverrides().find((r) => r.row === 1);
      assert.equal(row.heightPt, 22);
      assert.equal(row.customHeight, false);
    });
  });

  it('sheet format defaults use Chars and Pt suffixes on read and input', () => {
    withSheet((ws) => {
      ws.updateFormatDefaults({ defaultColumnWidthChars: 12, baseColumnWidthChars: 9, defaultRowHeightPt: 18 });
      const defaults = ws.getFormatDefaults();
      assert.deepEqual(
        [defaults.defaultColumnWidthChars, defaults.baseColumnWidthChars, defaults.defaultRowHeightPt],
        [12, 9, 18],
      );
      hasKeys(defaults, ['effectiveDefaultColumnWidthChars', 'effectiveDefaultRowHeightPt']);
      lacksKeys(defaults, [
        'defaultColWidth',
        'defaultColWidthChars',
        'baseColWidth',
        'baseColWidthChars',
        'defaultRowHeight',
        'effectiveDefaultColWidth',
        'effectiveDefaultColWidthChars',
        'effectiveDefaultRowHeight',
      ]);
    });
  });

  it('the width model reports the Normal font size in points', () => {
    withSheet((ws) => {
      const model = ws.getWidthModel();
      hasKeys(model, ['normalFontSizePt']);
      lacksKeys(model, ['normalFontSize']);
      assert.ok(model.normalFontSizePt > 0);
    });
  });

  it('font sizes are sizePt on records, patches and the default font', () => {
    withSheet((_ws, wb) => {
      const font = wb.ooxml.getFont(0);
      hasKeys(font, ['sizePt']);
      lacksKeys(font, ['size']);
      assert.equal(wb.getDefaultFont().sizePt, font.sizePt);
    });
  });

  it('chart fill, line and marker carry transparencyPct and sizePt', () => {
    withSheet((ws) => {
      const chart = addChart(ws, f.ChartPlotType.Line);
      chart.updateSeries(0, 0, {
        fill: { color: '#FF0000', transparencyPct: 25 },
        line: { color: '#0000FF', transparencyPct: 10, widthPt: 2 },
        marker: { symbol: f.ChartMarkerSymbol.Circle, sizePt: 7 },
      });
      const series = chart.getProperties().plots[0].series[0];
      assert.equal(series.line.transparencyPct, 10);
      assert.equal(series.marker.sizePt, 7);
      lacksKeys(series.line, ['transparency']);
      lacksKeys(series.marker, ['size']);
    });
  });

  it('chart plot carries gapWidthPct and firstSliceAngleDeg', () => {
    withSheet((ws) => {
      const chart = addChart(ws);
      chart.updatePlot(0, { gapWidthPct: 40 });
      const plot = chart.getProperties().plots[0];
      assert.equal(plot.gapWidthPct, 40);
      hasKeys(plot, ['firstSliceAngleDeg']);
      lacksKeys(plot, ['gapWidth', 'firstSliceAngle']);
      assert.throws(() => chart.updatePlot(0, { gapWidth: 40 }), TypeError);

      const pie = addChart(ws, f.ChartPlotType.Pie);
      pie.updatePlot(0, { firstSliceAngleDeg: 90 });
      assert.equal(pie.getProperties().plots[0].firstSliceAngleDeg, 90);
    });
  });

  it('shape rotation is degrees and line width is points, both converted at the boundary', () => {
    withSheet((ws) => {
      const shape = ws.addShape({
        presetGeometry: 'rect',
        anchor: { kind: f.AnchorKind.OneCell, fromRow: 2, widthEmu: 9144, heightEmu: 9144 },
        rotationDeg: 45,
        line: { widthPt: 2 },
      });
      const props = shape.getProperties();
      assert.equal(props.rotationDeg, 45);
      assert.equal(props.line.widthPt, 2);
      lacksKeys(props, ['rotation', 'lineWidthEmu']);
      shape.update({ rotationDeg: 90 });
      assert.equal(shape.getProperties().rotationDeg, 90);
    });
  });

  it('picture and image info report pixel size as widthPx and heightPx', () => {
    withSheet((ws) => {
      const info = f.getImageInfo(PNG_1X1);
      assert.equal(info.widthPx, 1);
      assert.equal(info.heightPx, 1);
      lacksKeys(info, ['pxWidth', 'pxHeight']);
      const props = ws.addPicture(PNG_1X1).getProperties();
      assert.equal(props.widthPx, 1);
      assert.equal(props.heightPx, 1);
      lacksKeys(props, ['pxWidth', 'pxHeight']);
    });
  });

  it('the image format is a picture property, not a drawing property, and the media path is not exposed', () => {
    withSheet((ws) => {
      const picture = ws.addPicture(PNG_1X1);
      const common = picture.getDrawingProperties();
      hasKeys(common, ['objectId', 'kind', 'name', 'description', 'hidden', 'anchor']);
      lacksKeys(common, ['format', 'imageFormat', 'mediaPath']);
      const props = picture.getProperties();
      hasKeys(props, ['format']);
      lacksKeys(props, ['imageFormat', 'mediaPath']);
    });
  });
});

describe('colours, fills and styles', () => {
  it('a colour is read as argb', () => {
    withSheet((_ws, wb) => {
      const color = wb.ooxml.getFont(0).color;
      hasKeys(color, ['argb']);
      lacksKeys(color, ['rgb']);
    });
  });

  it('a fill names foreground and background', () => {
    withSheet((_ws, wb) => {
      const fill = wb.ooxml.getFill(0);
      hasKeys(fill, ['foreground', 'background']);
      lacksKeys(fill, ['fg', 'bg']);
    });
  });

  it('a cell XF names its alignment horizontal and vertical and its named-style link styleXfIndex', () => {
    withSheet((_ws, wb) => {
      const xf = wb.ooxml.getCellXf(0);
      hasKeys(xf, ['horizontal', 'vertical', 'styleXfIndex']);
      lacksKeys(xf, ['horizontalAlign', 'verticalAlign', 'xfId']);
    });
  });

  it('a differential format nests the number format as numberFormat and types alignment and protection', () => {
    withSheet((_ws, wb) => {
      const index = wb.ooxml.addDxf({ numberFormat: '0.00' });
      const dxf = wb.ooxml.getDxf(index);
      assert.equal(dxf.numberFormat, '0.00');
      lacksKeys(dxf, ['numFmt', 'alignmentXml', 'protectionXml']);
    });
  });

  it('getCellFormat reports where the XF came from and leaves the XF index to ooxml', () => {
    withSheet((ws) => {
      ws.setText('A1', 'x');
      const format = ws.getCellFormat('A1');
      lacksKeys(format, ['xfIndex']);
      assert.equal(typeof ws.ooxml.getCellXfIndex('A1'), 'number');
      assert.ok(Object.values(f.EffectiveStyleSource).includes(format.source));
    });
  });

  it('the effective style names fontColor, numberFormat and protection', () => {
    withSheet((ws) => {
      ws.setText('A1', 'x');
      const style = ws.ooxml.getEffectiveStyle('A1');
      hasKeys(style, ['fontColor', 'numberFormat', 'protection']);
      lacksKeys(style, ['font', 'numFmtCode', 'locked', 'hidden']);
    });
  });

  it('the XF index family spells Xf and the reset restores the default', () => {
    withSheet((ws, wb) => {
      const xf = wb.ooxml.addCellXf({ ...wb.ooxml.getCellXf(0), wrapText: true });
      ws.ooxml.setColumnXfIndex('B', xf);
      ws.ooxml.setRowXfIndex('3', xf);
      assert.equal(ws.ooxml.getColumnXfIndex(1), xf);
      assert.equal(ws.ooxml.getRowXfIndex(2), xf);
      ws.ooxml.resetColumnXfIndex('B');
      ws.ooxml.resetRowXfIndex('3');
      assert.equal(ws.ooxml.getColumnXfIndex(1), null);
      assert.equal(ws.ooxml.getRowXfIndex(2), null);
      ws.ooxml.setCellXfIndex('A1', xf);
      ws.ooxml.resetCellXfIndex('A1');
      assert.notEqual(ws.ooxml.getCellXfIndex('A1'), xf);
    });
  });

  it('phonetic runs use start and end offsets and phonetic properties use fontIndex', () => {
    withSheet((ws) => {
      ws.setText('A1', 'tokyo');
      ws.setCellPhoneticRuns('A1', [{ start: 0, end: 2, text: 'to' }]);
      const run = ws.getCellPhoneticRuns('A1')[0];
      assert.equal(run.start, 0);
      assert.equal(run.end, 2);
      lacksKeys(run, ['sb', 'eb']);
      ws.setCellPhoneticProperties('A1', { fontIndex: 0 });
      const props = ws.getCellPhoneticProperties('A1');
      hasKeys(props, ['fontIndex']);
      lacksKeys(props, ['fontId']);
    });
  });

  it('function metadata names its signature', () => {
    withSheet(() => {
      for (const name of f.listFunctionNames()) {
        const meta = f.getFunctionMetadata(name);
        lacksKeys(meta, ['signatureTemplate']);
        assert.ok(meta.signature === null || typeof meta.signature === 'string', name);
        assert.ok(meta.description === null || typeof meta.description === 'string', name);
      }
    });
  });

  it('cell styles are read and written on the Workbook', () => {
    withSheet((_ws, wb) => {
      const styles = wb.listCellStyles();
      assert.ok(styles.length > 0);
      const normal = wb.getCellStyle(styles[0].name);
      assert.equal(normal.name, styles[0].name);
    });
  });

  it('number formats are added and read back by id', () => {
    withSheet((_ws, wb) => {
      const id = wb.ooxml.addNumberFormat('0.000');
      assert.equal(wb.ooxml.getNumberFormat(id), '0.000');
    });
  });

  it('hasPassthrough answers a boolean for a part path', () => {
    withSheet((_ws, wb) => {
      assert.equal(wb.ooxml.hasPassthrough('customXml/item1.xml'), false);
    });
  });
});

describe('page breaks', () => {
  it('setRowBreak upserts a RowBreak record with row, first, last and manual', () => {
    withSheet((ws) => {
      ws.setRowBreak(5, { first: 2, last: 9, manual: true });
      assert.deepEqual({ ...ws.listRowBreaks()[0] }, { row: 5, first: 2, last: 9, manual: true });
      ws.setRowBreak(5, { first: 0, last: 3, manual: false });
      assert.equal(ws.listRowBreaks().length, 1);
      assert.deepEqual({ ...ws.listRowBreaks()[0] }, { row: 5, first: 0, last: 3, manual: false });
    });
  });

  it('setColumnBreak upserts a ColumnBreak record with col, first, last and manual', () => {
    withSheet((ws) => {
      ws.setColumnBreak(3, { first: 1, last: 20, manual: true });
      assert.deepEqual({ ...ws.listColumnBreaks()[0] }, { col: 3, first: 1, last: 20, manual: true });
    });
  });

  it('a break with no options spans the whole sheet from its first line', () => {
    withSheet((ws) => {
      ws.setRowBreak(4, {});
      ws.setColumnBreak(2, {});
      assert.equal(ws.listRowBreaks()[0].row, 4);
      assert.equal(ws.listRowBreaks()[0].first, 0);
      assert.equal(ws.listColumnBreaks()[0].col, 2);
      assert.equal(ws.listColumnBreaks()[0].first, 0);
    });
  });

  it('setRowBreak(row, null) and setColumnBreak(col, null) remove the break and are no-ops when absent', () => {
    withSheet((ws) => {
      ws.setRowBreak(7, null);
      ws.setColumnBreak(7, null);
      ws.setRowBreak(7, { manual: true });
      ws.setColumnBreak(7, { manual: true });
      ws.setRowBreak(7, null);
      ws.setColumnBreak(7, null);
      assert.equal(ws.listRowBreaks().length, 0);
      assert.equal(ws.listColumnBreaks().length, 0);
    });
  });

  it('break options reject unknown keys with a TypeError', () => {
    withSheet((ws) => {
      present(ws, ['setRowBreak', 'setColumnBreak']);
      assert.throws(() => ws.setRowBreak(1, { min: 0 }), TypeError);
      assert.throws(() => ws.setColumnBreak(1, { id: 0 }), TypeError);
    });
  });

  it('pagination lists breaks as records and drops the parallel manual arrays', () => {
    withSheet((ws) => {
      for (let r = 0; r < 12; r += 1) ws.setText(r, 0, `r${r}`);
      ws.setRowBreak(3, { manual: true });
      const pagination = ws.getPagination();
      const hit = pagination.rowBreaks.find((b) => b.row === 3);
      assert.ok(hit, 'the manual break is listed');
      assert.equal(hit.manual, true);
      hasKeys(hit, ['row', 'first', 'last', 'manual']);
      assert.ok(Array.isArray(pagination.columnBreaks));
      lacksKeys(pagination, ['rowBreakManual', 'columnBreakManual']);
    });
  });

  it('clearBreaks removes the breaks of both axes', () => {
    withSheet((ws) => {
      ws.setRowBreak(2, { manual: true });
      ws.setColumnBreak(2, { manual: true });
      ws.clearBreaks();
      assert.equal(ws.listRowBreaks().length, 0);
      assert.equal(ws.listColumnBreaks().length, 0);
    });
  });
});

describe('table lifecycle on the sheet', () => {
  it('addTable returns a Table, and table, listTables and removeTable work on the owning sheet', () => {
    withSheet((ws) => {
      fillTable(ws);
      const table = ws.addTable({ range: 'A1:B3', name: 'Sales' });
      assert.equal(table.name, 'Sales');
      assert.equal(ws.table('Sales'), table);
      assert.equal(ws.table(0), table);
      assert.deepEqual(
        ws.listTables().map((t) => t.name),
        ['Sales'],
      );
      ws.removeTable('Sales');
      assert.equal(ws.listTables().length, 0);
    });
  });

  it('the Workbook no longer owns the table lifecycle but still finds a table by name', () => {
    withSheet((ws, wb) => {
      absent(wb, ['addTable', 'listTables', 'removeTable']);
      fillTable(ws);
      const table = ws.addTable({ range: 'A1:B3', name: 'Sales' });
      assert.equal(wb.table('Sales'), table);
    });
  });

  it('table indices are local to the sheet', () => {
    withSheet((ws, wb) => {
      fillTable(ws);
      const second = wb.addSheet('Second');
      second.setText('A1', 'H');
      second.setNumber('A2', 1);
      ws.addTable({ range: 'A1:B3', name: 'First' });
      const other = second.addTable({ range: 'A1:A2', name: 'Other' });
      assert.equal(second.table(0), other);
      assert.equal(second.listTables().length, 1);
      assert.equal(ws.listTables().length, 1);
      assert.equal(wb.table('Other'), other);
    });
  });

  it('a table range needs no sheet qualifier because the receiver is the sheet', () => {
    withSheet((ws) => {
      fillTable(ws);
      assert.doesNotThrow(() => ws.addTable({ range: 'A1:B3', name: 'Plain' }));
    });
  });

  it('a name used by another sheet is refused workbook-wide', () => {
    withSheet((ws, wb) => {
      fillTable(ws);
      ws.addTable({ range: 'A1:B3', name: 'Sales' });
      const second = wb.addSheet('Second');
      second.setText('A1', 'H');
      second.setNumber('A2', 1);
      assert.throws(
        () => second.addTable({ range: 'A1:A2', name: 'Sales' }),
        (e) => e.name === 'FormulonError',
      );
    });
  });

  it('a missing table is NotFound', () => {
    withSheet((ws, wb) => {
      expectCode(() => ws.table('Nope'), f.StatusCode.NotFound);
      expectCode(() => ws.removeTable('Nope'), f.StatusCode.NotFound);
      expectCode(() => wb.table('Nope'), f.StatusCode.NotFound);
    });
  });

  it('a table has one name and rename updates the name formulas use', () => {
    withSheet((ws, wb) => {
      fillTable(ws);
      const table = ws.addTable({ range: 'A1:B3', name: 'Sales' });
      table.rename('Rev');
      assert.equal(table.name, 'Rev');
      assert.equal(wb.table('Rev'), table);
      lacksKeys(table.getProperties(), ['displayName']);
      assert.throws(() => table.update({ displayName: 'Y' }), TypeError);
    });
  });

  it('a new table takes TableStyleMedium2 with row stripes when no default is set', () => {
    withSheet((ws) => {
      fillTable(ws);
      const info = ws.addTable({ range: 'A1:B3', name: 'Sales' }).getStyleOptions();
      assert.equal(info.styleName, 'TableStyleMedium2');
      assert.equal(info.showRowStripes, true);
    });
  });

  it('a new table takes the workbook default table style', () => {
    withSheet((ws, wb) => {
      wb.updateTableStyleDefaults({ tableStyleName: 'TableStyleLight9' });
      fillTable(ws);
      assert.equal(ws.addTable({ range: 'A1:B3', name: 'Sales' }).getStyleOptions().styleName, 'TableStyleLight9');
    });
  });

  it('table style options are a TableStyleOptions updated by patch and absent from the properties record', () => {
    withSheet((ws) => {
      fillTable(ws);
      const table = ws.addTable({ range: 'A1:B3', name: 'Sales' });
      table.updateStyleOptions({ showColumnStripes: true });
      table.updateStyleOptions({ showFirstColumn: true });
      const info = table.getStyleOptions();
      assert.equal(info.showColumnStripes, true);
      assert.equal(info.showFirstColumn, true);
      assert.equal(info.showRowStripes, true);
      hasKeys(info, ['styleName', 'showFirstColumn', 'showLastColumn', 'showRowStripes', 'showColumnStripes']);
      lacksKeys(table.getProperties(), [
        'styleName',
        'showFirstColumn',
        'showLastColumn',
        'showRowStripes',
        'showColumnStripes',
      ]);
      assert.throws(() => table.update({ styleName: 'TableStyleLight1' }), TypeError);
    });
  });

  it('a default pivot style applies to a new pivot', () => {
    withSheet((ws, wb) => {
      wb.updateTableStyleDefaults({ pivotStyleName: 'PivotStyleLight16' });
      const { pivot } = addPivot(ws, wb);
      assert.equal(pivot.getStyleOptions().styleName, 'PivotStyleLight16');
    });
  });
});

describe('reset methods', () => {
  it('resetRowHeight restores the default height of a row', () => {
    withSheet((ws) => {
      ws.setRowHeight('3', 40);
      assert.equal(
        ws.listRowOverrides().some((r) => r.row === 2 && r.heightPt === 40),
        true,
      );
      ws.resetRowHeight('3');
      assert.equal(
        ws.listRowOverrides().some((r) => r.row === 2 && r.heightPt !== null),
        false,
      );
    });
  });

  it('resetColumnWidth restores the default width of a column', () => {
    withSheet((ws) => {
      ws.setColumnWidth('B', 30);
      assert.equal(
        ws.listColumnOverrides().some((c) => c.widthChars === 30),
        true,
      );
      ws.resetColumnWidth('B');
      assert.equal(
        ws.listColumnOverrides().some((c) => c.widthChars === 30),
        false,
      );
    });
  });

  it('resetRowHeight and resetColumnWidth are no-ops on untouched spans', () => {
    withSheet((ws) => {
      ws.resetRowHeight('9');
      ws.resetColumnWidth('Z');
      assert.equal(ws.listRowOverrides().length, 0);
      assert.equal(ws.listColumnOverrides().length, 0);
    });
  });

  it('resetTheme is the only theme restore verb', () => {
    withSheet((_ws, wb) => {
      wb.resetTheme();
      absent(wb, ['clearTheme']);
    });
  });
});

describe('enum members', () => {
  it('PivotFilterKind spells the top-count and date-between members', () => {
    assert.equal(f.PivotFilterKind.ValueTopCount, 0);
    assert.equal(f.PivotFilterKind.DateBetween, 5);
    assert.equal(f.PivotFilterKind.ValueTop10, undefined);
    assert.equal(f.PivotFilterKind.LabelDate, undefined);
  });

  it('ChartPlotType.Extended replaces Chartex', () => {
    assert.equal(f.ChartPlotType.Extended, 16);
    assert.equal(f.ChartPlotType.Chartex, undefined);
  });

  it('ChartExtendedLayout names the eight extended chart layouts', () => {
    assert.deepEqual(
      { ...f.ChartExtendedLayout },
      {
        BoxWhisker: 0,
        ClusteredColumn: 1,
        Funnel: 2,
        ParetoLine: 3,
        RegionMap: 4,
        Sunburst: 5,
        Treemap: 6,
        Waterfall: 7,
      },
    );
  });

  it('ChartRole carries the ordinals of the series inputs, then the title and axis-title roles', () => {
    assert.deepEqual(
      { ...f.ChartRole },
      { Name: 0, Categories: 1, Values: 2, X: 3, Y: 4, Size: 5, Title: 6, AxisTitle: 7 },
    );
  });

  it('chart properties name the extended flag isExtended and the plot an extendedLayout', () => {
    withSheet((ws) => {
      const chart = addChart(ws);
      const props = chart.getProperties();
      hasKeys(props, ['isExtended']);
      lacksKeys(props, ['isChartex']);
      assert.equal(props.isExtended, false);
      hasKeys(props.plots[0], ['extendedLayout']);
      lacksKeys(props.plots[0], ['chartexLayout']);
      assert.equal(props.plots[0].extendedLayout, null);
    });
  });

  it('chart data reports unresolved inputs as a list whose role is a ChartRole ordinal', () => {
    withSheet((ws) => {
      const chart = addChart(ws);
      const data = chart.getData();
      assert.ok(Array.isArray(data.unresolved));
      for (const entry of data.unresolved) {
        assert.ok(Object.values(f.ChartRole).includes(entry.role));
      }
    });
  });
});

describe('chart axis visibility', () => {
  it('an axis reads visible and the read record round-trips into updateAxis', () => {
    withSheet((ws) => {
      const chart = addChart(ws);
      const axis = chart.getProperties().axes.find((a) => a.kind === f.ChartAxisKind.Value);
      assert.equal(axis.visible, true);
      lacksKeys(axis, ['deleted']);
      chart.updateAxis(axis.id, { visible: false });
      assert.equal(chart.getProperties().axes.find((a) => a.id === axis.id).visible, false);
      chart.updateAxis(axis.id, { visible: true });
      assert.equal(chart.getProperties().axes.find((a) => a.id === axis.id).visible, true);
      assert.throws(() => chart.updateAxis(axis.id, { deleted: true }), TypeError);
    });
  });
});

describe('pivot shape changes', () => {
  it('addFieldItem takes a label or a cache index and returns the item index', () => {
    withSheet((ws, wb) => {
      const { pivot } = addPivot(ws, wb);
      pivot.clearFieldItems(0);
      const byLabel = pivot.addFieldItem(0, 'East');
      const byCache = pivot.addFieldItem(0, { cacheIndex: 1 }, { visible: false });
      assert.equal(typeof byLabel, 'number');
      assert.notEqual(byCache, byLabel);
      assert.equal(pivot.listFieldItems(0).length, 2);
      expectCode(() => pivot.addFieldItem(0, 'East'), f.StatusCode.AlreadyExists);
    });
  });

  it('setAnchor takes one cell and refuses a multi-cell range', () => {
    withSheet((ws, wb) => {
      const { pivot } = addPivot(ws, wb);
      pivot.setAnchor('H1');
      expectCode(() => pivot.setAnchor('H1:I2'), f.StatusCode.InvalidArgument);
    });
  });

  it('updateFieldItem patches an item by position', () => {
    withSheet((ws, wb) => {
      const { pivot } = addPivot(ws, wb);
      pivot.updateFieldItem(0, 0, { visible: false });
      assert.equal(pivot.getFieldItem(0, 0).visible, false);
    });
  });

  it('subtotal functions are spelled out on members and on the field record', () => {
    withSheet((ws, wb) => {
      const { pivot } = addPivot(ws, wb);
      pivot.setFieldSubtotalFunctions(0, [f.PivotAggregation.Sum]);
      assert.deepEqual([...pivot.listFields()[0].subtotalFunctions], [f.PivotAggregation.Sum]);
      pivot.setFieldSubtotalFunctions(0, []);
      assert.deepEqual([...pivot.listFields()[0].subtotalFunctions], []);
      lacksKeys(pivot.listFields()[0], ['subtotalFns']);
    });
  });

  it('pivot style options name the style styleName', () => {
    withSheet((ws, wb) => {
      const { pivot } = addPivot(ws, wb);
      pivot.updateStyleOptions({ styleName: 'PivotStyleMedium9' });
      const info = pivot.getStyleOptions();
      assert.equal(info.styleName, 'PivotStyleMedium9');
      lacksKeys(info, ['name']);
    });
  });

  it('Pivot.rename still renames the table', () => {
    withSheet((ws, wb) => {
      const { pivot } = addPivot(ws, wb);
      pivot.rename('Renamed');
      assert.equal(pivot.getProperties().name, 'Renamed');
    });
  });
});

describe('restore options', () => {
  it('restoreDrawing with assignNewId adds a copy', () => {
    withSheet((ws) => {
      const picture = ws.addPicture(PNG_1X1);
      const bytes = picture.snapshot();
      const before = ws.listDrawings().length;
      ws.restoreDrawing(bytes, { assignNewId: true });
      assert.equal(ws.listDrawings().length, before + 1);
    });
  });
});
