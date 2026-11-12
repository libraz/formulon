// Every key of every patch record applies: each key is set through its public update method to a value that differs
// from the setup state and read back through the public getter. A meta-test ties the table to the API ledger
// (tools/api/names.yaml) and the C structs, so a new patch key cannot ship without a case.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import { describe, it } from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const pkgRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const repoRoot = path.resolve(pkgRoot, '../..');
// The staged entry named by the package's own "main" field.
async function loadPackage() {
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  if (!pkg.main) throw new Error('package.json is missing "main"');
  return import(pathToFileURL(path.resolve(pkgRoot, pkg.main)).href);
}

const f = await (await loadPackage()).createFormulon();
const SURFACES = ['wasm', 'native'];

/** Runs `body(ws, wb)` on a fresh default workbook and disposes it. */
function withSheet(body) {
  const wb = f.Workbook.createDefault();
  try {
    return body(wb.sheet(0), wb);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

const plain = (value) => JSON.parse(JSON.stringify(value) ?? 'null');
const rect = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });
const argb = (c) => (c == null ? null : c.argb >>> 0);
const addr = (a) => (a == null ? null : { row: a.row, col: a.col });

/** Set the opposite of what the getter reads before the patch. */
const FLIP = Symbol('flip');
/** A case: set `v` (or flip the pre-read bool); `clear` also sends `null` and expects `cleared` (default null). */
const k = (v, extra = {}) => ({ v, ...extra });

const RED = { kind: 1, argb: 0xffff0000, theme: 0, tint: 0, indexed: 0 };
const BLUE = { kind: 1, argb: 0xff0000ff, theme: 0, tint: 0, indexed: 0 };
const PERSON = '{00000000-0000-4000-8000-0000000000A1}';
const COMMENT = '{00000000-0000-4000-8000-000000000001}';

/** Sheet1: B1:D1 series names, A2:A5 categories, B2:D5 numbers. */
function fillGrid(ws) {
  ['S1', 'S2', 'S3'].forEach((name, c) => {
    ws.setText(0, c + 1, name);
  });
  ['a', 'b', 'c', 'd'].forEach((category, r) => {
    ws.setText(r + 1, 0, category);
    for (let c = 0; c < 3; c += 1) ws.setNumber(r + 1, c + 1, (r + 1) * (c + 2));
  });
}

function fillTable(ws) {
  ws.setText('A1', 'Region');
  ws.setText('B1', 'Amount');
  ws.setText('A2', 'East');
  ws.setNumber('B2', 10);
  ws.setText('A3', 'West');
  ws.setNumber('B3', 30);
}

function addTable(ws) {
  fillTable(ws);
  return ws.addTable({ range: 'A1:B3', name: 'Tbl', styleName: 'TableStyleMedium2' });
}

/** A pivot over A1:B3: Region on rows (items East, West), a Sum of Amount. */
function addPivot(ws, wb) {
  fillTable(ws);
  const cache = wb.addPivotCache();
  cache.addField('Region');
  cache.addField('Amount');
  cache.addFieldSharedItem(0, 'East');
  cache.addFieldSharedItem(0, 'West');
  for (const [region, amount] of [
    ['East', 10],
    ['West', 30],
  ]) {
    const record = cache.addRecord();
    cache.setRecordValue(record, 0, region);
    cache.setRecordValue(record, 1, amount);
  }
  const pivot = ws.addPivot({ name: 'P', cacheId: cache.cacheId, anchor: 'E1' });
  pivot.setFieldAxis('Region', f.PivotAxis.Row);
  pivot.setFieldAxis('Amount', f.PivotAxis.Value);
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: 1, aggregation: f.PivotAggregation.Sum });
  return { cache, pivot };
}

function addChart(ws, type = f.ChartPlotType.Bar) {
  fillGrid(ws);
  return ws.addChart({ type, range: 'Sheet1!A1:D5', legend: f.ChartLegendPosition.Right });
}

const literal = (...values) => ({ kind: f.ChartSourceKind.Literal, formula: '', literal: values });
const lit = (src) => plain(src?.literal);

// ---- Chart sub-records, reached through a series, a point or a marker --------------------------------------------

const firstSeries = (chart) => chart.getProperties().plots[0].series[0];
const lineChart = (ws) => ({ chart: addChart(ws, f.ChartPlotType.Line) });

/** A chart-paint patch type: `slot(series)` picks the record, `wrap(patch)` nests the patch under the series. */
function paintType(slot, wrap) {
  return {
    make: lineChart,
    update: ({ chart }, patch) => chart.updateSeries(0, 0, wrap(patch)),
    read: ({ chart }) => slot(firstSeries(chart)),
  };
}

const CHART_FILL = {
  ...paintType(
    (s) => s.fill,
    (p) => ({ fill: p }),
  ),
  make: (ws) => ({ chart: addChart(ws) }),
  cases: {
    color: k('#FF0000', { get: (r) => argb(r?.color), want: 0xffff0000, clear: true }),
    transparencyPct: k(40, {
      get: (r) => r?.transparencyPct,
      clear: true,
      patch: (v) => ({ color: '#FF0000', transparencyPct: v }),
    }),
    none: k(FLIP, { get: (r) => r?.none ?? false }),
  },
};

const CHART_LINE = {
  ...paintType(
    (s) => s.line,
    (p) => ({ line: p }),
  ),
  cases: {
    color: k('#00FF00', { get: (r) => argb(r?.color), want: 0xff00ff00, clear: true }),
    transparencyPct: k(25, {
      get: (r) => r?.transparencyPct,
      clear: true,
      patch: (v) => ({ color: '#00FF00', transparencyPct: v }),
    }),
    widthPt: k(3.5, { get: (r) => r?.widthPt, clear: true }),
    dash: k(f.ChartLineDash.Dot, { get: (r) => r?.dash, clear: true }),
    none: k(FLIP, { get: (r) => r?.none ?? false }),
  },
};

const CHART_MARKER = {
  ...paintType(
    (s) => s.marker,
    (p) => ({ marker: p }),
  ),
  cases: {
    symbol: k(f.ChartMarkerSymbol.Square, { get: (r) => r?.symbol }),
    sizePt: k(11, {
      get: (r) => r?.sizePt,
      clear: true,
      cleared: 5,
      patch: (v) => ({ symbol: f.ChartMarkerSymbol.Circle, sizePt: v }),
    }),
  },
};

const CHART_LABELS = {
  make: (ws) => {
    const chart = addChart(ws, f.ChartPlotType.Line);
    chart.updateSeries(0, 0, { labels: { showValue: true } });
    return { chart };
  },
  update: ({ chart }, patch) => chart.updateSeries(0, 0, { labels: patch }),
  read: ({ chart }) => firstSeries(chart).labels,
  cases: {
    showValue: k(FLIP),
    showCategory: k(FLIP),
    showSeries: k(FLIP),
    showPercent: k(FLIP),
    showLegendKey: k(FLIP),
    separator: k('; ', { clear: true }),
    position: k(f.ChartLabelPosition.Above, { clear: true }),
    numberFormat: k('0.00', { clear: true }),
    fontColor: k('#0000FF', { get: (r) => argb(r?.fontColor), want: 0xff0000ff, clear: true }),
    fontSizePt: k(14, { clear: true }),
    sourceLinked: k(FLIP),
    deleted: k(FLIP),
  },
};

const sourceCase = (...values) => k(literal(...values), { get: (r) => lit(r), want: values });

const CHART_SERIES = {
  make: lineChart,
  update: ({ chart }, patch) => chart.updateSeries(0, 0, patch),
  read: ({ chart }) => firstSeries(chart),
  cases: {
    name: k(literal('Fixed'), { get: (r) => lit(r.name), want: ['Fixed'], clear: true, cleared: [] }),
    categories: k(literal('p', 'q', 'r', 's'), {
      get: (r) => lit(r.categories),
      want: ['p', 'q', 'r', 's'],
      clear: true,
      cleared: [],
    }),
    values: { ...sourceCase(9, 8, 7, 6), get: (r) => lit(r.values) },
    x: {
      make: (ws) => ({ chart: addChart(ws, f.ChartPlotType.Scatter) }),
      ...sourceCase(5, 6, 7, 8),
      get: (r) => lit(r.x),
      clear: true,
      cleared: [],
    },
    y: {
      make: (ws) => ({ chart: addChart(ws, f.ChartPlotType.Scatter) }),
      ...sourceCase(4, 3, 2, 1),
      get: (r) => lit(r.y),
    },
    marker: k(
      { symbol: f.ChartMarkerSymbol.Diamond },
      { get: (r) => r.marker?.symbol, want: f.ChartMarkerSymbol.Diamond, clear: true },
    ),
    smooth: k(FLIP, { get: (r) => r.smooth ?? false }),
    labels: k(
      { showValue: true },
      { get: (r) => r.labels?.showValue ?? false, want: true, clear: true, cleared: false },
    ),
  },
};

const CHART_POINT = {
  make: lineChart,
  update: ({ chart }, patch) => chart.updatePoint(0, 0, 1, patch),
  read: ({ chart }) => firstSeries(chart).points.find((p) => p.index === 1),
  cases: {
    marker: k(
      { symbol: f.ChartMarkerSymbol.Star },
      { get: (r) => r?.marker?.symbol, want: f.ChartMarkerSymbol.Star, clear: true },
    ),
  },
};

const CHART_PLOT = {
  make: (ws) => ({ chart: addChart(ws) }),
  update: ({ chart }, patch) => chart.updatePlot(0, patch),
  read: ({ chart }) => chart.getProperties().plots[0],
  cases: {
    gapWidthPct: k(80),
    overlapPct: k(25),
    varyColors: k(FLIP),
    firstSliceAngleDeg: { make: (ws) => ({ chart: addChart(ws, f.ChartPlotType.Pie) }), ...k(90) },
    holeSizePct: { make: (ws) => ({ chart: addChart(ws, f.ChartPlotType.Doughnut) }), ...k(70) },
  },
};

const valueAxis = (chart) => chart.getProperties().axes.find((a) => a.kind === f.ChartAxisKind.Value);
const CHART_AXIS = {
  make: (ws) => ({ chart: addChart(ws) }),
  update: ({ chart }, patch) => chart.updateAxis(valueAxis(chart).id, patch),
  read: ({ chart }) => valueAxis(chart),
  cases: {
    visible: k(FLIP),
    title: k('Amount', { get: (r) => r.title?.text, clear: true }),
    min: k(1, { clear: true }),
    max: k(100, { clear: true }),
    majorUnit: k(5, { clear: true }),
    minorUnit: k(0.5, { clear: true }),
    numberFormat: k('0.00', { clear: true }),
    sourceLinked: k(FLIP),
    majorGridlines: k(FLIP),
    minorGridlines: k(FLIP),
    position: k(f.ChartAxisPosition.Right),
    reversed: k(FLIP),
    logBase: k(10, { clear: true }),
    crosses: k(f.ChartAxisCrosses.Max),
    crossesAt: k(3),
    tickLabelPosition: k(f.ChartTickLabelPosition.Low, { clear: true, cleared: f.ChartTickLabelPosition.NextTo }),
  },
};

const CHART_LEGEND = {
  make: (ws) => ({ chart: addChart(ws) }),
  update: ({ chart }, patch) => chart.updateLegend(patch),
  read: ({ chart }) => chart.getProperties().legend,
  cases: {
    position: k(f.ChartLegendPosition.Bottom),
    overlay: k(FLIP),
    hiddenEntries: k([1], { get: (r) => plain(r.hiddenEntries) }),
  },
};

const CHART_DISPLAY = {
  make: (ws) => ({ chart: addChart(ws) }),
  update: ({ chart }, patch) => chart.updateDisplay(patch),
  read: ({ chart }) => chart.getProperties(),
  cases: {
    blanksAs: k(f.BlanksAs.Span, { clear: true }),
    plotVisibleOnly: k(FLIP, { get: (r) => r.plotVisibleOnly ?? true, clear: true, cleared: true }),
    naAsBlank: k(FLIP, { get: (r) => r.naAsBlank ?? false, clear: true, cleared: false }),
    style: k(26, { clear: true }),
  },
};

// ---- Cell format -------------------------------------------------------------------------------------------------

/** A cell-format sub-patch type: `slot` names the CellFormat section the keys land in. */
function formatType(slot) {
  return {
    make: (ws) => ws,
    update: (ws, patch) => ws.updateCellFormat('B2', { [slot]: patch }),
    read: (ws) => ws.getCellFormat('B2')[slot],
  };
}

const FONT = {
  ...formatType('font'),
  cases: {
    name: k('Courier New'),
    sizePt: k(17, { clear: true, cleared: 11 }),
    bold: k(FLIP, { clear: true }),
    italic: k(FLIP, { clear: true }),
    strike: k(FLIP, { clear: true }),
    outline: k(FLIP, { clear: true }),
    shadow: k(FLIP, { clear: true }),
    condense: k(FLIP, { clear: true }),
    extend: k(FLIP, { clear: true }),
    underline: k(f.UnderlineStyle.Double),
    vertAlign: k(f.FontVerticalAlign.Superscript),
    family: k(f.FontFamily.Modern, { clear: true }),
    charset: k(128, { clear: true }),
    scheme: k(f.FontScheme.Minor),
    color: k('#123456', { get: (r) => argb(r.color), want: 0xff123456 }),
  },
};

const FILL = {
  ...formatType('fill'),
  cases: {
    pattern: k(f.FillPattern.LightGrid),
    foreground: k('#112233', {
      get: (r) => argb(r.foreground),
      want: 0xff112233,
      patch: (v) => ({ pattern: f.FillPattern.Solid, foreground: v }),
    }),
    background: k('#445566', {
      get: (r) => argb(r.background),
      want: 0xff445566,
      patch: (v) => ({ pattern: f.FillPattern.Solid, background: v }),
    }),
    gradient: k(
      {
        angleDeg: 45,
        stops: [
          { position: 0, color: RED },
          { position: 1, color: BLUE },
        ],
      },
      { get: (r) => r.gradient?.angleDeg, want: 45, clear: true },
    ),
  },
};

const BORDER_SIDE = {
  make: (ws) => ws,
  update: (ws, patch) => ws.updateCellFormat('B2', { border: { left: patch } }),
  read: (ws) => ws.getCellFormat('B2').border.left,
  cases: {
    style: k(f.BorderStyle.Thick),
    color: k('#AA0000', {
      get: (r) => argb(r.color),
      want: 0xffaa0000,
      patch: (v) => ({ style: f.BorderStyle.Thin, color: v }),
    }),
  },
};

const sideCase = (side) =>
  k({ style: f.BorderStyle.Medium }, { get: (r) => r[side].style, want: f.BorderStyle.Medium });
const BORDER = {
  ...formatType('border'),
  cases: {
    left: sideCase('left'),
    right: sideCase('right'),
    top: sideCase('top'),
    bottom: sideCase('bottom'),
    diagonal: sideCase('diagonal'),
    diagonalUp: k(FLIP),
    diagonalDown: k(FLIP),
  },
};

const ALIGNMENT = {
  ...formatType('alignment'),
  cases: {
    horizontal: k(f.HorizontalAlign.Center),
    vertical: k(f.VerticalAlign.Top),
    wrapText: k(FLIP),
    shrinkToFit: k(FLIP),
    indent: k(3, { patch: (v) => ({ horizontal: f.HorizontalAlign.Left, indent: v }) }),
    textRotation: k(45),
    readingOrder: k(f.ReadingOrder.RightToLeft),
    justifyLastLine: k(FLIP, { patch: (v) => ({ horizontal: f.HorizontalAlign.Justify, justifyLastLine: v }) }),
    relativeIndent: k(2),
  },
};

const PROTECTION = {
  ...formatType('protection'),
  cases: { locked: k(FLIP), hidden: k(FLIP) },
};

const CELL_FORMAT = {
  make: (ws) => ws,
  update: (ws, patch) => ws.updateCellFormat('B2', patch),
  read: (ws) => ws.getCellFormat('B2'),
  cases: {
    numberFormat: k('0.000'),
    font: k({ bold: true }, { get: (r) => r.font.bold, want: true }),
    fill: k(
      { pattern: f.FillPattern.Solid, foreground: '#FF0000' },
      { get: (r) => r.fill.pattern, want: f.FillPattern.Solid },
    ),
    border: k({ top: { style: f.BorderStyle.Thick } }, { get: (r) => r.border.top.style, want: f.BorderStyle.Thick }),
    alignment: k(
      { horizontal: f.HorizontalAlign.Right },
      { get: (r) => r.alignment.horizontal, want: f.HorizontalAlign.Right },
    ),
    protection: k({ locked: false }, { get: (r) => r.protection.locked, want: false }),
  },
};

// ---- Sheet and workbook ------------------------------------------------------------------------------------------

const ITERATIVE = {
  make: (ws, wb) => wb,
  update: (wb, patch) => wb.updateIterative(patch),
  read: (wb) => wb.getIterative(),
  cases: { enabled: k(FLIP), maxIterations: k(37), maxChange: k(0.25) },
};

const TABLE_STYLE_DEFAULTS = {
  make: (ws, wb) => wb,
  update: (wb, patch) => wb.updateTableStyleDefaults(patch),
  read: (wb) => wb.getTableStyleDefaults(),
  cases: {
    tableStyleName: k('TableStyleLight3', { clear: true }),
    pivotStyleName: k('PivotStyleLight12', { clear: true }),
  },
};

const sheetType = (update, read, cases) => ({
  make: (ws) => ws,
  update: (ws, patch) => ws[update](patch),
  read: (ws) => ws[read](),
  cases,
});

const PAGE_SETUP = sheetType('updatePageSetup', 'getPageSetup', {
  orientation: k(f.PageOrientation.Landscape, { clear: true }),
  paperSize: k(f.PaperSize.A3),
  scalePct: k(80),
  fitToWidth: k(2),
  fitToHeight: k(3),
  fitToPage: k(FLIP),
});

const PAGE_MARGINS = sheetType('updatePageMargins', 'getPageMargins', {
  leftIn: k(0.9),
  rightIn: k(0.8),
  topIn: k(1.1),
  bottomIn: k(1.2),
  headerIn: k(0.4),
  footerIn: k(0.45),
});

const PRINT_OPTIONS = sheetType('updatePrintOptions', 'getPrintOptions', {
  gridLines: k(FLIP),
  headings: k(FLIP),
  horizontalCentered: k(FLIP),
  verticalCentered: k(FLIP),
});

const HEADER_FOOTER = sheetType('updateHeaderFooter', 'getHeaderFooter', {
  oddHeader: k('&CHead', { clear: true }),
  oddFooter: k('&CFoot', { clear: true }),
  evenHeader: k('&CEvenHead', { clear: true, patch: (v) => ({ differentOddEven: true, evenHeader: v }) }),
  evenFooter: k('&CEvenFoot', { clear: true, patch: (v) => ({ differentOddEven: true, evenFooter: v }) }),
  firstHeader: k('&CFirstHead', { clear: true, patch: (v) => ({ differentFirst: true, firstHeader: v }) }),
  firstFooter: k('&CFirstFoot', { clear: true, patch: (v) => ({ differentFirst: true, firstFooter: v }) }),
  differentOddEven: k(FLIP),
  differentFirst: k(FLIP),
  scaleWithDoc: k(FLIP),
  alignWithMargins: k(FLIP),
});

const SHEET_VIEW = sheetType('updateView', 'getView', {
  zoomPct: k(150),
  freezeRows: k(2),
  freezeCols: k(3),
  showGridLines: k(FLIP),
  showRowColumnHeaders: k(FLIP),
  showZeros: k(FLIP),
  rightToLeft: k(FLIP),
  tabSelected: k(FLIP),
  viewMode: k(f.SheetViewMode.PageBreakPreview),
  topLeft: k('C5', { get: (r) => addr(r.topLeft), want: { row: 4, col: 2 } }),
  paneTopLeft: k('D6', {
    get: (r) => addr(r.paneTopLeft),
    want: { row: 5, col: 3 },
    patch: (v) => ({ splitXTwips: 3000, splitYTwips: 2000, paneTopLeft: v }),
  }),
  splitXTwips: k(3000),
  splitYTwips: k(2000),
  showFormulas: k(FLIP),
  showOutlineSymbols: k(FLIP),
});

const SHEET_FORMAT_DEFAULTS = sheetType('updateFormatDefaults', 'getFormatDefaults', {
  defaultColumnWidthChars: k(12.5, { clear: true }),
  defaultRowHeightPt: k(21, { clear: true }),
  baseColumnWidthChars: k(10),
  customHeight: k(FLIP),
  zeroHeight: k(FLIP),
});

const ALLOW_FLAGS = [
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
  'allowSelectLockedCells',
  'allowSelectUnlockedCells',
  'allowSort',
  'allowAutoFilter',
  'allowPivotTables',
];

const SHEET_PROTECTION = {
  make: (ws) => {
    ws.setSheetProtection({});
    return ws;
  },
  update: (ws, patch) => ws.updateSheetProtection(patch),
  read: (ws) => ws.getSheetProtection(),
  cases: {
    algorithmName: k('SHA-512', { clear: true }),
    hashValue: k('aGFzaA==', { clear: true }),
    saltValue: k('c2FsdA==', { clear: true }),
    spinCount: k(100000, { clear: true, cleared: 0 }),
    legacyPasswordHash: k('CC1A', { clear: true }),
    ...Object.fromEntries(ALLOW_FLAGS.map((key) => [key, k(FLIP)])),
  },
};

const PROTECTED_RANGE = {
  make: (ws) => {
    ws.addProtectedRange({ name: 'Totals', ranges: [rect(1, 1, 2, 2)] });
    return ws;
  },
  update: (ws, patch) => ws.updateProtectedRange(0, patch),
  read: (ws) => ws.listProtectedRanges()[0],
  cases: {
    name: k('Renamed'),
    ranges: k('C3:D4', { get: (r) => plain(r.ranges), want: [rect(2, 2, 3, 3)] }),
    legacyPasswordHash: k('CC1A', { clear: true }),
    algorithmName: k('SHA-512', { clear: true }),
    hashValue: k('aGFzaA==', { clear: true }),
    saltValue: k('c2FsdA==', { clear: true }),
    spinCount: k(100000),
    securityDescriptor: k('O:S-1-5-32-544G:S-1-5-32-544', { clear: true }),
  },
};

const VALIDATION = {
  make: (ws) => {
    const id = ws.addValidation({
      ranges: 'A1:A3',
      type: f.ValidationType.Whole,
      operator: f.ComparisonOperator.Between,
      formula1: '1',
      formula2: '10',
    });
    return { ws, id };
  },
  update: ({ ws, id }, patch) => ws.updateValidation(id, patch),
  read: ({ ws, id }) => ws.getValidation(id),
  cases: {
    ranges: k('C1:C3', { get: (r) => plain(r.ranges), want: [rect(0, 2, 2, 2)] }),
    type: k(f.ValidationType.Decimal),
    operator: k(f.ComparisonOperator.NotBetween),
    errorStyle: k(f.ValidationErrorStyle.Warning),
    allowBlank: k(FLIP),
    showInputMessage: k(FLIP),
    showErrorMessage: k(FLIP),
    showInCellDropdown: k(FLIP),
    formula1: k('2', { clear: true, want: '=2' }),
    formula2: k('20', { clear: true, want: '=20' }),
    errorTitle: k('Bad', { clear: true }),
    errorMessage: k('Not allowed', { clear: true }),
    promptTitle: k('Hint', { clear: true }),
    promptMessage: k('Enter 1 to 10', { clear: true }),
  },
};

// ---- Conditional formats -----------------------------------------------------------------------------------------

/** A sheet with one filler rule (so a second dxf exists) and the rule under test on A1:A2. */
const cfMaker = (rule) => (ws) => {
  ws.setNumber('A1', 1);
  ws.setNumber('A2', 5);
  ws.addCfRule({
    type: f.CfRuleType.Expression,
    ranges: 'E1:E2',
    formula1: 'TRUE',
    format: { font: { italic: true } },
  });
  const id = ws.addCfRule({ ranges: 'A1:A2', format: { font: { bold: true } }, ...rule });
  return { ws, id };
};
const withCf = (rule, spec) => ({ make: cfMaker(rule), ...spec });
const MIN_MAX = [{ type: f.CfValueObjectType.Min }, { type: f.CfValueObjectType.Max }];
const PERCENTS = ['33', '67'].map((value) => ({ type: f.CfValueObjectType.Percent, value }));
const asGet = (key) => (r) => r[key] ?? false;

const CF_RULE = {
  make: cfMaker({
    type: f.CfRuleType.CellIs,
    operator: f.ComparisonOperator.GreaterThan,
    formula1: '3',
  }),
  update: ({ ws, id }, patch) => ws.updateCfRule(id, patch),
  read: ({ ws, id }) => ws.getCfRule(id),
  cases: {
    priority: k(5),
    stopIfTrue: k(FLIP),
    dxfIndex: k(0, { clear: true }),
    ranges: k('B1:B2', { get: (r) => plain(r.ranges), want: [rect(0, 1, 1, 1)] }),
    formula1: k('4', { clear: true, want: '=4' }),
    formula2: withCf(
      {
        type: f.CfRuleType.CellIs,
        operator: f.ComparisonOperator.Between,
        formula1: '3',
        formula2: '9',
      },
      k('8', { clear: true, want: '=8' }),
    ),
    operator: k(f.ComparisonOperator.LessThan),
    rank: withCf({ type: f.CfRuleType.Top10, rank: 1 }, k(2)),
    percent: withCf({ type: f.CfRuleType.Top10, rank: 1 }, k(FLIP, { get: asGet('percent') })),
    bottom: withCf({ type: f.CfRuleType.Top10, rank: 1 }, k(FLIP, { get: asGet('bottom') })),
    aboveAverage: withCf({ type: f.CfRuleType.AboveAverage }, k(FLIP, { get: asGet('aboveAverage') })),
    equalAverage: withCf({ type: f.CfRuleType.AboveAverage }, k(FLIP, { get: asGet('equalAverage') })),
    stdDev: withCf({ type: f.CfRuleType.AboveAverage }, k(2)),
    text: withCf({ type: f.CfRuleType.ContainsText, text: 'ab' }, k('cd')),
    timePeriod: withCf({ type: f.CfRuleType.TimePeriod, timePeriod: f.DatePeriod.Today }, k(f.DatePeriod.LastMonth)),
    colorScale: withCf(
      { type: f.CfRuleType.ColorScale, colorScale: { thresholds: MIN_MAX, colors: ['#FF0000', '#00FF00'] } },
      k(
        { thresholds: MIN_MAX, colors: ['#0000FF', '#FFFF00'] },
        { get: (r) => r.colorScale.colors.map(argb), want: [0xff0000ff, 0xffffff00] },
      ),
    ),
    dataBar: withCf(
      { type: f.CfRuleType.DataBar, dataBar: { min: MIN_MAX[0], max: MIN_MAX[1], fill: '#0070C0' } },
      k({ min: MIN_MAX[0], max: MIN_MAX[1], fill: '#C00000' }, { get: (r) => argb(r.dataBar.fill), want: 0xffc00000 }),
    ),
    iconSet: withCf(
      { type: f.CfRuleType.IconSet, iconSet: { name: f.IconSetName.ThreeArrows, thresholds: PERCENTS } },
      k(
        { name: f.IconSetName.ThreeFlags, thresholds: PERCENTS },
        { get: (r) => r.iconSet.name, want: f.IconSetName.ThreeFlags },
      ),
    ),
  },
};

// ---- Tables, pivots ----------------------------------------------------------------------------------------------

const TABLE = {
  make: (ws) => ({ table: addTable(ws) }),
  update: ({ table }, patch) => table.update(patch),
  read: ({ table }) => table.getProperties(),
  cases: {
    range: k('A1:B4', { get: (r) => plain(r.range.rect), want: rect(0, 0, 3, 1) }),
    // Turning the header off deletes its cells and moves the range's top edge down one row.
    headerRow: k(false, {
      get: (r) => ({ headerRow: r.headerRow, firstRow: r.range.rect.firstRow }),
      want: { headerRow: false, firstRow: 1 },
    }),
    totalsRow: k(FLIP),
    columns: k([{}, { totalsRowFunction: f.TableTotalsRowFunction.Sum }], {
      get: (r) => r.columns[1].totalsRowFunction,
      want: f.TableTotalsRowFunction.Sum,
    }),
  },
};

/** Table Tbl over A1:B4 with a totals row: headers Item / Qty, data rows 2..3. Element 1 patches the Qty column. */
const TABLE_COLUMN = {
  make: (ws) => {
    ws.setText('A1', 'Item');
    ws.setText('B1', 'Qty');
    ws.setText('A2', 'r2');
    ws.setNumber('B2', 1);
    ws.setText('A3', 'r3');
    ws.setNumber('B3', 2);
    return { table: ws.addTable({ range: 'A1:B4', name: 'Tbl', totalsRow: true }) };
  },
  update: ({ table }, patch) => table.update({ columns: [{}, patch] }),
  read: ({ table }) => table.getProperties().columns[1],
  cases: {
    // The setup's totals row is blank, so its last column already defaults to Sum; patch a different function.
    totalsRowFunction: k(f.TableTotalsRowFunction.Max, { clear: true, cleared: f.TableTotalsRowFunction.None }),
    totalsRowLabel: k('Total', { clear: true }),
    totalsRowFormula: k('=5', {
      patch: (v) => ({ totalsRowFunction: f.TableTotalsRowFunction.Custom, totalsRowFormula: v }),
      clear: true,
      clearPatch: { totalsRowFunction: f.TableTotalsRowFunction.Max, totalsRowFormula: null },
    }),
    calculatedFormula: k('=[@Qty]*10', { clear: true }),
  },
};

const TABLE_STYLE_OPTIONS = {
  make: (ws) => ({ table: addTable(ws) }),
  update: ({ table }, patch) => table.updateStyleOptions(patch),
  read: ({ table }) => table.getStyleOptions(),
  cases: {
    styleName: k('TableStyleLight9', { clear: true }),
    showFirstColumn: k(FLIP),
    showLastColumn: k(FLIP),
    showRowStripes: k(FLIP),
    showColumnStripes: k(FLIP),
  },
};

const pivotType = (update, read, cases) => ({
  make: addPivot,
  update: (t, patch) => t.pivot[update](patch),
  read: (t) => t.pivot[read](),
  cases,
});

const PIVOT_STYLE_OPTIONS = pivotType('updateStyleOptions', 'getStyleOptions', {
  styleName: k('PivotStyleMedium5', { clear: true }),
  showRowHeaders: k(FLIP),
  showColumnHeaders: k(FLIP),
  showRowStripes: k(FLIP),
  showColumnStripes: k(FLIP),
  showLastColumn: k(FLIP),
});

const PIVOT_GRAND_TOTALS = pivotType('updateGrandTotals', 'getGrandTotals', { rows: k(FLIP), columns: k(FLIP) });

const PIVOT_OPTIONS = pivotType('updateOptions', 'getOptions', {
  showError: k(FLIP),
  errorCaption: k('Oops', { clear: true, patch: (v) => ({ showError: true, errorCaption: v }) }),
  showMissing: k(FLIP),
  missingCaption: k('none', { clear: true, patch: (v) => ({ showMissing: true, missingCaption: v }) }),
  showHeaders: k(FLIP),
  showDrill: k(FLIP),
  showDataTips: k(FLIP),
  gridDropZones: k(FLIP),
  showEmptyRow: k(FLIP),
  showEmptyColumn: k(FLIP),
  showDataDropDown: k(FLIP),
  printDrill: k(FLIP),
  fieldPrintTitles: k(FLIP),
  itemPrintTitles: k(FLIP),
  mergeItem: k(FLIP),
  indent: k(3),
  pageOverThenDown: k(FLIP),
  pageWrap: k(2),
  multipleFieldFilters: k(FLIP),
  preserveFormatting: k(FLIP),
  customListSort: k(FLIP),
  altText: k('Alt', { clear: true }),
  altTextSummary: k('Summary', { clear: true }),
  grandTotalCaption: k('Everything', { clear: true }),
  autoFormatting: k(FLIP),
  enableDrill: k(FLIP),
});

const PIVOT_CACHE_OPTIONS = {
  make: addPivot,
  update: (t, patch) => t.cache.updateOptions(patch),
  read: (t) => t.cache.getOptions(),
  cases: { saveData: k(FLIP), missingItemsLimit: k(5, { clear: true }), refreshOnLoad: k(FLIP) },
};

const PIVOT_FIELD = {
  make: addPivot,
  update: (t, patch) => t.pivot.updateField(0, patch),
  read: (t) => t.pivot.getField(0),
  cases: {
    customName: k('Territory', { clear: true }),
    numberFormat: k('0.00', { clear: true }),
    defaultSubtotal: k(FLIP),
    insertBlankRow: k(FLIP),
    fillDownLabels: k(FLIP),
    showAll: k(FLIP),
    subtotalTop: k(FLIP),
    sortOrder: k(f.PivotSortOrder.Descending),
    sortByField: k('Sum of Amount', {
      clear: true,
      patch: (v) => ({ sortOrder: f.PivotSortOrder.Ascending, sortByField: v }),
    }),
    showDetail: k(FLIP),
  },
};

const PIVOT_ITEM = {
  make: addPivot,
  update: (t, patch) => t.pivot.updateFieldItem(0, 0, patch),
  read: (t) => t.pivot.getFieldItem(0, 0),
  cases: { visible: k(FLIP), showDetail: k(FLIP), caption: k('Eastern', { clear: true }) },
};

// ---- Drawings ----------------------------------------------------------------------------------------------------

const TWO_CELL = { kind: f.AnchorKind.TwoCell, fromRow: 1, fromCol: 1, toRow: 5, toCol: 4 };
const ONE_CELL = { kind: f.AnchorKind.OneCell, fromRow: 1, fromCol: 1, widthEmu: 914400, heightEmu: 457200 };
const addShape = (ws, anchor = TWO_CELL) => ws.addShape({ name: 'Box', presetGeometry: 'rect', anchor });
const shapeMaker = (anchor) => (ws) => ({ shape: addShape(ws, anchor) });

const DRAWING = {
  make: shapeMaker(),
  update: ({ shape }, patch) => shape.update(patch),
  read: ({ shape }) => shape.getDrawingProperties(),
  cases: {
    name: k('Renamed'),
    description: k('Alt text', { clear: true }),
    hidden: k(FLIP),
  },
};

const anchorCase = (v, extra = {}) => k(v, extra);
const DRAWING_ANCHOR = {
  make: shapeMaker(),
  update: ({ shape }, patch) => shape.update({ anchor: patch }),
  read: ({ shape }) => shape.getDrawingProperties().anchor,
  cases: {
    kind: k(f.AnchorKind.OneCell),
    editAs: k(f.AnchorEditAs.OneCell),
    fromRow: anchorCase(2),
    fromCol: anchorCase(2),
    fromRowOffsetEmu: anchorCase(9525),
    fromColOffsetEmu: anchorCase(19050),
    toRow: anchorCase(7),
    toCol: anchorCase(6),
    toRowOffsetEmu: anchorCase(9525),
    toColOffsetEmu: anchorCase(19050),
    widthEmu: { make: shapeMaker(ONE_CELL), ...anchorCase(1828800) },
    heightEmu: { make: shapeMaker(ONE_CELL), ...anchorCase(914400) },
  },
};

const SHAPE = {
  make: shapeMaker(),
  update: ({ shape }, patch) => shape.update(patch),
  read: ({ shape }) => shape.getProperties(),
  cases: {
    rotationDeg: k(30),
    textBox: k(FLIP),
    presetGeometry: k('ellipse'),
    text: k('Hello'),
  },
};

const SLICER = {
  make: (ws) => {
    const table = addTable(ws);
    return { slicer: ws.addSlicer(table, 0, { anchor: { fromRow: 20, fromCol: 3, toRow: 30, toCol: 5 } }) };
  },
  update: ({ slicer }, patch) => slicer.update(patch),
  read: ({ slicer }) => slicer.getProperties(),
  cases: {
    columnCount: k(3),
    sortOrder: k(f.SlicerSortOrder.Descending),
    crossFilter: k(f.SlicerCrossFilter.None),
    rowHeightPt: k(30),
    caption: k('Pick one'),
    styleName: k('SlicerStyleLight3', { clear: true }),
  },
};

const SPARKLINE_COLORS = ['Series', 'Negative', 'Axis', 'Markers', 'First', 'Last', 'High', 'Low'];
const SPARKLINE_GROUP = {
  make: (ws) => {
    fillGrid(ws);
    ws.addSparklineGroup({ type: f.SparklineType.Line, sparklines: [{ source: 'Sheet1!A1:D1', location: 'F1' }] });
    return ws;
  },
  update: (ws, patch) => ws.updateSparklineGroup(0, patch),
  read: (ws) => ws.listSparklineGroups()[0],
  cases: {
    ...Object.fromEntries(
      SPARKLINE_COLORS.map((name) => [
        `color${name}`,
        k('#CC3300', { get: (r) => argb(r[`color${name}`]), want: 0xffcc3300, clear: true }),
      ]),
    ),
    manualMin: k(-2.5, {
      clear: true,
      patch: (v) => ({ minAxisType: f.SparklineAxisType.Custom, manualMin: v }),
    }),
    manualMax: k(7.5, {
      clear: true,
      patch: (v) => ({ maxAxisType: f.SparklineAxisType.Custom, manualMax: v }),
    }),
    lineWidthPt: k(2.25),
    type: k(f.SparklineType.Column),
    markers: k(FLIP),
    high: k(FLIP),
    low: k(FLIP),
    first: k(FLIP),
    last: k(FLIP),
    negative: k(FLIP),
    displayXAxis: k(FLIP),
    displayHidden: k(FLIP),
    rightToLeft: k(FLIP),
    minAxisType: k(f.SparklineAxisType.Group),
    maxAxisType: k(f.SparklineAxisType.Group),
    dateAxis: k(FLIP, { patch: (v) => ({ dateAxisRange: 'Sheet1!A2:D2', dateAxis: v }) }),
    emptyCellsAs: k(f.BlanksAs.Span),
    dateAxisRange: k('Sheet1!A3:D3', {
      get: (r) => plain(r.dateAxisRange?.rect),
      want: rect(2, 0, 2, 3),
      clear: true,
    }),
    sparklines: k(
      [
        { source: 'Sheet1!A2:D2', location: 'G1' },
        { source: 'Sheet1!A3:D3', location: 'G2' },
      ],
      { get: (r) => r.sparklines.length, want: 2 },
    ),
  },
};

const THREADED_COMMENT = {
  make: (ws, wb) => {
    wb.addPerson({ id: PERSON, displayName: 'Alice', userId: 'alice', providerId: 'None' });
    ws.addThreadedComment({
      id: COMMENT,
      text: 'hello world',
      cell: { row: 4, col: 2 },
      personId: PERSON,
      created: { year: 2026, month: 1, day: 2, hour: 3, minute: 4, second: 5 },
    });
    return ws;
  },
  update: (ws, patch) => ws.updateThreadedComment(COMMENT, patch),
  read: (ws) => ws.getThreadedComment(COMMENT),
  cases: {
    text: k('edited'),
    mentions: k([{ personId: PERSON, start: 0, end: 5 }], {
      get: (r) => r.mentions.map((m) => ({ start: m.start, end: m.end })),
      want: [{ start: 0, end: 5 }],
    }),
    done: k(FLIP),
  },
};

// ---- Every patch type, by its public JS name ---------------------------------------------------------------------

const TYPES = {
  ThreadedCommentPatch: THREADED_COMMENT,
  SheetProtectionPatch: SHEET_PROTECTION,
  PageSetupPatch: PAGE_SETUP,
  PageMarginsPatch: PAGE_MARGINS,
  PrintOptionsPatch: PRINT_OPTIONS,
  HeaderFooterPatch: HEADER_FOOTER,
  TableStyleOptionsPatch: TABLE_STYLE_OPTIONS,
  PivotOptionsPatch: PIVOT_OPTIONS,
  PivotCacheOptionsPatch: PIVOT_CACHE_OPTIONS,
  SheetFormatDefaultsPatch: SHEET_FORMAT_DEFAULTS,
  ChartFillPatch: CHART_FILL,
  ChartLinePatch: CHART_LINE,
  ChartMarkerPatch: CHART_MARKER,
  ChartLabelsPatch: CHART_LABELS,
  ChartSeriesPatch: CHART_SERIES,
  ChartPointPatch: CHART_POINT,
  ChartPlotPatch: CHART_PLOT,
  ChartAxisPatch: CHART_AXIS,
  ChartLegendPatch: CHART_LEGEND,
  ChartDisplayPatch: CHART_DISPLAY,
  PivotGrandTotalsPatch: PIVOT_GRAND_TOTALS,
  TableStyleDefaultsPatch: TABLE_STYLE_DEFAULTS,
  TablePatch: TABLE,
  TableColumnPatch: TABLE_COLUMN,
  SparklineGroupPatch: SPARKLINE_GROUP,
  ProtectedRangePatch: PROTECTED_RANGE,
  DrawingPatch: DRAWING,
  DrawingAnchorPatch: DRAWING_ANCHOR,
  ShapePatch: SHAPE,
  SlicerPatch: SLICER,
  SheetViewPatch: SHEET_VIEW,
  PivotFieldPatch: PIVOT_FIELD,
  PivotItemPatch: PIVOT_ITEM,
  PivotStyleOptionsPatch: PIVOT_STYLE_OPTIONS,
  IterativeSettingsPatch: ITERATIVE,
  DataValidationPatch: VALIDATION,
  ConditionalFormatPatch: CF_RULE,
  CellFormatPatch: CELL_FORMAT,
  FontPatch: FONT,
  FillPatch: FILL,
  BorderSidePatch: BORDER_SIDE,
  BorderPatch: BORDER,
  AlignmentPatch: ALIGNMENT,
  CellProtectionPatch: PROTECTION,
};

/** Keys with no public readback, by type; each value is the reason. */
const EXEMPT = {
  ChartSeriesPatch: {
    size: 'a bubble plot cannot be created through the public API (addChart and setType refuse it), so no series accepts a size',
  },
};

// ---- Runner ------------------------------------------------------------------------------------------------------

for (const [typeName, def] of Object.entries(TYPES)) {
  describe(`${typeName}: every key applies`, () => {
    for (const [key, spec] of Object.entries(def.cases)) {
      it(key, () => {
        withSheet((ws, wb) => {
          const target = (spec.make ?? def.make)(ws, wb);
          const get = spec.get ?? ((record) => record?.[key]);
          const before = plain(get(def.read(target)));
          const value = spec.v === FLIP ? !before : spec.v;
          const patch = typeof spec.patch === 'function' ? spec.patch(value) : (spec.patch ?? { [key]: value });
          const want = 'want' in spec ? spec.want : value;
          assert.notDeepEqual(before, plain(want), `${key}: the value must differ from the setup state`);
          def.update(target, patch);
          assert.deepEqual(plain(get(def.read(target))), plain(want), `${typeName}.${key} was not applied`);
          if (spec.clear) {
            def.update(target, spec.clearPatch ?? { [key]: null });
            assert.deepEqual(
              plain(get(def.read(target))),
              plain(spec.cleared ?? null),
              `${typeName}.${key} was not cleared`,
            );
          }
        });
      });
    }
  });
}

// ---- Coverage of the ledger --------------------------------------------------------------------------------------

const camel = (s) => s.replace(/_([a-z0-9])/g, (_, c) => c.toUpperCase());

/** C field groups that map to one JS key, per patch type. */
const GROUPS = {
  ConditionalFormatPatch: [
    [/^color_scale/, 'colorScale'],
    [/^data_bar/, 'dataBar'],
    [/^icon_set/, 'iconSet'],
    [/^op$/, 'operator'],
  ],
  DataValidationPatch: [[/^op$/, 'operator']],
  SheetViewPatch: [
    [/^pane_top_left/, 'paneTopLeft'],
    [/^top_left/, 'topLeft'],
  ],
  SparklineGroupPatch: [
    [/^date_axis_ref$/, 'dateAxisRange'],
    [/^items$/, 'sparklines'],
  ],
};

function jsKey(typeName, cKey) {
  for (const [re, name] of GROUPS[typeName] ?? []) if (re.test(cKey)) return name;
  return camel(cKey);
}

/** The `role: patch` records of the ledger that exist on a JS surface. */
async function ledgerPatches() {
  const text = await readFile(path.join(repoRoot, 'tools/api/names.yaml'), 'utf8');
  const out = [];
  for (const m of text.matchAll(/^ {2}(record\.\w+):\n((?: {4}.*\n)+)/gm)) {
    const body = m[2];
    if (!/role: patch/.test(body)) continue;
    const surfaces = (body.match(/surfaces: \[(.*?)\]/)?.[1] ?? '').split(',').map((s) => s.trim());
    if (!SURFACES.some((s) => surfaces.includes(s))) continue;
    out.push({ id: m[1], c: body.match(/\bc: (\w+)/)?.[1] ?? null, js: body.match(/\bjs: (\w+)/)?.[1] ?? null });
  }
  return out;
}

/** The field names of an `export interface <name>` in the d.ts. */
function interfaceKeys(dts, name) {
  const m = dts.match(new RegExp(`export interface ${name} \\{([\\s\\S]*?)\\n\\}`));
  return m ? [...m[1].matchAll(/^ {2}(?:readonly )?(\w+)\??:/gm)].map((x) => x[1]) : [];
}

/** The fields of a C patch struct that carry a sibling `<field>_op`. */
function structKeys(header, cName) {
  const end = header.indexOf(`} ${cName};`);
  assert.ok(end > 0, `${cName} is not declared in formulon_c.h`);
  const body = header.slice(header.lastIndexOf('typedef struct', end), end);
  return [...body.matchAll(/\b(\w+)_op;/g)].map((x) => x[1]);
}

describe('every key of every patch record has a case', async () => {
  const records = await ledgerPatches();
  const header = await readFile(path.join(repoRoot, 'src/c_api/formulon_c.h'), 'utf8');
  const dts = await readFile(path.join(repoRoot, 'src/wasm/formulon.d.ts'), 'utf8');

  it('the ledger lists patch records', () => {
    assert.ok(records.length >= 40, `only ${records.length} patch records found`);
  });

  for (const rec of records) {
    it(rec.id, () => {
      assert.ok(rec.js, `${rec.id} has no JS type name`);
      const def = TYPES[rec.js];
      assert.ok(def, `${rec.js} has no cases`);
      const keys = rec.c
        ? new Set(structKeys(header, rec.c).map((key) => jsKey(rec.js, key)))
        : new Set(interfaceKeys(dts, rec.js));
      assert.ok(keys.size > 0, `${rec.js} has no keys`);
      const missing = [...keys].filter((key) => !(key in def.cases) && !(key in (EXEMPT[rec.js] ?? {})));
      assert.deepEqual(missing, [], `${rec.js} keys without a case`);
      for (const key of Object.keys(def.cases)) {
        assert.ok(keys.has(key), `${rec.js}.${key} is not a key of the record`);
      }
    });
  }
});
