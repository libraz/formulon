// Acceptance of the typed enum tables and the value shapes of the public surface: exact ordinals, records that
// return enum members and native booleans instead of strings and 0/1, date inputs that accept a Date everywhere,
// leading `=` on formula reads, and the native value unions. FORMULON_NPM_ENTRY selects the staged subpath: "."
// (default) or "./threads".

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

const expectInvalid = (fn) => expectCode(fn, f.StatusCode.InvalidArgument);

/** A table is frozen and lists exactly `expected`, members declared in ordinal order. */
function assertTable(name, expected) {
  const table = f[name];
  assert.ok(table, `${name} is exported`);
  assert.ok(Object.isFrozen(table), `${name} is frozen`);
  assert.deepEqual({ ...table }, expected);
  assert.deepEqual(Object.keys(table), Object.keys(expected), `${name} members are declared in ordinal order`);
}

const inOrder = (...names) => Object.fromEntries(names.map((n, i) => [n, i]));

const PERSON = '{00000000-0000-4000-8000-0000000000A1}';

const rect = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });

/** A1 holds a region label, A2:A4 numbers 1..3 under a header, B1:B4 likewise. */
function fillColumns(ws) {
  ws.setText('A1', 'Region');
  ws.setText('B1', 'Amount');
  for (let i = 0; i < 3; i++) {
    ws.setNumber(i + 1, 0, i + 1);
    ws.setNumber(i + 1, 1, (i + 1) * 10);
  }
}

function addColumnChart(ws, kind = f.ChartPlotType.Bar) {
  fillColumns(ws);
  return ws.addChart({ type: kind, range: 'Sheet1!$A$1:$B$4', legend: f.ChartLegendPosition.Right });
}

/** A cache and a pivot over Region / Amount / Date; the date column is shared-item index 2. */
function buildPivot(wb) {
  const ws = wb.sheet(0);
  ws.setText('A1', 'Region');
  ws.setText('B1', 'Amount');
  ws.setText('C1', 'Date');
  const rows = [
    ['East', 10, 45292],
    ['West', 30, 45323],
    ['East', 5, 45658],
  ];
  rows.forEach(([region, amount, date], i) => {
    ws.setText(i + 1, 0, region);
    ws.setNumber(i + 1, 1, amount);
    ws.setNumber(i + 1, 2, date);
  });
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ range: 'A1:C4', sheet: 'Sheet1' });
  for (const name of ['Region', 'Amount', 'Date']) cache.addField(name);
  for (const item of ['East', 'West']) cache.addFieldSharedItem(0, item);
  rows.forEach(([region, amount, date]) => {
    const rec = cache.addRecord();
    cache.setRecordValue(rec, 0, region === 'East' ? 0 : 1);
    cache.setRecordValue(rec, 1, amount);
    cache.setRecordValue(rec, 2, date);
  });
  const pivot = ws.addPivot({ name: 'Pivot1', cacheId: cache.cacheId, anchor: { row: 0, col: 6 } });
  pivot.setFieldAxis('Region', f.PivotAxis.Row);
  const region = 0;
  pivot.setFieldAxis('Amount', f.PivotAxis.Value);
  const amount = 1;
  pivot.setFieldAxis('Date', f.PivotAxis.Hidden ?? 2);
  const date = 2;
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 });
  return { cache, pivot, region, amount, date };
}

describe('enum tables list their members at the exact ordinals', () => {
  it('ComparisonOperator is one table for conditional formats, validation and filters', () => {
    assertTable(
      'ComparisonOperator',
      inOrder(
        'Between',
        'NotBetween',
        'Equal',
        'NotEqual',
        'LessThan',
        'LessThanOrEqual',
        'GreaterThan',
        'GreaterThanOrEqual',
      ),
    );
    for (const gone of ['CfOperator', 'ValidationOperator', 'FilterOperator']) {
      assert.equal(f[gone], undefined, `${gone} is replaced by ComparisonOperator`);
    }
  });

  it('DatePeriod lists the conditional-format periods in schema order', () => {
    assertTable(
      'DatePeriod',
      inOrder(
        'Today',
        'Yesterday',
        'Tomorrow',
        'Last7Days',
        'ThisMonth',
        'LastMonth',
        'NextMonth',
        'ThisWeek',
        'LastWeek',
        'NextWeek',
      ),
    );
    assert.equal(f.CfTimePeriod, undefined);
  });

  it('periods shared by DatePeriod, DynamicFilterType and PivotFilterKind are spelled alike', () => {
    for (const period of Object.keys(f.DatePeriod)) {
      if (period === 'Last7Days') continue;
      assert.equal(typeof f.DynamicFilterType[period], 'number', `DynamicFilterType.${period}`);
      assert.equal(typeof f.PivotFilterKind[period], 'number', `PivotFilterKind.${period}`);
      assert.equal(f.PivotFilterKind[`Date${period}`], undefined, `PivotFilterKind.Date${period} lost its prefix`);
    }
    for (const name of ['NextQuarter', 'YearToDate', 'Q1', 'Q4', 'M1', 'M12']) {
      assert.equal(typeof f.PivotFilterKind[name], 'number', `PivotFilterKind.${name}`);
      assert.equal(f.PivotFilterKind[`Date${name}`], undefined, `PivotFilterKind.Date${name}`);
    }
    assert.equal(typeof f.PivotFilterKind.DateBetween, 'number', 'comparison kinds keep the Date prefix');
  });

  it('BlanksAs has Zero at 0 and serves charts and sparklines', () => {
    assertTable('BlanksAs', { Zero: 0, Gap: 1, Span: 2 });
    assert.equal(f.ChartBlanksAs, undefined);
    assert.equal(f.SparklineEmptyCells, undefined);
  });

  it('ChartLegendPosition puts the schema default Right at 0 and has no None member', () => {
    assertTable('ChartLegendPosition', { Right: 0, Bottom: 1, TopRight: 2, Left: 3, Top: 4 });
  });

  it('PivotSortOrder starts at Manual and SlicerSortOrder at Ascending', () => {
    assertTable('PivotSortOrder', { Manual: 0, Ascending: 1, Descending: 2 });
    assertTable('SlicerSortOrder', { Ascending: 0, Descending: 1 });
  });

  it('AnchorKind and AnchorEditAs start at Default, and ShapePaint is gone', () => {
    assertTable('AnchorKind', { Default: 0, OneCell: 1, TwoCell: 2, Absolute: 3 });
    assertTable('AnchorEditAs', { Default: 0, OneCell: 1, TwoCell: 2, Absolute: 3 });
    assert.equal('ShapePaint' in f, false);
  });

  it('an omitted anchor kind takes the operation default and a read reports the stored kind and edit-as', () => {
    withSheet((ws) => {
      const box = { fromRow: 1, fromCol: 1, toRow: 4, toCol: 4 };
      const shape = ws.addShape({ anchor: box });
      assert.equal(shape.getDrawingProperties().anchor.kind, f.AnchorKind.TwoCell);
      const oneCell = ws.addShape({ anchor: { kind: f.AnchorKind.OneCell, widthEmu: 9525, heightEmu: 9525 } });
      assert.equal(oneCell.getDrawingProperties().anchor.kind, f.AnchorKind.OneCell);
      for (const handle of ws.listDrawings()) {
        const { anchor } = handle.getDrawingProperties();
        assert.notEqual(anchor.kind, f.AnchorKind.Default);
        assert.notEqual(anchor.editAs, f.AnchorEditAs.Default);
      }
      const free = ws.addShape({ anchor: { ...box, editAs: f.AnchorEditAs.Absolute } });
      assert.equal(free.getDrawingProperties().anchor.editAs, f.AnchorEditAs.Absolute);
      assert.equal(free.getDrawingProperties().anchor.kind, f.AnchorKind.TwoCell);
    });
  });

  it('a shape without its own paint reads null fill and line, and a null patch removes the paint', () => {
    withSheet((ws) => {
      const anchor = { fromRow: 1, fromCol: 1, toRow: 4, toCol: 4 };
      const plain = ws.addShape({ anchor });
      assert.equal(plain.getProperties().fill, null);
      assert.equal(plain.getProperties().line, null);
      const painted = ws.addShape({
        anchor,
        fill: { color: '#336699' },
        line: { none: true },
      });
      assert.equal(painted.getProperties().fill.color.argb, 0xff336699);
      painted.update({ fill: null, line: null });
      assert.equal(painted.getProperties().fill, null);
      assert.equal(painted.getProperties().line, null);
    });
  });

  it('TableTotalsRowFunction follows the schema order', () => {
    assertTable(
      'TableTotalsRowFunction',
      inOrder('None', 'Sum', 'Min', 'Max', 'Average', 'Count', 'CountNumbers', 'StdDev', 'Var', 'Custom'),
    );
  });

  it('TableStyleElementType lists the 28 style regions in schema order', () => {
    assertTable(
      'TableStyleElementType',
      inOrder(
        'WholeTable',
        'HeaderRow',
        'TotalRow',
        'FirstColumn',
        'LastColumn',
        'FirstRowStripe',
        'SecondRowStripe',
        'FirstColumnStripe',
        'SecondColumnStripe',
        'FirstHeaderCell',
        'LastHeaderCell',
        'FirstTotalCell',
        'LastTotalCell',
        'FirstSubtotalColumn',
        'SecondSubtotalColumn',
        'ThirdSubtotalColumn',
        'FirstSubtotalRow',
        'SecondSubtotalRow',
        'ThirdSubtotalRow',
        'BlankRow',
        'FirstColumnSubheading',
        'SecondColumnSubheading',
        'ThirdColumnSubheading',
        'FirstRowSubheading',
        'SecondRowSubheading',
        'ThirdRowSubheading',
        'PageFieldLabels',
        'PageFieldValues',
      ),
    );
  });

  it('chart axis, chart data and width model tables have their documented members', () => {
    assertTable('ChartTickLabelPosition', inOrder('NextTo', 'High', 'Low', 'None'));
    assertTable('ChartDomain', inOrder('Category', 'Xy'));
    assertTable('WidthPlatform', inOrder('Win', 'Mac'));
  });

  it('PhoneticType and PhoneticAlignment keep the stored numbering', () => {
    assertTable('PhoneticType', inOrder('HalfwidthKatakana', 'FullwidthKatakana', 'Hiragana', 'NoConversion'));
    assertTable('PhoneticAlignment', inOrder('NoControl', 'Left', 'Center', 'Distributed'));
  });

  it('PageOrder is a closed enum and PaperSize names the eight known codes', () => {
    assertTable('PageOrder', { DownThenOver: 0, OverThenDown: 1 });
    assertTable('PaperSize', { Letter: 1, Legal: 5, A3: 8, A4: 9, A5: 11, B4: 12, B5: 13, A6: 70 });
  });

  it('FunctionAvailability, EffectiveStyleSource and ChartPointState are numeric tables', () => {
    assertTable(
      'FunctionAvailability',
      inOrder('Implemented', 'ImplementedUnverified', 'EnvironmentBound', 'UnavailableStub'),
    );
    assertTable('EffectiveStyleSource', inOrder('Cell', 'Row', 'Column', 'Default'));
    assertTable('ChartPointState', { Value: 0, Gap: 1, Connect: 2 });
  });

  it('IconSetName appends the four custom-only sets after the seventeen built-ins', () => {
    const names = Object.keys(f.IconSetName);
    assert.equal(names[0], 'ThreeArrows');
    assert.equal(f.IconSetName.FiveQuarters, 16);
    assert.deepEqual(
      Object.fromEntries(['ThreeStars', 'ThreeTriangles', 'FiveBoxes', 'NoIcons'].map((n) => [n, f.IconSetName[n]])),
      { ThreeStars: 17, ThreeTriangles: 18, FiveBoxes: 19, NoIcons: 20 },
    );
    assert.equal(names.length, 21);
  });

  it('removed operand-kind table is gone from the module', () => {
    assert.equal(f.PivotFilterOperandKind, undefined);
  });
});

describe('the merged comparison operator is accepted by every feature that compares', () => {
  const ALL = Object.entries({
    Between: 0,
    NotBetween: 1,
    Equal: 2,
    NotEqual: 3,
    LessThan: 4,
    LessThanOrEqual: 5,
    GreaterThan: 6,
    GreaterThanOrEqual: 7,
  });

  it('a conditional format cellIs rule reads back each member it was given', () => {
    withSheet((ws) => {
      for (const [name] of ALL) {
        const operator = f.ComparisonOperator[name];
        ws.addCfRule({ type: f.CfRuleType.CellIs, ranges: 'A1:A3', operator, formula1: '1', formula2: '9' });
        const rule = ws.listCfRules().at(-1);
        assert.equal(rule.operator, operator, name);
      }
    });
  });

  it('a data validation rule reads back each member it was given', () => {
    withSheet((ws) => {
      for (const [name] of ALL) {
        const operator = f.ComparisonOperator[name];
        ws.addValidation({ type: f.ValidationType.Whole, ranges: 'B1:B3', operator, formula1: '1', formula2: '9' });
        assert.equal(ws.listValidations().at(-1).operator, operator, name);
      }
    });
  });

  it('a custom auto filter reads back the six single-operand members', () => {
    withSheet((ws) => {
      fillColumns(ws);
      const single = ['Equal', 'NotEqual', 'LessThan', 'LessThanOrEqual', 'GreaterThan', 'GreaterThanOrEqual'];
      for (const name of single) {
        const operator = f.ComparisonOperator[name];
        ws.setAutoFilter({
          range: 'A1:B4',
          columns: [{ columnOffset: 0, kind: f.FilterKind.Custom, conditions: [{ operator, value: '2' }] }],
        });
        assert.equal(ws.getAutoFilter().columns[0].conditions[0].operator, operator, name);
      }
    });
  });

  it('a custom auto filter rejects the range operators because it always names one comparison', () => {
    withSheet((ws) => {
      fillColumns(ws);
      for (const name of ['Between', 'NotBetween']) {
        expectInvalid(() =>
          ws.setAutoFilter({
            range: 'A1:B4',
            columns: [
              {
                columnOffset: 0,
                kind: f.FilterKind.Custom,
                conditions: [{ operator: f.ComparisonOperator[name], value: '2' }],
              },
            ],
          }),
        );
      }
    });
  });

  it('an operator name where a table member is expected is a TypeError', () => {
    withSheet((ws) => {
      assert.throws(
        () => ws.addCfRule({ type: f.CfRuleType.CellIs, ranges: 'A1', operator: 'equal', formula1: '1' }),
        TypeError,
      );
      assert.throws(
        () => ws.addValidation({ type: f.ValidationType.Whole, ranges: 'B1', operator: 'between', formula1: '1' }),
        TypeError,
      );
    });
  });

  it('an out-of-range operator ordinal is InvalidArgument', () => {
    withSheet((ws) => {
      expectInvalid(() => ws.addCfRule({ type: f.CfRuleType.CellIs, ranges: 'A1', operator: 8, formula1: '1' }));
      expectInvalid(() => ws.addValidation({ type: f.ValidationType.Whole, ranges: 'B1', operator: 8, formula1: '1' }));
    });
  });

  it('auto filter sorting keeps a boolean descending flag and no direction enum', () => {
    withSheet((ws) => {
      fillColumns(ws);
      ws.setAutoFilter({
        range: 'A1:B4',
        sort: { range: 'A2:B4', conditions: [{ range: 'B2:B4', descending: true }] },
      });
      const condition = ws.getAutoFilter().sort.conditions[0];
      assert.equal(condition.descending, true);
      assert.equal(typeof condition.descending, 'boolean');
      ws.setAutoFilter({ range: 'A1:B4', sort: { range: 'A2:B4', conditions: [{ range: 'B2:B4' }] } });
      assert.equal(ws.getAutoFilter().sort.conditions[0].descending, false);
    });
  });
});

describe('date period rules use the reordered DatePeriod table', () => {
  function pinned(wb, ws) {
    wb.setPinnedNow({ year: 2020, month: 1, day: 15, hour: 12, minute: 0, second: 0 });
    // 2020-01-20 is later in the pinned month and in the following week.
    ws.setNumber('A1', 43850);
    wb.recalc();
  }

  it('ThisMonth matches a date later in the month that ThisWeek leaves out', () => {
    withSheet((ws, wb) => {
      pinned(wb, ws);
      ws.addCfRule({ type: f.CfRuleType.TimePeriod, ranges: 'A1', timePeriod: f.DatePeriod.ThisMonth });
      assert.equal(ws.evaluateCfRange('A1').length, 1);
    });
    withSheet((ws, wb) => {
      pinned(wb, ws);
      ws.addCfRule({ type: f.CfRuleType.TimePeriod, ranges: 'A1', timePeriod: f.DatePeriod.ThisWeek });
      assert.equal(ws.evaluateCfRange('A1').length, 0);
    });
    withSheet((ws, wb) => {
      pinned(wb, ws);
      ws.addCfRule({ type: f.CfRuleType.TimePeriod, ranges: 'A1', timePeriod: f.DatePeriod.NextWeek });
      assert.equal(ws.evaluateCfRange('A1').length, 1);
    });
  });

  it('a rule reads back the DatePeriod member it was given', () => {
    withSheet((ws) => {
      for (const [name, ordinal] of Object.entries(f.DatePeriod)) {
        ws.addCfRule({ type: f.CfRuleType.TimePeriod, ranges: 'A1', timePeriod: ordinal });
        assert.equal(ws.listCfRules().at(-1).timePeriod, ordinal, name);
      }
    });
  });
});

describe('reordered tables round-trip by member through save and load', () => {
  function reload(wb) {
    const bytes = wb.save();
    return f.Workbook.load(bytes);
  }

  it('a legend position survives save and load for every member', () => {
    for (const [name, position] of Object.entries(f.ChartLegendPosition)) {
      withSheet((ws, wb) => {
        const id = addColumnChart(ws);
        ws.chart(id).setLegend({ position, overlay: false, hiddenEntries: [] });
        assert.equal(ws.chart(id).getProperties().legend.position, position, `${name} in session`);
        const loaded = reload(wb);
        try {
          const chart = loaded.sheet(0).listCharts()[0];
          assert.equal(
            loaded
              .sheet(0)
              .chart(chart.objectId ?? chart)
              .getProperties().legend.position,
            position,
            name,
          );
        } finally {
          loaded.dispose();
        }
      });
    }
  });

  it('a chart without a legend reads null rather than a position member', () => {
    withSheet((ws) => {
      fillColumns(ws);
      const id = ws.addChart({ type: f.ChartPlotType.Bar, range: 'Sheet1!$A$1:$B$4', legend: null });
      assert.equal(ws.chart(id).getProperties().legend, null);
    });
  });

  it('a pivot field sort order survives save and load for every member', () => {
    for (const [name, order] of Object.entries(f.PivotSortOrder)) {
      withSheet((_ws, wb) => {
        const { pivot, region } = buildPivot(wb);
        pivot.updateField(region, { sortOrder: order });
        assert.equal(pivot.listFields()[region].sortOrder, order, `${name} in session`);
        const loaded = reload(wb);
        try {
          assert.equal(loaded.sheet(0).pivot(0).listFields()[region].sortOrder, order, name);
        } finally {
          loaded.dispose();
        }
      });
    }
  });

  it('a slicer sort order reads back the member it was given', () => {
    for (const [name, sortOrder] of Object.entries(f.SlicerSortOrder)) {
      withSheet((ws) => {
        fillColumns(ws);
        ws.addTable({ range: 'A1:B4', name: 'Sales' });
        ws.addSlicer('Sales', 0, {
          anchor: { fromRow: 4, fromCol: 3, toRow: 14, toCol: 5 },
          caption: 'Pick',
          sortOrder,
        });
        assert.equal(ws.listSlicers()[0].getProperties().sortOrder, sortOrder, name);
      });
    }
  });
});

describe('blanks handling is one table for charts and sparklines', () => {
  it('a chart reads back each BlanksAs member and null once removed', () => {
    withSheet((ws) => {
      const id = addColumnChart(ws);
      for (const [name, blanksAs] of Object.entries(f.BlanksAs)) {
        ws.chart(id).updateDisplay({ blanksAs });
        assert.equal(ws.chart(id).getProperties().blanksAs, blanksAs, name);
      }
      ws.chart(id).updateDisplay({ blanksAs: null });
      assert.equal(ws.chart(id).getProperties().blanksAs, null);
    });
  });

  it('a sparkline group reads back each BlanksAs member', () => {
    withSheet((ws) => {
      fillColumns(ws);
      for (const [name, emptyCellsAs] of Object.entries(f.BlanksAs)) {
        ws.addSparklineGroup({ emptyCellsAs, sparklines: [{ source: 'Sheet1!A2:A4', location: 'F1' }] });
        assert.equal(ws.listSparklineGroups().at(-1).emptyCellsAs, emptyCellsAs, name);
      }
    });
  });
});

describe('closed string sets are numeric tables on input and output', () => {
  function totalsTable(ws) {
    fillColumns(ws);
    return ws.addTable({ range: 'A1:B4', name: 'Sales', totalsRow: true });
  }

  it('a totals-row function reads back as its TableTotalsRowFunction member', () => {
    withSheet((ws) => {
      const table = totalsTable(ws);
      for (const [name, fn] of Object.entries(f.TableTotalsRowFunction)) {
        const totals =
          fn === f.TableTotalsRowFunction.Custom
            ? { totalsRowFunction: fn, totalsRowFormula: '=SUM(Sales[Amount])' }
            : { totalsRowFunction: fn };
        table.setColumnTotals(1, totals);
        const column = table.getProperties().columns[1];
        assert.equal(column.totalsRowFunction, fn, name);
        assert.equal(typeof column.totalsRowFunction, 'number');
      }
    });
  });

  it('a totals-row function given as a name or out of range is rejected', () => {
    withSheet((ws) => {
      const table = totalsTable(ws);
      assert.throws(() => table.setColumnTotals(1, { totalsRowFunction: 'sum' }), TypeError);
      expectInvalid(() => table.setColumnTotals(1, { totalsRowFunction: 10 }));
    });
  });

  it('a table style element reads back its TableStyleElementType member', () => {
    withSheet((_ws, wb) => {
      const elements = Object.values(f.TableStyleElementType).map((type) => ({ type }));
      wb.setTableStyle({ name: 'AllRegions', elements });
      const read = wb.getTableStyle('AllRegions');
      assert.deepEqual(
        read.elements.map((e) => e.type),
        Object.values(f.TableStyleElementType),
      );
    });
  });

  it('a table style element given as a name or out of range is rejected', () => {
    withSheet((_ws, wb) => {
      assert.throws(() => wb.setTableStyle({ name: 'S', elements: [{ type: 'headerRow' }] }), TypeError);
      expectInvalid(() => wb.setTableStyle({ name: 'S', elements: [{ type: 28 }] }));
    });
  });

  it('an axis reads back its ChartTickLabelPosition member and None is a value', () => {
    withSheet((ws) => {
      const id = addColumnChart(ws);
      const axis = ws
        .chart(id)
        .getProperties()
        .axes.find((a) => a.kind === f.ChartAxisKind.Value);
      for (const [name, tickLabelPosition] of Object.entries(f.ChartTickLabelPosition)) {
        ws.chart(id).updateAxis(axis.id, { tickLabelPosition });
        const read = ws
          .chart(id)
          .getProperties()
          .axes.find((a) => a.id === axis.id);
        assert.equal(read.tickLabelPosition, tickLabelPosition, name);
      }
      assert.equal(
        ws
          .chart(id)
          .getProperties()
          .axes.find((a) => a.id === axis.id).tickLabelPosition,
        f.ChartTickLabelPosition.None,
      );
    });
  });

  it('a tick label position given as a name or out of range is rejected', () => {
    withSheet((ws) => {
      const id = addColumnChart(ws);
      const axis = ws.chart(id).getProperties().axes[0];
      assert.throws(() => ws.chart(id).updateAxis(axis.id, { tickLabelPosition: 'low' }), TypeError);
      expectInvalid(() => ws.chart(id).updateAxis(axis.id, { tickLabelPosition: 4 }));
    });
  });

  it('chart data names its domain with ChartDomain', () => {
    withSheet((ws) => {
      const category = addColumnChart(ws);
      assert.equal(ws.chart(category).getData().domain, f.ChartDomain.Category);
      const xy = ws.addChart({ type: f.ChartPlotType.Scatter, range: 'Sheet1!$A$1:$B$4' });
      assert.equal(ws.chart(xy).getData().domain, f.ChartDomain.Xy);
    });
  });

  it('the width model reports its platform as a WidthPlatform member', () => {
    withSheet((ws) => {
      const model = ws.getWidthModel();
      assert.equal(typeof model.platform, 'number');
      assert.equal(model.platform, f.WidthPlatform.Win);
    });
  });

  it('a custom icon level takes IconSetName members including the custom-only sets', () => {
    withSheet((ws) => {
      fillColumns(ws);
      const rule = (icons) => ({
        type: f.CfRuleType.IconSet,
        ranges: 'B2:B4',
        iconSet: {
          name: f.IconSetName.ThreeArrows,
          thresholds: [
            { type: f.CfValueObjectType.Percent, value: '33' },
            { type: f.CfValueObjectType.Percent, value: '67' },
          ],
          icons,
        },
      });
      ws.addCfRule(
        rule([
          { iconSet: f.IconSetName.ThreeStars, iconIndex: 0 },
          { iconSet: f.IconSetName.FiveBoxes, iconIndex: 2 },
          { iconSet: f.IconSetName.NoIcons, iconIndex: 0 },
        ]),
      );
      const read = ws.listCfRules().at(-1).iconSet.icons;
      assert.deepEqual(
        read.map((i) => i.iconSet),
        [f.IconSetName.ThreeStars, f.IconSetName.FiveBoxes, f.IconSetName.NoIcons],
      );
      assert.throws(
        () => ws.addCfRule(rule([{ iconSet: '3Stars', iconIndex: 0 }, ...Array(2).fill({ iconSet: 0, iconIndex: 0 })])),
        TypeError,
      );
    });
  });

  it('a rule drawn set rejects the custom-only IconSetName members', () => {
    withSheet((ws) => {
      fillColumns(ws);
      expectInvalid(() =>
        ws.addCfRule({
          type: f.CfRuleType.IconSet,
          ranges: 'B2:B4',
          iconSet: {
            name: f.IconSetName.ThreeStars,
            thresholds: [
              { type: f.CfValueObjectType.Percent, value: '33' },
              { type: f.CfValueObjectType.Percent, value: '67' },
            ],
          },
        }),
      );
    });
  });
});

describe('numbers that were bare become enum members', () => {
  it('phonetic properties read back the PhoneticType and PhoneticAlignment given', () => {
    withSheet((ws) => {
      ws.setText('A1', '漢字');
      ws.setCellPhoneticRuns('A1', [{ start: 0, end: 2, text: 'かんじ' }]);
      ws.setCellPhoneticProperties('A1', {
        fontIndex: 0,
        type: f.PhoneticType.Hiragana,
        alignment: f.PhoneticAlignment.Distributed,
      });
      const read = ws.getCellPhoneticProperties('A1');
      assert.equal(read.type, f.PhoneticType.Hiragana);
      assert.equal(read.alignment, f.PhoneticAlignment.Distributed);
      assert.equal(read.stated, true);
      assert.equal('present' in read, false);
    });
  });

  it('phonetic properties given as names or out of range are rejected', () => {
    withSheet((ws) => {
      ws.setText('A1', '漢字');
      ws.setCellPhoneticRuns('A1', [{ start: 0, end: 2, text: 'かんじ' }]);
      const bad = (over) => () => ws.setCellPhoneticProperties('A1', { fontIndex: 0, type: 0, alignment: 0, ...over });
      assert.throws(bad({ type: 'hiragana' }), TypeError);
      assert.throws(bad({ alignment: 'left' }), TypeError);
      expectInvalid(bad({ type: 4 }));
      expectInvalid(bad({ alignment: 4 }));
    });
  });

  it('pagination reports the page order as a PageOrder member', () => {
    withSheet((ws) => {
      fillColumns(ws);
      assert.equal(ws.getPagination().pageOrder, f.PageOrder.DownThenOver);
      ws.ooxml.setPageSetupXml('<pageSetup pageOrder="overThenDown"/>');
      assert.equal(ws.getPagination().pageOrder, f.PageOrder.OverThenDown);
    });
  });

  it('a paper size stays a plain number: named codes and unknown codes round-trip', () => {
    withSheet((ws) => {
      for (const code of [...Object.values(f.PaperSize), 118]) {
        ws.updatePageSetup({ paperSize: code });
        assert.equal(ws.getPageSetup().paperSize, code);
      }
    });
  });

  it('the effective style source of a plain cell is an EffectiveStyleSource member', () => {
    withSheet((ws) => {
      const source = ws.ooxml.getEffectiveStyle('Z1').source;
      assert.ok(Object.values(f.EffectiveStyleSource).includes(source));
      assert.equal(source, f.EffectiveStyleSource.Default);
    });
  });

  it('function metadata reports availability as a FunctionAvailability member', () => {
    const availability = f.getFunctionMetadata('SUM').availability;
    assert.equal(availability, f.FunctionAvailability.Implemented);
  });

  it('a date-grouped pivot field reports its periods as a bitwise OR of PivotDatePeriod values', () => {
    withSheet((_ws, wb) => {
      const { pivot, date } = buildPivot(wb);
      const periods = f.PivotDatePeriod.Months | f.PivotDatePeriod.Years;
      pivot.setFieldDateGroup(date, { periods });
      assert.equal(pivot.listFields()[date].datePeriods, periods);
    });
  });

  it('theme script fonts take a FontScheme and reject None', () => {
    withSheet((_ws, wb) => {
      wb.setThemeScriptFont({ scheme: f.FontScheme.Major, script: 'Zzzz', typeface: 'Face Major' });
      wb.setThemeScriptFont({ scheme: f.FontScheme.Minor, script: 'Zzzz', typeface: 'Face Minor' });
      const major = wb.getThemeScriptFont({ scheme: f.FontScheme.Major, script: 'Zzzz' });
      assert.equal(major.scheme, f.FontScheme.Major);
      assert.equal(major.typeface, 'Face Major');
      assert.equal(wb.getThemeScriptFont({ scheme: f.FontScheme.Minor, script: 'Zzzz' }).typeface, 'Face Minor');
      assert.equal('major' in major, false);
      expectInvalid(() => wb.setThemeScriptFont({ scheme: f.FontScheme.None, script: 'Zzzz', typeface: 'X' }));
      assert.throws(() => wb.setThemeScriptFont({ major: true, script: 'Zzzz', typeface: 'X' }), TypeError);
    });
  });
});

describe('records return native booleans', () => {
  it('evaluateCfFormula returns true and false in both addressing forms', () => {
    withSheet((ws) => {
      ws.setNumber('A1', 5);
      assert.strictEqual(ws.evaluateCfFormula('A1', 'A1', 'A1>1'), true);
      assert.strictEqual(ws.evaluateCfFormula('A1', 'A1', 'A1>9'), false);
      assert.strictEqual(ws.evaluateCfFormula(0, 0, 0, 0, 'A1>1'), true);
      assert.strictEqual(ws.evaluateCfFormula(0, 0, 0, 0, 'A1>9'), false);
    });
  });

  it('a data bar match carries its bar with isNegative and gradient as booleans and border as a colour or null', () => {
    withSheet((ws) => {
      ws.setNumber('A1', -5);
      ws.setNumber('A2', 10);
      ws.addCfRule({
        type: f.CfRuleType.DataBar,
        ranges: 'A1:A2',
        dataBar: {
          min: { type: f.CfValueObjectType.Min },
          max: { type: f.CfValueObjectType.Max },
          fill: '#0070C0',
          border: '#000000',
        },
      });
      const cells = ws.evaluateCfRange('A1:A2');
      const bars = cells.map((c) => c.matches.find((m) => m.kind === f.CfMatchKind.DataBar));
      assert.equal(bars.length, 2);
      for (const match of bars) {
        for (const key of ['isNegative', 'gradient']) assert.equal(typeof match.bar[key], 'boolean', key);
        for (const key of ['barIsNegative', 'barBorderEngaged', 'barGradient'])
          assert.equal(match[key], undefined, key);
      }
      assert.equal(bars[0].bar.isNegative, true);
      assert.equal(bars[1].bar.isNegative, false);
      assert.equal(bars[1].bar.border.argb >>> 0, 0xff000000);
    });
  });

  it('a differential-format match reports its dxfIndex', () => {
    withSheet((ws, wb) => {
      ws.setNumber('A1', 5);
      const dxf = wb.ooxml.addDxf({ numberFormat: '0.00' });
      ws.addCfRule({
        type: f.CfRuleType.CellIs,
        ranges: 'A1',
        operator: f.ComparisonOperator.GreaterThan,
        formula1: '1',
        dxfIndex: dxf,
      });
      const match = ws.evaluateCfRange('A1')[0].matches[0];
      assert.strictEqual(match.kind, f.CfMatchKind.DifferentialFormat);
      assert.strictEqual(match.dxfIndex, dxf);
      assert.equal(match.dxfIndexEngaged, undefined);
    });
  });
});

describe('date inputs accept a Date wherever a civil time goes', () => {
  const local = new Date(2024, 0, 31, 9, 30, 15, 750);
  const civil = { year: 2024, month: 1, day: 31, hour: 9, minute: 30, second: 15 };

  it('setPinnedNow reads a Date through its local fields and drops milliseconds', () => {
    withSheet((_ws, wb) => {
      wb.setPinnedNow(local);
      assert.deepEqual({ ...wb.getPinnedNow() }, civil);
      wb.setPinnedNow(null);
      wb.setPinnedNow(civil);
      assert.deepEqual({ ...wb.getPinnedNow() }, civil);
    });
  });

  it('an Invalid Date is a RangeError in every date slot', () => {
    withSheet((ws, wb) => {
      const bad = new Date(Number.NaN);
      assert.throws(() => wb.setPinnedNow(bad), RangeError);
      assert.throws(() => ws.evaluateCfRange('A1', { today: bad }), RangeError);
      const { pivot, date, cache } = buildPivot(wb);
      assert.throws(() => pivot.setFieldDateGroup(date, { periods: f.PivotDatePeriod.Days, start: bad }), RangeError);
      assert.throws(() => cache.addFieldSharedItem(date, bad), RangeError);
      wb.addPerson({ id: PERSON, displayName: 'A', userId: 'a', providerId: 'None' });
      assert.throws(
        () =>
          ws.addThreadedComment({
            id: '{00000000-0000-4000-8000-000000000001}',
            text: 'x',
            cell: { row: 0, col: 0 },
            personId: PERSON,
            created: bad,
          }),
        RangeError,
      );
    });
  });

  it('a threaded comment takes a Date or a civil time for created and reads a civil time', () => {
    withSheet((ws, wb) => {
      wb.addPerson({ id: PERSON, displayName: 'A', userId: 'a', providerId: 'None' });
      const a = '{00000000-0000-4000-8000-000000000001}';
      const b = '{00000000-0000-4000-8000-000000000002}';
      ws.addThreadedComment({ id: a, text: 'one', cell: { row: 0, col: 0 }, personId: PERSON, created: local });
      ws.addThreadedComment({ id: b, text: 'two', cell: { row: 1, col: 0 }, personId: PERSON, created: civil });
      assert.deepEqual({ ...ws.getThreadedComment(a).created }, civil);
      assert.deepEqual({ ...ws.getThreadedComment(b).created }, civil);
    });
  });

  it('evaluateCfRange pins its clock from a Date or the equal civil time alike', () => {
    withSheet((ws) => {
      ws.setNumber('A1', 43845);
      ws.addCfRule({ type: f.CfRuleType.TimePeriod, ranges: 'A1', timePeriod: f.DatePeriod.Today });
      const clock = { year: 2020, month: 1, day: 15, hour: 12, minute: 0, second: 0 };
      assert.equal(ws.evaluateCfRange('A1', { today: new Date(2020, 0, 15, 12, 0, 0) }).length, 1);
      assert.equal(ws.evaluateCfRange('A1', { today: clock }).length, 1);
      assert.equal(ws.evaluateCfRange('A1', { today: new Date(2020, 0, 16, 12, 0, 0) }).length, 0);
    });
  });

  it('a pivot cache stores a Date as a date item and a number as a number', () => {
    withSheet((_ws, wb) => {
      const { cache, date } = buildPivot(wb);
      cache.addFieldSharedItem(date, new Date(2024, 0, 31));
      cache.addFieldSharedItem(date, 45322);
      const items = cache.listFieldSharedItems(date);
      assert.deepEqual(
        items.slice(-2).map((i) => [i.value, i.isDate]),
        [
          [45322, true],
          [45322, false],
        ],
      );
      const record = cache.addRecord();
      cache.setRecordValue(record, date, new Date(2024, 0, 31));
    });
  });

  it('a pivot date group takes Date bounds equal to the same civil bounds', () => {
    const group = (bounds) =>
      withSheet((_ws, wb) => {
        const { pivot, date } = buildPivot(wb);
        pivot.setFieldDateGroup(date, { periods: f.PivotDatePeriod.Months, ...bounds });
        return JSON.parse(JSON.stringify(pivot.listFields()[date]));
      });
    const viaDate = group({ start: new Date(2024, 0, 1), end: new Date(2025, 11, 31) });
    const viaCivil = group({
      start: { year: 2024, month: 1, day: 1, hour: 0, minute: 0, second: 0 },
      end: { year: 2025, month: 12, day: 31, hour: 0, minute: 0, second: 0 },
    });
    assert.deepEqual(viaDate, viaCivil);
  });

  it('setValue reads a Date through the same local wall clock', () => {
    withSheet((ws) => {
      ws.setValue('A1', new Date(2024, 0, 31, 12, 0, 0));
      assert.equal(ws.getValue('A1'), 45322.5);
    });
  });
});

describe('formula reads carry the leading equals sign and inputs take it either way', () => {
  it('a cell formula reads with =', () => {
    withSheet((ws) => {
      ws.setFormula('A1', '=1+2');
      ws.setFormula('A2', '3+4');
      assert.equal(ws.getFormula('A1'), '=1+2');
      assert.equal(ws.getFormula('A2'), '=3+4');
    });
  });

  it('conditional format and validation formulas read with = whichever spelling was stored', () => {
    withSheet((ws) => {
      ws.addCfRule({ type: f.CfRuleType.Expression, ranges: 'A1', formula1: 'A1>1' });
      ws.addCfRule({ type: f.CfRuleType.Expression, ranges: 'A2', formula1: '=A2>1' });
      const [a, b] = ws.listCfRules();
      assert.equal(a.formula1, '=A1>1');
      assert.equal(b.formula1, '=A2>1');
      ws.addValidation({ type: f.ValidationType.Custom, ranges: 'B1', formula1: 'B1>1' });
      ws.addValidation({ type: f.ValidationType.Custom, ranges: 'B2', formula1: '=B2>1' });
      const [c, d] = ws.listValidations();
      assert.equal(c.formula1, '=B1>1');
      assert.equal(d.formula1, '=B2>1');
    });
  });

  it('a formula-typed threshold reads with = and stores the same with or without it', () => {
    withSheet((ws) => {
      fillColumns(ws);
      const scale = (value) => ({
        type: f.CfRuleType.ColorScale,
        ranges: 'B2:B4',
        colorScale: {
          thresholds: [{ type: f.CfValueObjectType.Min }, { type: f.CfValueObjectType.Formula, value }],
          colors: ['#FFFFFF', '#FF0000'],
        },
      });
      ws.addCfRule(scale('B2'));
      ws.addCfRule(scale('=B2'));
      const [a, b] = ws.listCfRules();
      assert.equal(a.colorScale.thresholds[1].value, '=B2');
      assert.equal(b.colorScale.thresholds[1].value, '=B2');
    });
  });

  it('sparkline sources and the date axis range read as the same record from A1 text or a SheetRect', () => {
    withSheet((ws) => {
      fillColumns(ws);
      const source = { sheet: 0, rect: { firstRow: 1, firstCol: 0, lastRow: 3, lastCol: 0 } };
      const dates = { sheet: 0, rect: { firstRow: 1, firstCol: 1, lastRow: 3, lastCol: 1 } };
      ws.addSparklineGroup({
        dateAxis: true,
        dateAxisRange: 'Sheet1!B2:B4',
        sparklines: [{ source: 'Sheet1!A2:A4', location: 'F1' }],
      });
      ws.addSparklineGroup({
        dateAxis: true,
        dateAxisRange: dates,
        sparklines: [{ source, location: { row: 1, col: 5 } }],
      });
      const [a, b] = ws.listSparklineGroups();
      for (const group of [a, b]) {
        assert.deepEqual(JSON.parse(JSON.stringify(group.sparklines[0].source)), source);
        assert.deepEqual(JSON.parse(JSON.stringify(group.dateAxisRange)), dates);
      }
    });
  });

  it('a custom totals-row formula reads with = and stores the same either way', () => {
    withSheet((ws) => {
      fillColumns(ws);
      const table = ws.addTable({ range: 'A1:B4', name: 'Sales', totalsRow: true });
      table.setColumnTotals(1, {
        totalsRowFunction: f.TableTotalsRowFunction.Custom,
        totalsRowFormula: 'SUM(Sales[Amount])',
      });
      const bare = table.getProperties().columns[1].totalsRowFormula;
      table.setColumnTotals(1, {
        totalsRowFunction: f.TableTotalsRowFunction.Custom,
        totalsRowFormula: '=SUM(Sales[Amount])',
      });
      const prefixed = table.getProperties().columns[1].totalsRowFormula;
      assert.equal(bare, '=SUM([Amount])');
      assert.equal(prefixed, bare);
    });
  });
});

describe('error and value inputs accept the unions the setters accept', () => {
  it('setError takes an ExcelError read from another cell and still takes an ErrorCode', () => {
    withSheet((ws, wb) => {
      ws.setFormula('A1', '=1/0');
      wb.recalc();
      const error = ws.getValue('A1');
      assert.equal(error.code, f.ErrorCode.Div0);
      ws.setError('B1', error);
      ws.setError(1, 0, error);
      ws.setError('B3', f.ErrorCode.NA);
      assert.equal(ws.getValue('B1').code, f.ErrorCode.Div0);
      assert.equal(ws.getValue('A2').code, f.ErrorCode.Div0);
      assert.equal(ws.getValue('B3').code, f.ErrorCode.NA);
    });
  });

  it('validateValue accepts a native value, a Date and a tagged Value like setValue', () => {
    withSheet((ws) => {
      ws.addValidation({ type: f.ValidationType.List, ranges: 'B1', formula1: '"Open,Closed"' });
      assert.equal(ws.validateValue('B1', 'Open').valid, true);
      assert.equal(ws.validateValue('B1', 'Pending').valid, false);
      assert.equal(ws.validateValue(0, 1, 'Closed').valid, true);
      assert.equal(ws.validateValue('B1', { kind: f.ValueKind.Text, text: 'Open' }).valid, true);
      ws.addValidation({
        type: f.ValidationType.Whole,
        ranges: 'C1',
        operator: f.ComparisonOperator.GreaterThan,
        formula1: '10',
      });
      assert.equal(ws.validateValue('C1', 42).valid, true);
      assert.equal(ws.validateValue('C1', 3).valid, false);
      assert.equal(ws.validateValue('C1', new Date(2024, 0, 31)).valid, true);
    });
  });
});

describe('pivot filters carry value and valueHigh as number or string', () => {
  function withPivot(body) {
    withSheet((_ws, wb) => body(buildPivot(wb).pivot));
  }
  const region = { fieldIndex: 'Region' };

  it('a value filter reads back its numeric operand and no high bound', () => {
    withPivot((pivot) => {
      pivot.addFilter({ ...region, kind: f.PivotFilterKind.ValueGreaterThan, value: 20 });
      const filter = pivot.listFilters()[0];
      assert.equal(filter.value, 20);
      assert.equal(filter.valueHigh, null);
    });
  });

  it('a between filter reads back both numeric bounds', () => {
    withPivot((pivot) => {
      pivot.addFilter({ ...region, kind: f.PivotFilterKind.ValueBetween, value: 10, valueHigh: 20.5 });
      const filter = pivot.listFilters()[0];
      assert.equal(filter.value, 10);
      assert.equal(filter.valueHigh, 20.5);
    });
  });

  it('a top-count filter takes an integer count', () => {
    withPivot((pivot) => {
      pivot.addFilter({ ...region, kind: f.PivotFilterKind.ValueTopCount, value: 3 });
      assert.equal(pivot.listFilters()[0].value, 3);
    });
  });

  it('label filters read back string operands, the upper bound in valueHigh', () => {
    withPivot((pivot) => {
      pivot.addFilter({ ...region, kind: f.PivotFilterKind.LabelContains, value: 'as' });
      pivot.addFilter({ ...region, kind: f.PivotFilterKind.LabelBetween, value: 'A', valueHigh: 'M' });
      const [contains, between] = pivot.listFilters();
      assert.equal(contains.value, 'as');
      assert.equal(contains.valueHigh, null);
      assert.equal(between.value, 'A');
      assert.equal(between.valueHigh, 'M');
    });
  });

  it('the operand-kind fields are gone from reads and rejected on input', () => {
    withPivot((pivot) => {
      pivot.addFilter({ ...region, kind: f.PivotFilterKind.ValueGreaterThan, value: 5 });
      const filter = pivot.listFilters()[0];
      for (const key of [
        'valueKind',
        'valueInt',
        'valueDouble',
        'valueText',
        'valueHighKind',
        'valueHighInt',
        'valueHighDouble',
        'valueHighText',
      ]) {
        assert.equal(key in filter, false, key);
      }
      assert.throws(
        () => pivot.addFilter({ ...region, kind: f.PivotFilterKind.ValueGreaterThan, valueKind: 1, valueDouble: 5 }),
        TypeError,
      );
    });
  });

  it('an operand of the wrong type for the filter is rejected', () => {
    withPivot((pivot) => {
      assert.throws(() => pivot.addFilter({ ...region, kind: f.PivotFilterKind.ValueGreaterThan, value: 'big' }));
      assert.throws(() => pivot.addFilter({ ...region, kind: f.PivotFilterKind.LabelContains, value: 5 }));
    });
  });
});

describe('native values: toNative and pivot cells', () => {
  it('toNative maps each Value kind to its native JS value', () => {
    assert.equal(f.toNative({ kind: f.ValueKind.Number, number: 3 }), 3);
    assert.equal(f.toNative({ kind: f.ValueKind.Bool, boolean: false }), false);
    assert.equal(f.toNative({ kind: f.ValueKind.Text, text: 'abc' }), 'abc');
    assert.equal(f.toNative({ kind: f.ValueKind.Blank }), null);
    const error = f.toNative({ kind: f.ValueKind.Error, errorCode: f.ErrorCode.Div0 });
    assert.equal(error.code, f.ErrorCode.Div0);
  });

  it('a pivot layout cell carries a native value and no isDate flag', () => {
    withSheet((_ws, wb) => {
      const { pivot } = buildPivot(wb);
      const cells = pivot.getLayout().cells;
      assert.ok(cells.length > 0);
      for (const cell of cells) {
        assert.equal('isDate' in cell, false);
        assert.ok(cell.value === null || ['number', 'string', 'boolean', 'object'].includes(typeof cell.value));
        assert.equal(cell.value?.kind, undefined, 'a native value has no Value kind tag');
      }
      assert.ok(cells.some((c) => c.value === 'East'));
      assert.ok(cells.some((c) => c.value === 45));
    });
  });

  it('a shared item reads a native value with isDate telling dates from numbers', () => {
    withSheet((_ws, wb) => {
      const { cache, region, date } = buildPivot(wb);
      assert.equal(cache.listFieldSharedItems(region)[0].value, 'East');
      cache.addFieldSharedItem(date, new Date(2024, 0, 31));
      const item = cache.listFieldSharedItems(date).at(-1);
      assert.equal(item.value, 45322);
      assert.equal(item.isDate, true);
    });
  });
});
