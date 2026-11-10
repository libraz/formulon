// Table.update({ columns }): per-column totals settings and calculated-column formulas merged element by element, the single-cell write
// rules that keep a column's calculated mark and totals settings in step with its cells, and the save round trips.

import assert from 'node:assert/strict';
import test from 'node:test';

import { getModule } from './smoke_support.mjs';

async function withWorkbook(fn) {
  const f = await getModule();
  const wb = f.Workbook.createDefault();
  try {
    return await fn(wb, f);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

function rejectsWith(fn, code, label) {
  assert.throws(
    fn,
    (e) => {
      assert.equal(e.name, 'FormulonError', `${label}: expected a FormulonError, got ${e}`);
      assert.equal(e.code, code, `${label}: code ${e.code}`);
      return true;
    },
    label,
  );
}

// Headers a, b, c on row 2, data on rows 3..5 (a = 1..3, b = 10..30, c empty); table Tbl over A2:C5, or A2:C6 with a
// totals row. The first data row is Excel row 3, so `=A3*10` is the first-row spelling.
function calcTable(wb, totalsRow = false) {
  const ws = wb.sheet(0);
  ws.setValues('A2', [
    ['a', 'b'],
    [1, 10],
    [2, 20],
    [3, 30],
  ]);
  ws.setText('C2', 'c');
  const table = ws.addTable({ range: totalsRow ? 'A2:C6' : 'A2:C5', name: 'Tbl', totalsRow });
  return { ws, table };
}

const column = (table, index) => table.getProperties().columns[index];
const rectOf = (table) => ({ ...table.getProperties().range.rect });
// Element `index` of a `columns` array holds `element`; the columns before it are left alone with `{}`.
const at = (index, element) => [...Array(index).fill({}), element];
const calc = (index, calculatedFormula) => at(index, { calculatedFormula });

test('a column patch sets the calculated formula, fills the column and overwrites a constant', async () => {
  await withWorkbook((wb) => {
    const { ws, table } = calcTable(wb);
    assert.equal(column(table, 2).calculatedFormula, null);
    ws.setNumber('C4', 99);
    table.update({ columns: calc(2, '=A3*10') });
    assert.equal(column(table, 2).calculatedFormula, '=A3*10');
    assert.deepEqual(
      ['C3', 'C4', 'C5'].map((a1) => ws.getFormula(a1)),
      ['=A3*10', '=A4*10', '=A5*10'],
    );
    wb.recalc();
    assert.deepEqual(
      ['C3', 'C4', 'C5'].map((a1) => ws.getValue(a1)),
      [10, 20, 30],
    );
    assert.equal(column(table, 0).calculatedFormula, null);
    assert.equal(column(table, 1).calculatedFormula, null);
    assert.equal(ws.getValue('C2'), 'c', 'the header cell is not part of the fill');
  });
});

test('a calculated formula reads back in the first-row spelling whichever form it was given in', async () => {
  const cases = [
    ['=A3*10', '=A3*10'],
    ['A3*10', '=A3*10'],
    ['=[@a]*10', '=[@a]*10'],
    ['=Tbl[[#This Row],[a]]*10', '=[@a]*10'],
    ['=$A$3*10', '=$A$3*10'],
    ['=A$3*10+$B3', '=A$3*10+$B3'],
  ];
  for (const [input, read] of cases) {
    await withWorkbook((wb) => {
      const { table } = calcTable(wb);
      table.update({ columns: calc(2, input) });
      assert.equal(column(table, 2).calculatedFormula, read, input);
    });
  }
  await withWorkbook((wb) => {
    const { ws, table } = calcTable(wb);
    table.update({ columns: calc(2, '=$A$3*10') });
    wb.recalc();
    assert.equal(ws.getValue('C5'), 10, 'an absolute reference stays on A3 in every row');
    table.update({ columns: calc(2, '=[@a]*10') });
    wb.recalc();
    assert.deepEqual([ws.getValue('C3'), ws.getValue('C5')], [10, 30]);
  });
});

test('calculatedFormula null removes the mark and leaves the cells as they are', async () => {
  await withWorkbook((wb) => {
    const { ws, table } = calcTable(wb);
    table.update({ columns: calc(2, '=A3*10') });
    table.update({ columns: calc(2, null) });
    assert.equal(column(table, 2).calculatedFormula, null);
    assert.deepEqual(
      ['C3', 'C4', 'C5'].map((a1) => ws.getFormula(a1)),
      ['=A3*10', '=A4*10', '=A5*10'],
    );
    wb.recalc();
    assert.equal(ws.getValue('C5'), 30);
  });
});

test('changing the calculated formula refills the column, and re-applying the same formula keeps constants', async () => {
  await withWorkbook((wb) => {
    const { ws, table } = calcTable(wb);
    table.update({ columns: calc(2, '=A3*10') });
    ws.setNumber('C4', 99);
    table.update({ columns: calc(2, '=A3*10') });
    assert.equal(ws.getValue('C4'), 99, 'the same formula is no change, so the constant survives');
    table.update({ columns: calc(2, '=A$3*10') });
    assert.equal(ws.getFormula('C4'), '=A$3*10', 'a different formula fills every data cell');
    wb.recalc();
    assert.deepEqual(
      ['C3', 'C4', 'C5'].map((a1) => ws.getValue(a1)),
      [10, 10, 10],
    );
    table.update({ columns: calc(2, null) });
    ws.setNumber('C5', 7);
    table.update({ columns: calc(2, '=$A$3*10') });
    assert.equal(ws.getFormula('C5'), '=$A$3*10', 'setting a formula on an unmarked column fills it');
  });
});

test('the other keys of an element leave the mark, and an empty element is a no-op', async () => {
  await withWorkbook((wb, f) => {
    const { ws, table } = calcTable(wb, true);
    table.update({ columns: calc(2, '=A3*10') });
    table.update({ columns: [{}, {}, { totalsRowFunction: f.TableTotalsRowFunction.Max }] });
    assert.equal(column(table, 2).calculatedFormula, '=A3*10');
    assert.equal(column(table, 2).totalsRowFunction, f.TableTotalsRowFunction.Max);
    table.update({ columns: [{}, {}, {}] });
    assert.equal(column(table, 2).calculatedFormula, '=A3*10');
    assert.equal(column(table, 2).totalsRowFunction, f.TableTotalsRowFunction.Max);
    assert.equal(ws.getFormula('C4'), '=A4*10');
  });
});

test('an empty calculatedFormula is InvalidArgument and keeps the mark', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const { ws, table } = calcTable(wb);
    table.update({ columns: calc(2, '=A3*10') });
    rejectsWith(() => table.update({ columns: calc(2, '') }), StatusCode.InvalidArgument, 'empty formula');
    assert.equal(column(table, 2).calculatedFormula, '=A3*10');
    assert.equal(ws.getFormula('C4'), '=A4*10');
  });
});

test('unusable calculated formulas are refused and change nothing', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const { ws, table } = calcTable(wb);
    ws.setNumber('C4', 99);
    assert.throws(
      () => table.update({ columns: calc(2, '=SUM((') }),
      (e) => e.name === 'FormulonError',
    );
    for (const formula of ['=SEQUENCE(3)', '=FILTER(A3:A5,A3:A5>1)', '=A1048576+1']) {
      rejectsWith(() => table.update({ columns: calc(2, formula) }), StatusCode.InvalidArgument, formula);
    }
    assert.equal(column(table, 2).calculatedFormula, null);
    assert.equal(ws.getValue('C4'), 99);
    assert.equal(ws.getFormula('C3'), null);
  });
});

test('an array longer than the column count is refused and no element applies', async () => {
  await withWorkbook((wb, f) => {
    const { ws, table } = calcTable(wb, true);
    rejectsWith(
      () =>
        table.update({
          columns: [{}, { totalsRowFunction: f.TableTotalsRowFunction.Sum }, { calculatedFormula: '=A3*10' }, {}],
        }),
      f.StatusCode.IndexOutOfRange,
      'four elements for three columns',
    );
    assert.equal(column(table, 2).calculatedFormula, null);
    assert.equal(column(table, 1).totalsRowFunction, f.TableTotalsRowFunction.None);
    assert.equal(ws.getFormula('C3'), null);
    assert.equal(ws.getFormula('B6'), null);
  });
});

test('a shorter array patches the leading columns only, and an empty array changes nothing', async () => {
  await withWorkbook((wb, f) => {
    const { ws, table } = calcTable(wb, true);
    table.update({ columns: [{ totalsRowLabel: 'Total' }] });
    assert.equal(column(table, 0).totalsRowLabel, 'Total');
    assert.equal(column(table, 1).totalsRowLabel, null);
    table.update({
      columns: [{}, { totalsRowFunction: f.TableTotalsRowFunction.Sum }, { calculatedFormula: '=A3*10' }],
    });
    table.update({ columns: [] });
    assert.equal(column(table, 2).calculatedFormula, '=A3*10');
    assert.equal(column(table, 1).totalsRowFunction, f.TableTotalsRowFunction.Sum);
    assert.equal(ws.getFormula('B6'), '=SUBTOTAL(109,[b])');
  });
});

test('the columns of a patch are checked before the range or the totals flag changes', async () => {
  await withWorkbook((wb, { StatusCode }) => {
    const { table } = calcTable(wb);
    const before = rectOf(table);
    rejectsWith(
      () => table.update({ totalsRow: true, columns: [{}, {}, {}, {}] }),
      StatusCode.IndexOutOfRange,
      'too many elements with a flag',
    );
    assert.equal(table.getProperties().totalsRow, false);
    assert.deepEqual(rectOf(table), before);
    for (const formula of ['=SUM((', '=SEQUENCE(3)', '=A1048576+1']) {
      assert.throws(
        () => table.update({ totalsRow: true, columns: calc(2, formula) }),
        (e) => e.name === 'FormulonError',
      );
      assert.equal(table.getProperties().totalsRow, false, formula);
      assert.deepEqual(rectOf(table), before, formula);
    }
  });
});

test('several columns apply in one call and the totals row is not part of the fill', async () => {
  await withWorkbook((wb, f) => {
    const { ws, table } = calcTable(wb, true);
    table.update({
      columns: [
        { totalsRowLabel: 'Total' },
        { totalsRowFunction: f.TableTotalsRowFunction.Sum },
        { calculatedFormula: '=A3*10' },
      ],
    });
    assert.equal(column(table, 0).totalsRowLabel, 'Total');
    assert.equal(column(table, 1).totalsRowFunction, f.TableTotalsRowFunction.Sum);
    assert.equal(column(table, 2).calculatedFormula, '=A3*10');
    assert.equal(ws.getFormula('B6'), '=SUBTOTAL(109,[b])');
    assert.equal(ws.getFormula('C5'), '=A5*10');
    // The totals row holds the default total, not the calculated formula.
    assert.ok(!(ws.getFormula('C6') ?? '').includes('*10'), ws.getFormula('C6'));
  });
});

test('null clears each field of an element and a non-custom function drops the formula', async () => {
  await withWorkbook((wb, f) => {
    const { ws, table } = calcTable(wb, true);
    const T = f.TableTotalsRowFunction;
    table.update({ columns: [{}, { totalsRowFunction: T.Sum, totalsRowLabel: 'Sum' }] });
    assert.equal(ws.getFormula('B6'), '=SUBTOTAL(109,[b])');
    table.update({ columns: [{}, { totalsRowFunction: null }] });
    assert.equal(column(table, 1).totalsRowFunction, T.None);
    assert.equal(ws.getValue('B6'), 'Sum', 'no function leaves the label, which is shown as text');
    table.update({ columns: [{}, { totalsRowLabel: null }] });
    assert.equal(column(table, 1).totalsRowLabel, null);
    assert.equal(ws.getValue('B6'), null);

    table.update({ columns: [{}, { totalsRowFunction: T.Custom, totalsRowFormula: '=SUM(Tbl[b])*2' }] });
    assert.equal(column(table, 1).totalsRowFormula, '=SUM([b])*2');
    table.update({ columns: [{}, { totalsRowFunction: T.Max }] });
    assert.equal(column(table, 1).totalsRowFormula, null, 'a non-custom function reads its formula as null');
    table.update({ columns: [{}, { totalsRowFunction: T.Custom, totalsRowFormula: '=5' }] });
    assert.throws(() => table.update({ columns: [{}, { totalsRowFormula: null }] }), /required for a custom function/);
    assert.equal(column(table, 1).totalsRowFormula, '=5', 'a custom function keeps its formula');
  });
});

test('a custom function without a formula and an empty string are InvalidArgument', async () => {
  await withWorkbook((wb, { StatusCode, TableTotalsRowFunction }) => {
    const { table } = calcTable(wb, true);
    rejectsWith(
      () => table.update({ columns: [{}, { totalsRowFunction: TableTotalsRowFunction.Custom }] }),
      StatusCode.InvalidArgument,
      'custom without formula',
    );
    rejectsWith(() => table.update({ columns: [{ totalsRowLabel: '' }] }), StatusCode.InvalidArgument, 'empty label');
    rejectsWith(
      () => table.update({ columns: [{}, { totalsRowFormula: '' }] }),
      StatusCode.InvalidArgument,
      'empty formula',
    );
    assert.equal(column(table, 1).totalsRowFunction, TableTotalsRowFunction.None);
  });
});

test('null totals fields clear a shown cell, and hidden totals are metadata only', async () => {
  await withWorkbook((wb, f) => {
    const { ws, table } = calcTable(wb, true);
    table.update({ columns: [{}, { totalsRowFunction: f.TableTotalsRowFunction.Sum }] });
    assert.equal(ws.getFormula('B6'), '=SUBTOTAL(109,[b])');
    table.update({ columns: [{}, { totalsRowFunction: null }] });
    assert.equal(column(table, 1).totalsRowFunction, f.TableTotalsRowFunction.None);
    assert.equal(ws.getValue('B6'), null);
  });
  await withWorkbook((wb, f) => {
    const { ws, table } = calcTable(wb);
    table.update({ columns: [{}, { totalsRowFunction: f.TableTotalsRowFunction.Max }] });
    assert.equal(column(table, 1).totalsRowFunction, f.TableTotalsRowFunction.Max);
    assert.equal(ws.getValue('B6'), null, 'no totals row, so no cell');
    assert.equal(table.getProperties().range.rect.lastRow, 4);
  });
});

test('a table without data rows takes the calculated mark only', async () => {
  await withWorkbook((wb) => {
    const ws = wb.sheet(0);
    ws.setValues('A2', [['a', 'b', 'c']]);
    const table = ws.addTable({ range: 'A2:C3', name: 'Tbl', totalsRow: true });
    table.update({ columns: calc(2, '=A3*10') });
    assert.ok(column(table, 2).calculatedFormula.includes('*10'), column(table, 2).calculatedFormula);
    assert.ok(!(ws.getFormula('C3') ?? '').includes('*10'), 'the totals row is not a data row');
    table.update({ columns: calc(2, null) });
    assert.equal(column(table, 2).calculatedFormula, null);
  });
});

test('feeding getProperties() back to update changes nothing, constants in a calculated column included', async () => {
  await withWorkbook((wb, f) => {
    const { ws, table } = calcTable(wb, true);
    table.update({
      columns: [{}, { totalsRowFunction: f.TableTotalsRowFunction.Sum }, { calculatedFormula: '=A3*10' }],
    });
    ws.setNumber('C4', 99);
    const snapshot = () => {
      const props = table.getProperties();
      return {
        rect: { ...props.range.rect },
        headerRow: props.headerRow,
        totalsRow: props.totalsRow,
        columns: props.columns.map((c) => [c.name, c.totalsRowFunction, c.totalsRowLabel, c.calculatedFormula]),
        cells: ['C3', 'C4', 'C5', 'B6'].map((a1) => [ws.getFormula(a1), ws.getValue(a1)]),
      };
    };
    const before = snapshot();
    table.update(table.getProperties());
    assert.deepEqual(snapshot(), before);
    assert.equal(ws.getValue('C4'), 99);
  });
});

test('a malformed column patch element is a TypeError and changes nothing', async () => {
  await withWorkbook((wb, f) => {
    const { ws, table } = calcTable(wb);
    const bad = {
      'an unknown key': [{}, {}, { calculatedFormula: '=A3*10', bogus: 1 }],
      'the old totals nesting': [{}, { totals: { totalsRowFunction: f.TableTotalsRowFunction.Sum } }],
      'the old column key': [{ column: 2, calculatedFormula: '=A3*10' }],
      'a non-record element': [{}, 'sum'],
    };
    for (const [what, columns] of Object.entries(bad)) {
      assert.throws(() => table.update({ columns }), TypeError, what);
    }
    assert.equal(column(table, 2).calculatedFormula, null);
    assert.equal(column(table, 1).totalsRowFunction, f.TableTotalsRowFunction.None);
    assert.equal(ws.getFormula('C3'), null);
  });
});

test('a single-cell write keeps the calculated column and touches only that cell', async () => {
  const writes = {
    'a constant': (ws) => ws.setNumber('C4', 7),
    'a text': (ws) => ws.setText('C4', 'x'),
    'a cleared cell': (ws) => ws.setBlank('C4'),
    'another formula': (ws) => ws.setFormula('C4', '=B4+1'),
  };
  for (const [what, write] of Object.entries(writes)) {
    await withWorkbook((wb) => {
      const { ws, table } = calcTable(wb);
      table.update({ columns: calc(2, '=A3*10') });
      write(ws);
      assert.equal(column(table, 2).calculatedFormula, '=A3*10', what);
      assert.equal(ws.getFormula('C3'), '=A3*10', what);
      assert.equal(ws.getFormula('C5'), '=A5*10', what);
    });
  }
});

test('blanking the last non-blank cell of a calculated column removes the mark', async () => {
  await withWorkbook((wb) => {
    const { ws, table } = calcTable(wb);
    table.update({ columns: calc(2, '=A3*10') });
    ws.setBlank('C3');
    ws.setBlank('C5');
    assert.equal(column(table, 2).calculatedFormula, '=A3*10', 'one non-blank cell is left');
    ws.setBlank('C4');
    assert.equal(column(table, 2).calculatedFormula, null);
    ws.setBlank('A3');
    assert.equal(column(table, 0).calculatedFormula, null);
  });
});

test('blanking a totals cell does not drop the calculated column', async () => {
  await withWorkbook((wb) => {
    const { ws, table } = calcTable(wb, true);
    table.update({ columns: calc(2, '=A3*10') });
    ws.setText('C6', 'x');
    ws.setBlank('C6');
    assert.equal(column(table, 2).calculatedFormula, '=A3*10');
  });
});

test('a totals-cell write sets the column totals settings', async () => {
  await withWorkbook((wb, f) => {
    const { ws, table } = calcTable(wb, true);
    const none = f.TableTotalsRowFunction.None;
    const custom = f.TableTotalsRowFunction.Custom;

    ws.setNumber('B6', 5);
    assert.deepEqual([column(table, 1).totalsRowFunction, column(table, 1).totalsRowLabel], [none, '5']);
    assert.equal(ws.getValue('B6'), '5', 'a numeric label is stored as text');
    ws.setNumber('B6', 5.5);
    assert.equal(column(table, 1).totalsRowLabel, '5.5');

    ws.setText('A6', 'Total');
    assert.deepEqual([column(table, 0).totalsRowFunction, column(table, 0).totalsRowLabel], [none, 'Total']);

    ws.setFormula('B6', '=SUBTOTAL(109,Tbl[b])');
    assert.equal(column(table, 1).totalsRowFunction, f.TableTotalsRowFunction.Sum);
    assert.equal(column(table, 1).totalsRowFormula, null);

    ws.setFormula('B6', '=5');
    assert.deepEqual([column(table, 1).totalsRowFunction, column(table, 1).totalsRowFormula], [custom, '=5']);

    ws.setFormula('B6', '=SUBTOTAL(9,Tbl[b])');
    assert.deepEqual(
      [column(table, 1).totalsRowFunction, column(table, 1).totalsRowFormula],
      [custom, '=SUBTOTAL(9,[b])'],
    );

    ws.setFormula('B6', '=SUM(Tbl[b])*2');
    assert.deepEqual([column(table, 1).totalsRowFunction, column(table, 1).totalsRowFormula], [custom, '=SUM([b])*2']);

    ws.setBool('B6', true);
    assert.deepEqual([column(table, 1).totalsRowFunction, column(table, 1).totalsRowLabel], [none, 'TRUE']);

    ws.setBlank('B6');
    assert.deepEqual(
      [column(table, 1).totalsRowFunction, column(table, 1).totalsRowLabel, column(table, 1).totalsRowFormula],
      [none, null, null],
    );

    // A data cell is not a totals cell.
    table.update({ columns: [{}, { totalsRowFunction: f.TableTotalsRowFunction.Sum }] });
    ws.setText('B5', 'x');
    assert.equal(column(table, 1).totalsRowFunction, f.TableTotalsRowFunction.Sum);
  });
});

// Pairwise rows over write kind x column x prior state.
const TOTALS_STEPS = [
  ['blank', 2, 'label'],
  ['constantFormula', 2, 'function'],
  ['subtotalOwn', 1, 'none'],
  ['customFormula', 0, 'label'],
  ['number', 1, 'custom'],
  ['subtotalOtherColumn', 2, 'none'],
  ['constantFormula', 0, 'custom'],
  ['customFormula', 1, 'function'],
  ['subtotalOwn', 0, 'function'],
  ['bool', 1, 'label'],
  ['constantFormula', 1, 'none'],
  ['number', 0, 'none'],
  ['number', 2, 'label'],
  ['blank', 0, 'none'],
  ['bool', 0, 'none'],
  ['subtotalOtherColumn', 1, 'label'],
  ['constantFormula', 2, 'label'],
  ['customFormula', 2, 'none'],
  ['text', 2, 'none'],
  ['subtotalOwn', 2, 'custom'],
  ['bool', 2, 'function'],
  ['blank', 1, 'function'],
  ['subtotalOtherColumn', 0, 'custom'],
  ['text', 0, 'custom'],
  ['customFormula', 0, 'custom'],
  ['text', 1, 'none'],
  ['number', 2, 'function'],
  ['text', 1, 'function'],
  ['blank', 0, 'custom'],
  ['bool', 2, 'custom'],
  ['subtotalOtherColumn', 0, 'function'],
  ['subtotalOwn', 0, 'label'],
  ['text', 0, 'label'],
];
const NAMES = ['a', 'b', 'c'];
const COLS = ['A', 'B', 'C'];
const BUILT_INS = {
  109: 'Sum',
  105: 'Min',
  104: 'Max',
  101: 'Average',
  103: 'Count',
  102: 'CountNumbers',
  107: 'StdDev',
  110: 'Var',
};

function writeTotals(ws, col, kind) {
  const at = `${COLS[col]}6`;
  const own = `Tbl[${NAMES[col]}]`;
  const other = `Tbl[${NAMES[(col + 1) % 3]}]`;
  switch (kind) {
    case 'number':
      return ws.setNumber(at, 5.5);
    case 'text':
      return ws.setText(at, 'Total');
    case 'bool':
      return ws.setBool(at, true);
    case 'subtotalOwn':
      return ws.setFormula(at, `=SUBTOTAL(104,${own})`);
    case 'subtotalOtherColumn':
      return ws.setFormula(at, `=SUBTOTAL(109,${other})`);
    case 'constantFormula':
      return ws.setFormula(at, '=5');
    case 'customFormula':
      return ws.setFormula(at, `=SUM(${own})*2`);
    case 'blank':
      return ws.setBlank(at);
    default:
      throw new Error(kind);
  }
}

const PRIOR_WRITE = { none: 'blank', label: 'text', function: 'subtotalOwn', custom: 'constantFormula' };
const EXPECTED_LABEL = { number: '5.5', text: 'Total', bool: 'TRUE' };

// The column's totals settings judged against what its totals cell holds now.
function expectTotalsAgree(ws, table, f, col, where) {
  const read = column(table, col);
  const formula = ws.getFormula(`${COLS[col]}6`);
  const value = ws.getValue(`${COLS[col]}6`);
  if (formula !== null) {
    let expected = f.TableTotalsRowFunction.Custom;
    for (const [code, name] of Object.entries(BUILT_INS)) {
      if (formula === `=SUBTOTAL(${code},[${NAMES[col]}])`) expected = f.TableTotalsRowFunction[name];
    }
    assert.equal(read.totalsRowFunction, expected, `${where}: ${formula}`);
    if (expected === f.TableTotalsRowFunction.Custom) assert.equal(read.totalsRowFormula, formula, where);
    else assert.equal(read.totalsRowFormula, null, where);
  } else if (value === null) {
    assert.equal(read.totalsRowFunction, f.TableTotalsRowFunction.None, where);
    assert.equal(read.totalsRowLabel, null, where);
    assert.equal(read.totalsRowFormula, null, where);
  } else {
    assert.equal(read.totalsRowFunction, f.TableTotalsRowFunction.None, where);
    assert.equal(typeof value, 'string', `${where}: a value in a totals cell is stored as its label text`);
    assert.equal(read.totalsRowLabel, value, where);
  }
}

test('the totals settings agree with the totals cell after every write', async () => {
  const f = await getModule();
  TOTALS_STEPS.forEach(([kind, col, prior], i) => {
    const wb = f.Workbook.createDefault();
    try {
      const { ws, table } = calcTable(wb, true);
      writeTotals(ws, col, PRIOR_WRITE[prior]);
      expectTotalsAgree(ws, table, f, col, `step ${i} prior`);
      writeTotals(ws, col, kind);
      expectTotalsAgree(ws, table, f, col, `step ${i} write`);
      if (kind in EXPECTED_LABEL) assert.equal(column(table, col).totalsRowLabel, EXPECTED_LABEL[kind], `step ${i}`);
    } finally {
      wb.dispose();
    }
  });

  // The same writes as one running sequence: every column stays consistent after every step.
  const wb = f.Workbook.createDefault();
  try {
    const { ws, table } = calcTable(wb, true);
    TOTALS_STEPS.forEach(([kind, col], i) => {
      writeTotals(ws, col, kind);
      for (let c = 0; c < 3; c++) expectTotalsAgree(ws, table, f, c, `sequence ${i} column ${c}`);
    });
  } finally {
    wb.dispose();
  }
});

test('a calculated column and its totals survive an xlsx and an xlsb round trip', async () => {
  const f = await getModule();
  for (const format of [f.WorkbookFormat.Xlsx, f.WorkbookFormat.Xlsb]) {
    const wb = f.Workbook.createDefault();
    let bytes;
    try {
      const { table } = calcTable(wb, true);
      table.update({
        columns: [
          { totalsRowLabel: 'Total' },
          { totalsRowFunction: f.TableTotalsRowFunction.Sum },
          { calculatedFormula: '=A3*10' },
        ],
      });
      bytes = wb.save(format);
    } finally {
      wb.dispose();
    }
    const reloaded = f.Workbook.load(bytes);
    try {
      const ws = reloaded.sheet(0);
      const table = ws.table(0);
      assert.equal(column(table, 2).calculatedFormula, '=A3*10', `format ${format}`);
      assert.equal(column(table, 0).totalsRowLabel, 'Total', `format ${format}`);
      assert.equal(column(table, 1).totalsRowFunction, f.TableTotalsRowFunction.Sum, `format ${format}`);
      assert.deepEqual(
        ['C3', 'C4', 'C5'].map((a1) => ws.getFormula(a1)),
        ['=A3*10', '=A4*10', '=A5*10'],
        `format ${format}`,
      );
      assert.equal(ws.getFormula('B6'), '=SUBTOTAL(109,[b])', `format ${format}`);
      reloaded.recalc();
      assert.equal(ws.getValue('C5'), 30, `format ${format}`);
      assert.equal(ws.getValue('B6'), 60, `format ${format}`);
      if (format === f.WorkbookFormat.Xlsx) {
        ws.setNumber('C4', 7);
        assert.equal(column(table, 2).calculatedFormula, '=A3*10', 'a one-cell write keeps a reloaded mark');
      }
    } finally {
      reloaded.dispose();
    }
  }
});
