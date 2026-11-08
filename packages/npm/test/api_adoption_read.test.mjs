// The read model of the facade: values computed on read under CalcMode, native and tagged value shapes,
// the used range, row iteration, dates, buffer forms of `load`, workbook iteration and tab colours.
// FORMULON_PACKAGE=native runs it against the staged @libraz/formulon-native; otherwise FORMULON_NPM_ENTRY
// selects the staged @libraz/formulon subpath ("." by default, or "./threads").

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(testDir, '..', '..', '..');

async function loadPackage() {
  if (process.env.FORMULON_PACKAGE === 'native') {
    return import(pathToFileURL(path.join(repoRoot, 'packages', 'npm-native', 'dist', 'index.mjs')).href);
  }
  const pkgRoot = path.resolve(testDir, '..');
  const pkg = JSON.parse(await readFile(path.join(pkgRoot, 'package.json'), 'utf8'));
  const target = pkg.exports?.[process.env.FORMULON_NPM_ENTRY ?? '.']?.import;
  if (!target) throw new Error('package.json "exports" has no import target for the selected entry');
  return import(pathToFileURL(path.resolve(pkgRoot, target)).href);
}

const mod = await loadPackage();
const formulon = await mod.createFormulon();
const { CalcMode, ErrorCode, ExcelError, StatusCode, ValueKind, Workbook } = formulon;

async function table(name) {
  return JSON.parse(await readFile(path.join(repoRoot, 'tools', 'api', name), 'utf8'));
}

function withSheet(body, create = () => Workbook.createDefault()) {
  const wb = create();
  try {
    return body(wb.sheet(0), wb);
  } finally {
    wb.dispose();
  }
}

/** A case table's native value: `{error: <ErrorCode member>}` is that cell error. */
function expected(native) {
  if (Array.isArray(native)) return native.map(expected);
  if (native !== null && typeof native === 'object' && 'error' in native)
    return new ExcelError(ErrorCode[native.error]);
  return native;
}

function rejectsWith(call, code) {
  assert.throws(call, (e) => e.name === 'FormulonError' && e.code === code, `expected StatusCode ${code}`);
}

test('a read computes pending cells under Auto and AutoNoTable', () => {
  for (const mode of [CalcMode.Auto, CalcMode.AutoNoTable]) {
    withSheet((ws, wb) => {
      wb.setCalcMode(mode);
      ws.setValue('A1', 2);
      ws.setFormula('A2', '=A1*10');
      assert.equal(ws.getValue('A2'), 20);
      assert.equal(wb.hasDirtyCells(), false);
      ws.setValue('A1', 3);
      assert.deepEqual(ws.getValues('A2'), [[30]]);
    });
  }
});

test('a read under Manual leaves pending cells until recalc', () => {
  withSheet((ws, wb) => {
    wb.setCalcMode(CalcMode.Manual);
    ws.setValue('A1', 2);
    ws.setFormula('A2', '=A1*10');
    assert.equal(ws.getValue('A2'), null);
    assert.equal(wb.hasDirtyCells(), true);
    wb.recalc();
    assert.equal(ws.getValue('A2'), 20);
  });
});

test('a loaded workbook recomputes on read, or keeps its cached values under Manual', () => {
  const bytes = withSheet((ws, wb) => {
    wb.setCalcMode(CalcMode.Manual);
    wb.setCalcOnSave(false);
    ws.setValue('A1', 2);
    ws.setFormula('A2', '=A1*10');
    wb.recalc();
    ws.setValue('A1', 3);
    return wb.save();
  });
  withSheet(
    (ws, wb) => {
      // The file was saved under Manual, which the workbook loads with.
      wb.setCalcMode(CalcMode.Auto);
      assert.equal(ws.getValue('A2'), 30);
    },
    () => Workbook.load(bytes),
  );
  withSheet(
    (ws, wb) => {
      wb.setCalcMode(CalcMode.Manual);
      assert.equal(ws.getValue('A2'), 20);
    },
    () => Workbook.load(bytes),
  );
});

test('native reads follow the shared to-native cases', async () => {
  const cases = await table('to_native_cases.json');
  for (const c of cases.cases) {
    const want = expected(c.native);
    if (c.formula !== undefined) {
      assert.deepEqual(formulon.evaluateFormula(c.formula), want, c.formula);
      assert.deepEqual(formulon.toNative(formulon.evaluateTypedFormula(c.formula)), want, c.formula);
    } else {
      withSheet((ws) => assert.deepEqual(ws.getValue(c.cell), want, c.cell));
    }
  }
  for (const c of cases.facade_cases) {
    const want = expected(c.native);
    withSheet((ws) => {
      switch (c.call) {
        case 'evaluate_formula':
          assert.deepEqual(formulon.evaluateFormula(c.formula), want, c.id);
          assert.deepEqual(ws.evaluateFormula('J20', c.formula), want, c.id);
          break;
        case 'evaluate_formula_array':
          assert.deepEqual(ws.evaluateFormulaArray('J20', c.formula), want, c.id);
          break;
        case 'get_value':
          assert.deepEqual(ws.getValue(c.cell), want, c.id);
          break;
        case 'get_values':
          assert.deepEqual(ws.getValues(c.range), want, c.id);
          break;
        default:
          assert.fail(`unknown call ${c.call}`);
      }
    });
  }
});

test('the typed reads tag each value with its kind', () => {
  withSheet((ws) => {
    ws.setValue('A1', 1.5);
    ws.setValue('B1', 'x');
    ws.setValue('A2', true);
    ws.setFormula('B2', '=1/0');
    assert.deepEqual(ws.getTypedValue('A1'), { kind: ValueKind.Number, number: 1.5 });
    assert.deepEqual(ws.getTypedValue('C9'), { kind: ValueKind.Blank });
    assert.deepEqual(ws.getTypedValues('A1:C2'), [
      [{ kind: ValueKind.Number, number: 1.5 }, { kind: ValueKind.Text, text: 'x' }, { kind: ValueKind.Blank }],
      [
        { kind: ValueKind.Bool, boolean: true },
        { kind: ValueKind.Error, errorCode: ErrorCode.Div0 },
        { kind: ValueKind.Blank },
      ],
    ]);
    const error = ws.getValue('B2');
    assert.ok(error instanceof ExcelError);
    assert.equal(error.code, ErrorCode.Div0);
    assert.equal(String(error), '#DIV/0!');
  });
});

test('evaluateTypedFormula folds a 1x1 result and tags an array as an ArrayValue', () => {
  assert.deepEqual(formulon.evaluateTypedFormula('=1+1'), { kind: ValueKind.Number, number: 2 });
  assert.deepEqual(formulon.evaluateTypedFormula('=SEQUENCE(1,1)'), { kind: ValueKind.Number, number: 1 });
  const grid = formulon.evaluateTypedFormula('=SEQUENCE(2,2)');
  assert.deepEqual(grid, {
    kind: ValueKind.Array,
    rows: 2,
    cols: 2,
    cells: [1, 2, 3, 4].map((number) => ({ kind: ValueKind.Number, number })),
  });
  withSheet((ws) => {
    assert.deepEqual(ws.evaluateTypedFormula('J20', '=SEQUENCE(2,2)'), grid);
    assert.deepEqual(ws.evaluateTypedFormulaArray('J20', '=1+1'), [[{ kind: ValueKind.Number, number: 2 }]]);
  });
  assert.throws(() => formulon.evaluateFormula(1), TypeError);
});

test('toNative maps every Value a read returns and refuses a reference or LAMBDA value', () => {
  const { toNative } = formulon;
  assert.equal(toNative({ kind: ValueKind.Blank }), null);
  assert.equal(toNative({ kind: ValueKind.Text, text: '' }), '');
  assert.deepEqual(toNative({ kind: ValueKind.Error, errorCode: ErrorCode.NA }), new ExcelError(ErrorCode.NA));
  const cells = [
    { kind: ValueKind.Number, number: 1 },
    { kind: ValueKind.Error, errorCode: ErrorCode.Value },
  ];
  assert.deepEqual(toNative({ kind: ValueKind.Array, rows: 1, cols: 2, cells }), [
    [1, new ExcelError(ErrorCode.Value)],
  ]);
  assert.throws(() => toNative({ kind: ValueKind.Array, rows: 2, cols: 2, cells }), TypeError);
  assert.throws(() => toNative({ kind: ValueKind.Ref }), TypeError);
  assert.throws(() => toNative({ kind: ValueKind.Lambda }), TypeError);
  assert.throws(() => toNative(1), TypeError);
});

test('listCells and getCellsInRange return RangeCells with native values', () => {
  withSheet((ws) => {
    ws.setValue('A1', 2);
    ws.setFormula('B1', '=A1&"x"');
    const want = [
      { row: 0, col: 0, formula: null, value: 2, kind: ValueKind.Number },
      { row: 0, col: 1, formula: '=A1&"x"', value: '2x', kind: ValueKind.Text },
    ];
    assert.deepEqual(ws.listCells(), want);
    assert.deepEqual(ws.listCells('B1:B1'), [want[1]]);
    assert.deepEqual(ws.getCellsInRange('A1:B1'), { cells: want, nextCursor: null });
    const first = ws.getCellsInRange('A1:B1', null, 1);
    assert.deepEqual(first.cells, [want[0]]);
    assert.deepEqual(ws.getCellsInRange('A1:B1', first.nextCursor).cells, [want[1]]);
  });
});

test('getTypedValues reads a block that setValues stores back unchanged', () => {
  withSheet((ws) => {
    ws.setValues('A1', [
      [1, 'a', true],
      [null, { kind: ValueKind.Text, text: '=1' }, new ExcelError(ErrorCode.NA)],
    ]);
    const block = ws.getTypedValues('A1:C2');
    ws.setValues('E1', block);
    assert.deepEqual(ws.getTypedValues('E1:G2'), block);
    assert.equal(ws.getFormula('F2'), null);
    assert.equal(ws.getValue('F2'), '=1');
  });
});

test('getUsedRange bounds the cells holding a value or a formula', () => {
  withSheet((ws) => {
    assert.equal(ws.getUsedRange(), null);
    ws.updateCellFormat('D10', { numberFormat: '0.00' });
    assert.equal(ws.getUsedRange(), null);
    ws.setValue('C3', 1);
    ws.setFormula('B5', '=C3');
    ws.setValue('E2', 'x');
    assert.deepEqual(ws.getUsedRange(), { firstRow: 1, firstCol: 1, lastRow: 4, lastCol: 4 });
  });
});

test('iterRows yields native rows of the used range or of a given range', () => {
  withSheet((ws) => {
    assert.deepEqual([...ws.iterRows()], []);
    ws.setValues('B2', [
      [1, 'a'],
      [null, true],
    ]);
    const rows = ws.iterRows();
    assert.equal(typeof rows.next, 'function');
    assert.equal(rows[Symbol.iterator](), rows);
    assert.deepEqual(
      [...rows],
      [
        [1, 'a'],
        [null, true],
      ],
    );
    assert.deepEqual(
      [...ws.iterRows('A2:B3')],
      [
        [null, 1],
        [null, null],
      ],
    );
    assert.deepEqual([...ws.iterRows({ firstRow: 1, firstCol: 2, lastRow: 1, lastCol: 2 })], [['a']]);
  });
});

test('iterRows reads a range taller than one page', () => {
  withSheet((ws) => {
    ws.setValue('A1', 1);
    ws.setValue('B70000', 2);
    let count = 0;
    let last;
    for (const row of ws.iterRows()) {
      count += 1;
      last = row;
    }
    assert.equal(count, 70000);
    assert.deepEqual(last, [null, 2]);
  });
});

test('getDate follows the shared read cases', async () => {
  const cases = (await table('date_value_cases.json')).cases.filter(
    (c) => c.direction === 'read' && c.expect.js !== undefined,
  );
  assert.ok(cases.length > 0);
  for (const c of cases) {
    withSheet((ws, wb) => {
      wb.setDate1904(c.date1904);
      const want = c.expect.js;
      if (typeof c.input.serial === 'string') {
        // The cell cannot hold a non-finite number; the write is refused as the read would be.
        rejectsWith(() => ws.setNumber('A1', Number(c.input.serial)), StatusCode[want.error]);
        return;
      }
      if (c.input.serial !== undefined) ws.setNumber('A1', c.input.serial);
      else if (c.input.cell === 'text') ws.setText('A1', 'x');
      else if (c.input.cell === 'bool') ws.setBool('A1', true);
      else if (c.input.cell === 'error') ws.setError('A1', ErrorCode.NA);
      if (want.error !== undefined) {
        rejectsWith(() => ws.getDate('A1'), StatusCode[want.error]);
      } else if (want.date === undefined) {
        assert.equal(ws.getDate('A1'), want.value, c.id);
      } else {
        const d = want.date;
        const at = new Date(d.year, d.month - 1, d.day, d.hour, d.minute, d.second, d.millisecond);
        at.setFullYear(d.year);
        assert.equal(ws.getDate('A1').getTime(), at.getTime(), c.id);
      }
    });
  }
});

test('a Date written with setValue stores its serial and, on a General cell, a date format', async () => {
  const cases = (await table('date_value_cases.json')).cases.filter(
    (c) => c.direction === 'write' && c.expect.js !== undefined,
  );
  assert.ok(cases.length > 0);
  for (const c of cases) {
    const i = c.input;
    withSheet((ws, wb) => {
      wb.setDate1904(c.date1904);
      if (c.effective_format === 'custom') ws.updateCellFormat('A1', { numberFormat: '0.00' });
      ws.updateCellFormat('Z1', { numberFormat: '0.00' });
      const styled = ws.ooxml.getEffectiveStyle('Z1').xfIndex;
      if (c.effective_format === 'row_style') ws.ooxml.setRowXfIndex(0, 0, styled);
      if (c.effective_format === 'column_style') ws.ooxml.setColumnXfIndex(0, 0, styled);
      let date;
      if (i.kind === 'invalid_date') date = new Date(Number.NaN);
      else {
        const ms = i.subsecond?.js_ms ?? 0;
        date = new Date(i.year, i.month - 1, i.day, i.hour ?? 0, i.minute ?? 0, i.second ?? 0, ms);
        date.setFullYear(i.year);
      }
      const want = c.expect.js;
      if (want.error === 'RangeError') {
        assert.throws(() => ws.setValue('A1', date), RangeError, c.id);
        return;
      }
      if (want.error !== undefined) {
        rejectsWith(() => ws.setValue('A1', date), StatusCode[want.error]);
        return;
      }
      const before = wb.ooxml.getCellXf(ws.ooxml.getEffectiveStyle('A1').xfIndex).numFmtId;
      ws.setValue('A1', date);
      const serial = ws.getValue('A1');
      assert.ok(Math.abs(serial - want.serial) <= Math.abs(want.serial) * 1e-12, `${c.id}: ${serial}`);
      const after = wb.ooxml.getCellXf(ws.ooxml.getEffectiveStyle('A1').xfIndex).numFmtId;
      assert.equal(after, want.format_id ?? before, c.id);
    });
  }
});

test('load takes a Uint8Array, an ArrayBuffer or any ArrayBufferView', () => {
  const bytes = withSheet((ws, wb) => {
    ws.setValue('A1', 7);
    return wb.save();
  });
  const padded = new Uint8Array(bytes.length + 8);
  padded.set(bytes, 4);
  const forms = [
    bytes,
    bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength),
    new DataView(padded.buffer, 4, bytes.length),
    padded.subarray(4, 4 + bytes.length),
  ];
  for (const form of forms) {
    withSheet(
      (ws) => assert.equal(ws.getValue('A1'), 7),
      () => Workbook.load(form),
    );
  }
  assert.throws(() => Workbook.load([1, 2, 3]), TypeError);
});

test('a workbook iterates its sheet handles in order', () => {
  withSheet((ws, wb) => {
    const second = wb.addSheet('Two');
    assert.deepEqual([...wb], [ws, second]);
    assert.deepEqual([...wb], wb.listSheets());
  });
});

test('the tab colour takes a colour string or a Color and reads back a Color', () => {
  withSheet((ws) => {
    assert.equal(ws.getTabColor(), null);
    ws.setTabColor('#FF0000');
    assert.deepEqual(ws.getTabColor(), {
      kind: formulon.ColorKind.Rgb,
      argb: 0xffff0000,
      theme: 0,
      tint: 0,
      indexed: 0,
    });
    const theme = { kind: formulon.ColorKind.Theme, argb: 0, theme: 4, tint: 0.5, indexed: 0 };
    ws.setTabColor(theme);
    assert.deepEqual(ws.getTabColor(), theme);
    ws.setTabColor(null);
    assert.equal(ws.getTabColor(), null);
    assert.throws(() => ws.setTabColor('red'), TypeError);
  });
});
