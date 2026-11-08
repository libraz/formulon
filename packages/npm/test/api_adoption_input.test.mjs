// Input side of the facade: A1 text inside records, the positional arity guard and the method names in guard
// messages, the table-name refusal, trailing defaults, the defined-name scope, formula text versus setText, date
// inputs, and the getter -> setter round trips of tools/api/roundtrip_cases.json. FORMULON_PACKAGE=native runs it
// against the staged @libraz/formulon-native; otherwise FORMULON_NPM_ENTRY selects the staged @libraz/formulon
// subpath.

import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { fileURLToPath, pathToFileURL } from 'node:url';

const testDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(testDir, '..', '..', '..');
const native = process.env.FORMULON_PACKAGE === 'native';

async function loadPackage() {
  if (native) {
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
const { StatusCode, ValueKind, Workbook } = formulon;

const PNG_1X1 = Uint8Array.from(
  atob('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='),
  (c) => c.charCodeAt(0),
);
const PERSON = '{6B3C2E1A-0000-4000-8000-00000000A11C}';

function withSheet(body) {
  const wb = Workbook.createDefault();
  try {
    return body(wb.sheet(0), wb);
  } finally {
    if (!wb.disposed) wb.dispose();
  }
}

function thrown(call) {
  try {
    call();
  } catch (error) {
    return error;
  }
  assert.fail('expected the call to throw');
}

function rejectsWith(call, code, pattern) {
  const error = thrown(call);
  assert.equal(error.name, 'FormulonError', String(error));
  assert.equal(error.code, code, `${error.code}: ${error.message}`);
  if (pattern !== undefined) assert.match(error.message, pattern);
  return error;
}

const rect = (firstRow, firstCol, lastRow, lastCol) => ({ firstRow, firstCol, lastRow, lastCol });
const plain = (value) => JSON.parse(JSON.stringify(value));

function fillTableSource(ws) {
  ws.setText('A1', 'Region');
  ws.setText('B1', 'Amount');
  ws.setText('A2', 'East');
  ws.setNumber('B2', 10);
  ws.setText('A3', 'West');
  ws.setNumber('B3', 30);
}

function addCache(wb) {
  const cache = wb.addPivotCache();
  cache.setWorksheetSource({ ref: 'A1:B3', sheet: 'Sheet1' });
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

function addPivot(ws, cache, name = 'ByRegion', anchor = 'E1') {
  const pivot = ws.addPivot({ name, cacheId: cache.cacheId, anchor });
  pivot.addField({ sourceName: 'Region', axis: formulon.PivotAxis.Row });
  const amount = pivot.addField({ sourceName: 'Amount', axis: formulon.PivotAxis.Value });
  pivot.addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 });
  return pivot;
}

// ---- A1 text inside records ---------------------------------------------------------------------------------

test('a CF rule and a validation take their ranges as A1 strings, alone or in an array', () => {
  withSheet((ws) => {
    const { CfRuleType, ValidationType } = formulon;
    ws.addCfRule({ type: CfRuleType.Expression, ranges: 'A1:B2', formula1: 'TRUE' });
    ws.addCfRule({ type: CfRuleType.Expression, ranges: ['C1', rect(4, 4, 5, 5)], formula1: 'TRUE' });
    const rules = ws.listCfRules();
    assert.deepEqual(plain(rules[0].ranges), [rect(0, 0, 1, 1)]);
    assert.deepEqual(plain(rules[1].ranges), [rect(0, 2, 0, 2), rect(4, 4, 5, 5)]);
    ws.addValidation({ type: ValidationType.Whole, ranges: 'D1:D3', formula1: '1', formula2: '5' });
    assert.deepEqual(plain(ws.listValidations()[0].ranges), [rect(0, 3, 2, 3)]);
  });
});

test('a record field the raw layer reads stays visible when its ranges are rewritten', () => {
  withSheet((ws, wb) => {
    const { CfRuleType } = formulon;
    const dxf = wb.ooxml.addDxf({ font: { bold: true } });
    const hidden = { type: CfRuleType.Expression, ranges: 'A1:B2', formula1: 'TRUE' };
    Object.defineProperty(hidden, 'dxfIndex', { value: dxf, enumerable: false });
    ws.addCfRule(hidden);
    class Rule {
      get type() {
        return CfRuleType.Expression;
      }
      get ranges() {
        return 'C1';
      }
      get formula1() {
        return 'TRUE';
      }
      get dxfIndex() {
        return dxf;
      }
    }
    ws.addCfRule(new Rule());
    assert.deepEqual(
      ws.listCfRules().map((rule) => rule.dxfIndex),
      [dxf, dxf],
    );
  });
});

test('an auto filter takes A1 text in its range and in its nested sort state', () => {
  withSheet((ws) => {
    fillTableSource(ws);
    ws.setAutoFilter({
      range: 'A1:B3',
      sort: { ref: 'A2:B3', conditions: [{ ref: 'B2:B3', descending: true }] },
    });
    const filter = ws.getAutoFilter();
    assert.deepEqual(plain(filter.range), rect(0, 0, 2, 1));
    assert.deepEqual(plain(filter.sort.ref), rect(1, 0, 2, 1));
    assert.deepEqual(plain(filter.sort.conditions[0].ref), rect(1, 1, 2, 1));
  });
});

test('a table takes an A1 range on the receiving sheet, and its handle resolves one on its own sheet', () => {
  withSheet((ws) => {
    fillTableSource(ws);
    const table = ws.addTable({ range: 'A1:B3', name: 'Sales' });
    assert.deepEqual(plain(table.getProperties().range), { sheet: 0, rect: rect(0, 0, 2, 1) });
    assert.deepEqual(
      table.getProperties().columns.map((c) => c.name),
      ['Region', 'Amount'],
    );
    ws.setText('A4', 'North');
    table.update({ range: 'A1:B4' });
    assert.deepEqual(plain(table.getProperties().range), { sheet: 0, rect: rect(0, 0, 3, 1) });
    table.setAutoFilter({ range: 'A1:B4' });
    assert.deepEqual(plain(table.getAutoFilter().range), rect(0, 0, 3, 1));
    assert.equal(ws.addTable({ range: 'D1:E3', name: 'Other' }).getProperties().range.sheet, 0);
  });
});

test('a pivot anchor and a selection take one A1 cell, and the selection its ranges as A1 text', () => {
  withSheet((ws, wb) => {
    fillTableSource(ws);
    const cache = addCache(wb);
    addPivot(ws, cache);
    addPivot(ws, cache, 'Second', { row: 20, col: 6 });
    const anchors = ws.listPivots().map((p) => {
      const { rect: at } = p.getLayout();
      return [at.firstRow, at.firstCol];
    });
    assert.deepEqual(anchors, [
      [0, 4],
      [20, 6],
    ]);
    ws.setSelection({ active: 'C3', ranges: 'B2:D4' });
    assert.deepEqual(plain(ws.getSelection()), { activeRow: 2, activeCol: 2, ranges: [rect(1, 1, 3, 3)] });
    ws.setSelection({ active: { row: 0, col: 0 }, ranges: ['A1', 'C1:C2'] });
    assert.deepEqual(plain(ws.getSelection()), {
      activeRow: 0,
      activeCol: 0,
      ranges: [rect(0, 0, 0, 0), rect(0, 2, 1, 2)],
    });
  });
});

test('a data table, a reference and a print area take A1 text', () => {
  withSheet((ws, wb) => {
    assert.equal(wb.formatReference({ rect: 'B2:C3' }), 'B2:C3');
    assert.equal(wb.formatReference({ sheetName: 'Sheet1', rect: 'A1' }), 'Sheet1!A1');
    ws.setPrintArea('A1:C3');
    assert.deepEqual(plain(ws.getPrintArea()), [rect(0, 0, 2, 2)]);
    ws.setPrintArea(['A1:B2', 'D1']);
    assert.deepEqual(plain(ws.getPrintArea()), [rect(0, 0, 1, 1), rect(0, 3, 0, 3)]);
    ws.setNumber('A1', 1);
    ws.setNumber('A2', 2);
    ws.setNumber('A3', 3);
    ws.addDataTable({ rect: 'B2:B3', columnInput: 'A1' });
    assert.notEqual(ws.getDataTable('B2'), null);
  });
});

test('A1 text in a record field is refused with the method and the field it was given to', () => {
  withSheet((ws, wb) => {
    const { CfRuleType } = formulon;
    rejectsWith(
      () => ws.addCfRule({ type: CfRuleType.Expression, ranges: ['A1', 'Sheet1!B2'], formula1: 'TRUE' }),
      StatusCode.InvalidArgument,
      /^Sheet\.addCfRule ranges\[1\] "Sheet1!B2" names a sheet/,
    );
    rejectsWith(
      () => ws.addPivot({ name: 'P', cacheId: wb.addPivotCache().cacheId, anchor: 'A1:B2' }),
      StatusCode.InvalidArgument,
      /^Sheet\.addPivot anchor "A1:B2" is not one cell/,
    );
    rejectsWith(
      () => ws.setAutoFilter({ range: 'A1:B3', sort: { ref: 'Sheet1!A1:B3' } }),
      StatusCode.InvalidArgument,
      /^Sheet\.setAutoFilter sort\.ref /,
    );
    assert.throws(
      () => ws.setSelection({ active: 'A1', ranges: 7 }),
      (e) => e instanceof TypeError && /^Sheet\.setSelection ranges must be an array or one A1 string/.test(e.message),
    );
    assert.throws(
      () => ws.setAutoFilter({ range: 3 }),
      (e) => e instanceof TypeError && /^Sheet\.setAutoFilter range must be a CellRect or an A1 string/.test(e.message),
    );
    for (const text of ['A1 B2', 'XFE1', '']) {
      const error = thrown(() => ws.addCfRule({ type: CfRuleType.Expression, ranges: text, formula1: 'TRUE' }));
      assert.equal(error.name, 'FormulonError', `${JSON.stringify(text)}: ${error}`);
      assert.ok(error.parseError || error.code === StatusCode.InvalidArgument, `${JSON.stringify(text)}`);
    }
    assert.equal(ws.listCfRules().length, 0);
  });
});

// ---- Arity and guard messages ---------------------------------------------------------------------------------

function handlePrototypes(ws, wb) {
  fillTableSource(ws);
  const table = ws.addTable({ range: 'A1:B3', name: 'Tbl' });
  const cache = addCache(wb);
  const pivot = addPivot(ws, cache);
  const chart = ws.addChart({ kind: formulon.ChartInsertKind.Column, source: 'Sheet1!$A$1:$B$3', legend: null });
  const shape = ws.addShape({ presetGeometry: 'rect', anchorKind: 0, fromRow: 2, widthEmu: 9144, heightEmu: 9144 });
  const slicer = ws.addSlicer(table, 0, { fromRow: 20, fromCol: 3, toRow: 30, toCol: 5 });
  const picture = ws.addPicture(PNG_1X1);
  const proto = (value) => Object.getPrototypeOf(value);
  return {
    Workbook: [proto(wb), proto(wb.ooxml)],
    Sheet: [proto(ws), proto(ws.ooxml)],
    Table: [proto(table)],
    Pivot: [proto(pivot)],
    PivotCache: [proto(cache)],
    Drawing: [proto(proto(picture))],
    Picture: [proto(picture)],
    Chart: [proto(chart)],
    Shape: [proto(shape)],
    Slicer: [proto(slicer)],
  };
}

test('every facade method refuses more positional arguments than it declares, naming itself', () => {
  withSheet((ws, wb) => {
    const tooMany = new Array(16).fill(0);
    let checked = 0;
    for (const [owner, protos] of Object.entries(handlePrototypes(ws, wb))) {
      for (const proto of protos) {
        for (const key of [...Object.getOwnPropertyNames(proto), Symbol.iterator]) {
          if (key === 'constructor' || !Object.hasOwn(proto, key)) continue;
          const descriptor = Object.getOwnPropertyDescriptor(proto, key);
          if (typeof descriptor.value !== 'function') continue;
          const name = typeof key === 'symbol' ? '[Symbol.iterator]' : key;
          const error = thrown(() => descriptor.value.apply(undefined, tooMany));
          assert.ok(error instanceof TypeError, `${owner}.${name}: ${error}`);
          assert.match(
            error.message,
            new RegExp(`^${owner}\\.${name.replace(/[[\]]/g, '\\$&')} takes at most \\d+ arguments?$`),
          );
          checked += 1;
        }
      }
    }
    assert.ok(checked > 300, `${checked} methods`);
  });
});

test('the string form of a call takes one argument fewer per A1 cell or span', () => {
  withSheet((ws, wb) => {
    ws.setNumber('A1', 1);
    ws.setNumber(0, 1, 2);
    assert.throws(() => ws.setNumber('A1', 1, 2), {
      name: 'TypeError',
      message: 'Sheet.setNumber takes at most 2 arguments',
    });
    assert.throws(() => ws.setNumber(0, 0, 1, 2), {
      name: 'TypeError',
      message: 'Sheet.setNumber takes at most 3 arguments',
    });
    assert.throws(() => ws.getValue('A1', 0), {
      name: 'TypeError',
      message: 'Sheet.getValue takes at most 1 argument',
    });
    ws.setColumnWidth('B:C', 12);
    assert.throws(() => ws.setColumnWidth('B:C', 12, 1), { message: 'Sheet.setColumnWidth takes at most 2 arguments' });
    assert.throws(() => wb.listSheets(1), { message: 'Workbook.listSheets takes at most 0 arguments' });
    assert.throws(() => wb.ooxml.listCellXfs(1), { message: 'Workbook.listCellXfs takes at most 0 arguments' });
    assert.equal(ws.getValue('A1'), 1);
    assert.equal(ws.getValue('B1'), 2);
    assert.equal(wb[Symbol.dispose], wb.dispose);
  });
});

test('a sheet-qualified address names the call that would read it', () => {
  withSheet((ws) => {
    rejectsWith(() => ws.getValue('Sheet1!A1'), StatusCode.InvalidArgument, /wb\.sheet\("Sheet1"\)\.getValue\("A1"\)/);
    rejectsWith(
      () => ws.setNumber('Sheet1!B2', 1),
      StatusCode.InvalidArgument,
      /wb\.sheet\("Sheet1"\)\.setNumber\("B2"\)/,
    );
  });
});

test('a table name that reads as a cell is refused with that reason', () => {
  withSheet((ws) => {
    fillTableSource(ws);
    const error = rejectsWith(
      () => ws.addTable({ range: 'A1:B3', name: 'T2' }),
      StatusCode.InvalidArgument,
      /^invalid table name: "T2" is a cell address$/,
    );
    assert.match(error.context, /name=T2/);
    assert.equal(ws.addTable({ range: 'A1:B3', name: 'Sales' }).name, 'Sales');
  });
});

test('a guard message names the method it guards', () => {
  withSheet((ws) => {
    assert.throws(() => ws.setValue('A1', new Date(Number.NaN)), { name: 'RangeError', message: /^Sheet\.setValue / });
    assert.throws(() => ws.setValue('A1', Symbol('x')), { name: 'TypeError', message: /^Sheet\.setValue / });
    assert.throws(() => ws.setValues('A1', [[Symbol('x')]]), { name: 'TypeError', message: /^Sheet\.setValues / });
  });
});

// ---- Trailing defaults and the defined-name scope ----------------------------------------------------------

test('an omitted trailing argument takes its ledger default', () => {
  withSheet((ws, wb) => {
    ws.setNumber('A1', 1);
    ws.setValue('B1', '=A1+1');
    ws.setValue('C1', '=B1+1');
    const cells = (nodes) => nodes.map(({ row, col }) => [row, col]);
    assert.deepEqual(cells(ws.getPrecedents('C1')), [[0, 1]]);
    assert.deepEqual(cells(ws.getPrecedents(0, 2)), [[0, 1]]);
    assert.deepEqual(cells(ws.getDependents('A1')), [[0, 1]]);
    assert.deepEqual(cells(ws.getDependents(0, 0, undefined)), [[0, 1]]);
    if (typeof wb.recalcParallel === 'function') assert.equal(typeof wb.recalcParallel().cellsEvaluated, 'number');
  });
});

test('copySheet appends when insertBefore is omitted or null', () => {
  withSheet((ws, wb) => {
    wb.addSheet('Last');
    const appended = wb.copySheet(0);
    assert.equal(appended.index, 2);
    assert.equal(appended.name, 'Sheet1 (2)');
    const named = wb.copySheet(0, null, 'Copy');
    assert.equal(named.index, 3);
    assert.equal(named.name, 'Copy');
    assert.equal(wb.copySheet(ws, 1, 'Front').index, 1);
    assert.deepEqual(
      wb.listSheets().map((sheet) => sheet.name),
      ['Sheet1', 'Front', 'Last', 'Sheet1 (2)', 'Copy'],
    );
  });
});

test('evaluateCfRange dates timePeriod rules by the pinned clock, else today', () => {
  withSheet((ws, wb) => {
    const { CfRuleType, DatePeriod } = formulon;
    ws.addCfRule({ type: CfRuleType.TimePeriod, ranges: 'A1:A2', timePeriod: DatePeriod.Today });
    ws.setValue('A1', '=TODAY()');
    ws.setValue('A2', new Date(2020, 0, 15));
    const matched = () => ws.evaluateCfRange('A1:A2').map(({ row, col }) => [row, col]);
    assert.deepEqual(matched(), [[0, 0]]);
    assert.deepEqual(
      ws.evaluateCfRange('A1:A2', { today: new Date(2020, 0, 15, 12) }).map(({ row }) => row),
      [1],
    );
    assert.throws(() => ws.evaluateCfRange('A1:A2', { today: new Date(Number.NaN) }), RangeError);
    wb.setPinnedNow({ year: 2020, month: 1, day: 15, hour: 12, minute: 0, second: 0 });
    wb.recalc();
    assert.deepEqual(matched(), [
      [0, 0],
      [1, 0],
    ]);
  });
});

test('setDefinedName scopes a name to the sheet its scope names', () => {
  withSheet((ws, wb) => {
    wb.addSheet('Data');
    wb.setDefinedName({ name: 'ByIndex', formula: '=Sheet1!$A$1', scope: 0 });
    wb.setDefinedName({ name: 'ByName', formula: '=Data!$A$1', scope: 'Data' });
    wb.setDefinedName({ name: 'ByHandle', formula: '=Sheet1!$A$2', scope: ws });
    wb.setDefinedName({ name: 'Global', formula: '=Sheet1!$A$3', scope: null });
    assert.equal(wb.getDefinedName('ByIndex', 0).localSheetId, 0);
    assert.equal(wb.getDefinedName('ByName', 'Data').localSheetId, 1);
    assert.equal(wb.getDefinedName('ByHandle', ws).localSheetId, 0);
    assert.equal(wb.getDefinedName('Global').localSheetId, null);
    assert.equal(wb.getDefinedName('ByIndex'), null);
    rejectsWith(
      () => wb.setDefinedName({ name: 'Lost', formula: '=1', scope: 'Missing' }),
      StatusCode.NotFound,
      /Missing/,
    );
    assert.throws(() => wb.setDefinedName({ name: 'Far', formula: '=1', scope: 1.5 }), {
      name: 'RangeError',
      message: /^Workbook\.setDefinedName scope /,
    });
  });
});

// ---- Formula text and dates -----------------------------------------------------------------------------------

test('a native string starting with = is a formula in setValue and setValues, never in setText', () => {
  withSheet((ws) => {
    ws.setValues('A1', [[2, '=A1*3', '=1']]);
    assert.equal(ws.getFormula('B1'), '=A1*3');
    assert.deepEqual(ws.getValues('A1:C1'), [[2, 6, 1]]);
    ws.setText('D1', '=A1');
    assert.equal(ws.getFormula('D1'), null);
    assert.equal(ws.getValue('D1'), '=A1');
    ws.setValues('E1', [[{ kind: ValueKind.Text, text: '=A1' }]]);
    assert.equal(ws.getFormula('E1'), null);
    assert.equal(ws.getValue('E1'), '=A1');
  });
});

test('dates reach a cell through setValue and setValues, and the clock pins to a Date', () => {
  withSheet((ws, wb) => {
    ws.setValues('A1', [[new Date(2024, 2, 1), new Date(2024, 2, 1, 12)]]);
    assert.deepEqual(ws.getValues('A1:B1'), [[45352, 45352.5]]);
    assert.deepEqual(ws.getDate('B1'), new Date(2024, 2, 1, 12));
    assert.equal(formulon.convertCivilToSerial({ year: 2024, month: 3, day: 1 }), 45352);
    wb.setPinnedNow(new Date(2024, 2, 1, 6));
    ws.setFormula('C1', '=NOW()');
    wb.recalc();
    assert.equal(ws.getValue('C1'), 45352.25);
  });
});

// ---- Getter -> setter round trips -----------------------------------------------------------------------------

const SPARKLINES = () => ({
  type: formulon.SparklineType.Column,
  sparklines: [{ formula: 'Sheet1!A1:D1', location: 'F1' }],
});

// Per roundtrip_cases.json id: `make(ws, wb)` builds the object, `get(target)` reads the record and `put(target,
// record)` writes it back; a readd case also has `again(target)`, the record its re-added copy reads.
const DRIVERS = {
  cell_format: {
    make: (ws) => ws.updateCellFormat('B1', { font: { bold: true }, numberFormat: '0.00' }) ?? ws,
    get: (ws) => ws.getCellFormat('B1'),
    put: (ws, record) => ws.updateCellFormat('B1', record),
  },
  chart_legend: {
    make: (ws) => ws.addChart({ kind: formulon.ChartInsertKind.Column, source: 'Sheet1!$A$1:$B$3', legend: 0 }),
    get: (chart) => chart.getProperties().legend,
    put: (chart, record) => chart.updateLegend(record),
  },
  format_defaults: {
    make: (ws) => ws,
    get: (ws) => ws.getFormatDefaults(),
    put: (ws, record) => ws.updateFormatDefaults(record),
  },
  header_footer: {
    make: (ws) => {
      ws.updateHeaderFooter({ oddHeader: '&CTitle' });
      return ws;
    },
    get: (ws) => ws.getHeaderFooter(),
    put: (ws, record) => ws.updateHeaderFooter(record),
  },
  iterative: {
    make: (_ws, wb) => wb,
    get: (wb) => wb.getIterative(),
    put: (wb, record) => wb.updateIterative(record),
  },
  page_margins: {
    make: (ws) => ws,
    get: (ws) => ws.getPageMargins(),
    put: (ws, record) => ws.updatePageMargins(record),
  },
  page_setup: {
    make: (ws) => ws,
    get: (ws) => ws.getPageSetup(),
    put: (ws, record) => ws.updatePageSetup(record),
  },
  pivot_cache_options: {
    make: (_ws, wb) => addCache(wb),
    get: (cache) => cache.getOptions(),
    put: (cache, record) => cache.updateOptions(record),
  },
  pivot_field: {
    make: (ws, wb) => addPivot(ws, addCache(wb)),
    get: (pivot) => pivot.listFields()[0],
    put: (pivot, record) => pivot.updateField(0, record),
  },
  pivot_grand_totals: {
    make: (ws, wb) => addPivot(ws, addCache(wb)),
    get: (pivot) => pivot.getGrandTotals(),
    put: (pivot, record) => pivot.updateGrandTotals(record),
  },
  pivot_item: {
    make: (ws, wb) => {
      const pivot = addPivot(ws, addCache(wb));
      pivot.addFieldItem(0, 'East');
      return pivot;
    },
    get: (pivot) => pivot.getFieldItem(0, 0),
    put: (pivot, record) => pivot.updateFieldItem(0, 0, record),
  },
  pivot_options: {
    make: (ws, wb) => addPivot(ws, addCache(wb)),
    get: (pivot) => pivot.getOptions(),
    put: (pivot, record) => pivot.updateOptions(record),
  },
  pivot_style_info: {
    make: (ws, wb) => {
      const pivot = addPivot(ws, addCache(wb));
      pivot.updateStyleOptions({ styleName: 'PivotStyleLight16', showRowHeaders: true });
      return pivot;
    },
    get: (pivot) => pivot.getStyleOptions(),
    put: (pivot, record) => pivot.updateStyleOptions(record),
  },
  print_options: {
    make: (ws) => ws,
    get: (ws) => ws.getPrintOptions(),
    put: (ws, record) => ws.updatePrintOptions(record),
  },
  protected_range: {
    make: (ws) => {
      ws.addProtectedRange({ name: 'Totals', ranges: [rect(1, 1, 2, 2)] });
      return ws;
    },
    get: (ws) => ws.listProtectedRanges()[0],
    put: (ws, record) => ws.updateProtectedRange(0, record),
  },
  shape: {
    make: (ws) => ws.addShape({ presetGeometry: 'rect', anchorKind: 0, fromRow: 2, widthEmu: 9144, heightEmu: 9144 }),
    get: (shape) => shape.getProperties(),
    put: (shape, record) => shape.update(record),
  },
  sheet_protection: {
    make: (ws) => {
      ws.setSheetProtection({ allowSort: true, allowFormatCells: true });
      return ws;
    },
    get: (ws) => ws.getSheetProtection(),
    put: (ws, record) => ws.updateSheetProtection(record),
  },
  sheet_view: {
    make: (ws) => ws,
    get: (ws) => ws.getView(),
    put: (ws, record) => ws.updateView(record),
  },
  slicer: {
    make: (ws) => {
      const table = ws.addTable({ range: 'A1:B3', name: 'Tbl' });
      return ws.addSlicer(table, 0, { fromRow: 20, fromCol: 3, toRow: 30, toCol: 5 });
    },
    get: (slicer) => slicer.getProperties(),
    put: (slicer, record) => slicer.update(record),
  },
  sparkline_group: {
    make: (ws) => {
      ws.addSparklineGroup(SPARKLINES());
      return ws;
    },
    get: (ws) => ws.listSparklineGroups()[0],
    put: (ws, record) => ws.updateSparklineGroup(0, record),
  },
  table: {
    make: (ws) => ws.addTable({ range: 'A1:B3', name: 'Tbl' }),
    get: (table) => table.getProperties(),
    put: (table, record) => table.update(record),
  },
  table_style_info: {
    make: (ws) => ws.addTable({ range: 'A1:B3', name: 'Tbl' }),
    get: (table) => table.getStyleOptions(),
    put: (table, record) => table.updateStyleOptions(record),
  },
  table_style_defaults: {
    make: (_ws, wb) => wb,
    get: (wb) => wb.getTableStyleDefaults(),
    put: (wb, record) => wb.updateTableStyleDefaults(record),
  },
  threaded_comment: {
    make: (ws, wb) => {
      wb.addPerson({ id: PERSON, displayName: 'Alice', userId: 'alice', providerId: 'None' });
      ws.addThreadedComment({
        id: '{00000000-0000-4000-8000-000000000001}',
        row: 4,
        col: 2,
        personId: PERSON,
        created: { year: 2026, month: 1, day: 2, hour: 3, minute: 4, second: 5 },
        text: 'Hi',
      });
      return ws;
    },
    get: (ws) => ws.getThreadedComment('{00000000-0000-4000-8000-000000000001}'),
    put: (ws, record) => ws.updateThreadedComment('{00000000-0000-4000-8000-000000000001}', record),
  },
  cf_rule: {
    make: (ws) => {
      ws.addCfRule({ type: formulon.CfRuleType.Expression, ranges: 'A1:B2', formula1: 'A1>1' });
      return ws;
    },
    get: (ws) => ws.listCfRules()[0],
    put: (ws, record) => ws.addCfRule(record),
    again: (ws) => ws.listCfRules()[1],
    identity: ['id', 'priority'],
  },
  comment: {
    make: (ws) => {
      ws.setComment('A1', { author: 'Alice', text: 'Note' });
      return ws;
    },
    get: (ws) => ws.getComment('A1'),
    put: (ws, record) => ws.setComment('B1', record),
    again: (ws) => ws.getComment('B1'),
  },
  hyperlink: {
    make: (ws) => {
      ws.setHyperlink('A1', { target: 'https://example.com/', display: 'Example' });
      return ws;
    },
    get: (ws) => ws.getHyperlink('A1'),
    put: (ws, record) => ws.setHyperlink('B1', record),
    again: (ws) => ws.getHyperlink('B1'),
    identity: ['ref', 'rect'],
  },
  validation: {
    make: (ws) => {
      ws.addValidation({ type: formulon.ValidationType.Whole, ranges: 'D1:D3', formula1: '1', formula2: '5' });
      return ws;
    },
    get: (ws) => ws.listValidations()[0],
    put: (ws, record) => ws.addValidation(record),
    again: (ws) => ws.listValidations()[1],
  },
};

const without = (record, keys = []) =>
  Object.fromEntries(Object.entries(plain(record)).filter(([k]) => !keys.includes(k)));

test('every getter record goes back through its setter unchanged', async () => {
  const read = async (name) => JSON.parse(await readFile(path.join(repoRoot, 'tools', 'api', name), 'utf8'));
  const { cases } = await read('roundtrip_cases.json');
  const surface = native ? 'native' : 'wasm';
  const pending = new Set(
    (await read('roundtrip_pending.json')).pending.filter((p) => p.surfaces.includes(surface)).map((p) => p.id),
  );
  const rows = cases.filter((c) => c.surfaces.includes(surface) && !pending.has(c.id));
  assert.deepEqual(
    rows.map((c) => c.id).filter((id) => DRIVERS[id] === undefined),
    [],
  );
  const failures = [];
  for (const c of rows) {
    const driver = DRIVERS[c.id];
    withSheet((ws, wb) => {
      try {
        fillTableSource(ws);
        const target = driver.make(ws, wb);
        const before = plain(driver.get(target));
        assert.notEqual(before, null);
        driver.put(target, driver.get(target));
        if (c.kind === 'readd') {
          assert.deepEqual(without(driver.again(target), driver.identity), without(before, driver.identity));
        } else {
          assert.deepEqual(plain(driver.get(target)), before);
        }
      } catch (error) {
        failures.push(`${c.id}: ${error.message}`);
      }
    });
  }
  assert.deepEqual(failures, []);
});
