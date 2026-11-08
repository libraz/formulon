import test from 'node:test';
import { inflateRawSync } from 'node:zlib';
import {
  assert,
  expectFailure,
  getModule,
  loadStagedModule,
  makeIterativeWorkbook,
  path,
  pkgRoot,
  readFile,
} from './smoke_support.mjs';

// ---- Pivot mutator marshalling -----------------------------------------
//
// Every pivot mutator is driven once against a small pivot and checked
// through a read of the result -- the projected grid wherever the mutation
// reaches it -- so an argument dropped, reordered or mis-converted on the
// way to the C ABI shows up as a wrong grid rather than a green `.ok`. The
// completeness test below keeps the table in step with the declaration
// file.

// Region / Product carry shared items and their records store item
// indices; Amount and Date are plain numbers. Only Region is laid out, so
// the baseline grid is two regions, their sums and the grand total.
function buildPivot(mod) {
  const wb = mod.Workbook.createDefault();
  wb.setProfileId('win-365-ja_JP');
  const cache = wb.addPivotCache(0);
  const cacheId = cache.cacheId;
  wb.pivotCache(cacheId).setWorksheetSource({ range: 'A1:D5', sheet: 'Sheet1' });
  for (const name of ['Region', 'Product', 'Amount', 'Date']) wb.pivotCache(cacheId).addField(name);
  for (const item of ['East', 'West']) wb.pivotCache(cacheId).addFieldSharedItem(0, item);
  for (const item of ['Apple', 'Pear']) wb.pivotCache(cacheId).addFieldSharedItem(1, item);
  for (const values of [
    [0, 0, 10, 45292],
    [1, 0, 30, 45323],
    [0, 1, 5, 45658],
    [1, 1, 7, 45689],
  ]) {
    const rec = wb.pivotCache(cacheId).addRecord();
    for (const [field, v] of values.entries()) wb.pivotCache(cacheId).setRecordValue(rec, field, v);
  }
  const pivot = wb.sheet(0).addPivot({ name: 'Pivot1', cacheId: cacheId, anchor: { row: 0, col: 6 } });
  const addField = (spec) => wb.sheet(0).pivot(pivot).addField(spec);
  const region = addField({ sourceName: 'Region', axis: 0 });
  wb.sheet(0).pivot(pivot).addFieldItem(region, 'East', { visible: true });
  wb.sheet(0).pivot(pivot).addFieldItem(region, 'West', { visible: true });
  const product = addField({ sourceName: 'Product', axis: 2 });
  wb.sheet(0).pivot(pivot).addFieldItem(product, 'Apple', { visible: true });
  wb.sheet(0).pivot(pivot).addFieldItem(product, 'Pear', { visible: true });
  const amount = addField({ sourceName: 'Amount', axis: 2 });
  const date = addField({ sourceName: 'Date', axis: 2 });
  wb.sheet(0).pivot(pivot).setRowFieldOrder([region]);
  wb.sheet(0).pivot(pivot).addDataField({ name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 });
  return { mod, wb, cacheId, pivot, region, product, amount, date };
}

// The projected grid as one `|`-joined string per sheet row.
function pivotGrid(wb, pivot = 0) {
  const layout = wb.sheet(0).pivot(pivot).getLayout();
  const rows = new Map();
  for (const cell of layout.cells) {
    const v = cell.value;
    const text =
      typeof v === 'number'
        ? String(v)
        : typeof v === 'string'
          ? v
          : v !== null && typeof v === 'object'
            ? `#${v.code}`
            : '';
    if (!rows.has(cell.row)) rows.set(cell.row, new Array(layout.range.lastCol - layout.range.firstCol + 1).fill(''));
    rows.get(cell.row)[cell.col - layout.range.firstCol] = text;
  }
  return [...rows.keys()].sort((a, b) => a - b).map((row) => rows.get(row).join('|'));
}

const BASE_GRID = ['行ラベル|Sum of Amount', 'East|15', 'West|37', '総計|52'];
const TWO_LEVEL_GRID = [
  '行ラベル|Sum of Amount',
  'East|15',
  'Apple|10',
  'Pear|5',
  'West|37',
  'Apple|30',
  'Pear|7',
  '総計|52',
];
const DATE_SERIAL_GRID = ['行ラベル|Sum of Amount', '45292|10', '45323|30', '45658|5', '45689|7', '総計|52'];
const wrongKind = (err) => err instanceof TypeError || err instanceof RangeError;
const addNet = (b) => b.wb.pivotCache(b.cacheId).addCalculatedField('Net', 'Amount*2');
const addEw = (b) => b.wb.pivotCache(b.cacheId).addCalculatedItem(b.region, 'EW', 'East+West');
const twoLevels = (b) => b.wb.sheet(0).pivot(b.pivot).setRowFieldOrder([b.region, b.product]);
const dateRows = (b) => b.wb.sheet(0).pivot(b.pivot).setRowFieldOrder([b.date]);
const hideEast = (b) => b.wb.sheet(0).pivot(b.pivot).updateFieldItem(b.region, 0, { visible: false });
const WEST_ONLY = ['行ラベル|Sum of Amount', 'West|37', '総計|37'];
const GREATER_THAN_20 = Object.freeze({ axis: 0, fieldName: 'Region', type: 1, value: 20 });
const BETWEEN_10_20 = Object.freeze({
  axis: 0,
  fieldName: 'Region',
  type: 2,
  value: 10,
  valueHigh: 20,
});
// Points record 0's Region at the shared item a row has just appended.
const pointAtNewItem = (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 0, 2);

// The raw text of one stored part of a saved package.
function packagePart(bytes, name) {
  const buf = Buffer.from(bytes);
  let eocd = buf.length - 22;
  while (eocd >= 0 && buf.readUInt32LE(eocd) !== 0x06054b50) eocd -= 1;
  assert.ok(eocd >= 0, 'missing ZIP end record');
  let at = buf.readUInt32LE(eocd + 16);
  for (let i = buf.readUInt16LE(eocd + 10); i > 0; i -= 1) {
    const method = buf.readUInt16LE(at + 10);
    const size = buf.readUInt32LE(at + 20);
    const nameLen = buf.readUInt16LE(at + 28);
    const next = at + 46 + nameLen + buf.readUInt16LE(at + 30) + buf.readUInt16LE(at + 32);
    if (buf.toString('utf8', at + 46, at + 46 + nameLen) === name) {
      const local = buf.readUInt32LE(at + 42);
      const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28);
      const raw = buf.subarray(start, start + size);
      return (method === 0 ? raw : inflateRawSync(raw)).toString('utf8');
    }
    at = next;
  }
  assert.fail(`${name} not in the package`);
}

// One case per value kind the two cache value writers dispatch on.
const ERROR_NA = Object.freeze({ kind: 4, errorCode: 6 });
const ERROR_DIV0 = Object.freeze({ kind: 4, errorCode: 1 });
const SHARED_ITEM_CASES = [
  {
    label: 'number',
    act: (b) => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, 42.5),
    expect: (b) =>
      assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '42.5|10', 'East|5', 'West|37', '総計|52']),
  },
  {
    label: 'date',
    act: (b) => b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, new Date(2024, 0, 1)),
    expect: (b) => {
      const items = b.wb.pivotCache(b.cacheId).listFieldSharedItems(3);
      assert.equal(items[0].isDate, true);
      assert.equal(items[0].value, 45292);
    },
    noPointing: true,
  },
  {
    label: 'text',
    act: (b) => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, 'North'),
    expect: (b) =>
      assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'North|10', 'West|37', '総計|52']),
  },
  {
    label: 'boolean',
    act: (b) => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, true),
    expect: (b) =>
      assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', 'TRUE|10', '総計|52']),
  },
  {
    label: 'blank',
    act: (b) => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, null),
    expect: (b) =>
      assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '(空白)|10', '総計|52']),
  },
  {
    label: 'error',
    act: (b) => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, ERROR_NA),
    expect: (b) =>
      assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '#N/A|10', '総計|52']),
  },
];
const RECORD_VALUE_CASES = [
  {
    label: 'number',
    act: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, 1000),
    expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|1005', 'West|37', '総計|1042']),
  },
  {
    label: 'date',
    act: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 3, new Date(2024, 0, 9)),
    expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
  },
  {
    label: 'text',
    act: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 0, 'West'),
    expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|47', '総計|52']),
  },
  {
    label: 'boolean',
    act: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, true),
    expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
  },
  {
    label: 'blank',
    act: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, null),
    expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
  },
  {
    label: 'error',
    act: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, ERROR_DIV0),
    expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|#1', 'West|37', '総計|#1']),
  },
];

// [method, { setup, act, after, expect }]. `act` performs the one call
// under test; `expect` reads the result back.
const PIVOT_MUTATORS = [
  [
    'addPivotCache',
    {
      act: (b) => b.wb.addPivotCache(77),
      expect: (b, r) => {
        assert.equal(r.cacheId, 77);
        assert.equal(b.wb.pivotCache(77), r);
        assert.equal(b.wb.listPivotCaches().length, 2);
      },
    },
  ],
  [
    'Pivot.getCache',
    { act: (b) => b.wb.sheet(0).pivot(b.pivot).getCache(), expect: (b, r) => assert.equal(r.cacheId, b.cacheId) },
  ],
  [
    'removePivotCache',
    {
      setup: (b) => {
        b.spare = b.wb.addPivotCache(0);
      },
      act: (b) => b.wb.removePivotCache(b.spare),
      expect: (b) => {
        assert.equal(b.wb.listPivotCaches().length, 1);
        assert.deepEqual(
          b.wb.listPivotCaches().map((c) => c.cacheId),
          [b.cacheId],
        );
      },
    },
  ],
  [
    'PivotCache.setWorksheetSource',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).setWorksheetSource({ range: 'B2:E6', sheet: 'Sheet1' }),
      expect: (b) => {
        const source = b.wb.pivotCache(b.cacheId).getWorksheetSource();
        assert.equal(source.range.sheet, 0);
        assert.deepEqual({ ...source.range.rect }, { firstRow: 1, firstCol: 1, lastRow: 5, lastCol: 4 });
        assert.equal(source.name, null);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'PivotCache.addField',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addField('Extra'),
      expect: (b, r) => {
        assert.equal(r, 4);
        assert.equal(b.wb.pivotCache(b.cacheId).listFields()[4].name, 'Extra');
      },
    },
  ],
  [
    'PivotCache.clearFields',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).clearFields(),
      expect: (b) => {
        assert.equal(b.wb.pivotCache(b.cacheId).listFields().length, 0);
        // The data field now names a cache field that no longer exists.
        expectFailure(() => b.wb.sheet(0).pivot(b.pivot).getLayout());
      },
    },
  ],
  [
    'PivotCache.addFieldSharedItem',
    {
      act: (b) => SHARED_ITEM_CASES[0].act(b),
      after: pointAtNewItem,
      expect: (b) => SHARED_ITEM_CASES[0].expect(b),
    },
  ],
  [
    'PivotCache.clearFieldSharedItems',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).clearFieldSharedItems(0),
      // With no shared items the records' indices render as themselves.
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '0|15', '1|37', '総計|52']),
    },
  ],
  [
    'PivotCache.addRecord',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addRecord(),
      expect: (b, r) => {
        assert.equal(r, 4);
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|15', 'West|37', '(空白)|', '総計|52']);
      },
    },
  ],
  [
    'PivotCache.clearRecords',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).clearRecords(),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '総計|']),
    },
  ],
  [
    'PivotCache.setRecordValue',
    {
      act: (b) => RECORD_VALUE_CASES[0].act(b),
      expect: (b) => RECORD_VALUE_CASES[0].expect(b),
    },
  ],
  [
    'addPivot',
    {
      act: (b) => b.wb.sheet(0).addPivot({ name: 'Pivot2', cacheId: b.cacheId, anchor: { row: 20, col: 1 } }),
      expect: (b, r) => {
        assert.equal(r.index, 1);
        assert.equal(b.wb.sheet(0).pivot(1), r);
        const layout = b.wb.sheet(0).pivot(1).getLayout();
        assert.equal(layout.range.firstRow, 20);
        assert.equal(layout.range.firstCol, 1);
        assert.deepEqual(pivotGrid(b.wb, 0), BASE_GRID);
      },
    },
  ],
  [
    'removePivot',
    {
      act: (b) => b.wb.sheet(0).removePivot(b.pivot),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).listPivots().length, 0);
        expectFailure(() => b.wb.sheet(0).pivot(0).getLayout());
      },
    },
  ],
  [
    'Pivot.rename',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).rename('Renamed'),
      // A pivot's name is only observable in the part the writer emits.
      expect: (b) => {
        const saved = b.wb.save();
        assert.match(packagePart(saved, 'xl/pivotTables/pivotTable1.xml'), /<pivotTableDefinition [^>]*name="Renamed"/);
      },
    },
  ],
  [
    'Pivot.setAnchor',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).setAnchor({ row: 10, col: 2 }),
      expect: (b) => {
        const layout = b.wb.sheet(0).pivot(b.pivot).getLayout();
        assert.equal(layout.range.firstRow, 10);
        assert.equal(layout.range.firstCol, 2);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'Pivot.updateGrandTotals',
    {
      setup: (b) => b.wb.sheet(0).pivot(b.pivot).setColumnFieldOrder([b.product]),
      // Rows off drops the per-row total column; columns on keeps the
      // bottom total row, so a swapped pair would show the opposite.
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateGrandTotals({ rows: false, columns: true }),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), [
          'Sum of Amount|列ラベル|',
          '行ラベル|Apple|Pear',
          'East|10|5',
          'West|30|7',
          '総計|40|12',
        ]),
    },
  ],
  [
    'Pivot.updateStyleOptions',
    {
      act: (b) =>
        b.wb.sheet(0).pivot(b.pivot).updateStyleOptions({
          styleName: 'PivotStyleMedium9',
          showRowHeaders: false,
          showColumnHeaders: true,
          showRowStripes: true,
          showColumnStripes: false,
          showLastColumn: true,
        }),
      expect: (b) => {
        const style = b.wb.sheet(0).pivot(b.pivot).getStyleOptions();
        assert.equal(style.styleName, 'PivotStyleMedium9');
        assert.deepEqual(
          [
            style.showRowHeaders,
            style.showColumnHeaders,
            style.showRowStripes,
            style.showColumnStripes,
            style.showLastColumn,
          ],
          [false, true, true, false, true],
        );
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'Pivot.setReportLayout',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).setReportLayout(1),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).pivot(b.pivot).getReportLayout(), 1);
        assert.deepEqual(pivotGrid(b.wb), ['Region|Sum of Amount', 'East|15', 'West|37', '総計|52']);
      },
    },
  ],
  [
    'Pivot.addField',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).addField({ sourceName: 'Product', customName: 'Fruit', axis: 3 }),
      expect: (b, r) => {
        // Product is a cache column the fixture already added, so it is updated in place.
        assert.equal(r, b.product);
        assert.deepEqual(pivotGrid(b.wb), ['Fruit|(すべて)', '|', ...BASE_GRID]);
      },
    },
  ],
  [
    'Pivot.clearFields',
    {
      setup: hideEast,
      act: (b) => b.wb.sheet(0).pivot(b.pivot).clearFields(),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).pivot(b.pivot).listFields().length, 0);
        // The row order empties with the fields; East's hidden flag goes too.
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '|52']);
      },
    },
  ],
  [
    'Pivot.setFieldAxis',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).setFieldAxis(b.product, 3),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['Product|(すべて)', '|', ...BASE_GRID]),
    },
  ],
  [
    'Pivot.updateField',
    {
      // East now outsums West, so sorting ascending by the data field
      // differs from both label orders.
      setup: (b) => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, 40),
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { sortOrder: 1, sortByField: 'Sum of Amount' }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'West|37', 'East|45', '総計|82']),
    },
  ],
  [
    'Pivot.updateField',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { defaultSubtotal: false }),
      expect: (b) => assert.equal(b.wb.sheet(0).pivot(b.pivot).listFields()[b.region].defaultSubtotal, false),
    },
  ],
  [
    'Pivot.updateField',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { insertBlankRow: true }),
      expect: (b) => assert.equal(b.wb.sheet(0).pivot(b.pivot).listFields()[b.region].insertBlankRow, true),
    },
  ],
  [
    'Pivot.updateField',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { fillDownLabels: true }),
      expect: (b) => assert.equal(b.wb.sheet(0).pivot(b.pivot).listFields()[b.region].fillDownLabels, true),
    },
  ],
  [
    'Pivot.updateField',
    {
      setup: twoLevels,
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { showDetail: false }),
      expect: (b) => assert.equal(b.wb.sheet(0).pivot(b.pivot).getFieldItem(b.region, 0).showDetail, false),
    },
  ],
  [
    'Pivot.updateFieldItem',
    {
      setup: twoLevels,
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateFieldItem(b.region, 1, { showDetail: false }),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).pivot(b.pivot).getFieldItem(b.region, 1).showDetail, false);
        assert.equal(b.wb.sheet(0).pivot(b.pivot).getFieldItem(b.region, 0).showDetail, true);
      },
    },
  ],
  [
    'Pivot.updateFieldItem',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateFieldItem(b.region, 0, { caption: 'Eastern' }),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).pivot(b.pivot).getFieldItem(b.region, 0).caption, 'Eastern');
        assert.equal(b.wb.sheet(0).pivot(b.pivot).getFieldItem(b.region, 1).caption, null);
        b.wb.sheet(0).pivot(b.pivot).updateFieldItem(b.region, 0, { caption: '' });
        assert.equal(b.wb.sheet(0).pivot(b.pivot).getFieldItem(b.region, 0).caption, null);
      },
    },
  ],
  [
    'Pivot.updateOptions',
    {
      act: (b) =>
        b.wb.sheet(0).pivot(b.pivot).updateOptions({
          grandTotalCaption: 'All',
          autoFormatting: false,
          showError: true,
          errorCaption: 'n/a',
          mergeItem: true,
          indent: 3,
          pageWrap: 2,
          altText: 'Sales',
          altTextSummary: 'Sales by region\nand product',
        }),
      expect: (b) => {
        const o = b.wb.sheet(0).pivot(b.pivot).getOptions();
        assert.equal(o.grandTotalCaption, 'All');
        assert.equal(o.autoFormatting, false);
        assert.equal(o.showError, true);
        assert.equal(o.errorCaption, 'n/a');
        assert.equal(o.mergeItem, true);
        assert.equal(o.indent, 3);
        assert.equal(o.pageWrap, 2);
        assert.equal(o.altText, 'Sales');
        assert.equal(o.altTextSummary, 'Sales by region\nand product');
        assert.equal(pivotGrid(b.wb).at(-1), 'All|52');
      },
    },
  ],
  [
    'PivotCache.updateOptions',
    {
      act: (b) =>
        b.wb.pivotCache(b.cacheId).updateOptions({ saveData: false, missingItemsLimit: 0, refreshOnLoad: true }),
      expect: (b) => {
        const o = b.wb.pivotCache(b.cacheId).getOptions();
        assert.equal(o.saveData, false);
        assert.equal(o.missingItemsLimit, 0);
        assert.equal(o.refreshOnLoad, true);
      },
    },
  ],
  [
    'Pivot.updateField',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { customName: 'Area' }),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).pivot(b.pivot).listFields()[b.region].customName, 'Area');
        b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { customName: null });
        assert.equal(b.wb.sheet(0).pivot(b.pivot).listFields()[b.region].customName, null);
      },
    },
  ],
  [
    'Pivot.updateField',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { showAll: true }),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).pivot(b.pivot).listFields()[b.region].showAll, true);
        assert.equal(b.wb.sheet(0).pivot(b.pivot).listFields()[b.product].showAll, false);
        b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { showAll: false });
        assert.equal(b.wb.sheet(0).pivot(b.pivot).listFields()[b.region].showAll, false);
      },
    },
  ],
  [
    'Pivot.project',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).project(),
      expect: (b) => {
        const pending = b.wb.sheet(0).pivot(b.pivot).isProjectionPending();
        assert.equal(pending, false);
      },
    },
  ],
  [
    'PivotCache.addCalculatedField',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addCalculatedField('Net', 'Amount*2'),
      expect: (b, r) => {
        assert.equal(r, 4);
        assert.equal(b.wb.pivotCache(b.cacheId).listFields()[4].name, 'Net');
        assert.equal(b.wb.pivotCache(b.cacheId).listFields()[4].formula, '=Amount*2');
        assert.equal(b.wb.pivotCache(b.cacheId).listFields()[2].formula, null);
      },
    },
  ],
  [
    'PivotCache.setCalculatedFieldFormula',
    {
      setup: addNet,
      act: (b) => b.wb.pivotCache(b.cacheId).setCalculatedFieldFormula(4, 'Amount*3'),
      expect: (b) => assert.equal(b.wb.pivotCache(b.cacheId).listFields()[4].formula, '=Amount*3'),
    },
  ],
  [
    'PivotCache.removeCalculatedField',
    {
      setup: addNet,
      act: (b) => b.wb.pivotCache(b.cacheId).removeCalculatedField(4),
      expect: (b) => assert.equal(b.wb.pivotCache(b.cacheId).listFields().length, 4),
    },
  ],
  [
    'PivotCache.addCalculatedItem',
    {
      act: (b) => b.wb.pivotCache(b.cacheId).addCalculatedItem(b.region, 'EW', 'East+West'),
      expect: (b) => {
        assert.equal(b.wb.pivotCache(b.cacheId).listCalculatedItems().length, 1);
        const item = b.wb.pivotCache(b.cacheId).listCalculatedItems()[0];
        assert.equal(item.fieldIndex, b.region);
        assert.equal(item.name, 'EW');
        assert.equal(item.formula, '=Region[East]+Region[West]');
        assert.equal(item.scoped, false);
      },
    },
  ],
  [
    'PivotCache.setCalculatedItemFormula',
    {
      setup: addEw,
      act: (b) => b.wb.pivotCache(b.cacheId).setCalculatedItemFormula(0, 'East*2'),
      expect: (b) => assert.equal(b.wb.pivotCache(b.cacheId).listCalculatedItems()[0].formula, '=Region[East]*2'),
    },
  ],
  [
    'PivotCache.moveCalculatedItem',
    {
      setup: (b) => {
        addEw(b);
        b.wb.pivotCache(b.cacheId).addCalculatedItem(b.region, 'WE', 'West+East');
      },
      act: (b) => b.wb.pivotCache(b.cacheId).moveCalculatedItem(0, 1),
      expect: (b) => {
        assert.equal(b.wb.pivotCache(b.cacheId).listCalculatedItems()[0].name, 'WE');
        assert.equal(b.wb.pivotCache(b.cacheId).listCalculatedItems()[1].name, 'EW');
      },
    },
  ],
  [
    'PivotCache.removeCalculatedItem',
    {
      setup: addEw,
      act: (b) => b.wb.pivotCache(b.cacheId).removeCalculatedItem(0),
      expect: (b) => assert.equal(b.wb.pivotCache(b.cacheId).listCalculatedItems().length, 0),
    },
  ],
  [
    'Pivot.showDetails',
    {
      // The grand total cell is a data cell, so the drill-down lands on a new
      // sheet inserted before the pivot's sheet.
      act: (b) => b.wb.sheet(0).pivot(b.pivot).showDetails(3, 7),
      expect: (b, r) => {
        assert.equal(r, b.wb.sheet(0));
        assert.equal(r.index, 0);
        assert.equal(b.wb.listSheets().length, 2);
      },
    },
  ],
  [
    'Pivot.updateField',
    {
      setup: twoLevels,
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { subtotalTop: false }),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), [
          '行ラベル|Sum of Amount',
          'East|',
          'Apple|10',
          'Pear|5',
          'East 集計|15',
          'West|',
          'Apple|30',
          'Pear|7',
          'West 集計|37',
          '総計|52',
        ]),
    },
  ],
  [
    'Pivot.addFieldItem',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).addFieldItem(b.region, 'East', { visible: false }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
  [
    'Pivot.addFieldItem',
    {
      // Only the index-addressed form can name the blank item.
      setup: (b) => {
        b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, null);
        b.wb.pivotCache(b.cacheId).setRecordValue(0, 0, null);
      },
      act: (b) => b.wb.sheet(0).pivot(b.pivot).addFieldItem(b.region, { cacheIndex: 2 }, { visible: false }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
    },
  ],
  [
    'Pivot.clearFieldItems',
    {
      setup: hideEast,
      act: (b) => b.wb.sheet(0).pivot(b.pivot).clearFieldItems(b.region),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'Pivot.updateFieldItem',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateFieldItem(b.region, 1, { visible: false }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|15', '総計|15']),
    },
  ],
  [
    'Pivot.addFieldSubtotalFunction',
    {
      setup: twoLevels,
      act: (b) => b.wb.sheet(0).pivot(b.pivot).addFieldSubtotalFunction(b.region, 3),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), [
          '行ラベル|Sum of Amount',
          'East|10',
          'Apple|10',
          'Pear|5',
          'West|30',
          'Apple|30',
          'Pear|7',
          '総計|52',
        ]),
    },
  ],
  [
    'Pivot.clearFieldSubtotalFunctions',
    {
      setup: (b) => {
        twoLevels(b);
        b.wb.sheet(0).pivot(b.pivot).addFieldSubtotalFunction(b.region, 3);
      },
      act: (b) => b.wb.sheet(0).pivot(b.pivot).clearFieldSubtotalFunctions(b.region),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), TWO_LEVEL_GRID),
    },
  ],
  [
    'Pivot.setFieldDateGroup',
    {
      setup: dateRows,
      act: (b) => b.wb.sheet(0).pivot(b.pivot).setFieldDateGroup(b.date, { periods: 64, calendar: 1 }),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '令和6年|40', '令和7年|12', '総計|52']),
    },
  ],
  [
    'Pivot.clearFieldDateGroup',
    {
      setup: (b) => {
        dateRows(b);
        b.wb.sheet(0).pivot(b.pivot).setFieldDateGroup(b.date, { periods: 64, calendar: 0 });
      },
      act: (b) => b.wb.sheet(0).pivot(b.pivot).clearFieldDateGroup(b.date),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), DATE_SERIAL_GRID),
    },
  ],
  [
    'Pivot.updateField',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).updateField(b.amount, { numberFormat: '#,##0.00' }),
      // A field's numFmtId is only observable in the part the writer emits.
      expect: (b) => {
        expectFailure(
          () => b.wb.sheet(0).pivot(b.pivot).updateField(99, { numberFormat: '#,##0.00' }),
          b.mod.StatusCode.IndexOutOfRange,
        );
        const part = () => packagePart(b.wb.save(), 'xl/pivotTables/pivotTable1.xml');
        assert.match(part(), /<pivotField [^>]*numFmtId="4"/);
        b.wb.sheet(0).pivot(b.pivot).updateField(b.amount, { numberFormat: '0.00' });
        assert.match(part(), /<pivotField [^>]*numFmtId="2"/);
        b.wb.sheet(0).pivot(b.pivot).updateField(b.amount, { numberFormat: '' });
        assert.doesNotMatch(part(), /<pivotField [^>]*numFmtId=/);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'Pivot.setRowFieldOrder',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).setRowFieldOrder([b.product]),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'Apple|40', 'Pear|12', '総計|52']),
    },
  ],
  [
    'Pivot.setColumnFieldOrder',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).setColumnFieldOrder([b.product]),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), [
          'Sum of Amount|列ラベル||',
          '行ラベル|Apple|Pear|総計',
          'East|10|5|15',
          'West|30|7|37',
          '総計|40|12|52',
        ]),
    },
  ],
  [
    'Pivot.addDataField',
    {
      act: (b) =>
        b.wb.sheet(0).pivot(b.pivot).addDataField({ name: 'Max of Amount', fieldIndex: b.amount, aggregation: 3 }),
      expect: (b, r) => {
        assert.equal(r, 1);
        assert.deepEqual(pivotGrid(b.wb), [
          '行ラベル|Sum of Amount|Max of Amount',
          'East|15|10',
          'West|37|30',
          '総計|52|30',
        ]);
      },
    },
  ],
  [
    'Pivot.clearDataFields',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).clearDataFields(),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル', 'East', 'West', '総計']),
    },
  ],
  [
    'Pivot.setDataField',
    {
      act: (b) =>
        b.wb.sheet(0).pivot(b.pivot).setDataField(0, {
          name: 'Count of Amount',
          fieldIndex: b.amount,
          aggregation: 1,
          numberFormat: '0.00',
        }),
      expect: (b) => {
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Count of Amount', 'East|2', 'West|2', '総計|4']);
        const counts = b.wb
          .sheet(0)
          .pivot(b.pivot)
          .getLayout()
          .cells.filter((c) => typeof c.value === 'number');
        assert.deepEqual(
          counts.map((c) => c.numberFormat),
          ['0.00', '0.00', '0.00'],
        );
      },
    },
  ],
  [
    'Pivot.addFilter',
    {
      act: (b) => b.wb.sheet(0).pivot(b.pivot).addFilter(GREATER_THAN_20),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
  [
    'Pivot.clearFilters',
    {
      setup: (b) => b.wb.sheet(0).pivot(b.pivot).addFilter(GREATER_THAN_20),
      act: (b) => b.wb.sheet(0).pivot(b.pivot).clearFilters(),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'Pivot.removeFilter',
    {
      // Each filter alone keeps a different region, so removing the wrong
      // index leaves the wrong one.
      setup: (b) => {
        b.wb.sheet(0).pivot(b.pivot).addFilter(GREATER_THAN_20);
        b.wb.sheet(0).pivot(b.pivot).addFilter(BETWEEN_10_20);
      },
      act: (b) => b.wb.sheet(0).pivot(b.pivot).removeFilter(1),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
];

// The `{ name -> return type }` of the single-line methods an interface of the declaration file declares.
function declaredMethods(dts, name) {
  const start = dts.search(new RegExp(`^export interface ${name}(?: extends \\w+)? \\{$`, 'm'));
  assert.ok(start >= 0, `no interface ${name}`);
  const block = dts.slice(start, dts.indexOf('\n}\n', start));
  return [...block.matchAll(/^ {2}(\w+)\([^)]*\): ([\w[\]]+);/gm)].map((m) => [m[1], m[2]]);
}

test('the pivot mutator table covers every declared pivot mutator', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
  const mutating = new Set(['void', 'number', 'Sheet', 'Pivot', 'PivotCache']);
  const declared = [];
  for (const owner of ['Workbook', 'Sheet']) {
    for (const [name, ret] of declaredMethods(dts, owner)) {
      if (/^(?:[a-z]+Pivot|pivot)\w*$/.test(name) && !['pivot', 'pivotCache'].includes(name) && mutating.has(ret)) {
        declared.push(name);
      }
    }
  }
  for (const owner of ['Pivot', 'PivotCache']) {
    for (const [name, ret] of declaredMethods(dts, owner)) {
      if (mutating.has(ret) && !name.endsWith('Count')) declared.push(`${owner}.${name}`);
    }
  }
  declared.splice(0, declared.length, ...declared.filter((name) => !name.endsWith('Count')));
  assert.ok(declared.length > 40, `expected the pivot mutator family, found ${declared.length}`);
  const covered = PIVOT_MUTATORS.map(([name]) => name);
  assert.deepEqual(
    declared.filter((name) => !covered.includes(name)),
    [],
    'pivot mutators without a row',
  );
  assert.deepEqual(
    covered.filter((name) => !declared.includes(name)),
    [],
    'rows for methods that are not declared pivot mutators',
  );
});

test('pivot mutators forward their arguments to the C ABI', async () => {
  const mod = await getModule();
  const base = buildPivot(mod);
  try {
    assert.deepEqual(pivotGrid(base.wb), BASE_GRID);
  } finally {
    base.wb.dispose();
  }
  for (const [name, row] of PIVOT_MUTATORS) {
    const b = buildPivot(mod);
    try {
      row.setup?.(b);
      const r = row.act(b);
      row.after?.(b);
      row.expect(b, r);
    } catch (err) {
      err.message = `${name}: ${err.message}`;
      throw err;
    } finally {
      b.wb.dispose();
    }
  }
});

test('the cache value writers dispatch on every value kind', async () => {
  const mod = await getModule();
  for (const [writer, cases] of [
    ['addPivotCacheFieldSharedItem', SHARED_ITEM_CASES],
    ['setPivotCacheRecordValue', RECORD_VALUE_CASES],
  ]) {
    for (const row of cases) {
      const b = buildPivot(mod);
      try {
        row.act(b);
        if (writer === 'addPivotCacheFieldSharedItem' && !row.noPointing) pointAtNewItem(b);
        row.expect(b);
      } catch (err) {
        err.message = `${writer} (${row.label}): ${err.message}`;
        throw err;
      } finally {
        b.wb.dispose();
      }
    }
  }
});

test('a date option on the cache value writers is refused as a TypeError and writes nothing', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const items = b.wb.pivotCache(b.cacheId).listFieldSharedItems(0).length;
    for (const value of ['x', true, null, ERROR_NA]) {
      assert.throws(() => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, value, { date: true }), TypeError);
      assert.throws(() => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, value, { date: true }), TypeError);
    }
    assert.equal(b.wb.pivotCache(b.cacheId).listFieldSharedItems(0).length, items);
  } finally {
    b.wb.dispose();
  }
});

test('manual pivot sort order keeps item order and rejects bad arguments', async () => {
  const mod = await getModule();
  const { PivotSortOrder } = await loadStagedModule();
  const b = buildPivot(mod);
  try {
    b.wb.sheet(0).pivot(b.pivot).clearFieldItems(b.region);
    b.wb.sheet(0).pivot(b.pivot).addFieldItem(b.region, 'West', { visible: true });
    b.wb.sheet(0).pivot(b.pivot).addFieldItem(b.region, 'East', { visible: true });
    b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { sortOrder: PivotSortOrder.Manual });
    assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'West|37', 'East|15', '総計|52']);
    expectFailure(() => b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { sortOrder: 3 }));
    expectFailure(() =>
      b.wb
        .sheet(0)
        .pivot(b.pivot)
        .updateField(b.region, { sortOrder: PivotSortOrder.Manual, sortByField: 'Sum of Amount' }),
    );
    b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { sortOrder: PivotSortOrder.Ascending });
    assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('pivot filters and subtotal functions survive save and reload', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    b.wb.sheet(0).pivot(b.pivot).addFilter(GREATER_THAN_20);
    b.wb.sheet(0).pivot(b.pivot).addFieldSubtotalFunction(b.region, 2);
    b.wb.sheet(0).pivot(b.pivot).addFieldSubtotalFunction(b.region, 0);
    const region = b.wb.sheet(0).pivot(b.pivot).listFields()[b.region];
    assert.equal(region.defaultSubtotal, false);
    assert.deepEqual(region.subtotalFunctions, [0, 2]);
    assert.equal(b.wb.sheet(0).pivot(b.pivot).listFields()[b.amount].defaultSubtotal, true);
    assert.deepEqual(b.wb.sheet(0).pivot(b.pivot).listFields()[b.amount].subtotalFunctions, []);

    const saved = b.wb.save();
    assert.match(
      packagePart(saved, 'xl/pivotTables/pivotTable1.xml'),
      /<filters count="1"><filter fld="0" type="valueGreaterThan" evalOrder="-1" id="1" iMeasureFld="0">/,
    );
    const back = mod.Workbook.load(saved);
    try {
      back.setProfileId('win-365-ja_JP');
      assert.equal(back.sheet(0).pivot(0).listFilters().length, 1);
      const filter = back.sheet(0).pivot(0).listFilters()[0];
      assert.equal(filter.type, 1);
      assert.equal(filter.fieldName, 'Region');
      assert.equal(filter.value, 20);
      assert.deepEqual(pivotGrid(back), WEST_ONLY);
      const field = back.sheet(0).pivot(0).listFields()[0];
      assert.equal(field.defaultSubtotal, false);
      assert.deepEqual(field.subtotalFunctions, [0, 2]);
      back.sheet(0).pivot(0).clearFilters();
      assert.doesNotMatch(packagePart(back.save(), 'xl/pivotTables/pivotTable1.xml'), /<filters/);
    } finally {
      back.dispose();
    }
  } finally {
    b.wb.dispose();
  }
});

// ---- Failures of the fallible accessors --------------------------------
//
// Every accessor backed by a status-returning C ABI call throws on failure,
// so a rejected argument or a released handle cannot read as a
// legitimate zero count, empty name, or unspilled cell. Each row is
// [method, args that succeed, args the engine rejects (or null when only
// a released handle can fail it)]. A `Sheet` method takes the sheet index as
// its first element here; `callAccessor` routes it through `wb.sheet()`.
const STATUS_ACCESSORS = [
  ['listSheets', [], null],
  ['getMemoryUsage', [], null],
  ['listCells', [0], [99]],
  ['getCellXfIndex', [0, 0, 0], [99, 0, 0]],
  [
    'recalcViewport',
    [0, { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }],
    [99, { firstRow: 0, firstCol: 0, lastRow: 1, lastCol: 1 }],
  ],
  ['listCfRules', [0], [99]],
  ['listValidations', [0], [99]],
  ['listProtectedRanges', [0], [99]],
  ['listSparklineGroups', [0], [99]],
  ['getCellPhoneticRuns', [0, 0, 0], [99, 0, 0]],
  ['listExternalLinks', [], null],
  ['listDefinedNames', [], null],
  ['listTables', [0], [99]],
  ['listPassthroughs', [], null],
  ['listPivots', [0], [99]],
  ['listPivotCaches', [], null],
  ['PivotCache.listFields', null, [9999]],
  ['PivotCache.listFieldSharedItems', null, [9999, 0]],
  ['PivotCache.getRecordCount', null, [9999]],
  ['Pivot.listFields', null, [0, 99]],
  ['Pivot.listFieldItems', null, [0, 99, 0]],
  ['Pivot.getProperties', null, [0, 99]],
  ['Pivot.getReportLayout', null, [0, 99]],
  ['Pivot.listDataFields', null, [0, 99]],
  ['Pivot.listFilters', null, [0, 99]],
  ['PivotCache.listCalculatedItems', null, [9999]],
  ['listFonts', [], null],
  ['listFills', [], null],
  ['listBorders', [], null],
  ['listCellXfs', [], null],
  ['listDxfs', [], null],
  ['listCellStyles', [], null],
  ['listTableStyles', [], null],
  ['listCellStyleXfs', [], null],
  ['getCalcMode', [], null],
  ['getFullPrecision', [], null],
  ['getCalcOnSave', [], null],
  ['getCommentVisible', null, [99, 0, 0]],
  ['getRefMode', [], null],
  ['getDate1904', [], null],
  ['getProfileId', [], null],
  ['getPinnedNow', [], null],
  ['getSpillInfo', [0, 0, 0], [99, 0, 0]],
  ['getArrayFormula', [0, 0, 0], [99, 0, 0]],
  ['getColumnWidthPt', [0, 0, 0], [99, 0, 0]],
  ['getRowHeightPt', [0, 0], [99, 0]],
  ['convertColumnCharsToPt', [0, 8, 0], [99, 8, 0]],
  ['convertColumnPtToChars', [0, 10, 0], [99, 10, 0]],
  ['getPrecedents', [0, 0, 0, 1], [99, 0, 0, 1]],
  ['getDependents', [0, 0, 0, 1], [99, 0, 0, 1]],
];

// The stateless functions live on the module; each row is [function, args that succeed, args the engine rejects].
const MODULE_FUNCTIONS = [
  ['compareText', ['a', 'b', 0, 'win-365-en_US'], ['a', 'b', 0, 'nope']],
  ['localizeFunctionName', ['SUM', 'win-365-en_US'], ['NOPE_XYZ', 'win-365-en_US']],
  ['canonicalizeFunctionName', ['SUM', 'win-365-en_US'], ['NOPE_XYZ', 'win-365-en_US']],
  ['listFunctionNames', [], null],
  ['listLocaleErrorNames', ['win-365-en_US'], ['nope']],
  [
    'convertCivilToSerial',
    [{ year: 2026, month: 1, day: 1, hour: 0, minute: 0, second: 0 }],
    [{ year: 2026, month: 13, day: 1, hour: 0, minute: 0, second: 0 }],
  ],
];

// Calls a STATUS_ACCESSORS row on the workbook, or on the sheet its first argument names.
function callAccessor(wb, name, args) {
  if (name.startsWith('Pivot.'))
    return wb
      .sheet(args[0])
      .pivot(args[1])
      [name.slice(6)](...args.slice(2));
  if (name.startsWith('PivotCache.')) return wb.pivotCache(args[0])[name.slice(11)](...args.slice(1));
  if (name in wb) return wb[name](...args);
  if (name in wb.ooxml) return wb.ooxml[name](...args);
  const sheet = wb.sheet(args[0]);
  return name in sheet ? sheet[name](...args.slice(1)) : sheet.ooxml[name](...args.slice(1));
}

test('every numeric accessor is covered by the status table', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
  const handles = new Set(['Table', 'Pivot', 'PivotCache', 'Drawing', 'Picture', 'Chart', 'Shape', 'Slicer']);
  const declared = [];
  for (const part of dts.split(/^(?=export interface )/m)) {
    const owner = part.match(/^export interface (\w+)/)?.[1];
    for (const m of part.matchAll(/^ {2}([A-Za-z]\w*)\([^)]*\): (?:number|CalcMode|RefMode|PivotReportLayout);/gm)) {
      if (!/^(add|copy|restore|activeSheet)/.test(m[1])) declared.push(handles.has(owner) ? `${owner}.${m[1]}` : m[1]);
    }
  }
  assert.ok(declared.length > 10, `expected the numeric accessors, found ${declared.length}`);
  const covered = new Set([...STATUS_ACCESSORS, ...MODULE_FUNCTIONS].map(([name]) => name));
  assert.deepEqual(
    declared.filter((name) => !covered.has(name)),
    [],
    'numeric accessors missing from STATUS_ACCESSORS',
  );
});

test('fallible accessors succeed, then throw on a rejected argument and on a released handle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  for (const [name, okArgs, badArgs] of STATUS_ACCESSORS) {
    if (okArgs !== null) callAccessor(wb, name, okArgs);
    if (badArgs !== null) expectFailure(() => callAccessor(wb, name, badArgs));
  }
  assert.equal(wb.listSheets().length, 1);
  assert.equal(wb.getProfileId(), 'win-365-en_US');
  for (const [name, okArgs, badArgs] of MODULE_FUNCTIONS) {
    mod[name](...okArgs);
    if (badArgs !== null) expectFailure(() => mod[name](...badArgs));
  }
  assert.equal(mod.localizeFunctionName('SUM', 'win-365-ja_JP'), 'SUM');
  expectFailure(() => mod.localizeFunctionName('SUM', 'nope'));

  // `dispose()` releases the handle and keeps the wrapper callable.
  wb.dispose();
  wb.dispose();
  const dead = wb;
  assert.equal(dead.disposed, true);
  // Every accessor, including the ones no argument can fail, reports the
  // released handle instead of a default reading.
  for (const [name, okArgs, badArgs] of STATUS_ACCESSORS) {
    expectFailure(() => callAccessor(dead, name, okArgs ?? badArgs), mod.StatusCode.BindingInvalidHandle);
  }
  dead.dispose();
});

test('iterative progress callbacks are per Workbook, not per module', async () => {
  const Module = await getModule();
  const first = makeIterativeWorkbook(Module);
  const second = makeIterativeWorkbook(Module);
  try {
    let firstCalls = 0;
    let secondCalls = 0;
    first.setIterativeProgress(() => {
      firstCalls += 1;
      return true;
    });
    // Installing on `second` must not displace the registration on `first`.
    second.setIterativeProgress(() => {
      secondCalls += 1;
      return true;
    });

    first.recalc();
    assert.ok(firstCalls > 0, 'the first workbook kept its own callback');
    assert.equal(secondCalls, 0);

    firstCalls = 0;
    second.recalc();
    assert.ok(secondCalls > 0);
    assert.equal(firstCalls, 0);
  } finally {
    first.setIterativeProgress(null);
    second.setIterativeProgress(null);
    first.dispose();
    second.dispose();
  }
});

test('setPivotFieldDateGroup defaults its omitted fields', async () => {
  const mod = await getModule();
  const explicit = buildPivot(mod);
  const defaulted = buildPivot(mod);
  try {
    dateRows(explicit);
    dateRows(defaulted);
    const days = 8; // PivotDatePeriod.Days
    explicit.wb.sheet(0).pivot(explicit.pivot).setFieldDateGroup(explicit.date, {
      periods: days,
      calendar: 0,
      intervalDays: 1,
    });
    defaulted.wb.sheet(0).pivot(defaulted.pivot).setFieldDateGroup(defaulted.date, { periods: days });
    const field = defaulted.wb.sheet(0).pivot(defaulted.pivot).listFields()[defaulted.date];
    assert.equal(field.datePeriods, days);
    assert.deepEqual(field, explicit.wb.sheet(0).pivot(explicit.pivot).listFields()[explicit.date]);
    assert.deepEqual(pivotGrid(defaulted.wb), pivotGrid(explicit.wb));
  } finally {
    explicit.wb.dispose();
    defaulted.wb.dispose();
  }
});

test('pivot cache date shared items report as dates and reject dates the engine cannot store', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, new Date(2024, 0, 1));
    b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, new Date(2024, 1, 1, 12));
    b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, 7);
    assert.deepEqual(
      b.wb
        .pivotCache(b.cacheId)
        .listFieldSharedItems(3)
        .map((item) => item.isDate),
      [true, true, false],
    );
    expectFailure(() => b.wb.pivotCache(9999).listFieldSharedItems(3));
    const before = new Date(1899, 11, 30);
    expectFailure(() => b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, before));
    expectFailure(() => b.wb.pivotCache(b.cacheId).setRecordValue(0, 3, before));
    const invalid = new Date(Number.NaN);
    assert.throws(() => b.wb.pivotCache(b.cacheId).addFieldSharedItem(3, invalid), RangeError);
    assert.throws(() => b.wb.pivotCache(b.cacheId).setRecordValue(0, 3, invalid), RangeError);
  } finally {
    b.wb.dispose();
  }
});

test('a pivot cache shared item reads back by index', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const items = b.wb.pivotCache(b.cacheId).listFieldSharedItems(0);
    assert.equal(items[0].value, 'East');
    assert.equal(items[1].value, 'West');
    assert.equal(items.length, 2);
    expectFailure(() => b.wb.pivotCache(b.cacheId).listFieldSharedItems(99));
  } finally {
    b.wb.dispose();
  }
});

test('a pivot cache error value without errorCode throws and changes nothing', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const items = b.wb.pivotCache(b.cacheId).listFieldSharedItems(0).length;
    assert.throws(() => b.wb.pivotCache(b.cacheId).addFieldSharedItem(0, { kind: 4 }), TypeError);
    assert.equal(b.wb.pivotCache(b.cacheId).listFieldSharedItems(0).length, items);
    assert.throws(() => b.wb.pivotCache(b.cacheId).setRecordValue(0, 2, { kind: 4 }), TypeError);
    assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('new pivot calls reject bad arguments and map the Values level', async () => {
  const mod = await getModule();
  const { PIVOT_VALUES_FIELD } = await loadStagedModule();
  assert.equal(PIVOT_VALUES_FIELD, -2);
  const b = buildPivot(mod);
  try {
    const { wb, pivot, region, product } = b;
    for (const call of [
      () => wb.sheet(0).pivot(pivot).updateField(99, { defaultSubtotal: true }),
      () => wb.sheet(0).pivot(pivot).updateField(99, { insertBlankRow: true }),
      () => wb.sheet(0).pivot(pivot).updateField(99, { fillDownLabels: true }),
      () => wb.sheet(0).pivot(pivot).updateField(99, { showDetail: true }),
      () => wb.sheet(0).pivot(pivot).updateFieldItem(region, 99, { showDetail: true }),
      () => wb.sheet(0).pivot(pivot).updateFieldItem(region, 99, { caption: 'x' }),
      () => wb.sheet(0).pivot(9).updateOptions({ indent: 1 }),
      () => wb.pivotCache(9999).updateOptions({ saveData: true }),
      () => wb.sheet(0).pivot(pivot).updateField(99, { customName: 'x' }),
      () => wb.sheet(0).pivot(pivot).updateField(99, { showAll: true }),
      () => wb.sheet(0).pivot(9).project(),
      () => wb.sheet(0).pivot(pivot).showDetails(99, 99),
      () => wb.sheet(0).pivot(9).getOptions(),
      () => wb.pivotCache(9999).getOptions(),
      () => wb.sheet(0).pivot(9).isProjectionPending(),
      () => wb.pivotCache(9999).listFields(),
      () => wb.pivotCache(9999).listCalculatedItems(),
      () => wb.sheet(0).pivot(9).getDataFieldDefaultName(0, 0),
      () => wb.sheet(0).chart(12345).rename('x'),
    ]) {
      expectFailure(call);
    }
    // The Values level needs at least two data fields.
    wb.sheet(0).pivot(pivot).addDataField({ name: 'Count of Date', fieldIndex: b.date, aggregation: 1 });
    wb.sheet(0).pivot(pivot).setRowFieldOrder([region, PIVOT_VALUES_FIELD, product]);
    assert.deepEqual(wb.sheet(0).pivot(pivot).getRowFieldOrder(), [region, PIVOT_VALUES_FIELD, product]);
    expectFailure(() => wb.sheet(0).pivot(pivot).setColumnFieldOrder([-3]));
  } finally {
    b.wb.dispose();
  }
});

test('pivotSetOptions merges onto the current options and rejects bad values without changing anything', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const { wb, pivot } = b;
    const before = wb.sheet(0).pivot(pivot).getOptions();
    wb.sheet(0).pivot(pivot).updateOptions({});
    assert.deepEqual(wb.sheet(0).pivot(pivot).getOptions(), before);
    wb.sheet(0).pivot(pivot).updateOptions({ indent: 4 });
    const after = wb.sheet(0).pivot(pivot).getOptions();
    assert.equal(after.indent, 4);
    assert.deepEqual({ ...after, indent: before.indent }, before);

    // A caption of exactly 255 UTF-16 units is kept; one more is not. A surrogate pair counts as two.
    wb.sheet(0)
      .pivot(pivot)
      .updateOptions({ errorCaption: 'e'.repeat(255) });
    assert.equal(wb.sheet(0).pivot(pivot).getOptions().errorCaption.length, 255);
    wb.sheet(0)
      .pivot(pivot)
      .updateOptions({ errorCaption: '\u{1F600}'.repeat(127) });
    const kept = wb.sheet(0).pivot(pivot).getOptions();
    for (const bad of [
      { errorCaption: 'e'.repeat(256) },
      { missingCaption: '\u{1F600}'.repeat(128) },
      { grandTotalCaption: 'g'.repeat(256) },
      { altText: 'a'.repeat(256) },
      { indent: 127 },
      { pageWrap: 256 },
    ]) {
      expectFailure(() => wb.sheet(0).pivot(pivot).updateOptions(bad));
      assert.deepEqual(wb.sheet(0).pivot(pivot).getOptions(), kept, JSON.stringify(bad));
    }
    // A value of the wrong kind throws before anything is written.
    for (const bad of [{ indent: '3' }, { indent: 1.5 }, { indent: -1 }, { pageWrap: Number.NaN }]) {
      assert.throws(() => wb.sheet(0).pivot(pivot).updateOptions(bad), wrongKind, JSON.stringify(bad));
      assert.deepEqual(wb.sheet(0).pivot(pivot).getOptions(), kept, JSON.stringify(bad));
    }
    wb.sheet(0).pivot(pivot).updateOptions({ errorCaption: null });
    assert.equal(wb.sheet(0).pivot(pivot).getOptions().errorCaption, null);
    wb.sheet(0).pivot(pivot).updateOptions({ errorCaption: '' });
    assert.equal(wb.sheet(0).pivot(pivot).getOptions().errorCaption, null);
    assert.throws(() => wb.sheet(0).pivot(pivot).updateOptions(), TypeError);
    assert.deepEqual(pivotGrid(wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('pivotCacheSetOptions merges, bounds the missing-items limit and survives save and reload', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const { wb, cacheId } = b;
    const before = wb.pivotCache(cacheId).getOptions();
    wb.pivotCache(cacheId).updateOptions({ saveData: !before.saveData });
    assert.deepEqual(wb.pivotCache(cacheId).getOptions(), {
      ...before,
      saveData: !before.saveData,
    });
    for (const limit of [null, 0, 1, 32500, 1048576]) {
      wb.pivotCache(cacheId).updateOptions({ missingItemsLimit: limit });
      assert.equal(wb.pivotCache(cacheId).getOptions().missingItemsLimit, limit);
    }
    const kept = wb.pivotCache(cacheId).getOptions();
    expectFailure(() => wb.pivotCache(cacheId).updateOptions({ missingItemsLimit: 1048577 }));
    assert.deepEqual(wb.pivotCache(cacheId).getOptions(), kept);
    for (const bad of [-1, -2, 1.5, Number.NaN, '7']) {
      assert.throws(() => wb.pivotCache(cacheId).updateOptions({ missingItemsLimit: bad }), wrongKind, String(bad));
      assert.deepEqual(wb.pivotCache(cacheId).getOptions(), kept, String(bad));
    }
    assert.throws(() => wb.pivotCache(cacheId).updateOptions(), TypeError);

    wb.pivotCache(cacheId).updateOptions({ saveData: true, missingItemsLimit: 5, refreshOnLoad: true });
    const saved = wb.save();
    const back = mod.Workbook.load(saved);
    try {
      const o = back.pivotCache(back.listPivotCaches()[0]).getOptions();
      assert.deepEqual(
        { saveData: o.saveData, missingItemsLimit: o.missingItemsLimit, refreshOnLoad: o.refreshOnLoad },
        {
          saveData: true,
          missingItemsLimit: 5,
          refreshOnLoad: true,
        },
      );
    } finally {
      back.dispose();
    }
  } finally {
    b.wb.dispose();
  }
});

test('pivot options survive save and reload', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    b.wb
      .sheet(0)
      .pivot(b.pivot)
      .updateOptions({ grandTotalCaption: 'All', altText: 'Sales', indent: 2, mergeItem: true });
    b.wb.sheet(0).pivot(b.pivot).updateField(b.region, { showAll: true });
    b.wb.sheet(0).pivot(b.pivot).updateField(b.product, { customName: 'Item' });
    const saved = b.wb.save();
    const back = mod.Workbook.load(saved);
    try {
      const o = back.sheet(0).pivot(0).getOptions();
      assert.equal(o.grandTotalCaption, 'All');
      assert.equal(o.altText, 'Sales');
      assert.equal(o.indent, 2);
      assert.equal(o.mergeItem, true);
      assert.equal(back.sheet(0).pivot(0).listFields()[b.region].showAll, true);
      assert.equal(back.sheet(0).pivot(0).listFields()[b.product].customName, 'Item');
    } finally {
      back.dispose();
    }
  } finally {
    b.wb.dispose();
  }
});

test('pivotFieldSetCustomName refuses a name another field or data field holds', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const { wb, pivot, region } = b;
    wb.sheet(0).pivot(pivot).updateField(region, { customName: 'Region' });
    expectFailure(() => wb.sheet(0).pivot(pivot).updateField(region, { customName: 'Product' }));
    expectFailure(() => wb.sheet(0).pivot(pivot).updateField(region, { customName: 'Sum of Amount' }));
    expectFailure(() =>
      wb
        .sheet(0)
        .pivot(pivot)
        .updateField(region, { customName: 'n'.repeat(256) }),
    );
    expectFailure(() => wb.sheet(0).pivot(pivot).updateField(4096, { customName: 'x' }));
    assert.throws(() => wb.sheet(0).pivot(pivot).updateField(region), TypeError);
    assert.throws(() => wb.sheet(0).pivot(pivot).updateField(), TypeError);
    assert.equal(wb.sheet(0).pivot(pivot).listFields()[region].customName, 'Region');
  } finally {
    b.wb.dispose();
  }
});

test('pivotDataFieldDefaultName numbers a repeated source field and rejects bad arguments', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const { wb, pivot, amount } = b;
    const first = wb.sheet(0).pivot(pivot).getDataFieldDefaultName(amount, 0);
    assert.equal(first, '\u5408\u8a08 / Amount');
    assert.equal(wb.sheet(0).pivot(pivot).listDataFields().length, 1, 'reading a name adds nothing');
    const added = wb.sheet(0).pivot(pivot).addDataField({ fieldIndex: amount, aggregation: 0 });
    assert.equal(wb.sheet(0).pivot(pivot).listDataFields()[added].name, first);
    // The same source field again gets the numbered name.
    assert.equal(wb.sheet(0).pivot(pivot).getDataFieldDefaultName(amount, 0), '\u5408\u8a08 / Amount2');
    // Leaving out the entry that holds the name frees it, as a replacement would.
    assert.equal(wb.sheet(0).pivot(pivot).getDataFieldDefaultName(amount, 0, added), first);
    expectFailure(() => wb.sheet(0).pivot(pivot).getDataFieldDefaultName(99, 0));
    expectFailure(() => wb.sheet(0).pivot(pivot).getDataFieldDefaultName(amount, 99));
    expectFailure(() => wb.sheet(0).pivot(pivot).getDataFieldDefaultName(amount, 0, 99));
    expectFailure(() => wb.sheet(0).pivot(9).getDataFieldDefaultName(amount, 0));
    assert.throws(() => wb.sheet(0).pivot(pivot).getDataFieldDefaultName(amount), TypeError);
  } finally {
    b.wb.dispose();
  }
});

test('calculated fields and items reject bad names and formulas and change nothing', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const { wb, cacheId, region } = b;
    const fields = wb.pivotCache(cacheId).listFields().length;
    for (const [name, formula] of [
      ['', 'Amount*2'],
      ['Net', ''],
      ['All', 'Amount*2'],
      ['Blank', 'Amount*2'],
      ['N[et]', 'Amount*2'],
      ['Amount', 'Amount*2'],
      ['n'.repeat(256), 'Amount*2'],
      ['Net', 'Amount*'],
      ['Net', 'Missing*2'],
      ['Net', 'A1*2'],
      ['Net', `Amount+${'1+'.repeat(4100)}1`],
    ]) {
      expectFailure(() => wb.pivotCache(cacheId).addCalculatedField(name, formula));
    }
    assert.equal(wb.pivotCache(cacheId).listFields().length, fields);
    expectFailure(() => wb.pivotCache(9999).addCalculatedField('Net', 'Amount*2'));
    expectFailure(() => wb.pivotCache(cacheId).setCalculatedFieldFormula(2, 'Amount*2'));
    expectFailure(() => wb.pivotCache(cacheId).removeCalculatedField(2));
    assert.throws(() => wb.pivotCache(cacheId).addCalculatedField('Net'), TypeError);
    assert.throws(() => wb.pivotCache(cacheId).addCalculatedItem(region, 'EW'), TypeError);

    for (const [field, name, formula] of [
      [region, '', 'East+West'],
      [region, 'EW', ''],
      [region, 'All', 'East+West'],
      [region, 'East', 'East+West'],
      [region, 'EW', 'East+'],
      [region, 'EW', 'Nowhere+East'],
      [4096, 'EW', 'East+West'],
    ]) {
      expectFailure(() => wb.pivotCache(cacheId).addCalculatedItem(field, name, formula));
    }
    assert.equal(wb.pivotCache(cacheId).listCalculatedItems().length, 0);
    assert.deepEqual(wb.pivotCache(cacheId).listCalculatedItems(), []);
    expectFailure(() => wb.pivotCache(cacheId).setCalculatedItemFormula(0, 'East'));
    expectFailure(() => wb.pivotCache(cacheId).moveCalculatedItem(0, 0));
    expectFailure(() => wb.pivotCache(cacheId).removeCalculatedItem(0));
    assert.deepEqual(pivotGrid(wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('a calculated field can be placed as a data field and shows its formula over the sums', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const { wb, pivot, cacheId } = b;
    const net = wb.pivotCache(cacheId).addCalculatedField('Net', 'Amount*2');
    wb.sheet(0).pivot(pivot).addDataField({ name: 'Net', fieldIndex: net, aggregation: 0 });
    assert.equal(pivotGrid(wb).at(-1), '\u7dcf\u8a08|52|104');
    // A removal takes the data field with it.
    wb.pivotCache(cacheId).removeCalculatedField(net);
    assert.equal(wb.sheet(0).pivot(pivot).listDataFields().length, 1);
    assert.deepEqual(pivotGrid(wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('pivot filters take the full type range and an upper text bound', async () => {
  const mod = await getModule();
  const { PivotFilterKind } = await loadStagedModule();
  assert.equal(Object.keys(PivotFilterKind).length, 68);
  assert.equal(PivotFilterKind.ValueEqual, 6);
  assert.equal(PivotFilterKind.LabelBetween, 22);
  assert.equal(PivotFilterKind.M12, 67);
  const b = buildPivot(mod);
  try {
    const { wb, pivot } = b;
    const between = {
      axis: 0,
      fieldName: 'Region',
      type: PivotFilterKind.LabelBetween,
      value: 'A',
      valueHigh: 'F',
    };
    // Nothing the upper bound's rules reject is added.
    for (const bad of [
      { ...between, valueHigh: undefined },
      { ...between, type: PivotFilterKind.LabelContains },
      { ...between, type: PivotFilterKind.ValueEqual, value: 15 },
      { ...between, type: 68 },
    ]) {
      expectFailure(() => wb.sheet(0).pivot(pivot).addFilter(bad));
    }
    assert.equal(wb.sheet(0).pivot(pivot).listFilters().length, 0);
    assert.deepEqual(pivotGrid(wb), BASE_GRID);

    wb.sheet(0).pivot(pivot).addFilter(between);
    const at = wb.sheet(0).pivot(pivot).listFilters()[0];
    assert.equal(at.type, PivotFilterKind.LabelBetween);
    assert.equal(at.value, 'A');
    assert.equal(at.valueHigh, 'F');
    assert.deepEqual(pivotGrid(wb), ['\u884c\u30e9\u30d9\u30eb|Sum of Amount', 'East|15', '\u7dcf\u8a08|15']);

    wb.sheet(0).pivot(pivot).clearFilters();
    wb.sheet(0).pivot(pivot).addFilter({
      axis: 0,
      fieldName: 'Region',
      type: PivotFilterKind.ValueEqual,
      value: 15,
    });
    assert.equal(wb.sheet(0).pivot(pivot).listFilters()[0].valueHigh, null);
    assert.deepEqual(pivotGrid(wb), ['\u884c\u30e9\u30d9\u30eb|Sum of Amount', 'East|15', '\u7dcf\u8a08|15']);
  } finally {
    b.wb.dispose();
  }
});

test('setWorksheetSource(null) clears the source and an empty source is still a present one', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    const cache = wb.addPivotCache(0);
    cache.setWorksheetSource({ range: 'A1:B3', sheet: 'Sheet1' });
    assert.deepEqual(JSON.parse(JSON.stringify(cache.getWorksheetSource())), {
      range: { sheet: 0, rect: { firstRow: 0, firstCol: 0, lastRow: 2, lastCol: 1 } },
      name: null,
    });
    cache.setWorksheetSource(null);
    assert.equal(cache.getWorksheetSource(), null);
    cache.setWorksheetSource({});
    assert.deepEqual(JSON.parse(JSON.stringify(cache.getWorksheetSource())), { range: null, name: null });
  } finally {
    wb.dispose();
  }
});
