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
  wb.setExcelProfileId('win-365-ja_JP');
  const cacheId = wb.addPivotCache(0);
  wb.setPivotCacheWorksheetSource(cacheId, { present: true, ref: 'A1:D5', sheet: 'Sheet1' });
  for (const name of ['Region', 'Product', 'Amount', 'Date']) wb.addPivotCacheField(cacheId, name);
  for (const item of ['East', 'West']) wb.addPivotCacheFieldSharedItem(cacheId, 0, item);
  for (const item of ['Apple', 'Pear']) wb.addPivotCacheFieldSharedItem(cacheId, 1, item);
  for (const values of [
    [0, 0, 10, 45292],
    [1, 0, 30, 45323],
    [0, 1, 5, 45658],
    [1, 1, 7, 45689],
  ]) {
    const rec = wb.addPivotCacheRecord(cacheId);
    for (const [field, v] of values.entries()) wb.setPivotCacheRecordValue(cacheId, rec, field, v);
  }
  const pivot = wb.sheet(0).addPivot({ name: 'Pivot1', cacheId: cacheId, anchorRow: 0, anchorCol: 6 });
  const addField = (spec) => wb.sheet(0).addPivotField(pivot, spec);
  const region = addField({ sourceName: 'Region', axis: 0 });
  wb.sheet(0).addPivotFieldItem(pivot, region, 'East', { visible: true });
  wb.sheet(0).addPivotFieldItem(pivot, region, 'West', { visible: true });
  const product = addField({ sourceName: 'Product', axis: 2 });
  wb.sheet(0).addPivotFieldItem(pivot, product, 'Apple', { visible: true });
  wb.sheet(0).addPivotFieldItem(pivot, product, 'Pear', { visible: true });
  const amount = addField({ sourceName: 'Amount', axis: 2 });
  const date = addField({ sourceName: 'Date', axis: 2 });
  wb.sheet(0).setPivotRowFieldOrder(pivot, [region]);
  wb.sheet(0).addPivotDataField(pivot, { name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 });
  return { mod, wb, cacheId, pivot, region, product, amount, date };
}

// The projected grid as one `|`-joined string per sheet row.
function pivotGrid(wb, pivot = 0) {
  const layout = wb.sheet(0).getPivotLayout(pivot);
  const rows = new Map();
  for (const cell of layout.cells) {
    const v = cell.value;
    const text = v.kind === 1 ? String(v.number) : v.kind === 3 ? v.text : v.kind === 4 ? `#${v.errorCode}` : '';
    if (!rows.has(cell.row)) rows.set(cell.row, new Array(layout.rect.lastCol - layout.rect.firstCol + 1).fill(''));
    rows.get(cell.row)[cell.col - layout.rect.firstCol] = text;
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
const addNet = (b) => b.wb.addPivotCacheCalculatedField(b.cacheId, 'Net', 'Amount*2');
const addEw = (b) => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'EW', 'East+West');
const twoLevels = (b) => b.wb.sheet(0).setPivotRowFieldOrder(b.pivot, [b.region, b.product]);
const dateRows = (b) => b.wb.sheet(0).setPivotRowFieldOrder(b.pivot, [b.date]);
const hideEast = (b) => b.wb.sheet(0).updatePivotItem(b.pivot, b.region, 0, { visible: false });
const WEST_ONLY = ['行ラベル|Sum of Amount', 'West|37', '総計|37'];
const GREATER_THAN_20 = Object.freeze({ axis: 0, fieldName: 'Region', type: 1, valueKind: 1, valueDouble: 20 });
const BETWEEN_10_20 = Object.freeze({
  axis: 0,
  fieldName: 'Region',
  type: 2,
  valueKind: 1,
  valueDouble: 10,
  valueHighKind: 1,
  valueHighDouble: 20,
});
// Points record 0's Region at the shared item a row has just appended.
const pointAtNewItem = (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 0, 2);

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
    act: (b) => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, 42.5),
    expect: (b) =>
      assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '42.5|10', 'East|5', 'West|37', '総計|52']),
  },
  {
    label: 'date',
    act: (b) => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 3, 45292, { date: true }),
    expect: (b) => {
      const items = b.wb.listPivotCacheFieldSharedItems(b.cacheId, 3);
      assert.equal(items[0].isDate, true);
      assert.equal(items[0].value.number, 45292);
    },
    noPointing: true,
  },
  {
    label: 'text',
    act: (b) => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, 'North'),
    expect: (b) =>
      assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'North|10', 'West|37', '総計|52']),
  },
  {
    label: 'boolean',
    act: (b) => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, true),
    expect: (b) =>
      assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', 'TRUE|10', '総計|52']),
  },
  {
    label: 'blank',
    act: (b) => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, null),
    expect: (b) =>
      assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '(空白)|10', '総計|52']),
  },
  {
    label: 'error',
    act: (b) => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, ERROR_NA),
    expect: (b) =>
      assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '#N/A|10', '総計|52']),
  },
];
const RECORD_VALUE_CASES = [
  {
    label: 'number',
    act: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 2, 1000),
    expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|1005', 'West|37', '総計|1042']),
  },
  {
    label: 'date',
    act: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 3, 45300, { date: true }),
    expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
  },
  {
    label: 'text',
    act: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 0, 'West'),
    expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|47', '総計|52']),
  },
  {
    label: 'boolean',
    act: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 2, true),
    expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
  },
  {
    label: 'blank',
    act: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 2, null),
    expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
  },
  {
    label: 'error',
    act: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 2, ERROR_DIV0),
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
        assert.equal(r, 77);
        assert.equal(b.wb.listPivotCaches().length, 2);
      },
    },
  ],
  [
    'getPivotCacheId',
    { act: (b) => b.wb.sheet(0).getPivotCacheId(b.pivot), expect: (b, r) => assert.equal(r, b.cacheId) },
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
        assert.deepEqual([...b.wb.listPivotCaches()], [b.cacheId]);
      },
    },
  ],
  [
    'setPivotCacheWorksheetSource',
    {
      act: (b) => b.wb.setPivotCacheWorksheetSource(b.cacheId, { present: true, ref: 'B2:E6', sheet: 'Data' }),
      expect: (b) => {
        const source = b.wb.getPivotCacheWorksheetSource(b.cacheId);
        assert.equal(source.ref, 'B2:E6');
        assert.equal(source.sheet, 'Data');
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'addPivotCacheField',
    {
      act: (b) => b.wb.addPivotCacheField(b.cacheId, 'Extra'),
      expect: (b, r) => {
        assert.equal(r, 4);
        assert.equal(b.wb.listPivotCacheFields(b.cacheId)[4].name, 'Extra');
      },
    },
  ],
  [
    'clearPivotCacheFields',
    {
      act: (b) => b.wb.clearPivotCacheFields(b.cacheId),
      expect: (b) => {
        assert.equal(b.wb.listPivotCacheFields(b.cacheId).length, 0);
        // The data field now names a cache field that no longer exists.
        expectFailure(() => b.wb.sheet(0).getPivotLayout(b.pivot));
      },
    },
  ],
  [
    'addPivotCacheFieldSharedItem',
    {
      act: (b) => SHARED_ITEM_CASES[0].act(b),
      after: pointAtNewItem,
      expect: (b) => SHARED_ITEM_CASES[0].expect(b),
    },
  ],
  [
    'clearPivotCacheFieldSharedItems',
    {
      act: (b) => b.wb.clearPivotCacheFieldSharedItems(b.cacheId, 0),
      // With no shared items the records' indices render as themselves.
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '0|15', '1|37', '総計|52']),
    },
  ],
  [
    'addPivotCacheRecord',
    {
      act: (b) => b.wb.addPivotCacheRecord(b.cacheId),
      expect: (b, r) => {
        assert.equal(r, 4);
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|15', 'West|37', '(空白)|', '総計|52']);
      },
    },
  ],
  [
    'clearPivotCacheRecords',
    {
      act: (b) => b.wb.clearPivotCacheRecords(b.cacheId),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '総計|']),
    },
  ],
  [
    'setPivotCacheRecordValue',
    {
      act: (b) => RECORD_VALUE_CASES[0].act(b),
      expect: (b) => RECORD_VALUE_CASES[0].expect(b),
    },
  ],
  [
    'addPivot',
    {
      act: (b) => b.wb.sheet(0).addPivot({ name: 'Pivot2', cacheId: b.cacheId, anchorRow: 20, anchorCol: 1 }),
      expect: (b, r) => {
        assert.equal(r, 1);
        const layout = b.wb.sheet(0).getPivotLayout(1);
        assert.equal(layout.rect.firstRow, 20);
        assert.equal(layout.rect.firstCol, 1);
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
        expectFailure(() => b.wb.sheet(0).getPivotLayout(0));
      },
    },
  ],
  [
    'setPivotName',
    {
      act: (b) => b.wb.sheet(0).setPivotName(b.pivot, 'Renamed'),
      // A pivot's name is only observable in the part the writer emits.
      expect: (b) => {
        const saved = b.wb.save();
        assert.match(packagePart(saved, 'xl/pivotTables/pivotTable1.xml'), /<pivotTableDefinition [^>]*name="Renamed"/);
      },
    },
  ],
  [
    'setPivotAnchor',
    {
      act: (b) => b.wb.sheet(0).setPivotAnchor(b.pivot, { firstRow: 10, firstCol: 2, lastRow: 13, lastCol: 3 }),
      expect: (b) => {
        const layout = b.wb.sheet(0).getPivotLayout(b.pivot);
        assert.equal(layout.rect.firstRow, 10);
        assert.equal(layout.rect.firstCol, 2);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'setPivotGrandTotals',
    {
      setup: (b) => b.wb.sheet(0).setPivotColumnFieldOrder(b.pivot, [b.product]),
      // Rows off drops the per-row total column; columns on keeps the
      // bottom total row, so a swapped pair would show the opposite.
      act: (b) => b.wb.sheet(0).setPivotGrandTotals(b.pivot, { rows: false, columns: true }),
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
    'setPivotStyleInfo',
    {
      act: (b) =>
        b.wb.sheet(0).setPivotStyleInfo(b.pivot, {
          name: 'PivotStyleMedium9',
          showRowHeaders: false,
          showColHeaders: true,
          showRowStripes: true,
          showColStripes: false,
          showLastColumn: true,
        }),
      expect: (b) => {
        const style = b.wb.sheet(0).getPivotStyleInfo(b.pivot);
        assert.equal(style.name, 'PivotStyleMedium9');
        assert.deepEqual(
          [
            style.showRowHeaders,
            style.showColHeaders,
            style.showRowStripes,
            style.showColStripes,
            style.showLastColumn,
          ],
          [false, true, true, false, true],
        );
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'setPivotReportLayout',
    {
      act: (b) => b.wb.sheet(0).setPivotReportLayout(b.pivot, 1),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).getPivotReportLayout(b.pivot), 1);
        assert.deepEqual(pivotGrid(b.wb), ['Region|Sum of Amount', 'East|15', 'West|37', '総計|52']);
      },
    },
  ],
  [
    'addPivotField',
    {
      act: (b) => b.wb.sheet(0).addPivotField(b.pivot, { sourceName: 'Product', customName: 'Fruit', axis: 3 }),
      expect: (b, r) => {
        // Product is a cache column the fixture already added, so it is updated in place.
        assert.equal(r, b.product);
        assert.deepEqual(pivotGrid(b.wb), ['Fruit|(すべて)', '|', ...BASE_GRID]);
      },
    },
  ],
  [
    'clearPivotFields',
    {
      setup: hideEast,
      act: (b) => b.wb.sheet(0).clearPivotFields(b.pivot),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).listPivotFields(b.pivot).length, 0);
        // The row order empties with the fields; East's hidden flag goes too.
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '|52']);
      },
    },
  ],
  [
    'setPivotFieldAxis',
    {
      act: (b) => b.wb.sheet(0).setPivotFieldAxis(b.pivot, b.product, 3),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['Product|(すべて)', '|', ...BASE_GRID]),
    },
  ],
  [
    'setPivotFieldSort',
    {
      // East now outsums West, so sorting ascending by the data field
      // differs from both label orders.
      setup: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 2, 40),
      act: (b) => b.wb.sheet(0).setPivotFieldSort(b.pivot, b.region, 1, { byField: 'Sum of Amount' }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'West|37', 'East|45', '総計|82']),
    },
  ],
  [
    'updatePivotField',
    {
      act: (b) => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { defaultSubtotal: false }),
      expect: (b) => assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.region].defaultSubtotal, false),
    },
  ],
  [
    'updatePivotField',
    {
      act: (b) => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { insertBlankRow: true }),
      expect: (b) => assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.region].insertBlankRow, true),
    },
  ],
  [
    'updatePivotField',
    {
      act: (b) => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { fillDownLabels: true }),
      expect: (b) => assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.region].fillDownLabels, true),
    },
  ],
  [
    'setPivotFieldShowDetail',
    {
      setup: twoLevels,
      act: (b) => b.wb.sheet(0).setPivotFieldShowDetail(b.pivot, b.region, false),
      expect: (b) => assert.equal(b.wb.sheet(0).getPivotFieldItem(b.pivot, b.region, 0).showDetail, false),
    },
  ],
  [
    'updatePivotItem',
    {
      setup: twoLevels,
      act: (b) => b.wb.sheet(0).updatePivotItem(b.pivot, b.region, 1, { showDetail: false }),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).getPivotFieldItem(b.pivot, b.region, 1).showDetail, false);
        assert.equal(b.wb.sheet(0).getPivotFieldItem(b.pivot, b.region, 0).showDetail, true);
      },
    },
  ],
  [
    'updatePivotItem',
    {
      act: (b) => b.wb.sheet(0).updatePivotItem(b.pivot, b.region, 0, { caption: 'Eastern' }),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).getPivotFieldItem(b.pivot, b.region, 0).caption, 'Eastern');
        assert.equal(b.wb.sheet(0).getPivotFieldItem(b.pivot, b.region, 1).caption, null);
        b.wb.sheet(0).updatePivotItem(b.pivot, b.region, 0, { caption: '' });
        assert.equal(b.wb.sheet(0).getPivotFieldItem(b.pivot, b.region, 0).caption, null);
      },
    },
  ],
  [
    'setPivotOptions',
    {
      act: (b) =>
        b.wb.sheet(0).setPivotOptions(b.pivot, {
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
        const o = b.wb.sheet(0).getPivotOptions(b.pivot);
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
    'setPivotCacheOptions',
    {
      act: (b) => b.wb.setPivotCacheOptions(b.cacheId, { saveData: false, missingItemsLimit: 0, refreshOnLoad: true }),
      expect: (b) => {
        const o = b.wb.getPivotCacheOptions(b.cacheId);
        assert.equal(o.saveData, false);
        assert.equal(o.missingItemsLimit, 0);
        assert.equal(o.refreshOnLoad, true);
      },
    },
  ],
  [
    'updatePivotField',
    {
      act: (b) => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { customName: 'Area' }),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.region].customName, 'Area');
        b.wb.sheet(0).updatePivotField(b.pivot, b.region, { customName: null });
        assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.region].customName, null);
      },
    },
  ],
  [
    'updatePivotField',
    {
      act: (b) => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { showAll: true }),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.region].showAll, true);
        assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.product].showAll, false);
        b.wb.sheet(0).updatePivotField(b.pivot, b.region, { showAll: false });
        assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.region].showAll, false);
      },
    },
  ],
  [
    'projectPivot',
    {
      act: (b) => b.wb.sheet(0).projectPivot(b.pivot),
      expect: (b) => {
        const pending = b.wb.sheet(0).isPivotProjectionPending(b.pivot);
        assert.equal(pending, false);
      },
    },
  ],
  [
    'addPivotCacheCalculatedField',
    {
      act: (b) => b.wb.addPivotCacheCalculatedField(b.cacheId, 'Net', 'Amount*2'),
      expect: (b, r) => {
        assert.equal(r, 4);
        assert.equal(b.wb.listPivotCacheFields(b.cacheId)[4].name, 'Net');
        assert.equal(b.wb.listPivotCacheFields(b.cacheId)[4].formula, '=Amount*2');
        assert.equal(b.wb.listPivotCacheFields(b.cacheId)[2].formula, null);
      },
    },
  ],
  [
    'setPivotCacheCalculatedFieldFormula',
    {
      setup: addNet,
      act: (b) => b.wb.setPivotCacheCalculatedFieldFormula(b.cacheId, 4, 'Amount*3'),
      expect: (b) => assert.equal(b.wb.listPivotCacheFields(b.cacheId)[4].formula, '=Amount*3'),
    },
  ],
  [
    'removePivotCacheCalculatedField',
    {
      setup: addNet,
      act: (b) => b.wb.removePivotCacheCalculatedField(b.cacheId, 4),
      expect: (b) => assert.equal(b.wb.listPivotCacheFields(b.cacheId).length, 4),
    },
  ],
  [
    'addPivotCacheCalculatedItem',
    {
      act: (b) => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'EW', 'East+West'),
      expect: (b) => {
        assert.equal(b.wb.listPivotCacheCalculatedItems(b.cacheId).length, 1);
        const item = b.wb.listPivotCacheCalculatedItems(b.cacheId)[0];
        assert.equal(item.fieldIndex, b.region);
        assert.equal(item.name, 'EW');
        assert.equal(item.formula, '=Region[East]+Region[West]');
        assert.equal(item.scoped, false);
      },
    },
  ],
  [
    'setPivotCacheCalculatedItemFormula',
    {
      setup: addEw,
      act: (b) => b.wb.setPivotCacheCalculatedItemFormula(b.cacheId, 0, 'East*2'),
      expect: (b) => assert.equal(b.wb.listPivotCacheCalculatedItems(b.cacheId)[0].formula, '=Region[East]*2'),
    },
  ],
  [
    'movePivotCacheCalculatedItem',
    {
      setup: (b) => {
        addEw(b);
        b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'WE', 'West+East');
      },
      act: (b) => b.wb.movePivotCacheCalculatedItem(b.cacheId, 0, 1),
      expect: (b) => {
        assert.equal(b.wb.listPivotCacheCalculatedItems(b.cacheId)[0].name, 'WE');
        assert.equal(b.wb.listPivotCacheCalculatedItems(b.cacheId)[1].name, 'EW');
      },
    },
  ],
  [
    'removePivotCacheCalculatedItem',
    {
      setup: addEw,
      act: (b) => b.wb.removePivotCacheCalculatedItem(b.cacheId, 0),
      expect: (b) => assert.equal(b.wb.listPivotCacheCalculatedItems(b.cacheId).length, 0),
    },
  ],
  [
    'showPivotDetails',
    {
      // The grand total cell is a data cell, so the drill-down lands on a new
      // sheet inserted before the pivot's sheet.
      act: (b) => b.wb.sheet(0).showPivotDetails(b.pivot, 3, 7),
      expect: (b, r) => {
        assert.equal(r, b.wb.sheet(0));
        assert.equal(r.index, 0);
        assert.equal(b.wb.listSheets().length, 2);
      },
    },
  ],
  [
    'updatePivotField',
    {
      setup: twoLevels,
      act: (b) => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { subtotalTop: false }),
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
    'addPivotFieldItem',
    {
      act: (b) => b.wb.sheet(0).addPivotFieldItem(b.pivot, b.region, 'East', { visible: false }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
  [
    'addPivotFieldSharedItem',
    {
      // Only the index-addressed form can name the blank item.
      setup: (b) => {
        b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, null);
        b.wb.setPivotCacheRecordValue(b.cacheId, 0, 0, null);
      },
      act: (b) => b.wb.sheet(0).addPivotFieldSharedItem(b.pivot, b.region, 2, false),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
    },
  ],
  [
    'clearPivotFieldItems',
    {
      setup: hideEast,
      act: (b) => b.wb.sheet(0).clearPivotFieldItems(b.pivot, b.region),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'updatePivotItem',
    {
      act: (b) => b.wb.sheet(0).updatePivotItem(b.pivot, b.region, 1, { visible: false }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|15', '総計|15']),
    },
  ],
  [
    'addPivotFieldSubtotalFn',
    {
      setup: twoLevels,
      act: (b) => b.wb.sheet(0).addPivotFieldSubtotalFn(b.pivot, b.region, 3),
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
    'clearPivotFieldSubtotalFns',
    {
      setup: (b) => {
        twoLevels(b);
        b.wb.sheet(0).addPivotFieldSubtotalFn(b.pivot, b.region, 3);
      },
      act: (b) => b.wb.sheet(0).clearPivotFieldSubtotalFns(b.pivot, b.region),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), TWO_LEVEL_GRID),
    },
  ],
  [
    'setPivotFieldDateGroup',
    {
      setup: dateRows,
      act: (b) => b.wb.sheet(0).setPivotFieldDateGroup(b.pivot, b.date, { periods: 64, calendar: 1 }),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '令和6年|40', '令和7年|12', '総計|52']),
    },
  ],
  [
    'clearPivotFieldDateGroup',
    {
      setup: (b) => {
        dateRows(b);
        b.wb.sheet(0).setPivotFieldDateGroup(b.pivot, b.date, { periods: 64, calendar: 0 });
      },
      act: (b) => b.wb.sheet(0).clearPivotFieldDateGroup(b.pivot, b.date),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), DATE_SERIAL_GRID),
    },
  ],
  [
    'updatePivotField',
    {
      act: (b) => b.wb.sheet(0).updatePivotField(b.pivot, b.amount, { numberFormat: '4' }),
      // A field's numFmtId is only observable in the part the writer emits.
      expect: (b) => {
        expectFailure(
          () => b.wb.sheet(0).updatePivotField(b.pivot, 99, { numberFormat: '4' }),
          b.mod.StatusCode.IndexOutOfRange,
        );
        expectFailure(
          () => b.wb.sheet(0).updatePivotField(b.pivot, b.amount, { numberFormat: '0.00' }),
          b.mod.StatusCode.InvalidArgument,
        );
        const saved = b.wb.save();
        assert.match(packagePart(saved, 'xl/pivotTables/pivotTable1.xml'), /<pivotField [^>]*numFmtId="4"/);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'setPivotRowFieldOrder',
    {
      act: (b) => b.wb.sheet(0).setPivotRowFieldOrder(b.pivot, [b.product]),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'Apple|40', 'Pear|12', '総計|52']),
    },
  ],
  [
    'setPivotColumnFieldOrder',
    {
      act: (b) => b.wb.sheet(0).setPivotColumnFieldOrder(b.pivot, [b.product]),
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
    'addPivotDataField',
    {
      act: (b) =>
        b.wb.sheet(0).addPivotDataField(b.pivot, { name: 'Max of Amount', fieldIndex: b.amount, aggregation: 3 }),
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
    'clearPivotDataFields',
    {
      act: (b) => b.wb.sheet(0).clearPivotDataFields(b.pivot),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル', 'East', 'West', '総計']),
    },
  ],
  [
    'setPivotDataField',
    {
      act: (b) =>
        b.wb.sheet(0).setPivotDataField(b.pivot, 0, {
          name: 'Count of Amount',
          fieldIndex: b.amount,
          aggregation: 1,
          numberFormat: '2',
        }),
      expect: (b) => {
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Count of Amount', 'East|2', 'West|2', '総計|4']);
        const counts = b.wb
          .sheet(0)
          .getPivotLayout(b.pivot)
          .cells.filter((c) => c.value.kind === 1);
        assert.deepEqual(
          counts.map((c) => c.numberFormat),
          ['2', '2', '2'],
        );
      },
    },
  ],
  [
    'addPivotFilter',
    {
      act: (b) => b.wb.sheet(0).addPivotFilter(b.pivot, GREATER_THAN_20),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
  [
    'clearPivotFilters',
    {
      setup: (b) => b.wb.sheet(0).addPivotFilter(b.pivot, GREATER_THAN_20),
      act: (b) => b.wb.sheet(0).clearPivotFilters(b.pivot),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'removePivotFilter',
    {
      // Each filter alone keeps a different region, so removing the wrong
      // index leaves the wrong one.
      setup: (b) => {
        b.wb.sheet(0).addPivotFilter(b.pivot, GREATER_THAN_20);
        b.wb.sheet(0).addPivotFilter(b.pivot, BETWEEN_10_20);
      },
      act: (b) => b.wb.sheet(0).removePivotFilter(b.pivot, 1),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
];

test('the pivot mutator table covers every declared pivot mutator', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
  const declared = [...dts.matchAll(/^ {2}((?:[a-z]+Pivot|pivot)\w*)\([^)]*\): (?:void|number|Sheet);/gm)]
    .map((m) => m[1])
    .filter((name) => !name.endsWith('Count'));
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

test('a date flag on a value that is not a number is a TypeError and writes nothing', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const items = b.wb.listPivotCacheFieldSharedItems(b.cacheId, 0).length;
    for (const value of ['x', true, null, ERROR_NA]) {
      assert.throws(() => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, value, { date: true }), TypeError);
      assert.throws(() => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 2, value, { date: true }), TypeError);
    }
    assert.equal(b.wb.listPivotCacheFieldSharedItems(b.cacheId, 0).length, items);
  } finally {
    b.wb.dispose();
  }
});

test('manual pivot sort order keeps item order and rejects bad arguments', async () => {
  const mod = await getModule();
  const { PivotSortOrder } = await loadStagedModule();
  const b = buildPivot(mod);
  try {
    b.wb.sheet(0).clearPivotFieldItems(b.pivot, b.region);
    b.wb.sheet(0).addPivotFieldItem(b.pivot, b.region, 'West', { visible: true });
    b.wb.sheet(0).addPivotFieldItem(b.pivot, b.region, 'East', { visible: true });
    b.wb.sheet(0).setPivotFieldSort(b.pivot, b.region, PivotSortOrder.Manual);
    assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'West|37', 'East|15', '総計|52']);
    expectFailure(() => b.wb.sheet(0).setPivotFieldSort(b.pivot, b.region, 3));
    expectFailure(() =>
      b.wb.sheet(0).setPivotFieldSort(b.pivot, b.region, PivotSortOrder.Manual, { byField: 'Sum of Amount' }),
    );
    b.wb.sheet(0).setPivotFieldSort(b.pivot, b.region, PivotSortOrder.Ascending);
    assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('pivot filters and subtotal functions survive save and reload', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    b.wb.sheet(0).addPivotFilter(b.pivot, GREATER_THAN_20);
    b.wb.sheet(0).addPivotFieldSubtotalFn(b.pivot, b.region, 2);
    b.wb.sheet(0).addPivotFieldSubtotalFn(b.pivot, b.region, 0);
    const region = b.wb.sheet(0).listPivotFields(b.pivot)[b.region];
    assert.equal(region.defaultSubtotal, false);
    assert.deepEqual(region.subtotalFns, [0, 2]);
    assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.amount].defaultSubtotal, true);
    assert.deepEqual(b.wb.sheet(0).listPivotFields(b.pivot)[b.amount].subtotalFns, []);

    const saved = b.wb.save();
    assert.match(
      packagePart(saved, 'xl/pivotTables/pivotTable1.xml'),
      /<filters count="1"><filter fld="0" type="valueGreaterThan" evalOrder="-1" id="1" iMeasureFld="0">/,
    );
    const back = mod.Workbook.load(saved);
    try {
      back.setExcelProfileId('win-365-ja_JP');
      assert.equal(back.sheet(0).listPivotFilters(0).length, 1);
      const filter = back.sheet(0).listPivotFilters(0)[0];
      assert.equal(filter.type, 1);
      assert.equal(filter.fieldName, 'Region');
      assert.equal(filter.valueKind, 1);
      assert.equal(filter.valueDouble, 20);
      assert.deepEqual(pivotGrid(back), WEST_ONLY);
      const field = back.sheet(0).listPivotFields(0)[0];
      assert.equal(field.defaultSubtotal, false);
      assert.deepEqual(field.subtotalFns, [0, 2]);
      back.sheet(0).clearPivotFilters(0);
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
  ['listTables', [], null],
  ['listPassthroughs', [], null],
  ['listPivots', [0], [99]],
  ['listPivotCaches', [], null],
  ['listPivotCacheFields', null, [9999]],
  ['listPivotCacheFieldSharedItems', null, [9999, 0]],
  ['getPivotCacheRecordCount', null, [9999]],
  ['listPivotFields', null, [0, 99]],
  ['listPivotFieldItems', null, [0, 99, 0]],
  ['getPivotName', null, [0, 99]],
  ['getPivotReportLayout', null, [0, 99]],
  ['listPivotDataFields', null, [0, 99]],
  ['listPivotFilters', null, [0, 99]],
  ['listPivotCacheCalculatedItems', null, [9999]],
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
  ['getExcelProfileId', [], null],
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
];

// Calls a STATUS_ACCESSORS row on the workbook, or on the sheet its first argument names.
function callAccessor(wb, name, args) {
  return name in wb ? wb[name](...args) : wb.sheet(args[0])[name](...args.slice(1));
}

test('every numeric accessor is covered by the status table', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
  const declared = [...dts.matchAll(/^ {2}([A-Za-z]\w*)\([^)]*\): (?:number|CalcMode|RefMode|PivotReportLayout);/gm)]
    .map((m) => m[1])
    .filter((name) => !/^(add|copy|restore|showPivotDetails|getActiveSheet|getPivotCacheId)/.test(name));
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
  assert.equal(wb.getExcelProfileId(), 'win-365-en_US');
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
    explicit.wb.sheet(0).setPivotFieldDateGroup(explicit.pivot, explicit.date, {
      periods: days,
      calendar: 0,
      intervalDays: 1,
      start: Number.NaN,
      end: Number.NaN,
    });
    defaulted.wb.sheet(0).setPivotFieldDateGroup(defaulted.pivot, defaulted.date, { periods: days });
    const field = defaulted.wb.sheet(0).listPivotFields(defaulted.pivot)[defaulted.date];
    assert.equal(field.datePeriods, days);
    assert.deepEqual(field, explicit.wb.sheet(0).listPivotFields(explicit.pivot)[explicit.date]);
    assert.deepEqual(pivotGrid(defaulted.wb), pivotGrid(explicit.wb));
  } finally {
    explicit.wb.dispose();
    defaulted.wb.dispose();
  }
});

test('pivot cache date shared items report as dates and reject invalid serials', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    b.wb.addPivotCacheFieldSharedItem(b.cacheId, 3, 45292, { date: true });
    b.wb.addPivotCacheFieldSharedItem(b.cacheId, 3, 45323.5, { date: true });
    b.wb.addPivotCacheFieldSharedItem(b.cacheId, 3, 7);
    assert.deepEqual(
      b.wb.listPivotCacheFieldSharedItems(b.cacheId, 3).map((item) => item.isDate),
      [true, true, false],
    );
    expectFailure(() => b.wb.listPivotCacheFieldSharedItems(9999, 3));
    for (const bad of [-1, Number.NaN]) {
      expectFailure(() => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 3, bad, { date: true }));
      expectFailure(() => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 3, bad, { date: true }));
    }
  } finally {
    b.wb.dispose();
  }
});

test('a pivot cache shared item reads back by index', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const items = b.wb.listPivotCacheFieldSharedItems(b.cacheId, 0);
    assert.equal(items[0].value.text, 'East');
    assert.equal(items[1].value.text, 'West');
    assert.equal(items.length, 2);
    expectFailure(() => b.wb.listPivotCacheFieldSharedItems(b.cacheId, 99));
  } finally {
    b.wb.dispose();
  }
});

test('a pivot cache error value without errorCode throws and changes nothing', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const items = b.wb.listPivotCacheFieldSharedItems(b.cacheId, 0).length;
    assert.throws(() => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, { kind: 4 }), TypeError);
    assert.equal(b.wb.listPivotCacheFieldSharedItems(b.cacheId, 0).length, items);
    assert.throws(() => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 2, { kind: 4 }), TypeError);
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
      () => wb.sheet(0).updatePivotField(pivot, 99, { defaultSubtotal: true }),
      () => wb.sheet(0).updatePivotField(pivot, 99, { insertBlankRow: true }),
      () => wb.sheet(0).updatePivotField(pivot, 99, { fillDownLabels: true }),
      () => wb.sheet(0).setPivotFieldShowDetail(pivot, 99, true),
      () => wb.sheet(0).updatePivotItem(pivot, region, 99, { showDetail: true }),
      () => wb.sheet(0).updatePivotItem(pivot, region, 99, { caption: 'x' }),
      () => wb.sheet(0).setPivotOptions(9, { indent: 1 }),
      () => wb.setPivotCacheOptions(9999, { saveData: true }),
      () => wb.sheet(0).updatePivotField(pivot, 99, { customName: 'x' }),
      () => wb.sheet(0).updatePivotField(pivot, 99, { showAll: true }),
      () => wb.sheet(0).projectPivot(9),
      () => wb.sheet(0).showPivotDetails(pivot, 99, 99),
      () => wb.sheet(0).getPivotOptions(9),
      () => wb.getPivotCacheOptions(9999),
      () => wb.sheet(0).isPivotProjectionPending(9),
      () => wb.listPivotCacheFields(9999),
      () => wb.listPivotCacheCalculatedItems(9999),
      () => wb.sheet(0).getPivotDataFieldDefaultName(9, 0, 0),
      () => wb.sheet(0).setDrawingName(12345, 'x'),
    ]) {
      expectFailure(call);
    }
    // The Values level needs at least two data fields.
    wb.sheet(0).addPivotDataField(pivot, { name: 'Count of Date', fieldIndex: b.date, aggregation: 1 });
    wb.sheet(0).setPivotRowFieldOrder(pivot, [region, PIVOT_VALUES_FIELD, product]);
    assert.deepEqual(wb.sheet(0).getPivotRowFieldOrder(pivot), [region, PIVOT_VALUES_FIELD, product]);
    expectFailure(() => wb.sheet(0).setPivotColumnFieldOrder(pivot, [-3]));
  } finally {
    b.wb.dispose();
  }
});

test('pivotSetOptions merges onto the current options and rejects bad values without changing anything', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const { wb, pivot } = b;
    const before = wb.sheet(0).getPivotOptions(pivot);
    wb.sheet(0).setPivotOptions(pivot, {});
    assert.deepEqual(wb.sheet(0).getPivotOptions(pivot), before);
    wb.sheet(0).setPivotOptions(pivot, { indent: 4 });
    const after = wb.sheet(0).getPivotOptions(pivot);
    assert.equal(after.indent, 4);
    assert.deepEqual({ ...after, indent: before.indent }, before);

    // A caption of exactly 255 UTF-16 units is kept; one more is not. A surrogate pair counts as two.
    wb.sheet(0).setPivotOptions(pivot, { errorCaption: 'e'.repeat(255) });
    assert.equal(wb.sheet(0).getPivotOptions(pivot).errorCaption.length, 255);
    wb.sheet(0).setPivotOptions(pivot, { errorCaption: '\u{1F600}'.repeat(127) });
    const kept = wb.sheet(0).getPivotOptions(pivot);
    for (const bad of [
      { errorCaption: 'e'.repeat(256) },
      { missingCaption: '\u{1F600}'.repeat(128) },
      { grandTotalCaption: 'g'.repeat(256) },
      { altText: 'a'.repeat(256) },
      { indent: 127 },
      { pageWrap: 256 },
    ]) {
      expectFailure(() => wb.sheet(0).setPivotOptions(pivot, bad));
      assert.deepEqual(wb.sheet(0).getPivotOptions(pivot), kept, JSON.stringify(bad));
    }
    // A value of the wrong kind throws before anything is written.
    for (const bad of [{ indent: '3' }, { indent: 1.5 }, { indent: -1 }, { pageWrap: Number.NaN }]) {
      assert.throws(() => wb.sheet(0).setPivotOptions(pivot, bad), wrongKind, JSON.stringify(bad));
      assert.deepEqual(wb.sheet(0).getPivotOptions(pivot), kept, JSON.stringify(bad));
    }
    wb.sheet(0).setPivotOptions(pivot, { errorCaption: null });
    assert.equal(wb.sheet(0).getPivotOptions(pivot).errorCaption, null);
    wb.sheet(0).setPivotOptions(pivot, { errorCaption: '' });
    assert.equal(wb.sheet(0).getPivotOptions(pivot).errorCaption, null);
    assert.throws(() => wb.sheet(0).setPivotOptions(pivot), TypeError);
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
    const before = wb.getPivotCacheOptions(cacheId);
    wb.setPivotCacheOptions(cacheId, { saveData: !before.saveData });
    assert.deepEqual(wb.getPivotCacheOptions(cacheId), {
      ...before,
      saveData: !before.saveData,
    });
    for (const limit of [null, 0, 1, 32500, 1048576]) {
      wb.setPivotCacheOptions(cacheId, { missingItemsLimit: limit });
      assert.equal(wb.getPivotCacheOptions(cacheId).missingItemsLimit, limit);
    }
    const kept = wb.getPivotCacheOptions(cacheId);
    expectFailure(() => wb.setPivotCacheOptions(cacheId, { missingItemsLimit: 1048577 }));
    assert.deepEqual(wb.getPivotCacheOptions(cacheId), kept);
    for (const bad of [-1, -2, 1.5, Number.NaN, '7']) {
      assert.throws(() => wb.setPivotCacheOptions(cacheId, { missingItemsLimit: bad }), wrongKind, String(bad));
      assert.deepEqual(wb.getPivotCacheOptions(cacheId), kept, String(bad));
    }
    assert.throws(() => wb.setPivotCacheOptions(cacheId), TypeError);

    wb.setPivotCacheOptions(cacheId, { saveData: true, missingItemsLimit: 5, refreshOnLoad: true });
    const saved = wb.save();
    const back = mod.Workbook.load(saved);
    try {
      const o = back.getPivotCacheOptions(back.listPivotCaches()[0]);
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
    b.wb.sheet(0).setPivotOptions(b.pivot, { grandTotalCaption: 'All', altText: 'Sales', indent: 2, mergeItem: true });
    b.wb.sheet(0).updatePivotField(b.pivot, b.region, { showAll: true });
    b.wb.sheet(0).updatePivotField(b.pivot, b.product, { customName: 'Item' });
    const saved = b.wb.save();
    const back = mod.Workbook.load(saved);
    try {
      const o = back.sheet(0).getPivotOptions(0);
      assert.equal(o.grandTotalCaption, 'All');
      assert.equal(o.altText, 'Sales');
      assert.equal(o.indent, 2);
      assert.equal(o.mergeItem, true);
      assert.equal(back.sheet(0).listPivotFields(0)[b.region].showAll, true);
      assert.equal(back.sheet(0).listPivotFields(0)[b.product].customName, 'Item');
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
    wb.sheet(0).updatePivotField(pivot, region, { customName: 'Region' });
    expectFailure(() => wb.sheet(0).updatePivotField(pivot, region, { customName: 'Product' }));
    expectFailure(() => wb.sheet(0).updatePivotField(pivot, region, { customName: 'Sum of Amount' }));
    expectFailure(() => wb.sheet(0).updatePivotField(pivot, region, { customName: 'n'.repeat(256) }));
    expectFailure(() => wb.sheet(0).updatePivotField(pivot, 4096, { customName: 'x' }));
    assert.throws(() => wb.sheet(0).updatePivotField(pivot, region), TypeError);
    assert.throws(() => wb.sheet(0).updatePivotField(pivot), TypeError);
    assert.equal(wb.sheet(0).listPivotFields(pivot)[region].customName, 'Region');
  } finally {
    b.wb.dispose();
  }
});

test('pivotDataFieldDefaultName numbers a repeated source field and rejects bad arguments', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const { wb, pivot, amount } = b;
    const first = wb.sheet(0).getPivotDataFieldDefaultName(pivot, amount, 0);
    assert.equal(first, '\u5408\u8a08 / Amount');
    assert.equal(wb.sheet(0).listPivotDataFields(pivot).length, 1, 'reading a name adds nothing');
    const added = wb.sheet(0).addPivotDataField(pivot, { fieldIndex: amount, aggregation: 0 });
    assert.equal(wb.sheet(0).listPivotDataFields(pivot)[added].name, first);
    // The same source field again gets the numbered name.
    assert.equal(wb.sheet(0).getPivotDataFieldDefaultName(pivot, amount, 0), '\u5408\u8a08 / Amount2');
    // Leaving out the entry that holds the name frees it, as a replacement would.
    assert.equal(wb.sheet(0).getPivotDataFieldDefaultName(pivot, amount, 0, added), first);
    expectFailure(() => wb.sheet(0).getPivotDataFieldDefaultName(pivot, 99, 0));
    expectFailure(() => wb.sheet(0).getPivotDataFieldDefaultName(pivot, amount, 99));
    expectFailure(() => wb.sheet(0).getPivotDataFieldDefaultName(pivot, amount, 0, 99));
    expectFailure(() => wb.sheet(0).getPivotDataFieldDefaultName(9, amount, 0));
    assert.throws(() => wb.sheet(0).getPivotDataFieldDefaultName(pivot, amount), TypeError);
  } finally {
    b.wb.dispose();
  }
});

test('calculated fields and items reject bad names and formulas and change nothing', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const { wb, cacheId, region } = b;
    const fields = wb.listPivotCacheFields(cacheId).length;
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
      expectFailure(() => wb.addPivotCacheCalculatedField(cacheId, name, formula));
    }
    assert.equal(wb.listPivotCacheFields(cacheId).length, fields);
    expectFailure(() => wb.addPivotCacheCalculatedField(9999, 'Net', 'Amount*2'));
    expectFailure(() => wb.setPivotCacheCalculatedFieldFormula(cacheId, 2, 'Amount*2'));
    expectFailure(() => wb.removePivotCacheCalculatedField(cacheId, 2));
    assert.throws(() => wb.addPivotCacheCalculatedField(cacheId, 'Net'), TypeError);
    assert.throws(() => wb.addPivotCacheCalculatedItem(cacheId, region, 'EW'), TypeError);

    for (const [field, name, formula] of [
      [region, '', 'East+West'],
      [region, 'EW', ''],
      [region, 'All', 'East+West'],
      [region, 'East', 'East+West'],
      [region, 'EW', 'East+'],
      [region, 'EW', 'Nowhere+East'],
      [4096, 'EW', 'East+West'],
    ]) {
      expectFailure(() => wb.addPivotCacheCalculatedItem(cacheId, field, name, formula));
    }
    assert.equal(wb.listPivotCacheCalculatedItems(cacheId).length, 0);
    assert.deepEqual(wb.listPivotCacheCalculatedItems(cacheId), []);
    expectFailure(() => wb.setPivotCacheCalculatedItemFormula(cacheId, 0, 'East'));
    expectFailure(() => wb.movePivotCacheCalculatedItem(cacheId, 0, 0));
    expectFailure(() => wb.removePivotCacheCalculatedItem(cacheId, 0));
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
    const net = wb.addPivotCacheCalculatedField(cacheId, 'Net', 'Amount*2');
    wb.sheet(0).addPivotDataField(pivot, { name: 'Net', fieldIndex: net, aggregation: 0 });
    assert.equal(pivotGrid(wb).at(-1), '\u7dcf\u8a08|52|104');
    // A removal takes the data field with it.
    wb.removePivotCacheCalculatedField(cacheId, net);
    assert.equal(wb.sheet(0).listPivotDataFields(pivot).length, 1);
    assert.deepEqual(pivotGrid(wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('pivot filters take the full type range and an upper text bound', async () => {
  const mod = await getModule();
  const { PivotFilterType } = await loadStagedModule();
  assert.equal(Object.keys(PivotFilterType).length, 68);
  assert.equal(PivotFilterType.ValueEqual, 6);
  assert.equal(PivotFilterType.LabelBetween, 22);
  assert.equal(PivotFilterType.DateM12, 67);
  const b = buildPivot(mod);
  try {
    const { wb, pivot } = b;
    const between = {
      axis: 0,
      fieldName: 'Region',
      type: PivotFilterType.LabelBetween,
      valueKind: 2,
      valueText: 'A',
      valueHighKind: 2,
      valueHighText: 'F',
    };
    // Nothing the upper bound's rules reject is added.
    for (const bad of [
      { ...between, valueHighText: undefined },
      { ...between, type: PivotFilterType.LabelContains },
      { ...between, type: PivotFilterType.ValueEqual, valueKind: 1, valueDouble: 15 },
      { ...between, type: 68 },
    ]) {
      expectFailure(() => wb.sheet(0).addPivotFilter(pivot, bad));
    }
    assert.equal(wb.sheet(0).listPivotFilters(pivot).length, 0);
    assert.deepEqual(pivotGrid(wb), BASE_GRID);

    wb.sheet(0).addPivotFilter(pivot, between);
    const at = wb.sheet(0).listPivotFilters(pivot)[0];
    assert.equal(at.type, PivotFilterType.LabelBetween);
    assert.equal(at.valueText, 'A');
    assert.equal(at.valueHighText, 'F');
    assert.deepEqual(pivotGrid(wb), ['\u884c\u30e9\u30d9\u30eb|Sum of Amount', 'East|15', '\u7dcf\u8a08|15']);

    wb.sheet(0).clearPivotFilters(pivot);
    wb.sheet(0).addPivotFilter(pivot, {
      axis: 0,
      fieldName: 'Region',
      type: PivotFilterType.ValueEqual,
      valueKind: 1,
      valueDouble: 15,
    });
    assert.equal(wb.sheet(0).listPivotFilters(pivot)[0].valueHighText, null);
    assert.deepEqual(pivotGrid(wb), ['\u884c\u30e9\u30d9\u30eb|Sum of Amount', 'East|15', '\u7dcf\u8a08|15']);
  } finally {
    b.wb.dispose();
  }
});
