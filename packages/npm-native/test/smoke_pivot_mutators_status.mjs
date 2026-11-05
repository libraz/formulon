import test from 'node:test';
import { inflateRawSync } from 'node:zlib';
import { assert, getModule, path, pkgRoot, readFile } from './smoke_support.mjs';

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
  const must = (r, what) => {
    const status = typeof r.ok === 'boolean' ? r : r.status;
    assert.ok(status.ok, `${what}: ${JSON.stringify(status)}`);
    return r;
  };
  must(wb.setExcelProfileId('win-365-ja_JP'), 'setExcelProfileId');
  const cacheId = must(wb.addPivotCache(0), 'addPivotCache').index;
  must(wb.setPivotCacheWorksheetSource(cacheId, { present: true, ref: 'A1:D5', sheet: 'Sheet1' }), 'source');
  for (const name of ['Region', 'Product', 'Amount', 'Date']) must(wb.addPivotCacheField(cacheId, name), name);
  for (const item of ['East', 'West']) must(wb.addPivotCacheFieldSharedItemText(cacheId, 0, item), item);
  for (const item of ['Apple', 'Pear']) must(wb.addPivotCacheFieldSharedItemText(cacheId, 1, item), item);
  for (const values of [
    [0, 0, 10, 45292],
    [1, 0, 30, 45323],
    [0, 1, 5, 45658],
    [1, 1, 7, 45689],
  ]) {
    const rec = must(wb.addPivotCacheRecord(cacheId), 'record').index;
    for (const [field, v] of values.entries())
      must(wb.setPivotCacheRecordNumber(cacheId, rec, field, v), 'record value');
  }
  const pivot = must(wb.addPivot(0, 'Pivot1', cacheId, 0, 6), 'addPivot').index;
  const addField = (spec) => must(wb.addPivotField(0, pivot, spec), spec.sourceName).index;
  const region = addField({ sourceName: 'Region', axis: 0 });
  must(wb.addPivotFieldItem(0, pivot, region, 'East', true), 'item');
  must(wb.addPivotFieldItem(0, pivot, region, 'West', true), 'item');
  const product = addField({ sourceName: 'Product', axis: 2 });
  must(wb.addPivotFieldItem(0, pivot, product, 'Apple', true), 'item');
  must(wb.addPivotFieldItem(0, pivot, product, 'Pear', true), 'item');
  const amount = addField({ sourceName: 'Amount', axis: 2 });
  const date = addField({ sourceName: 'Date', axis: 2 });
  must(wb.setPivotRowFieldOrder(0, pivot, [region]), 'row order');
  must(wb.addPivotDataField(0, pivot, { name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 }), 'data field');
  return { wb, cacheId, pivot, region, product, amount, date };
}

// The projected grid as one `|`-joined string per sheet row.
function pivotGrid(wb, pivot = 0) {
  const layout = wb.getPivotLayout(0, pivot);
  assert.ok(layout.status.ok, JSON.stringify(layout.status));
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
const twoLevels = (b) => b.wb.setPivotRowFieldOrder(0, b.pivot, [b.region, b.product]);
const dateRows = (b) => b.wb.setPivotRowFieldOrder(0, b.pivot, [b.date]);
const hideEast = (b) => b.wb.setPivotFieldItemVisible(0, b.pivot, b.region, 0, false);
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
const pointAtNewItem = (b) => b.wb.setPivotCacheRecordNumber(b.cacheId, 0, 0, 2);

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

// [method, { setup, act, after, expect }]. `act` performs the one call
// under test; `expect` reads the result back.
const PIVOT_MUTATORS = [
  [
    'addPivotCache',
    {
      act: (b) => b.wb.addPivotCache(77),
      expect: (b, r) => {
        assert.equal(r.index, 77);
        assert.equal(b.wb.pivotCacheCount().value, 2);
      },
    },
  ],
  ['pivotCacheIdAt', { act: (b) => b.wb.pivotCacheIdAt(0), expect: (b, r) => assert.equal(r.index, b.cacheId) }],
  [
    'getPivotCacheId',
    { act: (b) => b.wb.getPivotCacheId(0, b.pivot), expect: (b, r) => assert.equal(r.index, b.cacheId) },
  ],
  [
    'removePivotCache',
    {
      setup: (b) => {
        b.spare = b.wb.addPivotCache(0).index;
      },
      act: (b) => b.wb.removePivotCache(b.spare),
      expect: (b) => {
        assert.equal(b.wb.pivotCacheCount().value, 1);
        assert.equal(b.wb.pivotCacheIdAt(0).index, b.cacheId);
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
        assert.equal(r.index, 4);
        assert.equal(b.wb.getPivotCacheFieldName(b.cacheId, 4).value, 'Extra');
      },
    },
  ],
  [
    'clearPivotCacheFields',
    {
      act: (b) => b.wb.clearPivotCacheFields(b.cacheId),
      expect: (b) => {
        assert.equal(b.wb.getPivotCacheFieldCount(b.cacheId).value, 0);
        // The data field now names a cache field that no longer exists.
        assert.equal(b.wb.getPivotLayout(0, b.pivot).status.ok, false);
      },
    },
  ],
  [
    'addPivotCacheFieldSharedItemNumber',
    {
      act: (b) => b.wb.addPivotCacheFieldSharedItemNumber(b.cacheId, 0, 42.5),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '42.5|10', 'East|5', 'West|37', '総計|52']),
    },
  ],
  [
    'addPivotCacheFieldSharedItemDate',
    {
      act: (b) => b.wb.addPivotCacheFieldSharedItemDate(b.cacheId, 3, 45292),
      expect: (b) => assert.equal(b.wb.isPivotCacheFieldSharedItemDate(b.cacheId, 3, 0).value, true),
    },
  ],
  [
    'addPivotCacheFieldSharedItemText',
    {
      act: (b) => b.wb.addPivotCacheFieldSharedItemText(b.cacheId, 0, 'North'),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'North|10', 'West|37', '総計|52']),
    },
  ],
  [
    'addPivotCacheFieldSharedItemBool',
    {
      act: (b) => b.wb.addPivotCacheFieldSharedItemBool(b.cacheId, 0, true),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', 'TRUE|10', '総計|52']),
    },
  ],
  [
    'addPivotCacheFieldSharedItemBlank',
    {
      act: (b) => b.wb.addPivotCacheFieldSharedItemBlank(b.cacheId, 0),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '(空白)|10', '総計|52']),
    },
  ],
  [
    'addPivotCacheFieldSharedItemError',
    {
      act: (b) => b.wb.addPivotCacheFieldSharedItemError(b.cacheId, 0, 6),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '#N/A|10', '総計|52']),
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
        assert.equal(r.index, 4);
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
    'setPivotCacheRecordNumber',
    {
      act: (b) => b.wb.setPivotCacheRecordNumber(b.cacheId, 0, 2, 1000),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|1005', 'West|37', '総計|1042']),
    },
  ],
  [
    'setPivotCacheRecordDate',
    {
      act: (b) => b.wb.setPivotCacheRecordDate(b.cacheId, 0, 3, 45300),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'setPivotCacheRecordText',
    {
      act: (b) => b.wb.setPivotCacheRecordText(b.cacheId, 0, 0, 'West'),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|47', '総計|52']),
    },
  ],
  [
    'setPivotCacheRecordBool',
    {
      act: (b) => b.wb.setPivotCacheRecordBool(b.cacheId, 0, 2, true),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
    },
  ],
  [
    'setPivotCacheRecordBlank',
    {
      act: (b) => b.wb.setPivotCacheRecordBlank(b.cacheId, 0, 2),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
    },
  ],
  [
    'setPivotCacheRecordError',
    {
      act: (b) => b.wb.setPivotCacheRecordError(b.cacheId, 0, 2, 1),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|#1', 'West|37', '総計|#1']),
    },
  ],
  [
    'addPivot',
    {
      act: (b) => b.wb.addPivot(0, 'Pivot2', b.cacheId, 20, 1),
      expect: (b, r) => {
        assert.equal(r.index, 1);
        const layout = b.wb.getPivotLayout(0, 1);
        assert.equal(layout.rect.firstRow, 20);
        assert.equal(layout.rect.firstCol, 1);
        assert.deepEqual(pivotGrid(b.wb, 0), BASE_GRID);
      },
    },
  ],
  [
    'removePivot',
    {
      act: (b) => b.wb.removePivot(0, b.pivot),
      expect: (b) => {
        assert.equal(b.wb.getPivotCount(0).value, 0);
        assert.equal(b.wb.getPivotLayout(0, 0).status.ok, false);
      },
    },
  ],
  [
    'setPivotName',
    {
      act: (b) => b.wb.setPivotName(0, b.pivot, 'Renamed'),
      // A pivot's name is only observable in the part the writer emits.
      expect: (b) => {
        const saved = b.wb.save();
        assert.ok(saved.status.ok, JSON.stringify(saved.status));
        assert.match(
          packagePart(saved.bytes, 'xl/pivotTables/pivotTable1.xml'),
          /<pivotTableDefinition [^>]*name="Renamed"/,
        );
      },
    },
  ],
  [
    'setPivotAnchor',
    {
      act: (b) => b.wb.setPivotAnchor(0, b.pivot, { firstRow: 10, firstCol: 2, lastRow: 13, lastCol: 3 }),
      expect: (b) => {
        const layout = b.wb.getPivotLayout(0, b.pivot);
        assert.equal(layout.rect.firstRow, 10);
        assert.equal(layout.rect.firstCol, 2);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'setPivotGrandTotals',
    {
      setup: (b) => b.wb.setPivotColFieldOrder(0, b.pivot, [b.product]),
      // Rows off drops the per-row total column; columns on keeps the
      // bottom total row, so a swapped pair would show the opposite.
      act: (b) => b.wb.setPivotGrandTotals(0, b.pivot, false, true),
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
        b.wb.setPivotStyleInfo(0, b.pivot, {
          name: 'PivotStyleMedium9',
          showRowHeaders: false,
          showColHeaders: true,
          showRowStripes: true,
          showColStripes: false,
          showLastColumn: true,
        }),
      expect: (b) => {
        const style = b.wb.getPivotStyleInfo(0, b.pivot);
        assert.ok(style.status.ok);
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
      act: (b) => b.wb.setPivotReportLayout(0, b.pivot, 1),
      expect: (b) => {
        assert.equal(b.wb.getPivotReportLayout(0, b.pivot).layout, 1);
        assert.deepEqual(pivotGrid(b.wb), ['Region|Sum of Amount', 'East|15', 'West|37', '総計|52']);
      },
    },
  ],
  [
    'addPivotField',
    {
      act: (b) => b.wb.addPivotField(0, b.pivot, { sourceName: 'Product', customName: 'Fruit', axis: 3 }),
      expect: (b, r) => {
        // Product is a cache column the fixture already added, so it is updated in place.
        assert.equal(r.index, b.product);
        assert.deepEqual(pivotGrid(b.wb), ['Fruit|(すべて)', '|', ...BASE_GRID]);
      },
    },
  ],
  [
    'clearPivotFields',
    {
      setup: hideEast,
      act: (b) => b.wb.clearPivotFields(0, b.pivot),
      expect: (b) => {
        assert.equal(b.wb.getPivotFieldCount(0, b.pivot).value, 0);
        // The row order empties with the fields; East's hidden flag goes too.
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '|52']);
      },
    },
  ],
  [
    'setPivotFieldAxis',
    {
      act: (b) => b.wb.setPivotFieldAxis(0, b.pivot, b.product, 3),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['Product|(すべて)', '|', ...BASE_GRID]),
    },
  ],
  [
    'setPivotFieldSort',
    {
      // East now outsums West, so sorting ascending by the data field
      // differs from both label orders.
      setup: (b) => b.wb.setPivotCacheRecordNumber(b.cacheId, 0, 2, 40),
      act: (b) => b.wb.setPivotFieldSort(0, b.pivot, b.region, 1, 'Sum of Amount'),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'West|37', 'East|45', '総計|82']),
    },
  ],
  [
    'setPivotFieldSubtotalTop',
    {
      setup: twoLevels,
      act: (b) => b.wb.setPivotFieldSubtotalTop(0, b.pivot, b.region, false),
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
      act: (b) => b.wb.addPivotFieldItem(0, b.pivot, b.region, 'East', false),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
  [
    'addPivotFieldSharedItem',
    {
      // Only the index-addressed form can name the blank item.
      setup: (b) => {
        b.wb.addPivotCacheFieldSharedItemBlank(b.cacheId, 0);
        b.wb.setPivotCacheRecordBlank(b.cacheId, 0, 0);
      },
      act: (b) => b.wb.addPivotFieldSharedItem(0, b.pivot, b.region, 2, false),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
    },
  ],
  [
    'clearPivotFieldItems',
    {
      setup: hideEast,
      act: (b) => b.wb.clearPivotFieldItems(0, b.pivot, b.region),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'setPivotFieldItemVisible',
    {
      act: (b) => b.wb.setPivotFieldItemVisible(0, b.pivot, b.region, 1, false),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|15', '総計|15']),
    },
  ],
  [
    'addPivotFieldSubtotalFn',
    {
      setup: twoLevels,
      act: (b) => b.wb.addPivotFieldSubtotalFn(0, b.pivot, b.region, 3),
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
        b.wb.addPivotFieldSubtotalFn(0, b.pivot, b.region, 3);
      },
      act: (b) => b.wb.clearPivotFieldSubtotalFns(0, b.pivot, b.region),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), TWO_LEVEL_GRID),
    },
  ],
  [
    'setPivotFieldDateGroup',
    {
      setup: dateRows,
      act: (b) => b.wb.setPivotFieldDateGroup(0, b.pivot, b.date, 64, 1),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '令和6年|40', '令和7年|12', '総計|52']),
    },
  ],
  [
    'clearPivotFieldDateGroup',
    {
      setup: (b) => {
        dateRows(b);
        b.wb.setPivotFieldDateGroup(0, b.pivot, b.date, 64, 0);
      },
      act: (b) => b.wb.clearPivotFieldDateGroup(0, b.pivot, b.date),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), DATE_SERIAL_GRID),
    },
  ],
  [
    'setPivotFieldNumberFormat',
    {
      act: (b) => b.wb.setPivotFieldNumberFormat(0, b.pivot, b.amount, '4'),
      // A field's numFmtId is only observable in the part the writer emits.
      expect: (b) => {
        assert.equal(b.wb.setPivotFieldNumberFormat(0, b.pivot, 99, '4').status, 2);
        assert.equal(b.wb.setPivotFieldNumberFormat(0, b.pivot, b.amount, '0.00').status, 2);
        const saved = b.wb.save();
        assert.ok(saved.status.ok, JSON.stringify(saved.status));
        assert.match(packagePart(saved.bytes, 'xl/pivotTables/pivotTable1.xml'), /<pivotField [^>]*numFmtId="4"/);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'setPivotRowFieldOrder',
    {
      act: (b) => b.wb.setPivotRowFieldOrder(0, b.pivot, [b.product]),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'Apple|40', 'Pear|12', '総計|52']),
    },
  ],
  [
    'setPivotColFieldOrder',
    {
      act: (b) => b.wb.setPivotColFieldOrder(0, b.pivot, [b.product]),
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
      act: (b) => b.wb.addPivotDataField(0, b.pivot, { name: 'Max of Amount', fieldIndex: b.amount, aggregation: 3 }),
      expect: (b, r) => {
        assert.equal(r.index, 1);
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
      act: (b) => b.wb.clearPivotDataFields(0, b.pivot),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル', 'East', 'West', '総計']),
    },
  ],
  [
    'setPivotDataField',
    {
      act: (b) =>
        b.wb.setPivotDataField(0, b.pivot, 0, {
          name: 'Count of Amount',
          fieldIndex: b.amount,
          aggregation: 1,
          numberFormat: '2',
        }),
      expect: (b) => {
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Count of Amount', 'East|2', 'West|2', '総計|4']);
        const counts = b.wb.getPivotLayout(0, b.pivot).cells.filter((c) => c.value.kind === 1);
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
      act: (b) => b.wb.addPivotFilter(0, b.pivot, GREATER_THAN_20),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
  [
    'clearPivotFilters',
    {
      setup: (b) => b.wb.addPivotFilter(0, b.pivot, GREATER_THAN_20),
      act: (b) => b.wb.clearPivotFilters(0, b.pivot),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'removePivotFilter',
    {
      // Each filter alone keeps a different region, so removing the wrong
      // index leaves the wrong one.
      setup: (b) => {
        b.wb.addPivotFilter(0, b.pivot, GREATER_THAN_20);
        b.wb.addPivotFilter(0, b.pivot, BETWEEN_10_20);
      },
      act: (b) => b.wb.removePivotFilter(0, b.pivot, 1),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
  [
    'setPivotFieldDefaultSubtotal',
    {
      act: (b) => b.wb.setPivotFieldDefaultSubtotal(0, b.pivot, b.region, false),
      expect: (b) => assert.equal(b.wb.getPivotField(0, b.pivot, b.region).defaultSubtotal, false),
    },
  ],
  [
    'setPivotFieldInsertBlankRow',
    {
      act: (b) => b.wb.setPivotFieldInsertBlankRow(0, b.pivot, b.region, true),
      expect: (b) => assert.equal(b.wb.getPivotField(0, b.pivot, b.region).insertBlankRow, true),
    },
  ],
  [
    'setPivotFieldFillDownLabels',
    {
      act: (b) => b.wb.setPivotFieldFillDownLabels(0, b.pivot, b.region, true),
      expect: (b) => assert.equal(b.wb.getPivotField(0, b.pivot, b.region).fillDownLabels, true),
    },
  ],
  [
    'setPivotFieldShowDetail',
    {
      act: (b) => b.wb.setPivotFieldShowDetail(0, b.pivot, b.region, false),
      expect: (b) => assert.equal(b.wb.getPivotFieldItem(0, b.pivot, b.region, 0).showDetail, false),
    },
  ],
  [
    'setPivotFieldItemShowDetail',
    {
      act: (b) => b.wb.setPivotFieldItemShowDetail(0, b.pivot, b.region, 1, false),
      expect: (b) => {
        assert.equal(b.wb.getPivotFieldItem(0, b.pivot, b.region, 0).showDetail, true);
        assert.equal(b.wb.getPivotFieldItem(0, b.pivot, b.region, 1).showDetail, false);
      },
    },
  ],
  [
    'setPivotFieldItemCaption',
    {
      act: (b) => b.wb.setPivotFieldItemCaption(0, b.pivot, b.region, 0, 'Eastern'),
      expect: (b) => {
        assert.equal(b.wb.getPivotFieldItem(0, b.pivot, b.region, 0).caption, 'Eastern');
        assert.equal(b.wb.getPivotFieldItem(0, b.pivot, b.region, 1).caption, null);
      },
    },
  ],
  [
    'showPivotDetails',
    {
      act: (b) => b.wb.showPivotDetails(0, b.pivot, 1, 7),
      expect: (b, r) => {
        assert.equal(r.index, 0);
        assert.equal(b.wb.getSheetCount().value, 2);
      },
    },
  ],
  [
    'setPivotOptions',
    {
      setup: (b) => {
        b.before = b.wb.getPivotOptions(0, b.pivot);
      },
      act: (b) =>
        b.wb.setPivotOptions(0, b.pivot, { errorCaption: 'oops', indent: 3, pageWrap: 2, showHeaders: false }),
      expect: (b) => {
        const o = b.wb.getPivotOptions(0, b.pivot);
        assert.equal(o.errorCaption, 'oops');
        assert.equal(o.indent, 3);
        assert.equal(o.pageWrap, 2);
        assert.equal(o.showHeaders, false);
        // Omitted keys keep their value.
        assert.equal(o.showDrill, b.before.showDrill);
      },
    },
  ],
  [
    'setPivotCacheOptions',
    {
      act: (b) => b.wb.setPivotCacheOptions(b.cacheId, { saveData: false, missingItemsLimit: 5 }),
      expect: (b) => {
        const o = b.wb.getPivotCacheOptions(b.cacheId);
        assert.equal(o.saveData, false);
        assert.equal(o.missingItemsLimit, 5);
        assert.equal(o.refreshOnLoad, false);
      },
    },
  ],
  [
    'addPivotCacheCalculatedField',
    {
      act: (b) => b.wb.addPivotCacheCalculatedField(b.cacheId, 'Double', 'Amount*2'),
      expect: (b, r) => {
        assert.equal(r.index, 4);
        assert.equal(b.wb.getPivotCacheFieldName(b.cacheId, 4).value, 'Double');
        assert.equal(b.wb.getPivotCacheFieldFormula(b.cacheId, 4).value, 'Amount*2');
        assert.equal(b.wb.getPivotCacheFieldFormula(b.cacheId, 2).value, '');
      },
    },
  ],
  [
    'setPivotCacheCalculatedFieldFormula',
    {
      setup: (b) => b.wb.addPivotCacheCalculatedField(b.cacheId, 'Double', 'Amount*2'),
      act: (b) => b.wb.setPivotCacheCalculatedFieldFormula(b.cacheId, 4, 'Amount*3'),
      expect: (b) => assert.equal(b.wb.getPivotCacheFieldFormula(b.cacheId, 4).value, 'Amount*3'),
    },
  ],
  [
    'removePivotCacheCalculatedField',
    {
      setup: (b) => b.wb.addPivotCacheCalculatedField(b.cacheId, 'Double', 'Amount*2'),
      act: (b) => b.wb.removePivotCacheCalculatedField(b.cacheId, 4),
      expect: (b) => assert.equal(b.wb.getPivotCacheFieldCount(b.cacheId).value, 4),
    },
  ],
  [
    'addPivotCacheCalculatedItem',
    {
      act: (b) => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'Both', 'East+West'),
      expect: (b) => {
        assert.equal(b.wb.getPivotCacheCalculatedItemCount(b.cacheId).value, 1);
        const item = b.wb.getPivotCacheCalculatedItem(b.cacheId, 0);
        assert.ok(item.status.ok);
        assert.equal(item.fieldIndex, b.region);
        assert.equal(item.name, 'Both');
        assert.equal(item.formula, 'Region[East]+Region[West]');
      },
    },
  ],
  [
    'setPivotCacheCalculatedItemFormula',
    {
      setup: (b) => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'Both', 'East+West'),
      act: (b) => b.wb.setPivotCacheCalculatedItemFormula(b.cacheId, 0, 'East'),
      expect: (b) => assert.equal(b.wb.getPivotCacheCalculatedItem(b.cacheId, 0).formula, 'Region[East]'),
    },
  ],
  [
    'movePivotCacheCalculatedItem',
    {
      setup: (b) => {
        b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'First', 'East');
        b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'Second', 'West');
      },
      act: (b) => b.wb.movePivotCacheCalculatedItem(b.cacheId, 0, 1),
      expect: (b) => {
        assert.equal(b.wb.getPivotCacheCalculatedItem(b.cacheId, 0).name, 'Second');
        assert.equal(b.wb.getPivotCacheCalculatedItem(b.cacheId, 1).name, 'First');
      },
    },
  ],
  [
    'removePivotCacheCalculatedItem',
    {
      setup: (b) => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'Both', 'East+West'),
      act: (b) => b.wb.removePivotCacheCalculatedItem(b.cacheId, 0),
      expect: (b) => assert.equal(b.wb.getPivotCacheCalculatedItemCount(b.cacheId).value, 0),
    },
  ],
  [
    'setPivotFieldCustomName',
    {
      act: (b) => b.wb.setPivotFieldCustomName(0, b.pivot, b.region, 'Area'),
      expect: (b) => {
        assert.equal(b.wb.getPivotField(0, b.pivot, b.region).customName, 'Area');
        assert.ok(b.wb.setPivotFieldCustomName(0, b.pivot, b.region, '').ok);
        assert.equal(b.wb.getPivotField(0, b.pivot, b.region).customName, '');
      },
    },
  ],
  [
    'setPivotFieldShowAll',
    {
      act: (b) => b.wb.setPivotFieldShowAll(0, b.pivot, b.product, true),
      expect: (b) => {
        assert.equal(b.wb.getPivotField(0, b.pivot, b.product).showAll, true);
        assert.ok(b.wb.setPivotFieldShowAll(0, b.pivot, b.product, false).ok);
        assert.equal(b.wb.getPivotField(0, b.pivot, b.product).showAll, false);
      },
    },
  ],
  [
    'projectPivot',
    {
      act: (b) => b.wb.projectPivot(0, b.pivot),
      expect: (b) => {
        assert.equal(b.wb.isPivotProjectionPending(0, b.pivot).value, false);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
];

test('the pivot mutator table covers every declared pivot mutator', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'index.d.ts'), 'utf8');
  const declared = [...dts.matchAll(/^ {2}((?:[a-z]+Pivot|pivot)\w*)\([^)]*\): (?:Status|AddStyleResult);/gm)].map(
    (m) => m[1],
  );
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
      const status = typeof r.ok === 'boolean' ? r : r.status;
      assert.ok(status.ok, `${name}: ${JSON.stringify(status)}`);
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

test('manual pivot sort order keeps item order and rejects bad arguments', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    b.wb.clearPivotFieldItems(0, b.pivot, b.region);
    b.wb.addPivotFieldItem(0, b.pivot, b.region, 'West', true);
    b.wb.addPivotFieldItem(0, b.pivot, b.region, 'East', true);
    const manual = b.wb.setPivotFieldSort(0, b.pivot, b.region, mod.PivotSortOrder.Manual, '');
    assert.ok(manual.ok, JSON.stringify(manual));
    assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'West|37', 'East|15', '総計|52']);
    assert.equal(b.wb.setPivotFieldSort(0, b.pivot, b.region, 3, '').ok, false);
    assert.equal(b.wb.setPivotFieldSort(0, b.pivot, b.region, mod.PivotSortOrder.Manual, 'Sum of Amount').ok, false);
    assert.ok(b.wb.setPivotFieldSort(0, b.pivot, b.region, mod.PivotSortOrder.Ascending, '').ok);
    assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('pivot filters and subtotal functions survive save and reload', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    assert.ok(b.wb.addPivotFilter(0, b.pivot, GREATER_THAN_20).ok);
    assert.ok(b.wb.addPivotFieldSubtotalFn(0, b.pivot, b.region, 2).ok);
    assert.ok(b.wb.addPivotFieldSubtotalFn(0, b.pivot, b.region, 0).ok);
    const region = b.wb.getPivotField(0, b.pivot, b.region);
    assert.equal(region.defaultSubtotal, false);
    assert.deepEqual(region.subtotalFns, [0, 2]);
    assert.equal(b.wb.getPivotField(0, b.pivot, b.amount).defaultSubtotal, true);
    assert.deepEqual(b.wb.getPivotField(0, b.pivot, b.amount).subtotalFns, []);

    const saved = b.wb.save();
    assert.ok(saved.status.ok, JSON.stringify(saved.status));
    assert.match(
      packagePart(saved.bytes, 'xl/pivotTables/pivotTable1.xml'),
      /<filters count="1"><filter fld="0" type="valueGreaterThan" evalOrder="-1" id="1" iMeasureFld="0">/,
    );
    const back = mod.Workbook.load(saved.bytes);
    try {
      assert.ok(back.setExcelProfileId('win-365-ja_JP').ok);
      assert.equal(back.getPivotFilterCount(0, 0).value, 1);
      const filter = back.getPivotFilter(0, 0, 0);
      assert.ok(filter.status.ok, JSON.stringify(filter.status));
      assert.equal(filter.type, 1);
      assert.equal(filter.fieldName, 'Region');
      assert.equal(filter.valueKind, 1);
      assert.equal(filter.valueDouble, 20);
      assert.deepEqual(pivotGrid(back), WEST_ONLY);
      const field = back.getPivotField(0, 0, 0);
      assert.equal(field.defaultSubtotal, false);
      assert.deepEqual(field.subtotalFns, [0, 2]);
      assert.ok(back.clearPivotFilters(0, 0).ok);
      assert.doesNotMatch(packagePart(back.save().bytes, 'xl/pivotTables/pivotTable1.xml'), /<filters/);
    } finally {
      back.dispose();
    }
  } finally {
    b.wb.dispose();
  }
});

// ---- Status on the fallible accessors ----------------------------------
//
// Every accessor backed by a status-returning C ABI call carries that
// status, so a rejected argument or a released handle cannot read as a
// legitimate zero count, empty name, or unspilled cell. Each row is
// [method, args that succeed, args the engine rejects (or null when only
// a released handle can fail it)].
const STATUS_ACCESSORS = [
  ['getSheetCount', [], null],
  ['getCellCount', [0], [99]],
  ['definedNameCount', [], null],
  ['getTableCount', [], null],
  ['passthroughCount', [], null],
  ['getPivotCount', [0], [99]],
  ['pivotCacheCount', [], null],
  ['getPivotCacheFieldCount', null, [9999]],
  ['getPivotCacheFieldSharedItemCount', null, [9999, 0]],
  ['getPivotCacheRecordCount', null, [9999]],
  ['getPivotCacheCalculatedItemCount', null, [9999]],
  ['getPivotFieldCount', null, [99, 0]],
  ['getPivotFieldItemCount', null, [99, 0, 0]],
  ['getPivotName', null, [99, 0]],
  ['getPivotDataFieldCount', null, [99, 0]],
  ['getPivotFilterCount', null, [99, 0]],
  ['getFontCount', [], null],
  ['getFillCount', [], null],
  ['getBorderCount', [], null],
  ['getCellXfCount', [], null],
  ['getDxfCount', [], null],
  ['cellStyleCount', [], null],
  ['tableStyleCount', [], null],
  ['getCellStyleXfCount', [], null],
  ['getCalcMode', [], null],
  ['getFullPrecision', [], null],
  ['getCalcOnSave', [], null],
  ['getCommentVisible', null, [99, 0, 0]],
  ['getRefMode', [], null],
  ['getDate1904', [], null],
  ['getExcelProfileId', [], null],
  ['localizeFunctionName', ['SUM', 'win-365-en_US'], ['NOPE_XYZ', 'win-365-en_US']],
  ['canonicalizeFunctionName', ['SUM', 'win-365-en_US'], ['NOPE_XYZ', 'win-365-en_US']],
  ['getPinnedNow', [], null],
  ['getSpillInfo', [0, 0, 0], [99, 0, 0]],
  ['getArrayFormula', [0, 0, 0], [99, 0, 0]],
  ['getColumnWidthPt', [0, 0, 0], [99, 0, 0]],
  ['getRowHeightPt', [0, 0], [99, 0]],
  ['convertColumnCharsToPt', [0, 0, 8], [99, 0, 8]],
  ['convertColumnPtToChars', [0, 0, 10], [99, 0, 10]],
  ['getPrecedents', [0, 0, 0, 1], [99, 0, 0, 1]],
  ['getDependents', [0, 0, 0, 1], [99, 0, 0, 1]],
  ['listFunctionNames', [], null],
];

// The payload a failed call must fall back to, keyed by what it returns.
function assertFailurePayload(name, r) {
  if (Array.isArray(r)) assert.equal(r.length, 0, name);
  else if ('now' in r) assert.equal(r.now, null, name);
  else if ('engaged' in r) assert.equal(r.engaged, false, name);
  else if ('arrayFormula' in r) assert.equal(r.arrayFormula, null, name);
  else if (typeof r.value === 'string') assert.equal(r.value, '', name);
  else assert.equal(r.value, 0, name);
}

test('every NumberResult accessor is covered by the status table', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'index.d.ts'), 'utf8');
  const declared = [...dts.matchAll(/^ {2}([A-Za-z]\w*)\([^)]*\): NumberResult\b/gm)].map((m) => m[1]);
  const covered = new Set(STATUS_ACCESSORS.map(([name]) => name));
  assert.deepEqual(
    declared.filter((name) => !covered.has(name)),
    [],
    'NumberResult accessors missing from STATUS_ACCESSORS',
  );
});

test('fallible accessors report their status on success, rejection and a released handle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  for (const [name, okArgs, badArgs] of STATUS_ACCESSORS) {
    if (okArgs !== null) {
      const r = wb[name](...okArgs);
      assert.ok(r.status.ok, `${name}: ${JSON.stringify(r.status)}`);
    }
    if (badArgs !== null) {
      const r = wb[name](...badArgs);
      assert.equal(r.status.ok, false, `${name} must report the rejected arguments`);
      assert.notEqual(r.status.status, 0, name);
      assertFailurePayload(name, r);
    }
  }
  assert.equal(wb.getSheetCount().value, 1);
  assert.equal(wb.getExcelProfileId().value, 'win-365-en_US');
  assert.equal(wb.localizeFunctionName('SUM', 'win-365-ja_JP').value, 'SUM');
  assert.equal(wb.localizeFunctionName('SUM', 'nope').status.ok, false);

  wb.dispose();
  const dead = wb;
  // Every accessor, including the ones no argument can fail, reports the
  // released handle instead of a default reading.
  for (const [name, okArgs, badArgs] of STATUS_ACCESSORS) {
    if (name.endsWith('FunctionName') || name === 'listFunctionNames') continue;
    const r = dead[name](...(okArgs ?? badArgs));
    assert.equal(r.status.ok, false, `${name} after release`);
    assert.equal(r.status.status, 7000, name);
    assertFailurePayload(name, r);
  }
});

test('list getters carry a status and stop at the first element failure', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    // A readable sheet: every list is empty but the enumeration succeeded,
    // which is what distinguishes it from a sheet that could not be read.
    for (const [name, call] of [
      ['listMerges', () => wb.listMerges(0)],
      ['listComments', () => wb.listComments(0)],
      ['listHyperlinks', () => wb.listHyperlinks(0)],
      ['getValidations', () => wb.getValidations(0)],
      ['getConditionalFormats', () => wb.getConditionalFormats(0)],
      ['getExternalLinks', () => wb.getExternalLinks()],
      ['listDrawingObjects', () => wb.listDrawingObjects(0)],
    ]) {
      const list = call();
      assert.ok(Array.isArray(list), `${name} must stay an array`);
      assert.equal(list.length, 0, name);
      assert.ok(list.status.ok, `${name}: ${JSON.stringify(list.status)}`);
    }

    // An unreadable sheet is reported, not silently rendered as "empty".
    for (const [name, call] of [
      ['listMerges', () => wb.listMerges(99)],
      ['listComments', () => wb.listComments(99)],
      ['listHyperlinks', () => wb.listHyperlinks(99)],
      ['getValidations', () => wb.getValidations(99)],
      ['getConditionalFormats', () => wb.getConditionalFormats(99)],
      ['listDrawingObjects', () => wb.listDrawingObjects(99)],
    ]) {
      const list = call();
      assert.ok(Array.isArray(list), `${name} must stay an array`);
      assert.equal(list.status.ok, false, `${name} must signal an unreadable sheet`);
      assert.notEqual(list.status.status, 0, name);
    }
  } finally {
    wb.dispose();
  }
});

test('pivotFieldSetDateGroup defaults its trailing three arguments', async () => {
  const mod = await getModule();
  const explicit = buildPivot(mod);
  const defaulted = buildPivot(mod);
  try {
    dateRows(explicit);
    dateRows(defaulted);
    const days = 8; // PivotDatePeriod.Days
    assert.ok(explicit.wb.setPivotFieldDateGroup(0, explicit.pivot, explicit.date, days, 0, 1, NaN, NaN).ok);
    const r = defaulted.wb.setPivotFieldDateGroup(0, defaulted.pivot, defaulted.date, days, 0);
    assert.ok(r.ok, JSON.stringify(r));
    const field = defaulted.wb.getPivotField(0, defaulted.pivot, defaulted.date);
    assert.equal(field.datePeriods, days);
    assert.deepEqual(field, explicit.wb.getPivotField(0, explicit.pivot, explicit.date));
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
    assert.ok(b.wb.addPivotCacheFieldSharedItemDate(b.cacheId, 3, 45292).ok);
    assert.ok(b.wb.addPivotCacheFieldSharedItemDate(b.cacheId, 3, 45323.5).ok);
    assert.ok(b.wb.addPivotCacheFieldSharedItemNumber(b.cacheId, 3, 7).ok);
    for (const [idx, expected] of [
      [0, true],
      [1, true],
      [2, false],
    ]) {
      const r = b.wb.isPivotCacheFieldSharedItemDate(b.cacheId, 3, idx);
      assert.ok(r.status.ok, JSON.stringify(r.status));
      assert.equal(r.value, expected);
    }
    assert.equal(b.wb.isPivotCacheFieldSharedItemDate(b.cacheId, 3, 3).status.ok, false);
    for (const bad of [-1, Number.NaN]) {
      assert.equal(b.wb.addPivotCacheFieldSharedItemDate(b.cacheId, 3, bad).ok, false);
      assert.equal(b.wb.setPivotCacheRecordDate(b.cacheId, 0, 3, bad).ok, false);
    }
  } finally {
    b.wb.dispose();
  }
});

test('a pivot cache shared item reads back by index', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const east = b.wb.getPivotCacheFieldSharedItemValue(b.cacheId, 0, 0);
    assert.ok(east.status.ok, JSON.stringify(east.status));
    assert.equal(east.value.text, 'East');
    assert.equal(b.wb.getPivotCacheFieldSharedItemValue(b.cacheId, 0, 1).value.text, 'West');
    const count = b.wb.getPivotCacheFieldSharedItemCount(b.cacheId, 0).value;
    const past = b.wb.getPivotCacheFieldSharedItemValue(b.cacheId, 0, count);
    assert.equal(past.status.ok, false);
  } finally {
    b.wb.dispose();
  }
});

test('a pivot cache error setter without errorCode throws and changes nothing', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const items = b.wb.getPivotCacheFieldSharedItemCount(b.cacheId, 0).value;
    assert.throws(() => b.wb.addPivotCacheFieldSharedItemError(b.cacheId, 0), TypeError);
    assert.equal(b.wb.getPivotCacheFieldSharedItemCount(b.cacheId, 0).value, items);
    assert.throws(() => b.wb.setPivotCacheRecordError(b.cacheId, 0, 2), TypeError);
    assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('pivot report options and the Values field order report bad arguments', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    assert.equal(mod.PIVOT_VALUES_FIELD, -2);
    // A single data field has no Values level to place.
    assert.equal(b.wb.setPivotRowFieldOrder(0, b.pivot, [b.region, mod.PIVOT_VALUES_FIELD]).ok, false);
    assert.equal(b.wb.setPivotFieldDefaultSubtotal(0, b.pivot, 99, true).ok, false);
    assert.equal(b.wb.setPivotFieldInsertBlankRow(0, b.pivot, 99, true).ok, false);
    assert.equal(b.wb.setPivotFieldFillDownLabels(0, b.pivot, 99, true).ok, false);
    assert.equal(b.wb.setPivotFieldShowDetail(0, b.pivot, 99, true).ok, false);
    assert.equal(b.wb.setPivotFieldItemShowDetail(0, b.pivot, b.region, 99, true).ok, false);
    assert.equal(b.wb.setPivotFieldItemCaption(0, b.pivot, b.region, 99, 'x').ok, false);
    // Captions collide with another item's label.
    assert.equal(b.wb.setPivotFieldItemCaption(0, b.pivot, b.region, 0, 'West').ok, false);
    assert.equal(b.wb.showPivotDetails(99, 0, 0, 0).status.ok, false);
    assert.equal(b.wb.setDrawingName(0, 9999, 'x').ok, false);
    assert.equal(b.wb.setDrawingName(0, 9999, '').ok, false);
  } finally {
    b.wb.dispose();
  }
});

test('pivot options read back every member and keep what a partial write omits', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const before = b.wb.getPivotOptions(0, b.pivot);
    assert.ok(before.status.ok, JSON.stringify(before.status));
    assert.equal(typeof before.errorCaption, 'string');
    assert.equal(typeof before.indent, 'number');
    const patch = {
      showError: true,
      errorCaption: '#n/a',
      showMissing: false,
      missingCaption: 'none',
      gridDropZones: true,
      mergeItem: true,
      pageOverThenDown: true,
      altText: 'Sales',
      altTextSummary: 'line one\nline two',
      grandTotalCaption: 'Everything',
      autoFormatting: false,
      enableDrill: false,
    };
    assert.ok(b.wb.setPivotOptions(0, b.pivot, patch).ok);
    const after = b.wb.getPivotOptions(0, b.pivot);
    for (const [key, value] of Object.entries(patch)) assert.equal(after[key], value, key);
    assert.equal(after.showHeaders, before.showHeaders);
    assert.equal(after.indent, before.indent);
    // A write of the whole read-back is a no-op, and '' clears a string.
    const { status, ...whole } = after;
    assert.ok(b.wb.setPivotOptions(0, b.pivot, { ...whole, errorCaption: '' }).ok);
    assert.equal(b.wb.getPivotOptions(0, b.pivot).errorCaption, '');
    assert.equal(b.wb.getPivotOptions(0, b.pivot).altText, 'Sales');
    assert.ok(b.wb.setPivotOptions(0, b.pivot, { grandTotalCaption: '' }).ok);
    assert.equal(b.wb.getPivotOptions(0, b.pivot).grandTotalCaption, '');
  } finally {
    b.wb.dispose();
  }
});

test('pivot options reject out-of-range values and wrong types without changing anything', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const base = b.wb.getPivotOptions(0, b.pivot);
    const unchanged = () => assert.deepEqual(b.wb.getPivotOptions(0, b.pivot), base);
    for (const bad of [{ indent: 127 }, { pageWrap: 256 }, { errorCaption: 'x'.repeat(256) }]) {
      assert.equal(b.wb.setPivotOptions(0, b.pivot, bad).ok, false, JSON.stringify(bad).slice(0, 40));
      unchanged();
    }
    for (const bad of [{ indent: 'a' }, { errorCaption: 5 }]) {
      assert.throws(() => b.wb.setPivotOptions(0, b.pivot, bad), TypeError, JSON.stringify(bad));
      unchanged();
    }
    for (const bad of [{ indent: 1.5 }, { indent: -1 }, { pageWrap: Number.NaN }]) {
      assert.throws(() => b.wb.setPivotOptions(0, b.pivot, bad), RangeError, JSON.stringify(bad));
      unchanged();
    }
    assert.throws(() => b.wb.setPivotOptions(0, b.pivot, 'x'), TypeError);
    assert.throws(() => b.wb.setPivotOptions(0, b.pivot), TypeError);
    assert.throws(() => b.wb.setPivotOptions('0', b.pivot, {}), TypeError);
    assert.throws(() => b.wb.setPivotOptions(2 ** 32, b.pivot, {}));
    unchanged();
    assert.equal(b.wb.setPivotOptions(99, 0, {}).ok, false);
    assert.equal(b.wb.getPivotOptions(99, 0).status.ok, false);

    const cacheBase = b.wb.getPivotCacheOptions(b.cacheId);
    for (const limit of [-2, 1048577]) {
      assert.equal(b.wb.setPivotCacheOptions(b.cacheId, { missingItemsLimit: limit }).ok, false, String(limit));
    }
    assert.throws(() => b.wb.setPivotCacheOptions(b.cacheId, { missingItemsLimit: 'a' }), TypeError);
    for (const bad of [{ missingItemsLimit: 1.5 }, { missingItemsLimit: Number.NaN }]) {
      assert.throws(() => b.wb.setPivotCacheOptions(b.cacheId, bad), RangeError, JSON.stringify(bad));
    }
    assert.throws(() => b.wb.setPivotCacheOptions(b.cacheId, null), TypeError);
    assert.throws(() => b.wb.setPivotCacheOptions('1', {}), TypeError);
    assert.deepEqual(b.wb.getPivotCacheOptions(b.cacheId), cacheBase);
    assert.equal(b.wb.setPivotCacheOptions(9999, {}).ok, false);
    assert.equal(b.wb.getPivotCacheOptions(9999).status.ok, false);
  } finally {
    b.wb.dispose();
  }
});

test('calculated fields and items reject bad arguments and change nothing', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const fields = b.wb.getPivotCacheFieldCount(b.cacheId).value;
    // Wrong types throw before anything runs.
    assert.throws(() => b.wb.addPivotCacheCalculatedField(b.cacheId, 'D', 5), TypeError);
    assert.throws(() => b.wb.addPivotCacheCalculatedField(b.cacheId, null, 'Amount'), TypeError);
    assert.throws(() => b.wb.addPivotCacheCalculatedField(b.cacheId, 'D'), TypeError);
    assert.throws(() => b.wb.addPivotCacheCalculatedField('1', 'D', 'Amount'), TypeError);
    assert.throws(() => b.wb.addPivotCacheCalculatedItem(b.cacheId, 0.5, 'N', 'East'), RangeError);
    assert.throws(() => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'N', undefined), TypeError);
    assert.throws(() => b.wb.movePivotCacheCalculatedItem(b.cacheId, 0, -1), RangeError);
    assert.throws(() => b.wb.getPivotCacheCalculatedItem(b.cacheId, 'a'), TypeError);
    // Rejected by the engine: empty or reserved names, bad formulas.
    for (const [name, formula] of [
      ['', 'Amount'],
      ['All', 'Amount'],
      ['A[1]', 'Amount'],
      ['Amount', 'Amount'],
      ['D', ''],
      ['D', 'Amount+'],
      ['D', 'Nope*2'],
      ['D', 'A1+Amount'],
      ['D', 'x'.repeat(8193)],
    ]) {
      assert.equal(
        b.wb.addPivotCacheCalculatedField(b.cacheId, name, formula).status.ok,
        false,
        `${name}=${formula.slice(0, 8)}`,
      );
    }
    assert.equal(b.wb.addPivotCacheCalculatedField(9999, 'D', 'Amount').status.ok, false);
    assert.equal(b.wb.getPivotCacheFieldCount(b.cacheId).value, fields);
    // Not a calculated field.
    assert.equal(b.wb.setPivotCacheCalculatedFieldFormula(b.cacheId, 0, 'Amount').ok, false);
    assert.equal(b.wb.removePivotCacheCalculatedField(b.cacheId, 0).ok, false);
    assert.equal(b.wb.removePivotCacheCalculatedField(b.cacheId, 2 ** 31).ok, false);
    assert.equal(b.wb.getPivotCacheFieldFormula(b.cacheId, 99).status.ok, false);
    assert.equal(b.wb.getPivotCacheFieldFormula(9999, 0).status.ok, false);
    // A calculated field a second one names cannot be removed.
    assert.ok(b.wb.addPivotCacheCalculatedField(b.cacheId, 'Base', 'Amount').status.ok);
    assert.ok(b.wb.addPivotCacheCalculatedField(b.cacheId, 'Top', 'Base*2').status.ok);
    assert.equal(b.wb.removePivotCacheCalculatedField(b.cacheId, fields).ok, false);
    assert.equal(b.wb.getPivotCacheFieldCount(b.cacheId).value, fields + 2);

    // Items: unplaced fields, duplicates, and out-of-range indices.
    assert.equal(b.wb.addPivotCacheCalculatedItem(b.cacheId, b.date, 'N', 'Amount').ok, false);
    assert.equal(b.wb.addPivotCacheCalculatedItem(b.cacheId, 99, 'N', 'East').ok, false);
    assert.equal(b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'East', 'West').ok, false);
    assert.equal(b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'N', 'Nope').ok, false);
    assert.equal(b.wb.addPivotCacheCalculatedItem(9999, b.region, 'N', 'East').ok, false);
    assert.equal(b.wb.getPivotCacheCalculatedItemCount(b.cacheId).value, 0);
    assert.equal(b.wb.getPivotCacheCalculatedItem(b.cacheId, 0).status.ok, false);
    assert.equal(b.wb.setPivotCacheCalculatedItemFormula(b.cacheId, 0, 'East').ok, false);
    assert.equal(b.wb.movePivotCacheCalculatedItem(b.cacheId, 0, 0).ok, false);
    assert.equal(b.wb.removePivotCacheCalculatedItem(b.cacheId, 0).ok, false);
    assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('pivot field custom names, show-all and default data field names report bad arguments', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    assert.throws(() => b.wb.setPivotFieldCustomName(0, b.pivot, b.region, null), TypeError);
    assert.throws(() => b.wb.setPivotFieldCustomName(0, b.pivot, b.region), TypeError);
    assert.throws(() => b.wb.setPivotFieldCustomName(0, b.pivot, -1, 'x'), RangeError);
    assert.throws(() => b.wb.setPivotFieldShowAll(0, b.pivot, b.region), TypeError);
    assert.equal(b.wb.setPivotFieldCustomName(0, b.pivot, 99, 'x').ok, false);
    assert.equal(b.wb.setPivotFieldCustomName(0, b.pivot, b.region, 'x'.repeat(256)).ok, false);
    // Another field's source name or display name is taken.
    assert.equal(b.wb.setPivotFieldCustomName(0, b.pivot, b.region, 'Product').ok, false);
    assert.equal(b.wb.setPivotFieldCustomName(0, b.pivot, b.region, 'Sum of Amount').ok, false);
    // The field's own source name is accepted.
    assert.ok(b.wb.setPivotFieldCustomName(0, b.pivot, b.region, 'Region').ok);
    assert.equal(b.wb.setPivotFieldShowAll(0, b.pivot, 99, true).ok, false);

    // The first derived name is free; once taken, a new one is numbered
    // unless the entry that holds it is the one being replaced.
    assert.equal(b.wb.getPivotDataFieldDefaultName(0, b.pivot, b.amount, 0).value, '合計 / Amount');
    assert.ok(b.wb.addPivotDataField(0, b.pivot, { fieldIndex: b.amount, aggregation: 0 }).status.ok);
    assert.equal(b.wb.getPivotDataField(0, b.pivot, 1).name, '合計 / Amount');
    assert.equal(b.wb.getPivotDataFieldDefaultName(0, b.pivot, b.amount, 0).value, '合計 / Amount2');
    assert.equal(b.wb.getPivotDataFieldDefaultName(0, b.pivot, b.amount, 0, undefined).value, '合計 / Amount2');
    assert.equal(b.wb.getPivotDataFieldDefaultName(0, b.pivot, b.amount, 0, 1).value, '合計 / Amount');
    assert.equal(b.wb.getPivotDataFieldDefaultName(0, b.pivot, b.amount, 0, 99).status.ok, false);
    assert.equal(b.wb.getPivotDataFieldCount(0, b.pivot).value, 2);
    assert.equal(b.wb.getPivotDataFieldDefaultName(0, b.pivot, 99, 0).status.ok, false);
    assert.equal(b.wb.getPivotDataFieldDefaultName(0, b.pivot, b.amount, 99).status.ok, false);
    assert.equal(b.wb.getPivotDataFieldDefaultName(99, 0, 0, 0).status.ok, false);
    assert.throws(() => b.wb.getPivotDataFieldDefaultName(0, b.pivot, b.amount, 'sum'), TypeError);
    assert.throws(() => b.wb.getPivotDataFieldDefaultName(0, b.pivot, b.amount, 0, null), TypeError);
    assert.throws(() => b.wb.getPivotDataFieldDefaultName(0, b.pivot, b.amount), TypeError);
    assert.equal(b.wb.getPivotDataField(0, b.pivot, 0).name, 'Sum of Amount');
    assert.equal(b.wb.projectPivot(99, 0).ok, false);
    assert.equal(b.wb.isPivotProjectionPending(99, 0).status.ok, false);
    assert.throws(() => b.wb.projectPivot(0), TypeError);
  } finally {
    b.wb.dispose();
  }
});

test('label-between filters carry valueHighText', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const spec = {
      axis: 0,
      fieldName: 'Region',
      type: mod.PivotFilterType.LabelBetween,
      valueKind: 2,
      valueText: 'A',
      valueHighKind: 2,
    };
    // The upper bound is required for the between types.
    assert.equal(b.wb.addPivotFilter(0, b.pivot, spec).ok, false);
    assert.ok(b.wb.addPivotFilter(0, b.pivot, { ...spec, valueHighText: 'F' }).ok);
    const at = b.wb.getPivotFilter(0, b.pivot, 0);
    assert.equal(at.type, 22);
    assert.equal(at.valueText, 'A');
    assert.equal(at.valueHighText, 'F');
    // And rejected for every other type.
    assert.equal(
      b.wb.addPivotFilter(0, b.pivot, { ...spec, type: mod.PivotFilterType.LabelEqual, valueHighText: 'F' }).ok,
      false,
    );
    assert.throws(() => b.wb.addPivotFilter(0, b.pivot, { ...spec, valueHighText: 5 }), TypeError);
    assert.equal(b.wb.getPivotFilterCount(0, b.pivot).value, 1);
  } finally {
    b.wb.dispose();
  }
});
