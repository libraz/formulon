import test from 'node:test';
import { inflateRawSync } from 'node:zlib';
import { assert, getModule, isFormulonError, path, pkgRoot, readFile } from './smoke_support.mjs';

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
  const pivot = wb.sheet(0).addPivot({ name: 'Pivot1', cacheId, anchorRow: 0, anchorCol: 6 });
  const addField = (spec) => wb.sheet(0).addPivotField(pivot, spec);
  const region = addField({ sourceName: 'Region', axis: 0 });
  wb.sheet(0).addPivotFieldItem(pivot, region, 'East');
  wb.sheet(0).addPivotFieldItem(pivot, region, 'West');
  const product = addField({ sourceName: 'Product', axis: 2 });
  wb.sheet(0).addPivotFieldItem(pivot, product, 'Apple');
  wb.sheet(0).addPivotFieldItem(pivot, product, 'Pear');
  const amount = addField({ sourceName: 'Amount', axis: 2 });
  const date = addField({ sourceName: 'Date', axis: 2 });
  wb.sheet(0).setPivotRowFieldOrder(pivot, [region]);
  wb.sheet(0).addPivotDataField(pivot, { name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 });
  return { wb, mod, cacheId, pivot, region, product, amount, date };
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
        assert.throws(() => b.wb.sheet(0).getPivotLayout(b.pivot), isFormulonError);
      },
    },
  ],
  [
    'addPivotCacheFieldSharedItem (number)',
    {
      act: (b) => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, 42.5),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '42.5|10', 'East|5', 'West|37', '総計|52']),
    },
  ],
  [
    'addPivotCacheFieldSharedItem (date)',
    {
      act: (b) => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 3, 45292, { date: true }),
      expect: (b) => assert.equal(b.wb.listPivotCacheFieldSharedItems(b.cacheId, 3)[0].isDate, true),
    },
  ],
  [
    'addPivotCacheFieldSharedItem',
    {
      act: (b) => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, 'North'),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'North|10', 'West|37', '総計|52']),
    },
  ],
  [
    'addPivotCacheFieldSharedItem (bool)',
    {
      act: (b) => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, true),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', 'TRUE|10', '総計|52']),
    },
  ],
  [
    'addPivotCacheFieldSharedItem (blank)',
    {
      act: (b) => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, null),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '(空白)|10', '総計|52']),
    },
  ],
  [
    'addPivotCacheFieldSharedItem (error)',
    {
      act: (b) => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, { kind: b.mod.ValueKind.Error, errorCode: 6 }),
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
      act: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 2, 1000),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|1005', 'West|37', '総計|1042']),
    },
  ],
  [
    'setPivotCacheRecordValue (date)',
    {
      act: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 3, 45300, { date: true }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'setPivotCacheRecordValue (text)',
    {
      act: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 0, 'West'),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|47', '総計|52']),
    },
  ],
  [
    'setPivotCacheRecordValue (bool)',
    {
      act: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 2, true),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
    },
  ],
  [
    'setPivotCacheRecordValue (blank)',
    {
      act: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 2, null),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
    },
  ],
  [
    'setPivotCacheRecordValue (error)',
    {
      act: (b) => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 2, { kind: b.mod.ValueKind.Error, errorCode: 1 }),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|#1', 'West|37', '総計|#1']),
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
        assert.throws(() => b.wb.sheet(0).getPivotLayout(0), isFormulonError);
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
    'updatePivotField (subtotalTop)',
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
    'updatePivotItem (visible)',
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
    'updatePivotField (numberFormat)',
    {
      act: (b) => b.wb.sheet(0).updatePivotField(b.pivot, b.amount, { numberFormat: '4' }),
      // A field's numFmtId is only observable in the part the writer emits.
      expect: (b) => {
        assert.throws(
          () => b.wb.sheet(0).updatePivotField(b.pivot, 99, { numberFormat: '4' }),
          (e) => isFormulonError(e, 5054),
        );
        assert.throws(
          () => b.wb.sheet(0).updatePivotField(b.pivot, b.amount, { numberFormat: '0.00' }),
          (e) => isFormulonError(e, 2),
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
  [
    'updatePivotField (defaultSubtotal)',
    {
      act: (b) => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { defaultSubtotal: false }),
      expect: (b) => assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.region].defaultSubtotal, false),
    },
  ],
  [
    'updatePivotField (insertBlankRow)',
    {
      act: (b) => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { insertBlankRow: true }),
      expect: (b) => assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.region].insertBlankRow, true),
    },
  ],
  [
    'updatePivotField (fillDownLabels)',
    {
      act: (b) => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { fillDownLabels: true }),
      expect: (b) => assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.region].fillDownLabels, true),
    },
  ],
  [
    'setPivotFieldShowDetail',
    {
      act: (b) => b.wb.sheet(0).setPivotFieldShowDetail(b.pivot, b.region, false),
      expect: (b) => assert.equal(b.wb.sheet(0).getPivotFieldItem(b.pivot, b.region, 0).showDetail, false),
    },
  ],
  [
    'updatePivotItem (showDetail)',
    {
      act: (b) => b.wb.sheet(0).updatePivotItem(b.pivot, b.region, 1, { showDetail: false }),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).getPivotFieldItem(b.pivot, b.region, 0).showDetail, true);
        assert.equal(b.wb.sheet(0).getPivotFieldItem(b.pivot, b.region, 1).showDetail, false);
      },
    },
  ],
  [
    'updatePivotItem (caption)',
    {
      act: (b) => b.wb.sheet(0).updatePivotItem(b.pivot, b.region, 0, { caption: 'Eastern' }),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).getPivotFieldItem(b.pivot, b.region, 0).caption, 'Eastern');
        assert.equal(b.wb.sheet(0).getPivotFieldItem(b.pivot, b.region, 1).caption, null);
      },
    },
  ],
  [
    'showPivotDetails',
    {
      act: (b) => b.wb.sheet(0).showPivotDetails(b.pivot, 1, 7),
      expect: (b, r) => {
        assert.equal(r, b.wb.sheet(0));
        assert.equal(r.index, 0);
        assert.equal(b.wb.listSheets().length, 2);
      },
    },
  ],
  [
    'setPivotOptions',
    {
      setup: (b) => {
        b.before = b.wb.sheet(0).getPivotOptions(b.pivot);
      },
      act: (b) =>
        b.wb.sheet(0).setPivotOptions(b.pivot, { errorCaption: 'oops', indent: 3, pageWrap: 2, showHeaders: false }),
      expect: (b) => {
        const o = b.wb.sheet(0).getPivotOptions(b.pivot);
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
        assert.equal(r, 4);
        const fields = b.wb.listPivotCacheFields(b.cacheId);
        assert.equal(fields[4].name, 'Double');
        assert.equal(fields[4].formula, '=Amount*2');
        assert.equal(fields[2].formula, null);
      },
    },
  ],
  [
    'setPivotCacheCalculatedFieldFormula',
    {
      setup: (b) => b.wb.addPivotCacheCalculatedField(b.cacheId, 'Double', 'Amount*2'),
      act: (b) => b.wb.setPivotCacheCalculatedFieldFormula(b.cacheId, 4, 'Amount*3'),
      expect: (b) => assert.equal(b.wb.listPivotCacheFields(b.cacheId)[4].formula, '=Amount*3'),
    },
  ],
  [
    'removePivotCacheCalculatedField',
    {
      setup: (b) => b.wb.addPivotCacheCalculatedField(b.cacheId, 'Double', 'Amount*2'),
      act: (b) => b.wb.removePivotCacheCalculatedField(b.cacheId, 4),
      expect: (b) => assert.equal(b.wb.listPivotCacheFields(b.cacheId).length, 4),
    },
  ],
  [
    'addPivotCacheCalculatedItem',
    {
      act: (b) => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'Both', 'East+West'),
      expect: (b) => {
        assert.equal(b.wb.listPivotCacheCalculatedItems(b.cacheId).length, 1);
        const item = b.wb.listPivotCacheCalculatedItems(b.cacheId)[0];
        assert.equal(item.fieldIndex, b.region);
        assert.equal(item.name, 'Both');
        assert.equal(item.formula, '=Region[East]+Region[West]');
      },
    },
  ],
  [
    'setPivotCacheCalculatedItemFormula',
    {
      setup: (b) => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'Both', 'East+West'),
      act: (b) => b.wb.setPivotCacheCalculatedItemFormula(b.cacheId, 0, 'East'),
      expect: (b) => assert.equal(b.wb.listPivotCacheCalculatedItems(b.cacheId)[0].formula, '=Region[East]'),
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
        assert.equal(b.wb.listPivotCacheCalculatedItems(b.cacheId)[0].name, 'Second');
        assert.equal(b.wb.listPivotCacheCalculatedItems(b.cacheId)[1].name, 'First');
      },
    },
  ],
  [
    'removePivotCacheCalculatedItem',
    {
      setup: (b) => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'Both', 'East+West'),
      act: (b) => b.wb.removePivotCacheCalculatedItem(b.cacheId, 0),
      expect: (b) => assert.equal(b.wb.listPivotCacheCalculatedItems(b.cacheId).length, 0),
    },
  ],
  [
    'updatePivotField (customName)',
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
    'updatePivotField (showAll)',
    {
      act: (b) => b.wb.sheet(0).updatePivotField(b.pivot, b.product, { showAll: true }),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.product].showAll, true);
        b.wb.sheet(0).updatePivotField(b.pivot, b.product, { showAll: false });
        assert.equal(b.wb.sheet(0).listPivotFields(b.pivot)[b.product].showAll, false);
      },
    },
  ],
  [
    'projectPivot',
    {
      act: (b) => b.wb.sheet(0).projectPivot(b.pivot),
      expect: (b) => {
        assert.equal(b.wb.sheet(0).isPivotProjectionPending(b.pivot), false);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
];

test('the pivot mutator table covers every declared pivot mutator', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'index.d.ts'), 'utf8');
  const declaredAll = [...dts.matchAll(/^ {2}((?:[a-z]+Pivot|pivot)\w*)\([^)]*\): (?:void|number|Sheet);/gm)].map(
    (m) => m[1],
  );
  const declared = declaredAll.filter((name) => !/^(?:get|is|list)/.test(name));
  assert.ok(declared.length > 40, `expected the pivot mutator family, found ${declared.length}`);
  const covered = PIVOT_MUTATORS.map(([name]) => name.replace(/ \(.*\)$/, ''));
  assert.deepEqual(
    declared.filter((name) => !covered.includes(name)),
    [],
    'pivot mutators without a row',
  );
  assert.deepEqual(
    covered.filter((name) => !declaredAll.includes(name)),
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

test('manual pivot sort order keeps item order and rejects bad arguments', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    b.wb.sheet(0).clearPivotFieldItems(b.pivot, b.region);
    b.wb.sheet(0).addPivotFieldItem(b.pivot, b.region, 'West');
    b.wb.sheet(0).addPivotFieldItem(b.pivot, b.region, 'East');
    b.wb.sheet(0).setPivotFieldSort(b.pivot, b.region, mod.PivotSortOrder.Manual);
    assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'West|37', 'East|15', '総計|52']);
    assert.throws(() => b.wb.sheet(0).setPivotFieldSort(b.pivot, b.region, 3), isFormulonError);
    assert.throws(
      () => b.wb.sheet(0).setPivotFieldSort(b.pivot, b.region, mod.PivotSortOrder.Manual, { byField: 'Sum of Amount' }),
      isFormulonError,
    );
    b.wb.sheet(0).setPivotFieldSort(b.pivot, b.region, mod.PivotSortOrder.Ascending);
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

// ---- Failures on the fallible accessors --------------------------------
//
// Every accessor backed by a fallible C ABI call throws, so a rejected
// argument or a released handle cannot read as a legitimate zero count,
// empty name, or unspilled cell. Each row is
// [method, args that succeed, args the engine rejects (or null when only
// a released handle can fail it)].
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

// Calls a STATUS_ACCESSORS row on the workbook, or on the sheet its first argument names.
function callAccessor(wb, name, args) {
  return name in wb ? wb[name](...args) : wb.sheet(args[0])[name](...args.slice(1));
}

test('every row of the accessor table names a declared Workbook or Sheet method', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'index.d.ts'), 'utf8');
  const declared = new Set([...dts.matchAll(/^ {2}([A-Za-z]\w*)\(/gm)].map((m) => m[1]));
  assert.deepEqual(
    STATUS_ACCESSORS.map(([name]) => name).filter((name) => !declared.has(name)),
    [],
    'STATUS_ACCESSORS rows for methods that are not declared',
  );
});

test('fallible accessors succeed, throw on a rejection and on a released handle', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  for (const [name, okArgs, badArgs] of STATUS_ACCESSORS) {
    if (okArgs !== null) {
      assert.doesNotThrow(() => callAccessor(wb, name, okArgs), name);
    }
    if (badArgs !== null) {
      assert.throws(
        () => callAccessor(wb, name, badArgs),
        isFormulonError,
        `${name} must report the rejected arguments`,
      );
    }
  }
  assert.equal(wb.listSheets().length, 1);
  assert.equal(wb.getExcelProfileId(), 'win-365-en_US');
  assert.equal(mod.localizeFunctionName('SUM', 'win-365-ja_JP'), 'SUM');
  assert.throws(() => mod.localizeFunctionName('SUM', 'nope'), isFormulonError);
  assert.equal(mod.listFunctionNames().includes('SUM'), true);
  assert.throws(() => mod.listLocaleErrorNames('nope'), isFormulonError);
  assert.throws(() => mod.compareText('a', 'b', 0, 'nope'), isFormulonError);

  wb.dispose();
  const dead = wb;
  // Every accessor, including the ones no argument can fail, reports the
  // released handle instead of a default reading.
  for (const [name, okArgs, badArgs] of STATUS_ACCESSORS) {
    assert.throws(
      () => callAccessor(dead, name, okArgs ?? badArgs),
      (e) => isFormulonError(e, 7000),
      `${name} after release`,
    );
  }
});

test('list getters return arrays and throw on an unreadable sheet', async () => {
  const mod = await getModule();
  const wb = mod.Workbook.createDefault();
  try {
    // A readable sheet: every list is empty but the enumeration succeeded,
    // which is what distinguishes it from a sheet that could not be read.
    for (const [name, call] of [
      ['listMerges', () => wb.sheet(0).listMerges()],
      ['listComments', () => wb.sheet(0).listComments()],
      ['listHyperlinks', () => wb.sheet(0).listHyperlinks()],
      ['listDrawingObjects', () => wb.sheet(0).listDrawingObjects()],
    ]) {
      const list = call();
      assert.ok(Array.isArray(list), `${name} must stay an array`);
      assert.equal(list.length, 0, name);
    }

    // An unreadable sheet is reported, not silently rendered as "empty".
    for (const [name, call] of [
      ['listMerges', () => wb.sheet(99).listMerges()],
      ['listComments', () => wb.sheet(99).listComments()],
      ['listHyperlinks', () => wb.sheet(99).listHyperlinks()],
      ['listDrawingObjects', () => wb.sheet(99).listDrawingObjects()],
    ]) {
      assert.throws(call, isFormulonError, `${name} must signal an unreadable sheet`);
    }
  } finally {
    wb.dispose();
  }
});

test('setPivotFieldDateGroup defaults its optional group fields', async () => {
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
    for (const bad of [-1, Number.NaN]) {
      assert.throws(() => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 3, bad, { date: true }), isFormulonError);
      assert.throws(() => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 3, bad, { date: true }), isFormulonError);
    }
    // A date flag only qualifies a number.
    assert.throws(() => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 3, 'x', { date: true }), TypeError);
    assert.throws(() => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 3, true, { date: true }), TypeError);
    assert.equal(b.wb.listPivotCacheFieldSharedItems(b.cacheId, 3).length, 3);
  } finally {
    b.wb.dispose();
  }
});

test('a pivot cache shared item list reads back in order', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const items = b.wb.listPivotCacheFieldSharedItems(b.cacheId, 0);
    assert.deepEqual(
      items.map((item) => [item.value.text, item.isDate]),
      [
        ['East', false],
        ['West', false],
      ],
    );
    assert.throws(() => b.wb.listPivotCacheFieldSharedItems(b.cacheId, 99), isFormulonError);
  } finally {
    b.wb.dispose();
  }
});

test('a pivot cache error value without errorCode throws and changes nothing', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const items = b.wb.listPivotCacheFieldSharedItems(b.cacheId, 0).length;
    const partial = { kind: mod.ValueKind.Error };
    assert.throws(() => b.wb.addPivotCacheFieldSharedItem(b.cacheId, 0, partial), TypeError);
    assert.equal(b.wb.listPivotCacheFieldSharedItems(b.cacheId, 0).length, items);
    assert.throws(() => b.wb.setPivotCacheRecordValue(b.cacheId, 0, 2, partial), TypeError);
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
    assert.throws(
      () => b.wb.sheet(0).setPivotRowFieldOrder(b.pivot, [b.region, mod.PIVOT_VALUES_FIELD]),
      isFormulonError,
    );
    assert.throws(() => b.wb.sheet(0).updatePivotField(b.pivot, 99, { defaultSubtotal: true }), isFormulonError);
    assert.throws(() => b.wb.sheet(0).updatePivotField(b.pivot, 99, { insertBlankRow: true }), isFormulonError);
    assert.throws(() => b.wb.sheet(0).updatePivotField(b.pivot, 99, { fillDownLabels: true }), isFormulonError);
    assert.throws(() => b.wb.sheet(0).setPivotFieldShowDetail(b.pivot, 99, true), isFormulonError);
    assert.throws(() => b.wb.sheet(0).updatePivotItem(b.pivot, b.region, 99, { showDetail: true }), isFormulonError);
    assert.throws(() => b.wb.sheet(0).updatePivotItem(b.pivot, b.region, 99, { caption: 'x' }), isFormulonError);
    // Captions collide with another item's label.
    assert.throws(() => b.wb.sheet(0).updatePivotItem(b.pivot, b.region, 0, { caption: 'West' }), isFormulonError);
    assert.throws(() => b.wb.sheet(99).showPivotDetails(0, 0, 0), isFormulonError);
    assert.throws(() => b.wb.sheet(0).setDrawingName(9999, 'x'), isFormulonError);
    assert.throws(() => b.wb.sheet(0).setDrawingName(9999, ''), isFormulonError);
  } finally {
    b.wb.dispose();
  }
});

test('pivot options read back every member and keep what a partial write omits', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const before = b.wb.sheet(0).getPivotOptions(b.pivot);
    assert.equal(before.errorCaption, null);
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
    b.wb.sheet(0).setPivotOptions(b.pivot, patch);
    const after = b.wb.sheet(0).getPivotOptions(b.pivot);
    for (const [key, value] of Object.entries(patch)) assert.equal(after[key], value, key);
    assert.equal(after.showHeaders, before.showHeaders);
    assert.equal(after.indent, before.indent);
    // A write of the whole read-back is a no-op, and null or '' clears a string.
    b.wb.sheet(0).setPivotOptions(b.pivot, { ...after, errorCaption: null });
    assert.equal(b.wb.sheet(0).getPivotOptions(b.pivot).errorCaption, null);
    assert.equal(b.wb.sheet(0).getPivotOptions(b.pivot).altText, 'Sales');
    b.wb.sheet(0).setPivotOptions(b.pivot, { grandTotalCaption: '' });
    assert.equal(b.wb.sheet(0).getPivotOptions(b.pivot).grandTotalCaption, null);
  } finally {
    b.wb.dispose();
  }
});

test('pivot options reject out-of-range values and wrong types without changing anything', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const base = b.wb.sheet(0).getPivotOptions(b.pivot);
    const unchanged = () => assert.deepEqual(b.wb.sheet(0).getPivotOptions(b.pivot), base);
    for (const bad of [{ indent: 127 }, { pageWrap: 256 }, { errorCaption: 'x'.repeat(256) }]) {
      assert.throws(
        () => b.wb.sheet(0).setPivotOptions(b.pivot, bad),
        isFormulonError,
        JSON.stringify(bad).slice(0, 40),
      );
      unchanged();
    }
    for (const bad of [{ indent: 'a' }, { errorCaption: 5 }]) {
      assert.throws(() => b.wb.sheet(0).setPivotOptions(b.pivot, bad), TypeError, JSON.stringify(bad));
      unchanged();
    }
    for (const bad of [{ indent: 1.5 }, { indent: -1 }, { pageWrap: Number.NaN }]) {
      assert.throws(() => b.wb.sheet(0).setPivotOptions(b.pivot, bad), RangeError, JSON.stringify(bad));
      unchanged();
    }
    assert.throws(() => b.wb.sheet(0).setPivotOptions(b.pivot, 'x'), TypeError);
    assert.throws(() => b.wb.sheet(0).setPivotOptions(b.pivot), TypeError);
    // A string is a pivot name, even one that reads as a number.
    assert.throws(
      () => b.wb.sheet(0).setPivotOptions(String(b.pivot), {}),
      (e) => isFormulonError(e, mod.StatusCode.NotFound),
    );
    assert.throws(() => b.wb.sheet(0).setPivotOptions(true, {}), TypeError);
    assert.throws(() => b.wb.sheet(2 ** 32).setPivotOptions(b.pivot, {}));
    unchanged();
    assert.throws(() => b.wb.sheet(99).setPivotOptions(0, {}), isFormulonError);
    assert.throws(() => b.wb.sheet(99).getPivotOptions(0), isFormulonError);

    const cacheBase = b.wb.getPivotCacheOptions(b.cacheId);
    assert.throws(() => b.wb.setPivotCacheOptions(b.cacheId, { missingItemsLimit: 1048577 }), isFormulonError);
    assert.throws(() => b.wb.setPivotCacheOptions(b.cacheId, { missingItemsLimit: 'a' }), TypeError);
    for (const bad of [{ missingItemsLimit: 1.5 }, { missingItemsLimit: Number.NaN }, { missingItemsLimit: -2 }]) {
      assert.throws(() => b.wb.setPivotCacheOptions(b.cacheId, bad), RangeError, JSON.stringify(bad));
    }
    assert.throws(() => b.wb.setPivotCacheOptions(b.cacheId, null), TypeError);
    assert.throws(() => b.wb.setPivotCacheOptions('1', {}), TypeError);
    assert.deepEqual(b.wb.getPivotCacheOptions(b.cacheId), cacheBase);
    assert.throws(() => b.wb.setPivotCacheOptions(9999, {}), isFormulonError);
    assert.throws(() => b.wb.getPivotCacheOptions(9999), isFormulonError);
  } finally {
    b.wb.dispose();
  }
});

test('calculated fields and items reject bad arguments and change nothing', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    const fields = b.wb.listPivotCacheFields(b.cacheId).length;
    // Wrong types throw before anything runs.
    assert.throws(() => b.wb.addPivotCacheCalculatedField(b.cacheId, 'D', 5), TypeError);
    assert.throws(() => b.wb.addPivotCacheCalculatedField(b.cacheId, null, 'Amount'), TypeError);
    assert.throws(() => b.wb.addPivotCacheCalculatedField(b.cacheId, 'D'), TypeError);
    assert.throws(() => b.wb.addPivotCacheCalculatedField('1', 'D', 'Amount'), TypeError);
    assert.throws(() => b.wb.addPivotCacheCalculatedItem(b.cacheId, 0.5, 'N', 'East'), RangeError);
    assert.throws(() => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'N', undefined), TypeError);
    assert.throws(() => b.wb.movePivotCacheCalculatedItem(b.cacheId, 0, -1), RangeError);
    assert.throws(() => b.wb.listPivotCacheCalculatedItems('a'), TypeError);
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
      assert.throws(
        () => b.wb.addPivotCacheCalculatedField(b.cacheId, name, formula),
        isFormulonError,
        `${name}=${formula.slice(0, 8)}`,
      );
    }
    assert.throws(() => b.wb.addPivotCacheCalculatedField(9999, 'D', 'Amount'), isFormulonError);
    assert.equal(b.wb.listPivotCacheFields(b.cacheId).length, fields);
    // Not a calculated field.
    assert.throws(() => b.wb.setPivotCacheCalculatedFieldFormula(b.cacheId, 0, 'Amount'), isFormulonError);
    assert.throws(() => b.wb.removePivotCacheCalculatedField(b.cacheId, 0), isFormulonError);
    assert.throws(() => b.wb.removePivotCacheCalculatedField(b.cacheId, 2 ** 31), isFormulonError);
    assert.throws(() => b.wb.listPivotCacheFields(9999), isFormulonError);
    // A calculated field a second one names cannot be removed.
    b.wb.addPivotCacheCalculatedField(b.cacheId, 'Base', 'Amount');
    b.wb.addPivotCacheCalculatedField(b.cacheId, 'Top', 'Base*2');
    assert.throws(() => b.wb.removePivotCacheCalculatedField(b.cacheId, fields), isFormulonError);
    assert.equal(b.wb.listPivotCacheFields(b.cacheId).length, fields + 2);

    // Items: unplaced fields, duplicates, and out-of-range indices.
    assert.throws(() => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.date, 'N', 'Amount'), isFormulonError);
    assert.throws(() => b.wb.addPivotCacheCalculatedItem(b.cacheId, 99, 'N', 'East'), isFormulonError);
    assert.throws(() => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'East', 'West'), isFormulonError);
    assert.throws(() => b.wb.addPivotCacheCalculatedItem(b.cacheId, b.region, 'N', 'Nope'), isFormulonError);
    assert.throws(() => b.wb.addPivotCacheCalculatedItem(9999, b.region, 'N', 'East'), isFormulonError);
    assert.equal(b.wb.listPivotCacheCalculatedItems(b.cacheId).length, 0);
    assert.throws(() => b.wb.listPivotCacheCalculatedItems(9999), isFormulonError);
    assert.throws(() => b.wb.setPivotCacheCalculatedItemFormula(b.cacheId, 0, 'East'), isFormulonError);
    assert.throws(() => b.wb.movePivotCacheCalculatedItem(b.cacheId, 0, 0), isFormulonError);
    assert.throws(() => b.wb.removePivotCacheCalculatedItem(b.cacheId, 0), isFormulonError);
    assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
  } finally {
    b.wb.dispose();
  }
});

test('pivot field custom names, show-all and default data field names report bad arguments', async () => {
  const mod = await getModule();
  const b = buildPivot(mod);
  try {
    assert.throws(() => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { customName: 5 }), TypeError);
    assert.throws(() => b.wb.sheet(0).updatePivotField(b.pivot, b.region), TypeError);
    assert.throws(() => b.wb.sheet(0).updatePivotField(b.pivot, -1, { customName: 'x' }), RangeError);
    assert.throws(() => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { showAll: null }), TypeError);
    assert.throws(() => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { axis: 0 }), TypeError);
    assert.throws(() => b.wb.sheet(0).updatePivotField(b.pivot, 99, { customName: 'x' }), isFormulonError);
    assert.throws(
      () => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { customName: 'x'.repeat(256) }),
      isFormulonError,
    );
    // Another field's source name or display name is taken.
    assert.throws(() => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { customName: 'Product' }), isFormulonError);
    assert.throws(
      () => b.wb.sheet(0).updatePivotField(b.pivot, b.region, { customName: 'Sum of Amount' }),
      isFormulonError,
    );
    // The field's own source name is accepted.
    b.wb.sheet(0).updatePivotField(b.pivot, b.region, { customName: 'Region' });
    assert.throws(() => b.wb.sheet(0).updatePivotField(b.pivot, 99, { showAll: true }), isFormulonError);

    // The first derived name is free; once taken, a new one is numbered
    // unless the entry that holds it is the one being replaced.
    assert.equal(b.wb.sheet(0).getPivotDataFieldDefaultName(b.pivot, b.amount, 0), '合計 / Amount');
    b.wb.sheet(0).addPivotDataField(b.pivot, { fieldIndex: b.amount, aggregation: 0 });
    assert.equal(b.wb.sheet(0).listPivotDataFields(b.pivot)[1].name, '合計 / Amount');
    assert.equal(b.wb.sheet(0).getPivotDataFieldDefaultName(b.pivot, b.amount, 0), '合計 / Amount2');
    assert.equal(b.wb.sheet(0).getPivotDataFieldDefaultName(b.pivot, b.amount, 0, undefined), '合計 / Amount2');
    assert.equal(b.wb.sheet(0).getPivotDataFieldDefaultName(b.pivot, b.amount, 0, 1), '合計 / Amount');
    assert.throws(() => b.wb.sheet(0).getPivotDataFieldDefaultName(b.pivot, b.amount, 0, 99), isFormulonError);
    assert.equal(b.wb.sheet(0).listPivotDataFields(b.pivot).length, 2);
    assert.throws(() => b.wb.sheet(0).getPivotDataFieldDefaultName(b.pivot, 99, 0), isFormulonError);
    assert.throws(() => b.wb.sheet(0).getPivotDataFieldDefaultName(b.pivot, b.amount, 99), isFormulonError);
    assert.throws(() => b.wb.sheet(99).getPivotDataFieldDefaultName(0, 0, 0), isFormulonError);
    assert.throws(() => b.wb.sheet(0).getPivotDataFieldDefaultName(b.pivot, b.amount, 'sum'), TypeError);
    assert.throws(() => b.wb.sheet(0).getPivotDataFieldDefaultName(b.pivot, b.amount, 0, null), TypeError);
    assert.throws(() => b.wb.sheet(0).getPivotDataFieldDefaultName(b.pivot, b.amount), TypeError);
    assert.equal(b.wb.sheet(0).listPivotDataFields(b.pivot)[0].name, 'Sum of Amount');
    assert.throws(() => b.wb.sheet(99).projectPivot(0), isFormulonError);
    assert.throws(() => b.wb.sheet(99).isPivotProjectionPending(0), isFormulonError);
    assert.throws(() => b.wb.sheet(0).projectPivot(), TypeError);
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
    assert.throws(() => b.wb.sheet(0).addPivotFilter(b.pivot, spec), isFormulonError);
    b.wb.sheet(0).addPivotFilter(b.pivot, { ...spec, valueHighText: 'F' });
    const at = b.wb.sheet(0).listPivotFilters(b.pivot)[0];
    assert.equal(at.type, 22);
    assert.equal(at.valueText, 'A');
    assert.equal(at.valueHighText, 'F');
    // And rejected for every other type.
    assert.throws(
      () =>
        b.wb.sheet(0).addPivotFilter(b.pivot, { ...spec, type: mod.PivotFilterType.LabelEqual, valueHighText: 'F' }),
      isFormulonError,
    );
    assert.throws(() => b.wb.sheet(0).addPivotFilter(b.pivot, { ...spec, valueHighText: 5 }), TypeError);
    assert.equal(b.wb.sheet(0).listPivotFilters(b.pivot).length, 1);
  } finally {
    b.wb.dispose();
  }
});
