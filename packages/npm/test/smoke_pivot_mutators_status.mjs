import test from 'node:test';
import { inflateRawSync } from 'node:zlib';
import { assert, getModule, makeIterativeWorkbook, path, pkgRoot, readFile } from './smoke_support.mjs';

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
  const cacheId = must(wb.pivotCacheCreate(0), 'pivotCacheCreate').index;
  must(wb.pivotCacheSetWorksheetSource(cacheId, { present: true, ref: 'A1:D5', sheet: 'Sheet1' }), 'source');
  for (const name of ['Region', 'Product', 'Amount', 'Date']) must(wb.pivotCacheFieldAdd(cacheId, name), name);
  for (const item of ['East', 'West']) must(wb.pivotCacheFieldAddSharedItemText(cacheId, 0, item), item);
  for (const item of ['Apple', 'Pear']) must(wb.pivotCacheFieldAddSharedItemText(cacheId, 1, item), item);
  for (const values of [
    [0, 0, 10, 45292],
    [1, 0, 30, 45323],
    [0, 1, 5, 45658],
    [1, 1, 7, 45689],
  ]) {
    const rec = must(wb.pivotCacheRecordAdd(cacheId), 'record').index;
    for (const [field, v] of values.entries())
      must(wb.pivotCacheRecordSetNumber(cacheId, rec, field, v), 'record value');
  }
  const pivot = must(wb.pivotCreate(0, 'Pivot1', cacheId, 0, 6), 'pivotCreate').index;
  const addField = (spec) => must(wb.pivotFieldAdd(0, pivot, spec), spec.sourceName).index;
  const region = addField({ sourceName: 'Region', axis: 0 });
  must(wb.pivotFieldAddItem(0, pivot, region, 'East', true), 'item');
  must(wb.pivotFieldAddItem(0, pivot, region, 'West', true), 'item');
  const product = addField({ sourceName: 'Product', axis: 2 });
  must(wb.pivotFieldAddItem(0, pivot, product, 'Apple', true), 'item');
  must(wb.pivotFieldAddItem(0, pivot, product, 'Pear', true), 'item');
  const amount = addField({ sourceName: 'Amount', axis: 2 });
  const date = addField({ sourceName: 'Date', axis: 2 });
  must(wb.pivotSetRowFieldOrder(0, pivot, [region]), 'row order');
  must(wb.pivotDataFieldAdd(0, pivot, { name: 'Sum of Amount', fieldIndex: amount, aggregation: 0 }), 'data field');
  return { wb, cacheId, pivot, region, product, amount, date };
}

// The projected grid as one `|`-joined string per sheet row.
function pivotGrid(wb, pivot = 0) {
  const layout = wb.pivotLayout(0, pivot);
  assert.ok(layout.status.ok, JSON.stringify(layout.status));
  const rows = new Map();
  for (const cell of layout.cells) {
    const v = cell.value;
    const text = v.kind === 1 ? String(v.number) : v.kind === 3 ? v.text : v.kind === 4 ? `#${v.errorCode}` : '';
    if (!rows.has(cell.row)) rows.set(cell.row, new Array(layout.cols).fill(''));
    rows.get(cell.row)[cell.col - layout.left] = text;
  }
  return [...rows.keys()].sort((a, b) => a - b).map((row) => rows.get(row).join('|'));
}

const BASE_GRID = ['行ラベル|Sum of Amount', 'East|15', 'West|37', '総計|52'];
const TWO_LEVEL_GRID = [
  '行ラベル|Sum of Amount',
  'Apple|10',
  'Pear|5',
  'East|15',
  'Apple|30',
  'Pear|7',
  'West|37',
  '総計|52',
];
const DATE_SERIAL_GRID = ['行ラベル|Sum of Amount', '45292|10', '45323|30', '45658|5', '45689|7', '総計|52'];
const twoLevels = (b) => b.wb.pivotSetRowFieldOrder(0, b.pivot, [b.region, b.product]);
const dateRows = (b) => b.wb.pivotSetRowFieldOrder(0, b.pivot, [b.date]);
const hideEast = (b) => b.wb.pivotFieldSetItemVisible(0, b.pivot, b.region, 0, false);
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
const pointAtNewItem = (b) => b.wb.pivotCacheRecordSetNumber(b.cacheId, 0, 0, 2);

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
    'pivotCacheCreate',
    {
      act: (b) => b.wb.pivotCacheCreate(77),
      expect: (b, r) => {
        assert.equal(r.index, 77);
        assert.equal(b.wb.pivotCacheCount().value, 2);
      },
    },
  ],
  ['pivotCacheIdAt', { act: (b) => b.wb.pivotCacheIdAt(0), expect: (b, r) => assert.equal(r.index, b.cacheId) }],
  [
    'pivotCacheRemove',
    {
      setup: (b) => {
        b.spare = b.wb.pivotCacheCreate(0).index;
      },
      act: (b) => b.wb.pivotCacheRemove(b.spare),
      expect: (b) => {
        assert.equal(b.wb.pivotCacheCount().value, 1);
        assert.equal(b.wb.pivotCacheIdAt(0).index, b.cacheId);
      },
    },
  ],
  [
    'pivotCacheSetWorksheetSource',
    {
      act: (b) => b.wb.pivotCacheSetWorksheetSource(b.cacheId, { present: true, ref: 'B2:E6', sheet: 'Data' }),
      expect: (b) => {
        const source = b.wb.pivotCacheGetWorksheetSource(b.cacheId);
        assert.equal(source.ref, 'B2:E6');
        assert.equal(source.sheet, 'Data');
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'pivotCacheFieldAdd',
    {
      act: (b) => b.wb.pivotCacheFieldAdd(b.cacheId, 'Extra'),
      expect: (b, r) => {
        assert.equal(r.index, 4);
        assert.equal(b.wb.pivotCacheFieldName(b.cacheId, 4).value, 'Extra');
      },
    },
  ],
  [
    'pivotCacheFieldClear',
    {
      act: (b) => b.wb.pivotCacheFieldClear(b.cacheId),
      expect: (b) => {
        assert.equal(b.wb.pivotCacheFieldCount(b.cacheId).value, 0);
        // The data field now names a cache field that no longer exists.
        assert.equal(b.wb.pivotLayout(0, b.pivot).status.ok, false);
      },
    },
  ],
  [
    'pivotCacheFieldAddSharedItemNumber',
    {
      act: (b) => b.wb.pivotCacheFieldAddSharedItemNumber(b.cacheId, 0, 42.5),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '42.5|10', 'East|5', 'West|37', '総計|52']),
    },
  ],
  [
    'pivotCacheFieldAddSharedItemText',
    {
      act: (b) => b.wb.pivotCacheFieldAddSharedItemText(b.cacheId, 0, 'North'),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'North|10', 'West|37', '総計|52']),
    },
  ],
  [
    'pivotCacheFieldAddSharedItemBool',
    {
      act: (b) => b.wb.pivotCacheFieldAddSharedItemBool(b.cacheId, 0, true),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', 'TRUE|10', '総計|52']),
    },
  ],
  [
    'pivotCacheFieldAddSharedItemBlank',
    {
      act: (b) => b.wb.pivotCacheFieldAddSharedItemBlank(b.cacheId, 0),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '(空白)|10', '総計|52']),
    },
  ],
  [
    'pivotCacheFieldAddSharedItemError',
    {
      act: (b) => b.wb.pivotCacheFieldAddSharedItemError(b.cacheId, 0, 6),
      after: pointAtNewItem,
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '#N/A|10', '総計|52']),
    },
  ],
  [
    'pivotCacheFieldClearSharedItems',
    {
      act: (b) => b.wb.pivotCacheFieldClearSharedItems(b.cacheId, 0),
      // With no shared items the records' indices render as themselves.
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '0|15', '1|37', '総計|52']),
    },
  ],
  [
    'pivotCacheRecordAdd',
    {
      act: (b) => b.wb.pivotCacheRecordAdd(b.cacheId),
      expect: (b, r) => {
        assert.equal(r.index, 4);
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|15', 'West|37', '(空白)|0', '総計|52']);
      },
    },
  ],
  [
    'pivotCacheRecordClear',
    {
      act: (b) => b.wb.pivotCacheRecordClear(b.cacheId),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '総計|0']),
    },
  ],
  [
    'pivotCacheRecordSetNumber',
    {
      act: (b) => b.wb.pivotCacheRecordSetNumber(b.cacheId, 0, 2, 1000),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|1005', 'West|37', '総計|1042']),
    },
  ],
  [
    'pivotCacheRecordSetText',
    {
      act: (b) => b.wb.pivotCacheRecordSetText(b.cacheId, 0, 0, 'West'),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|47', '総計|52']),
    },
  ],
  [
    'pivotCacheRecordSetBool',
    {
      act: (b) => b.wb.pivotCacheRecordSetBool(b.cacheId, 0, 2, true),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|6', 'West|37', '総計|43']),
    },
  ],
  [
    'pivotCacheRecordSetBlank',
    {
      act: (b) => b.wb.pivotCacheRecordSetBlank(b.cacheId, 0, 2),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
    },
  ],
  [
    'pivotCacheRecordSetError',
    {
      act: (b) => b.wb.pivotCacheRecordSetError(b.cacheId, 0, 2, 1),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|#1', 'West|37', '総計|#1']),
    },
  ],
  [
    'pivotCreate',
    {
      act: (b) => b.wb.pivotCreate(0, 'Pivot2', b.cacheId, 20, 1),
      expect: (b, r) => {
        assert.equal(r.index, 1);
        const layout = b.wb.pivotLayout(0, 1);
        assert.equal(layout.top, 20);
        assert.equal(layout.left, 1);
        assert.deepEqual(pivotGrid(b.wb, 0), BASE_GRID);
      },
    },
  ],
  [
    'pivotRemove',
    {
      act: (b) => b.wb.pivotRemove(0, b.pivot),
      expect: (b) => {
        assert.equal(b.wb.pivotCount(0).value, 0);
        assert.equal(b.wb.pivotLayout(0, 0).status.ok, false);
      },
    },
  ],
  [
    'pivotSetName',
    {
      act: (b) => b.wb.pivotSetName(0, b.pivot, 'Renamed'),
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
    'pivotSetAnchor',
    {
      act: (b) => b.wb.pivotSetAnchor(0, b.pivot, 10, 2, 4, 2),
      expect: (b) => {
        const layout = b.wb.pivotLayout(0, b.pivot);
        assert.equal(layout.top, 10);
        assert.equal(layout.left, 2);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'pivotSetGrandTotals',
    {
      setup: (b) => b.wb.pivotSetColFieldOrder(0, b.pivot, [b.product]),
      // Rows off drops the per-row total column; columns on keeps the
      // bottom total row, so a swapped pair would show the opposite.
      act: (b) => b.wb.pivotSetGrandTotals(0, b.pivot, false, true),
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
    'pivotSetLayout',
    {
      act: (b) => b.wb.pivotSetLayout(0, b.pivot, 1),
      expect: (b) => {
        assert.equal(b.wb.pivotGetLayout(0, b.pivot).layout, 1);
        assert.deepEqual(pivotGrid(b.wb), ['Region|Sum of Amount', 'East|15', 'West|37', '総計|52']);
      },
    },
  ],
  [
    'pivotFieldAdd',
    {
      act: (b) => b.wb.pivotFieldAdd(0, b.pivot, { sourceName: 'Product', customName: 'Fruit', axis: 3 }),
      expect: (b, r) => {
        assert.equal(r.index, 4);
        assert.deepEqual(pivotGrid(b.wb), ['Fruit|(すべて)', '|', ...BASE_GRID]);
      },
    },
  ],
  [
    'pivotFieldClear',
    {
      setup: hideEast,
      act: (b) => b.wb.pivotFieldClear(0, b.pivot),
      expect: (b) => {
        assert.equal(b.wb.pivotFieldCount(0, b.pivot).value, 0);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'pivotFieldSetAxis',
    {
      act: (b) => b.wb.pivotFieldSetAxis(0, b.pivot, b.product, 3),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['Product|(すべて)', '|', ...BASE_GRID]),
    },
  ],
  [
    'pivotFieldSetSort',
    {
      // East now outsums West, so sorting ascending by the data field
      // differs from both label orders.
      setup: (b) => b.wb.pivotCacheRecordSetNumber(b.cacheId, 0, 2, 40),
      act: (b) => b.wb.pivotFieldSetSort(0, b.pivot, b.region, true, 'Sum of Amount'),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'West|37', 'East|45', '総計|82']),
    },
  ],
  [
    'pivotFieldSetSubtotalTop',
    {
      setup: twoLevels,
      act: (b) => b.wb.pivotFieldSetSubtotalTop(0, b.pivot, b.region, true),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), [
          '行ラベル|Sum of Amount',
          'East|15',
          'Apple|10',
          'Pear|5',
          'West|37',
          'Apple|30',
          'Pear|7',
          '総計|52',
        ]),
    },
  ],
  [
    'pivotFieldAddItem',
    {
      act: (b) => b.wb.pivotFieldAddItem(0, b.pivot, b.region, 'East', false),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
  [
    'pivotFieldAddItemAt',
    {
      // Only the index-addressed form can name the blank item.
      setup: (b) => {
        b.wb.pivotCacheFieldAddSharedItemBlank(b.cacheId, 0);
        b.wb.pivotCacheRecordSetBlank(b.cacheId, 0, 0);
      },
      act: (b) => b.wb.pivotFieldAddItemAt(0, b.pivot, b.region, 2, false),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|5', 'West|37', '総計|42']),
    },
  ],
  [
    'pivotFieldClearItems',
    {
      setup: hideEast,
      act: (b) => b.wb.pivotFieldClearItems(0, b.pivot, b.region),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'pivotFieldSetItemVisible',
    {
      act: (b) => b.wb.pivotFieldSetItemVisible(0, b.pivot, b.region, 1, false),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'East|15', '総計|15']),
    },
  ],
  [
    'pivotFieldAddSubtotalFn',
    {
      setup: twoLevels,
      act: (b) => b.wb.pivotFieldAddSubtotalFn(0, b.pivot, b.region, 3),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), [
          '行ラベル|Sum of Amount',
          'Apple|10',
          'Pear|5',
          'East|10',
          'Apple|30',
          'Pear|7',
          'West|30',
          '総計|52',
        ]),
    },
  ],
  [
    'pivotFieldClearSubtotalFns',
    {
      setup: (b) => {
        twoLevels(b);
        b.wb.pivotFieldAddSubtotalFn(0, b.pivot, b.region, 3);
      },
      act: (b) => b.wb.pivotFieldClearSubtotalFns(0, b.pivot, b.region),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), TWO_LEVEL_GRID),
    },
  ],
  [
    'pivotFieldSetDateGroup',
    {
      setup: dateRows,
      act: (b) => b.wb.pivotFieldSetDateGroup(0, b.pivot, b.date, 3, 1, -1, -1),
      expect: (b) =>
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', '令和6年|40', '令和7年|12', '総計|52']),
    },
  ],
  [
    'pivotFieldClearDateGroup',
    {
      setup: (b) => {
        dateRows(b);
        b.wb.pivotFieldSetDateGroup(0, b.pivot, b.date, 3, 0, -1, -1);
      },
      act: (b) => b.wb.pivotFieldClearDateGroup(0, b.pivot, b.date),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), DATE_SERIAL_GRID),
    },
  ],
  [
    'pivotFieldSetNumberFormat',
    {
      act: (b) => b.wb.pivotFieldSetNumberFormat(0, b.pivot, b.amount, '4'),
      // A field's numFmtId is only observable in the part the writer emits.
      expect: (b) => {
        assert.equal(b.wb.pivotFieldSetNumberFormat(0, b.pivot, 99, '4').status, 2);
        assert.equal(b.wb.pivotFieldSetNumberFormat(0, b.pivot, b.amount, '0.00').status, 2);
        const saved = b.wb.save();
        assert.ok(saved.status.ok, JSON.stringify(saved.status));
        assert.match(packagePart(saved.bytes, 'xl/pivotTables/pivotTable1.xml'), /<pivotField [^>]*numFmtId="4"/);
        assert.deepEqual(pivotGrid(b.wb), BASE_GRID);
      },
    },
  ],
  [
    'pivotSetRowFieldOrder',
    {
      act: (b) => b.wb.pivotSetRowFieldOrder(0, b.pivot, [b.product]),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Sum of Amount', 'Apple|40', 'Pear|12', '総計|52']),
    },
  ],
  [
    'pivotSetColFieldOrder',
    {
      act: (b) => b.wb.pivotSetColFieldOrder(0, b.pivot, [b.product]),
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
    'pivotDataFieldAdd',
    {
      act: (b) => b.wb.pivotDataFieldAdd(0, b.pivot, { name: 'Max of Amount', fieldIndex: b.amount, aggregation: 3 }),
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
    'pivotDataFieldClear',
    {
      act: (b) => b.wb.pivotDataFieldClear(0, b.pivot),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), ['行ラベル', 'East', 'West', '総計']),
    },
  ],
  [
    'pivotDataFieldSet',
    {
      act: (b) =>
        b.wb.pivotDataFieldSet(0, b.pivot, 0, {
          name: 'Count of Amount',
          fieldIndex: b.amount,
          aggregation: 1,
          numberFormat: '2',
        }),
      expect: (b) => {
        assert.deepEqual(pivotGrid(b.wb), ['行ラベル|Count of Amount', 'East|2', 'West|2', '総計|4']);
        const counts = b.wb.pivotLayout(0, b.pivot).cells.filter((c) => c.value.kind === 1);
        assert.deepEqual(
          counts.map((c) => c.numberFormat),
          ['2', '2', '2'],
        );
      },
    },
  ],
  [
    'pivotFilterAdd',
    {
      act: (b) => b.wb.pivotFilterAdd(0, b.pivot, GREATER_THAN_20),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
  [
    'pivotFilterClear',
    {
      setup: (b) => b.wb.pivotFilterAdd(0, b.pivot, GREATER_THAN_20),
      act: (b) => b.wb.pivotFilterClear(0, b.pivot),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), BASE_GRID),
    },
  ],
  [
    'pivotFilterRemoveAt',
    {
      // Each filter alone keeps a different region, so removing the wrong
      // index leaves the wrong one.
      setup: (b) => {
        b.wb.pivotFilterAdd(0, b.pivot, GREATER_THAN_20);
        b.wb.pivotFilterAdd(0, b.pivot, BETWEEN_10_20);
      },
      act: (b) => b.wb.pivotFilterRemoveAt(0, b.pivot, 1),
      expect: (b) => assert.deepEqual(pivotGrid(b.wb), WEST_ONLY),
    },
  ],
];

test('the pivot mutator table covers every declared pivot mutator', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
  const declared = [...dts.matchAll(/^ {2}(pivot\w*)\([^)]*\): (?:Status|AddStyleResult);/gm)].map((m) => m[1]);
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
    base.wb.delete();
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
      b.wb.delete();
    }
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
  ['sheetCount', [], null],
  ['cellCount', [0], [99]],
  ['definedNameCount', [], null],
  ['tableCount', [], null],
  ['passthroughCount', [], null],
  ['pivotCount', [0], [99]],
  ['pivotCacheCount', [], null],
  ['pivotCacheFieldCount', null, [9999]],
  ['pivotCacheFieldSharedItemCount', null, [9999, 0]],
  ['pivotCacheRecordCount', null, [9999]],
  ['pivotFieldCount', null, [99, 0]],
  ['pivotDataFieldCount', null, [99, 0]],
  ['pivotFilterCount', null, [99, 0]],
  ['fontCount', [], null],
  ['fillCount', [], null],
  ['borderCount', [], null],
  ['xfCount', [], null],
  ['dxfCount', [], null],
  ['cellStyleCount', [], null],
  ['cellStyleXfCount', [], null],
  ['calcMode', [], null],
  ['excelProfileId', [], null],
  ['localizeFunctionName', ['SUM', 0], ['NOPE_XYZ', 0]],
  ['canonicalizeFunctionName', ['SUM', 0], ['NOPE_XYZ', 0]],
  ['pinnedNow', [], null],
  ['spillInfo', [0, 0, 0], [99, 0, 0]],
  ['getColumnWidthPt', [0, 0, 0], [99, 0, 0]],
  ['getRowHeightPt', [0, 0], [99, 0]],
  ['columnCharsToPt', [0, 0, 8], [99, 0, 8]],
  ['columnPtToChars', [0, 0, 10], [99, 0, 10]],
  ['precedents', [0, 0, 0, 1], [99, 0, 0, 1]],
  ['dependents', [0, 0, 0, 1], [99, 0, 0, 1]],
  ['functionNames', [], null],
];

// The payload a failed call must fall back to, keyed by what it returns.
function assertFailurePayload(name, r) {
  if (Array.isArray(r)) assert.equal(r.length, 0, name);
  else if ('now' in r) assert.equal(r.now, null, name);
  else if ('engaged' in r) assert.equal(r.engaged, false, name);
  else if (typeof r.value === 'string') assert.equal(r.value, '', name);
  else assert.equal(r.value, 0, name);
}

test('every NumberResult accessor is covered by the status table', async () => {
  const dts = await readFile(path.join(pkgRoot, 'dist', 'formulon.d.ts'), 'utf8');
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
  assert.equal(wb.sheetCount().value, 1);
  assert.equal(wb.excelProfileId().value, 'win-365-ja_JP');
  assert.equal(wb.localizeFunctionName('SUM', 1).value, 'SUM');
  assert.equal(wb.localizeFunctionName('SUM', 99).status.ok, false);

  wb.delete();
  // `delete()` detaches the embind wrapper outright; a failed load is the
  // wrapper that outlives its handle on this surface.
  const dead = mod.Workbook.loadBytes(new Uint8Array([0]));
  assert.equal(dead.isValid(), false);
  // Every accessor, including the ones no argument can fail, reports the
  // released handle instead of a default reading.
  for (const [name, okArgs, badArgs] of STATUS_ACCESSORS) {
    if (name.endsWith('FunctionName') || name === 'functionNames') continue;
    const r = dead[name](...(okArgs ?? badArgs));
    assert.equal(r.status.ok, false, `${name} after release`);
    assert.equal(r.status.status, 7000, name);
    assertFailurePayload(name, r);
  }
  dead.delete();
});

test('iterative progress callbacks are per Workbook, not per module', async () => {
  const Module = await getModule();
  const first = makeIterativeWorkbook(Module);
  const second = makeIterativeWorkbook(Module);
  try {
    let firstCalls = 0;
    let secondCalls = 0;
    assert.ok(
      first.setIterativeProgress(() => {
        firstCalls += 1;
        return true;
      }).ok,
    );
    // Installing on `second` must not displace the registration on `first`.
    assert.ok(
      second.setIterativeProgress(() => {
        secondCalls += 1;
        return true;
      }).ok,
    );

    assert.ok(first.recalc().ok);
    assert.ok(firstCalls > 0, 'the first workbook kept its own callback');
    assert.equal(secondCalls, 0);

    firstCalls = 0;
    assert.ok(second.recalc().ok);
    assert.ok(secondCalls > 0);
    assert.equal(firstCalls, 0);
  } finally {
    first.setIterativeProgress(null);
    second.setIterativeProgress(null);
    first.delete();
    second.delete();
  }
});
